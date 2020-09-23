VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_RunningThreads As Collection   'Collection to hold ThreadData objects for each thread
Private m_fStoppingWorkers As Boolean    'Currently tearing down, so don't start anything new
Private m_EventHandle As Long            'Synchronization handle
Private m_CS As CRITICAL_SECTION         'Critical section to avoid conflicts when signalling threads
Private m_pCS As Long                    'Pointer to m_CS structure

'Called to create a new thread worker thread.
'CLSID can be obtained from a ProgID via CLSIDFromProgID
'Data contains the data for the new thread
'fStealData should be True if the data is large. If this
'  is set, then Data will be Empty on return. If Data
'  contains an object reference, then the object should
'  be created on this thread.
'fReturnThreadHandle must explicitly be set to True to
'  return the created thread handle. This handle can be
'  used for calls like SetThreadPriority and must be
'  closed with CloseHandle.
Friend Function CreateWorkerThread(CLSID As CLSID, Data As Variant, Optional ByVal fStealData As Boolean = False, Optional ByVal fReturnThreadHandle As Boolean = False) As Long
Dim TPD As ThreadProcData
Dim IID_IUnknown As VBGUID
Dim ThreadID As Long
Dim ThreadHandle As Long
Dim pStream As IUnknown
Dim ThreadData As ThreadData
Dim fCleanUpOnFailure As Boolean
Dim hProcess As Long
Dim pUnk As IUnknown
    If m_fStoppingWorkers Then Err.Raise 5, , "Can't create new worker while shutting down"
    CleanCompletedThreads 'We need to clean up sometime, this is as good a time as any
    With TPD
        Set ThreadData = New ThreadData
        .CLSID = CLSID
        .EventHandle = m_EventHandle
        With IID_IUnknown
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        .pMarshalStream = CoMarshalInterThreadInterfaceInStream(IID_IUnknown, Me)
        .ThreadDonePointer = ThreadData.ThreadDonePointer
        .ThreadDataCookie = ObjPtr(ThreadData)
        .pCritSect = m_pCS
        ThreadData.SetData Data, fStealData
        Set ThreadData.Controller = Me
        m_RunningThreads.Add ThreadData, CStr(.ThreadDataCookie)
    End With
    ThreadHandle = CreateThread(0, 0, AddressOf ThreadProc.ThreadStart, VarPtr(TPD), 0, ThreadID)
    If ThreadHandle = 0 Then
        fCleanUpOnFailure = True
    Else
        'Turn ownership of the thread handle over to
        'the ThreadData object
        ThreadData.ThreadHandle = ThreadHandle
        'Make sure we've been notified by ThreadProc before continuing to
        'guarantee that the new thread has gotten the data they need out
        'of the ThreadProcData structure
        WaitForSingleObject m_EventHandle, INFINITE
        If TPD.hr Then
            fCleanUpOnFailure = True
        ElseIf fReturnThreadHandle Then
            hProcess = GetCurrentProcess
            DuplicateHandle hProcess, ThreadHandle, hProcess, CreateWorkerThread
        End If
    End If
    If fCleanUpOnFailure Then
        'Failure, clean up stream by making a reference and releasing it
        CopyMemory pStream, TPD.pMarshalStream, 4
        Set pStream = Nothing
        'Tell the thread its done using the normal mechanism
        InterlockedIncrement TPD.ThreadDonePointer
        'There's no reason to keep the new thread data
        CleanCompletedThreads
    End If
    If TPD.hr Then Err.Raise TPD.hr
End Function

'Called after a thread is created to provide a mechanism
'to stop execution and retrieve initial data for running
'the thread.  Should be called in ThreadLaunch_Go with:
'Controller.RegisterNewThread ThreadDataCookie, VarPtr(m_Notify), Controller, Data
Public Sub RegisterNewThread(ByVal ThreadDataCookie As Long, ByVal ThreadSignalPointer As Long, ByRef ThreadControl As ThreadControl, Optional Data As Variant)
Dim ThreadData As ThreadData
Dim fInCriticalSection As Boolean
    Set ThreadData = m_RunningThreads(CStr(ThreadDataCookie))
    ThreadData.ThreadSignalPointer = ThreadSignalPointer
    ThreadData.GetData Data
    'The new thread should not own the controlling thread because
    'the controlling thread has to teardown after all of the worker
    'threads are done running code, which can't happen if we happen
    'to release the last reference to ThreadControl in a worker
    'thread.  ThreadData is already holding an extra reference on
    'this object, so it is guaranteed to remain alive until
    'ThreadData is signalled.
    Set ThreadControl = Nothing
    If m_fStoppingWorkers Then
        'This will only happen when StopWorkerThreads is called
        'almost immediately after CreateWorkerThread.  We could
        'just let this signal happen in the StopWorkerThreads loop,
        'but this allows a worker thread to be signalled immediately.
        'See note in SignalThread about CriticalSection usage.
        ThreadData.SignalThread m_pCS, fInCriticalSection
        If fInCriticalSection Then LeaveCriticalSection m_pCS
    End If
End Sub

'Call StopWorkerThreads to signal all worker threads
'and spin until they terminate. Any calls to an object
'passed via the Data parameter in CreateWorkerThread
'will succeed.
Friend Sub StopWorkerThreads()
Dim ThreadData As ThreadData
Dim fInCriticalSection As Boolean
Dim fSignal As Boolean
Dim fHaveOleThreadhWnd As Boolean
Dim OleThreadhWnd As Long
    If m_fStoppingWorkers Then Exit Sub
    m_fStoppingWorkers = True
    fSignal = True
    Do
        For Each ThreadData In m_RunningThreads
            If ThreadData.ThreadCompleted Then
                m_RunningThreads.Remove CStr(ObjPtr(ThreadData))
            ElseIf fSignal Then
                'See note in SignalThread about CriticalSection usage.
                ThreadData.SignalThread m_pCS, fInCriticalSection
            End If
        Next
        If fInCriticalSection Then
            LeaveCriticalSection m_pCS
            fInCriticalSection = False
        Else
            'We can turn this off indefinitely because new threads
            'which arrive at RegisterNewThread while stopping workers
            'are signalled immediately
            fSignal = False
        End If
        If m_RunningThreads.Count = 0 Then Exit Do
        'We need to clear the message queue here in order to allow
        'any pending RegisterNewThread messages to come through.
        If Not fHaveOleThreadhWnd Then
            OleThreadhWnd = FindOLEhWnd
            fHaveOleThreadhWnd = True
        End If
        SpinOlehWnd OleThreadhWnd, False
        Sleep 0
    Loop
    m_fStoppingWorkers = False
End Sub

'Releases ThreadData objects for all threads
'that are completed. Cleaning happens automatically
'when you call SignalWorkerThreads, StopWorkerThreads,
'and RegisterNewThread.
Friend Sub CleanCompletedThreads()
Dim ThreadData As ThreadData
    For Each ThreadData In m_RunningThreads
        If ThreadData.ThreadCompleted Then
            m_RunningThreads.Remove CStr(ObjPtr(ThreadData))
        End If
    Next
End Sub

'Call to tell all running worker threads to
'terminated. If the thread has not yet called
'RegisterNewThread, then it will not be signalled.
'Unlike StopWorkerThreads, this does not block
'while the workers actually terminate.
'SignalWorkerThreads must be called by the owner
'of this class before the ThreadControl instance
'is released.
Friend Sub SignalWorkerThreads()
Dim ThreadData As ThreadData
Dim fInCriticalSection As Boolean
    For Each ThreadData In m_RunningThreads
        If ThreadData.ThreadCompleted Then
            m_RunningThreads.Remove CStr(ObjPtr(ThreadData))
        Else
            'See note in SignalThread about CriticalSection usage.
            ThreadData.SignalThread m_pCS, fInCriticalSection
        End If
    Next
    If fInCriticalSection Then LeaveCriticalSection m_pCS
End Sub

Private Sub Class_Initialize()
    Set m_RunningThreads = New Collection
    m_EventHandle = CreateEvent(0, 0, 0, vbNullString)
    m_pCS = VarPtr(m_CS)
    InitializeCriticalSection m_pCS
End Sub

Private Sub Class_Terminate()
    CleanCompletedThreads                   'Just in case, this generally does nothing.
    Debug.Assert m_RunningThreads.Count = 0 'Each worker should have a reference to this class
    CloseHandle m_EventHandle
    DeleteCriticalSection m_pCS
End Sub
