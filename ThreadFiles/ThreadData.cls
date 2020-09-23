VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ThreadDone As Long
Private m_ThreadSignal As Long
Private m_ThreadHandle As Long
Private m_Data As Variant
Private m_Controller As ThreadControl
Friend Function ThreadCompleted() As Boolean
Dim ExitCode As Long
    ThreadCompleted = m_ThreadDone
    If ThreadCompleted Then
        'Since code runs on the worker thread after the
        'ThreadDone pointer is incremented, there is a chance
        'that we are signalled, but the thread hasn't yet
        'terminated.  In this case, just claim we aren't done
        'yet to make sure that code on all worker threads is
        'actually completed before ThreadControl terminates.
        If m_ThreadHandle Then
            If GetExitCodeThread(m_ThreadHandle, ExitCode) Then
                If ExitCode = STILL_ACTIVE Then
                    ThreadCompleted = False
                    Exit Function
                End If
            End If
            CloseHandle m_ThreadHandle
            m_ThreadHandle = 0
        End If
    End If
End Function
Friend Property Get ThreadDonePointer() As Long
    ThreadDonePointer = VarPtr(m_ThreadDone)
End Property
Friend Property Let ThreadSignalPointer(ByVal RHS As Long)
    m_ThreadSignal = RHS
End Property
Friend Property Let ThreadHandle(ByVal RHS As Long)
    'This takes over ownership of the ThreadHandle
    m_ThreadHandle = RHS
End Property
Friend Sub SignalThread(ByVal pCritSect As Long, ByRef fInCriticalSection As Boolean)
    'm_ThreadDone and m_ThreadSignal must be checked/modified inside
    'a critical section because m_ThreadDone could change on some
    'threads while we are signalling, causing m_ThreadSignal to point
    'to invalid memory, as well as other problems. The parameters to this
    'function are provided to ensure that the critical section is entered
    'only when necessary. If fInCriticalSection is set, then the caller
    'must call LeaveCriticalSection on pCritSect. This is left up to the
    'caller since this function is designed to be called on multiple instances
    'in a tight loop. There is no point in repeatedly entering/leaving the
    'critical section.
    If m_ThreadSignal Then
        If Not fInCriticalSection Then
            EnterCriticalSection pCritSect
            fInCriticalSection = True
        End If
        If m_ThreadDone = 0 Then
            InterlockedIncrement m_ThreadSignal
        End If
        'No point in signalling twice
        m_ThreadSignal = 0
    End If
End Sub
Friend Property Set Controller(ByVal RHS As ThreadControl)
    Set m_Controller = RHS
End Property
Friend Sub SetData(Data As Variant, ByVal fStealData As Boolean)
    If IsEmpty(Data) Or IsMissing(Data) Then Exit Sub
    If fStealData Then
        CopyMemory ByVal VarPtr(m_Data), ByVal VarPtr(Data), 16
        CopyMemory ByVal VarPtr(Data), 0, 2
    ElseIf IsObject(Data) Then
        Set m_Data = Data
    Else
        m_Data = Data
    End If
End Sub
Friend Sub GetData(Data As Variant)
    'This is called only once. Always steal.
    'Before stealing, make sure there's
    'nothing lurking in Data
    Data = Empty
    CopyMemory ByVal VarPtr(Data), ByVal VarPtr(m_Data), 16
    CopyMemory ByVal VarPtr(m_Data), 0, 2
End Sub

Private Sub Class_Terminate()
    'This shouldn't happen, but just in case
    If m_ThreadHandle Then CloseHandle m_ThreadHandle
End Sub
