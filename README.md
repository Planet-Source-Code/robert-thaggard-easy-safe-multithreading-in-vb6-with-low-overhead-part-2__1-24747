<div align="center">

## Easy, Safe Multithreading in Vb6 with Low Overhead \- Part 2


</div>

### Description

Use the ThreadingAPI type library to safely CreateProcess in vb6. by Matthew Curland
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |1999-05-07 10:41:18
**By**             |[Robert Thaggard](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-thaggard.md)
**Level**          |Advanced
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Easy, Safe22236752001\.zip](https://github.com/Planet-Source-Code/robert-thaggard-easy-safe-multithreading-in-vb6-with-low-overhead-part-2__1-24747/archive/master.zip)





### Source Code

<p><font size="4"><b>Easy, Stable VB6 Multithreading with Low Overhead - Part 2<br>
</b>Calling CreateThread Safely Within a DLL</font></p>
<p>I found some better, straight vb code for the tutorial I was going to do for
this part so I thought it would be better than using c++. The code does the same
thing.</p>
<p>Part 2 of thesse tutorials is based off Matthew Curland's &quot;Apartment
Threading in VB6, Safely and Externally&quot;. This uses a precompiled type
library to easily call CreateThread from a global name space (no declaration
required).&nbsp; While this might be use activex, it still isn't using nearly as
many system resources as Srideep's solution.</p>
<p>In addition to safely calling CreateThread from vb there are some thread
classes that are used for doing the work with class ids rather than function
addresses. (Launch, Worker, ThreadControl, ThreadData,
and ThreadLaunch)</p>
<p>I will list all the classes below. I also decide4d not to add syntax
highlighting because that took too long. Also please realize that I did not
write these. Matthew Curland did. So vote for him, not me.</p>
<table border="0" bgcolor="#C0C0C0">
 <tr>
  <td bgcolor="#FFFF00">ThreadControl.cls</td>
 </tr>
 <tr>
  <td><pre>Option Explicit
Private m_RunningThreads As Collection  'Collection to hold ThreadData objects for each thread
Private m_fStoppingWorkers As Boolean  'Currently tearing down, so don't start anything new
Private m_EventHandle As Long      'Synchronization handle
Private m_CS As CRITICAL_SECTION     'Critical section to avoid conflicts when signalling threads
Private m_pCS As Long          'Pointer to m_CS structure
'Called to create a new thread worker thread.
'CLSID can be obtained from a ProgID via CLSIDFromProgID
'Data contains the data for the new thread
'fStealData should be True if the data is large. If this
' is set, then Data will be Empty on return. If Data
' contains an object reference, then the object should
' be created on this thread.
'fReturnThreadHandle must explicitly be set to True to
' return the created thread handle. This handle can be
' used for calls like SetThreadPriority and must be
' closed with CloseHandle.
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
  If m_fStoppingWorkers Then Err.Raise 5, , &quot;Can't create new worker while shutting down&quot;
  CleanCompletedThreads 'We need to clean up sometime, this is as good a time as any
  With TPD
    Set ThreadData = New ThreadData
    .CLSID = CLSID
    .EventHandle = m_EventHandle
    With IID_IUnknown
      .Data4(0) = &amp;HC0
      .Data4(7) = &amp;H46
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
'the thread. Should be called in ThreadLaunch_Go with:
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
  'thread. ThreadData is already holding an extra reference on
  'this object, so it is guaranteed to remain alive until
  'ThreadData is signalled.
  Set ThreadControl = Nothing
  If m_fStoppingWorkers Then
    'This will only happen when StopWorkerThreads is called
    'almost immediately after CreateWorkerThread. We could
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
  CleanCompletedThreads          'Just in case, this generally does nothing.
  Debug.Assert m_RunningThreads.Count = 0 'Each worker should have a reference to this class
  CloseHandle m_EventHandle
  DeleteCriticalSection m_pCS
End Sub
</pre></td>
 </tr>
</table>
<p>&nbsp;
<table border="0" bgcolor="#C0C0C0">
 <tr>
  <td bgcolor="#FFFF00">Launch.cls</td>
 </tr>
 <tr>
  <td><pre>Option Explicit
Private Controller As ThreadControl
Public Sub LaunchThreads()
Dim CLSID As CLSID
  CLSID = CLSIDFromProgID(&quot;DllThreads.Worker&quot;)
  Controller.CreateWorkerThread CLSID, 3000, True
  Controller.CreateWorkerThread CLSID, 5000, True
  Controller.CreateWorkerThread CLSID, 7000
End Sub
Public Sub FinishThreads()
  Controller.StopWorkerThreads
End Sub
Public Sub CleanCompletedThreads()
  Controller.CleanCompletedThreads
End Sub
Private Sub Class_Initialize()
  Set Controller = New ThreadControl
End Sub
Private Sub Class_Terminate()
  Controller.StopWorkerThreads
  Set Controller = Nothing
End Sub</pre></td>
 </tr>
</table><br>
<table border="0" bgcolor="#C0C0C0">
 <tr>
  <td bgcolor="#FFFF00">ThreadLaunch.cls</td>
 </tr>
 <tr>
  <td><pre>Option Explicit
'Just an interface definition
Public Function Go(Controller As ThreadControl, ByVal ThreadDataCookie As Long) As Long
End Function
'The rest of this is a comment
#If False Then
'A worker thread should include the following code.
'The Instancing for a worker should be set to 5 - MultiUse
Implements ThreadLaunch
Private m_Notify As Long
Public Function ThreadLaunch_Go(Controller As ThreadControl, ByVal ThreadDataCookie As Long) As Long
Dim Data As Variant
  Controller.RegisterNewThread ThreadDataCookie, VarPtr(m_Notify), Controller, Data
  'TODO: Process Data while
  'regularly calling HaveBeenNotified to
  'see if the thread should terminate.
  If HaveBeenNotified Then
    'Clean up and return
  End If
End Function
Private Function HaveBeenNotified() As Boolean
  HaveBeenNotified = m_Notify
End Function
#End If</pre></td>
 </tr>
</table><br>
<table border="0" bgcolor="#C0C0C0">
 <tr>
  <td bgcolor="#FFFF00">Worker.cls</td>
 </tr>
 <tr>
  <td><pre>Option Explicit
Implements ThreadLaunch
Private m_Notify As Long
Public Function ThreadLaunch_Go(Controller As ThreadControl, ByVal ThreadDataCookie As Long) As Long
Dim Data As Variant
Dim SleepTime As Long
  Controller.RegisterNewThread ThreadDataCookie, VarPtr(m_Notify), Controller, Data
  ThreadLaunch_Go = Data
  SleepTime = Data
  While SleepTime &gt; 0
    Sleep 100
    SleepTime = SleepTime - 100
    If HaveBeenNotified Then
      MsgBox &quot;Notified&quot;
      Exit Function
    End If
  Wend
  MsgBox &quot;Done Sleeping: &quot; &amp; Data
End Function
Private Function HaveBeenNotified() As Boolean
  HaveBeenNotified = m_Notify
End Function</pre></td>
 </tr>
</table>
<p>&nbsp;</p>
<table border="0" bgcolor="#C0C0C0">
 <tr>
  <td bgcolor="#FFFF00">ThreadData.cls</td>
 </tr>
 <tr>
  <td><pre>Option Explicit
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
    'terminated. In this case, just claim we aren't done
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
</pre></td>
 </tr>
</table>
<p>The type library (ThreadAPI) used to call CreateThread safely is in the zip. </p>

