Attribute VB_Name = "ThreadProc"
Option Explicit
Public Type ThreadProcData
    pMarshalStream As Long
    EventHandle As Long
    CLSID As CLSID
    hr As Long
    ThreadDataCookie As Long
    ThreadDonePointer As Long
    pCritSect As Long
End Type
Private Const FailBit As Long = &H80000000
Public Function ThreadStart(ThreadProcData As ThreadProcData) As Long
Dim hr As Long
Dim pUnk As IUnknown
Dim TL As ThreadLaunch
Dim TC As ThreadControl
Dim ThreadDataCookie As Long
Dim IID_IUnknown As VBGUID
Dim pMarshalStream As Long
Dim ThreadDonePointer As Long
Dim pCritSect As Long
    'Extreme care must be taken in this function to
    'not do any real VB code until an object has been
    'created on this thread by VB.
    hr = CoInitialize(0)
    With ThreadProcData
        ThreadDonePointer = .ThreadDonePointer
        If hr And FailBit Then
            .hr = hr
            PulseEvent .EventHandle
            Exit Function
        End If
        With IID_IUnknown
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        hr = CoCreateInstance(.CLSID, Nothing, CLSCTX_INPROC_SERVER, IID_IUnknown, pUnk)
        If hr And FailBit Then
            .hr = hr
            PulseEvent .EventHandle
            CoUninitialize
            Exit Function
        End If
        'If we made it this far, then we can start using normal VB calls
        'because we have an initialized object on this thread
        On Error Resume Next
        Set TL = pUnk
        Set pUnk = Nothing
        If Err Then
            .hr = Err
            PulseEvent .EventHandle
            CoUninitialize
            Exit Function
        End If
        ThreadDataCookie = .ThreadDataCookie
        pMarshalStream = .pMarshalStream
        pCritSect = .pCritSect
        'The controlling thread can continue at this point.
        'The event must be pulsed here because CoGetInterfaceAndReleaseStream
        'blocks if WaitForSingleObject is still running.
        PulseEvent .EventHandle
        Set TC = CoGetInterfaceAndReleaseStream(pMarshalStream, IID_IUnknown)
        'An error is not expected here.  If it happens, then
        'we have no way of passing it back out because the structure
        'may already be popped from the stack, meaning that we can't
        'use ThreadProcData.hr
        If Err Then
            'Note: Incrementing the ThreadDonePointer call needs
            'to be protected by a critical section once the
            'ThreadSignalPointer has been passed to ThreadControl
            'Before that time, there is no conflict.
            InterlockedIncrement ThreadDonePointer
            Set TL = Nothing
            CoUninitialize
            Exit Function
        End If
        'Launch the background thread and wait for it to finish
        'Note: TC is released by ThreadControl.RegisterNewThread
        ThreadStart = TL.Go(TC, ThreadDataCookie)
        'Tell the controlling thread that this thread is done.
        EnterCriticalSection pCritSect
        InterlockedIncrement ThreadDonePointer
        LeaveCriticalSection pCritSect
        'Release TL after the critical section. This
        'prevents ThreadData.SignalThread from
        'signalling a pointer to released memory.
        Set TL = Nothing
    End With
    CoUninitialize
End Function

