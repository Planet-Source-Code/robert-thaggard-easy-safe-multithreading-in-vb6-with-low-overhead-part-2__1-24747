VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
        'Clean up and  return
    End If
End Function
Private Function HaveBeenNotified() As Boolean
    HaveBeenNotified = m_Notify
End Function
#End If
