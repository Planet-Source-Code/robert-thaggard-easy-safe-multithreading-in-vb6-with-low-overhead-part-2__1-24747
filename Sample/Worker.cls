VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Worker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ThreadLaunch
Private m_Notify As Long

Public Function ThreadLaunch_Go(Controller As ThreadControl, ByVal ThreadDataCookie As Long) As Long
Dim Data As Variant
Dim SleepTime As Long
    Controller.RegisterNewThread ThreadDataCookie, VarPtr(m_Notify), Controller, Data
    ThreadLaunch_Go = Data
    SleepTime = Data
    While SleepTime > 0
        Sleep 100
        SleepTime = SleepTime - 100
        If HaveBeenNotified Then
            MsgBox "Notified"
            Exit Function
        End If
    Wend
    MsgBox "Done Sleeping: " & Data
End Function
Private Function HaveBeenNotified() As Boolean
    HaveBeenNotified = m_Notify
End Function
