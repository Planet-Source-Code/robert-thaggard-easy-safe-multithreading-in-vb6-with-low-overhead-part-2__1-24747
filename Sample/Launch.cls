VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Launch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Controller As ThreadControl
Attribute Controller.VB_VarHelpID = -1
Public Sub LaunchThreads()
Dim CLSID As CLSID
    CLSID = CLSIDFromProgID("DllThreads.Worker")
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
End Sub
