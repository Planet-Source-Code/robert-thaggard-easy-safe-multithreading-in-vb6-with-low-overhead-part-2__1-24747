VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "CleanCompletedThreads"
      Height          =   855
      Left            =   990
      TabIndex        =   3
      Top             =   3060
      Width           =   2385
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop While Launching"
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   2130
      Width           =   2385
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Threads"
      Height          =   705
      Left            =   960
      TabIndex        =   1
      Top             =   1260
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Launch Threads"
      Height          =   765
      Left            =   960
      TabIndex        =   0
      Top             =   390
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As Object
Private Sub Command1_Click()
    Set x = CreateObject("DllThreads.Launch")
    x.LaunchThreads
End Sub

Private Sub Command2_Click()
    Set x = Nothing
End Sub

Private Sub Command3_Click()
    Set x = CreateObject("DllThreads.Launch")
    x.LaunchThreads
    x.FinishThreads
End Sub

Private Sub Command4_Click()
    If Not x Is Nothing Then
        x.CleanCompletedThreads
    End If
End Sub
