VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "主要的"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   3000
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   1200
      Top             =   1200
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
    db.delTimeout
End Sub

Private Sub Timer2_Timer()
    db.autoSend
    
End Sub
