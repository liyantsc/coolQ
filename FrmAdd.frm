VERSION 5.00
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登记信息"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "FrmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "登记信息"
      Height          =   4320
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4425
      Begin VB.CommandButton 增加 
         Caption         =   "增加"
         Height          =   360
         Left            =   585
         TabIndex        =   6
         Top             =   3690
         Width           =   3585
      End
      Begin VB.TextBox editContent 
         Height          =   1920
         Left            =   585
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1575
         Width           =   3660
      End
      Begin VB.TextBox group 
         Height          =   375
         Left            =   585
         TabIndex        =   4
         Top             =   990
         Width           =   3615
      End
      Begin VB.TextBox qq 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   405
         Width           =   3570
      End
      Begin VB.Label Label2 
         Caption         =   "群:"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "QQ:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub 增加_Click()
    Dim clsSplit As New clsSplit
    Dim content As String
    
    content = clsSplit.parse(editContent.Text, Trim(group.Text), Trim(qq.Text))
    If (content <> "") Then
        MsgBox content
    End If

    
End Sub
