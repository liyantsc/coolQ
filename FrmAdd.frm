VERSION 5.00
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ǽ���Ϣ"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   9165
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ���� 
      Caption         =   "����"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ����_Click()
    Dim clsSplit As New clsSplit
    Dim content As String
    
    content = clsSplit.parse(Text1.Text, "94369629", "279135138")
    MsgBox content
    
    
    
End Sub
