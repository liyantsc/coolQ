VERSION 5.00
Begin VB.Form FrmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2295
   Icon            =   "FrmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   2295
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton �Ǽ���Ϣ 
      Caption         =   "�Ǽ���Ϣ"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton �鿴��Ϣ 
      Caption         =   "�鿴��Ϣ"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton �������� 
      Caption         =   "��������"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox db.queryInfo("279135138")
End Sub

Private Sub Form_Load()
    FrmMain.Hide
End Sub

Private Sub �鿴��Ϣ_Click()
    FrmList.Show
End Sub

Private Sub �Ǽ���Ϣ_Click()
    FrmAdd.Show
    
End Sub

Private Sub ��������_Click()
    FrmGroupSet.Show
    
End Sub
