VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "����"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    FrmSetting.Show
    
End Sub

Private Sub Command2_Click()
    db.filterContent "@��Ѱ��@��������7:30@18710022429@��������7:30������ѧ԰��������������վ��ԤԼ18710022429"

End Sub

Private Sub Form_Load()
    MainModule.eventStartup
End Sub
