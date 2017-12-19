VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9675
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   2280
      TabIndex        =   1
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   7455
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim clsSplit As New clsSplit
    clsSplit.parse Text1.Text, "279135138"
End Sub
