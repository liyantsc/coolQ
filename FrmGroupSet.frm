VERSION 5.00
Begin VB.Form FrmGroupSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ⱥ��������"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6480
   Icon            =   "FrmGroupSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6480
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "�����Ⱥ��"
      Height          =   4575
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton ˢ�� 
         Caption         =   "ˢ��"
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton ���� 
         Caption         =   "����"
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Ⱥ�ţ�"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Ѿ��ӹܵ�Ⱥ"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ListBox List1 
         Height          =   4200
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "˫��ɾ��"
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmGroupSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub query()
    List1.Clear
    
    Dim rs As New ADODB.Recordset
    rs.Open "select * from groupinfo", db.getDb
    While Not rs.EOF
        List1.AddItem rs.Fields("groupid")
        rs.MoveNext
    Wend
End Sub

Private Sub List1_DblClick()
    Dim groupid As String

    groupid = List1.List(List1.ListIndex)

    db.delGroupId groupid
    
    query
End Sub

Private Sub ˢ��_Click()
    query
End Sub

Private Sub ����_Click()
    Dim groupid As String
    groupid = Text1.Text
    If groupid = "" Then Exit Sub
    
    db.addGroupId groupid
    
    query
End Sub
