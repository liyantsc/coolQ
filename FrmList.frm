VERSION 5.00
Begin VB.Form FrmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "拼车信息"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8685
   Icon            =   "FrmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8685
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "全部拼车信息"
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      Begin VB.ListBox List1 
         Height          =   4200
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "刷新"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   4800
      Width           =   2175
   End
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command2_Click()
    query

End Sub

Private Sub Form_Load()
    FrmMain.Hide
    query
End Sub


Private Sub query()
    List1.Clear
    
    Dim rs As New ADODB.Recordset
    rs.Open "select * from msgdata order by type+[time]", db.getDb

    While Not rs.EOF
        List1.AddItem rs.Fields("qq") + "|" + rs.Fields("group") + "|" + rs.Fields("time") + "|" + rs.Fields("content")
        rs.MoveNext
        
    Wend
End Sub

Private Sub List1_DblClick()
    Dim content As String
    content = List1.List(List1.ListIndex)
    Dim data() As String
    Dim qq As String
    Dim group As String
    data = split(content, "|")
    qq = data(0)
    group = data(1)
    db.delMsg group, qq
    
    query
End Sub
