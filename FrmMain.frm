VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmMain 
   Caption         =   "主要的"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   6300
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   1320
      Left            =   1665
      TabIndex        =   0
      Top             =   3330
      Width           =   1455
      ExtentX         =   2566
      ExtentY         =   2328
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer3 
      Interval        =   999
      Left            =   1215
      Top             =   2565
   End
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

Private strQQ As String
Private strGroup As String

Private Sub Form_Load()
    web.Silent = True
End Sub

Private Sub Timer1_Timer()
    db.delTimeout
End Sub

Private Sub Timer2_Timer()
    Timer2.Interval = 0
    db.autoSend
    Timer2.Interval = 5000
End Sub

Private Sub Timer3_Timer()
    db.waitSend
    
End Sub

Public Sub queryLuKuang(url As String, qq As String, group As String)
    strQQ = qq
    strGroup = group
    web.Silent = True
    web.Navigate url
End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, url As Variant)
    On Error GoTo errh
    Dim splitTool As New clsSplit
    Dim content As String
    content = splitTool.parseLuKuang(web.Document, strQQ, strGroup)
    If content <> "" Then
        Debug.Print content
        db.addMessage strGroup, content
    End If
errh:
End Sub
