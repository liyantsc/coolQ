VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "调试"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置"
      Height          =   495
      Left            =   990
      TabIndex        =   0
      Top             =   1080
      Width           =   825
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
    Dim content As String
    content = Format("2017-10-22 上午9:00", "yyyy-mm-dd hh:mm")
    MsgBox content
End Sub

Private Sub Form_Load()
    MainModule.eventStartup
End Sub


Private Function replaceWeek(content As String) As String
    Dim day As Integer
    day = Weekday(Now)
    
    content = Replace(content, "周一", fomratDateToDay(DateAdd("d", Abs(2 - day), Now)))
    content = Replace(content, "周二", fomratDateToDay(DateAdd("d", Abs(3 - day), Now)))
    content = Replace(content, "周三", fomratDateToDay(DateAdd("d", Abs(4 - day), Now)))
    content = Replace(content, "周四", fomratDateToDay(DateAdd("d", Abs(5 - day), Now)))
    content = Replace(content, "周五", fomratDateToDay(DateAdd("d", Abs(6 - day), Now)))
    content = Replace(content, "周六", fomratDateToDay(DateAdd("d", Abs(7 - day), Now)))
    content = Replace(content, "周日", fomratDateToDay(DateAdd("d", Abs(1 - day), Now)))
    
    replaceWeek = content
End Function

Private Function fomratDateToDay(content As String) As String
    fomratDateToDay = Format(content, "yyyy-mm-dd ")
End Function
