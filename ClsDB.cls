VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Conn As New ADODB.Connection

Const TIME_INTERVAL = 30
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Function getNextTime()
    getNextTime = Format(DateAdd("n", TIME_INTERVAL, Now), "yyyy-mm-dd hh:mm")
End Function



Public Sub openDb()
    On Error GoTo errh
    Dim path As String
    If Conn.State = 1 Then
        Exit Sub
    End If
    initPath
    path = App.path + "\app\com.liyan.pinche\msg.mdb"
    If Dir(path) <> "" Then
        Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Persist Security Info=False"
    Else

    End If
    Exit Sub
errh:
End Sub

Private Sub initPath()
    On Error Resume Next
    MkDir App.path + "\app"
    MkDir App.path + "\app\com.liyan.pinche"
End Sub
Public Sub writeLog(content As String)
    Dim fileNum As Integer
    Dim myDate As Date
    myDate = Now
    fileNum = FreeFile
    Open App.path + "\" + Format(myDate, "yyyy-mm-dd") + ".txt" For Append As #fileNum
    Print #fileNum, Format(myDate, "hh:mm:ss") + vbTab + content
    Close #fileNum
End Sub
Public Function getDb() As ADODB.Connection
    openDb
    Set getDb = Conn
End Function
'Private Sub createDb(path As String)
'    Dim MyTable As TableDef, MyField As Field
'
'    Set MyDatabase = CreateDatabase(path, dbLangGeneral)
'    Set MyTable = MyDatabase.CreateTableDef("msgdata")
'    Set MyField = MyTable.CreateField("qq", dbText, 255)
'    MyTable.Fields.Append MyField
'    Set MyField = MyTable.CreateField("mobile", dbText, 255)
'    MyTable.Fields.Append MyField
'    Set MyField = MyTable.CreateField("group", dbText, 255)
'    MyTable.Fields.Append MyField
'    Set MyField = MyTable.CreateField("type", dbText, 255)
'    MyTable.Fields.Append MyField
'    Set MyField = MyTable.CreateField("time", dbText, 255)
'    MyTable.Fields.Append MyField
'    Set MyField = MyTable.CreateField("content", dbText, 255)
'    MyTable.Fields.Append MyField
'    MyDatabase.TableDefs.Append MyTable
'
'
'    Set MyTable = MyDatabase.CreateTableDef("groupinfo")
'    Set MyField = MyTable.CreateField("groupid", dbText, 255)
'    MyTable.Fields.Append MyField
'    MyDatabase.TableDefs.Append MyTable
'End Sub

Public Function delMsg(group As String, qq As String) As Boolean
    On Error GoTo errh
    getDb.Execute "delete from msgdata where [qq]='" + qq + "' and [group]='" + group + "'"
    delMsg = True
    Exit Function
errh:
    delMsg = False
End Function

Public Function delGroupId(groupid As String) As Boolean
    On Error GoTo errh
    getDb.Execute "delete from groupinfo where [groupid]='" + groupid + "'"
    delGroupId = True
    Exit Function
errh:
    delGroupId = False
End Function

Public Function addGroupId(groupid As String) As Boolean
    On Error GoTo errh
    delGroupId groupid
    
    Dim nextTime As String
    nextTime = getNextTime
    
    getDb.Execute "insert into groupinfo ([groupid],[time]) values('" + groupid + "','" + nextTime + "')"
    addGroupId = True
    Exit Function
errh:
    addGroupId = False
End Function

Public Function insertCheMsg(group As String, qq As String, mobile As String, time As String, content As String) As Boolean
    On Error GoTo errh
    
    delMsg group, qq
    Dim sql As String
    sql = "insert into msgdata([qq],[mobile],[group],[time],[type],[content]) values('" + qq + "','" + mobile + "','" + group + "','" + time + "','0','" + content + "')"
    getDb.Execute sql
    insertCheMsg = True
    Exit Function
errh:
    insertCheMsg = False
    writeLog sql
End Function

Public Function insertRenMsg(group As String, qq As String, mobile As String, time As String, content As String) As Boolean
    On Error GoTo errh
    delMsg group, qq
    getDb.Execute "insert into msgdata([qq],[mobile],[group],[time],[type],[content]) values('" + qq + "','" + mobile + "','" + group + "','" + time + "','1','" + content + "')"
    insertRenMsg = True
    Exit Function
errh:
    insertRenMsg = False
End Function

Public Function queryInfo(groupid As String) As String
    On Error GoTo errh
    Dim rs As New ADODB.Recordset
    Dim nextTime As String
    nextTime = getNextTime
    
    rs.Open "select * from msgdata where [group]='" + groupid + "' order by type+time", getDb
    If Not (rs.EOF And rs.BOF) Then
        Dim index As Integer
        Dim content As String
        index = 0
        Dim strTime As String
        
        While Not rs.EOF
            index = index + 1

            content = content & getContent(index, rs)

            rs.MoveNext
        Wend

        Dim begin As String

        begin = "===拼车消息(" & index & "条)按发车先后==" & vbCrLf & vbCrLf
        Dim endstr As String
        endstr = "下次转发时间:" & Format(getNextTime, "hh点mm分") & ",车满请发送:车满"
        Dim allContent As String
        allContent = begin & content & endstr
        queryInfo = allContent
        
    End If
    updateNextTime groupid, nextTime
    rs.Close
    Exit Function
errh:
    writeLog Err.Description
    
    
End Function

Private Function updateNextTime(groupid As String, nextTime As String)
    On Error GoTo errh
    Conn.Execute "update groupinfo set [time]='" + nextTime + "' where groupid='" + groupid + "'"
    Exit Function
errh:
    writeLog "更新时间出错"
End Function


Public Function queryItem(group As String, qq As String) As String
    On Error GoTo errh
    Dim rs As New ADODB.Recordset
    
    Dim content As String
    rs.Open "select * from msgdata where [group]='" + group + "' and [qq]='" + qq + "'", getDb
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        content = rs.Fields("content") & ""
        rs.Close
    End If
    queryItem = content
    Exit Function
errh:
    
End Function

Public Function search(group As String, key As String) As String
    On Error GoTo errh
    Dim rs As New ADODB.Recordset
    writeLog "搜索" + key
    rs.Open "select * from msgdata where [group]='" + group + "' and content like '%" + key + "%'order by type+time", getDb
    If Not (rs.EOF And rs.BOF) Then
        Dim index As Integer
        Dim content As String
        index = 0
        Dim strTime As String
        
        While Not rs.EOF
            index = index + 1

            content = content & getContent(index, rs)

            rs.MoveNext
        Wend

        Dim begin As String
        begin = "[" + key + "]搜索结果:" + vbCrLf + vbCrLf + "===拼车消息(" & index & "条)按发车先后==" & vbCrLf & vbCrLf
        Dim allContent As String
        allContent = begin & content
        search = allContent
    End If
    rs.Close
errh:
End Function


Private Function getContent(index As Integer, rs As ADODB.Recordset) As String
    Dim strTime As String
    Dim strtype As String
    Dim emoji As String
    
    strTime = rs.Fields("time") + ""
    If Format(strTime, "yyyy-mm-dd") = Format(Now, "yyyy-mm-dd") Then
        strTime = "今" & Format(strTime, "hh:mm")
    Else
        strTime = "明" & Format(strTime, "hh:mm")
    End If
    strtype = rs.Fields("type")


    If strtype = "0" Then
        emoji = "[CQ:emoji,id=128663]"
    Else
        emoji = "[CQ:emoji,id=128107]"
    End If
    
    getContent = index & "" & emoji & strTime & "" & "[CQ:emoji,id=9742]" & rs.Fields("mobile") & vbCrLf & rs.Fields("content") & vbCrLf & vbCrLf

End Function

Public Function delTimeout() As Boolean
    On Error Resume Next
    Dim time As String
    time = Format(Now, "yyyy-mm-dd hh:mm")
    getDb.Execute "delete from msgdata where time<'" + time + "'"
End Function

Public Function autoSend() As Boolean
    On Error GoTo errh
    Dim time As String
    Dim content As String
    Dim rs As New ADODB.Recordset
    
    Dim group As String
    time = Format(Now, "yyyy-mm-dd hh:mm")
    rs.Open "select * from groupinfo where time<='" + time + "'", db.getDb
    While Not rs.EOF
        group = rs.Fields("groupid") + ""
        content = queryInfo(group)
        If content <> "" Then
            sendMesssage group, content
            Sleep 500
        Else
            sendMesssage group, "大家好，有发布拼车信息的吗?" + vbCrLf + vbCrLf + "如果您想让我帮您24小时自动转发" + vbCrLf + vbCrLf + "您可以参照以下格式" + vbCrLf + vbCrLf + "举例:@车寻人@18:00@1358176****@六里桥回*。余3座"
        End If
        rs.MoveNext
    Wend
    rs.Close
errh:
End Function

Public Function queryGroup(group As String) As Boolean
    On Error GoTo errh
    Dim rs As New ADODB.Recordset
    rs.Open "select * from groupinfo where [groupid]='" + group + "'", getDb
    If Not (rs.EOF And rs.BOF) Then
        queryGroup = True
    End If
    rs.Close
errh:
End Function

Public Function sendMesssage(group As String, content As String)
    If content <> "" Then
        Dim ZC() As Byte
        ZC = StrTByte(content)
        CQ_sendGroupMsg MainModule.ac, group / 10000, VarPtr(ZC(0))
    End If
End Function

