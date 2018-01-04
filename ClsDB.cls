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

Const TIME_INTERVAL = 60
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private data As New Dictionary


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

Public Function delMsgInfo(group As String, qq As String) As Boolean
    On Error GoTo errh
    getDb.Execute "delete from msginfo where [qq]='" + qq + "'"
    delMsgInfo = True
    Exit Function
errh:
    delMsgInfo = False
End Function


Public Function insertMsgInfo(group As String, qq As String, content As String) As Boolean
    On Error GoTo errh
    
    delMsgInfo group, qq
    Dim sql As String
    Dim time As String
    time = Format(Now, "yyyy-mm-dd" + " 23:59")
    sql = "insert into msginfo([qq],[group],[time],[content]) values('" + qq + "','" + group + "','" + time + "','" + content + "')"
    getDb.Execute sql
    insertMsgInfo = True
    Exit Function
errh:
    insertMsgInfo = False
    writeLog sql
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

            content = content & getContent(index, rs) + vbCrLf + vbCrLf

            rs.MoveNext
        Wend

        Dim begin As String

        begin = "===拼车消息(" & index & "条)按发车先后==" & vbCrLf & vbCrLf
        Dim endstr As String
        endstr = "下次转发时间:" & Format(getNextTime, "hh点mm分") & ",取消登记请发送 删除"
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
    If key = "" Then
        search = ""
        Exit Function
    End If
    key = Trim(key)
    rs.Open "select * from msgdata where [group]='" + group + "' and " & splitKeyWord(key) & " order by type+time", getDb
    If Not (rs.EOF And rs.BOF) Then
        Dim index As Integer
        Dim content As String
        index = 0
        Dim strTime As String
        
        While Not rs.EOF
            index = index + 1

            If index <= 5 Then
                content = content & getContent(index, rs) & vbCrLf & vbCrLf
            End If

            rs.MoveNext
        Wend

        Dim begin As String
        begin = "[" + key + "]搜索结果:" + vbCrLf + vbCrLf + "===拼车消息(" & index & "条,只显示前5条)按发车先后==" & vbCrLf & vbCrLf
        Dim allContent As String
        allContent = begin & content
        
        search = allContent & searchOther(key)
    Else
        search = searchOther(key)
    End If
    rs.Close
errh:
End Function

Private Function splitKeyWord(key As String) As String
    On Error GoTo errh
    If key = "" Then
        Exit Function
    End If
    Dim data() As String
    data = split(key, " ")
    Dim i As Integer
    Dim filter As String
    For i = 0 To UBound(data)
        If data(i) <> "" Then
            If filter <> "" Then
                filter = filter & " and "
            End If
            filter = filter & " content like '%" + data(i) + "%' "
        End If
    Next i
    splitKeyWord = filter
errh:
    
End Function

Private Function searchOther(key As String) As String
    On Error GoTo errh
    If key = "" Then
        Exit Function
    End If
    Dim rs As New ADODB.Recordset
    rs.Open "select * from msginfo where " & splitKeyWord(key), getDb
    If Not (rs.EOF And rs.BOF) Then
        Dim index As Integer
        Dim content As String
        index = 0
        Dim strTime As String
        
        While Not rs.EOF
            index = index + 1

            If index <= 5 Then
            
                content = content & index & "、" + rs.Fields("content") + vbCrLf + vbCrLf
            End If

            rs.MoveNext
        Wend

        Dim begin As String
        begin = "[" + key + "]还有以下信息可以参考:" + vbCrLf + vbCrLf + "===(" & index & "条,最多显示5条)===" & vbCrLf & vbCrLf
        Dim allContent As String
        Dim endstr As String
        endstr = "以上信息仅供参考"
        allContent = begin & content + endstr

        
        searchOther = allContent
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
    ElseIf Format(strTime, "yyyy-mm-dd") = Format(DateAdd("d", Now, 1), "yyyy-mm-dd") Then
        strTime = "明" & Format(strTime, "hh:mm")
    Else
        strTime = strTime
    End If
    strTime = Replace(strTime, "上午", "")
    strTime = Replace(strTime, "下午", "")
    strtype = rs.Fields("type")


    If strtype = "0" Then
        emoji = "[CQ:emoji,id=128663]"
    Else
        emoji = "[CQ:emoji,id=128107]"
    End If
    
    getContent = index & "" & emoji & strTime & "" & "[CQ:emoji,id=9742]" & rs.Fields("mobile") & vbCrLf & rs.Fields("content")

End Function

Public Function delTimeout() As Boolean
    On Error Resume Next
    Dim time As String
    time = Format(Now, "yyyy-mm-dd hh:mm")
    getDb.Execute "delete from msgdata where time<'" + time + "'"
    getDb.Execute "delete from msginfo where time<'" + time + "'"
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
            If isLingCheng = False Then
                addMessage group, content
            End If
        Else
            If isLingCheng = False Then
                addMessage group, "大家好，有发布拼车信息的吗?"
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
errh:
End Function

Private Function isLingCheng() As Boolean
    Dim hour As Integer
    hour = Format(Now, "hh")
    '10点以后，或者凌晨，就不发送消息了
    If hour > 22 Or (hour >= 0 And hour <= 7) Then
        isLingCheng = True
    End If
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
    content = filterContent(content)
    If content <> "" Then
        Dim ZC() As Byte
        ZC = StrTByte(content)
        CQ_sendGroupMsg MainModule.ac, group / 10000, VarPtr(ZC(0))
    End If
End Function


Public Function addMessage(group As String, content As String)
    On Error Resume Next
    Dim clsMsg As New MsgObj
    clsMsg.target = group
    clsMsg.content = content
    Randomize Timer
    data.Add group & GetTickCount, clsMsg
End Function

Public Sub waitSend()
    On Error Resume Next
    Dim msg As MsgObj
    If data.Count > 0 Then
        Set msg = data.Item(data.Keys(0))
        data.Remove data.Keys(0)
        sendMesssage msg.target, msg.content
    End If
End Sub


Public Function filterContent(content As String) As String
    On Error GoTo errh
    Dim filter1 As String
    Dim filter2 As String
    Dim i As Integer
    
    
    filter1 = "*电话*手机*预约*加群*日赚*红包群*加Q*加q*包补*诚聘*代购*放心购买*QQ群发*有意私聊*接单*任何费用*收徒*淘宝交易*招聘*要的加*低价*业务电话*转让*工资日结*现金交易*代刷*返红包*业务QQ*代价*刷钻*开钻*供货*招收*带价*免费兼职*赶快加入*如有打扰*提供货源*负责售后*火爆出售中*请+q*有需求的可以*的赚钱*寻大量*找代理*招代理*寻合作*正规开票*群里不回*网赚*可购买*欢迎加盟*推荐奖金*刷Q*兼职*点击网址*有需要做*要求能上网*急聘*免费刷*包赚*群号码*要的联系*卡盟*微信公众平台*精品课程"
    filter2 = "*中共*民主*共产党*中国共产党*胡锦涛*反共*反党*反国家*游行*集会*强奸*走私*强暴*套套*摇头丸*白粉*冰毒*海洛因*假/钞*人民币*假币*操*我日*干你*我操*操你*操你妈*干你妈*傻逼*大逼*高潮*鸡鸡*鸡巴*做爱*打炮*打洞*插入*抽插*贱人*贱逼*骚逼*牛逼*妈的*他妈的*"
    
    Dim data() As String
    data = split(filter1, "*")
    For i = 0 To UBound(data)
        content = Replace(content, data(i), "")
    Next i
    data = split(filter2, "*")
    For i = 0 To UBound(data)
        content = Replace(content, data(i), "")
    Next i
    filterContent = content
errh:
End Function

