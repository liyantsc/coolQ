VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function parse(content As String, group As String, qq As String) As String


    content = Trim(content)
    content = Replace(content, vbTab, "") '
    content = Replace(content, "'", "")
    content = Replace(content, vbNullChar, "")
    
    content = db.filterContent(content)
    
    If content = "" Then
        Exit Function
    End If
    
    If InStr(content, "找人") > 0 Or InStr(content, "车寻人") Or InStr(content, "私家车") Then
        db.insertMsgInfo group, qq, content
    End If
    
    If Not (db.queryGroup(group)) Then
        Exit Function
    End If
    
    
    On Error GoTo errh:
    Dim data() As String
    Dim time As String
    Dim mobile As String
    Dim dateTime As Date
    
    Dim keyword As String

    If content = "拼车信息" Or content = "拼车" Then
        parse = db.queryInfo(group)
        If parse = "" Then
            parse = getNoDataMsg(qq)
        Else
            db.writeLog parse
        End If
        Exit Function
    End If
    If content = "京昆" Or UCase(content) = "G5" Or content = "京昆高速" Then
        FrmMain.queryLuKuang "http://gsgl.00cha.com/G5.html", qq, group
        Exit Function
    End If
    If content = "京港澳" Or UCase(content) = "G4" Or content = "京港澳高速" Then
        FrmMain.queryLuKuang "http://gsgl.00cha.com/G4.html", qq, group
        Exit Function
    End If
    If content = "京哈" Or UCase(content) = "G1" Or content = "京哈高速" Then
        FrmMain.queryLuKuang "http://gsgl.00cha.com/G1.html", qq, group
        Exit Function
    End If
    If content = "京沪" Or UCase(content) = "G2" Or content = "京沪高速" Then
        FrmMain.queryLuKuang "http://gsgl.00cha.com/G2.html", qq, group
        Exit Function
    End If
    If content = "京台" Or UCase(content) = "G3" Or content = "京台高速" Then
        FrmMain.queryLuKuang "http://gsgl.00cha.com/G3.html", qq, group
        Exit Function
    End If
    If content = "京藏" Or UCase(content) = "G6" Or content = "京藏高速" Then
        FrmMain.queryLuKuang "http://gsgl.00cha.com/G6.html", qq, group
        Exit Function
    End If
    If content = "京新" Or UCase(content) = "G7" Or content = "京新高速" Then
        FrmMain.queryLuKuang "http://gsgl.00cha.com/G7.html", qq, group
        Exit Function
    End If
    If content = "查限行" Or content = "查限号" Or content = "限号查询" Or content = "查询限号" Or content = "限号" Or content = "限行" Then
        FrmMain.queryLuKuang XIANXING_URL, qq, group
        Exit Function
    End If
    If Left(content, 3) = "查路况" Then
        keyword = Mid(content, 4)
        parse = queryLuKuang(qq, group, keyword)
'        FrmMain.searchLuKuang LUKUANG_URL, qq, group, keyword
        Exit Function
    End If
    If Left(content, 1) = "搜" Then
        keyword = Mid(content, 2)
        If keyword <> "" Then
            parse = db.search(group, keyword)
            If parse = "" Then
                parse = getNoDataMsg(qq)
            Else
                parse = "[CQ:at,qq=" & qq & "]" + vbCrLf + vbCrLf + parse
            End If
        End If
        Exit Function
    End If
    If content = "车满" Or content = "删除" Or content = "#删除" Or content = "#车满" Then
        parse = db.queryItem(group, qq)
        If parse <> "" Then
            db.delMsg group, qq
            parse = "[CQ:at,qq=" & qq & "]" + vbCrLf + "删除成功，您的信息将不再转发"
        Else
            parse = "[CQ:at,qq=" & qq & "]" + vbCrLf + "虽然我是机器人，但也不能这么调戏我"
        End If
        Exit Function
    End If
    
    parse = parseBySplitChar(content, group, qq, "@")
    If parse = "" Then
        parse = parseBySplitChar(content, group, qq, vbCrLf)
    End If
errh:
    If parse = "" Then
        If InStr(content, "找人") > 0 Or InStr(content, "寻人") > 0 Then
            parse = getCheErrMsg(qq, time, mobile)
        ElseIf InStr(content, "找车") > 0 Or InStr(content, "寻车") > 0 Then
            parse = getPersonErrMsg(qq, time, mobile)
        End If
    End If
End Function

Private Function parseBySplitChar(content As String, group As String, qq As String, splitChar As String) As String
    On Error GoTo errh
    Dim data() As String
    Dim time As String
    Dim mobile As String
    Dim dateTime As Date
    If InStr(content, "车找人") > 0 Or InStr(content, "车寻人") > 0 Then
        data = split(content, splitChar)
        If UBound(data) < 2 Then
            parseBySplitChar = ""
            Exit Function
        End If
    End If
    If InStr(content, "人寻车") > 0 Or InStr(content, "人找车") > 0 Or InStr(content, "找车") > 0 Then
        data = split(content, splitChar)
        If UBound(data) < 2 Then
            parseBySplitChar = ""
            Exit Function
        End If
    End If
    data = split(content, splitChar)
    If UBound(data) = 0 Then
        Exit Function
    End If
    Dim command As String
    Dim i As Integer
    Dim startIndex As Integer
    If splitChar = "@" Then
        startIndex = 1
    Else
        startIndex = 0
    End If
    command = Trim(data(startIndex))
    command = Replace(command, vbCr, "")
    command = Replace(command, vbLf, "")
    command = Replace(command, vbCrLf, "")
    time = getTime(data(startIndex + 1))
    mobile = getUserPhone(content)
    content = ""
    For i = startIndex + 2 To UBound(data)
        content = content & data(i)
    Next i
    content = Replace(content, splitChar + mobile, "")
    content = Replace(content, mobile, "")
    If command = "车寻人" Or command = "车找人" Or command = "私家车找人" Then
        If time <> "" And mobile <> "" And content <> "" Then
            db.insertCheMsg group, qq, mobile, time, content
            
            parseBySplitChar = "[CQ:at,qq=" & qq & "]" + "[CQ:face,id=76][CQ:face,id=76][CQ:face,id=76]" + vbCrLf + "登记成功 发车时间:" & time + vbCrLf + "取消登记请发送 删除"
'        Else
'            parseBySplitChar = getCheErrMsg(qq, time, mobile)
        End If
        Exit Function
    End If
    If command = "人寻车" Or command = "人找车" And content <> "" Then
        If time <> "" And mobile <> "" Then
            db.insertRenMsg group, qq, mobile, time, content
            parseBySplitChar = "[CQ:at,qq=" & qq & "]" + "[CQ:face,id=76][CQ:face,id=76][CQ:face,id=76]" + vbCrLf + "登记成功 坐车时间:" & time + vbCrLf + "取消登记请发送 删除"
'        Else
'            parseBySplitChar = getPersonErrMsg(qq, time, mobile)
        End If
        Exit Function
    End If
errh:
    
End Function


Private Function getPersonErrMsg(qq As String, time As String, mobile As String) As String
    Dim modStr As String
    modStr = "您可以参考以下模板重新发送" + vbCrLf + vbCrLf + "@人寻车@18:00@13912345678@出发地点到目的地"
    If time = "" Then
        getPersonErrMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + "您的格式有误，机器人不会自动转发" + vbCrLf + modStr
        Exit Function
    End If
    If mobile = "" Then
        getPersonErrMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + "您的格式有误，机器人不会自动转发" + vbCrLf + modStr
        Exit Function
    End If
    getPersonErrMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + "您的格式有误，机器人不会自动转发" + vbCrLf + modStr
End Function

Private Function getCheErrMsg(qq As String, time As String, mobile As String) As String
    Dim modStr As String
    modStr = "您可以参考以下模板重新发送" + vbCrLf + vbCrLf + "@车寻人@18:00@13912345678@出发地点到目的地"
    If time = "" Then
        getCheErrMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + "您的格式有误，机器人不会自动转发" + vbCrLf + modStr
        Exit Function
    End If
    If mobile = "" Then
        getCheErrMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + "您的格式有误，机器人不会自动转发" + vbCrLf + modStr
        Exit Function
    End If
    getCheErrMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + "您的格式有误，机器人不会自动转发" + vbCrLf + modStr
End Function

Private Function getNoDataMsg(qq As String) As String
    getNoDataMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + "非常抱歉,暂时没有相关信息"
End Function

Public Function queryLuKuang(qq As String, group As String, keyword As String) As String
    On Error GoTo errh
    Dim xmlHttp
    Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
    Randomize Timer
    Dim temp As Long
    temp = Rnd() * 9999999
    xmlHttp.Open "GET", "http://eye.bjjtw.gov.cn/Web-T_bjjt_new/query.do?serviceType=jam&acode=110000&type=3&cls=0&rcd=40&t=" & temp, True
    xmlHttp.send (Null)
    While xmlHttp.ReadyState <> 4
        DoEvents
    Wend
    Dim body As String
    body = xmlHttp.responseText
    If body = "" Then
        Exit Function
    End If

    Dim content As String
    Dim bFind As Boolean
    
    Dim dict As New Dictionary
    
    
    Dim ScriptControl As Object, roads As Object, JscriptCode$
    JscriptCode = "function toObject(json) {eval(""var o=""+json);return o;}"
    Set ScriptControl = CreateObject("MSScriptControl.ScriptControl")
    With ScriptControl
        .Language = "Javascript"
        .Timeout = -1
        .AddCode JscriptCode
        Set roads = .Run("toObject", body)
        Dim item
        Dim time As String
        Dim txtInfo
        Dim roaditem
        Dim roadsub
        Dim itemInfo As String
        Dim findMain As Boolean
        Dim findS As Boolean
        Dim findE As Boolean
        
        For Each item In roads.roads
            Set txtInfo = item.txtInfo
            findMain = findKeyWord(item.road, keyword)
            For Each roaditem In txtInfo
                For Each roadsub In roaditem.section
                    findS = findKeyWord(roadsub.sName, keyword)
                    findE = findKeyWord(roadsub.eName, keyword)

                    If findMain = True Or findS = True Or findE = True Then
                        itemInfo = item.road & " " & item.time
                        
                        itemInfo = itemInfo + vbCrLf + getDirection(roaditem.direction) + vbCrLf
                        itemInfo = itemInfo & roadsub.sName & "-→" & roadsub.eName & " " & roadsub.info & vbCrLf
                        If dict.Exists(itemInfo) = False Then
                            dict.Add itemInfo, itemInfo
                        End If
                    End If
                Next
            Next

        Next
        Dim s
        For Each s In dict
            If s <> "" Then
                If content <> "" Then
                    content = content & "- - - - - - - - - - - -" & vbCrLf
                End If
                content = content & s
            End If
        Next
    End With
    If content <> "" Then
        queryLuKuang = "[CQ:at,qq=" & qq & "]" & vbCrLf & content
    Else
        queryLuKuang = "[CQ:at,qq=" & qq & "]" & vbCrLf & "没有查到相关路况"
    End If
errh:
End Function

Private Function findKeyWord(value As String, keyword As String) As Boolean
    On Error GoTo errh
    Dim bFind As Boolean
    Dim data() As String
    Dim i As Integer
    bFind = False
    If Trim(keyword) <> "" Then  '如果不是查全部
         data = split(keyword, " ")
         bFind = False
         For i = 0 To UBound(data)
             If data(i) <> "" Then
                 If InStr(value, data(i)) Then
                     bFind = True
                     Exit For
                 End If
             End If
         Next i
     Else
         bFind = True
     End If
     findKeyWord = bFind
errh:
End Function
Private Function getDirection(direction As String)
    On Error Resume Next
    Select Case UCase(direction)
        Case "NS"
            getDirection = "北向南方向"
        Case "WE"
            getDirection = "西向东方向"
        Case "SN"
            getDirection = "南向北方向"
        Case "EW"
            getDirection = "东向西方向"
    End Select
End Function

Public Function getUserPhone(ByVal TempStr As String) As String
    On Error GoTo errh
    Dim RegGex As New RegExp
    Dim ms As MatchCollection
    Dim m As Match
    Dim i As Integer
    Dim temp As String
    Dim Result As String
    TempStr = Replace(TempStr, vbCr, "")
    TempStr = Replace(TempStr, vbLf, "")
    TempStr = Replace(TempStr, vbCrLf, "")
    RegGex.IgnoreCase = True
    RegGex.Global = True
    RegGex.Pattern = "13[0123456789]{1}\d{8}|15[0123456789]\d{8}|18[0123456789]{1}\d{8}|17[0123456789]{1}\d{8}|17[0123456789]{1}\d{8}"
    Set ms = RegGex.Execute(TempStr)
    i = 0
    For Each m In ms
        DoEvents
        temp = m.value
        Exit For
    Next
    getUserPhone = temp
errh:
End Function

Public Function parseLuKuang(Document As Object, qq As String, group As String, url As Variant) As String

    On Error GoTo errh
    Dim body
    If IsEmpty(Document) Then
        Exit Function
    End If
    Dim content As String
    Dim count As Integer
    
    For Each body In Document.getElementsByTagName("div")
        If body.classname = "sj" Then
            If content <> "" Then
                content = content + vbCrLf + vbCrLf
            End If
            content = content + body.outertext
            count = count + 1
            If count >= 8 Then
                Exit For
            End If
        End If
    Next
    If content <> "" Then
        parseLuKuang = "[CQ:at,qq=" & qq & "]" & vbCrLf & content
    End If
errh:
    
End Function
Public Function parseXianXing(Document As Object, qq As String, group As String, url As Variant) As String
    On Error GoTo errh
    If url <> XIANXING_URL Then
        Exit Function
    End If
    Dim body
    If IsEmpty(Document) Then
        Exit Function
    End If
    Dim content As String
    Dim count As Integer
    
    For Each body In Document.getElementsByTagName("div")
        If body.classname = "xianhao" Then
            If content <> "" Then
                content = content + vbCrLf + vbCrLf
            End If
            content = content + body.outertext
        End If
    Next
    If content <> "" Then
        content = Replace(content, "  详情>>", "")
        parseXianXing = "[CQ:at,qq=" & qq & "]" & vbCrLf & "限行时间：凌晨7:00至晚上20:00；限行范围：五环路(含)以内道路；车牌尾号为字母：按0号管理" + vbCrLf + content
    End If
errh:
    
End Function
Private Function getTime(content As String) As String
    On Error GoTo errh:
    content = Trim(content)
    content = Replace(content, vbCr, "")
    content = Replace(content, vbLf, "")
    content = Replace(content, vbCrLf, "")
    If (InStr(content, ":") > 0 Or InStr(content, ".") > 0 Or InStr(content, "：") > 0) Or InStr(content, "点") > 0 Then
        Dim time As String
        Dim isTommorrow As Boolean
        Dim isTommorrowTommorrow As Boolean
        
        If InStr(content, "明") Then
            isTommorrow = True
        End If
        If InStr(content, "后天") Then
            isTommorrowTommorrow = True
        End If
        time = replaceWeek(content)

        If InStr(time, "-") Then
            time = Format(time, "yyyy-mm-dd hh:mm")
            Dim data() As String
            data = split(time, " ")
            time = data(0) & " " + getHourMinuteTime(data(1))
        Else
            time = Format(Format(Now, "yyyy-mm-dd") + " " + getHourMinuteTime(time), "yyyy-mm-dd hh:mm")
        End If
        If isTommorrow Then
            time = DateAdd("d", 1, time)
        End If
        If isTommorrowTommorrow Then
            time = DateAdd("d", 2, time)
        End If
        time = Format(time, "yyyy-mm-dd hh:mm")
        time = getHourMinuteTime(time)

    End If
    Dim strDate As Date
    strDate = CDate(time)
    getTime = Format(strDate, "yyyy-mm-dd hh:mm")
    Exit Function
errh:
    getTime = ""
End Function

Private Function getHourMinuteTime(time As String) As String
    On Error GoTo errh:
    time = Replace(time, "一", "1")
    time = Replace(time, "二", "2")
    
    time = Replace(time, "三", "3")
    
    time = Replace(time, "四", "4")
    time = Replace(time, "五", "5")
    time = Replace(time, "六", "6")
    time = Replace(time, "七", "7")
    time = Replace(time, "八", "8")
    time = Replace(time, "九", "9")
    time = Replace(time, "十", "0")
    time = Replace(time, ".", ":")
    time = Replace(time, "：", ":")
    time = Replace(time, "点", ":")
    time = Replace(time, "半", "30")
    time = Replace(time, "今天", "上午")
    time = Replace(time, "今晚", "下午")
    time = Replace(time, "晚上", "下午")
    time = Replace(time, "明天", "")
    time = Replace(time, "明晚", "下午")
    time = Replace(time, "后天", "")
    time = Replace(time, "明早", "上午")
    time = Replace(time, "左右", "")
    time = Replace(time, "晚", "下午")
    time = Replace(time, "明", "下午")
    time = Replace(time, "今", "下午")
    time = Replace(time, "早晨", "上午")
    time = Replace(time, "早上", "上午")
    time = Replace(time, "中午", "下午")
    time = Replace(time, "本", "")
    time = Replace(time, "分", "")
    time = Replace(time, "早", "上午")
    time = Replace(time, "下", "下午")
    time = Replace(time, "下午午", "下午")

    
    Dim data() As String
    data = split(time, ":")
    If UBound(data) <> 1 Then
        Exit Function
    End If
    
    If data(1) = "" Then
        time = time + "00"
    End If
    getHourMinuteTime = time
errh:
    
End Function

Private Function replaceWeek(content As String) As String
    Dim day As Integer
    day = Weekday(Now)
    content = Replace(content, "周一", fomratDateToDay(DateAdd("d", Abs(2 - day), Now)))
    content = Replace(content, "周二", fomratDateToDay(DateAdd("d", Abs(3 - day), Now)))
    content = Replace(content, "周三", fomratDateToDay(DateAdd("d", Abs(4 - day), Now)))
    content = Replace(content, "周四", fomratDateToDay(DateAdd("d", Abs(5 - day), Now)))
    content = Replace(content, "周五", fomratDateToDay(DateAdd("d", Abs(6 - day), Now)))
    content = Replace(content, "周六", fomratDateToDay(DateAdd("d", Abs(7 - day), Now)))
    content = Replace(content, "周日", fomratDateToDay(DateAdd("d", Abs(8 - day), Now)))
    
    content = Replace(content, "星期一", fomratDateToDay(DateAdd("d", Abs(2 - day), Now)))
    content = Replace(content, "星期二", fomratDateToDay(DateAdd("d", Abs(3 - day), Now)))
    content = Replace(content, "星期三", fomratDateToDay(DateAdd("d", Abs(4 - day), Now)))
    content = Replace(content, "星期四", fomratDateToDay(DateAdd("d", Abs(5 - day), Now)))
    content = Replace(content, "星期五", fomratDateToDay(DateAdd("d", Abs(6 - day), Now)))
    content = Replace(content, "星期六", fomratDateToDay(DateAdd("d", Abs(7 - day), Now)))
    content = Replace(content, "星期日", fomratDateToDay(DateAdd("d", Abs(8 - day), Now)))
    
    replaceWeek = content
End Function

Private Function fomratDateToDay(content As String) As String
    fomratDateToDay = Format(content, "yyyy-mm-dd ")
End Function


'Private Function getTodayTime(ByVal time As String) As String
'    On Error GoTo errh
'    Dim temp As String
'    temp = time
'    time = Replace(time, "今天", "上午")
'    time = Replace(time, "今晚", "下午")
'    time = Replace(time, "晚上", "下午")
'    time = Replace(time, "晚", "下午")
'    time = Replace(time, "早晨", "上午")
'    time = Replace(time, "早上", "上午")
'
'    time = Replace(time, "左右", "")
'    time = Format(time, "hh:mm")
'    Dim dateTime As Date
'    dateTime = CDate(time)
'    getTodayTime = Format(Date, "yyyy-mm-dd hh:mm")
'    Exit Function
'errh:
'    getTodayTime = ""
'End Function
'Private Function getTommorrowTime(ByVal time As String) As String
'    On Error GoTo errh
'    Dim temp As String
'    temp = time
'    time = Replace(time, "明天", "")
'    time = Replace(time, "明晚", "下午")
'    time = Replace(time, "晚上", "下午")
'    time = Replace(time, "早晨", "上午")
'    time = Replace(time, "早上", "上午")
'    time = Replace(time, "左右", "")
'    Dim dateTime As Date
'    dateTime = CDate(time)
'    getTodayTime = Format(Date, "yyyy-mm-dd hh:mm")
'    getTommorrowTime = time
'    Exit Function
'errh:
'    getTommorrowTime = ""
'End Function
'
'Private Function insertData(name As String, content As String) As Boolean
'    Dim rs As New ADODB.Recordset
'    rs.Open "select id from msgdata where name='" & name & "' and content='" + content + "'", conn, adOpenStatic, adLockReadOnly
'    If Not (rs.EOF And rs.BOF) Then
'
'    Else
'        conn.Execute "insert into msgdata (name, content) values ('" & name & "','" & content & "')"
'    End If
'End Function

  '加密
'Public Function Encrypt(ByVal strSource As String, ByVal key As Byte) As String
'    Dim i     As Long
'    Dim j     As Byte
'    Dim temps     As String
'    Dim s     As String
'    Dim arr()     As Byte
'    arr = StrConv(strSource, vbFromUnicode)
'    For i = 0 To UBound(arr)
'    j = arr(i) Xor key
'    temps = Right("00" & Hex(j), 2)
'    s = s + temps
'    Next
'    Encrypt = s
'End Function
''解密
'Public Function decrypt(ByVal strSource As String, ByVal key As Byte) As String
'    Dim i     As Long
'    Dim j     As Long
'    Dim temps     As String
'    Dim s     As String
'    Dim arr     As Variant
'    i = Len(strSource)
'    If i Mod 2 = 1 Then
'        '待解密的字串不符合要求
'        decrypt = ""
'        Exit Function
'    End If
'    Dim buff()     As Byte
'    Dim k     As Long
'    k = 0
'    For i = 1 To Len(strSource) Step 2
'        temps = Mid(strSource, i, 2)
'        j = Val("&H" & temps)
'        j = j Xor key
'        ReDim Preserve buff(k)
'        buff(k) = j
'        k = k + 1
'    Next
'    decrypt = StrConv(buff, vbUnicode)
'End Function
'Public Function GetContentString() As String
'
'   Dim rs As New ADODB.Recordset
'
'    rs.Open "select * from msg order by sendtime asc", conn, adOpenStatic, adLockReadOnly
'
'    Dim content As String
'
'    content = "信息由机器自动统计" + vbCrLf
'
'    Dim index As Integer
'
'    index = 1
'
'    If Not (rs.EOF And rs.BOF) Then
'        While Not rs.EOF
'            content = content + CStr(index) + "、" + "<高铁>" + rs.Fields("sendtime") + "  电话：" + rs.Fields("username") + "  " + rs.Fields("content") + vbCrLf + "                              " + vbCrLf
'
'            rs.MoveNext
'            index = index + 1
'        Wend
'    End If
'    GetContentString = content
'End Function
