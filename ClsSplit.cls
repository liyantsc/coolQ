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
    content = Replace(content, vbCr, "")
    content = Replace(content, vbLf, "")
    content = Replace(content, vbTab, "") '
    content = Replace(content, "'", "")
    content = Replace(content, vbNullChar, "")
    content = Replace(content, vbCrLf, "")
    
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
    


    If content = "拼车信息" Or content = "拼车" Then
        parse = db.queryInfo(group)
        If parse = "" Then
            parse = getNoDataMsg(qq)
        Else
            db.writeLog parse
        End If
        Exit Function
    End If
    If Left(content, 1) = "搜" Then
        Dim keyword As String
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
    If content = "车满" Or content = "删除" Then
        parse = db.queryItem(group, qq)
        If parse <> "" Then
            db.delMsg group, qq
            parse = "[CQ:at,qq=" & qq & "]" + vbCrLf + "拼车信息:" + parse + vbCrLf + vbCrLf + "删除成功"
        End If
        Exit Function
    End If
    
    If InStr(content, "车找人") > 0 Or InStr(content, "车寻人") > 0 Or InStr(content, "人寻车") > 0 Or InStr(content, "人找车") > 0 Then
        data = split(content, "@")
        If UBound(data) < 2 Then
            parse = getErrMsg(qq)
            Exit Function
        End If
    End If
    data = split(content, "@")
    If UBound(data) = 0 Then
        Exit Function
    End If
    Dim command As String
    command = Trim(data(1))
    If command = "车寻人" Or command = "车找人" Then
        time = getTime(data(2))
        mobile = getUserPhone(content)
        If time <> "" And mobile <> "" Then
            db.insertCheMsg group, qq, mobile, time, content
            
            parse = "[CQ:at,qq=" & qq & "] 登记成功" & vbCrLf & "发车时间:" & time
        Else
            parse = getErrMsg(qq)
        End If
        Exit Function
    End If
    If command = "人寻车" Or command = "人找车" Then
        time = getTime(data(2))
        mobile = getUserPhone(content)
        If time <> "" And mobile <> "" Then
            db.insertRenMsg group, qq, mobile, time, content
            parse = "[CQ:at,qq=" & qq & "] 登记成功" & vbCrLf & "坐车时间:" & time
        Else
            parse = getErrMsg(qq)
        End If
        Exit Function
    End If
errh:
    
End Function

Private Function getErrMsg(qq As String) As String
    getErrMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + vbCrLf + "您的信息我无法识别" + vbCrLf + "您可以参照以下格式" + vbCrLf + "举例:@车寻人@18:00@1761030****@六里桥回*。余3座"
End Function

Private Function getNoDataMsg(qq As String) As String
    getNoDataMsg = "[CQ:at,qq=" & qq & "]" + vbCrLf + vbCrLf + "非常抱歉,暂时没有相关信息"
End Function

Public Function getUserPhone(ByVal TempStr As String) As String
    On Error GoTo errh
    Dim RegGex As New RegExp
    Dim ms As MatchCollection
    Dim m As Match
    Dim i As Integer
    Dim temp As String
    Dim Result As String
    RegGex.IgnoreCase = True
    RegGex.Global = True
    RegGex.Pattern = "13[0123456789]{1}\d{8}|15[0123456789]\d{8}|18[0123456789]{1}\d{8}|17[0123456789]{1}\d{8}"
    Set ms = RegGex.Execute(TempStr)
    i = 0
    For Each m In ms
        DoEvents
        temp = m.Value
        Exit For
    Next
    getUserPhone = temp
errh:
End Function
Private Function getTime(content As String) As String
    On Error GoTo errh:
    content = Trim(content)
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
        
        time = Replace(time, "一", "1")
        time = Replace(time, "二", "2")
        
        time = Replace(time, "三", "3")
        
        time = Replace(time, "四", "4")
        time = Replace(time, "五", "5")
        time = Replace(time, "六", "6")
        time = Replace(time, "七", "7")
        time = Replace(time, "八", "8")
        time = Replace(time, "九", "9")
        time = Replace(time, ".", ":")
        time = Replace(time, "：", ":")
        time = Replace(time, "点", ":")
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
        

        
        Dim data() As String
        data = split(time, ":")
        If UBound(data) <> 1 Then
            getTime = ""
            Exit Function
        End If
        
        If data(1) = "" Then
            time = time + "00"
        End If
        If InStr(time, "-") Then
            time = Format(time, "yyyy-mm-dd hh:mm")
        Else
            time = Format(Format(Now, "yyyy-mm-dd") + " " + time, "yyyy-mm-dd hh:mm")
        End If
        If isTommorrow Then
            time = DateAdd("d", 1, time)
        End If
        If isTommorrowTommorrow Then
            time = DateAdd("d", 2, time)
        End If
        time = Format(time, "yyyy-mm-dd hh:mm")
    End If
    getTime = time
    Exit Function
errh:
    getTime = ""
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
    content = Replace(content, "周日", fomratDateToDay(DateAdd("d", Abs(1 - day), Now)))
    
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
