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
    If content = "" Then
        Exit Function
    End If
    
    If db.queryGroup(group) Then
        db.writeLog "�������йܵ�Ⱥ"
        Exit Function
    End If
    
    
    On Error GoTo errh:
    Dim data() As String
    Dim time As String
    Dim mobile As String
    Dim dateTime As Date
    
    db.writeLog content

    If content = "ƴ����Ϣ" Then
        parse = db.queryInfo(group)
        If parse = "" Then
            parse = "[CQ:at,qq=" & qq & "]" + vbCrLf + vbCrLf + "�ǳ���Ǹ����ʱ��û���˷���ƴ����Ϣ"
        End If
        Exit Function
    End If
    If Left(content, 2) = "�� " Then
        data = split(content, " ")
        If UBound(data) = 1 Then
            parse = db.search(group, data(1))
            If parse = "" Then
                parse = "[CQ:at,qq=" & qq & "]" + vbCrLf + vbCrLf + "�ǳ���Ǹ��û���ҵ����ƴ����Ϣ"
            Else
                parse = "[CQ:at,qq=" & qq & "]" + vbCrLf + vbCrLf + parse
            End If
        End If

        Exit Function
    End If
    If content = "����" Or content = "ɾ��" Then
        parse = db.queryItem(group, qq)
        If parse <> "" Then
            db.delMsg group, qq
            parse = "[CQ:at,qq=" & qq & "]" + vbCrLf + "ƴ����Ϣ:" + parse + vbCrLf + vbCrLf + "ɾ���ɹ�,����ת��" + vbCrLf + vbCrLf + db.queryInfo(group)
        End If
        Exit Function
    End If
    
    If InStr(content, "������") > 0 Or InStr(content, "��Ѱ��") > 0 Or InStr(content, "��Ѱ��") > 0 Or InStr(content, "���ҳ�") > 0 Then
        data = split(content, "@")
        If UBound(data) < 2 Then
            parse = "[CQ:at,qq=" & qq & "]��Ⱥ��ƴ����Ϣ��ʽҪ�����ĸ�ʽ���󣬻����˲����Զ�ת��" + vbCrLf + vbCrLf + "�����԰��˸�ʽ���·�:@��Ѱ��@ʱ��@�绰@����" + vbCrLf + vbCrLf + "����:@��Ѱ��@18:00@1358176****@�����Ż����ء���3��"
            Exit Function
        End If
    End If
    data = split(content, "@")
    If UBound(data) = 0 Then
        Exit Function
    End If
    
    If data(1) = "��Ѱ��" Or data(1) = "������" Then
        time = getTime(data(2))
        mobile = getUserPhone(content)
        If time <> "" And mobile <> "" Then
            db.insertCheMsg group, qq, mobile, time, content
            
            parse = "[CQ:at,qq=" & qq & "] �Ǽǳɹ�" & vbCrLf & "����ʱ��Ǽ�Ϊ" & time
        Else
            parse = "[CQ:at,qq=" & qq & "]��Ⱥ��ƴ����Ϣ��ʽҪ�����ĸ�ʽ���󣬻����˲����Զ�ת��" + vbCrLf + vbCrLf + "�����԰��˸�ʽ���·�:@��Ѱ��@ʱ��@�绰@����" + vbCrLf + vbCrLf + "����:@��Ѱ��@18:00@1358176****@�����Ż����ء���3��"
        End If
        Exit Function
    End If
    If data(1) = "��Ѱ��" Or data(1) = "���ҳ�" Then
        time = getTime(data(2))
        mobile = getUserPhone(content)
        If time <> "" And mobile <> "" Then
            db.insertRenMsg group, qq, mobile, time, content
            parse = "�Ǽǳɹ�"
        Else
            parse = "[CQ:at,qq=" & qq & "]��Ⱥ��ƴ����Ϣ��ʽҪ�����ĸ�ʽ���󣬻����˲����Զ�ת��" + vbCrLf + vbCrLf + "�����԰��˸�ʽ���·�:@��Ѱ��@ʱ��@�绰@����" + vbCrLf + vbCrLf + "����:@��Ѱ��@18:00@1358176****@�����Ż����ء���3��"
        End If
        Exit Function
    End If
errh:
    
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
    If (InStr(content, ":") > 0 Or InStr(content, ".") > 0 Or InStr(content, "��") > 0) Or InStr(content, "��") > 0 Then
        Dim time As String
        Dim isTommorrow As Boolean

        
        If InStr(content, "��") Then
            isTommorrow = True
        End If
        time = content
        time = Replace(time, ".", ":")
        time = Replace(time, "��", ":")
        time = Replace(time, "��", ":")
        time = Replace(time, "����", "����")
        time = Replace(time, "����", "����")
        time = Replace(time, "����", "����")
        time = Replace(time, "��", "����")
        time = Replace(time, "�糿", "����")
        time = Replace(time, "����", "����")
        
        time = Replace(time, "����", "")
        time = Replace(time, "����", "����")
        time = Replace(time, "����", "����")
        time = Replace(time, "����", "")
        
        Dim data() As String
        data = split(time, ":")
        If UBound(data) <> 1 Then
            getTime = ""
            Exit Function
        End If
        
        If data(1) = "" Then
            time = time + "00"
        End If
        
        
        time = Format(time, "hh:mm")
        time = Format(Now, "yyyy-mm-dd") + " " + time
        If isTommorrow Then
            time = DateAdd("d", 1, time)
        End If
        time = Format(time, "yyyy-mm-dd hh:mm")
    End If
    getTime = time
    Exit Function
errh:
    getTime = ""
End Function

'Private Function getTodayTime(ByVal time As String) As String
'    On Error GoTo errh
'    Dim temp As String
'    temp = time
'    time = Replace(time, "����", "����")
'    time = Replace(time, "����", "����")
'    time = Replace(time, "����", "����")
'    time = Replace(time, "��", "����")
'    time = Replace(time, "�糿", "����")
'    time = Replace(time, "����", "����")
'
'    time = Replace(time, "����", "")
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
'    time = Replace(time, "����", "")
'    time = Replace(time, "����", "����")
'    time = Replace(time, "����", "����")
'    time = Replace(time, "�糿", "����")
'    time = Replace(time, "����", "����")
'    time = Replace(time, "����", "")
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

  '����
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
''����
'Public Function decrypt(ByVal strSource As String, ByVal key As Byte) As String
'    Dim i     As Long
'    Dim j     As Long
'    Dim temps     As String
'    Dim s     As String
'    Dim arr     As Variant
'    i = Len(strSource)
'    If i Mod 2 = 1 Then
'        '�����ܵ��ִ�������Ҫ��
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
'    content = "��Ϣ�ɻ����Զ�ͳ��" + vbCrLf
'
'    Dim index As Integer
'
'    index = 1
'
'    If Not (rs.EOF And rs.BOF) Then
'        While Not rs.EOF
'            content = content + CStr(index) + "��" + "<����>" + rs.Fields("sendtime") + "  �绰��" + rs.Fields("username") + "  " + rs.Fields("content") + vbCrLf + "                              " + vbCrLf
'
'            rs.MoveNext
'            index = index + 1
'        Wend
'    End If
'    GetContentString = content
'End Function