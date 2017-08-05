
<%  
'On Error Resume Next
'======================================================================================
'C9 静态文章发布系统ACCESS数据库
'官方网站:http://www.csc9.cn
'======================================================================================
'----------------------------------------------------------
Dim dbdir,SaveMenu
    SaveMenu="sys_Download"       '保存目录
    dbdir   ="./sys_dbase/db.mdb"
	
	
	
'==================================================
'数据库连接
'2009-08-09 Kami
'==================================================
set conn=server.CreateObject("adodb.connection")
    conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.MapPath(dbdir)
if err then
  response.Write "连接数据库出错..."
  response.End()
end if

'.mdb数据库路径
'执行一条SQL语句 rs函数
'============
function rsfun(sql,i)
  select case i
    case "1"
	  set rsa=server.CreateObject("adodb.recordset")
	  rsa.open sql,conn,1,1
	  set rsfun=rsa
	  set rsa=nothing
	case "3"
	  set rsa=server.CreateObject("adodb.recordset")
	  rsa.open sql,conn,1,3
	  set rsfun=rsa
	  set rsa=nothing
  end select 
end function
'============
'关闭数据库
'============
sub connclose
conn.close
set conn=nothing
end sub
%>






<%
'==================================================
'MD5加密函数
'2009-08-09 Kami
'==================================================
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)

Private Function LShift(lvalue, iShiftBits)
If iShiftBits = 0 Then
LShift = lvalue
Exit Function
ElseIf iShiftBits = 31 Then
If lvalue And 1 Then
LShift = &H80000000
Else
LShift = 0
End If
Exit Function
ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
Err.Raise 6
End If

If (lvalue And m_l2Power(31 - iShiftBits)) Then
LShift = ((lvalue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
Else
LShift = ((lvalue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
End If
End Function

Private Function RShift(lvalue, iShiftBits)
If iShiftBits = 0 Then
RShift = lvalue
Exit Function
ElseIf iShiftBits = 31 Then
If lvalue And &H80000000 Then
RShift = 1
Else
RShift = 0
End If
Exit Function
ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
Err.Raise 6
End If

RShift = (lvalue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

If (lvalue And &H80000000) Then
RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End If
End Function

Private Function RotateLeft(lvalue, iShiftBits)
RotateLeft = LShift(lvalue, iShiftBits) Or RShift(lvalue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
Dim lX4
Dim lY4
Dim lX8
Dim lY8
Dim lResult

lX8 = lX And &H80000000
lY8 = lY And &H80000000
lX4 = lX And &H40000000
lY4 = lY And &H40000000

lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

If lX4 And lY4 Then
lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
ElseIf lX4 Or lY4 Then
If lResult And &H40000000 Then
lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
Else
lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
End If
Else
lResult = lResult Xor lX8 Xor lY8
End If

AddUnsigned = lResult
End Function

Private Function md5_F(x, y, z)
md5_F = (x And y) Or ((Not x) And z)
End Function

Private Function md5_G(x, y, z)
md5_G = (x And z) Or (y And (Not z))
End Function

Private Function md5_H(x, y, z)
md5_H = (x Xor y Xor z)
End Function

Private Function md5_I(x, y, z)
md5_I = (y Xor (x Or (Not z)))
End Function

Private Sub md5_FF(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub

Private Sub md5_GG(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub

Private Sub md5_HH(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub

Private Sub md5_II(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
Dim lMessageLength
Dim lNumberOfWords
Dim lWordArray()
Dim lBytePosition
Dim lByteCount
Dim lWordCount

Const MODULUS_BITS = 512
Const CONGRUENT_BITS = 448

lMessageLength = Len(sMessage)

lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
ReDim lWordArray(lNumberOfWords - 1)

lBytePosition = 0
lByteCount = 0
Do Until lByteCount >= lMessageLength
lWordCount = lByteCount \ BYTES_TO_A_WORD
lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
lByteCount = lByteCount + 1
Loop

lWordCount = lByteCount \ BYTES_TO_A_WORD
lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)

ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lvalue)
Dim lByte
Dim lCount

For lCount = 0 To 3
lByte = RShift(lvalue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
Next
End Function

Public Function MD5(sMessage)
m_lOnBits(0) = CLng(1)
m_lOnBits(1) = CLng(3)
m_lOnBits(2) = CLng(7)
m_lOnBits(3) = CLng(15)
m_lOnBits(4) = CLng(31)
m_lOnBits(5) = CLng(63)
m_lOnBits(6) = CLng(127)
m_lOnBits(7) = CLng(255)
m_lOnBits(8) = CLng(511)
m_lOnBits(9) = CLng(1023)
m_lOnBits(10) = CLng(2047)
m_lOnBits(11) = CLng(4095)
m_lOnBits(12) = CLng(8191)
m_lOnBits(13) = CLng(16383)
m_lOnBits(14) = CLng(32767)
m_lOnBits(15) = CLng(65535)
m_lOnBits(16) = CLng(131071)
m_lOnBits(17) = CLng(262143)
m_lOnBits(18) = CLng(524287)
m_lOnBits(19) = CLng(1048575)
m_lOnBits(20) = CLng(2097151)
m_lOnBits(21) = CLng(4194303)
m_lOnBits(22) = CLng(8388607)
m_lOnBits(23) = CLng(16777215)
m_lOnBits(24) = CLng(33554431)
m_lOnBits(25) = CLng(67108863)
m_lOnBits(26) = CLng(134217727)
m_lOnBits(27) = CLng(268435455)
m_lOnBits(28) = CLng(536870911)
m_lOnBits(29) = CLng(1073741823)
m_lOnBits(30) = CLng(2147483647)

m_l2Power(0) = CLng(1)
m_l2Power(1) = CLng(2)
m_l2Power(2) = CLng(4)
m_l2Power(3) = CLng(8)
m_l2Power(4) = CLng(16)
m_l2Power(5) = CLng(32)
m_l2Power(6) = CLng(64)
m_l2Power(7) = CLng(128)
m_l2Power(8) = CLng(256)
m_l2Power(9) = CLng(512)
m_l2Power(10) = CLng(1024)
m_l2Power(11) = CLng(2048)
m_l2Power(12) = CLng(4096)
m_l2Power(13) = CLng(8192)
m_l2Power(14) = CLng(16384)
m_l2Power(15) = CLng(32768)
m_l2Power(16) = CLng(65536)
m_l2Power(17) = CLng(131072)
m_l2Power(18) = CLng(262144)
m_l2Power(19) = CLng(524288)
m_l2Power(20) = CLng(1048576)
m_l2Power(21) = CLng(2097152)
m_l2Power(22) = CLng(4194304)
m_l2Power(23) = CLng(8388608)
m_l2Power(24) = CLng(16777216)
m_l2Power(25) = CLng(33554432)
m_l2Power(26) = CLng(67108864)
m_l2Power(27) = CLng(134217728)
m_l2Power(28) = CLng(268435456)
m_l2Power(29) = CLng(536870912)
m_l2Power(30) = CLng(1073741824)


Dim x
Dim k
Dim AA
Dim BB
Dim CC
Dim DD
Dim a
Dim b
Dim c
Dim d

Const S11 = 7
Const S12 = 12
Const S13 = 17
Const S14 = 22
Const S21 = 5
Const S22 = 9
Const S23 = 14
Const S24 = 20
Const S31 = 4
Const S32 = 11
Const S33 = 16
Const S34 = 23
Const S41 = 6
Const S42 = 10
Const S43 = 15
Const S44 = 21

x = ConvertToWordArray(sMessage)

a = &H67452301
b = &HEFCDAB89
c = &H98BADCFE
d = &H10325476

For k = 0 To UBound(x) Step 16
AA = a
BB = b
CC = c
DD = d

md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
md5_FF b, c, d, a, x(k + 15), S14, &H49B40821

md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
md5_GG d, a, b, c, x(k + 10), S22, &H2441453
md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A

md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665

md5_II a, b, c, d, x(k + 0), S41, &HF4292244
md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
md5_II c, d, a, b, x(k + 6), S43, &HA3014314
md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
md5_II b, c, d, a, x(k + 9), S44, &HEB86D391

a = AddUnsigned(a, AA)
b = AddUnsigned(b, BB)
c = AddUnsigned(c, CC)
d = AddUnsigned(d, DD)
Next

'MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
MD5=LCase(WordToHex(b) & WordToHex(c)) 
End Function
%> 



<%  

'==============================
'取出中间内容函数 midstr(str,str1,str2)
'str 内容 str1,开始 str2,结束
'==============================
function midstr(str,str1,str2)
  dim midstr1,midstr2
  midstr1=split(str,str1)
  if ubound(midstr1)>0 then
    midstr2=split(midstr1(1),str2)
	if ubound(midstr2)>0 then
	  midstr=midstr2(0)
	else
	  midstr="0"
	end if
  else
    midstr="0"
  end if
end function


'==================================================
'字符函数 funstr
'2009-06-05 Crazy
'==================================================
Function funstr(str)	 
	str = trim(str) 	 
	str = replace(str, "<", "&lt;", 1, -1, 1)
	str = replace(str, ">", "&gt;", 1, -1, 1)
	str = replace(str,"'","‘")
	funstr = str
End Function


'==================================================
'读取文件内容
'==================================================
function openfile(url)
  fileurl=server.MapPath(url)
  set fso=server.CreateObject("scripting.filesystemobject") '定义FSO
  set mofile=fso.opentextfile(fileurl,1) '以读的方式打开文件
  mo_top=mofile.readall() '读取全部内容
  mofile.close
  openfile=mo_top
end function

'==================================================
'写入文件内容
'==================================================
sub createfile(url,str)
  fileurl=server.MapPath(url)
  set fso=server.CreateObject("scripting.filesystemobject") 
  set mofile=fso.createtextfile(fileurl,true)
  mofile.write str
  mofile.close
end sub

'==================================================
'检查文件夹是否存在,如果不存在则创建
'==================================================
sub createfolder(folder)
  folderurl=server.MapPath(folder)
  set fso=server.CreateObject("scripting.filesystemobject") 
  if fso.folderexists(folderurl) then
  
  else
    fso.createfolder(folderurl)
  end if
set fso=nothing '''*****
end sub


'==================================================
'弹出对话框
'==================================================
sub getshow(str,url)
  if str="" and url <>"" then
    response.Write "<script language='javascript'>window.document.location.href='"& url &"'</script>"
    response.End()
  end if
  if str<>"" and url="" then
    response.Write "<script language='javascript'>alert('"&str&"');history.go(-1)</script>"
    response.End()
  end if
  if str<>"" and url<>"" then
    response.Write "<script language='javascript'>alert('"&str&"');window.document.location.href='"& url &"'</script>"
    response.End()
  end if
end sub


'==================================================
'清除HTML代码函数
'==================================================
Function ClearHtml(Content) 
Content=Zxj_ReplaceHtml("&#[^>]*;", "", Content) 
Content=Zxj_ReplaceHtml("</?marquee[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?object[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?param[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?embed[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?table[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml(" ","",Content) 
Content=Zxj_ReplaceHtml("</?tr[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?th[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?p[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?a[^>]*>","",Content) 
'Content=Zxj_ReplaceHtml("</?img[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?tbody[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?li[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?span[^>]*>","",Content) 

Content=Zxj_ReplaceHtml("</?div[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?th[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?td[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?script[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("(javascript|jscript|vbscript|vbs):", "", Content) 
Content=Zxj_ReplaceHtml("on(mouse|exit|error|click|key)", "", Content) 
Content=Zxj_ReplaceHtml("<\\?xml[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("<\/?[a-z]+:[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?font[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?b[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?u[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?i[^>]*>","",Content)
Content=Zxj_ReplaceHtml("</?strong[^>]*>","",Content) 

ClearHtml=Content 
End Function 



Function Zxj_ReplaceHtml(patrn, strng,content) 
IF IsNull(content) Then 
content="" 
End IF 
Set regEx = New RegExp ' 建立正则表达式。 
regEx.Pattern = patrn ' 设置模式。 
regEx.IgnoreCase = true ' 设置忽略字符大小写。 
regEx.Global = true ' 设置全局可用性。 
Zxj_ReplaceHtml=regEx.Replace(content,strng) ' 执行正则匹配 
End Function

'匹配第一项符合的要求把值输出
function reimgone(patrn,str)
  reimgone=""
  Set re = New RegExp
  re.Pattern = patrn
  re.IgnoreCase = true
  re.Global = true
  set reexe=re.Execute(str)
  
    if isnull(reexe(0)) then
      reimgone=""
    else
	  reimgone=reexe(0)
    end if
end function

'获取网络数据函数
 Function Gethttppage(Path,crzchars)      
      T = Getbody(Path)
      Gethttppage=Bytestobstr(T,crzchars)
 End Function

 Function Getbody(Url)          
 On Error Resume Next
 Set Retrieval = Createobject("Microsoft.Xmlhttp") 
   
 Retrieval.Open "Get", Url, False, "", "" 
 Retrieval.Send 
 Getbody = Retrieval.Responsebody
 Set Retrieval = Nothing 
 End Function 

Function BytesToBstr(body,Cset)         
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText 
        objstream.Close
        set objstream = nothing
End Function

'获取网络数据函数结束


'==================================================
'函数名：DefiniteUrl
'作 用：将相对地址转换为绝对地址
'参 数：PrimitiveUrl ------要转换的相对地址
'参 数：ConsultUrl ------当前网页地址
'==================================================
Function DefiniteUrl(Byval PrimitiveUrl,Byval ConsultUrl)
Dim ConTemp,PriTemp,Pi,Ci,PriArray,ConArray
If PrimitiveUrl="" or ConsultUrl="" or PrimitiveUrl="$False$" Then
DefiniteUrl="$False$"
Exit Function
End If
If Left(ConsultUrl,7)<>"HTTP://" And Left(ConsultUrl,7)<>"http://" Then
ConsultUrl= "http://" & ConsultUrl
End If
ConsultUrl=Replace(ConsultUrl,"://",":\\")
If Right(ConsultUrl,1)<>"/" Then
If Instr(ConsultUrl,"/")>0 Then
If Instr(Right(ConsultUrl,Len(ConsultUrl)-InstrRev(ConsultUrl,"/")),".")>0 then 
Else
ConsultUrl=ConsultUrl & "/"
End If
Else
ConsultUrl=ConsultUrl & "/"
End If
End If
ConArray=Split(ConsultUrl,"/")
If Left(PrimitiveUrl,7) = "http://" then
DefiniteUrl=Replace(PrimitiveUrl,"://",":\\")
ElseIf Left(PrimitiveUrl,1) = "/" Then
DefiniteUrl=ConArray(0) & PrimitiveUrl
ElseIf Left(PrimitiveUrl,2)="./" Then
DefiniteUrl=ConArray(0) & Right(PrimitiveUrl,Len(PrimitiveUrl)-1)
ElseIf Left(PrimitiveUrl,3)="../" then
Do While Left(PrimitiveUrl,3)="../"
PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-3)
Pi=Pi+1
Loop 
For Ci=0 to (Ubound(ConArray)-1-Pi)
If DefiniteUrl<>"" Then
DefiniteUrl=DefiniteUrl & "/" & ConArray(Ci)
Else
DefiniteUrl=ConArray(Ci)
End If
Next
DefiniteUrl=DefiniteUrl & "/" & PrimitiveUrl
Else
If Instr(PrimitiveUrl,"/")>0 Then
PriArray=Split(PrimitiveUrl,"/")
If Instr(PriArray(0),".")>0 Then
If Right(PrimitiveUrl,1)="/" Then
DefiniteUrl="http:\\" & PrimitiveUrl
Else
If Instr(PriArray(Ubound(PriArray)-1),".")>0 Then 
DefiniteUrl="http:\\" & PrimitiveUrl
Else
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
End If
End If 
Else
If Right(ConsultUrl,1)="/" Then 
DefiniteUrl=ConsultUrl & PrimitiveUrl
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
End If
End If
Else
If Instr(PrimitiveUrl,".")>0 Then
If Right(ConsultUrl,1)="/" Then
If right(PrimitiveUrl,3)=".cn" or right(PrimitiveUrl,3)="com" or right(PrimitiveUrl,3)="net" or right(PrimitiveUrl,3)="org" Then
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
Else
DefiniteUrl=ConsultUrl & PrimitiveUrl
End If
Else
If right(PrimitiveUrl,3)=".cn" or right(PrimitiveUrl,3)="com" or right(PrimitiveUrl,3)="net" or right(PrimitiveUrl,3)="org" Then
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl
End If
End If
Else
If Right(ConsultUrl,1)="/" Then
DefiniteUrl=ConsultUrl & PrimitiveUrl & "/"
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl & "/"
End If 
End If
End If
End If
If Left(DefiniteUrl,1)="/" then
DefiniteUrl=Right(DefiniteUrl,Len(DefiniteUrl)-1)
End if
If DefiniteUrl<>"" Then
DefiniteUrl=Replace(DefiniteUrl,"//","/")
DefiniteUrl=Replace(DefiniteUrl,":\\","://")
Else
DefiniteUrl="$False$"
End If
End Function



'==================================================
'过程名：SaveRemoteFile
'作 用：保存远程的文件到本地
'参 数：LocalFileName ------ 本地文件名
'参 数：RemoteFileUrl ------ 远程文件URL
'==================================================
Function SaveRemoteFile(LocalFileName,RemoteFileUrl)
    on error resume next
	LocalFileName=strToPath(LocalFileName)
	LocalFileName=server.MapPath(LocalFileName)
	C_LocalFileName=LocalFileName
	call creatfolder(C_LocalFileName)   '创建目录
    response.Write("<br>&nbsp;Downning: "&RemoteFileUrl)
    response.Flush()
If not IsFileExist(LocalFileName) then
dim Ads,Retrieval,GetRemoteData
Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
    With Retrieval
      .Open "Get", RemoteFileUrl, False, "", ""
      .Send
      GetRemoteData = .ResponseBody
    End With
Set Retrieval = Nothing
Set Ads = Server.CreateObject("Adodb.Stream")
    With Ads
      .Type = 1
      .Open
      .Write GetRemoteData
      .SaveToFile LocalFileName,2
      .Cancel()
      .Close()
    End With
Set Ads=nothing
'-------------------
end if
end Function

'========= 转换字符 ==============
function strToPath(str)
    rem 检查过滤字符
    dim dist,ToStr
	    ToStr="_"
	    dist =">|<|:|*|?"
        dists=split(dist,"|")
		for i=0 to ubound(dists)
		    if dists(i)<>"" then str=replace(str,dists(i),ToStr)
        next	
    strToPath=str
end function


'检测文件是否存在
Function IsFileExist(filespec)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec)) Then
      IsFileExist=true
   Else
      IsFileExist=false
   End If
   set fso=nothing
End Function


 '创建目录,支持多级
Function creatfolder(w_path)
    '文件夹不存在则生成
	on error resume next
 	set fso=CREATEOBJECT("SCRIPTING.FILESYSTEMOBJECT")
	    w_path=replace(w_path,"/","\")
		w_path=split(w_path,"\")
		for i=0 to ubound(w_path)
		    if w_path(i)<>"" and instr(w_path(i),".")=0 then
		       w_paths=w_paths&w_path(i)&"\"
			   if fso.folderexists(w_paths)=false then fso.createfolder(w_paths)
	        end if
		next   
		creatfolder=true            
	set fso=nothing
End Function


%>






















































<%  
'C9静态文章发布系统 采集插件1.0
   rel=funstr(request.QueryString("rel")) 
On Error Resume Next
   times=date()
   response.Buffer=true
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网络采集下载系统 v.10 [ 卡mi.伊凡 http://cm.ivan.blog.163.com ]</title>
<style type="text/css">
<!--
*{
	font-family:Verdana, Arial, Helvetica, sans-serif;
}
body {
	background:#F7F7F7;
	font-size:12px;
}
input{
	vertical-align:middle;
}
img{
	border:none;
	vertical-align:middle;
}
a{
	color:#333333;
}
a:hover{
	color:#FF3300;
	text-decoration:none;
}
.main{
	width:640px;
	margin:40px auto 0px;
	border:4px solid #EEE;
	background:#FFF;
	padding-bottom:10px;
}
.main .title{
	width:600px;
	height:50px;
	margin:0px auto;
	background:url(images/login_toptitle.jpg) -10px 0px no-repeat;
	text-indent:326px;
	line-height:46px;
	font-size:14px;
	letter-spacing:2px;
	color:#F60;
	font-weight:bold;
}
.main .login{
	width:560px;
	margin:20px auto 0px;
	overflow:hidden;
}
.main .login .inputbox{
	width:230px;
	float:left;
	background:url(images/login_input_hr.gif) right center no-repeat;
}
.main .login .inputbox dl{
	width:230px;
	height:38px;
	clear:both;
}
.main .login .inputbox dl dt{
	float:left;
	width:60px;
	height:31px;
	line-height:31px;
	text-align:right;
	font-weight:bold;
}
.main .login .inputbox dl dd{
	width:160px;
	float:right;
	padding-top:1px;
}
.main .login .inputbox dl dd input{
	font-size:12px;
	font-weight:bold;
	border:1px solid #888;
	padding:4px;
}
.main .login .butbox{
	float:left;
	width:200px;
	margin-left:15px;
}
.main .login .butbox dl{
	width:200px;
}
.main .login .butbox dl dt{
	width:160px;
	height:38px;
	padding-top:1px;
}
.main .login .butbox dl dt input{
	width:75px;
	height:61px;
	border:1px solid #888;
	cursor:pointer;
	border:1px solid #888;
	background-color:#F93;
	color: #FFFFFF;
}
.main .login .butbox dl dd{
	height:21px;
	line-height:21px;
}
.main .login .butbox dl dd a{
	margin:5px;
}
.main .msg{
	width:560px;
	margin:10px auto;
	clear:both;
	line-height:17px;
	padding:6px;
	border:1px solid #FC9;
	background:#FFFFCC;
	color:#666;
}
.copyright{
	width:640px;
	text-align:right;
	margin:10px auto;
	font-size:10px;
	color:#999999;
}
.copyright a{
	font-weight:bold;
	color:#F63;
	text-decoration:none;
}
.copyright a:hover{
	color:#000;
}
-->
</style>

<script type="text/javascript">

// JavaScript Document
// JavaScript Document
//在鼠标显示一个层，该层的内空为div2的内容 
function showdiv(divname){ 
var div3 = document.getElementById(divname); //将要弹出的层 
div3.style.display="block"; //div3初始状态是不可见的，设置可为可见 
//window.event代表事件状态，如事件发生的元素，键盘状态，鼠标位置和鼠标按钮状. 
//clientX是鼠标指针位置相对于窗口客户区域的 x 坐标，其中客户区域不包括窗口自身的控件和滚动条。 
div3.style.left=event.clientX+10; //鼠标目前在X轴上的位置，加10是为了向右边移动10个px方便看到内容 
div3.style.top=event.clientY+5; 
div3.style.position="absolute"; //必须指定这个属性，否则div3层无法跟着鼠标动 
//var div2 =document.getElementById('div2'); 
//div3.innerText=div2.innerHTML; 
} 
//关闭层div3的显示 
function closediv(divname){ 
var div3 = document.getElementById(divname); 
div3.style.display="none"; 
}

function winshow(pagename,w,h){
  window.open(pagename,null,"width="+w+",height="+h);
}
function checkbox(obj,num){
  var id;
  for (i=1;i<=num;i++){
	id=obj+i;
	if(document.getElementById(id).checked==""){
	  document.getElementById(id).checked="checked";
	}
	else{
	  document.getElementById(id).checked="";
	}
  }
}



var xhr;
function getXHR() {
 try {
  xhr=new ActiveXObject("Msxml2.XMLHTTP");
 } catch (e) {
 try {
  xhr=new ActiveXObject("Microsoft.XMLHTTP");
 } catch (e) {
 xhr=false;
}
}
if(!xhr&&typeof XMLHttpRequest!='undefined') {
xhr=new XMLHttpRequest();
}
return xhr;
}

function openXHR(method,url) {
getXHR();
xhr.open(method,url,true);
xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
xhr.onreadystatechange=function() {
if(xhr.readyState!=4)return;
  document.write(xhr.responseText);//responseBody
}
xhr.send(null);
}

function loadXML(method,url) {
getXHR();
xhr.open(method,url,true);
xhr.setRequestHeader("Content-Type","text/xml");
xhr.setRequestHeader("Content-Type","GBK");
xhr.onreadystatechange=function() {
if(xhr.readyState!=4) return;
   document.write(xhr.responseText);
  //$("abc").innerHTML = xhr.responseText;
}
xhr.send(null);
}

function $(idValue)
{
  return document.getElementById(idValue);
}

</script>
</head>

<body>
<%if len(session("adminuser"))-4<0 then
'================================  登陆页面 ==========================================================
%>

<style type="text/css">
<!--
*{
	padding:0px;
	margin:0px;
	font-family:Verdana, Arial, Helvetica, sans-serif;
}
body {
	margin: 0px;
	background:#F7F7F7;
	font-size:12px;
}
-->
</style>
<script type="text/javascript" language="javascript">
<!--
	window.onload = function (){
		userid = document.getElementById("username");
		userid.focus();
	}
-->
</script>
	<div class="main">
		<div class="title"></div>

		<div class="login">
		<form action="?" method="post">
            <div class="inputbox">
				<dl>
					<dt>用户名：</dt>
					<dd><input type="text" name="username" id="username" size="20" onfocus="this.style.borderColor='#F93'" onblur="this.style.borderColor='#888'" />
					</dd>
				</dl>
				<dl>
					<dt>密码：</dt>
					<dd><input type="password" name="password" size="20" onfocus="this.style.borderColor='#F93'" onblur="this.style.borderColor='#888'" />
					</dd>
				</dl>
					
          </div>
            <div class="butbox">
            <dl>
					<dt><input name="submit" type="submit" value="登陆" /></dt>
			  </dl>
			</div>
		</form>
		</div>
	</div>
	
<%  
'======================================================================================
'C9 静态文章发布系统
'官方网站:http://www.csc9.cn
'======================================================================================
username=funstr(request.Form("username"))
password=funstr(request.Form("password"))
'获取表单传递过来的值,使用funstr函数进行处理,过滤掉无效字符

if username<>"" and password<>"" then
'----------------------------------------------------
if len(usernmae)-20>0 or len(username)-4<0 then
  getshow "用户名不得大于20,小于4",""
end if
'判断用户名是否大于20或小于4
if len(password)-20>0 or len(password)-4<0 then
  getshow "密码不得大于20,小于4",""
end if
'判断密码是否大于20或小于4
sql="select * from user where password='"&md5(password)&"' and username='"&username&"'"
'sql语句同时查询帐户和密码
set rs=rsfun(sql,1)

if not rs.eof then
'判断记录集不在最后一条记录之后为真.
  session("adminuser")=username
  session("adminvip")=rs("vip")
'session变量赋值
  'response.Cookies("adminuser")=username
  'response.Cookies("adminuser").Expires=DateAdd("d",1,NOW)
  response.Redirect "index.asp"
'转向index.asp文件
else
  getshow "用户名或密码错误!","index.asp"
end if
'----------------------------------------------------
end if
connclose
%>




<%
else
'================================  操作页面 ==========================================================
%>


<table width="780" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E2F5BC" class="sysinfo">
<tr align="center" height="24">
  <td colspan="3" align="left" bgcolor="#EDF9D5"> <b>&nbsp;提示:</b>&nbsp;&nbsp;<a href="?rel=add">添加采集规则</a>&nbsp;|&nbsp;<a href="?">规则列表</a>&nbsp;|&nbsp;<a href="?rel=exit">退出系统</a></td>
</tr></table>
<div style=" border:#E2F5BC 1px solid; border-top:0;padding:2px;width:774px; margin:auto;">


<% if rel="" then  %>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E2F5BC" class="sysinfo">
<tr align="center" bgcolor="#FFFFFF"> 
<td align="left" bgcolor="#FFFFFF">保存目录</td>
<td align="left" bgcolor="#FFFFFF">采集地址</td>
<td width="200" align="center" bgcolor="#FFFFFF">操作</td>
</tr>
<%  
set rs=rsfun("select * from getinfo order by id desc",1)
do while not rs.eof
%>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="left" bgcolor="#FFFFFF"><%= rs("names") %></td>
<td height="24" align="left" bgcolor="#FFFFFF"><a href="<%= rs("urls") %>" target="_blank"><%= rs("urls") %></a></td>
<td height="24" align="center" bgcolor="#FFFFFF"><a href="?rel=getces&id=<%= rs("id") %>">测试</a> <a href="?rel=get0&id=<%= rs("id") %>">采集</a> <a href="?rel=dao1&id=<%= rs("id") %>">导出</a> <a href="?rel=dao2&id=<%= rs("id") %>">导入</a> <a href="?rel=info&id=<%= rs("id") %>">修改</a> <a href="?rel=del&id=<%= rs("id") %>">删除</a></td>
</tr>
<%  
rs.movenext
loop
%>
</table>
<% end if %>


<%
'----------退出系统-------------
if rel="exit" then
session.Abandon()
response.Redirect("?")
end if
%>





<% 
if rel="get0" then 

  id=funstr(request.QueryString("id"))
  set rs1=rsfun("select * from getinfo where id="&id,1)
  if not rs1.eof then
    urls=rs1("urls")
	str=Gethttppage(urls,rs1("bian"))
	
	
	str=replace(str,"'","‘")
	
	urlintervalsarr=split(rs1("urlintervals"),"[c9]")
	urlintervals=midstr(str,urlintervalsarr(0),urlintervalsarr(1))
	
	url=""
	rulesarr=split(rs1("rules"),"[c9]")
	rules=split(urlintervals,rulesarr(0))
	for rulesi=1 to ubound(rules)
	  rules2=split(rules(rulesi),rulesarr(1))
	  
	  if rs1("urlprefixs")<>"" then rules2(0)=rs1("urlprefixs")&rules2(0)
	  if rs1("urlincludes")<>"" then
	    urlincludes=split(rules2(0),rs1("urlincludes"))
		if ubound(urlincludes)>0 then url=url&",,"&rules2(0)
	  else
	    url=url&",,"&rules2(0)
url=replace(url,""" target=""_blank","")

	  end if
	  
	next

	createfolder("sys_dbase/getinfo/")
	urlfile="sys_dbase/getinfo/"&id&".txt"
	createfile urlfile,url
	
  end if
  set rs1=nothing

  
  getshow "","?rel=get1&id="&id
end if
%>

<%  
if rel="get1" then
  id=funstr(request.QueryString("id"))
  urlfile="sys_dbase/getinfo/"&id&".txt"
  urlinfo=openfile(urlfile)
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E2F5BC" class="sysinfo">
<form id="form2" name="form2" method="post" action="?rel=get2">
<%urlpage=split(urlinfo,",,")%>
<tr align="center" height="24">
  <td align="left" bgcolor="#FFFFFF"><a href="javascript:checkbox('checkbox',<%=ubound(urlpage) %>);">全选/取消</a>
  <input type="hidden" name="id" value="<%= id %>" /><input type="hidden" name="num" value="<%= ubound(urlpage) %>" /> <input type="submit" name="Submit13" value="批量采集" /></td>
</tr>
<% for i=1 to ubound(urlpage)%>
<tr align="center" height="24">
  <td align="left" bgcolor="#FFFFFF"><input name="checkbox<%= i %>" type="checkbox" value="<%= urlpage(i) %>" checked="CHECKED" />
  <a href="<%= urlpage(i) %>" target="_blank"><%= urlpage(i) %></a></td>
</tr>
<% next%>
</form>
</table>

<%  
end if
%>

<%  
if rel="get2" then
  id=funstr(request.Form("id"))
  num=funstr(request.Form("num"))
  i=1
  urlpagearr=""
  do while num-i>0
    urlstr="checkbox"&i
    urlpage=request.Form(urlstr)
	if urlpage="" then
    else
      urlpagearr=urlpagearr&",,"&urlpage
    end if
    i=i+1
  loop
  
  urlfile="sys_dbase/getinfo/"&id&"_.txt"
  createfile urlfile,urlpagearr
  getshow "","?rel=get3&num=1&id="&id
end if
%>

<%  
if rel="get3" then
  id=funstr(request.QueryString("id"))
  num=funstr(request.QueryString("num"))
  urlfile="sys_dbase/getinfo/"&id&"_.txt"
  urlfiles=openfile(urlfile)
  urlfiles=replace(urlfiles,"'","‘")
  numarr=split(urlfiles,",,")
  
  if int(ubound(numarr))<int(num) then
    response.Write "全部信息采集完毕!"
	response.End()
  else
  
  set rs1=rsfun("select * from getinfo where id="&id,1)
  if not rs1.eof then
    names=rs1("names")   '用于存放目录
    urlces=numarr(num)
	str=Gethttppage(urlces,rs1("bian"))

	tags=split(rs1("tags"),",")
	for tagsi=0 to ubound(tags)
	  tags2=split(tags(tagsi),"@")
	  str=replace(str,tags2(0),tags2(1))
	next

	titlesarr=split(rs1("titles"),"[c9]")
	titles=midstr(str,titlesarr(0),titlesarr(1))
        if left(titles,1)=" " then
           titles=right(titles,len(titles)-1)
        end if
	
	mgsarr=split(rs1("mgs"),"[c9]")
	mgs=midstr(str,mgsarr(0),mgsarr(1))
	
'------------- 保存文件类型 --------------------
Dim T_mgs,U_mgs
    U_mgs=mgs
    T_mgs=split(U_mgs,".")
	F_mgs=T_mgs(ubound(T_mgs))
	
	if len(F_mgs)>3 or len(F_mgs)<2 then T_mgs="html"
	if rs1("audits") then mgs=clearhtml(mgs)

    mgs=replace(mgs,"'","‘")
'-----------------------------------------
    titles=replace(titles,"\","_")
    titles=replace(titles,"/","_")
    Ntitles= SaveMenu&"\"&names&"\"&titles
    response.Write "<br>&nbsp;&nbsp;完成采集:" &urlces
	response.Write "<br>&nbsp;&nbsp;正保存文件... "&Ntitles&"."&F_mgs
	response.Flush()
	SaveRemoteFile "./"&Ntitles&"."&F_mgs,mgs
	
    imgurl=""
	fckarr=split(LCase(mgs),"<img")
    if imgurl="" and ubound(fckarr)>0 then
	  
      imgurl=reimgone("src=[\""]?(.[^<]*)(gif|jpg|png|bmp)",mgs)
	  imgurl=replace(imgurl,"src=","")
	  imgurl=replace(imgurl,"""","")
	  imgurl=DefiniteUrl(imgurl,urlces)
    end if
	
  set rsp=rsfun("insert into page(title,keywords,dis,user,web,vip,mg,times,imgurl,tou,huan,ding,audits)values('"&titles&"','"&titles&"','"&titles&"','"&rs1("users")&"','"&rs1("source")&"','0','"&mgs&"','"&times&"','"&imgurl&"','0','0','0','1')",3)
  '存储记录
  
  set rsp=nothing
      end if
  set rs1=nothing
  
  getshow "","?rel=pageid&num="&num&"&id="&id
  end if
  
end if
%>

<%  
'发布或修改后直接开始采集下一个.
if rel="pageid" then
  id=funstr(request.QueryString("id"))
  num=funstr(request.QueryString("num"))
  getshow "","?rel=get3&num="&(num+1)&"&id="&id
end if
%>

<!--测试采集-->
<%  
if rel="getces" then
   titles="尚未采集到标题<br />"
      mgs="尚未采集到内容<br />"
  id=funstr(request.QueryString("id"))
  set rs1=rsfun("select * from getinfo where id="&id,1)
  if not rs1.eof then
    urlces=rs1("urlces")
	str=Gethttppage(urlces,rs1("bian"))

	
	tags=split(rs1("tags"),",")
	for tagsi=0 to ubound(tags)
	  tags2=split(tags(tagsi),"@")
	  str=replace(str,tags2(0),tags2(1))
	next
	
	titlesarr=split(rs1("titles"),"[c9]")
	titles=midstr(str,titlesarr(0),titlesarr(1))
	
	mgsarr=split(rs1("mgs"),"[c9]")
	mgs=midstr(str,mgsarr(0),mgsarr(1))
	
	if rs1("audits") then mgs=clearhtml(mgs)
  end if
  set rs1=nothing
  
  response.Write "<font color='#FF0000'><b>标题:</b></font>"&titles&"<br />"
  response.Write "<font color='#FF0000'><b>内容:</b></font>"&mgs
  
end if
%>
<!--测试采集结束-->

<% if rel="add" then %>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E2F5BC" class="sysinfo">
<form id="form1" name="form1" method="post" action="?rel=add_info">

<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">保存目录:</td>
<td width="48%" align="left" bgcolor="#FFFFFF"><input type="text" name="names" /></td>
<td width="37%" align="left" bgcolor="#FFFFFF">采集规则标识</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">采集地址:</td>
<td align="left" bgcolor="#FFFFFF"><input name="urls" type="text" size="40" value="http://" /></td>
<td align="left" bgcolor="#FFFFFF">指定采集地址</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">目标编码:</td>
<td align="left" bgcolor="#FFFFFF">
<select name="bian">
<option value="utf-8">utf-8</option>
<option value="gb2312" selected="selected">gb2312</option>
<option value="gbk">gbk</option>
<option value="big5">big5</option>
</select></td>
<td align="left" bgcolor="#FFFFFF">指定目标编码方式</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">采集区间:</td>
<td align="left" bgcolor="#FFFFFF"><textarea name="urlintervals" cols="50" rows="3"></textarea></td>
<td align="left" bgcolor="#FFFFFF">列表区间识别符: [c9],如:&lt;td&gt;[c9]&lt;/td&gt;</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">地址规则:</td>
<td align="left" bgcolor="#FFFFFF"><textarea name="rules" cols="50" rows="3"></textarea></td>
<td align="left" bgcolor="#FFFFFF">解析出对应的文章地址,用 [c9] 标识! </td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">地址包含:</td>
<td align="left" bgcolor="#FFFFFF"><input name="urlincludes" type="text" size="40"/></td>
<td align="left" bgcolor="#FFFFFF">信息地址必须包含</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">地址补全:</td>
<td align="left" bgcolor="#FFFFFF"><input name="urlprefixs" type="text" size="40"/></td>
<td align="left" bgcolor="#FFFFFF">如采集到的地址为相对地址,在此设置补全地址.</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">信息测试:</td>
<td align="left" bgcolor="#FFFFFF"><input name="urlces" type="text" size="40"/></td>
<td align="left" bgcolor="#FFFFFF">测试此信息地址</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">替换字符:</td>
<td align="left" bgcolor="#FFFFFF"><textarea name="tags" cols="50" rows="3"></textarea></td>
<td align="left" bgcolor="#FFFFFF">替换文章信息页字符,替换规则如下:<br />
替换字符1@<br />
替换成字符1,替换字符2@替换成字符2</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">保存名称:</td>
<td align="left" bgcolor="#FFFFFF"><input name="titles" type="text" value="&lt;title&gt;[c9]&lt;/title&gt;" size="40" /></td>
<td align="left" bgcolor="#FFFFFF">名称规则:&lt;title&gt;[c9]&lt;/title&gt;</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">下载地址:</td>
<td align="left" bgcolor="#FFFFFF"><textarea name="mgs" cols="50" rows="3"></textarea></td>
<td align="left" bgcolor="#FFFFFF">下载地址规则标识为: [c9] </td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">其他选项:</td>
<td align="left" bgcolor="#FFFFFF"><input type="checkbox" name="html" value="1" />去除HTML标签</td>
<td align="left" bgcolor="#FFFFFF"></td>
</tr>
<tr align="center" height="24">
<td align="left" bgcolor="#EDF9D5">&nbsp;</td>
<td align="left" bgcolor="#EDF9D5"><input type="submit" name="Submit3" value="确认添加规则" /></td>
<td align="left" bgcolor="#EDF9D5">&nbsp;</td>
</form>
</table>
<% end if %>


<%  
if rel="add_info" then
  names=request.Form("names")
  urls=request.Form("urls")
  urlintervals=request.Form("urlintervals")
  urlintervals = replace(urlintervals,"'","‘")
  rules=request.Form("rules")
  rules = replace(rules,"'","‘")
  urlincludes=request.Form("urlincludes")
  urlprefixs=request.Form("urlprefixs")
  urlces=request.Form("urlces")
  tags=request.Form("tags")
  titles=request.Form("titles")
  mgs=request.Form("mgs")
  mgs = replace(mgs,"'","‘")
  users=request.Form("users")
  sources=request.Form("source")
  html=request.Form("html")
  bian=request.Form("bian")
  if html="" then html=0

  set rs=rsfun("insert into getinfo(names,urls,urlintervals,rules,urlincludes,urlprefixs,tags,titles,mgs,users,source,audits,urlces,bian)values('"&names&"','"&urls&"','"&urlintervals&"','"&rules&"','"&urlincludes&"','"&urlprefixs&"','"&tags&"','"&titles&"','"&mgs&"','"&users&"','"&sources&"','"&html&"','"&urlces&"','"&bian&"')",3)
  set rs=nothing

  if err<>0 then
  response.Write(err.description)
  response.End()
  end if
  
  getshow "添加记录成功",""
end if
%>

<% if rel="info" then 
id=request.QueryString("id")
set rs1=rsfun("select * from getinfo where id="&id,1)
if not rs1.eof then
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E2F5BC" class="sysinfo">

<form id="form1" name="form1" method="post" action="?rel=info_info">

<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td width="15%" align="center" bgcolor="#FFFFFF">保存目录:</td>
<td width="48%" align="left" bgcolor="#FFFFFF"><input name="names" type="text" id="names" value="<% =rs1("names") %>" /></td>
<td width="37%" align="left" bgcolor="#FFFFFF">采集规则标识</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">采集地址:</td>
<td align="left" bgcolor="#FFFFFF"><input name="urls" type="text" size="40" value="<% =rs1("urls") %>" /></td>
<td align="left" bgcolor="#FFFFFF">指定采集地址</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">目标编码:</td>
<td align="left" bgcolor="#FFFFFF">
<select name="bian">
<option value="<% =rs1("bian") %>" selected="selected"><% =rs1("bian") %></option>
<option value="utf-8">utf-8</option>
<option value="gb2312">gb2312</option>
<option value="gbk">gbk</option>
<option value="big5">big5</option>
</select></td>
<td align="left" bgcolor="#FFFFFF">指定编码方式</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">采集区间:</td>
<td align="left" bgcolor="#FFFFFF"><textarea name="urlintervals" cols="50" rows="3"><% str = replace(rs1("urlintervals"),"‘","'") %><% =str %></textarea></td>
<td align="left" bgcolor="#FFFFFF">列表区间规则识别符号: [c9]<br />
  如: &lt;td&gt;文章列表&lt;/td&gt;<br />
  用&lt;td&gt;[c9]&lt;/td&gt;标识</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">地址规则:</td>
<td align="left" bgcolor="#FFFFFF"><textarea name="rules" cols="50" rows="3"><% str = replace(rs1("rules"),"‘","'") %><% =str %></textarea></td>
<td align="left" bgcolor="#FFFFFF">对采集区间获取的代码进行分析<br />
解析出对应的文章地址<br />
用 [c9] 标识! </td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">地址包含:</td>
<td align="left" bgcolor="#FFFFFF"><input name="urlincludes" type="text" size="40" value="<% =rs1("urlincludes") %>"/></td>
<td align="left" bgcolor="#FFFFFF">信息地址必须包含</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">地址补全:</td>
<td align="left" bgcolor="#FFFFFF"><input name="urlprefixs" type="text" size="40" value="<% =rs1("urlprefixs") %>"/></td>
<td align="left" bgcolor="#FFFFFF">如采集到的地址为相对地址,在此设置补全地址.</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">信息测试:</td>
<td align="left" bgcolor="#FFFFFF"><input name="urlces" type="text" size="40" value="<% =rs1("urlces") %>"/></td>
<td align="left" bgcolor="#FFFFFF">测试此信息地址</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">替换字符:</td>
<td align="left" bgcolor="#FFFFFF"><textarea name="tags" cols="50" rows="3"><% =rs1("tags") %></textarea></td>
<td align="left" bgcolor="#FFFFFF">替换文章信息页字符,替换规则如下:<br />
替换字符1@<br />
替换成字符
1,
替换字符2@替换成字符2</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">保存名称:</td>
<td align="left" bgcolor="#FFFFFF"><input name="titles" type="text" value="<% =rs1("titles") %>" size="40" /></td>
<td align="left" bgcolor="#FFFFFF">名称规则:&lt;title&gt;[c9]&lt;/title&gt;</td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">下载地址:</td>
<td align="left" bgcolor="#FFFFFF"><textarea name="mgs" cols="50" rows="3"><% str = replace(rs1("mgs"),"‘","'") %><% =str %></textarea></td>
<td align="left" bgcolor="#FFFFFF">下载地址规则标识为: [c9] </td>
</tr>
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="center" bgcolor="#FFFFFF">其他选项:</td>
<td align="left" bgcolor="#FFFFFF"><% 
html=""
if rs1("audits") then html="checked='checked'" %>
<input type="checkbox" name="html" value="1" <% =html %> />去除HTML标签</td>
<td align="left" bgcolor="#FFFFFF"></td>
</tr>
<tr align="center" height="24">
<td align="left" bgcolor="#EDF9D5">&nbsp;</td>
<td align="left" bgcolor="#EDF9D5"><input type="submit" name="Submit2" value="确认添加规则" />
  <input name="id" type="hidden" value="<% =rs1("id") %>" /></td>
<td align="left" bgcolor="#EDF9D5">&nbsp;</td>
</form>
</table>

<% 
end if
set rs1=nothing
end if %>

<%  
if rel="info_info" then
  id=request.Form("id")
  names=request.Form("names")
  urls=request.Form("urls")
  urlintervals=request.Form("urlintervals")
  urlintervals = replace(urlintervals,"'","‘")
  rules=request.Form("rules")
  rules = replace(rules,"'","‘")
  urlincludes=request.Form("urlincludes")
  urlprefixs=request.Form("urlprefixs")
  urlces=request.Form("urlces")
  tags=request.Form("tags")
  titles=request.Form("titles")
  mgs=request.Form("mgs")
  mgs = replace(mgs,"'","‘")
  users=request.Form("users")
  sources=request.Form("source")
  html=request.Form("html")
  bian=request.Form("bian")
  
  if html="" then html=0
  
  set rs=rsfun("update getinfo set names='"&names&"',urls='"&urls&"',urlintervals='"&urlintervals&"',rules='"&rules&"',urlincludes='"&urlincludes&"',urlprefixs='"&urlprefixs&"',tags='"&tags&"',titles='"&titles&"',mgs='"&mgs&"',users='"&users&"',source='"&sources&"',audits='"&html&"',urlces='"&urlces&"',bian='"&bian&"' where id="&id,3)
  set rs=nothing
  if err<>0 then
  response.Write(err.description)
  response.End()
  end if
  getshow "修改记录成功","?"
end if
%>

<% 
if rel="del" then
  id=request.QueryString("id")
  set rs=rsfun("delete from getinfo where id="&id,3)
  getshow "删除记录成功",""
  set rs=nothing
end if
%>

<% 
if rel="dao1" then 
  id=request.QueryString("id")
set rs1=rsfun("select * from getinfo where id="&id,1)
if not rs1.eof then

mg=""
mg=mg&rs1("urls")&"{c9}"
mg=mg&rs1("urlintervals")&"{c9}"
mg=mg&rs1("rules")&"{c9}"
mg=mg&rs1("urlincludes")&"{c9}"
mg=mg&rs1("urlprefixs")&"{c9}"
mg=mg&rs1("urlces")&"{c9}"
mg=mg&rs1("tags")&"{c9}"
mg=mg&rs1("titles")&"{c9}"
mg=mg&rs1("mgs")&"{c9}"
mg=mg&rs1("bian")&"{c9}"
if rs1("audits") then
  audits="1"
else
  audits="0"
end if
mg=mg&audits
mg = replace(mg,"‘","'")
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E2F5BC" class="sysinfo">
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="left" bgcolor="#FFFFFF">
<textarea name="textarea" cols="100" rows="30" style="width:99%"><% =mg %>
</textarea>
</td>
</tr>

</table>


<% 
end if
end if %>

<% 
if rel="dao2" then 
id=request.QueryString("id")
%>
<form id="form1" name="form1" method="post" action="?rel=dao2_info">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E2F5BC" class="sysinfo">
<tr align="center" height="24" bgcolor="#FFFFFF"> 
<td align="left" bgcolor="#FFFFFF">
<textarea rows="30" cols="100" style="width:99%" name="mg"></textarea>
</td>
</tr>
<tr align="center" height="24">
<td align="left" bgcolor="#EDF9D5"> 
 <input name="id" type="hidden" value="<% =id %>" />
<input type="submit" name="Submit" value="确认导入规则" /></td>

</table>
</form>
<% end if %>

<%  
if rel="dao2_info" then
  id=request.Form("id")
  mg=request.Form("mg")
  mg = replace(mg,"'","‘")
  mgs=split(mg,"{c9}")
  
  set rs=rsfun("update getinfo set urls='"&mgs(0)&"',urlintervals='"&mgs(1)&"',rules='"&mgs(2)&"',urlincludes='"&mgs(3)&"',urlprefixs='"&mgs(4)&"',urlces='"&mgs(5)&"',tags='"&mgs(6)&"',titles='"&mgs(7)&"',mgs='"&mgs(8)&"',bian='"&mgs(9)&"',audits='"&mgs(10)&"',times='"&times&"' where id="&id,3)
  set rs=nothing
  getshow "导入成功","?"
end if
%>


</div>




<%end if%>

</body>
</html>
<% connclose %>