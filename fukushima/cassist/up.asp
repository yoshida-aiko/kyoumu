=================== upload.vbs start =========================
Set bobj = WScript.CreateObject("basp21")  
Set bsocket = WScript.CreateObject("basp21.socket")   
rc1 = bsocket.Connect("localhost", 80, 10)
bobj.debugclear
bobj.debug rc1
fname = "d:\basp21.gif"   ' アップロードするファイル名
cmd2 = "--sep" & vbCRLF & _ 
        "Content-Disposition: form-data; name=""file1""; " & _
        "filename=""" & fname & """" & vbCRLF & _ 
        "Content-Type: text/plain" & vbCRLF & vbCRLF
cmd3 = bobj.BinaryRead(fname)   ' 送信ファイルをバイナリで読みます
cmd4 = vbCRLF & "--sep--" & vbCrLf
conlen = len(cmd2) + Ubound(cmd3) + 1 + len(cmd4)  ' 長さを求めます
cmd1 =  "POST /fileup.asp HTTP/1.0" & vbCRLF & _ 
        "Content-Type: multipart/form-data; " & _
        "boundary=""sep""" & vbCRLF & _ 
        "Content-Length:" & conlen & vbCrLf & vbCrLf
rc1 = bsocket.write (cmd1)
rc1 = bsocket.write (cmd2)
rc1 = bsocket.write (cmd3)  ' バイナリもそのまま送れます
rc1 = bsocket.write (cmd4)
bobj.debug cmd1 & cmd2
rc1 = bsocket.read(data)    ' 結果を受信します
bobj.debug data
=================== upload.vbs end  =========================
=================== fileup.asp start =========================
<%
a=Request.TotalBytes
b=Request.BinaryRead(a)
set obj=server.createobject("basp21")
c=obj.binarywrite(b,"d:\fup1.txt")  ' 内容確認
name=obj.Form(b,"yourname")
f1=obj.FormFileName(b,"file1")
fsize1=obj.FormFileSize(b,"file1")
newf1=Mid(f1,InstrRev(f1,"\")+1)
l1=obj.FormSaveAs(b,"file1","d:\temp\fup\" & newf1)
%>
<HTML><HEAD><TITLE>File Upload Test</TITLE>
<BODY><H1>Testing</H1><BR>
<%= name %>さん、アップロードされました<BR>
file1= <%= newf1 %><BR>
len1= <%= l1 %><BR>
=================== fileup.asp end   =========================
