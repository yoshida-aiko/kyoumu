<%
	On Error Resume Next
    Err.Clear

	Dim w_cCn,w_rRs,w_SQL,w_Index
	Dim w_StartCD,w_EndCD,w_Name,w_CheckDel

	w_StartCD = Request.Form("txtStartCD")
	w_EndCD = Request.Form("txtEndCD")
	w_Name = Request.Form("txtName")
	w_CheckDel = Request.Form("checkDel")
	w_cboName = Request.Form("cboName")
	w_FileName = Request.Form("txtFileName")
	w_SQL = Request.Form("SQL")
	
	if w_FileName = "" then
		w_FileName= "Sample"
	end if
	
' �I�u�W�F�N�g�̒�`   
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")
	
    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_�Ј�",w_cCn,2,2
    
'--------------------�S�p�𔼊p�ɕϊ�----------------------------------

	Set bobj = Server.CreateObject("basp21")
	w_StartCD = bobj.StrConv(w_StartCD,8)	'�S�p�����p�ϊ�
	w_EndCD = bobj.StrConv(w_EndCD,8)	'�S�p�����p�ϊ�

    Set w_rRs = w_cCn.Execute(w_SQL)

' SQL���s���̃G���[����
	if Err then
		Session.Contents("SQLerror")=Err.description
		Response.Redirect "SQLerror.asp"
	end if

	On Error Goto 0
    
' �Y������Ј������邩�ǂ����̔���
	if w_rRs.EOF=true then
		Response.Redirect "NOexport.asp"
	end if

'*********************************************************************
'				���݂̃X�N���v�g��URL�p�X�𓾂�
'*********************************************************************
	Function GetURLPath()
		On Error Resume Next
		Dim strURL, nP	  
		strURL = "http://" & _
		  Request.ServerVariables("SERVER_NAME")
		If Request.ServerVariables("SERVER_PORT") <> "80" Then
		  strURL = strURL & ":80"
		End If
		strURL = strURL & "/" & Request.ServerVariables("SCRIPT_NAME")
		nP = InStrRev(strURL, "/")
		If nP > 0 Then
		  strURL = Left(strURL, nP)
		End If
		GetURLPath = strURL
	End Function
%>