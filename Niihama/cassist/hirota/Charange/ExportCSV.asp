<%	
	On Error Resume Next
    Err.Clear

	Dim w_cCn,w_rRs,w_SQL,w_Index
	Dim w_StartCD,w_EndCD,w_Name,w_CheckDel

	w_StartCD = Request.Form("txtStartCD")
	w_EndCD = Request.Form("txtEndCD")
	w_Name = Request.Form("txtName")
	w_CheckDel = Request.Form("checkDel")
	w_SQL = Request.Form("SQL")
	
'--------------------�S�p�𔼊p�ɕϊ�----------------------------------
	Set bobj = Server.CreateObject("basp21")
	w_StartCD = bobj.StrConv(w_StartCD,8)	'�S�p�����p�ϊ�
	
' �I�u�W�F�N�g�̒�`   
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")
	
    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_�Ј�",w_cCn,2,2
    
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
'************************************************************************************************
'											CSV�o�͏���
'************************************************************************************************
	FileName="\\WEBSVR_2\infogram\hirota\No02\Sample.csv"
   'FileName=Server.MapPath("Sample.csv") C:\infogram\hirota\No02\Sample.csv 
	Set g_file = Server.CreateObject("Scripting.fileSystemObject")	'�t�@�C���V�X�e���I�u�W�F�N�g�̍쐬

	
	Set f_ExportFile = g_file.CreateTextFile(FileName, True)	'�t�@�C���̍쐬
	w_Index = 0
	
' CSV�t�@�C���ɏ�����
	Do while not w_rRs.EOF
	w_Index = w_Index + 1
		f_ExportFile.WriteLine(w_rRs("�Ј�CD") & "," & w_rRs("�Ј�����") & "," & w_rRs("���N����") _
				& "," & w_rRs("�d�b�ԍ�1") & "," & w_rRs("�d�b�ԍ�2")	 & "," & w_rRs("�X��") & "," _
							& w_rRs("�Z��1") & "," & w_rRs("�Z��2") & "," & w_rRs("���l"))	'�P�s���C�g
		w_rRs.MoveNext
	Loop
	f_ExportFile.Close						'�t�@�C���̃N���[�Y

' �o�̓p�X�ƃ��R�[�h�J�E���g�̑��M
	Session.Contents("Path")=FileName
	Response.Redirect "FinishCSV.asp?Count=" & w_Index

	
    w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
%>