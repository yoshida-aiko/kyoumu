<%
	On Error Resume Next
    Err.Clear
    
' �I�u�W�F�N�g��`   
	Set g_cCn = Server.CreateObject("ADODB.Connection")
	Set g_rRs = Server.CreateObject("ADODB.Recordset")
	Set g_file = Server.CreateObject("Scripting.fileSystemObject")	'�t�@�C���V�X�e���I�u�W�F�N�g�̍쐬

    g_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    g_rRs.Open "M_�Ј�",g_cCn,2,2

	Set g_rRs = g_cCn.Execute(Session.Contents("g_CSV"))
	
' SQL���s���̃G���[����
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	
	On Error Goto 0
	
'************************************************************************************************
'											CSV�o�͏���
'************************************************************************************************
	FileName="\\WEBSVR_2\infogram\hirota\No02\Sample.csv"
	'FileName=Server.MapPath("Sample.csv") C:\infogram\hirota\No02\Sample.csv 
	
	Set fs_test = g_file.CreateTextFile(FileName, True)	'�t�@�C���̍쐬
	w_Index = 0
	
' CSV�t�@�C���ɏ�����
	Do while not g_rRs.EOF
	w_Index = w_Index + 1
		fs_test.WriteLine(g_rRs("�Ј�CD") & "," & g_rRs("�Ј�����") & "," & g_rRs("���N����") _
				& "," & g_rRs("�d�b�ԍ�1") & "," & g_rRs("�d�b�ԍ�2")	 & "," & g_rRs("�X��") & "," _
							& g_rRs("�Z��1") & "," & g_rRs("�Z��2") & "," & g_rRs("���l"))	'�P�s���C�g
		g_rRs.MoveNext
	Loop
	fs_test.Close						'�t�@�C���̃N���[�Y

' �o�̓p�X�ƃ��R�[�h�J�E���g�̑��M
	Session.Contents("Path")=FileName
	Response.Redirect "FinishCSV.asp?Count=" & w_Index

' �I�u�W�F�N�g�̊J��
    w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
	
%>