<!--#include file="adovbs.inc"-->
<%
' ***** �`�r�o�p�c�a���ʊ֐��Q�@����� �Ɠ� *****


'Const MailSvName = "smtp.nifty.com"
Const MailSvName = "arbby.arbby.co.jp"
Const BaspMailLog = "d:\basp\email_log.txt"
Const ArbbyEAdd = "info@arbby.com"


'***** �Z�b�V�����I�[�v�� *****
Function gf_ConnOpen(pCon)
	On Error Resume Next
	Err.Clear
	
	Set pCon = Server.CreateObject("ADODB.Connection")
	pCon.Open "10.0.1.4/cassist", "cassist", "cassist"
	If Err.Number <> 0 Then
		gf_ConnOpen = False
		Exit Function
	End If
	gf_ConnOpen = True
End Function

'***** �Z�b�V�����N���[�Y *****
Function gf_ConnClose(pCon)
	On Error Resume Next
	Err.Clear
	
	pCon.Close
	Set pCon = Nothing
	If Err.Number <> 0 Then
		gf_ConnClose = False
		Exit Function
	End If
	gf_ConnClose = True
End Function

'***** ���R�[�h�Z�b�g�I�[�v�� *****
Function gf_RSOpen(pRS, pCon, pSql, pCType, pLType)
	On Error Resume Next
	Err.Clear

	Set pRS = Server.CreateObject("ADODB.Recordset")
	pRS.Open pSql, pCon, pCType, pLType
	If Err.Number <> 0 Then
		gf_RSOpen = False
		Exit Function
	End If
	gf_RSOpen = True
End Function

'***** ���R�[�h�Z�b�g�N���[�Y *****
Function gf_RSClose(pRS)
	On Error Resume Next
	Err.Clear
	
	pRS.Close
	Set pRS = Nothing
	If Err.Number <> 0 Then
		gf_RSClose = False
		Exit Function
	End If
	gf_RSClose = True
End Function


'***** �n�k�d�Z�b�V�����I�[�v�� *****
Function gf_ConnOpenOLE(oraSess, pCon)
	On Error Resume Next
	Err.Clear

	Set oraSess = CreateObject("OracleInProcServer.XOraSession")
	Set pCon = oraSess.DbOpenDatabase("CASSIST", "CASSIST/CASSIST", clng(3))
	
	If Err.Number <> 0 Then
		gf_ConnOpenOLE = False
		Exit Function
	End If
	gf_ConnOpenOLE = True

	Set orsSess = Nothing
End Function

'***** �n�k�d�Z�b�V�����N���[�Y *****
Function gf_ConnCloseOLE(oraSess, pCon)
	On Error Resume Next
	Err.Clear

	pCon.Close
	'oraSess.Close

	Set pCon = Nothing
	Set oraSess = Nothing

	If Err.Number <> 0 Then
		gf_ConnCloseOLE = False
		Exit Function
	End If
	gf_ConnClosedOLE = True
End Function

'***** �n�k�d���R�[�h�Z�b�g�I�[�v�� *****
Function gf_RSOpenOLE(pRS, pCon, tSql)
	On Error Resume Next
	Err.Clear

	Set pRS = pCon.DbCreateDynaset(tSql,clng(0))

	If Err.Number <> 0 Then
		gf_RSOpenOLE = False
		Exit Function
	End If
	gf_RSOpenOLE = True
End Function

'***** �n�k�d���R�[�h�Z�b�g�N���[�Y *****
Function gf_RSCloseOLE(pRS)
	On Error Resume Next
	Err.Clear

	pRS.Close

	Set pRS = Nothing
	
	If Err.Number <> 0 Then
		gf_RSCloseOLE = False
		Exit Function
	End If
	gf_RSClosedOLE = True
End Function


'***** �a�`�r�o�Q�P���[�����M *****
Function gf_BaspEmail(mto, mfrom, subj, body)
	Set basp = Server.CreateObject("basp21")
	
	gf_BaspEmail = basp.SendMailEx(BaspMailLog, MailSvName, mto, mfrom, subj, body, "")
End Function

%>
