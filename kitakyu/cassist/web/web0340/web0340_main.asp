<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �l���C�I���Ȗڌ���
' ��۸���ID : web/web0340/web0340_main.asp
' �@	  �\: ���y�[�W �\������\��
'-------------------------------------------------------------------------
' ��	  ��:�����R�[�h 	��		SESSION("KYOKAN_CD")
'			 �N�x			��		SESSION("NENDO")
' ��	  ��:
' ��	  �n:
' ��	  ��:
'-------------------------------------------------------------------------
' ��	  ��: 2001/07/25 �O�c
' ��	  �X: 2001/08/28 �ɓ����q �w�b�_���؂藣���Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public m_bErrFlg			'�װ�׸�

	Dim 	m_sKyokanCd 	'//�����R�[�h
	Dim 	m_sGakunen		'//�w�N
	Dim 	m_sClass		'//�N���X
	Dim 	m_sKBN			'//�敪
	Dim 	m_sGRP			'//�O���[�v�敪

'///////////////////////////���C������/////////////////////////////

	'Ҳ�ٰ�ݎ��s
	Call Main

'///////////////////////////�@�d�m�c�@/////////////////////////////

Sub Main()
'********************************************************************************
'*	[�@�\]	�{ASP��Ҳ�ٰ��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]
'********************************************************************************

	Dim w_iRet				'// �߂�l
	Dim w_sSQL				'// SQL��
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle = "�L�����p�X�A�V�X�g"
	w_sMsgTitle = "�l���C�I���Ȗڌ���"
	w_sMsg = ""
	w_sRetURL = "../../login/default.asp"
	w_sTarget = "_top"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// �ް��ް��ڑ�
		w_iRet = gf_OpenDatabase()
		If w_iRet <> 0 Then
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'// �����`�F�b�N�Ɏg�p
		SESSION("PRJ_No") = "WEB0340"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(SESSION("PRJ_No"))
'
'		 '// �S�C�`�F�b�N
'		If gf_Tannin(session("NENDO"),session("KYOKAN_CD"),1) <> 0 Then
'			m_bErrFlg = True
'			m_sErrMsg = "�S�C�ȊO�̓��͂͂ł��܂���B"
'			Exit Do
'		End If

		'2001/12/01 Modd ---->
'		'// �y�[�W��\��
'		Call showPage()

		Call s_GetParam 		'�n���ꂽ�������擾

		'�S�����Ă��邩�ǂ������`�F�b�N
		If f_chkTantoKyokan = True Then
			'// �S�����Ă���ꍇ�A�ڍ׃y�[�W��\��
			Call showPage
		Else
			'// �S�����Ă��Ȃ��ꍇ�A�G���[�y�[�W��\��
			Call showErrPage
		End If
		'2001/12/01 Add <----

		Exit Do
	Loop

	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// �I������
	Call gs_CloseDatabase
End Sub

Sub showPage()
'********************************************************************************
'*	[�@�\]	HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]
'********************************************************************************
%>
<html>

<head>
<title>�l���C�I���Ȗڌ���</title>
</head>

<frameset rows="138px,1px,*" frameborder="no">
	<frame src="white.asp?txtMsg=<%=Request("txtMsg")%>" scrolling="yes" noresize name="middle">
	<frame src="../../common/bar.html" scrolling="no" noresize name="bar">
	<frame src="web0340_bottom.asp?<%=Server.HTMLEncode(request.form.item)%>" scrolling="yes" noresize name="bottom">
</frameset>

</html>
<%
End Sub

'2001/12/01 Add ---->

Sub s_GetParam()
'********************************************************************************
'*	[�@�\]�@�p�����[�^�擾
'*	[����]�@ �Ȃ�
'*	[�ߒl]�@�Ȃ�
'*	[����]
'********************************************************************************

	m_sKyokanCd = SESSION("KYOKAN_CD")
	m_sGakunen = request("txtGakunen")
	m_sClass = request("txtClass")
	m_sKBN = Cint(request("txtKBN"))
	m_sGRP = Cint(request("txtGRP"))

End Sub

Function f_chkTantoKyokan()
'********************************************************************************
'*	[�@�\]�@�S���������ǂ����`�F�b�N
'*	[����]�@ �Ȃ�
'*	[�ߒl]�@True:�S�����������Ă���AFalse:�S�����������Ă��Ȃ�
'*	[����]
'********************************************************************************
	Dim w_iRet			'�߂�l
	Dim w_sSQL			'SQL
	Dim w_oRecord		'���R�[�h

	f_chkTantoKyokan = False

	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "	T27_KYOKAN_CD"
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "	T27_TANTO_KYOKAN,"
	w_sSQL = w_sSQL & vbCrLf & "	T16_RISYU_KOJIN "
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "	T27_NENDO      = T16_NENDO AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_KAMOKU_CD  = T16_KAMOKU_CD AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_GAKUNEN    = T16_HAITOGAKUNEN AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_NENDO      = " & SESSION("NENDO") & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_GAKUNEN    = " & m_sGakunen & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_CLASS      = " & m_sClass & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_KYOKAN_CD  = '" & m_sKyokanCd & "' AND "
	w_sSQL = w_sSQL & vbCrLf & "	T16_HISSEN_KBN = " & C_HISSEN_SEN & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T16_SELECT_FLG = " & C_SENTAKU_YES & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T16_KAMOKU_KBN = " & m_sKBN & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T16_GRP        = " & m_sGRP

	Set w_oRecord = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset_OpenStatic(w_oRecord, w_sSQL)
'response.write(w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit Function
	End If

	'//�S�����Ă��Ȃ��ꍇ
	If w_oRecord.EOF = True Then
		Exit Function
	End If

	w_oRecord.Close
	Set w_oRecord = Nothing

	f_chkTantoKyokan = True

End Function

Sub showErrPage()
'********************************************************************************
'*	[�@�\]	Html���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]
'********************************************************************************
%>
<Html>
<head>
<Title>�l���C�I���Ȗڌ���G���[�y�[�W</Title>
<link rel=stylesheet href=../../common/style.css type=text/css>
</head>
<Body>
	<center>
		<br><br><br>
		<span class="msg">�S�����Ă���Ȗڂ͂���܂���B</span>
	</center>
</Body>
</Html>
<%
End Sub

'2001/12/01 Add <----

%>

