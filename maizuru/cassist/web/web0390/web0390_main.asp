<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���x���ʉȖڌ���
' ��۸���ID : web/web0390/web0390_main.asp
' �@	  �\: ���y�[�W �\������\��
'-------------------------------------------------------------------------
' ��	  ��:�����R�[�h 	��		SESSION("KYOKAN_CD")
'			 �N�x			��		SESSION("NENDO")
' ��	  ��:
' ��	  �n:
' ��	  ��:
'-------------------------------------------------------------------------
' ��	  ��: 2001/10/26 �J�e �ǖ�
' ��	  �X: 2001/12/01 �c�� ��K�@�S�����Ă��Ȃ��Ȗڂ�ύX�ł��Ȃ��悤�ɏC��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public	m_bErrFlg			'�װ�׸�
'///////////////////////////���C������/////////////////////////////
	Dim 	m_iNendo		'//�N�x
	Dim 	m_sKyokanCd 	'//�����R�[�h
	Dim 	m_sGakunen		'//�w�N
	Dim 	m_sClass		'//�N���X
	Dim		m_sKamokuCD		'//�ȖڃR�[�h


	'Ҳ�ٰ�ݎ��s
	Call Main()

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
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���x���ʉȖڌ���"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"

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
		session("PRJ_No") = "WEB0390"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
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

		Call s_GetParam() 		'�n���ꂽ�������擾

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
	Call gs_CloseDatabase()
End Sub

Sub showPage()
'********************************************************************************
'*	[�@�\]	Html���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
%>
<Html>

<head>

<Title>�l���C�I���Ȗڌ���</Title>
<frameset rows="138px,1px,*" frameborder="no">
	<frame src="white.asp?txtMsg=<%=Request("txtMsg")%>" scrolling="yes" noresize name="middle">
	<frame src="../../common/bar.html" scrolling="no" noresize name="bar">
	<frame src="web0390_bottom.asp?<%=Server.htmlEncode(request.form.item)%>" scrolling="yes" noresize name="bottom">
</frameset>
</head>

</Html>
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

	m_sKyokanCd = session("KYOKAN_CD")
	m_sGakunen = request("txtGakunen")
	m_sClass = request("txtClass")
	m_sKamokuCD = request("cboKamokuCode")

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

	f_chkTantoKyokan = false

	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "     T27_KYOKAN_CD"
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "     T27_TANTO_KYOKAN T27"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "     T27_NENDO = " & SESSION("NENDO") & " "
	w_sSQL = w_sSQL & vbCrLf & " AND "
	w_sSQL = w_sSQL & vbCrLf & " T27_GAKUNEN = " & request("txtGakunen") & " "
	w_sSQL = w_sSQL & vbCrLf & " AND "
	w_sSQL = w_sSQL & vbCrLf & " T27_KYOKAN_CD = '" & m_sKyokanCd & "' "
	w_sSQL = w_sSQL & vbCrLf & " AND "
	w_sSQL = w_sSQL & vbCrLf & " T27_KAMOKU_CD = '" & request("cboKamokuCode") & "' "
	w_sSQL = w_sSQL & vbCrLf & " GROUP BY T27_KYOKAN_CD"
	w_sSQL = w_sSQL & vbCrLf & " ORDER BY T27_KYOKAN_CD"

	Set w_oRecord = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset_OpenStatic(w_oRecord, w_sSQL)

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
		<span class="msg">�S�����Ă���Ȗڂł͂���܂���B</span>
	</center>
</Body>
</Html>
<%
End Sub

'2001/12/01 Add <----

%>