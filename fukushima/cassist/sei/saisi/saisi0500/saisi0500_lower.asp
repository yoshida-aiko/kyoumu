<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��:
' ��۸���ID :
' �@      �\:
'-------------------------------------------------------------------------
' ��      ��:
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2003/02/24 hirota
'*************************************************************************/

%>
<!--#include file="../../Common/com_All.asp"-->
<%

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

	Public gRs					'//���R�[�h
	Public msURL
	Public m_bErrFlg
	Public m_sLoad

	Public m_iGakunen			'//�w�N
	Public m_iClassNo			'//�N���XNO
	Public m_iSyoriNen          '//�N�x
	Public m_iKyokanCd          '//��������
	Public m_iGakka				'//�w��
	Public m_sHyoka()			'//�]���L���z��
	Public m_sClass				'//�N���X
	Public m_sClassNM			'//�N���X��

'///////////////////////////���C������/////////////////////////////

	'Ҳ�ٰ�ݎ��s
	Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()

	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    On Error Resume Next
    Err.Clear

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�s���i�w���ꗗ"
    w_sMsg=""
    w_sRetURL = C_RetURL & C_ERR_RETURL
    w_sTarget = "fTopMain"

    m_bErrFlg = False

    Do
		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'// �p�����[�^�擾
		Call s_GetParameter()

		'// �\���{�^��������
		If m_sLoad = "load" then

			'// �ް��ް��ڑ�
			If gf_OpenDatabase() <> 0 Then
				'�ް��ް��Ƃ̐ڑ��Ɏ��s
				m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
				Exit Do
			End If

			'// �]���L���擾
			If Not f_GetHyoka() then
				m_sErrMsg = "�]���`���擾�Ɏ��s���܂����B"
				Exit Do
			End If

			'// �N���X�f�[�^�擾
			If Not f_GetClassData() then
				m_sErrMsg = "�N���X�f�[�^�擾�Ɏ��s���܂����B"
				Exit Do
			End If

		End If

		'// �y�[�W��\��
		Call showPage()

		m_bErrFlg = True
        Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If Not m_bErrFlg Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
    End If

	'// �I������
    Call gf_closeObject(gRs)
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[�@�\]	�p�����[�^�擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_GetParameter()

	m_sLoad     = Request("mode")
	m_iSyoriNen = Session("NENDO")
	m_sClass    = Request("hidClass")
	m_iGakunen  = Request("hidGakunen")
	m_sClassNM  = Request("hidClassNM")

End Sub

'********************************************************************************
'*	[�@�\]	�]���L���擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetHyoka()

	Dim w_sSQL
	Dim w_lRecCnt
	Dim w_iRet
	Dim wRs

	On Error Resume Next
	Err.Clear

	f_GetHyoka = False

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD, "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI_R "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & " 	AND M01_DAIBUNRUI_CD = " & C_HYOKA_FUKA
	w_sSQL = w_sSQL & " ORDER BY "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD "

	w_iRet = gf_GetRecordset(wRs, w_sSQL)

	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit Function
	End If

	If Not wRs.EOF then
		w_lRecCnt = wRs.RecordCount											'//���R�[�h�J�E���g
		wRs.MoveFirst

		'//�]���L���z��Z�b�g
		Do While Not wRs.EOF
			Redim Preserve m_sHyoka(wRs("M01_SYOBUNRUI_CD"))							'//�]���L���z���`
			m_sHyoka(wRs("M01_SYOBUNRUI_CD")) = wRs("M01_SYOBUNRUIMEI_R")	'//�]���L��
			wRs.MoveNext
		Loop
	End If

	'//���R�[�h���
    Call gf_closeObject(wRs)

	f_GetHyoka = True

End Function

'********************************************************************************
'*	[�@�\]	�N���X�f�[�^�擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetClassData()

	Dim w_sSQL
	Dim w_iRet

	On Error Resume Next
	Err.Clear

	f_GetClassData = False

	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T11.T11_SIMEI, "
	w_sSQL = w_sSQL & " 	T11.T11_GAKUSEI_NO, "
	w_sSQL = w_sSQL & " 	T13.T13_SYUSEKI_NO1, "
	w_sSQL = w_sSQL & " 	T13.T13_GAKUNEN , "
	w_sSQL = w_sSQL & " 	T13.T13_CLASS , "
	w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO , "
	w_sSQL = w_sSQL & " 	T120.*, "
	w_sSQL = w_sSQL & " 	M04.M04_KYOKANMEI_SEI, "
	w_sSQL = w_sSQL & " 	M04.M04_KYOKANMEI_MEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI T11, "
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN T13, "
	w_sSQL = w_sSQL & " 	T120_SAISIKEN T120, "
	w_sSQL = w_sSQL & " 	M04_KYOKAN M04 "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	T13.T13_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & " 	AND T13.T13_CLASS        = '" & m_sClass & "'"
	w_sSQL = w_sSQL & " 	AND T13.T13_GAKUNEN      =  " & m_iGakunen
	w_sSQL = w_sSQL & " 	AND T13.T13_GAKUSEI_NO   = T120.T120_GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND T11.T11_GAKUSEI_NO   = T120.T120_GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND T120.T120_NENDO      = M04.M04_NENDO "
	w_sSQL = w_sSQL & " 	AND T120.T120_KYOUKAN_CD = M04.M04_KYOKAN_CD(+) "
	w_sSQL = w_sSQL & " ORDER BY "
	w_sSQL = w_sSQL & " 	T13.T13_GAKUNEN, "
	w_sSQL = w_sSQL & " 	T13.T13_CLASS, "
	w_sSQL = w_sSQL & " 	T13.T13_SYUSEKI_NO1, "
	w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO, "
	w_sSQL = w_sSQL & " 	T120.T120_NENDO DESC, "
	w_sSQL = w_sSQL & " 	T120.T120_KAMOKU_CD "

	w_iRet = gf_GetRecordset(gRs, w_sSQL)

	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit Function
	End If

	f_GetClassData = True

End Function

'********************************************************************************
'*	[�@�\]	HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear
	'---------- HTML START ----------
%>
<html>
<head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>�s���i�w���ꗗ</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
    //************************************************************
    //  [�@�\]  �t�H�[�����[�h��
    //  [����]  
    //  [�ߒl]  
    //  [����]
    //************************************************************
	function jf_windowload(){
		<% If m_sLoad = "load" then %>
			<% If gRs.EOF then %>
				alert("�Ώۃf�[�^�͑��݂��܂���B");
				parent._TOP.document.body.style.cursor = "default";
				return;
			<% End If %>
			with(document.frm){
				target = "_TOP";
				action = "saisi0500_head.asp";
				submit();
			}
		<% End If %>
	}
	window.onload = jf_windowload;
	//-->
	</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<center>

<form name="frm" action="" target="main" Method="POST">

<%
If m_sLoad = "load" then

	Dim w_Class
	Dim w_sSeiseki
	Dim w_sSaisi

	w_Class = ""

	Do While Not gRs.EOF

		'�e�[�u���w�i�F�ݒ�
		Call gs_cellPtn(w_Class)

		'�]���L�� + ���]��
		w_sSeiseki = m_sHyoka(gRs("T120_OLD_HYOKA_FUKA_KBN")) & gRs("T120_SEISEKI")
		'w_sSaisi   = m_sHyoka(gRs("T120_HYOKA_FUKA_KBN")) & gRs("T120_SAISI_SEISEKI")
		w_sSaisi   = gRs("T120_SAISI_SEISEKI")
%>
	<table class="hyo" border="1">
		<tr>
			<td width="60"  class="<%= w_Class %>" nowrap align="right"><%= gRs("T13_SYUSEKI_NO1") %></td>
			<td width="150" class="<%= w_Class %>" nowrap><%= gRs("T11_SIMEI") %></td>
			<td width="150" class="<%= w_Class %>" nowrap><%= gRs("T120_KAMOKUMEI") %></td>
			<td width="50"  class="<%= w_Class %>" nowrap align="center"><%= gRs("T120_NENDO") %></td>
			<td width="70"  class="<%= w_Class %>" nowrap align="right"><%= gRs("T120_KEKASU") & " / " & gRs("T120_JUNJIKAN") %></td>
			<td width="40"  class="<%= w_Class %>" nowrap align="right"><%= w_sSeiseki %></td>
			<td width="40"  class="<%= w_Class %>" nowrap align="right"><%= w_sSaisi %></td>
			<td width="100" class="<%= w_Class %>" nowrap><%= gRs("M04_KYOKANMEI_SEI") & " " & gRs("M04_KYOKANMEI_MEI") %></td>
		</tr>
	</table>
<%
		gRs.MoveNext
	Loop
Else
%>

	<br><br><br>
	<CENTER><span class="msg">���@�\���{�^���������Ă������� </span></CENTER>

<% End If %>
<input type="hidden" name="hidClass" value="<%= m_sClass %>">
<input type="hidden" name="hidGakunen" value="<%= m_iGakunen %>">
<input type="hidden" name="hidClassNM" value="<%= m_sClassNM %>">
</form>

</center>

</body>
</html>
<%
'---------- HTML END   ----------
End Sub
%>