<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍��ڍ�
' ��۸���ID : gak/gak0300/kojin_ue.asp
' �@      �\: �������ꂽ�w���̏ڍׂ�\������
'-------------------------------------------------------------------------
' ��      �� 	Session("GAKUSEI_NO")  = �w���ԍ�
'            	Session("HyoujiNendo") = �\���N�x
' ��      ��
' ��      �n
'           
'           
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public m_bErrFlg		'�װ�׸�
	Public m_Rs				'ں��޾�ĵ�޼ު��

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

Sub Main()
'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�w����񌟍�����"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'//�\�����ڂ��擾
		w_iRet = f_GetDetail()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//�����\��
        if m_TxtMode = "" then
            Call showPage()
            Exit Do
        end if

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// �I������
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �\�����ڂ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:����I��	1:�C�ӂ̃G���[  99:�V�X�e���G���[
'*  [����]  
'********************************************************************************
Function f_GetDetail()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetail = 1

	Do
		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO, "
		w_sSql = w_sSql & " 	A.T13_GAKUSEKI_NO,  "
		w_sSql = w_sSql & " 	A.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	A.T13_CLASS,  "
		w_sSql = w_sSql & " 	B.T11_SIMEI "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A, "
		w_sSql = w_sSql & " 	T11_GAKUSEKI B "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO(+) AND "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' AND "
		w_sSql = w_sSql & " 	A.T13_NENDO		 =  " & Session("HyoujiNendo")

		iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetail = 99
			Exit Do
		End If

		'//����I��
		f_GetDetail = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()


	m_GAKUSEI_NO  = ""
	m_GAKUSEKI_NO = ""
	m_GAKUNEN     = ""
	m_CLASS       = ""
	m_SIMEI       = ""

	if Not m_Rs.Eof then
		m_GAKUSEI_NO  = m_Rs("T13_GAKUSEI_NO")
		m_GAKUSEKI_NO = m_Rs("T13_GAKUSEKI_NO")
		m_GAKUNEN     = m_Rs("T13_GAKUNEN")
		m_CLASS       = m_Rs("T13_CLASS")
		m_SIMEI       = m_Rs("T11_SIMEI")
	End if

%>
	<html>
	<head>
	<title>�w�Ѓf�[�^�Q��</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<%call gs_title("�w����񌟍�","�ڍ�")%>

	<br>
	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td>

				<table border="1" class="disp">
					<tr>
						<% if gf_empItem(C_T13_GAKUSEI_NO) then %>
							<td class="disph" nowrap width="100" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_5NEN)%></td>
							<td class="disp" nowrap width="80"><%= m_GAKUSEI_NO %>&nbsp;</td>
						<% End if %>
						<% if gf_empItem(C_T13_GAKUSEKI_NO) then %>
							<td class="disph" nowrap width="100" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_1NEN)%></td>
							<td class="disp" nowrap width="80"><%= m_GAKUSEKI_NO %>&nbsp;</td>
						<% End if %>
					</tr>
				</table>

			</td>
		</tr>
		<tr>
			<td>

				<table border="1" class="disp">
					<tr>
						<% if gf_empItem(C_T13_GAKUNEN) then %>
							<td class="disph" nowrap width="100" height="16">�w�@�@�N</td>
							<td class="disp" nowrap width="80"><%= m_GAKUNEN %>&nbsp;</td>
						<% End if %>
						<% if gf_empItem(C_T13_CLASS) then %>
							<td class="disph" nowrap width="100" height="16">�N �� �X</td>
							<td class="disp" nowrap width="80"><%= m_CLASS %>�g&nbsp;</td>
						<% End if %>
						<% if gf_empItem(C_T11_SIMEI) then %>
							<td class="disph" nowrap width="100" height="16">���@�@��</td>
							<td class="disp" nowrap><%= m_SIMEI %>&nbsp;</td>
						<% End if %>
					</tr>
				</table>

			</td>
		</tr>
	</table>

	</div>
	</form>
	</body>
	</html>
<% End Sub %>