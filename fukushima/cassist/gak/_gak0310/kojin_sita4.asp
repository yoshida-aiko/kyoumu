<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍��ڍ�
' ��۸���ID : gak/gak0300/kojin_sita4.asp
' �@      �\: �������ꂽ�w���̏ڍׂ�\������(���l�E����)
'-------------------------------------------------------------------------
' ��      ��	Session("GAKUSEI_NO")  = �w���ԍ�
'            	Session("HyoujiNendo") = �\���N�x
'           
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
    Public m_bErrFlg        '�װ�׸�
	Public m_Rs				'ں��޾�ĵ�޼ު��
	Public m_HyoujiFlg		'�\���׸�

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
		w_iRet = f_GetDetailBikou()
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
Function f_GetDetailBikou()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailBikou = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T11_SYUMITOKUGI,  "
		w_sSql = w_sSql & " 	A.T11_SOGOSYOKEN,  "
		w_sSql = w_sSql & " 	A.T11_KODOSYOKEN,  "
		w_sSql = w_sSql & " 	A.T11_KOJIN_BIK "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T11_GAKUSEKI A "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & "  	A.T11_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "

		iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailBikou = 99
			Exit Do
		End If

		'//����I��
		f_GetDetailBikou = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  ���R�I�����ڂ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDetailFreeData(ByRef p_Rs)
	Dim w_iRet2
	Dim w_sSQL2

	On Error Resume Next
	Err.Clear

	f_GetDetailFreeData = 1

	Do

		w_sSql2 = w_sSql2 & ""
		w_sSql2 = w_sSql2 & " SELECT "
		w_sSql2 = w_sSql2 & "		A.M58_JIYU_MEI,	"
		w_sSql2 = w_sSql2 & "		A.M58_JIYUBUNRUI_MEI, "
		w_sSql2 = w_sSql2 & "		A.M58_NENDO, "
		w_sSql2 = w_sSql2 & "		B.T74_GAKUSEI_NO, "
		w_sSql2 = w_sSql2 & "		B.T74_YOBI1, "
		w_sSql2 = w_sSql2 & "		B.T74_YOBI2 "
		w_sSql2 = w_sSql2 & " FROM "
		w_sSql2 = w_sSql2 & "		MM58_JIYU_JYOHOU A , "
		w_sSql2 = w_sSql2 & "		TT74_JIYU_JYOHOU B "
		w_sSql2 = w_sSql2 & " WHERE "
		w_sSql2 = w_sSql2 & "		A.M58_NENDO = " & Session("HyoujiNendo")
		w_sSql2 = w_sSql2 & " AND "
		w_sSql2 = w_sSql2 & " 	B.T74_NENDO = A.M58_NENDO "
		w_sSql2 = w_sSql2 & " AND "
		w_sSql2 = w_sSql2 & "		A.M58_JIYUBUNRUI_CD = B.T74_JIYUBUNRUI_CD "
		w_sSql2 = w_sSql2 & " AND "
		w_sSql2 = w_sSql2 & "		A.M58_JIYU_CD = B.T74_JIYU_CD "
		w_sSql2 = w_sSql2 & " AND "
		w_sSql2 = w_sSql2 & "  B.T74_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		w_sSql2 = w_sSql2 & " AND "
		w_sSql2 = w_sSql2 & "  A.M58_JIYU_TYPE = " & C_JIYU_USE_YES

		iRet2 = gf_GetRecordset(p_Rs, w_sSQL2)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailBikou = 99
			Exit Do
		End If

		'//����I��
		f_GetDetailFreeData = 0
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

	Dim w_Rs '���R�[�h�Z�b�g

	On Error Resume Next
	Err.Clear

	'// �ϐ�������
	m_HyoujiFlg = 0 		'<!-- �\���t���O�i0:�Ȃ�  1:����j

	m_SYUMITOKUGI = ""
	m_SOGOSYOKEN  = ""
	m_KODOSYOKEN  = ""
	m_KOJIN_BIK   = ""

	if Not m_Rs.EOF then
		m_SYUMITOKUGI = m_Rs("T11_SYUMITOKUGI")
		m_SOGOSYOKEN  = m_Rs("T11_SOGOSYOKEN")
		m_KODOSYOKEN  = m_Rs("T11_KODOSYOKEN")
		m_KOJIN_BIK   = m_Rs("T11_KOJIN_BIK")
	End if

	'// ���R�I�����ڂ��擾
	Call f_GetDetailFreeData(w_Rs)


%>

	<html>
	<head>
	<title>�w�Ѓf�[�^�Q��</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<style type="text/css">
	<!--
		a:link { color:#cc8866; text-decoration:none; }
		a:visited { color:#cc8866; text-decoration:none; }
		a:active { color:#888866; text-decoration:none; }
		a:hover { color:#888866; text-decoration:underline; }
		b { color:#88bbbb; font-weight: bold; font-size:14px}
	//-->
	</style>
	<script language="javascript">
	<!--
		function sbmt(m,i) {
			document.forms[0].mode.value = m;
			document.forms[0].id.value = i;
			document.forms[0].submit();
		}
	//-->
	</script>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">����{���</a></td>
			<td nowrap><a href="kojin_sita1.asp">���l���</a></td>
			<td nowrap><a href="kojin_sita2.asp">�����w���</a></td>
			<td nowrap><a href="kojin_sita3.asp">���w�N���</a></td>
			<td nowrap><b>�����̑��\�����</b></td>
			<td nowrap><a href="kojin_sita5.asp">���ٓ����</a></td>
		</tr>
	</table>
	<br>
				<% if gf_empItem(C_T_JIYUSENTAKU) then %>
					<table class="hyo" border="1" width="600">
						<tr>
							<th class="header" width="150" height="16">����</th>
							<th class="header" width="100" height="16">��</th>
							<th class="header" width="125" height="16">���l</th>
							<th class="header" width="125" height="16">���l2</th>
						</tr>
					<%
							'// ���R�񐔕�,��
								Do Until w_Rs.Eof

									'// ���R���ږ��y�уf�[�^���擾
									Call gs_cellPtn(w_cell)
								%>
									<tr>
										<%' if gf_empItem(C_T_JIYUSENTAKU) then %>
											<td class="<%=w_cell%>"><%= w_Rs("M58_JIYU_MEI") %></td>
											<td class="<%=w_cell%>"><%= w_Rs("M58_JIYUBUNRUI_MEI") %></td>
											<td class="<%=w_cell%>"><%= w_Rs("T74_YOBI1") %></td>
											<td class="<%=w_cell%>"><%= w_Rs("T74_YOBI2") %></td>
										<%' End if %>
									</tr>
								<%
									w_Rs.MoveNext
								Loop
					 %>
					</table>
				<% End if %>

	<% if m_HyoujiFlg = 0 then %>
		<BR>
		�\���ł���f�[�^������܂���<BR>
		<BR>
	<% End if %>

	<BR>
	<input type="button" class="button" value="�@����@" onClick="parent.window.close();">

	</div>
	</form>
	</body>
	</html>
<% End Sub %>