<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍��ڍ�
' ��۸���ID : gak/gak0310/kojin_sita0.asp
' �@      �\: �������ꂽ�w���̏ڍׂ�\������(��{���)
'-------------------------------------------------------------------------
' ��      ��	Session("GAKUSEI_NO")  = �w���ԍ�
'            	Session("Nendo") = �\���N�x
'
' ��      ��
' ��      �n
'
'
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/12/01 ���c
' ��      �X: '�c��̓R���X�g�ύX��OK�ł��B�I2001.12.01 Add ���c
' ��      �X: 2011/04/05 iwata �w���ʐ^�f�[�^���@Session����łȂ��A�f�[�^�x�[�X����擾����B
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public m_bErrFlg        '�װ�׸�
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

'2011.04.05 ins
    '// �摜�f�[�^�擾�p oo4o �Z�b�V�����쐬
   'Del 20170515
   ' Set Session("OraDatabasePh") = OraSession.GetDatabaseFromPool(100)

		'//�\�����ڂ��擾
		w_iRet = f_GetDetailKihon()
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

'2011.04.05 ins
	'** oo4o �ڑ��v�[���p��
    'Del 20170515
	'   Session("OraDatabasePh").DestroyDatabasePool

End Sub

'********************************************************************************
'*  [�@�\]  �\�����ڂ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:����I��	1:�C�ӂ̃G���[  99:�V�X�e���G���[
'*  [����]
'********************************************************************************
Function f_GetDetailKihon()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailKihon = 1

	Do
		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO,  "	'�w���ԍ�
		w_sSql = w_sSql & " 	B.T11_SIMEI,  "			'����
		w_sSql = w_sSql & " 	B.T11_SIMEI_KD, " 		'�����J�i
		w_sSql = w_sSql & " 	B.T11_SIMEI_ROMA,  "	'�������[�}��
		w_sSql = w_sSql & " 	B.T11_SIMEI_KYU,  "		'������
		w_sSql = w_sSql & " 	B.T11_SIMEI_KD_KYU, " 	'�������J�i
		w_sSql = w_sSql & " 	B.T11_SIMEI_ROMA_KYU,  "	'���������[�}��
		w_sSql = w_sSql & "		B.T11_KAIMEI_DATE,	"	'�ŏI��������
		w_sSql = w_sSql & " 	B.T11_HON_ZIP,  "		'�{�ЗX�֔ԍ�
		w_sSql = w_sSql & " 	B.T11_HON_JUSYO,  "		'�{�ЏZ��
		w_sSql = w_sSql & " 	B.T11_GEN_ZIP,  "		'���Z���X�֔ԍ�
		w_sSql = w_sSql & " 	B.T11_GEN_JUSYO,  "		'���Z��
		w_sSql = w_sSql & " 	B.T11_GEN_TEL,  "		'���Z���d�b�ԍ�

		w_sSql = w_sSql & " 	B.T11_GEN_SITYOSON_CD,  "		'���s�����R�[�h
		w_sSql = w_sSql & " 	B.T11_HON_SITYOSON_CD  "		'�{�Ўs�����R�[�h

'/***BLOB�^�Ή��̈� �R�����g		w_sSql = w_sSql & " 	D.T09_IMAGE "			'�ʐ^
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A, "
		w_sSql = w_sSql & " 	T11_GAKUSEKI B, "
		w_sSql = w_sSql & " 	M02_GAKKA    C, "
		w_sSql = w_sSql & " 	T09_GAKU_IMG D, "
		w_sSql = w_sSql & " 	M01_KUBUN E  "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 		A.T13_GAKUSEI_NO   = B.T11_GAKUSEI_NO(+) "
		w_sSql = w_sSql & " 	AND	A.T13_NENDO		   = C.M02_NENDO(+) "
		w_sSql = w_sSql & " 	AND A.T13_GAKKA_CD 	   = C.M02_GAKKA_CD(+) "
		w_sSql = w_sSql & " 	AND A.T13_NENDO		   = E.M01_NENDO "
		w_sSql = w_sSql & " 	AND E.M01_DAIBUNRUI_CD = " & C_ZAISEKI				'�ݐЋ敪
		w_sSql = w_sSql & " 	AND A.T13_ZAISEKI_KBN  = E.M01_SYOBUNRUI_CD(+) "
		w_sSql = w_sSql & " 	AND A.T13_GAKUSEI_NO   = D.T09_GAKUSEI_NO(+) "
		w_sSql = w_sSql & " 	AND A.T13_GAKUSEI_NO   = '" & Session("GAKUSEI_NO") & "'"
		w_sSql = w_sSql & " 	AND A.T13_NENDO 	   =  " & Session("HyoujiNendo")

		iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailKihon = 99
			Exit Do
		End If

		'//����I��
		f_GetDetailKihon = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �ʐ^�����邩���� (BLOB�^�Ή��ׁ̈j
'*  [����]  �Ȃ�
'*  [�ߒl]  True: False
'*  [����]
'********************************************************************************
Function f_Photoimg(pGAKUSEI_NO)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_Photoimg = False

	'// NULL�Ȃ甲����(False)
	if trim(pGAKUSEI_NO) = "" then Exit Function

	Do

	    w_sSQL = ""
	    w_sSQL = w_sSQL & " SELECT "
	    w_sSQL = w_sSQL & " T09_GAKUSEI_NO "
	    w_sSQL = w_sSQL & " FROM T09_GAKU_IMG "
	    w_sSQL = w_sSQL & " WHERE T09_GAKUSEI_NO = '" & cstr(pGAKUSEI_NO) & "'"

		iRet = gf_GetRecordset(w_ImgRs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			Exit Do
		End If
		'// EOF�Ȃ甲����(False)
		if w_ImgRs.Eof then	Exit Do

		'//����I��
		f_Photoimg = True
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

	On Error Resume Next
	Err.Clear

	m_HyoujiFlg = 0 		'<!-- �\���t���O�i0:�Ȃ�  1:����j

	m_GAKUSEI_NO   = ""
	m_SIMEI        = ""
	m_SIMEI_KD     = ""
	m_SIMEI_GAIJI  = ""
	m_SIMEI_ROMA   = ""
	m_SIMEI_KYU    = ""
	m_SIMEI_KD_KYU = ""
	m_SIMEI_ROMA_KYU = ""
	m_KAIMEI_DATE = ""
	m_HON_ZIP      = ""
	m_HON_JUSYO    = ""
	m_GEN_ZIP      = ""
	m_GEN_JUSYO    = ""
	m_GEN_TEL      = ""

	m_Ken = ""
	m_SITYOSONCD = ""
	m_SITYOSONMEI = ""
	m_TYOIKIMEI = ""

	m_Ken2 = ""
	m_SITYOSONCD2 = ""
	m_SITYOSONMEI2 = ""
	m_TYOIKIMEI2 = ""

	m_HONSITYOSONCD = m_Rs("T11_HON_SITYOSON_CD")
	m_GENSITYOSONCD = m_Rs("T11_GEN_SITYOSON_CD")

	Call gf_ComvZip2(m_HONSITYOSONCD,m_Rs("T11_HON_ZIP"),m_Ken,m_SITYOSONCD,m_SITYOSONMEI,m_TYOIKIMEI,Session("HyoujiNendo"))
	Call gf_ComvZip2(m_GENSITYOSONCD,m_Rs("T11_GEN_ZIP"),m_Ken2,m_SITYOSONCD2,m_SITYOSONMEI2,m_TYOIKIMEI2,Session("HyoujiNendo"))

 	if Not m_Rs.EOF then
		m_GAKUSEI_NO   = m_Rs("T13_GAKUSEI_NO")
		m_SIMEI        = m_Rs("T11_SIMEI")
		m_SIMEI_KD     = m_Rs("T11_SIMEI_KD")
		m_SIMEI_ROMA   = m_Rs("T11_SIMEI_ROMA")
		m_SIMEI_KYU    =  m_Rs("T11_SIMEI_KYU")
		m_SIMEI_KD_KYU =  m_Rs("T11_SIMEI_KD_KYU")
		m_SIMEI_ROMA_KYU = m_Rs("T11_SIMEI_ROMA_KYU")
		m_KAIMEI_DATE = m_Rs("T11_KAIMEI_DATE")
		m_HON_ZIP      = m_Rs("T11_HON_ZIP")
		m_HON_JUSYO	=  m_Rs("T11_HON_JUSYO")

		m_HON_JUSYO = Replace(m_HON_JUSYO,m_Ken,"")
		m_HON_JUSYO = Replace(m_HON_JUSYO,m_SITYOSONMEI,"")

		'/* �Z���Ɍ��A�s�����������݂��Ă����ꍇ�폜���čēx�t�������B*/ Add 2002.04.30 okada
		m_HON_JUSYO	=  m_Ken & m_SITYOSONMEI & m_HON_JUSYO     'm_SITYOSONMEI & Replace(m_Rs("T11_HON_JUSYO"),m_SITYOSONMEI,"")
		m_GEN_ZIP      = m_Rs("T11_GEN_ZIP")

		m_GEN_JUSYO    = m_Rs("T11_GEN_JUSYO")
		m_GEN_JUSYO    = Replace(m_GEN_JUSYO,m_Ken2,"")
		m_GEN_JUSYO    = Replace(m_GEN_JUSYO,m_SITYOSONMEI2,"")
		'/* �Z���Ɍ��A�s�����������݂��Ă����ꍇ�폜���čēx�t�������B*/ Add 2002.04.30 okada
		m_GEN_JUSYO    = m_Ken2 & m_SITYOSONMEI2 & m_GEN_JUSYO'm_SITYOSONMEI2 & Replace(m_Rs("T11_GEN_JUSYO"),m_SITYOSONMEI2,"")

		m_GEN_TEL      = m_Rs("T11_GEN_TEL")
	End if

%>
	<html>
	<head>
	<title>�w�Ѓf�[�^�Q��</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
	<style type="text/css">
	<!--
		a:link { color:#cc8866; text-decoration:none; }
		a:visited { color:#cc8866; text-decoration:none; }
		a:active { color:#888866; text-decoration:none; }
		a:hover { color:#888866; text-decoration:underline; }
		b { color:#88bbbb; font-weight: bold; font-size:14px}
	//-->
	</style>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">
	<br><br>
	<div align="center">

	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><b>����{���</b></td>
			<td nowrap><a href="kojin_sita1.asp">���l���</a></td>
			<td nowrap><a href="kojin_sita2.asp">�����w���</a></td>
			<td nowrap><a href="kojin_sita3.asp">���w�N���</a></td>
			<td nowrap><a href="kojin_sita4.asp">�����̑��\�����</a></td>
			<td nowrap><a href="kojin_sita5.asp">���ٓ����</a></td>
		</tr>
	</table>
	<br>

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td valign="top">

				<br>

<% '- ��{��� - %>

				<table class="disp" border="1" width="220">
					<% '- �w���ԍ� - %>
					<% if gf_empItem(C_T13_GAKUSEI_NO) then %>
						<tr>
							<td class="disph" width="100" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_5NEN)%></td>
							<td class="disp"><%= m_GAKUSEI_NO %>&nbsp</td>
						</tr>
					<% End if %>

					<% '- ���� - %>
					<% if gf_empItem(C_T11_SIMEI) then %>
						<tr>
							<td class="disph" height="16">���@�@��</td>
							<td class="disp"><%= m_SIMEI %>&nbsp</td>
						</tr>
					<% End if %>

					<% '- �����J�i - %>
					<% if gf_empItem(C_T11_SIMEI_KD) then %>
						<tr>
							<td class="disph" height="16">�����J�i</td>
							<td class="disp"><%= m_SIMEI_KD %>&nbsp</td>
						</tr>
					<% End if %>

					<% '- �������[�}�� - %>
					<% if gf_empItem(C_T11_SIMEI_ROMA) then %>
						<tr>
							<td class="disph" height="16">�������[�}��</td>
							<td class="disp"><%= m_SIMEI_ROMA %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" rowspan="2">

				<br>
				<table class="disp" border="1" width="230">
					<% '- ������ - %>
					<% if gf_empItem(C_T11_SIMEI_KYU) then %>
						<tr>
							<td class="disph" width="110" height="16">���@���@��</td>
							<td class="disp"><%= m_SIMEI_KYU %>&nbsp</td>
						</tr>
					<% End if %>

					<% '- �������J�i - %>
					<% if gf_empItem(C_T11_SIMEI_KD) then %>
						<tr>
							<td class="disph" width="110" height="16">�������J�i</td>
							<td class="disp"><%= m_SIMEI_KD_KYU %>&nbsp</td>
						</tr>
					<% End if %>

					<% '- ���������[�}�� - %>
					<% if gf_empItem(C_T11_SIMEI_ROMA) then %>
						<tr>
							<td class="disph" width="110" height="16">���������[�}��</td>
							<td class="disp"><%= m_SIMEI_ROMA_KYU %>&nbsp</td>
						</tr>
					<% End if %>

					<% '- �ŏI�������� - %>
					<% if gf_empItem(C_T11_KAIMEI_DATE) then %>
						<tr>
							<td class="disph" width="110" height="16">�ŏI��������</td>
							<td class="disp"><%= m_KAIMEI_DATE %>&nbsp</td>
						</tr>
					<% End if %>
				</table>
				<br>

				<div align="center">
				�y �ʁ@�@�^ �z
				<table border="1" class="disp">
					<tr><td class="disp">
						<%
						'// ��ʐ^�����邩��Ɍ�������
						w_bRet = ""
						w_bRet = f_Photoimg(Session("GAKUSEI_NO"))
						if w_bRet = True then
							' 2011.04.05 upd DispBinary => DispBinaryRec �ɕύX
							%><IMG SRC="DispBinaryRec.asp?gakuNo=<%= Session("GAKUSEI_NO") %>" width="88" height="136" border="0"><%
						Else
							%><IMG SRC="images/Img0000000000.gif" width="100" height="120" border="0"><%
						End if
						%><br>
					</td></tr>
				</table>
				</div>

			</td>
		</tr>
		<tr>
			<td valign="top">
<% '- �Z�� - %>
				<br>
				<table border="1" width="260" class="disp">
					<% if gf_empItem(C_T11_HON_ZIP) then %>
						<tr>
							<td class="disph" width="100" height="16">�{�Ё�</td>
							<td class="disp"><%= m_HON_ZIP %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_HON_JUSYO) then %>
						<tr>
							<td class="disph" height="16" rowspan="3">�{�ЏZ��</td>
							<td class="disp"><%= m_HON_JUSYO %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

				<BR>
				<table class="disp" border="1" width="260">
					<% if gf_empItem(C_T11_GEN_ZIP) then %>
						<tr>
							<td class="disph" width="100" height="16">���Z����</td>
							<td class="disp"><%= m_GEN_ZIP %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_GEN_JUSYO) then %>
						<tr>
							<td class="disph" height="16">�Z�@�@��</td>
							<td class="disp"><%= m_GEN_JUSYO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_GEN_TEL) then %>
						<tr>
							<td class="disph" height="16">���Z��TEL</td>
							<td class="disp"><%= m_GEN_TEL %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
		</tr>
	</table>

	<BR>
	<input type="button" class="button" value="�@����@" onClick="parent.window.close();">

	</div>
	</form>
	</body>
	</html>
<% End Sub %>