<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍��ڍ�
' ��۸���ID : gak/gak0300/kojin_sita1.asp
' �@      �\: �������ꂽ�w���̏ڍׂ�\������(�l���)
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
	Public m_SEIBETU		'����
	Public m_BLOOD			'���t�^
	Public m_RH				'RH
	Public m_HOG_ZOKU		'�ی�ґ���
	Public m_HOS_ZOKU		'�ۏؐl����
	Public m_RYOSEI_KBN		'�����敪
	Public m_RYUNEN_FLG		'�i���敪

	Public m_HyoujiFlg		'�\���׸�
	Public m_KakoRs			'ں��޾�ĵ�޼ު��(�ߋ��׽)
	Public mHyoujiNendo		'�\���N�x

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
		'//�ߋ��̃N���X���擾
		w_iRet = f_GetDetailKakoClass()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'//�\�����ڂ��擾
		w_iRet = f_GetDetailGakunen()
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
'*  [�@�\]  �ߋ��̃N���X���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:����I��	1:�C�ӂ̃G���[  99:�V�X�e���G���[
'*  [����]  
'********************************************************************************
Function f_GetDetailKakoClass()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailKakoClass = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	T13.T13_NENDO, "
		w_sSql = w_sSql & " 	T13.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	T13.T13_CLASS "
		w_sSql = w_sSql & " FROM T13_GAKU_NEN T13 "
		w_sSql = w_sSql & " WHERE  "
		w_sSql = w_sSql & " 	T13.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		w_sSql = w_sSql & " 	ORDER BY T13.T13_NENDO DESC "

		iRet = gf_GetRecordset(m_KakoRs, w_sSql)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailKakoClass = 99
			Exit Do
		End If

		if m_KakoRs.Eof then
			msMsg = "�w�N���擾���ɃG���[���������܂���"
			f_GetDetailKakoClass = 99
			Exit Do
		End if

		'//����I��
		f_GetDetailKakoClass = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �\�����ڂ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:����I��	1:�C�ӂ̃G���[  99:�V�X�e���G���[
'*  [����]  
'********************************************************************************
Function f_GetDetailGakunen()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	'// �\������N�x�����߂�
	wSelNendo = request("selNendo")
	if gf_IsNull(wSelNendo) then
		mHyoujiNendo = Session("HyoujiNendo")
	Else
		mHyoujiNendo = wSelNendo
	End if

	f_GetDetailGakunen = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T13_NENDO, "
		w_sSql = w_sSql & " 	A.T13_GAKUSEKI_NO,"
		w_sSql = w_sSql & " 	A.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	B.M02_GAKKAMEI, "
		w_sSql = w_sSql & " 	A.T13_SYUSEKI_NO1,  "
		w_sSql = w_sSql & " 	A.T13_CLASS,  "
		w_sSql = w_sSql & " 	A.T13_SYUSEKI_NO2,  "
		w_sSql = w_sSql & " 	A.T13_RYOSEI_KBN,  "
		w_sSql = w_sSql & " 	A.T13_RYUNEN_FLG,  "
		w_sSql = w_sSql & " 	A.T13_SINTYO,  "
		w_sSql = w_sSql & " 	A.T13_TAIJYU,  "
		w_sSql = w_sSql & " 	A.T13_CLUB_1,  "
		w_sSql = w_sSql & " 	A.T13_CLUB_2,  "
		w_sSql = w_sSql & " 	A.T13_TOKUKATU, "
		w_sSql = w_sSql & " 	A.T13_TOKUKATU_DET,  "
		w_sSql = w_sSql & " 	A.T13_NENSYOKEN, "
		w_sSql = w_sSql & " 	A.T13_NENBIKO "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A, "
		w_sSql = w_sSql & " 	M02_GAKKA    B "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 	 A.T13_GAKKA_CD   = B.M02_GAKKA_CD(+) "
		w_sSql = w_sSql & "  AND A.T13_NENDO      = B.M02_NENDO(+) "
		w_sSql = w_sSql & "  AND A.T13_NENDO      = " & mHyoujiNendo
		w_sSql = w_sSql & "  AND A.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "

		iRet = gf_GetRecordset(m_Rs, w_sSql)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailGakunen = 99
			Exit Do
		End If

		'//�����敪���擾
		if Not gf_GetKubunName(C_NYURYO,m_Rs("T13_RYOSEI_KBN"),Session("HyoujiNendo"),m_RYOSEI_KBN) then Exit Do

		'//�i���敪���擾
		Select Case m_Rs("T13_RYUNEN_FLG")
			Case 0: m_RYUNEN_FLG = ""
			Case 1: m_RYUNEN_FLG = C_SHINKYU_NO
		End Select

		'//����I��
		f_GetDetailGakunen = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  ���������擾����
'*  [����]  p_sClubCd:����CD
'*  [�ߒl]  f_GetClubName�F������
'*  [����]  
'********************************************************************************
Function f_GetClubName(p_sClubCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubName = ""
	w_sClubName = ""

	Do

		'//����CD����̎�
		If trim(gf_SetNull2String(p_sClubCd)) = "" Then
			Exit Do
		End If

		'//���������擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & mHyoujiNendo
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD=" & p_sClubCd

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//������
			w_sClubName = rs("M17_BUKATUDOMEI")
		End If

		Exit Do
	Loop

	'//�߂�l���
	f_GetClubName = w_sClubName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

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

	m_NENDO		 	= ""
	m_GAKUSEKI_NO	= ""
	m_GAKUNEN		= ""
	m_GAKKAMEI	 	= ""
	m_SYUSEKI_NO1	= ""
	m_CLASS		 	= ""
	m_SYUSEKI_NO2	= ""
	m_SINTYO		= ""
	m_TAIJYU		= ""
	m_CLUB_1		= ""
	m_CLUB_2		= ""
	m_TOKUKATU	 	= ""
	m_TOKUKATU_DET  = ""
	m_NENSYOKEN	 	= ""
	m_F_NENBIKO	 	= ""

	if Not m_Rs.Eof Then
		m_NENDO		 	= m_Rs("T13_NENDO")
		m_GAKUSEKI_NO	= m_Rs("T13_GAKUSEKI_NO")
		m_GAKUNEN		= m_Rs("T13_GAKUNEN")
		m_GAKKAMEI	 	= m_Rs("M02_GAKKAMEI")
		m_SYUSEKI_NO1	= m_Rs("T13_SYUSEKI_NO1")
		m_CLASS		 	= m_Rs("T13_CLASS")
		m_SYUSEKI_NO2	= m_Rs("T13_SYUSEKI_NO2")
		m_SINTYO		= m_Rs("T13_SINTYO")
		m_TAIJYU		= m_Rs("T13_TAIJYU")
		m_CLUB_1		= f_GetClubName(gf_SetNull2String(m_Rs("T13_CLUB_1")))
		m_CLUB_2		= f_GetClubName(gf_SetNull2String(m_Rs("T13_CLUB_2")))
		m_TOKUKATU	 	= m_Rs("T13_TOKUKATU")
		m_TOKUKATU_DET  = m_Rs("T13_TOKUKATU_DET")
		m_NENSYOKEN	 	= m_Rs("T13_NENSYOKEN")
		m_F_NENBIKO	 	= m_Rs("T13_NENBIKO")
	End if

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
		//**************************************
		//*   �N�x�ڸ��ޯ�����ύX���ꂽ�Ƃ�
		//**************************************
		function jf_ChangSelect(){

			document.frm.submit();

		}

	//-->
	</script>
	</head>

	<body>
	<form action="kojin_sita3.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">����{���</a></td>
			<td nowrap><a href="kojin_sita1.asp">���l���</a></td>
			<td nowrap><a href="kojin_sita2.asp">�����w���</a></td>
			<td nowrap><b>���w�N���</b></td>
			<td nowrap><a href="kojin_sita4.asp">�����l�E����</a></td>
			<td nowrap><a href="kojin_sita5.asp">���ٓ����</a></td>
		</tr>
	</table>
	<br>

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td colspan="3">
				<span class="msg"><font size="2">�� �����N�x��ύX����ƁA�ߋ��̊w�N�������邱�Ƃ��ł��܂�<BR></font></span>
			</td>
		</tr>
		<tr>
			<td valign="top" align="left">

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_NENDO) then %>
						<tr>
							<td class="disph" width="100">�����N�x</td>
							<td class="disp"><select name="selNendo" onChange="jf_ChangSelect();">
												<% do until m_KakoRs.Eof 
													wSelected = ""
													if Cint(mHyoujiNendo) = Cint(m_KakoRs("T13_NENDO")) then
														wSelected = "selected"
													End if
													%>
													<option value="<%=m_KakoRs("T13_NENDO")%>" <%=wSelected%>><%=m_KakoRs("T13_NENDO")%>�N�x
												<% m_KakoRs.MoveNext : Loop %>
											</select></td>
						</tr>
<!--
						<tr>
							<td class="disph" width="100" height="16">�����N�x</td>
							<td class="disp"><%= m_NENDO %>&nbsp</td>
						</tr>
-->
					<% End if %>
					<% if gf_empItem(C_T13_GAKUSEKI_NO) then %>
						<tr>
							<td class="disph" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_1NEN)%></td>
							<td class="disp"><%= m_GAKUSEKI_NO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_GAKUNEN) then %>
						<tr>
							<td class="disph" height="16">�w�@�@�N</td>
							<td class="disp"><%= m_GAKUNEN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_COURCE_CD) then %>
						<tr>
							<td class="disph" height="16">�����w��</td>
							<td class="disp"><%= m_GAKKAMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLASS) then %>
						<tr>
							<td class="disph" height="16">�N���X</td>
							<td class="disp"><%= m_CLASS %>�g&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO1) then %>
						<tr>
							<td class="disph" height="16">�o�Ȕԍ�(�w��)</td>
							<td class="disp"><%= m_SYUSEKI_NO1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO2) then %>
						<tr>
							<td class="disph" height="16">�o�Ȕԍ�(�N���X)</td>
							<td class="disp"><%= m_SYUSEKI_NO2 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_RYOSEI_KBN) then %>
						<tr>
							<td class="disph" height="16">�����敪</td>
							<td class="disp"><%= m_RYOSEI_KBN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_RYUNEN_FLG) then %>
						<tr>
							<td class="disph" height="16">�i���敪</td>
							<td class="disp"><%= m_RYUNEN_FLG %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" align="left">

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_SINTYO) then %>
						<tr>
							<td class="disph" width="100" height="16">�g�@�@��</td>
							<td class="disp"><%= m_SINTYO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TAIJYU) then %>
						<tr>
							<td class="disph" height="16">�́@�@�d</td>
							<td class="disp"><%= m_TAIJYU %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_1) then %>
						<tr>
							<td class="disph" height="16">�N���u�����P</td>
							<td class="disp"><%= m_CLUB_1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_2) then %>
						<tr>
							<td class="disph" height="16">�N���u�����Q</td>
							<td class="disp"><%= m_CLUB_2 %>&nbsp</td>
						</tr>
					<% End if %>
<!--
					<% if gf_empItem(C_T13_TOKUKATU) then %>
						<tr>
							<td class="disph" height="16">���ʊ���</td>
							<td class="disp"><%= m_TOKUKATU %>&nbsp</td>
						</tr>
-->
					<% End if %>
					<% if gf_empItem(C_T13_TOKUKATU_DET) then %>
						<tr>
							<td class="disph" height="16">���ʊ����ڍ�</td>
							<td class="disp"><%= m_TOKUKATU_DET %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" align="left">

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_NENSYOKEN) then %>
						<tr><td class="disph" width="220" height="16">�w����Q�l�ƂȂ鏔����</td></tr>
						<tr><td class="disp" valign="top" height="220"><%= m_NENSYOKEN %><br><br></td></tr>
					<% End if %>
					<% if gf_empItem(C_T13_NENBIKO) then %>
						<tr><td class="disph" width="100" height="16">�� �l</td></tr>
						<tr><td class="disp" valign="top" height="100"><%= m_F_NENBIKO %></td></tr>
					<% End if %>
				</table>

			</td>
		</tr>
	</table>

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