<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍��ڍ�
' ��۸���ID : gak/gak0300/kojin_sita2.asp
' �@      �\: �������ꂽ�w���̏ڍׂ�\������(���w���)
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
	Public m_NYUGAKU_KBN	'���w�敪
	Public m_HyoujiFlg		'�\���׸�
	Public m_TYUGAKKOMEI	'���w�Z��
	Public m_NYU_GAKKA		'�w�Ȗ�
	Public m_KURABUMEI		'�N���u��


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
		w_iRet = f_GetDetailNyugaku()
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
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
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
Function f_GetDetailNyugaku()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailNyugaku = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql &  vbCrLf &" SELECT "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUNENDO,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKU_KBN, " 
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKUBI,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKUBI,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYU_GAKKA, "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_JUKEN_NO,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYU_SEISEKI, " 
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYUGAKKO_CD, "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYUSOTUGYOBI,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYU_CLUB, "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYU_CLUB_SYOSAI "
		w_sSql = w_sSql &  vbCrLf &" FROM  "
		w_sSql = w_sSql &  vbCrLf &" 	T11_GAKUSEKI A "
		w_sSql = w_sSql &  vbCrLf &" WHERE "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_GAKUSEI_NO  = '" & Session("GAKUSEI_NO") & "' "

'		w_sSql = ""
'		w_sSql = w_sSql &  vbCrLf &" SELECT "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUNENDO,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKU_KBN, " 
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKUBI,  "
'		w_sSql = w_sSql &  vbCrLf &" 	C.M02_GAKKAMEI,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_JUKEN_NO,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYU_SEISEKI, " 
'		w_sSql = w_sSql &  vbCrLf &" 	D.M13_TYUGAKKOMEI,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYUSOTUGYOBI,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYU_CLUB, "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYU_CLUB_SYOSAI "
'		w_sSql = w_sSql &  vbCrLf &" FROM  "
'		w_sSql = w_sSql &  vbCrLf &" 	T11_GAKUSEKI A, "
'		w_sSql = w_sSql &  vbCrLf &" 	M02_GAKKA    C, "
'		w_sSql = w_sSql &  vbCrLf &" 	M13_TYUGAKKO D "
'		w_sSql = w_sSql &  vbCrLf &" WHERE "
'		w_sSql = w_sSql &  vbCrLf &" 		A.T11_TYUGAKKO_CD = D.M13_TYUGAKKO_CD(+) "
'		w_sSql = w_sSql &  vbCrLf &" 	AND A.T11_NYU_GAKKA   = C.M02_GAKKA_CD(+) "
'		w_sSql = w_sSql &  vbCrLf &" 	AND D.M13_NENDO       =  " & Session("HyoujiNendo")
'		w_sSql = w_sSql &  vbCrLf &" 	AND C.M02_NENDO       =  " & Session("HyoujiNendo")
'		w_sSql = w_sSql &  vbCrLf &" 	AND A.T11_GAKUSEI_NO  = '" & Session("GAKUSEI_NO") & "' "

		iRet = gf_GetRecordset(m_Rs, w_sSql)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailNyugaku = 99
			Exit Do
		End If

		'//���w�敪���擾
		if Not gf_GetKubunName(C_NYUGAKU,m_Rs("T11_NYUGAKU_KBN"),Session("HyoujiNendo"),m_NYUGAKU_KBN) then Exit Do

		'//���w�Z�����擾
		if Not f_GetTyugakkoMei(m_Rs("T11_TYUGAKKO_CD"),m_TYUGAKKOMEI) then Exit Do

		'//�w�Ȗ����擾
		if Not f_GetGakkaMei(m_Rs("T11_NYU_GAKKA"),m_NYU_GAKKA) then Exit Do

		'//�N���u�����擾
		if Not f_GetKurabuMei(m_Rs("T11_TYU_CLUB"),m_KURABUMEI) then Exit Do

		'//����I��
		f_GetDetailNyugaku = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  ���w�Z�����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  True: False
'*  [����]  
'********************************************************************************
Function f_GetTyugakkoMei(pKey,pTYUGAKKOMEI)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetTyugakkoMei = False

	'// NULL�Ȃ甲����(False)
	if trim(pKey) = "" then Exit Function

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
    w_sSQL = w_sSQL & " 	M13_TYUGAKKOMEI "
    w_sSQL = w_sSQL & " FROM M13_TYUGAKKO "
    w_sSQL = w_sSQL & " WHERE M13_TYUGAKKO_CD = '" & pKey & "'"
    w_sSQL = w_sSQL & " 	AND M13_NENDO = " & Session("HyoujiNendo")

	iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit Function
	End If

	'// EOF�Ȃ甲����(False)
	if w_Rs.Eof then 
		f_GetTyugakkoMei = True
		Exit Function
	End if

	'// ���w�Z��
	pTYUGAKKOMEI = w_Rs("M13_TYUGAKKOMEI")

    '// �I������
    If Not IsNull(w_Rs) Then gf_closeObject(w_Rs)

	'//����I��
	f_GetTyugakkoMei = True

End Function


'********************************************************************************
'*  [�@�\]  �w�Ȗ����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  True: False
'*  [����]  
'********************************************************************************
Function f_GetGakkaMei(pKey,pNYU_GAKKA)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetGakkaMei = False

	'// NULL�Ȃ甲����(False)
	if trim(pKey) = "" then Exit Function

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
    w_sSQL = w_sSQL & " 	M02_GAKKAMEI "
    w_sSQL = w_sSQL & " FROM M02_GAKKA "
    w_sSQL = w_sSQL & " WHERE M02_GAKKA_CD = '" & pKey & "'"
    w_sSQL = w_sSQL & " 	AND M02_NENDO = " & Session("HyoujiNendo")

	iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit Function
	End If

	'// EOF�Ȃ甲����(False)
	if w_Rs.Eof then 
		f_GetGakkaMei = True
		Exit Function
	End if

	'// �w�Ȗ�
	pNYU_GAKKA = w_Rs("M02_GAKKAMEI")

    '// �I������
    If Not IsNull(w_Rs) Then gf_closeObject(w_Rs)

	'//����I��
	f_GetGakkaMei = True

End Function


'********************************************************************************
'*  [�@�\]  �N���u�����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  True: False
'*  [����]  
'********************************************************************************
Function f_GetKurabuMei(pKey,pKURABUMEI)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetGakkaMei = False

	'// NULL�Ȃ甲����(False)
	if trim(pKey) = "" then Exit Function

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
    w_sSQL = w_sSQL & " 	M17_BUKATUDOMEI "
    w_sSQL = w_sSQL & " FROM M17_BUKATUDO "
    w_sSQL = w_sSQL & " WHERE M17_BUKATUDO_CD = '" & pKey & "'"
    w_sSQL = w_sSQL & " 	AND M17_NENDO = " & Session("HyoujiNendo")

	iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit Function
	End If

	'// EOF�Ȃ甲����(False)
	if w_Rs.Eof then 
		f_GetKurabuMei = True
		Exit Function
	End if

	'// �N���u��
	pKURABUMEI = w_Rs("M17_BUKATUDOMEI")

    '// �I������
    If Not IsNull(w_Rs) Then gf_closeObject(w_Rs)

	'//����I��
	f_GetKurabuMei = True

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

	m_NYUNENDO 		  = ""
	m_F_NYUGAKUBI 	  = ""
'	m_GAKKAMEI 		  = ""
	m_JUKEN_NO		  = ""
	m_NYU_SEISEKI 	  = ""
'	m_TYUGAKKOMEI 	  = ""
	m_TYUSOTUGYOBI 	  = ""
'	m_BUKATUDOMEI 	  = ""
	m_TYU_CLUB_SYOSAI = ""

	if Not m_Rs.Eof Then
		m_NYUNENDO 		  = m_Rs("T11_NYUNENDO") 
		m_NYUGAKUBI 	  = m_Rs("T11_NYUGAKUBI") 
'		m_GAKKAMEI 		  = m_Rs("M02_GAKKAMEI") 
		m_JUKEN_NO		  = m_Rs("T11_JUKEN_NO") 
		m_NYU_SEISEKI 	  = m_Rs("T11_NYU_SEISEKI") 
'		m_TYUGAKKOMEI 	  = m_Rs("M13_TYUGAKKOMEI") 
		m_TYUSOTUGYOBI 	  = m_Rs("T11_TYUSOTUGYOBI") 
'		m_BUKATUDOMEI 	  = m_Rs("M17_BUKATUDOMEI") 
		m_TYU_CLUB_SYOSAI = m_Rs("T11_TYU_CLUB_SYOSAI") 
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
	-->
	</style>
	<script language="javascript">
	<!--
		function sbmt(m,i) {
			document.forms[0].mode.value = m;
			document.forms[0].id.value = i;
			document.forms[0].submit();
		}
	-->
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
			<td nowrap><b>�����w���</b></td>
			<td nowrap><a href="kojin_sita3.asp">���w�N���</a></td>
			<td nowrap><a href="kojin_sita4.asp">�����l�E����</a></td>
			<td nowrap><a href="kojin_sita5.asp">���ٓ����</a></td>
		</tr>
	</table>
	<br>
	

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td width="60">&nbsp</td>
			<td valign="top" align="left">

				<br>
				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T11_NYUGAKU_KBN) then %>
						<tr>
							<td class="disph" width="100" height="16">���w�敪</td>
							<td class="disp"><%= m_NYUGAKU_KBN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_NYUNENDO) then %>
						<tr>
							<td class="disph" height="16">���w�N�x</td>
							<td class="disp"><%= m_NYUNENDO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_NYUGAKUBI) then %>
						<tr>
							<td class="disph" height="16">�� �w ��</td>
							<td class="disp"><%= m_NYUGAKUBI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_NYU_GAKKA) then %>
						<tr>
							<td class="disph" height="16">�w�@�@��</td>
							<td class="disp"><%= m_NYU_GAKKA %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" align="left">

				<br>
				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T11_JUKEN_NO) then %>
						<tr>
							<td class="disph" width="100" height="16">�󌱔ԍ�</td>
							<td class="disp"><%= m_JUKEN_NO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_NYU_SEISEKI) then %>
						<tr>
							<td class="disph" height="16">���w����</td>
							<td class="disp"><%= m_NYU_SEISEKI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_TYUGAKKO_CD) then %>
						<tr>
							<td class="disph" height="16">���w�Z��</td>
							<td class="disp"><%= m_TYUGAKKOMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TYUSOTUGYOBI) then %>
						<tr>
							<td class="disph" height="16">�� �� ��</td>
							<td class="disp"><%= m_TYUSOTUGYOBI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_TYU_CLUB) then %>
						<tr>
							<td class="disph" height="16">�N �� �u</td>
							<td class="disp"><%= m_KURABUMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_TYU_CLUB_SYOSAI) then %>
						<tr>
							<td class="disph" height="16">�N���u�ڍ�</td>
							<td class="disp"><%= m_TYU_CLUB_SYOSAI %>&nbsp</td>
						</tr>
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