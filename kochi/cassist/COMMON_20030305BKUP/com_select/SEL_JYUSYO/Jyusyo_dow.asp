<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �s�����������
' ��۸���ID : Common/com_select/SEL_JYUSYO/Jyusyo_dow.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:	
' 	           	hidSchMode	= �����׸�
'   	        hidKenCd	= ���R�[�h
'       	    hidShiCd	= �s�����R�[�h
'           	txtCyouiki	= ����R�[�h
' 
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/30 ���i
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

	Public  m_bErrFlg           '�װ�׸�
	Public  m_bSchMode			'�����׸�
	Public  m_KenCd				'���R�[�h
	Public  m_ShiCd				'�s�����R�[�h
	Public  m_Cyouiki 			'����R�[�h
	Public  m_CyouRs			'ں��޾�ĵ�޼ު��(��)

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

    Dim w_iRet              '// �߂�l

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�������o�^"
    w_sMsg=""
    w_sRetURL="../../login/top.asp"
    w_sTarget="_parent"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

		m_bSchMode = request("hidSchMode")	'�����׸�
		m_KenCd    = request("hidKenCd")	'���R�[�h
		m_ShiCd    = request("hidShiCd")	'�s�����R�[�h
		m_Cyouiki  = request("txtCyouiki")	'����R�[�h

		'// ���f�[�^���擾
		if Not f_GetCyou() then Exit Do

        '// �y�[�W��\��
        Call showPage()

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Rs)

    '// �I������
    Call gs_CloseDatabase()

End Sub

'******************************************************************
'�@�@�@�\�F���f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'******************************************************************
Function f_GetCyou()

	On Error Resume Next
	Err.Clear

	f_GetCyou = False

	
	if m_bSchMode then

		w_iWhereFlg = 0		'where�傪�O�ɂ���Flg

		'// ����WHERE��
		w_sKenSQL = ""
		if m_KenCd <> "" then
			w_sKenSQL = " M12_KEN_CD = '" & gf_fmtZero(m_KenCd,2) & "' "
			w_iWhereFlg = 1
		End if

		'// �s������WHERE��
		w_sShiSQL = ""
		if m_ShiCd <> "" then
			'// ���ł�WHERE������������"AND"�łȂ�
			if w_iWhereFlg = 1 then 
				w_sShiSQL = w_sShiSQL & " AND "
			End if
			w_sShiSQL = w_sShiSQL & " M12_SITYOSON_CD = '" & m_ShiCd & "' "
			w_iWhereFlg = 1
		End if

		'// ����w���WHERE��
		w_sCyouikiSQL = ""
		if m_Cyouiki <> "" then
			'// ���ł�WHERE������������"AND"�łȂ�
			if w_iWhereFlg = 1 then 
				w_sShiSQL = w_sShiSQL & " AND "
			End if
			w_sCyouikiSQL = w_sCyouikiSQL & " M12_TYOIKIMEI like '%" & m_Cyouiki & "%' "
		End if

		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	M12_YUBIN_BANGO, "
		w_sSQL = w_sSQL & " 	M12_SITYOSONMEI, "
		w_sSQL = w_sSQL & " 	M12_RENBAN, "
		w_sSQL = w_sSQL & " 	M12_TYOIKIMEI "
		w_sSQL = w_sSQL & " FROM  "
		w_sSQL = w_sSQL & " 	M12_SITYOSON "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & w_sKenSQL
		w_sSQL = w_sSQL & w_sShiSQL
		w_sSQL = w_sSQL & w_sCyouikiSQL

		iRet = gf_GetRecordset(m_CyouRs, w_sSQL)
		If iRet <> 0 Then
			m_bErrFlg = True
			'ں��޾�Ă̎擾���s
			Exit Function
		End If

	End if

	f_GetCyou = True

End Function


Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
<html>

<head>
<link rel=stylesheet href="../../style.css" type=text/css>
<!--#include file="../../jsCommon.htm"-->
<script language="JavaScript">
<!--
    //************************************************************
    //  [�@�\]  �I�[�v�i�[�Ƀf�[�^��Ԃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
	function jf_ReturnDate(pZipCode,pJusyoNum,pRenban) {

		parent.opener.document.frm.txtYUBINBANGO.value = pZipCode;		// �X�֔ԍ�
		pJusyo1 = eval("document.frm.hidSITYOSONMEI_"+pJusyoNum+".value");
		pJusyo2 = eval("document.frm.hidTYOIKIMEI_"+pJusyoNum+".value");
		parent.opener.document.frm.txtJUSYO1.value  = pJusyo1;			// �Z���P
		parent.opener.document.frm.txtJUSYO2.value  = pJusyo2;			// �Z���Q
		parent.opener.document.frm.txtJUSYO3.value  = "";				// �Z���R

		parent.opener.document.frm.txtRenban.value  = pRenban;			// �A��(�L�[)
		parent.opener.document.frm.txtKenCd.value   = document.frm.hidKenCd.value;	// ���R�[�h(�L�[)
		parent.opener.document.frm.txtSityoCd.value = document.frm.hidShiCd.value;	// �s�����R�[�h(�L�[)

		parent.window.close();

	}
//-->
</script>
</head>

<body>
<form name="frm" method="post">
<div align="center">

<% if m_bSchMode then %>
<table>
	<tr>
		<td align="center">

			<table border="1" class="hyo">
				<tr>
					<th align="center" class="header" width="80">�X�֔ԍ�</th>
					<th align="center" class="header" width="266">���於</th>
				</tr>
				<% if m_CyouRs.Eof then %>
					<tr>
						<td colspan="2" class="CELL1">�Y������f�[�^������܂���B</td>
					</tr>
				<% End if %>
				<% 
					i = 1		'�A��
					Do Until m_CyouRs.Eof
					
					%>
					<tr>
						<td class="CELL1" align="center"><a href="javascript:jf_ReturnDate('<%= m_CyouRs("M12_YUBIN_BANGO") %>','<%=i%>','<%= m_CyouRs("M12_RENBAN") %>');"><%= m_CyouRs("M12_YUBIN_BANGO") %></a>
							<input type="hidden" name="hidSITYOSONMEI_<%=i%>" value="<%= m_CyouRs("M12_SITYOSONMEI") %>">
							<input type="hidden" name="hidTYOIKIMEI_<%=i%>"   value="<%= m_CyouRs("M12_TYOIKIMEI") %>"></td>
						<td class="CELL1" nowrap><%= m_CyouRs("M12_SITYOSONMEI") & m_CyouRs("M12_TYOIKIMEI") %></td>
					</tr>
					<%
						i = i + 1
						m_CyouRs.MoveNext
					Loop
				%>
			</table>

		</td>
	</tr>
	<tr>
		<td align="right">
			<input type="button" class="button" value="�L�����Z��" onClick="parent.window.close();">
		</td>
	</tr>
</table>
<% End if %>

<input type="hidden" name="hidKenCd" value="<%= m_KenCd %>">
<input type="hidden" name="hidShiCd" value="<%= m_ShiCd %>">
</div>
</form>
</body>
</html>
<%
End Sub
%>