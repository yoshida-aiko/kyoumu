<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �s�����������
' ��۸���ID : Common/com_select/SEL_JYUSYO/Jyusyo_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:	hidKenCd = ���R�[�h�i�������g�փT�u�~�b�g�j
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

    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_KenRs				'ں��޾�ĵ�޼ު�āi���}�X�^�j
    Public  m_KenCd				'���R�[�h
    Public  m_ShiCd				'�s�R�[�h
    Public  m_ShiRs				'ں��޾�ĵ�޼ު�āi�s����)
    Public  m_Cyouiki 			'����R�[�h
    Public  m_JUSYO1			'���s��(��������)
    Public  m_JUSYO2            '��    (��������)
    Public  m_NoHitFlg			'����˯ĂȂ��׸�

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
	m_NoHitFlg = 0

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

		'//-- ���[�U�[���͏��� --
		m_JUSYO1 = request("JUSYO1")		'���s��
		m_JUSYO2 = request("JUSYO2")		'��

'		m_JUSYO1 = Session("m_JUSYO1")
'		m_JUSYO2 = Session("m_JUSYO2")

		'//----------------------//
		m_KenCd   = request("hidKenCd")		'���R�[�h
		m_Cyouiki = request("txtCyouiki")	'����R�[�h

		'// ���[�U�[���͏������猧�R�[�h���擾
		if Not f_GetKenCode() then Exit Do

		'// ���f�[�^���擾
		if Not f_GetKenMaster() then Exit Do

		'// ���[�U�[���͏�������s�����R�[�h���擾
		if Not f_GetShicyosonCode() then Exit Do

		'// �s�������擾
		if Not f_GetShicyoson() then Exit Do
        
        '// �y�[�W��\��
        Call showPage()

		'// ����ݍ폜
'		Session("m_JUSYO1") = ""
'		Session("m_JUSYO2") = ""

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ں��޾��CLOSE
    Call gf_closeObject(m_KenRs)
    Call gf_closeObject(m_ShiRs)

    '// �I������
    Call gs_CloseDatabase()

End Sub

'******************************************************************
'�@�@�@�\�F���[�U�[���͏������猧�R�[�h���擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'******************************************************************
Function f_GetKenCode()

	On Error Resume Next
	Err.Clear

	f_GetKenCode = False

	'// �Z���P��Null����Ȃ�������
	if m_JUSYO1 <> "" then

		'// ��CD���擾����
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	M16_KEN_CD "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	M16_KEN "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 		M16_NENDO  =  " & Session("NENDO")
		w_sSQL = w_sSQL & " 	AND M16_KENMEI = '" & m_JUSYO1 & "'"

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			m_bErrFlg = True
			'ں��޾�Ă̎擾���s
			Exit Function
		End If

		'// Null����Ȃ�������ϐ��ɓ����
		if Not w_Rs.Eof then
			m_KenCd = w_Rs("M16_KEN_CD")	
		Else
			m_NoHitFlg = 1						'// ˯ĂȂ��׸�
		End if

	End if

    Call gf_closeObject(w_Rs)
	f_GetKenCode = True

End Function

'******************************************************************
'�@�@�@�\�F���f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'******************************************************************
Function f_GetKenMaster()

	On Error Resume Next
	Err.Clear

	f_GetKenMaster = False

	'// ���}�X�^���擾����
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M16_KEN_CD,  "
	w_sSQL = w_sSQL & " 	M16_KENMEI "
	w_sSQL = w_sSQL & " FROM  "
	w_sSQL = w_sSQL & " 	M16_KEN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M16_NENDO = " & Session("NENDO")

	iRet = gf_GetRecordset(m_KenRs, w_sSQL)
	If iRet <> 0 Then
		m_bErrFlg = True
		'ں��޾�Ă̎擾���s
		Exit Function
	End If

	if m_KenRs.Eof then
		Call gs_SetErrMsg("�s���{�����擾���ɁA�G���[���������܂���")
		m_bErrFlg = True
		Exit Function
	End if

	f_GetKenMaster = True

End Function

'******************************************************************
'�@�@�@�\�F���[�U�[���͏�������s�����R�[�h���擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'******************************************************************
Function f_GetShicyosonCode()

	On Error Resume Next
	Err.Clear

	f_GetShicyosonCode = False

	if m_NoHitFlg = 1 then

		'// ���R�[�h���킩���Ă�����WHERE�����ɉ�����
		if m_KenCd <> "" then
			w_sKenSQL = " AND M12_KEN_CD = '" & m_KenCd & "'"
		End if
		
		'// �s����CD���擾����
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	M12_KEN_CD,  "
		w_sSQL = w_sSQL & " 	M12_SITYOSON_CD,  "
		w_sSQL = w_sSQL & " 	M12_SITYOSONMEI "
		w_sSQL = w_sSQL & " FROM  "
		w_sSQL = w_sSQL & " 	M12_SITYOSON "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 	M12_SITYOSONMEI Like '%" & m_JUSYO1 & "%' " 
		w_sSQL = w_sSQL & 		w_sKenSQL
		w_sSQL = w_sSQL & " Group by "
		w_sSQL = w_sSQL & " 	M12_KEN_CD,  "
		w_sSQL = w_sSQL & " 	M12_SITYOSON_CD,  "
		w_sSQL = w_sSQL & " 	M12_SITYOSONMEI "

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			m_bErrFlg = True
			'ں��޾�Ă̎擾���s
			Exit Function
		End If

		'// Null����Ȃ�������ϐ��ɓ����
		if Not w_Rs.Eof then
			m_KenCd = w_Rs("M12_KEN_CD")
			m_ShiCd = w_Rs("M12_SITYOSON_CD")
			m_NoHitFlg = 0
		End if

	End if

    Call gf_closeObject(w_Rs)
	f_GetShicyosonCode = True

End Function

'******************************************************************
'�@�@�@�\�F�s�����f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'******************************************************************
Function f_GetShicyoson()

	On Error Resume Next
	Err.Clear

	f_GetShicyoson = False

	'// ���R�[�h��NUll����Ȃ�������
	if m_KenCd <> "" then

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_KEN_CD,  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSON_CD,  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSONMEI "
		w_sSQL = w_sSQL & vbCrLf & " FROM  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSON "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_KEN_CD = '" & gf_fmtZero(m_KenCd,2) & "' "
		w_sSQL = w_sSQL & vbCrLf & " Group by "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_KEN_CD,  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSON_CD,  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSONMEI "

		iRet = gf_GetRecordset(m_ShiRs, w_sSQL)

		If iRet <> 0 Then
			m_bErrFlg = True
			'ں��޾�Ă̎擾���s
			Exit Function
		End If

	End if

	f_GetShicyoson = True

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

	'// �Z���Q�̒l�𒬈�ɂ����
	if m_JUSYO2 <> "" then
		m_Cyouiki = m_JUSYO2
	End if

	'// �I�����[�h���ɃT�u�~�b�g����
	if m_NoHitFlg = 1 AND m_JUSYO2 = "" then
		w_jFName = ""
	Else
		if m_JUSYO1 & m_JUSYO2 <> "" then
			w_jFName = " onLoad='jf_OnLoadSubmit();'"
		End if
	End if

	'// �l�X�P��IE�ɂ����textbox�̃T�C�Y��ς���
	if session("browser") = "IE" then
		w_FormSize = "61"
	Else
		w_FormSize = "44"
	End if

%>
<html>

<head>
<link rel=stylesheet href="../../style.css" type=text/css>
<!--#include file="../../jsCommon.htm"-->
<script language="JavaScript">
<!--


    //************************************************************
    //  [�@�\]  ���[�h���ɃT�u�~�b�g����
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
	function jf_OnLoadSubmit(){
		document.frm.action         = "Jyusyo_dow.asp";
		document.frm.target         = "dow";
		document.frm.submit();
	}


    //************************************************************
    //  [�@�\]  �����I�΂ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
	function jf_KenSelect(){
		w_KenCd = document.frm.selKen.value;
		document.frm.hidKenCd.value = w_KenCd;
		document.frm.action         = "Jyusyo_top.asp";
		document.frm.target         = "top";
		document.frm.submit();
	}

    //************************************************************
    //  [�@�\]  �����{�^�����I�΂ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
	function jf_ShiSelect(){
		if( document.frm.selKen.value == "" && document.frm.selShi.value == "" ){
			window.alert("�����w������Ă�������");
			document.frm.selKen.focus();
	        return false;
		}

		w_ShiCd = document.frm.selShi.value;
		document.frm.hidShiCd.value = w_ShiCd;
		document.frm.action         = "Jyusyo_dow.asp";
		document.frm.target         = "dow";
		document.frm.submit();
	}

    //************************************************************
    //  [�@�\]  ���挟���������ꂽ��
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
	function jf_Cyouiki(){

		if( f_Trim(document.frm.txtCyouiki.value) == ""){
			window.alert("����w������Ă�������");
			document.frm.txtCyouiki.focus();
	        return false;
		}

		document.frm.action         = "Jyusyo_dow.asp";
		document.frm.target         = "dow";
		document.frm.submit();
	}

//-->
</script>
</head>

<body <%= w_jFName %>>
<form name="frm" method="post">
<div align="center">

<% 
    call gs_title("�s��������","���@��")
%>
<table>
	<tr>
		<td align="center">

			<table border="1" class="hyo">
				<tr>
					<th align="center" class="header">����</th>
					<th align="center" class="header">�s������</th>
				<tr>
				<tr>
					<td class="CELL1">
						<select name="selKen" onChange="jf_KenSelect();">
							<option value="">�@�@�@�@
							<%
								Do Until m_KenRs.Eof
									'// �ϐ���"selected"������
									w_Selected = ""
									if Cint(m_KenCd) = Cint(m_KenRs("M16_KEN_CD")) then
										w_Selected = "Selected"
									End if
								%>
								<option value="<%= m_KenRs("M16_KEN_CD") %>" <%=w_Selected%>><%= m_KenRs("M16_KENMEI") %>
								<%
									m_KenRs.MoveNext
								Loop
							%>
						</select>
					</td>
					<td class="CELL1">
						<select name="selShi" style="width:230px;">
							<option value="">�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@
							<%
								if m_KenCd <> "" then
									Do Until m_ShiRs.Eof
									'// �ϐ���"selected"������
									w_Selected = ""
									if m_ShiCd = m_ShiRs("M12_SITYOSON_CD") then
										w_Selected = "selected"
									End if
									%>
									<option value="<%= m_ShiRs("M12_SITYOSON_CD") %>" <%=w_Selected%>><%= m_ShiRs("M12_SITYOSONMEI") %>
									<%
										m_ShiRs.MoveNext
									Loop
								End if
							%>
						</select>
						<input type="button" class="button" value="����" onClick="jf_ShiSelect();">
					</td>
				</tr>
			</table>

		</td>
	</tr>
	<tr>
		<td align="center">

			<table border="1" class="hyo">
				<tr><th align="center" class="header">�������於</th></tr>
				<tr><td class="CELL1"><input type="text" name="txtCyouiki" value="<%= m_Cyouiki %>" size="<%=w_FormSize%>">
						<input type="button" class="button" value="����" onClick="return jf_Cyouiki();"></td></tr>
			</table>

		</td>
	</tr>
</table>

</div>
<input type="hidden" name="hidKenCd" value="<%= m_KenCd %>">
<input type="hidden" name="hidShiCd">
<input type="hidden" name="hidSchMode" value="True">
</form>
</body>
</html>
<%
End Sub
%>