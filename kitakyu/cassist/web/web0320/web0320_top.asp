<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �g�p���ȏ��o�^
' ��۸���ID : web/web0320/web0320_top.asp
' �@	  �\: ��y�[�W �g�p���ȏ��o�^�̌������s��
'-------------------------------------------------------------------------
' ��	  ��:�����R�[�h 	��		SESSION���i�ۗ��j
'			:�N�x			��		SESSION���i�ۗ��j
' ��	  ��:�Ȃ�
' ��	  �n:�����R�[�h 	��		SESSION���i�ۗ��j
'			:�N�x			��		SESSION���i�ۗ��j
' ��	  ��:
'			�������\��
'				�R���{�{�b�N�X�͋󔒂ŕ\��
'			���\���{�^���N���b�N��
'				���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��	  ��: 2001/08/01 �O�c �q�j
' ��	  �X: 2001/08/07 ���{ ����	   NN�Ή��ɔ����\�[�X�ύX
' ��	  �X: 2001/08/18 �ɓ��@���q ���N�x�̊w����񂪂Ȃ����͎��N�x�̓��͂��o���Ȃ��悤�ɂ���
' ��	  �X: 2001/08/22 �ɓ��@���q ������I���ł���悤�ɕύX
' ��	  �X: 2001/12/01 �c�� ��K �����w�Ȃ݂̂�ύX����悤�ɏC��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public	m_bErrFlg			'�װ�׸�

	'�s�����I��p��Where����
	Public m_iNendo 		'�N�x
	Public m_sKyokanCd		'�����R�[�h
	Public m_sKyokanSimei	'��������
	Public m_bJinendoGakki	'���N�x�w�����

	Public m_iGakunen
	Public m_sGakkaCd
	Public m_sGakunenWhere
	Public m_sGakkaWhere

	Public	m_Rs
	Public	m_iMax			'�ő�y�[�W
	Public	m_iDsp			'�ꗗ�\���s��

	Public m_sSyozokuGakka		'//2001/12/01 Add ���O�C�����������̏�������w��

'///////////////////////////���C������/////////////////////////////

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
	w_sMsgTitle="�g�p���ȏ��o�^"
	w_sMsg=""
	w_sRetURL="../../login/default.asp" 	
	w_sTarget="_top"


	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = False
	
	m_iDsp = C_PAGE_LINE
	
	Do
		'// �ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		Call s_SetParam()
		
		'//���N�x��񂪂��邩�`�F�b�N
		w_iRet = f_GetJinendoGakki(m_bJinendoGakki)
		If w_iRet  = False Then
			m_bErrFlg = True
			exit do
		End If
		
		'//�w�N�̃R���{���擾
		Call s_MakeGakunenWhere()
		
		'//�w�Ȃ̃R���{���擾
		Call s_MakeGakkaWhere2()
		
		'// �y�[�W��\��
		Call showPage()
		Exit Do
	Loop
	
	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// �I������
	Call gf_closeObject(m_Rs)
	Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_SetParam()
	
	'If Request("txtNendo") = "" Then
	If Request("hidYear") = "" Then
		m_iNendo   = session("NENDO")
		
		'//���N�x��񂪂���ꍇ�́A���N�x�̋��ȏ��o�^�������ݒ�̑ΏۂƂ���
		If m_bJinendoGakki = True Then
			m_iNendo   = m_iNendo + 1
		End If
	Else
		'm_iNendo = Request("txtNendo")
		m_iNendo = Request("hidYear")
	End If
	
	'm_iGakunen = Request("txtGakunenCd")
	'm_sGakkaCd = Request("txtGakkaCD")
	
	m_iGakunen = Request("hidGakunen")
	m_sGakkaCd = Request("hidGakka")
	
End Sub

'********************************************************************************
'*	[�@�\]	���N�x�̊w����񂪂��邩�ǂ����`�F�b�N����
'*	[����]	�Ȃ�
'*	[�ߒl]	p_bJinendoGakki=true:�w����񂠂�
'*			p_bJinendoGakki=false:�w�����Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetJinendoGakki(p_bJinendoGakki)
	Dim w_iRet				'// �߂�l
	Dim w_sSQL				'// SQL��
	dim w_Rs

	on error resume next
	err.clear

	f_GetJinendoGakki = False
	p_bJinendoGakki = False

	'//���N�x�̊w����񂪂��邩�ǂ���
	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI"
	w_sSQL = w_sSQL & vbCrLf & " FROM M01_KUBUN"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_NENDO=" & cint(SESSION("NENDO"))+1
	w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & C_KAISETUKI

	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		m_bErrFlg = True
		Exit Function
	End If

	'//�f�[�^����������
	If w_Rs.EOF = False Then
		p_bJinendoGakki = True
	End If

	Call gf_closeObject(w_Rs)

	f_GetJinendoGakki = True

End Function

'********************************************************************************
'*	[�@�\]	�w�N�R���{�Ɋւ���WHERE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_MakeGakunenWhere()

	m_sGakunenWhere = ""
	m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iNendo
	m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

End Sub

'********************************************************************************
'*	[�@�\]	�w�ȃR���{�Ɋւ���WHRE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_MakeGakkaWhere()
	m_sGakkaWhere=""
	m_sGakkaWhere = " M02_NENDO = " & m_iNendo
	m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD <> '00' "

End Sub

'********************************************************************************
'*  [�@�\]  �w�ȃR���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeGakkaWhere2()
	'2001/12/01 Add ---->
	Dim w_sSQL				'//SQL��
	Dim w_iRet				'//�߂�l

	Dim w_oRecord			'//�����w�Ȏ擾�̂���

	'//�����w�Ȃ̎擾
	w_sSQL = ""
	w_sSQL = w_sSQL & "SELECT "
	w_sSQL = w_sSQL & "M04_GAKKA_CD "
	w_sSQL = w_sSQL & "From "
	w_sSQL = w_sSQL & "M04_KYOKAN "
	w_sSQL = w_sSQL & "Where "
	w_sSQL = w_sSQL & "M04_NENDO = " & m_iNendo & " "
	w_sSQL = w_sSQL & "And "
	w_sSQL = w_sSQL & "M04_KYOKAN_CD = '" & Session("KYOKAN_CD") & "'"
'Response.Write w_ssql
	w_iRet = gf_GetRecordset(w_oRecord, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit Sub
	End If

	If w_oRecord.EOF <> True Then
		m_sSyozokuGakka = w_oRecord("M04_GAKKA_CD")
	Else
		m_sSyozokuGakka =""
	End If

	'//����
	w_oRecord.Close
	Set w_oRecord = Nothing

	'2001/12/01 Add <----

'	m_sGakkaWhere=""
'	m_sGakkaWhere = " M02_NENDO = " & m_iNendo
'	'm_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD <> '00' "
'	m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD = '" & m_sSyozokuGakka & "' "	'2001/12/01 Mod
	
	if m_sSyozokuGakka = "00" then
		m_sGakkaWhere=""
		m_sGakkaWhere = " M02_NENDO = " & m_iNendo
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD <> '00' "	
	else
		m_sGakkaWhere=""
		m_sGakkaWhere = " M02_NENDO = " & m_iNendo
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD = '" & m_sSyozokuGakka & "' "	'2001/12/01 Mod
	end if
	
End Sub

Function f_ComboSet_Gakka(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' �@	�\:ComboBox�Z�b�g
' ��	�l:OK=True/NG=False
' ��	��:p_oCombo - ComboBox
'		   p_sTableName - �e�[�u����
'		   p_sWhere - Where����(WHERE��͗v��Ȃ�)
'		   p_sSelectOption - <SELECT id=select1 name=select1>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
'		   p_bWhite - �擪�ɋ󔒂����邩
'		   p_sSelectCD - �W���I�����������R�[�h(""�Ȃ�I���Ȃ�)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��	�l:�����w�Ȃ���ʑ����w�Ȃ̏ꍇ�͑S�w�Ȃ���
'*************************************************************************************
	Dim w_sId			'ID�t�B�[���h��
	Dim w_sName 		'���̃t�B�[���h��
	Dim w_sTableName	'���̃e�[�u����
	Dim w_rst

	f_ComboSet_Gakka = False
	do 
	''�}�X�^����SELECT����t�B�[���h�����擾
	If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
		Exit Do
	End If

	''�}�X�^SELECT
	If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
		Exit Do
	End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
	Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

	'�󔒂�Option�̑��
	If p_bWhite Then
		response.Write " <Option Value="&C_CBO_NULL&">�@�@�@�@�@ "& Chr(13)
	End If
	
	Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
	
	'// ��ʑ����w�Ȃ̏ꍇ�͑S�w�Ȃ�I���\
	If m_sSyozokuGakka = "00" Then
		response.write(" <Option Value='" & C_CLASS_ALL & "'")
		If CStr(p_sSelectCD) = CStr(C_CLASS_ALL) Then
			response.write " Selected "
		End If
		response.Write(">" & "�S�w��" & Chr(13))
	Else
		'Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
	End If
	
	Response.write("</select>" & chr(13))

	If Not w_rst Is Nothing Then
		w_rst.Close
		Set w_rst = Nothing
	End If
   
	f_ComboSet_Gakka = True
	Exit Do
	Loop
End Function

'****************************************************
'[�@�\] �f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'		(���X�g�_�E���{�b�N�X�I��\���p)
'[����] pData1 : �f�[�^�P
'		pData2 : �f�[�^�Q
'[�ߒl] f_Selected : "SELECTED" OR ""
'					
'****************************************************
Function f_Selected(pData1,pData2)

	f_Selected = ""

	If IsNull(pData1) = False And IsNull(pData2) = False Then
		If trim(cStr(pData1)) = trim(cstr(pData2)) Then
			f_Selected = "selected" 
		Else
		End If
	End If

End Function

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
<title>�g�p���ȏ��o�^</title>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//	[�@�\]	�\���{�^���N���b�N���̏���
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Search(){
		with(document.frm){
			
			// ������NULL����������
			// ���N�x
			if( f_Trim(txtNendo.value) == "" ){
				alert("�N�x�̑I�����s���Ă�������");
				txtNendo.focus();
				return ;
			}
			
			hidYear.value = txtNendo.value;
			hidGakunen.value = txtGakunenCd.value;
			hidGakka.value = txtGakkaCD.value;
			
			action="web0320_main.asp";
			target="main";
			submit();
		}
	
	}
	//************************************************************
	//	[�@�\]	�o�^�{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Touroku(){
		
		if(parent.main.document.frm){
			document.frm.txtPageCD.value = parent.main.document.frm.txtPageCD.value;
		}
		
		with(document.frm){
			hidYear.value = txtNendo.value;
			hidGakunen.value = txtGakunenCd.value;
			hidGakka.value = txtGakkaCD.value;
			
			action="touroku.asp";
			target="<%=C_MAIN_FRAME%>";
			txtMode.value = "Touroku";
			submit();
		}
	}
	
	//************************************************************
	//	[�@�\]	�N�x���ύX���ꂽ�Ƃ��A�{��ʂ��ĕ\��
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_ReLoadMyPage(){
		with(document.frm){
			hidYear.value = txtNendo.value;
			hidGakunen.value = txtGakunenCd.value;
			hidGakka.value = txtGakkaCD.value;
			
			action="./web0320_top.asp";
			target="_self";
			submit();
		}
	}
	
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>

	<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<center>
	<form name="frm" method="POST">
	<% call gs_title("�g�p���ȏ��o�^","��@��") %>
	<br>
	<table border="0">
	<tr>
	<td valign="bottom">
		<table border="0" cellpadding="1" cellspacing="1">
		<tr>
		<td align="left" class="search">
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
			<td Nowrap>
				<select name="txtNendo" onchange = 'javascript:f_ReLoadMyPage()'>
					<%If m_bJinendoGakki = True Then%>
						<%w_iNen=Session("NENDO")%>
						<!--<option VALUE="<%= w_iNen + 1 %>" <%=f_Selected(Request("txtNendo"),w_iNen + 1)%> ><%= w_iNen + 1 %>
						<option VALUE="<%= w_iNen %>"	  <%=f_Selected(Request("txtNendo"),w_iNen)%> ><%= w_iNen %>
						-->
						<option VALUE="<%= w_iNen + 1 %>" <%=f_Selected(Request("hidYear"),w_iNen + 1)%> ><%= w_iNen + 1 %>
						<option VALUE="<%= w_iNen %>"	  <%=f_Selected(Request("hidYear"),w_iNen)%> ><%= w_iNen %>
					<%Else%>
						<option VALUE="<%= m_iNendo %>" 			><%= m_iNendo %>
					<%End If%>

				</select>
			</td>
			<td>�N�x&nbsp;&nbsp;</td>

			<td>�w�N</td>
			<td nowrap align="left">
				<% call gf_ComboSet("txtGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere," style='width:40px;' ",True,m_iGakunen) %>
			</td>

			<td nowrap>�w��</td>
			<td nowrap align="left">

			<%	'���ʊ֐�����w�ȂɊւ���R���{�{�b�N�X���o�͂���
				call f_ComboSet_Gakka("txtGakkaCD",C_CBO_M02_GAKKA,m_sGakkaWhere,"style='width:175px;' ",True,m_sGakkaCd)%>
			</td>

			</tr><tr>
			<td Nowrap align="right" colspan=6>
			<input class="button" type="button" value="�@�\�@���@" onClick = "javascript:f_Search();">
			</td>
			</tr>
			</table>
		</td>
		</tr>
		</table>
	</td>
	<td valign="top">
	<a href="javascript:f_Touroku();">�V�K�o�^�͂�����</a><br><img src="../../image/sp.gif" height="10"><br>
	</td>
	</tr>
	</table>
		<input type="hidden" name="txtMode" value="">
		
		<input type="hidden" name="hidYear">
		<input type="hidden" name="hidGakunen">
		<input type="hidden" name="hidGakka">
		<input type="hidden" name="txtPageCD">
	</form>
	</center>
	</body>
</html>

<%
End Sub
%>
