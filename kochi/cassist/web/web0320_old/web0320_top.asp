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

'	 m_iNendo	 = session("NENDO")

'	If Request("SKyokanCd1") <> "" Then
'		m_sKyokanCd = Request("SKyokanCd1")
'	Else
'	 	m_sKyokanCd = session("KYOKAN_CD")
'	End If

	m_iDsp = C_PAGE_LINE

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

		Call s_SetParam()

		'//���N�x��񂪂��邩�`�F�b�N
		w_iRet = f_GetJinendoGakki(m_bJinendoGakki)
		If w_iRet  = False Then
			m_bErrFlg = True
			exit do
		End If

'//�f�o�b�O
'Call s_DebugPrint

'		'//�����������擾
'		 w_iRet = f_KyokanSimei()
'		 If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//�w�N�̃R���{���擾
		Call s_MakeGakunenWhere()

		'//�w�Ȃ̃R���{���擾
		Call s_MakeGakkaWhere()

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

	If Request("txtNendo") = "" Then
		m_iNendo   = session("NENDO")

		'//���N�x��񂪂���ꍇ�́A���N�x�̋��ȏ��o�^�������ݒ�̑ΏۂƂ���
		If m_bJinendoGakki = True Then
			m_iNendo   = m_iNendo + 1
		End If
	Else
		m_iNendo = Request("txtNendo")
	End If

	m_iGakunen = Request("txtGakunenCd")
	m_sGakkaCd = Request("txtGakkaCD")

End Sub

'********************************************************************************
'*	[�@�\]	�f�o�b�O�p
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iNendo   = " & m_iNendo	& "<br>"
	response.write "m_iGakunen = " & m_iGakunen & "<br>"
	response.write "m_sGakkaCd = " & m_sGakkaCd & "<br>"

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

		// ������NULL����������
		// ���N�x
		if( f_Trim(document.frm.txtNendo.value) == "" ){
			window.alert("�N�x�̑I�����s���Ă�������");
			document.frm.txtNendo.focus();
			return ;
		}

		document.frm.action="web0320_main.asp";
		document.frm.target="main";
		document.frm.submit();
	
	}
	//************************************************************
	//	[�@�\]	�o�^�{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Touroku(){

		document.frm.action="./touroku.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.txtMode.value = "Touroku";
		document.frm.submit();
	
	}

	//************************************************************
	//	[�@�\]	�����Q�ƑI����ʃE�B���h�E�I�[�v��
	//	[����]
	//	[�ߒl]
	//	[����]
	//************************************************************
	function KyokanWin(p_iInt,p_sKNm) {
		var obj=eval("document.frm."+p_sKNm)

		URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
		nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=600,top=0,left=0");
		nWin.focus();
		return true;	
	}

	//************************************************************
	//	[�@�\]	�N���A�{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function fj_Clear(){

		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}

	//************************************************************
	//	[�@�\]	�N�x���ύX���ꂽ�Ƃ��A�{��ʂ��ĕ\��
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_ReLoadMyPage(){

		document.frm.action="./web0320_top.asp";
		document.frm.target="_self";
		document.frm.submit();

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
						<option VALUE="<%= w_iNen + 1 %>" <%=f_Selected(Request("txtNendo"),w_iNen + 1)%> ><%= w_iNen + 1 %>
						<option VALUE="<%= w_iNen %>"	  <%=f_Selected(Request("txtNendo"),w_iNen)%> ><%= w_iNen %>
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
				call gf_ComboSet_Gakka("txtGakkaCD",C_CBO_M02_GAKKA,m_sGakkaWhere,"style='width:175px;' ",True,m_sGakkaCd)%>
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
	</form>
	</center>
	</body>
	</html>

<%
	'---------- HTML END   ----------
End Sub
%>
