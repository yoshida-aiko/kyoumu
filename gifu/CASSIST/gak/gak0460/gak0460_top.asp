<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w���v�^�������o�^
' ��۸���ID : gak/gak0460/gak0460_top.asp
' �@      �\: ��y�[�W �w���v�^�������o�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/18 �O�c �q�j
' ��      �X: 2001/08/07 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'           : 2001/08/09 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'           : 2001/08/30 �ɓ� ���q     ����������2�d�ɕ\�����Ȃ��悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�s�����I��p��Where����
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '�����R�[�h
    Public m_sGakuNo        '�����R���{�{�b�N�X�ɓ���l
    Public m_sGakuNoWhere   '�����R���{�{�b�N�X��where����

    Public m_Rs
    Public m_iDsp          '�ꗗ�\���s��
	
	Public m_sNendoWhere
	Public m_sNendo
	Public m_RsN
	Public m_sOption
	
	Public m_sGakunen
	Public m_sClass
	Public m_sClassNm
	
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
	Dim w_iRet              '// �߂�l
	Dim w_sSQL              '// SQL��
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="�w���v�^�������o�^"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_top"
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = False
	
	m_iNendo    = session("NENDO")
	m_sKyokanCd = session("KYOKAN_CD")
	m_sGakuNo   = request("txtGakuNo")
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
		
		'�ݒ�N���X�R���{�쐬
		If f_NendoWhere() <> 0 Then m_bErrFlg = True : Exit Do
		
		'//�w�N�̑Ώۂ̃f�[�^�擾
		'If f_getData() <> 0 Then m_bErrFlg = True : Exit Do
		
		Call f_GakuNoWhere()
		
		'// �y�[�W��\��
		Call showPage()
		Exit Do
	Loop
	
	'// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// �I������
	Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*  [�@�\]  �w�N�̑Ώۂ̃f�[�^�擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_getData()
	On Error Resume Next
	Err.Clear
	f_getData = 1
	
	Do
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & "     M05_CLASS "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & "     M05_NENDO = '" & m_iNendo & "' "
		w_sSQL = w_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "
		
		Set m_Rs = Server.CreateObject("ADODB.Recordset")
		
		If gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp) <> 0 Then
			'ں��޾�Ă̎擾���s
			f_getData = 99
			m_bErrFlg = True
			Exit Do 
		End If
		
		f_getData = 0
		Exit Do
	Loop
	
	'// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If

End Function

'********************************************************************************
'*  [�@�\]  �ݒ�N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_NendoWhere()
	
	On Error Resume Next
	Err.Clear
	f_NendoWhere = 1
	
	Do
		m_sNendoWhere=""
		m_sNendoWhere = " M05_NENDO > " & m_iNendo - 5 & "  AND "
		m_sNendoWhere = m_sNendoWhere & " M05_NENDO <= " & m_iNendo & "  AND "
		m_sNendoWhere = m_sNendoWhere & " M05_TANNIN = '" & m_sKyokanCd & "' "
		
		m_sNendo = request("txtNendo")
		
		If request("txtNendo") = C_CBO_NULL Then m_sNendo = ""
		
		If m_sNendo <> "" Then
			
			w_sSQL = ""
			w_sSQL = w_sSQL & " SELECT "
			w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
			w_sSQL = w_sSQL & " FROM "
			w_sSQL = w_sSQL & "     M05_CLASS "
			w_sSQL = w_sSQL & " WHERE"
			w_sSQL = w_sSQL & "     M05_NENDO = '" & m_sNendo & "' "
			w_sSQL = w_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "
			
			Set m_RsN = Server.CreateObject("ADODB.Recordset")
			
			If gf_GetRecordsetExt(m_RsN, w_sSQL, m_iDsp) <> 0 Then
				'ں��޾�Ă̎擾���s
				f_NendoWhere = 99
				m_bErrFlg = True
				Exit Do 
			End If
			
			m_sGakunen	= m_RsN("M05_GAKUNEN")
			m_sClass	= m_RsN("M05_CLASSNO")
			m_sClassNm	= m_RsN("M05_CLASSMEI")
			
		End If
		
		f_NendoWhere = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub f_GakuNoWhere()
	
	m_sGakuNoWhere=""
	m_sOption=""
	
	If m_sNendo <> "" Then
		If m_RsN.EOF Then
			m_sOption = " DISABLED "
			m_sGakuNoWhere  = " T11_GAKUSEI_NO = '' "
		Else
			m_sGakuNoWhere = " T11.T11_GAKUSEI_NO = T13.T13_GAKUSEI_NO AND "
			'm_sGakuNoWhere = m_sGakuNoWhere & " T11.T11_NYUNENDO = T13.T13_NENDO - T13.T13_GAKUNEN + 1 AND "
			m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_GAKUNEN = " & m_sGakunen & " AND "
			m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_CLASS = " & m_sClass & " AND "
			m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_NENDO = " & m_sNendo & " "
		End If
	Else
		m_sOption = " DISABLED "
		m_sGakuNoWhere  = " T11_GAKUSEI_NO = '' "
	End IF

End Sub

'********************************************************************************
'*		�w�N�A�N���X���Z�b�g
'********************************************************************************
Sub f_Syosai()
	
	If m_sNendo = "" Then
		response.write "<td width='30' Nowrap>�@</td>"
		response.write "<td width='90' Nowrap>�@</td>"
	Else
		response.write "<td width='30' Nowrap align='right'>" & m_sGakunen & "�N</td>"
		response.write "<td width='90' Nowrap>" & m_sClassNm & "</td>"
	End If
	
End Sub

'Sub f_GakuNoWhere()
''********************************************************************************
''*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����
''*  [����]  �Ȃ�
''*  [�ߒl]  �Ȃ�
''*  [����]  
''********************************************************************************
'
'    m_sGakuNoWhere=""
'
'    m_sGakuNoWhere = " T11_GAKUSEI_NO = T13_GAKUSEI_NO AND "
'    m_sGakuNoWhere = m_sGakuNoWhere & " T13_GAKUNEN = " & m_Rs("M05_GAKUNEN") & " AND "
'    m_sGakuNoWhere = m_sGakuNoWhere & " T13_CLASS = " & m_Rs("M05_CLASSNO") & " AND "
'    m_sGakuNoWhere = m_sGakuNoWhere & " T13_NENDO = " & m_iNendo & " "
'
'End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	On Error Resume Next
	Err.Clear
%>
<html>
<head>
<title>�w���v�^�������o�^</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [�@�\]  �N�x���C�����ꂽ�Ƃ��A�ĕ\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){
		document.frm.action="gak0460_top.asp";
        document.frm.target="topFrame";
        document.frm.submit();
    }
    
    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){
		// ���N�x
		if( f_Trim(document.frm.txtNendo.value) == "" ){
			alert("�N�x�̑I�����s���Ă�������");
			document.frm.txtNendo.focus();
			return ;
		}
		
		// ���N�x
		if( f_Trim(document.frm.txtNendo.value) == "<%=C_CBO_NULL%>" ){
			alert("�N�x�̑I�����s���Ă�������");
			document.frm.txtNendo.focus();
			return ;
		}
		
		// ���w��
		if(f_Trim(document.frm.txtGakuNo.value) == "" ){
			if(document.frm.txtGakuNo.length == 1) {
				alert("�w��N�x�̊w���̃f�[�^������܂���");
				document.frm.txtNendo.focus();
			}else{
				alert("�w���̑I�����s���Ă�������");
				document.frm.txtGakuNo.focus();
			}
			
			return ;
		}
		
		// ���w��
		if(f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
			if(document.frm.txtGakuNo.length == 1) {
				alert("�w��N�x�̊w���̃f�[�^������܂���");
				document.frm.txtNendo.focus();
			}else{
				alert("�w���̑I�����s���Ă�������");
				document.frm.txtGakuNo.focus();
			}
			return ;
		}
		
		document.frm.action="gak0460_main.asp";
		document.frm.target="main";
		document.frm.submit();
	}
	
	/*
	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���̏���
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//************************************************************
	function f_Search2(){
		// ���w��
		if( f_Trim(document.frm.txtGakuNo.value) == "" ){
			alert("�w���̑I�����s���Ă�������");
			document.frm.txtGakuNo.focus();
			return ;
		}
		
		// ���w��
		if( f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
			alert("�w���̑I�����s���Ă�������");
			document.frm.txtGakuNo.focus();
			return ;
		}
		
		document.frm.action="gak0460_main.asp";
		document.frm.target="main";
		document.frm.submit();
	}
	*/
	
	//************************************************************
	//  [�@�\]  �N���A�{�^���N���b�N���̏���
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_Clear(){
		document.frm.txtGakuNo.value = "";
	}
	
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>

<body>
<center>
<form name="frm" METHOD="post" onClick="return false;">

<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
<td valign="top" align="center">
<%call gs_title("�w���v�^�������o�^","�o�@�^")%>
<br>
	<table border="0">
		<tr>
			<td class="search">
				<table border="0" cellpadding="1" cellspacing="1">
					<tr>
						<td align="left">
							<table border="0" cellpadding="1" cellspacing="1">
								<tr valign="bottom">
									<td Nowrap>
										<%call gf_ComboSet("txtNendo",C_CBO_M05_CLASS_N,m_sNendoWhere,"style='width:70px;' onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sNendo)%>�N�x</td>
									</td>
									<!--td Nowrap align="center">�@�N���X�@</td-->
									<% Call f_Syosai() %>
									<!--td Nowrap><%=m_Rs("M05_GAKUNEN")%>�N</td-->
									<!--td Nowrap><%=m_Rs("M05_CLASSMEI")%></td-->
									<td Nowrap align="center">�@���@���@
										<%call gf_PluComboSet("txtGakuNo",C_CBO_T11_GAKUSEKI_N,m_sGakuNoWhere,"style='width:250px;' "& m_sOption,True,m_sGakuNo)%>
										<!--%call gf_PluComboSet("txtGakuNo",C_CBO_T11_GAKUSEKI_N,m_sGakuNoWhere, "style='width:250px;'",True,m_sGakuNo)%-->
									</td>
								</tr>
								
								<tr>
									<td colspan="4" align="right">
										<input type="button" class="button" value=" �N�@���@�A " onclick="javasript:f_Clear();">
										<input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</td>
</tr>
</table>

<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
<input type="hidden" name="txtClass" value="<%=m_sClass%>">
<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">

</form>

</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
