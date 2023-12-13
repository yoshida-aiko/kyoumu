<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: ���юQ��
'* ��۸���ID : sei/sei0800/default_top.asp
'* �@      �\: 
'*-------------------------------------------------------------------------
'* ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*           :session("PRJ_No")      '���������̃L�[
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'* ��      ��:
'*           �������\��
'*               �R���{�{�b�N�X�͊w�N��\��
'*           ���\���{�^���N���b�N��
'*               ���̃t���[���Ɏw�肵�������̗��N�Y���҈ꗗ��\��������
'*-------------------------------------------------------------------------
'* ��      ��: 2003/05/15 �A�c
'* ��      �X: 2015/03/19 ���{ Win7�Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
	Dim m_bErrFlg			'�װ�׸�

    '�I��p��Where����
	Dim m_sGakunenWhere		'�w�N�̏���
	Dim m_sClassWhere		'�����̏���

    Dim m_sClassOption		'�N���X�R���{�̃I�v�V����
	Dim m_iNendo			'�����N�x
	Dim m_iGakunen
    Dim m_iClassNo
    
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
	Dim w_sSQL              '// SQL��
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle= "�L�����p�X�A�V�X�g"
	w_sMsgTitle= "���N�Y���҈ꗗ"
	w_sMsg= ""
	w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget= ""

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// �ް��ް��ڑ�
		if gf_OpenDatabase() <> 0 Then
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'// ���Ұ�SET
		Call s_SetParam()

		'�w�N�R���{�A�N���X�R���{�Ɋւ���WHERE���쐬����
		Call s_MakeGakunenWhere() 

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

Sub s_SetParam()
'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************

	m_iNendo    = Session("NENDO")			'//�����N�x
    m_iGakunen  = Request("cboGakunenCD")	'//�w�N
    m_iClassNo  = Request("cboClassCD")		'//�N���X

End Sub

Sub s_MakeGakunenWhere()
'********************************************************************************
'*  [�@�\]  �w�N�R���{�A�N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	'�w�N
	m_sGakunenWhere = ""
	m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iNendo
	m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

	'�N���X
	m_sClassWhere = ""
	m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iNendo

	If m_iGakunen = "" Then
		m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = 1"							'//�����\������1�N1�g��\��
	Else
		m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_iGakunen)		'//�I���N���X��\��
	End If

End Sub

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
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���̏���
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_Search(){
		with(document.frm){
			action="sei0800_listbottom.asp";
			target="main";
			submit();
		}
	}

	//************************************************************
	//  [�@�\]  �w�N���ύX���ꂽ�Ƃ��A�{��ʂ��ĕ\��
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
    function f_ReLoadMyPage(){
		with(document.frm){
			action="default_top.asp";
			target="topFrame";
			submit();
		}
	}

//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" method="post">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
	<tr>
		<td valign="top" align="center">
		<%call gs_title("���юQ��","�Q�@��")%>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td class="search">
					<table border="0" cellpadding="1" cellspacing="1">
						<tr>
							<td align="left">
								<table border="0" cellpadding="1" cellspacing="1">
									<tr>
										<td align="left" nowrap>�w�N</td>
										<td align="left" nowrap><% call gf_ComboSet("cboGakunenCD",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' ",False,m_iGakunen) %>�N</td>
										<td align="left" nowrap>�N���X</td>
										<!-- 2015.03.19 Upd width:80��180 -->
										<td align="left" nowrap><% call gf_ComboSet("cboClassCD",C_CBO_M05_CLASS,m_sClassWhere,"style='width:180px;' " & m_sClassOption,False,m_iClassNo) %></td>
									    <td valign="bottom"><input type="button" value="�@�\�@���@" onClick = "javascript:f_Search()" class="button"></td>
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

</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
