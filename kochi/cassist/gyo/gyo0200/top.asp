<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �s�������ꗗ
' ��۸���ID : gyo/gyo0200/top.asp
' �@      �\: ��y�[�W �s�������̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           cboGyojiDate      :�s�����t
'           chkGyojiCd      :�s���R�[�h
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͌���\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��s���ꗗ��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/26 ���{ ����
' ��      �X: 2001/07/27 �ɓ����q�@M40_CALENDER�e�[�u���폜�ɑΉ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bErrMsg           '�װү����
    
    '���I��p��Where����
    Public m_sGyojiMWhere        '���̏���
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen      ':�����N�x
    Public  m_iKyokanCd      ':�����R�[�h
    
    '�f�[�^�擾�p
    Public  m_iDate             ':�����̓��t(yyyy/mm/dd)
    Public  m_iDay              ':�����̓�

    Public  m_iTuki		'//����

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
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�s�������ꗗ"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False


        '// ���Ұ�SET
		Call s_SetParam()

        '// ���tSET
		Call s_SetDate()
         

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_bErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
        
        '���R���{�Ɋւ���WHERE���쐬����
        'Call s_MakeGyojiMWhere() 
        
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
    Call gs_CloseDatabase()
End Sub

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub

Sub s_SetDate()
'********************************************************************************
'*  [�@�\]  �����̓��t��ݒ�i�R���{�{�b�N�X�p�j
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

'    m_iDate = gf_YYYY_MM_DD(date(),"/")
'    m_iDay = Day(m_iDate)

	m_iTuki = month(date())

End Sub

'****************************************************
'[�@�\]	�f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'		(���X�g�_�E���{�b�N�X�I��\���p)
'[����]	pData1 : �f�[�^�P
'		pData2 : �f�[�^�Q
'[�ߒl]	f_Selected : "SELECTED" OR ""
'					
'****************************************************
Function f_Selected(pData1,pData2)

	If IsNull(pData1) = False And IsNull(pData2) = False Then
		If trim(cStr(pData1)) = trim(cstr(pData2)) Then
			f_Selected = "selected"	
		Else 
			f_Selected = ""	
		End If
	End If

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
<link rel=stylesheet href="../../common/style.css" type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

    }

    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_BackClick(){

        document.frm.action="../../menu/sansyo.asp";
        document.frm.target="_parent";
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

        document.frm.action="main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" action="./main.asp" target="main" Method="POST">
    <input type="hidden" name="txtMode">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
    <%call gs_title("�s�������ꗗ","��@��")%>
<%
If m_sMode = "" Then
%>
    <table border="0">
    <tr>
    <td class=search>
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left" >
            <table border="0" cellpadding="1" cellspacing="1">
            <tr>
            <td align="left" >
            <% 'call gf_calComboSet("cboGyojiDate",C_CBO_M40_CALENDER,m_sGyojiMWhere,"",False,m_iDate,"mm") %>
			<select name="cboGyojiDate">
				<option value="4"  <%=f_Selected("4" ,cstr(m_iTuki))%> >4
				<option value="5"  <%=f_Selected("5" ,cstr(m_iTuki))%> >5
				<option value="6"  <%=f_Selected("6" ,cstr(m_iTuki))%> >6
				<option value="7"  <%=f_Selected("7" ,cstr(m_iTuki))%> >7
				<option value="8"  <%=f_Selected("8" ,cstr(m_iTuki))%> >8
				<option value="9"  <%=f_Selected("9" ,cstr(m_iTuki))%> >9
				<option value="10" <%=f_Selected("10",cstr(m_iTuki))%> >10
				<option value="11" <%=f_Selected("11",cstr(m_iTuki))%> >11
				<option value="12" <%=f_Selected("12",cstr(m_iTuki))%> >12
				<option value="1"  <%=f_Selected("1" ,cstr(m_iTuki))%> >1
				<option value="2"  <%=f_Selected("2" ,cstr(m_iTuki))%> >2
				<option value="3"  <%=f_Selected("3" ,cstr(m_iTuki))%> >3
			</select>��
            </td>
            <td align="left" >&nbsp;&nbsp;&nbsp;<input type="checkbox" name="chkGyojiCd">�s���̂ݕ\��</td>
            </tr>
            </table>
        </td>
        </tr>
        </table>
    </td>
    <td valign="bottom">
    <input type="button" value="�@�\�@���@" onClick="javascript:f_Search()" class=button>
    </td>
    </tr>
    </table>
<%
End IF
%>
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
