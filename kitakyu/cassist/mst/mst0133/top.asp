<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �i�H���񌟍�
' ��۸���ID : mst/mst0133/top.asp
' �@      �\: ��y�[�W �A�E��}�X�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           txtSinroCD      :�i�H�敪
'           txtSingakuCd    :�i�w�敪
'           txtSinroName        :�A�E�於�́i�ꕔ�j

'           :txtSyusyokuName        :�A�E�於�́i�ꕔ�j '/2001/07/31�ǉ�
'           :txtMode                :���[�h             '/2001/07/31�ǉ�
'           :txtFLG                 :                   '/2001/07/31�ǉ�
'           :txtSNm                 :                   '/2001/07/31�ǉ�
'           :txtNendo               :�N�x               '/2001/07/31�ǉ�
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��A�E���\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/15 �≺�@�K��Y
' ��      �X: 2001/07/31 ���{ ����  �����E���n�ǉ�
'           :                       �i�H�於�̃e�L�X�g�{�b�N�XMAXLENGTH�ǉ�
'           :                       �ϐ��������K���Ɋ�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    '�s�����I��p��Where����
    Public m_sSinroWhere    '�i�H�̏���
    Public m_sSingakuWhere  '�i�w�R���{�̏���
    Public m_sSingakuOption '�i�w�R���{�̃I�v�V����
    Public m_sSyusyokuName  ':�A�E�於�́i�ꕔ�j
    Public m_iSinroCD       ':�i�H�敪      '/2001/07/31�ύX
    Public m_iSingakuCd     ':�i�w�敪      '/2001/07/31�ύX
    Public m_iNendo         ':�N�x
    Public m_sMode          ':���[�h
    Public m_iFLG
    Public m_sSNm

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
    w_sMsgTitle="�i�H���񌟍�"
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

        '// ���Ұ�SET
        Call s_SetParam()
        '�i�H�Ɋւ���WHRE���쐬����
        Call f_MakeSinroWhere() 
        '�i�w&�A�E�Ɋւ���WHRE���쐬����
        Call f_MakeSingakuWhere()   

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
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSinroCD = Request("txtSinroCD")              ':�i�H�敪
    m_iSingakuCd = Request("txtSingakuCD")          ':�i�w�敪
    m_sSyusyokuName = Request("txtSyusyokuName")    ':�A�E�於�́i�ꕔ�j
    m_sMode = request("txtMode")                    ':���[�h    
    m_iNendo = Session("NENDO")                     ':�N�x
    m_iFLG = request("txtFLG")
    m_sSNm = request("txtSNm")
End Sub


Sub f_MakeSinroWhere()
'********************************************************************************
'*  [�@�\]  �i�H�R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sSinroWhere=""

    m_sSinroWhere = " M01_DAIBUNRUI_CD = " & C_SINRO & "  AND "
    m_sSinroWhere = m_sSinroWhere & " M01_NENDO = " & m_iNendo & ""

End Sub

Sub f_MakeSingakuWhere()
'********************************************************************************
'*  [�@�\]  �i�w�R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sSingakuWhere=""
    m_sSingakuOption=""

	'// �i�w
    If m_iSinroCD = "1" Then
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_SINGAKU & "  AND "
        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
	'// �A�E
	ElseIf m_iSinroCD = "2" Then
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_GYOSYU_KBN & "  AND "
        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
	'// ���̑�
    Else
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = 0 "
        m_sSingakuOption = " DISABLED "
    End IF

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

<title>�i�H���񌟍�</title>
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
    //  [�@�\]  �i�H���C�����ꂽ�Ƃ��A�ĕ\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="top.asp";
        document.frm.target="_self";
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

        document.frm.action="./main.asp";
        document.frm.target="main";
        document.frm.txtMode.value = "Search";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �N���A�{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Clear(){

        document.frm.txtSinroCD.value = "@@@";
        document.frm.txtSingakuCd.value = "@@@";
        document.frm.txtSyusyokuName.value = "";
    
    }

    //-->
    </SCRIPT>

    <link rel=stylesheet href="../../common/style.css" type=text/css>

    </HEAD>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<center>

<form name="frm" Method="POST"  onSubmit="return false" onClick="return false;">

<%
call gs_title("�i�H���񌟍�","��@��")
%>
<br>
    <table border="0">
    <tr>
    <td>

        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left" class=search>

                <table border="0" cellpadding="1" cellspacing="1">
                <tr>
                <td Nowrap align="left">
            �i�H�敪<img src="../../image/sp.gif" width="15">
<%          '���ʊ֐�����i�H�Ɋւ���R���{�{�b�N�X���o�͂���i�N�x�����j
            call gf_ComboSet("txtSinroCD",C_CBO_M01_KUBUN,m_sSinroWhere,"onchange = 'javascript:f_ReLoadMyPage()' ",True,m_iSinroCD)%>

                </td>
                <td Nowrap align="left">��ʋ敪

<%          '���ʊ֐�����i�w�Ɋւ���R���{�{�b�N�X���o�͂���i�N�x�A�i�H�敪�������j�i�i�H�敪�����͂���Ă��Ȃ��Ƃ��́ADISABLED�ƂȂ�j
            call gf_ComboSet("txtSingakuCd",C_CBO_M01_KUBUN,m_sSingakuWhere,m_sSingakuOption & " style='width:100px;'",True,m_iSingakuCd)%>
                </td>
                </tr>

                <tr>
                <td align="left" colspan="2" nowrap>
                �i�H�於��
                <input type="text" name="txtSyusyokuName" size="20" Value="<%=m_sSyusyokuName%>" maxlength="60">   <!--'//2001/07/31�C��-->
	            <font size="2">���i�H�於�̂̈ꕔ�Ō������܂�</font>
                </td>
                </tr>
				<tr>
					<td valign="bottom" align="right" colspan="2">
			        <input type="button" class="button" value=" �N�@���@�A " onclick="javasript:f_Clear();">
					<input class="button" type="button" value="�@�\�@���@" onClick = "javascript:f_Search()">
					</td>
				</tr>
                </table>
	        </td>
        </tr>
        </table>
    </td>
  </tr>
</table>
<input type="hidden" name="txtFLG" value="<%=m_iFLG%>">
<input type="hidden" name="txtSNm" value="<%=m_sSNm%>">
<input type="hidden" name="txtMode" value="<%=m_sMode%>">
<input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
</form>

</center>

</body>

</html>






<%
    '---------- HTML END   ----------
End Sub
%>
