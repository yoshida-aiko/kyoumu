<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �N���X�ʎ��Ǝ��Ԉꗗ
' ��۸���ID : jik/jik0210/top.asp
' �@      �\: ��y�[�W �N���X�ʎ��Ǝ��Ԃ̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           cboGakunenCd      :�w�N�R�[�h
'           cboClassCd      :�N���X�R�[�h
'           txtMode         :���샂�[�h
'                               BLANK   :�����\��
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͊w�N�ƃN���X��\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ����ƈꗗ��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/06 ���{ ����
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�I��p��Where����
    Public m_sGakunenWhere        '�w�N�̏���
    Public m_sClassWhere        '�N���X�̏���
    Public m_sClassOption          ':�N���X�R���{�̃I�v�V����
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen      ':�����N�x
    Public  m_iKyokanCd      ':�����R�[�h
    Public  m_iGakunen      ':�w�N�R�[�h
    
    '�f�[�^�擾�p

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
    w_sMsgTitle="�N���X�ʎ��Ǝ��Ԉꗗ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False


        '// ���Ұ�SET
        Call s_SetParam()
         

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

        
        '�w�N�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeGakunenWhere() 
        '�N���X�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeClassWhere() 
        
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
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iSyoriNen = Session("NENDO")
    'm_iSyoriNen = 2001      '//�e�X�g�p
    m_iKyokanCd = Session("KYOKAN_CD")

    m_iGakunen = ""
    m_iGakunen = Request("cboGakunenCd")


End Sub

'Sub s_MakeGakunenWhere()
'********************************************************************************
'*  [�@�\]  �w�N�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
'
'    
'    m_sGakuenWhere = ""
'    
'    m_sGakuenWhere = m_sGakuenWhere & " M05_NENDO = " & m_iSyorinen
'    'm_sGakuenWhere = m_sGakuenWhere & " M05_NENDO = " & 2000  '//�e�X�g�p
'End Sub

Sub s_MakeClassWhere()
'********************************************************************************
'*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sClassWhere = ""
    m_sClassOption = ""
    
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen
    if m_iGakunen <> "" Then
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & m_iGakunen
    end if
    
    if m_iGakunen = "" Then
        m_sClassOption = " DISABLED "
    end if
    
End Sub

Sub s_SetGakCbo()
'********************************************************************************
'*  [�@�\]  �w�N�R���{��Select��\��������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_iCount

    For w_iCount = 1 To 5
        response.write "<option value=" & w_iCount
            If CStr(m_iGakunen) = CStr(w_iCount) Then
                response.write " Selected "
            End If
        response.write " >" & w_iCount
    Next

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
    //************************************************************
    //  [�@�\]  �������I�����ꂽ�Ƃ��A�ĕ\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();
    
    }



    //-->
    </SCRIPT>

</head>
<body>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
    <input type="hidden" name="txtMode">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
    <%call gs_title("�N���X�ʎ��Ǝ��Ԉꗗ","��@��")%>
<%
If m_sMode = "" Then
%>
    <table border="0">
    <tr>
    <td>
        <table border="0" class=search cellpadding="1" cellspacing="1">
        <tr>
        <td align="left" class=search>
            <table border="0" cellpadding="1" cellspacing="1">
            <tr>
            <td align="left" class=search>
            �w�N
            </td>
            <td align="left" class=search>
            <% 'call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS,m_sGakuenWhere,"onchange = 'javascript:f_ReLoadMyPage()' ",False,m_iGakunen) %>
            <select name="cboGakunenCd" onchange = 'javascript:f_ReLoadMyPage()'>
                <option>
                <%Call s_SetGakCbo()%>
            </select>
            �N��
            </td>
            <td align="left" class=search>
            �N���X
            <% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,m_sClassOption,False,"") %>
            </td>
            </tr>
            </table>
        </td>
        <td><input type="button" value="�\��" onClick="javascript:f_Search()" class=button></td>
        </tr>
        </table>
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
