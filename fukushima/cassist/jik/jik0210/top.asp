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
'           :session("PRJ_No")      '���������̃L�[
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           cboGakunenCd      :�w�N�R�[�h
'           cboClassCd      :�N���X�R�[�h
'           txtMode         :���샂�[�h
'                           (BLANK) :�����\��
'                           Reload  :�����[�h
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͊w�N�ƃN���X��\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������̎��ƈꗗ��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/06 ���{ ����
' ��      �X: 2001/07/30 ���{ ����  �߂��URL�ύX
'           : 2001/08/09 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'           : 2015/03/20 ���{ ��H  Win7�Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�I��p��Where����
    Public m_sGakunenWhere      '�w�N�̏���
    Public m_sClassWhere        '�N���X�̏���
    Public m_sGakunenOption     ':�w�N�R���{�̃I�v�V����
    Public m_sClassOption       ':�N���X�R���{�̃I�v�V����
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iGakunen          ':�w�N�R�[�h
    Public  m_sMode             ':���샂�[�h
    
    '�f�[�^�擾�p
    Public  m_iTanninG          ':�S�C�i�w�N�j
    Public  m_iTanninC          ':�S�C�i�N���X�j

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
    w_sMsgTitle= "�N���X�ʎ��Ǝ��Ԉꗗ"
    w_sMsg= ""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget= ""


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

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '�S�C�w�N�E�N���X�擾
        Call f_GetTannin() 
        '// �w�N�R�[�h�擾
        Call s_SetGakunenCd()
        
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
    m_iKyokanCd = Session("KYOKAN_CD")

    m_sMode = Request("txtMode")

End Sub

Sub s_SetGakunenCd()
'********************************************************************************
'*  [�@�\]  �w�N�R�[�h��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    if m_sMode = "" Then
        m_iGakunen = m_iTanninG
    else
        m_iGakunen = Request("cboGakunenCd")
    end if
    
End Sub

Sub s_MakeGakunenWhere()
'********************************************************************************
'*  [�@�\]  �w�N�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    
    m_sGakunenWhere = ""
    m_sGakunenOption = ""
    
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyorinen
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"
    
    if m_sMode = "" Then
        m_sGakunenOption = m_iTanninG
    else
        m_sGakunenOption = m_iGakunen
    end if
    
End Sub

Sub s_MakeClassWhere()
'********************************************************************************
'*  [�@�\]  �N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sClassWhere = ""
    m_sClassOption = ""
    
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen
    if m_sMode = "" Then
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & m_iTanninG
    else
        if m_iGakunen <> "" Then
            m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & m_iGakunen
        end if
    end if

    if m_sMode = "" Then
        m_sClassOption = m_iTanninC
    end if
    
End Sub

'********************************************************************************
'*  [�@�\]  �S�C�w�N�E�N���X�̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾�����A99:���s
'*  [����]  
'********************************************************************************
Function f_GetTannin()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    f_GetTannin = 0
    m_iTanninG = ""
    m_iTanninC = ""

    Do

        '// �w�N�E�N���X�}�X�^���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "M05_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "M05_CLASSNO "
        w_sSQL = w_sSQL & vbCrLf & "FROM "
        w_sSQL = w_sSQL & vbCrLf & "M05_CLASS "
        w_sSQL = w_sSQL & vbCrLf & "WHERE "
        'w_sSQL = w_sSQL & vbCrLf & "M05_NENDO = 2200"
        w_sSQL = w_sSQL & vbCrLf & "M05_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND M05_TANNIN = '" & m_iKyokanCd & "'"
        
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
        
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            f_GetSentaku = 99
            Exit Do 'GOTO LABEL_f_GetTannin_END
        Else
        End If
        
        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            'm_bErrFlg = True
            'm_sErrMsg = "�Ώ�ں��ނȂ�"
            'f_GetTannin = 1
            m_iTanninG = 1
            m_iTanninC = 1

            Exit Do 'GOTO LABEL_f_GetTannin_END
        End If

            '// �擾�����l���i�[
            m_iTanninG = w_Rs("M05_GAKUNEN")
            m_iTanninC = w_Rs("M05_CLASSNO")
        '// ����I��
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs)

'// LABEL_f_GetTannin_END
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
<link rel="stylesheet" href="../../common/style.css" type="text/css">
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
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

        document.frm.action="top.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "Reload";
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
    <%call gs_title("�N���X�ʎ��Ǝ��Ԉꗗ","��@��")%>
<br>
    <table border="0">
    <tr>
    <td class="search">
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left">
            <table border="0" cellpadding="1" cellspacing="1">
            <tr valign="middle">
            <td align="left">
            �N���X
            </td>
            <td align="left">
            <% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage();' style='width:40px;' ",False,m_sGakunenOption) %>
            </td>
            <td align="left">
            �N
            </td>
            <td><img src="../../image/sp.gif" height="10"></td>
            <td align="left">
			<!-- 2015.03.20 Upd width:80->180 -->
            <% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"style='width:180px;' ",False,m_sClassOption) %>
            </td>
		        <td colspan="6" align="right">�@
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

</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
