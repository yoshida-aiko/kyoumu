<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������ԋ����\��o�^
' ��۸���ID : skn/skn0120/top.asp
' �@      �\: ��y�[�W �\��o�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtSikenKbn      :�����敪
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtSikenKbn      :�����敪
'           txtSikenCd      :�����R�[�h�i���́E�ǎ���//A:1,B:2�j
'           txtMode         :���샂�[�h
'                               BLANK   :�����\��
'                               Reroad  :�i�����I����j�ĕ\��
'                               Search  :����
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͎������̂�\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�����������ɂ��Ȃ������\���\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/18 ���u �m��
' ��      �X: 2001/07/24 �{��
'           : 2001/08/02 ���{ ����  '�����R���{�\���ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�����I��p��Where����
    Public m_sSikenWhere        '�����̏���
    Public m_sSikenOption       '�����R���{�̃I�v�V����
    Public  m_sSikenCdWhere     '�����R���{�̃I�v�V�����i�����R�[�h�j
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iSikenKbn      ':�����敪
    Public  m_iSyoriNen      ':�����N�x
    Public  m_iKyokanCd      ':�����R�[�h

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
    w_sMsgTitle="�������ԋ����\��o�^"
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

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '//���݂̓��t�Ɉ�ԋ߂������敪���擾
        '//(�����\���͌��݂̓��t�Ɉ�ԋ߂������ł̎��Ԋ��ꗗ��\������)
        'If trim(m_iSikenKbn) = "" Then
        If m_sTxtMode = "" Then
            w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_SEISEKI_KIKAN,0)
            'w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_JISSI_KIKAN,0)
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If
        End If
        
        '�����R���{�Ɋւ���WHERE���쐬����
        Call s_MakeSikenWhere() 
        
        '�����R���{�Ɋւ���WHERE���쐬����
        Call s_MakeSikenCdWhere() 
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

    m_iSikenKbn = ""
    m_iSikenKbn = Request("txtSikenKbn")
    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub


Sub s_MakeSikenWhere()
'********************************************************************************
'*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    
    m_sSikenWhere = ""
    
    m_sSikenWhere = m_sSikenWhere & " M01_NENDO = " & m_iSyorinen
    m_sSikenWhere = m_sSikenWhere & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
    m_sSikenWhere = m_sSikenWhere & " AND M01_SYOBUNRUI_CD <= 4 "						'<!--8/16�C��

'response.write("<BR>m_sSikenWhere = " & m_sSikenWhere)

End Sub

Sub s_MakeSikenCdWhere()
'********************************************************************************
'*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����i�����R�[�h�j
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************


    m_sSikenCdWhere = ""
    m_sSikenOption = ""
    
'--2001/07/15 CONST�ɕύX
        'C_SIKEN_JITURYOKU = 5  '���͎���
        'C_SIKEN_TUISI = 6      '�ǎ���

    If cint(m_iSikenKbn) = Cint(C_SIKEN_JITURYOKU) or cint(m_iSikenKbn) = cInt(C_SIKEN_TUISI)  Then
        m_sSikenCdWhere = m_sSikenCdWhere & " M27_NENDO = " & m_iSyoriNen
        m_sSikenCdWhere = m_sSikenCdWhere & " AND M27_SIKEN_KBN = " & m_iSikenKbn
    else
        m_sSikenOption = " DISABLED "
    End If

'   response.write("<BR>m_sSikenCdWhere = " & m_sSikenCdWhere)
    
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
<link rel=stylesheet href=../../common/style.css type=text/css>
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

        document.frm.action="../../menu/siken.asp";
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

        document.frm.txtMode.value = "Search";
        document.frm.action="skn0130_main.asp";
//        document.frm.action="default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        //document.frm.target="main";
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

        document.frm.txtMode.value = "";
        document.frm.action="skn0130_main.asp";
        document.frm.target="main";
        document.frm.submit();

        document.frm.action="skn0130_top.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" action="" target="main" Method="POST">
<input type="hidden" name="txtMode">

<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
<%call gs_title("�������{�Ȗړo�^","��@��")%>
<br>
    <table border="0">
    <tr>
    <td class=search >
		<table border="0" cellpadding="1" cellspacing="1">
		<tr>
		<td align="left" >
		�@<br>
		</td>
		<td align="left" >
		<!--% call gf_ComboSet("txtSikenKbn",C_CBO_M01_KUBUN,m_sSikenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:150px;'",False,m_iSikenKbn) %-->
        <% call gf_ComboSet("txtSikenKbn",C_CBO_M01_KUBUN,m_sSikenWhere," style='width:150px;' ",false,m_iSikenKbn) %>
		</td>
		<td align="left" >
		�@<br>
		</td>
		<td valign="bottom" align="right"><input class="button" type="button" onclick="javascript:f_Search();" value="�@�\�@���@"></td>
<!-- 
		<td align="left" >&nbsp;&nbsp;
		<% call gf_ComboSet("txtSikenCd",C_CBO_M27_SIKEN,m_sSikenCdWhere,m_sSikenOption & " style='width:120px;' ",True,"") %>
		</td>
//-->
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




