<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����Q�ƑI�����
' ��۸���ID : Common/com_select/SEL_KYOKAN/SEL_KYOKAN_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/19 �O�c
' ��      �X: 2001/08/08 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iNendo             '�N�x
    Public m_sKyokanCd          '��������
    Public m_iI                 '
    Public m_sKNm               '������
    Public m_sGakkaCd               '�����w��

    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

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
    w_sMsgTitle="�����Q�ƑI�����"
    w_sMsg=""
    w_sRetURL="../../../../default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iI        = request("txtI")
    m_sKNm      = request("txtKNm")
    m_sGakkaCd      = request("txtGakka")

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If
        
        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Rs)
    '// �I������
    Call gs_CloseDatabase()
End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
<HTML>
<BODY>

<link rel="stylesheet" href="../../style.css" type="text/css">
    <title>�����Q�ƑI�����</title>

    <!--#include file="../../jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [�@�\]  �\���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Hyouji(){

        //���X�g����submit
        document.frm.target = "main" ;
        document.frm.action = "SEL_KYOKAN_main.asp";
        document.frm.submit();

    }
    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" action="post" onClick = "return;false;">

<br>
<% 
    call gs_title("�����Q�ƑI�����","��@��")
%>
</TD>
</TR>
</TABLE>

<table border="0" width="100%">
    <tr>
<!--
        <td width="100%">
            <table border="0" class="search" cellpadding="1" cellspacing="1" width="100%">
                <tr>
-->
                    <td align="left" class="search" nowrap>
                        <table border="0" cellpadding="1" cellspacing="1" width="100%">
                            <tr>
                                <td align="center" width="16%" Nowrap>�w�@��</td>
                                <td width="84%" Nowrap colspan="3">
                                <%
                                call gf_ComboSet("Gakka",C_CBO_M02_GAKKA,"M02_NENDO='" & m_iNendo & "'","",True,m_sGakkaCd)
                                %>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" width="16%" Nowrap>�����敪</td>
                                <td width="34%" Nowrap>
                                <%
                                call gf_ComboSet("KkanKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD='" & C_KYOKAN &"' AND M01_NENDO='" & m_iNendo & "'","",True,"")
                                %>
                                </td>
                                <td align="center" width="26%" Nowrap>���Ȍn��敪</td>
                                <td width="24%" Nowrap>
                                <%
                                'call gf_ComboSet("KkeiKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD='" & C_KYOKA_KEIRETU &"' AND M01_NENDO='" & m_iNendo & "'","",True,"")
                                call gf_ComboSet("KkeiKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_KYOKA_KEIRETU &" AND M01_NENDO=" & m_iNendo & " AND M01_SYOBUNRUIMEI IS NOT NULL ","",True,"")


                                %>
                                </td>
                            </tr>
                        </table>
                    </td>
<!--
                </tr>
            </table>
        </td>
-->
    </tr>
    <tr>
        <td align="right"><input class="button" type="button" value="�@�\�@���@" onClick = "javascript:f_Hyouji()"></td>
    </tr>
</table>
    <input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
    <input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
    <input type="hidden" name="txtI"        value="<%=m_iI%>">
    <input type="hidden" name="txtKNm"      value="<%=m_sKNm%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>