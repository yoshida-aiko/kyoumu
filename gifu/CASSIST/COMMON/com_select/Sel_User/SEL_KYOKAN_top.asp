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
    Const DebugFlg = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iNendo             '�N�x
    Public m_sKyokanCd          '��������
    Public m_iI                 '
    Public m_sKNm               '������
    Public m_sGakkaCd               '�����w��

    Public m_sUserKbn
    Public m_sSimei
    Public m_sGakkaOption
    Public m_sKeiretuOption
    Public m_sKyokaKbnOption

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
    w_sMsgTitle="���p�ґI�����"
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

	m_sUserKbn = Replace(request("UserKbn"),"@@@","")
	m_sSimei   = request("txtSimei")

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

		'//�Ώێҋ敪�������ȊO�̏ꍇ�́A�����敪����I���ł��Ȃ��悤�ɂ���
		Call s_CtrlDisabled()

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

'********************************************************************************
'*  [�@�\]  �R���{�{�b�N�X��DISABLED���Z�b�g
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_CtrlDisabled()

	m_sGakkaOption=""
	m_sKeiretuOption=""
	m_sKyokaKbnOption=""

	'//�����ȊO�̏ꍇ
	If cint(gf_SetNull2Zero(m_sUserKbn)) <> C_USER_KYOKAN Then
		m_sGakkaOption="DISABLED"
		m_sKeiretuOption="DISABLED"
		m_sKyokaKbnOption="DISABLED"
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
<HTML>
<BODY>

<link rel="stylesheet" href="../../style.css" type="text/css">
    <title>���p�ґI�����</title>

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
    //************************************************************
    //  [�@�\]  �۰�ގ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./SEL_KYOKAN_top.asp";
        document.frm.target="_self";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" action="post" onsubmit="return false">

<br>
<% 
    call gs_title("���p�ґI�����","��@��")
%>
</TD>
</TR>
</TABLE>

<table border="0" width="100%">
    <tr>
        <td width="100%">
            <table border="0" class="search" cellpadding="1" cellspacing="1" width="100%">
                <tr>
                    <td align="left" class="search" nowrap>
                        <table border="0" cellpadding="1" cellspacing="1" width="100%">
                            <tr>
                                <td align="left" width="16%" Nowrap>���p�ҋ敪</td>
                                <td width="34%" Nowrap colspan="1">
                                <%
                                call gf_ComboSet("UserKbn",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_USER & " AND M01_NENDO=" & m_iNendo & " AND M01_SYOBUNRUI_CD>0","onchange='javascript:f_ReLoadMyPage()'",True,Request("UserKbn"))
                                %>
                                </td>

                                <td align="left" width="16%" Nowrap>�w�@��</td>
                                <td width="34%" Nowrap colspan="3">
                                <%
                                call gf_ComboSet("Gakka",C_CBO_M02_GAKKA,"M02_NENDO='" & m_iNendo & "'",m_sGakkaOption,True,m_sGakkaCd)
                                %>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" width="16%" Nowrap>�����敪</td>
                                <td width="34%" Nowrap>
                                <%
                                call gf_ComboSet("KkanKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD='" & C_KYOKAN &"' AND M01_NENDO='" & m_iNendo & "'",m_sKyokaKbnOption,True,"")
                                %>
                                </td>
                                <td align="left" width="16%" Nowrap>���Ȍn��敪</td>
                                <td width="34%" Nowrap>
                                <%
                                call gf_ComboSet("KkeiKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_KYOKA_KEIRETU &" AND M01_NENDO=" & m_iNendo & " AND M01_SYOBUNRUIMEI IS NOT NULL ",m_sKeiretuOption,True,"")
                                %>
                                </td>

                            </tr>
                            <tr>

                                <td align="left" width="16%" Nowrap>����</td>
                                <td width="34%"  colspan="1" Nowrap><input type="text" name="txtSimei" size="25" value="<%=Request("txtSimei")%>"></td>
                                <td align="left" width="16%" Nowrap><br></td>
						        <td align="right" width="34%" Nowrap ><input class="button" type="button" value="�@�\�@���@" onClick = "javascript:f_Hyouji()"></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<span class="CAUTION">�� ���p�ҋ敪�������̏ꍇ�̂݁A�w�ȁA�����敪�A���Ȍn��敪���I���\�ƂȂ�܂��B </span>
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