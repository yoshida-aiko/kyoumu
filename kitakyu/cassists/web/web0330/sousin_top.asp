<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A���f����
' ��۸���ID : web/web0330/sousin_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
'            ���[�h         ��      txtMode
'                                   �V�K = NEW
'                                   �X�V = UPDATE
'            ����           ��      txtkenmei
'            ���e           ��      txtNaiyou
'            �J�n��         ��      txtKaisibi
'            ������         ��      txtSyuryobi
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 2001/09/01 �ɓ����q �����ȊO�����p�ł���悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_sNendo             '�N�x
    Public m_sKyokanCd          '��������
    Public m_stxtMode           '���[�h
    Public m_stxtNo             '�����ԍ�
    Public m_sKenmei            '����
    Public m_sNaiyou            '���e
    Public m_sKaisibi           '�J�n��
    Public m_sSyuryoubi         '������
    Public m_rs

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
    w_sMsgTitle="�A���f����"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_stxtMode = request("txtMode")

    m_sKenmei   = request("txtKenmei")
    m_sNaiyou   = request("txtNaiyou")
    m_sKaisibi  = request("txtKaisibi")
    m_sSyuryoubi= request("txtSyuryoubi")
    m_sNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_stxtNo    = request("txtNo")

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

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'//�Ώێҋ敪�������ȊO�̏ꍇ�́A�����敪����I���ł��Ȃ��悤�ɂ���
		Call s_CtrlDisabled()

	    If m_stxtMode = "NEW" Then
	        Call showPage()
	        Exit Do
	    End If

        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
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

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>�A���f����</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
    //************************************************************
    //  [�@�\]  �\���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Hyouji(){
		document.frm.BtnCtrl.value=""

        //���X�g����submit
        document.frm.target = "main" ;
        document.frm.action = "sousin_main.asp";
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

        document.frm.action="./sousin_top.asp";
        document.frm.target="_self";
        document.frm.submit();

    }

	//-->
    </SCRIPT>

<center>

<FORM NAME="frm" action="post" onsubmit="return false">

<br>
<% 
    Select Case m_stxtMode
        Case "NEW"
            call gs_title("�A���f����","�V�@�K")
        Case "UPD"
            call gs_title("�A���f����","�C�@��")
    End Select      
%>
<font>���@�t�@��@�I�@��</font>
<br>
</TD>
</TR>
</TABLE>

<br>

<table class=hyo width=46%>
    <tr>
        <td align="center" width=20%><font color="white">���@��</font></td>
        <td class=detail width=80%><%=m_sKenmei%></td>
    </tr>
</table>
<BR>
<div align="center"><span class=CAUTION>�� �\���{�^�����N���b�N���A���t��ꗗ��\�����܂��B<br>
										�� �R���{���������I�������t��ꗗ�̍i���݂��s�Ȃ������ł��܂��B<br>
										�� �Ώێҋ敪�������̏ꍇ�̂݁A�w�ȁA���Ȍn��敪�A�����敪���I���\�ƂȂ�܂��B
</span></div>
<table border="0" width="78%">
    <tr>
<!--
        <td width="80%">
            <table border="0" class=search cellpadding="1" cellspacing="1" width="100%">
                <tr>
-->
                    <td align="left" class=search>
                        <table border="0" cellpadding="1" cellspacing="1" width="100%">
                            <tr>
                                <td align="left" width="10%" Nowrap>�Ώێҋ敪</td>
                                <td width="20%" Nowrap colspan="1">�F
                                <%
                                call gf_ComboSet("UserKbn",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_USER & " AND M01_NENDO=" & m_sNendo & " AND M01_SYOBUNRUI_CD>0","onchange='javascript:f_ReLoadMyPage()'",True,Request("UserKbn"))
                                %>
                                </td>

                                <td align="left" width=10% Nowrap>�w�@��</td>
                                <td widt=20% Nowrap>�F
                                <%
                                call gf_ComboSet("Gakka",C_CBO_M02_GAKKA,"M02_NENDO='" & m_sNendo & "'",m_sGakkaOption,True,"")
                                %>
                                </td>
                                <td align="left" width=10% Nowrap>���Ȍn��敪</td>
                                <td width=20% Nowrap>�F
                                <%
                                call gf_ComboSet("KkeiKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_KYOKA_KEIRETU &" AND M01_NENDO=" & m_sNendo & " AND M01_SYOBUNRUIMEI IS NOT NULL ",m_sKeiretuOption,True,"")
                                %>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" width=10% Nowrap>�����敪</td>
                                <td width=20% Nowrap>�F
                                <%
                                call gf_ComboSet("KkanKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD='" & C_KYOKAN &"' AND M01_NENDO='" & m_sNendo & "'",m_sKyokaKbnOption,True,"")
                                %>
                                </td>

                                <td align="left" width="10%" Nowrap>����</td>
                                <td width="20%"  colspan="1" Nowrap>�F <input type="text" name="txtSimei" size="25" value="<%=Request("txtSimei")%>"></td>

                                <td align="left" width="10%" Nowrap><br></td>
						        <td width="30%" valign="bottom" align="left" colspan="2">�@
						        <input class=button type="button" value="�@�\�@���@" onClick = "javascript:f_Hyouji()">
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
</table>
    <INPUT TYPE=HIDDEN  NAME=txtNo          value="<%=m_stxtNo%>">
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="<%=m_stxtMode%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_sNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtKenmei      value="<%=m_sKenmei%>">
    <INPUT TYPE=HIDDEN  NAME=txtNaiyou      value="<%=m_sNaiyou%>">
    <INPUT TYPE=HIDDEN  NAME=txtKaisibi     value="<%=m_sKaisibi%>">
    <INPUT TYPE=HIDDEN  NAME=txtSyuryoubi   value="<%=m_sSyuryoubi%>">

    <INPUT TYPE=HIDDEN  NAME=BtnCtrl value="<%=Request("BtnCtrl")%>">

</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>