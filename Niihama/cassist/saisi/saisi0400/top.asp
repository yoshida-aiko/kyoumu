<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �ǎ���u�҈ꗗ
' ��۸���ID : saisi/saisi0400/top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_sNendo             '�N�x
    Public m_PgMode             '�����ʃt���O
    Public m_sMsgTitle          '����

    Public m_Rs

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
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    On Error Resume Next

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�Ď���u�҈ꗗ"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    'm_PgMode=request("p_mode")
	'Select Case m_PgMode
	'	Case "P_HAN0100"
	'	    m_sMsgTitle="���шꗗ�\"
	'	Case "P_KKS0200"
	'	    m_sMsgTitle="���ۈꗗ�\"
	'	Case "P_KKS0210"
	'	    m_sMsgTitle="�x���ꗗ�\"
	'	Case "P_KKS0220"
	'	    m_sMsgTitle="�s�����ۈꗗ�\"
	'	Case Else
	'End Select
	'w_sMsgTitle = m_sMsgTitle

    Err.Clear

    m_bErrFlg = False

    m_sNendo    = session("NENDO")
    m_iDsp = C_PAGE_LINE

    Do

		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop

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
    <title>�Ď���u�҈ꗗ</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Touroku(){

        //���X�g����submit
        //document.frm.target="_parent";
        //document.frm.action="regist.asp";
        //document.frm.txtMode.value ="NEW";
        //document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript>
    <form name="frm" method="post">
    <center>
<%call gs_title("�Ď���u�҈ꗗ","��@��")%>
<br>
    <!--INPUT TYPE=HIDDEN NAME=txtMode     VALUE=""-->
    <INPUT TYPE=HIDDEN NAME=txtNendo    VALUE="<%=m_sNendo%>">
    <!--INPUT TYPE=HIDDEN NAME=txtKyokanCd VALUE="<%=m_sKyokanCd%>"-->

    </center>

    </form>
    </body>
    </html>
<%
    '---------- HTML END   ----------
End Sub
%>
