<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �}�C�y�[�W�̂��m�点
' ��۸���ID : login/view.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/23 �O�c
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iMax           ':�ő�y�[�W
    Public m_iDsp                       '// �ꗗ�\���s��
    Public m_stxtNo         '�����ԍ�
    Public m_stxtSEIMEI     '���M�҂̐���
    Public m_sKyokanCd  
    Public m_iNendo 
    Public m_rs

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
    w_sMsgTitle="�A�������o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sKyokanCd     = session("KYOKAN_CD")
    m_iNendo        = session("NENDO")
    m_iDsp          = C_PAGE_LINE

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '//�f�[�^�̎擾�A�\��
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If
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

Function f_GetData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
	Dim w_user

    On Error Resume Next
    Err.Clear
    f_GetData = 1

	w_user = session("LOGIN_ID")
	'���[�U�������ł���΁A����CD����
	If m_sKyokanCd <> "" then w_user = m_sKyokanCd

    Do
        '//�ϐ��̒l���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     A.T52_NAIYO,A.T52_INS_DATE,B.M10_USER_NAME "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T52_JYUGYO_HENKO A,M10_USER B "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "     A.T52_KYOKAN_CD = '" & w_user & "' AND "
        w_sSQL = w_sSQL & "     A.T52_INS_USER = B.M10_USER_ID AND "
        w_sSQL = w_sSQL & "     B.M10_NENDO = " & m_iNendo & " "

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do
        End If

    f_GetData = 0

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_sClass

%>
<HTML>
<head>
<link rel=stylesheet href="../common/style.css" type=text/css>
    <title>���Ԋ��ύX�A��</title>

    <!--#include file="../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  ����{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Close(){

        window.close()

    }
    //-->
    </SCRIPT>
</head>

<BODY>
<center>
<FORM NAME="frm" action="post">

<% call gs_title("���Ԋ��ύX�A��","�ځ@��") %>
<br>
<table width="400" border=1 CLASS="hyo">
    <TR>
        <TH CLASS="header" width=65%>���@�M�@��</TD>
        <TH CLASS="header" width=35%>�o�@�^�@��</TD>
    </TR>
    <TR>
        <TH CLASS="header" width=100% colspan=2>�@���@�e�@</TD>
    </TR>
	<%
	    m_rs.MoveFirst
	    Do Until m_rs.EOF
			%>
		    <TR>
		        <TD CLASS="CELL1" ><%=m_rs("M10_USER_NAME")%></TD>
		        <TD CLASS="CELL1" ><%=m_rs("T52_INS_DATE")%></TD>
		    </TR>
		    <TR>
		        <TD CLASS="CELL2" colspan=2><%=m_rs("T52_NAIYO")%></TD>
		    </TR>
			<%
	    m_rs.MoveNext
    Loop
	%>
    </TABLE>

	<br>
    <table border="0" width="350">
	    <tr>
		    <td valign="top" align="center">
		    <input type="button" value="����" class=button onclick="javascript:f_Close()">
		    </td>
	    </tr>
    </table>

</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>