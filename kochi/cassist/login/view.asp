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
' ��      �X: 2001/08/07 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'           : 2001/08/10 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
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
    Public m_sKenmei    
    Public m_sNaiyou    
    Public m_sKyokanCd  
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

    m_stxtNo        = request("txtNo")
    m_stxtSEIMEI    = request("txtSEIMEI")

	'T46�ɂ́A���[�U�[ID�œo�^����Ă���̂ŁA�����R�[�h�ł͊Y�����Ȃ�
	'���[�U�[ID�Œ��o�A�X�V����悤�ɕύX�@2001/12/11 �ɓ�	
    'm_sKyokanCd     = session("KYOKAN_CD")
    m_sKyokanCd     = session("LOGIN_ID")

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
Dim w_sSQL
Dim w_Srs           '�ڍחp�̃��R�[�h�Z�b�g

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
        '//�ϐ��̒l���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT DISTINCT"
        w_sSQL = w_sSQL & " T46_KENMEI,T46_NAIYO "
        w_sSQL = w_sSQL & "FROM "
        w_sSQL = w_sSQL & " T46_RENRAK "
        w_sSQL = w_sSQL & "WHERE "
        w_sSQL = w_sSQL & " T46_NO = '" & cInt(m_stxtNo) & "'"

        Set w_Srs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_Srs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

    '//�擾�����l��ϐ��ɑ��
    m_sKenmei   = w_Srs("T46_KENMEI")
    m_sNaiyou   = w_Srs("T46_NAIYO")

    '//�m�F�t���O���P�ɂ���B
        w_sSQL = ""
        w_sSQl = w_sSQL & " UPDATE T46_RENRAK SET "
        w_sSQL = w_sSQL & "     T46_KAKNIN = 1 "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "     T46_NO = " & cInt(m_stxtNo) & ""
        w_sSQL = w_sSQL & " AND T46_KYOKAN_CD = '" & m_sKyokanCd & "'"

        iRet = gf_ExecuteSQL(w_sSQL)
        If iRet <> 0 Then
            msMsg = Err.description
            f_GetData = 99
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
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link rel="stylesheet" href="../common/style.css" type="text/css">
    <title>���m�点</title>

    <!--#include file="../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [�@�\]  �I�����[�h��
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_open(){

        //���X�g����submit
        //document.frm.target = "<%=C_MAIN_FRAME%>_low" ;
        //document.frm.action = "top_lwr.asp";
        //document.frm.submit();

	opener.location.reload();
        window.focus();

    }

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
<BODY onload="f_open()">
<center>
<FORM NAME="frm" action="post">
<br>
<% 
call gs_title("���@�m�@��@��","�ځ@��")
%>
<br>
<table width="300" border="1" CLASS="hyo">
    <TR>
        <TH CLASS="header" width="60">����</TH>
        <TD CLASS="detail"><%=m_sKenmei%></TH>
    </TR>
    <TR>
        <TH CLASS="header">���e</TD>
        <TD CLASS="detail"><%=m_sNaiyou%></TD>
    </TR>
    <TR>
        <TH CLASS="header">���M��</TD>
        <TD CLASS="detail"><%=m_stxtSEIMEI%></TD>
    </TR>
</table>
<br>
<table border="0" width="350">
    <tr>
    <td valign="top" align="center">
    <input type="button" value="����" class="button" onclick="javascript:f_Close()">
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