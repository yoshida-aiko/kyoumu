<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Ԋ������A��
' ��۸���ID : web/web0310/web0310_DEL.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/24 �O�c
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp          '// �ꗗ�\���s��
    Public  m_rs
    Public  m_stxtMode      '���[�h
    Dim     m_iNendo
    Dim     m_sKyokanCd
    Dim     m_sNo
    Dim     m_sDelNo
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
    w_sMsgTitle="���Ԋ������A��"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sNo   = request("Delchk")
    m_sDelNo    = request("txtDelNo")
    m_stxtMode = request("txtMode")
    m_iDsp = C_PAGE_LINE

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

        Select Case m_stxtMode

            Case "","DELKNIN"
                '//���X�g�̈ꗗ�f�[�^�̏ڍ׎擾
                w_iRet = f_GetData()
                If w_iRet <> 0 Then
                    '�ް��ް��Ƃ̐ڑ��Ɏ��s
                    m_bErrFlg = True
                    Exit Do
                End If
                '// �y�[�W��\��
                Call showPage()
                Exit Do

            Case "Delete"

                w_iRet = f_DeleteData()
                If w_iRet <> 0 Then
                    '�ް��ް��Ƃ̐ڑ��Ɏ��s
                    m_bErrFlg = True
                    Exit Do
                End If
                '// �y�[�W��\��
                Call DEL_showPage()
                Exit Do
        End Select

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

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do

        '//���X�g�̕\��
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT DISTINCT "
        m_sSQL = m_sSQL & "     T52_NO,T52_NAIYO "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T52_JYUGYO_HENKO "
        m_sSQL = m_sSQL & " WHERE "
        If m_stxtMode = "" Then
            m_sSQL = m_sSQL & "     T52_NO IN (" & Trim(m_sNo) & ") "
        ElseIf m_stxtMode = "DELKNIN" Then
            m_sSQL = m_sSQL & "     T52_NO = '" & Trim(m_sDelNo) & "' "
        End If

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
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

Function f_DeleteData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_DeleteData = 1

    Do
        '//���X�g�̕\��
        m_sSQL = ""
        m_sSQL = m_sSQL & " DELETE FROM T52_JYUGYO_HENKO "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T52_NO IN (" & Trim(m_sDelNo) & ") "

        iRet = gf_ExecuteSQL(m_sSQL)
        If iRet <> 0 Then
            msMsg = Err.description
            f_DeleteData = 99
            Exit Do
        End If

        f_DeleteData = 0

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

    On Error Resume Next
    Err.Clear
%>

<html>
    <head>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �폜�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_delete(){

        if (!confirm("<%=C_SAKUJYO_KAKUNIN%>")) {
           return ;
        }
        document.frm.action="web0310_DEL.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "Delete";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
</head>
<body>

<center>

<%call gs_title("���Ԋ������A��","��@��")%>
<br>
��@���@���@�e
<br><br>
    <table border="1" class=hyo width="80%">
<form name="frm" action="post">

    <tr>
    <th class=header>�����ԍ�</th>
    <th class=header>���@�e</th>
    </tr>

<%
    m_rs.MoveFirst
    Do Until m_rs.EOF
%>
    <tr>
    <td align="center" class=detail width="20%"><%=m_rs("T52_NO")%></td>
    <td class=detail width="80%"><%=m_rs("T52_NAIYO")%></td>
    </tr>
<%
    m_rs.MoveNext
    Loop
 %>

    </table>
<br>
�ȏ�̓��e���폜���܂��B
<br><br>
<table border="0">
<tr>
<td align=left>
<input type="button" class=button value="�@��@���@" Onclick="javascript:f_delete()">
</td>
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_iNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">

<%
    If m_stxtMode = "" Then
%>
        <INPUT TYPE=HIDDEN  NAME=txtDelNo       value="<%=m_sNo%>">
<%
    ElseIf m_stxtMode = "DELKNIN" Then
%>
        <INPUT TYPE=HIDDEN  NAME=txtDelNo       value="<%=m_sDelNo%>">
<%
        End If
%>

</form>
<form action="default.asp" target="<%=C_MAIN_FRAME%>" method="post">
<td align=right>
<input type="submit" class=button value="�L�����Z��">
</td>
</form>
</tr>
</table>

</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub

Sub DEL_showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>

    <html>
    <head>
    <title>���Ԋ������A��</title>
    <link rel=stylesheet href=../../font.css type=text/css>

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

        location.href = "default.asp"
        return;
    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>