<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �g�p���ȏ��o�^
' ��۸���ID : web/WEB0320/delete.asp
' �@      �\: �o�^����Ă��鋳�ȏ��̍폜���s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           w_sDelKyokasyoCD    :�I�����ꂽ���ȏ��R�[�h
'           txtPageCD       :�\���ϕ\���Ő��i�������g�Ɉ����n�������j
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��A�E�E�i�w���\��
'           �����ցA�߂�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ��A�E�E�i�w��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/16 �≺�@�K��Y
' ��      �X: 2001/08/01 �O�c�@�q�j
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_sMode
    Public  m_sNendo        ':�N�x
    Public  m_sNo           


    '�y�[�W�֌W
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp                      '// �ꗗ�\���s��

'   call gs_viewForm(request.form)
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

        '// ���Ұ�SET
        Call s_SetParam()

        '// �폜�̎��s
        Call S_delete()

        '// �y�[�W��\��
        Call showPage()

End Sub


'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_sMode      = Request("txtMode")
    m_sNendo     = Request("txtNendo")              ':�N�x
    m_sNo        = Request("txtNo")

End Sub


'********************************************************************************
'*  [�@�\]  �폜���s
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub S_delete()
Dim i
Dim w_iKyokasyoCD
Dim w_iRet              '// �߂�l
Dim w_sSQL              '// SQL��
Dim w_sWHERE            '// WHERE��
Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    w_iKyokasyoCD = ""
    w_sSQL = ""

    w_iKyokasyoCD = Request("deleteNO")

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E�}�X�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

        w_sSQL = w_sSQL & vbCrLf & " delete "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & " T47_KYOKASYO T47 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        '���o�����̍쐬
        w_sSQL = w_sSQL & vbCrLf & " T47.T47_NO in (" & m_sNo & ")"
        w_sSQL = w_sSQL & vbCrLf & " and T47.T47_NENDO = " & m_sNendo
        
'response.write ("<BR>w_sSQL = " & w_sSQL)

        if gf_ExecuteSQL(w_sSQL) <> 0 then
            'ں��޾�Ă̎擾���s
            '//۰��ޯ�
            Call gs_RollbackTrans()
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If


    '���ׂĂ̏������������I��
    Exit do
    Loop

    '//�Я�
    Call gs_CommitTrans()

    
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// �I������
    Call gs_CloseDatabase()

End sub

'********************************************************************************
'*  [�@�\]  HTML��\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
%>
<html>
<link rel=stylesheet href=../common/style.css type=text/css>
    <head>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    function gonext() {
    alert('<%=C_SAKUJYO_OK_MSG%>');
            document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>
<body onLoad="gonext();">

<center>
<form action="default.asp" name="frm" target=<%=C_MAIN_FRAME%> method="post">
<img src="../../image/sp.gif" width="20" height="1">
<input type="hidden" name="txtMode" value="DELETE">
<input type="hidden" name="SKyokanCd1" value="<%=Request("SKyokanCd1")%>">
</form>

</center>

</body>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>