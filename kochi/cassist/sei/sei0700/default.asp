<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�
' ��۸���ID : gak/gak0300/default.asp
' �@      �\: �t���[���y�[�W �w����񌟍����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg          '�װ�׸�
    Public  m_PgMode           '�������[�h
    Public  m_sMsgTitle        '����

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
    w_sMsgTitle="���шꗗ�\"
    w_sRetURL="../../login/default.asp"
    w_sTarget="_parent"

    m_PgMode=request("p_mode")
	Select Case m_PgMode
		Case "P_HAN0100"
		    m_sMsgTitle="���шꗗ�\"
		Case "P_KKS0200"
		    m_sMsgTitle="���ۈꗗ�\"
		Case "P_KKS0210"
		    m_sMsgTitle="�x���ꗗ�\"
		Case "P_KKS0220"
		    m_sMsgTitle="�s�����ۈꗗ�\"
		Case "P_HAN0111"
		    m_sMsgTitle="�]�_�ꗗ�\"
		Case Else
	End Select
	w_sMsgTitle = m_sMsgTitle

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do

        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �����`�F�b�N�Ɏg�p
		'session("PRJ_No") = "GAK0300"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

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


'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
    '---------- HTML START  ----------
    On Error Resume Next
    Err.Clear

%>

<html>
	<head>
	<title><%=m_sMsgTitle%></title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<script language=javascript>
	</script>
	<frameset rows="65,*" border="1" framespacing="0" frameborder="no"> 
		<frame src="top.asp?p_mode=<%=m_PgMode%>" name="fTop" marginwidth="0" noresize scrolling="no" frameborder="no">
		<frame src="main.asp?p_mode=<%=m_PgMode%>&txtMode=Search" name="fMain" marginwidth="0" marginheight="0" scrolling="auto" frameborder="0">
	</frameset>
	</head>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>

