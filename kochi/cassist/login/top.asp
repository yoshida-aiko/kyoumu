<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �i�g�b�v�y�[�W�j
' ��۸���ID : login/top.asp
' �@      �\: �t���[���y�[�W ���T�̎��Ԋ��Ƃ��m�点�̎Q�Ƃ��s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/07/19 ���{ ����
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_bErrFlg           '�װ�׸�

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////


'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sRetURL           '// �װү���ޗp�߂��URL
    Dim w_sTarget           '// �װү���ޗp�߂���ڰ�
    Dim w_sWinTitle         '// �װү���ޗp����
    Dim w_sMsgTitle         '// �װү���ޗp����
    
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�g�b�v�y�[�W"
    w_sMsg=""
    w_sRetURL="../default.asp"
    w_sTarget="_parent"

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
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		Call showPage()         '// �y�[�W��\��

        '// ����I��
        Exit Do
    LOOP

   '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, m_sErrMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()

End Sub


'********************************************************************************
'*  [�@�\]  HTML�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()

'	if gf_IsNull(trim(Session("KYOKAN_CD"))) then
		w_FremeSize = "0,*"
		w_FrameUrl  = "about:blank;"
'	Else
'		w_FremeSize = "300,*"
'		w_FrameUrl  = "jikanwari.asp"
'	End if

%>

<html>

<head>
<title>���������V�X�e���FCampus Assist �g�b�v�y�[�W</title>
</head>

<frameset rows="<%=w_FremeSize%>" frameborder="0">
	<frame src="<%=w_FrameUrl%>" scrolling="auto" noresize name="<%=C_MAIN_FRAME%>_up">
	<frame src="top_lwr.asp"   scrolling="auto" noresize name="<%=C_MAIN_FRAME%>_low">
</frameset>

</html>

<% End Sub %>