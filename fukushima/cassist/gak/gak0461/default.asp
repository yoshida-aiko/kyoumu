<%@Language=VBScript %>
<%
'******************************************************************
'�V�X�e����     �F���������V�X�e��
'���@���@��     �F�������������o�^
'�v���O����ID   �Fgak/gak0461/default.asp
'�@�@�@�@�\     �F�t���[���y�[�W �w�Јψ������͂̕\�����s��
'------------------------------------------------------------------
'���@�@�@��     �F
'�ρ@�@�@��     �F
'���@�@�@�n     �F
'���@�@�@��     �F
'------------------------------------------------------------------
'��@�@�@��     �F2001.07.18    �O�c�@�q�j
'�ρ@�@�@�X     �F2001/08/30 �ɓ� ���q     ����������2�d�ɕ\�����Ȃ��悤�ɕύX
'******************************************************************
Public m_sMode
Public m_iNendo
Public m_sNendo
Public m_sKyokanCd
Public m_sGakuNo
Public m_sGakunen
Public m_sClass
Public m_sClassNm
'*******************�@ASP���ʃ��W���[���錾�@**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******�@�� �W �� �[ �� �� ���@********
'******�@���C�������@********

    'Ҳ�ٰ�ݎ��s
    Call Main()

'******�@�d�@�m�@�c�@********

Sub Main()
'******************************************************************
'�@�@�@�\�F�{ASP��Ҳ�ٰ��
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    '******���ʊ֐�******
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E�}�X�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_parent"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sMode = request("txtMode")
    m_iNendo = session("NENDO")
    m_sNendo = request("txtNendo")
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakuNo = request("GakuseiNo")
    m_sGakunen = request("txtGakunen")
    m_sClass = request("txtClass")
    m_sClassNm = request("txtClassNm")

'response.write m_sMode &"<<br>"
'response.write m_iNendo &"<<br>"
'response.write m_sNendo &"<<br>"
'response.write m_sKyokanCd &"<<br>"
'response.write m_sGakuNo &"<<br>"
'response.write m_sGakunen  &"<<br>"
'response.write m_sClass &"<<br>"
'response.end

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
		session("PRJ_No") = "GAK0461"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '// �S�C�`�F�b�N
	  If gf_Tannin(m_iNendo,m_sKyokanCd,5) <> 0 Then
	            m_bErrFlg = True
	            m_sErrMsg = "�S�C�ȊO�̓��͂͂ł��܂���B"
	            Exit Do
	  End If

'--------2001/08/30 ito --------------
'		If m_sGakuNo <> "" Then
'			Call showPageBack()
'	        Exit Do
'		End If

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
    <title>�������������o�^</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    </script>
    <frameset rows=160,1,* frameborder="0">
        <frame src="gak0461_top.asp?txtGakuNo=<%=Request("txtGakuNo")%>&txtNendo=<%=m_sNendo%>" scrolling="auto" noresize name="topFrame">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        <frame src="default2.asp" scrolling="auto" noresize name="main">
    </frameset>
    </head>
</html>
<%
End Sub

Sub showPageBack()
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
    <title>�������������o�^</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    </script>
    <frameset rows=180,1,* frameborder="0" FRAMESPACING="0" border="0">
        <frame src="gak0461_top.asp?txtGakuNo=<%=m_sGakuNo%>&txtNendo=<%=m_sNendo%>" scrolling="auto" noresize name="topFrame">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        <frame src="gak0461_main.asp?txtGakuNo=<%=m_sGakuNo%>&txtGakunen=<%=m_sGakunen%>&txtClass=<%=m_sClass%>&txtNendo=<%=m_sNendo%>&txtClassNm=<%=m_sClassNm%>" scrolling="auto" noresize name="main">
    </frameset>
    </head>
</html>
<%
End Sub
%>