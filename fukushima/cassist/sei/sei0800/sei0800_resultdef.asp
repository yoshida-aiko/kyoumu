<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���юQ�Ɓi�������j
' ��۸���ID : sei/sei0800/default.asp
' �@      �\: 
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2003/05/13 �A�c
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�
	Dim     m_sGakuseiNo
	Dim     m_sGakunen
	Dim     m_sClass
	Dim     m_sGakuseiNM

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
	w_sMsgTitle="���юQ��"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_parent"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do
		'// �ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'// �����`�F�b�N�Ɏg�p
'		Session("PRJ_No") = "SEI0800"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		Call s_SetParam()

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
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()

	m_sGakuseiNo = Request("hidGakuseiNo")
	m_sGakunen   = Request("hidGakunen")
	m_sClass     = Request("hidClass")
	m_sGakuseiNM = Request("hidGakuseiNM")

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
<!--#include file="../../Common/scroll.js"-->
<title>���юQ��</title>
</head>

<frameset rows=137px,1,* frameborder="0" framespacing="0">
	<frame src="sei0800_resulttop.asp?hidGakunen=<%=m_sGakunen%>&hidClass=<%=m_sClass%>&hidGakuseiNM=<%=m_sGakuseiNM%>&hidGakuseiNo=<%=m_sGakuseiNo%>"    scrolling="auto" name="topFrame" noresize>
    <frame src="../../common/bar.html"    scrolling="auto" name="bar"      noresize>
	<frame src="sei0800_resultbottom.asp?hidGakuseiNo=<%=m_sGakuseiNo%>" scrolling="auto"  name="main"     noresize>
</frameset>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
