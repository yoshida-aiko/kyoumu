<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���шꗗ
' ��۸���ID : sei/sei0200/sei0200_main.asp
' �@      �\: �t���[���y�[�W ���шꗗ�̓o�^���s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/10/22 �J�e�@�ǖ�
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
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

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���шꗗ"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_parent"


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
		session("PRJ_No") = "SEI0200"

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

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

dim w_sItem

	w_sItem = ""
	w_sItem = w_sItem & "?txtSikenKBN=" & request("txtSikenKBN")
	w_sItem = w_sItem & "&txtHyojiKBN=" & request("txtHyojiKBN")
	w_sItem = w_sItem & "&txtGakuNo=" & request("txtGakuNo")
	w_sItem = w_sItem & "&txtClassNo=" & request("txtClassNo")
	w_sItem = w_sItem & "&txtKBN=" & request("txtKBN")
	w_sItem = w_sItem & "&txtNendo=" & request("txtNendo")
	w_sItem = w_sItem & "&txtKyokanCd=" & request("txtKyokanCd")
	w_sItem = w_sItem & "&txtKengen=" & request("txtKengen")

if request("txtGakkaNo") = C_SEI0200_ACCESS_TANNIN then
	w_sItem = w_sItem & "&txtClassNo=" & request("txtClassNo")
ElseIf request("txtGakkaNo") = C_SEI0200_ACCESS_GAKKA then
	w_sItem = w_sItem & "&txtGakkaNo=" & request("txtGakkaNo")
End if
%>

<html>

<head>
<title>���шꗗ</title>
</head>

<frameset rows=310,1,* frameborder="0" framespacing="0">
	<frame src="sei0200_middle.asp<%=w_sItem%>" scrolling="no"  name="stop">
    <frame src="../../common/bar.html" scrolling="auto" name="bar" noresize>
	<frame src="sei0200_bottom.asp<%=w_sItem%>" scrolling="auto"  name="smain" >
</frameset>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
