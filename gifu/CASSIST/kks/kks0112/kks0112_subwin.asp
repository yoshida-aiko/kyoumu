<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo���ꗗ
' ��۸���ID : kks/kks0110/kks0111_detail.asp
' �@      �\: �t���[���y�[�W ���Əo���\��
'-------------------------------------------------------------------------
' ��      ��:
' ��      ��:
' ��      �n:
' ��      ��:
'           
'-------------------------------------------------------------------------
' ��      ��: 2002/05/07 shin
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

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()
	Dim w_iRet              '// �߂�l
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���Əo������"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"
	
    On Error Resume Next
    Err.Clear
	
    m_bErrFlg = False
	
    Do
		'// �ް��ް��ڑ�
		w_iRet = gf_OpenDatabase()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			w_sMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			'm_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If
		
		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = "KKS0112"
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'// �y�[�W��\��
		Call showPage()
		Exit Do
    Loop
	
	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
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
<title>���Əo������</title>

<script language="javascript1.2">
<!--

window.onload=init;

function init(){
	if(!document.all){
		document.all={}
		if(!frames["topFrame"].document.body)frames["topFrame"].document.body={scrollLeft:0,scrollTop:0}
			frames["topFrame"].setInterval('if(parent.frames["main"].pageXOffset!=(document.body.scrollLeft=self.pageXOffset))document.body.onscroll()',10)
		if(!frames["main"].document.body)frames["main"].document.body={scrollLeft:0,scrollTop:0}
			frames["main"].setInterval('if(parent.frames["topFrame"].pageXOffset!=(document.body.scrollLeft=self.pageXOffset))document.body.onscroll()',10)
	}
	
	if(document.all){
		frames["topFrame"].document.body.onscroll=function(){
			frames["main"].scrollTo(frames["topFrame"].document.body.scrollLeft,frames["main"].document.body.scrollTop)
		}
		
		frames["main"].document.body.onscroll=function(){
			frames["topFrame"].scrollTo(frames["main"].document.body.scrollLeft,frames["topFrame"].document.body.scrollTop)
		}
	}
}

//-->
</script>

</head>
<frameset cols="200px,1,*" border="1" frameborder="no" onBlur="window.focus();">
	<frame src="kks0112_subwin_left.asp?<%=Request.QueryString%>" scrolling="no" name="leftFrame" border="1">
	<frame src="../../common/bar.html" scrolling="no" name="barH" noresize>
	
    <frameset rows="80px,1,*" border="1" frameborder="no" border="1" onload="init();">
		<frame src="kks0112_subwin_top.asp?<%=Request.QueryString%>" scrolling="yes" name="topFrame">
		<frame src="../../common/bar.html" scrolling="no" name="barW" noresize>
		<frame src="kks0112_subwin_bottom.asp?<%=Request.QueryString%>" scrolling="yes" name="main">
	</frameset>
    
</frameset>

</head>
</html>
<%
End Sub
%>