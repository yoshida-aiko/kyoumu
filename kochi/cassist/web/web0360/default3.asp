<%@ Language=VBScript %>
<!--#include file="../../Common/com_All.asp"-->
<html>
<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//************************************************************
//  [�@�\]  �y�[�W���[�h������
//  [����]
//  [�ߒl]
//  [����]
//************************************************************
function window_onload() {
}

//************************************************************
//  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
//  [����]  �Ȃ�
//  [�ߒl]  �Ȃ�
//  [����]
//
//************************************************************
function f_Back(){
	//�L�����Z�����A������ʂɖ߂�
	//��t���[���ĕ\��
	parent.topFrame.location.href="./web0360_top.asp?txtClubCd=<%=Request("txtClubCd")%>"
	//���t���[���ĕ\��
	parent.main.location.href="./web0360_main.asp?txtClubCd=<%=Request("txtClubCd")%>"

}

//-->
</SCRIPT>

</head>
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">

<center>
<br><br>
<span class="msg"><%=C_BRANK_VIEW_MSG%></span>
<br><br>
<input class="button" type="button" onclick="javascript:f_Back();" value="�@�߁@��@">
</center>

</form>
</body>
</html>