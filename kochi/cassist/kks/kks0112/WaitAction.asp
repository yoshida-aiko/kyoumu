<%@ Language=VBScript %>
<%
'********************************************************************************
'*	[���і�]		
'*	[��۸��і�]	���s���� with ���b�Z�[�W�\��
'*===============================================================================
'*	[�@�\]	�������s���Ƀ��b�Z�[�W�y�[�W��\������
'*	[����]	txtURL		:���ɌĂяo��URL
'*			txtMsg		:�y�[�W�ɕ\�����郁�b�Z�[�W
'*			���̑�		:�e�����ɂ���
'*	[�ϐ�]	�Ȃ�
'*	[���n]	�����n����Ă������̂����̂܂ܓn��
'*	[����]	�w�胁�b�Z�[�W�̃y�[�W��\�����A�w�肳�ꂽURL���Ăяo��
'*			URL�Ăяo�����ɂ͈����n���ꂽ�l�����̂܂ܓn��
'*
'*	[�쐬��]	
'*	[�C����]	----/--/--
'********************************************************************************
%>
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	Public msURL

'///////////////////////////���C������/////////////////////////////

	'Ҳ�ٰ�ݎ��s
	Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*	[�@�\]	�{ASP��Ҳ�ٰ��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub Main()

	'//�����n���ꂽ̫�ѓ��e�����̂܂܈��p���w�肳�ꂽURL��\��
	'msURL = Request("txtURL") & "?" & Request.Form.Item
	msURL = Request("txtURL")
	
	'// �y�[�W��\��
	Call showPage()

End Sub

'********************************************************************************
'*	[�@�\]	HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub showPage()
	'---------- HTML START ----------
	%>
	<HTML>
	<HEAD>
	<META>
	<TITLE></TITLE>
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//	[�@�\]	�y�[�W���[�h������
	//	[����]
	//	[�ߒl]
	//	[����]
	//************************************************************
	function window_onload() {

		var szURL;
		szURL = "<%=msURL %>";
		//document.location.href = szURL;
        document.frm.target = "main";
        document.frm.action = szURL
        document.frm.submit();
        return;

	}

	//-->
	</SCRIPT>
	</HEAD>
	<BODY LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
	<br><br><br>
	<CENTER><span class="msg"><%=Request("txtMsg") %></span></CENTER>
	<P>�@<P>
	<%
	For Each I_Name In Request.Form
	Response.Write "<input type='hidden' name='" & I_Name & "' value='" & Request.Form(I_Name) & "'>" & vbCrLf
	Next
	%>
	</form>
	</BODY>
	</HTML>
	<%
	'---------- HTML END   ----------
End Sub
%>