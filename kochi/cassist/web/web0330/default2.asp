<html>
<link rel=stylesheet href="../../common/style.css" type=text/css>

	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	//************************************************************
	//	[�@�\]	�ꗗ�փ{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_itiran(){

		//���X�g����submit
		document.frm.target = "fTopMain" ;
		document.frm.action = "default.asp";
		document.frm.txtMode.value = "";
		document.frm.submit();

	}

	//************************************************************
	//	[�@�\]	�O��ʂփ{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Back(){

		//���X�g����submit
		document.frm.target = "fTopMain" ;
		document.frm.action = "regist.asp";
		document.frm.submit();

	}

	//-->
	</SCRIPT>

<body>
<form name="frm" action="regist.asp" target="fTopMain" Method="POST">
<table border="0" width="100%">
	<tr>
		<td align="center">
			<input type="button" value="�O �� �� ��" class=button onclick="javascript:f_Back()">
			<input type="button" value="�L�����Z��" class=button onclick="javascript:f_itiran()">
		</td>
	</tr>
</table>
    <INPUT TYPE=HIDDEN  NAME=txtNo          value="<%=request("txtNo")%>">
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="<%=request("txtMode")%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=request("txtNendo")%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=request("txtKyokanCd")%>">
    <INPUT TYPE=HIDDEN  NAME=txtKenmei      value="<%=request("txtKenmei")%>">
    <INPUT TYPE=HIDDEN  NAME=txtNaiyou      value="<%=request("txtNaiyou")%>">
    <INPUT TYPE=HIDDEN  NAME=txtKaisibi     value="<%=request("txtKaisibi")%>">
    <INPUT TYPE=HIDDEN  NAME=txtSyuryoubi   value="<%=request("txtSyuryoubi")%>">
</form>
</body>

</html>