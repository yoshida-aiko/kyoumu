<html>
<link rel=stylesheet href="../../common/style.css" type=text/css>

	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	//************************************************************
	//	[機能]	一覧へボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_itiran(){

		//リスト情報をsubmit
		document.frm.target = "fTopMain" ;
		document.frm.action = "default.asp";
		document.frm.txtMode.value = "";
		document.frm.submit();

	}

	//************************************************************
	//	[機能]	前画面へボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Back(){

		//リスト情報をsubmit
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
			<input type="button" value="前 画 面 へ" class=button onclick="javascript:f_Back()">
			<input type="button" value="キャンセル" class=button onclick="javascript:f_itiran()">
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