<%@ Language=VBScript %>
<!--#include file="../../Common/com_All.asp"-->
<html>
<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//************************************************************
//  [機能]  ページロード時処理
//  [引数]
//  [戻値]
//  [説明]
//************************************************************
function window_onload() {
}

//************************************************************
//  [機能]  戻るボタンが押されたとき
//  [引数]  なし
//  [戻値]  なし
//  [説明]
//
//************************************************************
function f_Back(){
	//キャンセル時、初期画面に戻る
	//上フレーム再表示
	parent.topFrame.location.href="./web0360_top.asp?txtClubCd=<%=Request("txtClubCd")%>"
	//下フレーム再表示
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
<input class="button" type="button" onclick="javascript:f_Back();" value="　戻　る　">
</center>

</form>
</body>
</html>