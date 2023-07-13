<%@ Language=VBScript %>
<%call main

Sub main()
%>

	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>

	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//	[機能]	ページロード時処理
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function window_onload() {
		document.body.style.cursor = "wait";
	}

	//-->
	</SCRIPT>

	</head>

	<body onload="window_onload();">
	<center>
	<br><br><br>
	<span class="msg">しばらくお待ちください。</span>
	</center>
	</body>

	</html>
<%
End Sub 
%>