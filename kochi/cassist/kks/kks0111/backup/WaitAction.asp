<%@ Language=VBScript %>
<%
'********************************************************************************
'*	[ｼｽﾃﾑ名]		
'*	[ﾌﾟﾛｸﾞﾗﾑ名]	実行処理 with メッセージ表示
'*===============================================================================
'*	[機能]	処理実行時にメッセージページを表示する
'*	[引数]	txtURL		:次に呼び出すURL
'*			txtMsg		:ページに表示するメッセージ
'*			その他		:各処理により可変
'*	[変数]	なし
'*	[引渡]	引き渡されてきたものをそのまま渡す
'*	[説明]	指定メッセージのページを表示し、指定されたURLを呼び出す
'*			URL呼び出し時には引き渡された値をそのまま渡す
'*
'*	[作成日]	
'*	[修正日]	----/--/--
'********************************************************************************
%>
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	Public msURL

'///////////////////////////メイン処理/////////////////////////////

	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*	[機能]	本ASPのﾒｲﾝﾙｰﾁﾝ
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub Main()

	'//引き渡されたﾌｫｰﾑ内容をそのまま引継ぎ指定されたURLを表示
	'msURL = Request("txtURL") & "?" & Request.Form.Item
	msURL = Request("txtURL")
	
	'// ページを表示
	Call showPage()

End Sub

'********************************************************************************
'*	[機能]	HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
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
	//	[機能]	ページロード時処理
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function window_onload() {

		var szURL;
		szURL = "<%=msURL %>";
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
	<P>　<P>
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