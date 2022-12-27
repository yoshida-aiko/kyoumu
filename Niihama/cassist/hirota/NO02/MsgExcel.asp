<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<body>
	<h3 align=center>■ 出力エラー ■</h3>
		<hr><br>
	<h4 align=center><font color=red>Excelの起動に失敗しました。<br>
	※データベースのデータを出力することが出来ませんでした。</font></h4>
	<p align=center>
		エラー：<%= Session.Contents("g_Err") %>
	</p>
	<p align=center>
		<form action="EXCEL.asp" target="Right">
			<input type="submit" value="戻 る">
		</form>
	</p>
</body>
</html>
		document.write"<html>"
		document.write"<head>"
			document.write"<title>社員管理</title>"
			document.write"<base target=Right>"
		document.write"</head>"
		<!-- <BODY BGCOLOR=#F5F5F5> -->
		document.write"<body>"
			document.write"<h3 align=center>★ EXCEL出力 ★</h3>"
				document.write"<hr>"
			document.write"<h2 align=center><font color=red>EXCEL出力が完了しました！</font></h2>"

		document.write"<table align=center>"
			document.write"<tr>"
				document.write"<td>"
					document.write"出力件数"
				document.write"</td>"
				document.write"<td>"
					document.write"："
				document.write"</td>"
				document.write"<td>"
				document.write"</td>"
			document.write"</tr>"
		document.write"</table>"
		document.write"</td>"
		document.write"</tr>"
		document.write"</table>"
		document.write"<br>"
		document.write"<table align=center width=20%>"
			document.write"<tr>"
				document.write"<td align=center>"
				document.write"<form action=EXCEL.asp id=form1 name=form1>"
					document.write"<p align=center><input type=submit value=戻る id=submit1 name=submit1>"
				document.write"</td>"
				document.write"</form>"
				document.write"<form action=INitiran.asp target=Right id=form2 name=form2>"
				document.write"<td align=center valign=bottom>"
					document.write"<input type=submit value=一覧 id=submit2 name=submit2>"
				document.write"</td>"
				document.write"</form>"
			document.write"<tr>"
		document.write"</table>"
		document.write"</body>"
		document.write"</html>"