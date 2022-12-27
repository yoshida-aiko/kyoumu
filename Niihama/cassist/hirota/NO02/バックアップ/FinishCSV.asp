<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>★ CSV出力 ★</h3>
		<hr>
	<h2 align=center><font color=red>CSV出力が完了しました！</font></h2>
<table align=center>
	<tr>
		<td>
			<table align=center>
				<tr>
					<td>
						出力場所
					</td>
					<td>
						：
					</td>
					<td>
						<%= Session.Contents("Path") %>
					</td>
				</tr>
				<tr>
					<td>
						出力件数
					</td>
					<td>
						：
					</td>
					<td>
						<%= Request.QueryString("Count") %> 件
					</td>
				</tr>
				<tr>
					<td>
						出力CSVファイル
					</td>
					<td>
						：
					</td>
					<td>
						<a href="\\WEBSVR_2\infogram\hirota\No02\Sample.csv">ダウンロード</a>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<table align=center width=20%>
	<tr>
		<form action="CSV.asp" id=form1 name=form1>
			<td align=center><p align=center><input type="submit" value="戻 る"></td>
		</form>
		<form action="INitiran.asp" target="Right">
			<td align=center valign=bottom><input type="submit" value="一 覧"></td>
		</form>
	<tr>
</table>
</body>
</html>
