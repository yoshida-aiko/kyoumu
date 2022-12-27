<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>★ EXCEL出力 ★</h3>
<hr>
	<h2 align=center><font color=red>EXCEL出力が完了しました！</font></h2>
<table align=center>
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
</table>
<br>
<table align=center width=20%>
	<tr>
		<form action="EXCEL.asp" id=form1 name=form1>
			<td align=center><p align=center><input type="submit" value="戻 る" id=submit1 name=submit1>	</td>
		</form>
		<form action="INitiran.asp" target="Right" id=form2 name=form2>
			<td align=center valign=bottom><input type="submit" value="一 覧" id=submit2 name=submit2></td>
		</form>
	<tr>
</table>
</body>
</html>
