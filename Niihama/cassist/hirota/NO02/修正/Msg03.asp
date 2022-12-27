<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>■ 重複データ ■</h3>
		<hr><br>
	<h4 align=center><font color=red>※データベースデータ重複メッセージ</font></h4>
<table align=center>
<tr>
	<td>
		<ul>
		<li>以前削除されたレコードがデータベースに記憶されています。
		<p align=center>社員CD <%= Request.QueryString("CD") %> を上書きしてもよろしいですか？</p>
		</ul>
	</td>
</tr>
</table>

<form action="SQLexe.asp" method="Post">
	<input type="hidden" name="社員CD" value="<%= Request.QueryString("CD") %>">
	<input type="hidden" name="社員名称" value="<%= Session.Contents("社員名称") %>">
	<input type="hidden" name="生年月日" value="<%= Session.Contents("生年月日") %>">
	<input type="hidden" name="電話番号1" value="<%= Session.Contents("電話番号1") %>">
	<input type="hidden" name="電話番号2" value="<%= Session.Contents("電話番号2") %>">
	<input type="hidden" name="郵便" value="<%= Session.Contents("郵便") %>">
	<input type="hidden" name="住所1" value="<%= Session.Contents("住所1") %>">
	<input type="hidden" name="住所2" value="<%= Session.Contents("住所2") %>">
	<input type="hidden" name="備考" value="<%= Session.Contents("備考") %>">
	<input type="hidden" name="FLG" value="2">
	<table align="center" width=30%>
		<tr>
			<td align=center><input type="submit" value="O K"></td>
</FORM>
	<form action="SHINKI.asp" target="Right" method="Post">
			<td align=center><input type="submit" value="登録キャンセル" id=submit1 name=submit1></td>
	</form>
		</tr>
</table>
</body>
</html>
