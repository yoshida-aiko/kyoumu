<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<% ErrorCD = Request.QueryString("WStaff") %>
<% Select Case ErrorCD %>
<% Case "1" %>
	<body>
		<h3 align=center>■ 登録エラー ■</h3>
			<hr><br>
		<h4 align=center><font color=red>※データベースに登録することが出来ませんでした。</font></h4>
	<table align=center>
	<tr>
		<td>
			<ul>
			<li>社員CD <%= Session.Contents("ErrorCD") %> はすでに使われています。
			</ul>
		</td>
	</tr>
	</table>
	<table align=center>
		<tr>
			<td>
				社員CDが重複しているので登録することは出来ません。<br>
				違う社員CDを入力し、再度登録して下さい。
			</td>
		</tr>
	</table>
	<p align=center><input type="button" value="戻る" onClick=history.go(-1) id=button2 name=button2>

<% Case "2" %>

		<body>
			<h3 align=center>■ 重複データ ■</h3>
				<hr><br>
			<h4 align=center><font color=red>※データベースに重複データがあります。</font></h4>
		<table align=center>
		<tr>
			<td>
				<ul>
				<li>以前削除されたレコードがデータベースに記憶されています。
				<p align=center>社員CD <%= Session.Contents("社員CD") %> を上書きしてもよろしいですか？</p>
				</ul>
			</td>
		</tr>
		</table>

		<form action="SQLexe.asp" method="Post" id=form1 name=form1>
			<input type="hidden" name="CD" value="<%= Session.Contents("社員CD") %>">
			<input type="hidden" name="NAME" value="<%= Session.Contents("社員名称") %>">
			<input type="hidden" name="BIRTHDAY" value="<%= Session.Contents("生年月日") %>">
			<input type="hidden" name="TELL1" value="<%= Session.Contents("電話番号1") %>">
			<input type="hidden" name="TELL2" value="<%= Session.Contents("電話番号2") %>">
			<input type="hidden" name="POST" value="<%= Session.Contents("郵便") %>">
			<input type="hidden" name="ADDRESS1" value="<%= Session.Contents("住所1") %>">
			<input type="hidden" name="ADDRESS2" value="<%= Session.Contents("住所2") %>">
			<input type="hidden" name="BIKOU" value="<%= Session.Contents("備考") %>">
			<% Session.Contents("FLG") = "UPDATE" %>
			<table align="center" width=30%>
				<tr>
					<td align=center><input type="submit" value="O K" id=submit2 name=submit2></td>
		</FORM>
			<form action="ADDNEW.asp" target="Right" method="Post" id=form2 name=form2>
					<td align=center><input type="submit" value="登録キャンセル" id=submit1 name=submit1></td>
			</form>
				</tr>
		</table>
<% End Select %>
</body>
</html>