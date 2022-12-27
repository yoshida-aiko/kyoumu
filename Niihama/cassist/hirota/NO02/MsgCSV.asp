<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>■ 入力項目エラー ■</h3>
		<hr><br>
	<h4 align=center><font color=red>※データベースのデータを出力することが出来ませんでした。</font></h4>
<table align=center width=95%>
	<tr>
		<td valign=Top>

		</td>
		<td>
			入力項目にエラーがありました。下記の条件を満たすものはデータベースのデータを出力することが出来ません。
			もう一度よく確かめてもう再度入力して下さい。
		</td>
	</tr>
</table>
<br>
<table align=center>
		<ul>
		<h4 align=center><font color=red>社員CDには必ず整数を入力して下さい。</font></h4>
		<li><font color=red>社員CD</font>にハイフン( - )は入力しないで下さい。
		<li><font color=red>社員CD</font>にドット( . )は入力しないで下さい。
		<li><font color=red>社員CD</font>にカンマ( , )は入力しないで下さい。
		<li><font color=red>社員CD</font>に\マーク入力しないで下さい。
		</ul>
</table>
<table width=80% align=center>
		※ 以上の項目をもう一度確かめたうえで、再度入力し、出力処理を行って下さい。それでも出力されない場合は<a href="">管理者</a>に問い合わせてください。
</table>
<form action="CSV.asp" target="Right" method="Post">
<p align=center><input type="submit" value="戻 る">
</form>
</body>
</html>