<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>■ 登録エラー ■</h3>
		<hr><br>
	<h4 align=center><font color=red>※データベースに登録することが出来ませんでした。</font></h4>
<table align=center>
<tr>
	<td>
		<ul>
		<li>社員CD <%= Request.QueryString("CD") %> はすでに使われています。
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
<p align=center><input type="button" value="戻る" onClick=history.go(-1)>

</body>
</html>
