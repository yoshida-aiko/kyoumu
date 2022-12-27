<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>■ 実行エラー ■</h3>
		<hr><br>
	<h4 align=center><font color=red>※データベースにアクセスすることが出来ませんでした。</font></h4>
	<p align=center>エラー：<%= Session.Contents("SQLerror") %></p>
	<p align=center><input type="button" value="戻る" onclick=history.go(-1)>
</body>
</html>
