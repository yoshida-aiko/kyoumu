<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
<%
	w_sFLG = Request.QueryString("FLG")
	Select Case w_sFLG
		Case "1"
			Response.Write "<h3 align=center>★ 新 規 登 録 完 了 ★</h3>"
			Response.Write "<hr>"
			Response.Write "<h3 align=center><font color=red>社員データを登録しました！</font></h3>"

		Case "2"
			Response.Write "<h3 align=center>★ 修 正 完 了 ★</h3>"
			Response.Write "<hr>"
			Response.Write "<h3 align=center><font color=red>社員データを修正しました！</font></h3>"

		Case "3"
			Response.Write "<h3 align=center>★ 削　除 完 了 ★</h3>"
			Response.Write "<hr>"
			Response.Write "<h3 align=center><font color=red>社員データを削除しました！</font></h3>"
	end Select
 %>
<form action="INitiran.asp" target="Right" id=form1 name=form1>
	<p align=center><input type="submit" value="一 覧" id=submit1 name=submit1></p>
</form>

</body>
</html>
