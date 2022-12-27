<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->

<% w_sFLG = Request.QueryString("FLG")
   Select Case w_sFLG
   
 '----------------------------------------条件あり------------------------------------
	Case "1" %>
		<body>
			<h3 align=center>★ EXCEL出力 ★</h3>
				<hr>
			<h3 align=center>指定した条件で社員データを出力してもよろしいですか？</h3>

		<table align=center width=20%>
			<tr>
				<td align=center>
				<form action="EndCSV.asp" target="Right" method="Post" id=form1 name=form1>
					<p align=center><input type="submit" value="O K" id=submit1 name=submit1>
				</td>
				</form>
				<form action="CSV.asp" method="Post" target="Right" id=form2 name=form2>
				<td align=center>
					<input type="submit" value="キャンセル" id=submit2 name=submit2>
				</td>
				</form>
			<tr>
		</table>

<%
'----------------------------------------条件なし------------------------------------
	Case "2" %>
		<body>
			<h3 align=center>★ EXCEL出力 ★</h3>
				<hr>
			<h3 align=center>条件範囲の指定がありません。<br>すべての社員データを出力してもよろしいですか？</h3>
		<br>
		<table align=center width=20%>
			<tr>
				<form action="EndCSV.asp" target="Right" method="Post" id=form3 name=form3>
					<td align=center>
						<p align=center><input type="submit" value="O K" id=submit1 name=submit1>
					</td>
				</form>
				<form action="CSV.asp" target="Right" id=form6 name=form6>
					<td align=center>
						<input type="submit" value="戻 る" id=submit4 name=submit4>
					</td>
				</form>
			<tr>
		</table>
		</body>
		</html>
<%
'----------------------------------------該当者なし------------------------------------
	Case "3" %>
		<body>
			<h3 align=center>■ 出力エラー ■</h3>
			<hr>
		<table align=center>
		<tr>
		<td>
			<h4 align=center><font color=red>※ 出力エラーメッセージ</font></h4>
		</td>
		</tr>
		</table>
		<table align=center>
		<tr>
			<td align=center>
				<ul>
				<li>条件に該当する社員はいませんでした。
				</ul>
			</td>
		</tr>
		</table>


		<table align="center" width=20%>
			<tr><form action="CSV.asp" target="Right" method="Post" id=form4 name=form4>
				<td align=center>
					<input type="submit" value="戻 る" id=submit3 name=submit3>
				</td>
				</form>
				<form action="INitiran.asp" target="Right" id=form5 name=form5>
				<td align=center>
						<input type="submit" value="一覧" id=submit1 name=submit1>
				</td>
				</form>
			</tr>
		</table>

<% End Select %>
</body>
</html>
