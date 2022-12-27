
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
<table align=center width=95%>
	<tr>
		<td valign=Top></td>
		<td>
			入力項目にエラーがありました。下記の条件を満たすものはデータベースに登録することが出来ません。
			もう一度よく確かめてもう再度入力して下さい。
		</td>
	</tr>
</table>
<br>
<table align=center width=85%>
	<ul>
	<li><font color=red>社員CD</font>と<font color=red>社員名称</font>は必ず記入してください。<br>
	(ただし<font color=red>社員CD</font>、<font color=red>社員名称</font>以外の項目に関しては記入しなくても構いません。)
	<li><font color=red>生年月日</font>を選ぶ場合は必ず年、月、日をすべて選択して下さい。
	<li><font color=red>生年月日</font>は存在しない日にちを入力しないで下さい。
	<li><font color=red>社員名称</font>、<font color=red>住所</font>、<font color=red>備考</font>にHTMLタグなどは入力しないで下さい。
	<li><font color=red>電話番号</font>を入力する時は、ハイフン( - )区切りで入力して下さい。
	<li><font color=red>郵便番号</font>を入力する時は、 3桁 - 4桁、もしくは 3桁 で入力して下さい。
	</ul>
	※ 以上の項目をもう一度確かめたうえで、再度登録してください。それでも登録されない場合は<a href="">管理者</a>に問い合わせてください。
</table>

<% Response.Write Request.Form("Msg") %>

<p align=center><input type="button" value="戻る" onclick=history.go(-1)>

</body>
</html>
