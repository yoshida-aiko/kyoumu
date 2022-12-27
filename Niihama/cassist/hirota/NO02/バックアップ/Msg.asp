<html>
<head>
	<title>社員管理</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<%
w_sFLG = Request.QueryString("FLG")
Select Case w_sFLG
	Case "1"
%>
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
			<li><font color=red>生年月日</font>を選ぶ場合は必ず年、月、日をすべて選択して下さい。
			<li><font color=red>生年月日</font>は存在しない日にちを入力しないで下さい。
			<li><font color=red>社員名称</font>、<font color=red>住所</font>、<font color=red>備考</font>にHTMLタグなどは入力しないで下さい。
			<li><font color=red>電話番号</font>を入力する時は、ハイフン( - )区切りで入力して下さい。
			<li><font color=red>郵便番号</font>を入力する時は、 3桁 - 4桁、もしくは 3桁 で入力して下さい。
			</ul>
			※ 以上の項目をもう一度確かめたうえで、再度登録してください。それでも登録されない場合は<a href="">管理者</a>に問い合わせてください。
		</table>

		<% Response.Write Request.Form("Msg") %>

		<p align=center><input type="button" value="戻る" onclick=history.go(-1) id=button1 name=button1>

<% Case "2" %>
		<body>
			<h3 align=center>■ 登録エラー ■</h3>
				<hr><br>
			<h4 align=center><font color=red>※データベースに登録することが出来ませんでした。</font></h4>
		<table align=center>
		<tr>
			<td>
				<ul>
				<li>社員CD <%= Session.Contents("w_sCD") %> はすでに使われています。
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

<% Case "3" %>

		<body>
			<h3 align=center>■ 重複データ ■</h3>
				<hr><br>
			<h4 align=center><font color=red>※データベースデータ重複メッセージ</font></h4>
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
			<input type="hidden" name="社員CD" value="<%= Session.Contents("社員CD") %>">
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
					<td align=center><input type="submit" value="O K" id=submit2 name=submit2></td>
		</FORM>
			<form action="SHINKI.asp" target="Right" method="Post" id=form2 name=form2>
					<td align=center><input type="submit" value="登録キャンセル" id=submit1 name=submit1></td>
			</form>
				</tr>
		</table>
<% Case "4" %>
		<body>
			<h3 align=center>■ 登録エラー ■</h3>
				<hr><br>
			<h4 align=center><font color=red>※入力項目にエラーがあります。<br>
					データベースに登録することが出来ませんでした。</font></h4>
		<br>
		<table align=center width=85%>
			<ul>
			<li><font color=red>社員CD</font>と<font color=red>社員名称</font>は必ず記入してください。<br>
			<li><font color=red>社員CD</font>、<font color=red>社員名称</font>以外の項目に関しては記入しなくても構いません。
			</ul>
			※ 再度登録処理を行ってください。それでも登録されない場合は<a href="">管理者</a>に問い合わせてください。
		</table>

		<% Response.Write Request.Form("Msg") %>

		<p align=center><input type="button" value="戻る" onclick=history.go(-1) id=button1 name=button1>

<% End Select %>
</body>
</html>

