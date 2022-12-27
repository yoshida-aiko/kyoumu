<html>
<head>
	<title>社員管理</title>
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>

<!--#INCLUDE FILE="include01.asp"-->

<hr>
<br>
<h5 align=center><font color=Green>このデータを登録してもよろしいですか？</font></h5>
<form action="SQLexe.asp" method="post" id=form1 name=form1>
	<input type="hidden" name="FLG" value="<%= g_sFLG %>">
	<table border=1 align="center" CELLPADDING="5" CELLSPACING="1">
		<tr>
			<td>
				社員CD</td><td align=center><%= FixZero(w_sCD,4) %>
				<input type="hidden" name=社員CD value=<%= w_sCD %>>
			</td>
		</tr>
		<tr>
			<td>
				社員名称</td><td align=center><%= w_sName %>
				<input type="hidden" name=社員名称 value=<%= w_sName %>>
			</td>
		</tr>
		<tr>
			<td>
				生年月日</td><td align=center><%= w_sBirthday %>
				<input type="hidden" name=生年月日 value="<%= w_sBirth %>">
			</td>
		</tr>
		<tr>
			<td>
				電話番号1</td><td align=center><%= w_sTelphone1 %>
				<input type="hidden" name=電話番号1 value="<%= w_sTel1 %>">
			</td>
		</tr>
		<tr>
			<td>
				電話番号2</td><td align=center><%= w_sTelphone2 %>
				<input type="hidden" name=電話番号2 value="<%= w_sTel2 %>">
			</td>
		</tr>
		<tr>
			<td>
				郵便</td><td align=center><%= w_sPostPost %>
				<input type="hidden" name=郵便 value="<%= w_sPost %>">
			</td>
		</tr>
		<tr>
			<td>
				住所</td><td align=center><%= w_sAdd %>
				<input type="hidden" name=住所1 value="<%= w_sAddress1 %>">
				<input type="hidden" name=住所2 value="<%= w_sAddress2 %>">
			</td>
		</tr>
		<tr>
			<td>
				備考</td><td align=center><pre><%= w_sIndex %></pre>
				<input type="hidden" name=備考 value="<%= w_sBikou %>">
			</td>
		</tr>
	</table>
	<h5 align=center><font color=Green>よければOKボタンを押してください。</font></h5>
	<table align="center" width=20%>
		<tr>
			<td align=center>
				<INPUT TYPE="submit" VALUE=" O K " onclick="Message()">
			</td>
</FORM>
			<td align=center>
				<INPUT TYPE="button" VALUE="キャンセル" onClick=history.go(-1) id=button1 name=button1>
			</td>
		</tr>
	</table>
</body>
</html>

<%

'*******************************************************************
'　　ゼロ埋め
'*******************************************************************
function FixZero(n, l) 'as string
	FixZero = right(string(l, "0") & n, l)
end function

%>
