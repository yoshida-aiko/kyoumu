<%
	On Error Resume Next
	
	Dim w_cCn,w_rRs,w_SQL
	
' オブジェクトの定義
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_社員",w_cCn,2,2
    
' リクエストされた社員データを抽出
	w_SQL="SELECT * FROM M_社員 WHERE 社員CD=" & Request.Form("社員CD")    

	Set w_rRs = w_cCn.Execute(w_SQL)
	
' SQL実行時のエラー処理
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	On Error Goto 0

'********************************************************************************
'	ゼロ埋め : format()
'********************************************************************************
	function FixZero(n, l) 'as string
		FixZero = right(string(l, "0") & n, l)
	end function
%>

<HTML>

<HEAD>
<TITLE>社員管理</TITLE>

	<base target="Right">

</HEAD>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>★　削除画面　★</h3>
	<HR>
	<h5 align=center><font color=Blue>削除すると元に戻すことは出来ません。<br>以下の社員データを削除してもよそしいですか？</h5>
<form action="SQLexe.asp" method="post" target="Right">
<input type="hidden" name="FLG" value="3">
	<table border=1 align="center" CELLPADDING="8" CELLSPACING="1">
		<tr>
			<td>
				社員CD</td><td align=center><%= FixZero(w_rRs("社員CD"),4) %>
				<input type="hidden" name=社員CD value=<%= w_rRs("社員CD") %>>
			</td>
		</tr>
		<tr>
			<td>
				社員名称</td><td align=center><%= w_rRs("社員名称") %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("生年月日"))=false then %>
				生年月日</td><td align=center><%= w_rRs("生年月日") %>
			<% else %>
				生年月日</td><td align=center><font color="red">記入無し</font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("電話番号1"))=false then %>
				電話番号1</td><td align=center><%= w_rRs("電話番号1") %>
			<% else %>
				電話番号1</td><td align=center><font color="red">記入無し</font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("電話番号2"))=false then %>
				電話番号2</td><td align=center><%= w_rRs("電話番号2") %>
			<% else %>
				電話番号2</td><td align=center><font color="red">記入無し</font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("郵便"))=false then %>
				郵便</td><td align=center><%= w_rRs("郵便") %>
			<% else %>
				郵便</td><td align=center><font color="red">記入無し</font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("住所1"))=false then %>
				<% if Isnull(w_rRs("住所2"))=false then %>
					住所</td><td align=center><%= w_rRs("住所1") %><br><%= w_rRs("住所2") %>
				<% else %>
					住所</td><td align=center><%= w_rRs("住所1") %><br>
				<% end if %>
			<% elseif isnull(w_rRs("住所2"))=false then %>
					住所</td><td align=center><br><%= w_rRs("住所2") %>
			<% else %>
					住所</td><td align=center><font color="red">記入無し</font><br>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("備考"))=false then %>
				備考</td><td align=center><pre><%= w_rRs("備考") %></pre>
			<% else %>
				備考</td><td align=center><font color="red">記入無し</font>
			<% end if %>
			</td>
		</tr>
	</table>

<h5 align=center><font color=Blue size="2">よければOKボタンを押して下さい。</font></h5>
	<table align=center width=20%>
		<tr>
			<td align=center>
				<input type="submit" value=" O K ">
			</td>
</form>
		<form action="INitiran.asp" target="Right">
			<td align=center>
				<input type="submit" value="キャンセル">
			</td>
		</form>
		</tr>
	</table>

</BODY>
</HTML>

<%
	w_rRs.Close
	Set w_rRs = Nothing
	w_cCn.Close
	Set w_cCn = Nothing
%>
