<%
	On Error Resume Next

	Dim w_sCD,w_sName,w_sYEAR,w_sMONTH,w_sDAY,w_sTel1,w_sTel2,w_sTel3,w_sTel4,w_sTel5,w_sTel6
	Dim w_sPost1,w_sPost2,w_sAddress1,w_sAddress2,w_sBikou
	
' 受信した社員データを変数に格納
	w_sName = Request.Form("txtName")
	w_sYEAR = Request.Form("txtYEAR")
	w_sMONTH = Request.Form("txtMONTH")
	w_sDAY = Request.Form("txtDAY")
	w_sTel1 = Request.Form("txtTel1")
	w_sTel2 = Request.Form("txtTel2")
	w_sTel3 = Request.Form("txtTel3")
	w_sTel4 = Request.Form("txtTel4")
	w_sTel5 = Request.Form("txtTel5")
	w_sTel6 = Request.Form("txtTel6")
	w_sPost1 = Request.Form("txtPost1")
	w_sPost2 = Request.Form("txtPost2")
	w_sAddress1 = Request.Form("txtAddress1")
	w_sAddress2 = Request.Form("txtAddress2")
	w_sBikou = Request.Form("txtBikou")
	
' ******************************************************************
' 　社員名称の入力判断
' ******************************************************************
	if w_sName = "" then
		Response.Redirect "Msg.asp"
	end if
	
	Dim SQL
'*******************************************************************
'　　ゼロ埋め
'*******************************************************************
	function FixZero(n, l) 'as string
		FixZero = right(string(l, "0") & n, l)
	end function
	
'*******************************************************************
'　　タグが入力されたかどうかを判定
'*******************************************************************
	function f_CheckVALUE(p_VALUE)
	f_CheckVALUE = false
    If InStr(p_VALUE, "<") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, ">") <> 0 Then
        Exit Function
    End If
    f_CheckVALUE = true
    end function
%>

<html>
<head>
	<title>社員管理</title>
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>

<h3 align=center>★ 修正確認画面 ★</h3>

<p><hr></p>
<br>
<h5 align=center><font color=Orange>以下の社員データに修正してもよろしいですか？</font></h5>
<form action="SQLexe.asp" method="post">
<input type="hidden" name="FLG" value="2">
	<table border=1 align="center" CELLPADDING="5" CELLSPACING="1">
		<tr>
			<td>
				社員CD</td><td align=center><%= FixZero(Request.Form("CD"),4) %>
				<input type="hidden" name=社員CD value=<%= Request.Form("CD") %>>
			</td>
		</tr>
		<tr>
			<td>
				<%  if f_CheckVALUE(w_sName)=false then
					Response.Redirect "Msg.asp"
				end if %>
				社員名称</td><td align=center><%= w_sName %>
				<input type="hidden" name=社員名称 value=<%= w_sName %>>
			</td>
		</tr>
		<tr>
			<td>
			<% if w_sYEAR <> "" AND w_sMONTH <> "" AND w_sDAY <> "" then %>
				<% if IsDate(w_sYEAR & "/" & w_sMONTH & "/" & w_sDAY)=false then %>
						<% Response.Redirect "Msg.asp" %>
				<% else %>
						生年月日</td><td align=center><%= w_sYEAR & "年" & w_sMONTH & "月" & w_sDAY & "日" %>
						<input type="hidden" name=生年月日 value="<%= "'" & w_sYEAR & "/" & w_sMONTH & "/" & w_sDAY & "'" %>">
				<% end if %>
			<% elseif w_sYEAR ="" AND w_sMONTH = "" AND w_sDAY ="" then %>
				生年月日</td><td align=center><font color="red">記入無し</font>
				<input type="hidden" name=生年月日 value=NULL>
			<% else %>
				<% Response.Redirect "Msg.asp" %>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if w_sTel1 <> "" AND w_sTel2 <> "" AND w_sTel3 <> "" then %>
				電話番号1</td><td align=center><%= w_sTel1 & "-" & w_sTel2 & "-" & w_sTel3 %>
				<input type="hidden" name=電話番号1 value="<%= "'" & w_sTel1 & "-" & w_sTel2 & "-" & w_sTel3 & "'" %>">
			<% elseif w_sTel1 = "" AND w_sTel2 = "" AND w_sTel3 = "" then %>
				電話番号1</td><td align=center><font color="red">記入無し</font>
				<input type="hidden" name=電話番号1 value=NULL>
			<% else
				Response.Redirect "Msg.asp" %>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if w_sTel4 <> "" AND w_sTel5 <> "" AND w_sTel6 <> "" then %>
				電話番号2</td><td align=center><%= w_sTel4 & "-" & w_sTel5 & "-" & w_sTel6 %>
				<input type="hidden" name=電話番号2 value="<%= "'" & w_sTel4 & "-" & w_sTel5 & "-" & w_sTel6 & "'" %>">
			<% elseif w_sTel4 = "" AND w_sTel5 = "" AND w_sTel6 = "" then %>
				電話番号2</td><td align=center><font color="red">記入無し</font>
				<input type="hidden" name=電話番号2 value=NULL>
			<% else
				Response.Redirect "Msg.asp" %>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
				<%   if w_sPost1 = "" then
						if w_sPost2 = "" then %>
							郵便</td><td align=center><font color="red">記入無し</font>
							<input type="hidden" name=郵便 value=NULL>
				<%		else
							Response.Redirect "Msg.asp"
						end if
				   elseif w_sPost2 = "" then %>
						郵便</td><td align=center><%= w_sPost1 %>
						<input type="hidden" name=郵便 value="<%= "'" & w_sPost1 & "'" %>">
				<% else %>
				<%	if Len(w_sPost1) < 3 or Len(w_sPost2) < 4 then
							Response.Redirect "Msg.asp"
						end if %>
						郵便</td><td align=center><%= w_sPost1 & " - " & w_sPost2 %>
						<input type="hidden" name=郵便 value="<%= "'" & w_sPost1 & "-" & w_sPost2 & "'" %>">
				<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if w_sAddress1 <> "" then %>
				<% if w_sAddress2 <> "" then %>
					<%  if f_CheckVALUE(w_sAddress1)=false or f_CheckVALUE(w_sAddress2)=false then
							Response.Redirect "Msg.asp"
						end if %>
					住所</td><td align=center><%= w_sAddress1 %><br><%= w_sAddress2 %>
					<input type="hidden" name=住所1 value="<%= "'" & w_sAddress1 & "'" %>">
					<input type="hidden" name=住所2 value="<%= "'" & w_sAddress2 & "'" %>">
				<% else %>
					<%  if f_CheckVALUE(w_sAddress1)=false then
							Response.Redirect "Msg.asp"
						end if %>
					住所</td><td align=center><%= w_sAddress1 %><br>
					<input type="hidden" name=住所1 value="<%= "'" & w_sAddress1 & "'" %>">
					<input type="hidden" name=住所2 value=NULL>	
				<% end if %>
			<% elseif w_sAddress2 <> "" then %>
				<%  if f_CheckVALUE(w_sAddress2)=false then
							Response.Redirect "Msg.asp"
					end if %>
					住所</td><td align=center><br><%= w_sAddress2 %>
					<input type="hidden" name=住所1 value=NULL>	
					<input type="hidden" name=住所2 value="<%= "'" & w_sAddress2 & "'" %>">
			<% else %>
					住所</td><td align=center><font color="red">記入無し</font><br>
					<input type="hidden" name=住所1 value=NULL>
					<input type="hidden" name=住所2 value=NULL>	
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if w_sBikou <> "" then %>
				<%  if f_CheckVALUE(w_sBikou)=false then
						Response.Redirect "Msg.asp"
					end if %>
				備考</td><td align=center><%= w_sBikou %>
				<input type="hidden" name=備考 value="<%= "'" & w_sBikou & "'" %>">
			<% else %>
				備考</td><td align=center><font color="red">記入無し</font>
				<input type="hidden" name=備考 value=NULL>
			<% end if %>
			</td>
		</tr>
	</table>

<h5 align=center><font color=Orange>よければOKボタンを押してください。</font></h5>

<table align="center" width=20%>
	<tr>
		<td align=center>
			<INPUT TYPE="submit" VALUE=" O K " id=submit1 name=submit1>
		</td>
		</FORM>
		<td align=center>
			<INPUT TYPE="button" VALUE="キャンセル" onClick=history.go(-1) id=button1 name=button1>
		</td>
	</tr>
</table>
</body>
</html>
