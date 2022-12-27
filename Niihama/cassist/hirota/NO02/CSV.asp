<html>
<head>
<title>社員管理</title>
	<base target="Right">
	
<SCRIPT LANGUAGE="VBS">

Sub Submit_OnClick()
if CSV.txtStartCD.value = "" then
	if CSV.txtEndCD.value="" then
		if CSV.txtName.value="" then
			Msgbox "名前に不適応なデータが入力されています。",16,"入力エラー"
		end if
	end if
end if
end Sub

</SCRIPT>

</head>

<body>

<h3 align=center>★ 社員マスタCSV出力 ★</h3>

<p><HR></p>
<br>
<br>
<br>
<form action="SYUTUcsv.asp" method="post" name="CSV">
<h4 align=center>以下の条件で、社員マスタをCSV出力します。</h4>
<br>
<table CELLSPACING="0" CELLPADDING="12" ALIGN="CENTER">
<tr>
	<td>◆　社員CD</td>
	<td><input type=text name="txtStartCD"size=15 style="ime-mode:inactive" maxlength=4>
		　～　<input type=text name=txtEndCD size=15 style="ime-mode:inactive" maxlength=4></td>
</tr>
<tr>
	<td>◆　社員名称</td>
	<td>
		<input type=text name="txtName"size=42 maxlength="30" style="ime-mode:active">
	</td>
</tr>
<tr>
	<td>
	</td>
	<td>
		<font color=red>※あいまい検索</font>
	</td>
</tr>

<tr>
	<td>◆　削除フラグ</td>
	<td><input type="checkbox" name="checkDel" value=1>削除済みのデータは出力しない。
</tr>
</TABLE>
<br>
<br>
<br>
<br>
<br>

<hr>
<br>
<table align="center" width=20%>
	<tr>
		<td align=center>
			<INPUT TYPE="submit" VALUE="出 力" name="Submit"></form>
		</td>
		<td align=center>
			<form action="INitiran.asp" target="Right"><INPUT TYPE="submit" VALUE="一 覧" id=submit2 name=submit2></form>
		</td>
	</tr>
</table>
	</td>
</tr>
</form>
</body>
</html>