<html>
<head>
<title>社員管理</title>
	<base target="Right">

<!--#INCLUDE FILE="CheckVALUE.asp"-->

<SCRIPT LANGUAGE="VBS">
' ファイル名の入力チェック
Function CheckFileName()

	Dim CheckNameStr
	Dim NameStr
	Dim i
	CheckFileName=false
	CheckNameStr = "\/:;,*?<>|"
	if EXPORT.txtFileName.value = "" then
		Exit Function
	else
		For i = 1 to Len(EXPORT.txtFileName.value)
			NameStr=Mid(EXPORT.txtFileName.value,i,1)
			if Instr(CheckNameStr,NameStr) <> 0 then
				Exit Function
			end if
		Next
	end if
	CheckFileName=true
	
End Function
</SCRIPT>

</head>
<body>

<h3 align=center>★ 社員マスタEXCEL出力 ★</h3>

<p><HR></p>


<br>
<h4 align=center>以下の条件で、社員マスタをEXCEL出力します。</h4>
<br>
<form action="ExportEXCEL.asp" method="post" name="EXPORT">
<table CELLSPACING="0" CELLPADDING="12" ALIGN="CENTER">
	<tr>
		<td>◆　社員CD</td>
		<td><input type=text name="txtStartCD"size=15 style="ime-mode:inactive" maxlength=4>
			　～　<input type=text name=txtEndCD size=15 style="ime-mode:inactive" maxlength=4></td>
	</tr>
	<tr>
		<td>◆　社員名称</td>
		<td><input type=text name="txtName"size=42 maxlength="30" style="ime-mode:active"></td>
	</tr>
	<tr>
		<td></td>
		<td><font color=red>※あいまい検索</font></td>
	</tr>
	<tr>
		<td>◆　削除フラグ</td>
		<td><input type="checkbox" name="checkDel" value="1">削除済みのデータは出力しない。</td>
	</tr>
	<tr>
		<td>◆　出力先</td>
		<td><select name="cboName">
			<Option value="C:\WINDOWS\ﾃﾞｽｸﾄｯﾌﾟ\">ﾃﾞｽｸﾄｯﾌﾟ
			<Option value="Y:\宮井\">宮井さん<Option value="Y:\廣田\">廣田<Option value="Y:\矢野\">矢野
			<Option value="Y:\内田\">内田
			</select>
		</td>
	</tr>
	<tr>
		<td>◆　保存ファイル名</td>
		<td><input type="text" name="txtFileName" size=10 value="Sample" Maxlength="10" style="ime-mode:active">.xls<br></td>
	</tr>
</TABLE>
<input type="hidden" name="SQL">
<br>
<br>
<br>
<hr>
<br>
<table align="center" width=20%>
	<tr>
		<td align=center><INPUT TYPE="submit" VALUE="出 力" name="Submit"></form></td>
		<td align=center><form action="INitiran.asp" target="Right">
				<INPUT TYPE="submit" VALUE="一 覧" id=submit2 name=submit2></form></td>
	</tr>
</table>
	</td>
</tr>
</form>
</body>
</html>