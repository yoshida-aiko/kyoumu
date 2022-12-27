<html>
<head>
	<title>社員管理</title>
	<base target="Right">
<SCRIPT LANGUAGE="VBS">
Sub Prev_OnClick
	Prev.w_PageCount.value = "Prev"
End Sub
Sub Nexts_OnClick
	Nexts.w_PageCount.value = "Next"
	document.write
	exit sub
End Sub
</SCRIPT>
</head>
<body>
	<h3 align=center>★ 社員マスタメンテ ★</h3>
<hr>
<br>
<table align=center border=1 width="80%" bordercolor=#c0c0c0>
	<tr>
		<td width="17%" align=middle>
			<P>社員CD</P></td>
		<td width="63%" align=middle>社員名称</td>
	</tr>
</table>
<center>
	<IFRAME SRC="STAFFLIST.asp" name="INitiran" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="60%" marginheight=0></IFRAME>
</center>
<br><hr>

<table align=center width="80%">
	<tr>
		<form action="STAFFLIST.asp" target="INitiran" name="Prev">
		<input type="hidden" name="w_PageCount" value="">
		<td align="left">
			<input type="submit" value="≪前の10件" name="Prev">
		</td>
		</form>
		<td>
		<form action="ADDNEW.asp" target="Right" method="post" id=form1 name=form1>
			<td align=middle>
				<input type="submit" value="新 規" id=submit1 name=submit1>
			</td>
		</form>
		</td>
		<form action="Top.htm" target="Right" id=form2 name=form2>
			<td align=middle>
				<input type="submit" value="戻 る" id=submit2 name=submit2>
			</td>
		</form>
		<form action = "STAFFLIST.asp" target="INitiran" name="Nexts">
		<input type="hidden" name="w_PageCount" value="">
		<td align="right">
			<input type="submit" value="次の10件≫" name="Nexts">
		</td>
		</form>
	</tr>
</table>
</body>
</html>