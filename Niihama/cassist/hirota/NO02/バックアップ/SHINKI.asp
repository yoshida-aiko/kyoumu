<HTML>

<HEAD>
<TITLE>社員管理</TITLE>

	<base target="Right">

<script language="JavaScript">
<!--
//カスタマイズされたonkeydownイベントハンドラ
function customOnKeyDown( usrType )
{
 actionKey = ( document.layers ) ? usrType.which : event.keyCode ;
 // alert ( actionKey ) ; デバック用

 if ( ( 47 < actionKey && actionKey < 58 ) || ( 95 < actionKey && actionKey < 106 ) )
 {
  return true ;
 }
// BackSpace Key, Delete Key, 左矢印, 右矢印, Enter Key の場合、入力を許可
 else if ( actionKey == 8 || actionKey == 9 ||
             actionKey == 37 || actionKey == 39 ||
             actionKey == 46 ) 
 {
  return true ;
 }
 {
  return false ;
 }
}

//onkeydownイベントハンドラを置き換える
function setOnKeyDown( obj )
{
 obj.onkeydown = customOnKeyDown ;
}
// -->

//右クリックの禁止
function forbidIt(){ 
	if(document.all){ 
		if(event.button == 2){ 
			alert("右クリックは禁止！");
		}
	}else if(document.layers){
		if(myEvent.which == 3){ 
			alert("右クリックは禁止！");
		}
	}
}
if(document.layers)document.captureEvents(Event.MOUSEDOWN);
document.onmousedown = forbidIt;
//----->
</script>

<!--#INCLUDE FILE="CheckStaff.asp"-->

</HEAD>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<BODY>
	<h3 align=center>★　新規登録画面　★</h3>
	<HR>

<FORM ACTION="STAFFKAKUNIN.asp" METHOD="POST" target="Right" name="Tanpyou">
<input type="hidden" name="FLG" value="1">
<TABLE CELLSPACING="0" CELLPADDING="10" ALIGN="CENTER">
<TR>
	<TD>◆　社員CD</TD><TD><INPUT TYPE="text" NAME="txtCD" SIZE=22 maxlength=4 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )">
	<font color=red size=0.5>*必須入力</font></TD>
</TR>
<TR>
	<TD>◆　社員名称</TD><TD><INPUT TYPE="text" NAME="txtName" SIZE=36 maxlength=20 style="ime-mode:active">
	<font color=red size=0.5>*必須入力</font></TD>
</TR>
<TR>
	<TD>◆　生年月日</TD><TD>
	<select name="txtYEAR">
		<Option selected></Option>
		<% For i = (YEAR(now)-90) to YEAR(now) %>
			<Option value=<%= i %>><%= i %>
		<% Next %>
	</select> 年
	<select name="txtMONTH">
		<Option selected></Option>
		<% For i = 1 to 12 %>
			<Option value=<%= i %>><%= i %>
		<% Next %>
	</select> 月
	<select name="txtDAY">
		<Option selected></Option>
		<% For i = 1 to 31 %>
			<Option value=<%= i %>><%= i %>
		<% Next %>
	</select> 日
</TD>
</TR>
<TR>
	<TD>◆　電話番号1</TD><TD><INPUT TYPE="text" NAME="txtTel1" SIZE=7 maxlength=6 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"> - 
							<INPUT TYPE="text" NAME="txtTel2" SIZE=7 maxlength=7 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"> - 
							<INPUT TYPE="text" NAME="txtTel3" SIZE=7 maxlength=6 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )">
	</TD>
</TR>
<TR>
	<TD>◆　電話番号2</TD><TD><INPUT TYPE="text" NAME="txtTel4" SIZE=7 maxlength=6 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel5" SIZE=7 maxlength=7 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel6" SIZE=7 maxlength=6 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )">
	</TD>
</TR>
<TR>
	<TD>◆　郵便</TD><TD><INPUT TYPE="text" NAME="txtPost1" SIZE=7 maxlength=3 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )">
								- <INPUT TYPE="text" NAME="txtPost2" SIZE=7 maxlength=4 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"></TD>
</TR>
<TR>
	<TD>◆　住所</TD><TD><INPUT TYPE="text" NAME="txtAddress1" SIZE=50 maxlength=30 style="ime-mode:active"><BR>
       						<INPUT TYPE="text" NAME="txtAddress2" SIZE=50 maxlength=30 style="ime-mode:active"></TD>
</TR>
<TR>
	<TD>◆　備考</TD><TD><TEXTAREA ROWS="5" COLS="35" NAME="txtBikou" maxlength=50 style="ime-mode:active"></TEXTAREA></TD>
</TR>
</TABLE>

	<br>

<table align="center" width=20%>
	<tr>
		<td align=center>
			<INPUT TYPE="submit" VALUE="更 新" name="Submit"></FORM>
		</td>
		<td align=center>
			<form action="INitiran.asp" target="Right"><INPUT TYPE="submit" VALUE="一 覧"></form>
		</td>
	</tr>
</table>

</BODY>

</HTML>