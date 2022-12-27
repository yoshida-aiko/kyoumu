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
<SCRIPT LANGUAGE="VBS">
Function MsgOK()
	MsgStr=Msgbox("修正してもよろしいですか？",vbOkCancel + vbInformation,"登録")
	if MsgStr=vbCancel then
		window.event.returnValue=false
		Exit Function
	end if
End Function
</SCRIPT>
</HEAD>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>★　修正画面　★</h3>
	<HR>
<%
	On Error Resume Next
	
	Dim w_cCn,w_rRs,SQL,w_Birth,w_Tel1,w_Tel2,w_Post
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_社員",w_cCn,2,2
    
	SQL="SELECT * FROM M_社員 WHERE 社員CD=" & Request.Form("社員CD") & " AND 使用FLG=1 ORDER BY 1 ASC"
	
	Set w_rRs=w_cCn.Execute(SQL)
	
'*******************************************************************
'　　ゼロ埋め
'*******************************************************************
	function FixZero(n, l) 'as string
		FixZero = right(string(l, "0") & n, l)
	end function

%>

<FORM ACTION="SQLexe.asp" METHOD="POST" target="Right" name="Tanpyou">
<% Session.Contents("FLG")="UPDATE" %>
<input type="hidden" name="FLG" value="2">
<TABLE CELLSPACING="0" CELLPADDING="10" ALIGN="CENTER">
<TR>
	<TD>◆　社員CD</TD><TD><INPUT TYPE="text" NAME="txtCD" SIZE=22 disabled value="<%= FixZero(w_rRs("社員CD"),4) %>">
						<input type="hidden" name="CD" value=<%= w_rRs("社員CD") %>>
	<font color=red size=0.5>*必須入力</font></TD>
</TR>
<TR>
	<TD>◆　社員名称</TD><TD><INPUT TYPE="text" NAME="txtName" SIZE=36 
						value="<%= w_rRs("社員名称") %>" maxlength="20" style="ime-mode:active">
	<font color=red size=0.5>*必須入力</font></TD>
</TR>
<TR>
	<TD>◆　生年月日</TD><TD>
	<select name="txtYEAR">
		<% if IsNull(w_rRs("生年月日"))=false then
			w_Birth=split(w_rRs("生年月日"),"/") %>
			<% For i = (YEAR(now)-90) to YEAR(now) %>
				<% if Clng(i) = Clng(w_Birth(0)) then %>
					<Option Selected value="<%= i %>"><%= i %>
				<% else %>
					<Option value=<%= i %>><%= i %>
				<% end if %>
			<% Next %>
		<% else %>
			<Option selected value="">
			<% For i = (YEAR(now)-90) to YEAR(now) %>
				<Option value=<%= i %>><%= i %>
			<% Next %>
		<% end if %>
	</select> 年
		
	<select name="txtMONTH">
	<% if IsNull(w_rRs("生年月日"))=false then
		w_Birth=split(w_rRs("生年月日"),"/") %>
		<% For i = 1 to 12 %>
			<% if Clng(i) = Clng(w_Birth(1)) then %>
				<Option Selected value="<%= i %>"><%= i %>
			<% else %>
				<Option value=<%= i %>><%= i %>
			<% end if %>
		<% Next %>
	<% else %>
		<Option selected><%= NULL %></Option>
		<% For i = 1 to 12 %>
			<Option value=<%= i %>><%= i %>
		<% Next %>
	<% end if %>
	</select> 月
		
	<select name="txtDAY">
	<% if IsNull(w_rRs("生年月日"))=false then
		w_Birth=split(w_rRs("生年月日"),"/") %>
		<% For i = 1 to 31 %>
			<% if Clng(i) = Clng(w_Birth(2)) then %>
				<Option Selected value="<%= i %>"><%= i %>
			<% else %>
				<Option value=<%= i %>><%= i %>
			<% end if %>
		<% Next %>
	<% else %>
		<Option selected></Option>
		<% For i = 1 to 31 %>
			<Option value=<%= i %>><%= i %>
		<% Next %>
	<% end if %>
	</select> 日
</TD>
</TR>

<% if isnull(w_rRs("電話番号1"))=false then
		w_Tel1=split(w_rRs("電話番号1"),"-",-1,1) %>
<TR>
	<TD>◆　電話番号1</TD><TD><INPUT TYPE="text" NAME="txtTel1" SIZE=7 value="<%= w_Tel1(0) %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel2" SIZE=7 value="<%= w_Tel1(1) %>" maxlength=7
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel3" SIZE=7 value="<%= w_Tel1(2) %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
						</TD>
</TR>
<% else %>
<TR>
	<TD>◆　電話番号1</TD><TD><INPUT TYPE="text" NAME="txtTel1" SIZE=7 value="<%= NULL %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel2" SIZE=7 value="<%= NULL %>" maxlength=7
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel3" SIZE=7 value="<%= NULL %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
						</TD>
</TR>
<% end if %>

<% if isnull(w_rRs("電話番号2"))=false then
		w_Tel2=split(w_rRs("電話番号2"),"-",-1,1) %>
<TR>
	<TD>◆　電話番号2</TD><TD><INPUT TYPE="text" NAME="txtTel4" SIZE=7 value="<%= w_Tel2(0) %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel5" SIZE=7 value="<%= w_Tel2(1) %>" maxlength=7
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel6" SIZE=7 value="<%= w_Tel2(2) %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
						</TD>
</TR>
<% else %>
<TR>
	<TD>◆　電話番号2</TD><TD><INPUT TYPE="text" NAME="txtTel4" SIZE=7 value="<%= NULL %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel5" SIZE=7 value="<%= NULL %>" maxlength=7
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel6" SIZE=7 value="<%= NULL %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
						</TD>
</TR>
<% end if %>

<% if isnull(w_rRs("郵便"))=false then
	w_Post=split(w_rRs("郵便"),"-",-1,1)
	if instr(w_rRs("郵便"),"-")=0 then %>
		<TR>
			<TD>◆　郵便</TD><TD><INPUT TYPE="text" NAME="txtPost1" SIZE=7 value="<%= w_Post(0) %>" maxlength=3
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
								- <INPUT TYPE="text" NAME="txtPost2" SIZE=7 value="<%= NULL %>" maxlength=4
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
							</TD>
		</TR>
	<% else %>
		<TR>
			<TD>◆　郵便</TD><TD><INPUT TYPE="text" NAME="txtPost1" SIZE=7 value="<%= w_Post(0) %>" maxlength=3
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
							- <INPUT TYPE="text" NAME="txtPost2" SIZE=7 value="<%= w_Post(1) %>" maxlength=4
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
							</TD>
		</TR>
	<% end if %>
<% else %>
<TR>
	<TD>◆　郵便</TD><TD><INPUT TYPE="text" NAME="txtPost1" SIZE=7 value="<%= NULL %>" maxlength=3
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
								- <INPUT TYPE="text" NAME="txtPost2" SIZE=7 value="<%= NULL %>" maxlength=4
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
					</TD>
</TR>
<% end if %>

<TR>
	<TD>◆　住所</TD><TD><INPUT TYPE="text" NAME="txtAddress1" SIZE=50 value="<%= w_rRs("住所1") %>"
					 maxlength=20 style="ime-mode:active"><BR>
       						<INPUT TYPE="text" NAME="txtAddress2" SIZE=50 value="<%= w_rRs("住所2") %>"
       							 maxlength=20 style="ime-mode:active"></TD>
</TR>
<TR>
	<TD>◆　備考</TD><TD>
	<TEXTAREA ROWS="5" COLS="35" NAME="txtBikou" style="ime-mode:active"><%= w_rRs("備考") %></TEXTAREA></TD>
</TR>
</TABLE>
<input type="hidden" name="NAME">
<input type="hidden" name="BIRTHDAY">
<input type="hidden" name="TELL1">
<input type="hidden" name="TELL2">
<input type="hidden" name="POST">
<input type="hidden" name="ADDRESS1">
<input type="hidden" name="ADDRESS2">
<input type="hidden" name="BIKOU">
<p></p>

<table align=center width=20%>
	<tr>
		<td align=center>
			<INPUT TYPE="submit" VALUE="更 新" name="Submit">
		</td>
</form>
			<td align=center>
		<form action="INitiran.asp" target="Right">
				<INPUT TYPE="submit" VALUE="一 覧">
			</td>
		</FORM>
	</tr>
<%
	w_rRs.Close
	Set w_rRs = Nothing
	w_cCn.Close
	Set w_cCn = Nothing
%>
</BODY>

</HTML>

