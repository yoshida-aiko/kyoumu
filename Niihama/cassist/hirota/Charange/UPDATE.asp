<HTML>

<HEAD>

<TITLE>�Ј��Ǘ�</TITLE>

	<base target="Right">

<script language="JavaScript">
<!--
//�J�X�^�}�C�Y���ꂽonkeydown�C�x���g�n���h��
function customOnKeyDown( usrType )
{
 actionKey = ( document.layers ) ? usrType.which : event.keyCode ;
 // alert ( actionKey ) ; �f�o�b�N�p

 if ( ( 47 < actionKey && actionKey < 58 ) || ( 95 < actionKey && actionKey < 106 ) )
 {
  return true ;
 }
// BackSpace Key, Delete Key, �����, �E���, Enter Key �̏ꍇ�A���͂�����
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

//onkeydown�C�x���g�n���h����u��������
function setOnKeyDown( obj )
{
 obj.onkeydown = customOnKeyDown ;
}

// -->

//�E�N���b�N�̋֎~
function forbidIt(){ 
	if(document.all){ 
		if(event.button == 2){ 
			alert("�E�N���b�N�͋֎~�I");
		}
	}else if(document.layers){
		if(myEvent.which == 3){ 
			alert("�E�N���b�N�͋֎~�I");
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
	MsgStr=Msgbox("�C�����Ă���낵���ł����H",vbOkCancel + vbInformation,"�o�^")
	if MsgStr=vbCancel then
		window.event.returnValue=false
		Exit Function
	end if
End Function
</SCRIPT>
</HEAD>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>���@�C����ʁ@��</h3>
	<HR>
<%
	On Error Resume Next
	
	Dim w_cCn,w_rRs,SQL,w_Birth,w_Tel1,w_Tel2,w_Post
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_�Ј�",w_cCn,2,2
    
	SQL="SELECT * FROM M_�Ј� WHERE �Ј�CD=" & Request.Form("�Ј�CD") & " AND �g�pFLG=1 ORDER BY 1 ASC"
	
	Set w_rRs=w_cCn.Execute(SQL)
	
'*******************************************************************
'�@�@�[������
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
	<TD>���@�Ј�CD</TD><TD><INPUT TYPE="text" NAME="txtCD" SIZE=22 disabled value="<%= FixZero(w_rRs("�Ј�CD"),4) %>">
						<input type="hidden" name="CD" value=<%= w_rRs("�Ј�CD") %>>
	<font color=red size=0.5>*�K�{����</font></TD>
</TR>
<TR>
	<TD>���@�Ј�����</TD><TD><INPUT TYPE="text" NAME="txtName" SIZE=36 
						value="<%= w_rRs("�Ј�����") %>" maxlength="20" style="ime-mode:active">
	<font color=red size=0.5>*�K�{����</font></TD>
</TR>
<TR>
	<TD>���@���N����</TD><TD>
	<select name="txtYEAR">
		<% if IsNull(w_rRs("���N����"))=false then
			w_Birth=split(w_rRs("���N����"),"/") %>
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
	</select> �N
		
	<select name="txtMONTH">
	<% if IsNull(w_rRs("���N����"))=false then
		w_Birth=split(w_rRs("���N����"),"/") %>
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
	</select> ��
		
	<select name="txtDAY">
	<% if IsNull(w_rRs("���N����"))=false then
		w_Birth=split(w_rRs("���N����"),"/") %>
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
	</select> ��
</TD>
</TR>

<% if isnull(w_rRs("�d�b�ԍ�1"))=false then
		w_Tel1=split(w_rRs("�d�b�ԍ�1"),"-",-1,1) %>
<TR>
	<TD>���@�d�b�ԍ�1</TD><TD><INPUT TYPE="text" NAME="txtTel1" SIZE=7 value="<%= w_Tel1(0) %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel2" SIZE=7 value="<%= w_Tel1(1) %>" maxlength=7
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel3" SIZE=7 value="<%= w_Tel1(2) %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
						</TD>
</TR>
<% else %>
<TR>
	<TD>���@�d�b�ԍ�1</TD><TD><INPUT TYPE="text" NAME="txtTel1" SIZE=7 value="<%= NULL %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel2" SIZE=7 value="<%= NULL %>" maxlength=7
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel3" SIZE=7 value="<%= NULL %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
						</TD>
</TR>
<% end if %>

<% if isnull(w_rRs("�d�b�ԍ�2"))=false then
		w_Tel2=split(w_rRs("�d�b�ԍ�2"),"-",-1,1) %>
<TR>
	<TD>���@�d�b�ԍ�2</TD><TD><INPUT TYPE="text" NAME="txtTel4" SIZE=7 value="<%= w_Tel2(0) %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel5" SIZE=7 value="<%= w_Tel2(1) %>" maxlength=7
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel6" SIZE=7 value="<%= w_Tel2(2) %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
						</TD>
</TR>
<% else %>
<TR>
	<TD>���@�d�b�ԍ�2</TD><TD><INPUT TYPE="text" NAME="txtTel4" SIZE=7 value="<%= NULL %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel5" SIZE=7 value="<%= NULL %>" maxlength=7
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel6" SIZE=7 value="<%= NULL %>" maxlength=6
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
						</TD>
</TR>
<% end if %>

<% if isnull(w_rRs("�X��"))=false then
	w_Post=split(w_rRs("�X��"),"-",-1,1)
	if instr(w_rRs("�X��"),"-")=0 then %>
		<TR>
			<TD>���@�X��</TD><TD><INPUT TYPE="text" NAME="txtPost1" SIZE=7 value="<%= w_Post(0) %>" maxlength=3
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
								- <INPUT TYPE="text" NAME="txtPost2" SIZE=7 value="<%= NULL %>" maxlength=4
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
							</TD>
		</TR>
	<% else %>
		<TR>
			<TD>���@�X��</TD><TD><INPUT TYPE="text" NAME="txtPost1" SIZE=7 value="<%= w_Post(0) %>" maxlength=3
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
							- <INPUT TYPE="text" NAME="txtPost2" SIZE=7 value="<%= w_Post(1) %>" maxlength=4
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
							</TD>
		</TR>
	<% end if %>
<% else %>
<TR>
	<TD>���@�X��</TD><TD><INPUT TYPE="text" NAME="txtPost1" SIZE=7 value="<%= NULL %>" maxlength=3
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
								- <INPUT TYPE="text" NAME="txtPost2" SIZE=7 value="<%= NULL %>" maxlength=4
									 style="ime-mode:inactive" onFocus="setOnKeyDown( this )">
					</TD>
</TR>
<% end if %>

<TR>
	<TD>���@�Z��</TD><TD><INPUT TYPE="text" NAME="txtAddress1" SIZE=50 value="<%= w_rRs("�Z��1") %>"
					 maxlength=20 style="ime-mode:active"><BR>
       						<INPUT TYPE="text" NAME="txtAddress2" SIZE=50 value="<%= w_rRs("�Z��2") %>"
       							 maxlength=20 style="ime-mode:active"></TD>
</TR>
<TR>
	<TD>���@���l</TD><TD>
	<TEXTAREA ROWS="5" COLS="35" NAME="txtBikou" style="ime-mode:active"><%= w_rRs("���l") %></TEXTAREA></TD>
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
			<INPUT TYPE="submit" VALUE="�X �V" name="Submit">
		</td>
</form>
			<td align=center>
		<form action="INitiran.asp" target="Right">
				<INPUT TYPE="submit" VALUE="�� ��">
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

