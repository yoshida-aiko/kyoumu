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

</HEAD>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<BODY>
	<h3 align=center>���@�V�K�o�^��ʁ@��</h3>
	<HR>

<FORM ACTION="STAFFKAKUNIN.asp" METHOD="POST" target="Right" name="Tanpyou">
<input type="hidden" name="FLG" value="1">
<TABLE CELLSPACING="0" CELLPADDING="10" ALIGN="CENTER">
<TR>
	<TD>���@�Ј�CD</TD><TD><INPUT TYPE="text" NAME="txtCD" SIZE=22 maxlength=4 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )">
	<font color=red size=0.5>*�K�{����</font></TD>
</TR>
<TR>
	<TD>���@�Ј�����</TD><TD><INPUT TYPE="text" NAME="txtName" SIZE=36 maxlength=20 style="ime-mode:active">
	<font color=red size=0.5>*�K�{����</font></TD>
</TR>
<TR>
	<TD>���@���N����</TD><TD>
	<select name="txtYEAR">
		<Option selected></Option>
		<% For i = (YEAR(now)-90) to YEAR(now) %>
			<Option value=<%= i %>><%= i %>
		<% Next %>
	</select> �N
	<select name="txtMONTH">
		<Option selected></Option>
		<% For i = 1 to 12 %>
			<Option value=<%= i %>><%= i %>
		<% Next %>
	</select> ��
	<select name="txtDAY">
		<Option selected></Option>
		<% For i = 1 to 31 %>
			<Option value=<%= i %>><%= i %>
		<% Next %>
	</select> ��
</TD>
</TR>
<TR>
	<TD>���@�d�b�ԍ�1</TD><TD><INPUT TYPE="text" NAME="txtTel1" SIZE=7 maxlength=6 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"> - 
							<INPUT TYPE="text" NAME="txtTel2" SIZE=7 maxlength=7 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"> - 
							<INPUT TYPE="text" NAME="txtTel3" SIZE=7 maxlength=6 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )">
	</TD>
</TR>
<TR>
	<TD>���@�d�b�ԍ�2</TD><TD><INPUT TYPE="text" NAME="txtTel4" SIZE=7 maxlength=6 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel5" SIZE=7 maxlength=7 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"> - 
						<INPUT TYPE="text" NAME="txtTel6" SIZE=7 maxlength=6 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )">
	</TD>
</TR>
<TR>
	<TD>���@�X��</TD><TD><INPUT TYPE="text" NAME="txtPost1" SIZE=7 maxlength=3 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )">
								- <INPUT TYPE="text" NAME="txtPost2" SIZE=7 maxlength=4 style="ime-mode:inactive"
 onFocus="setOnKeyDown( this )"></TD>
</TR>
<TR>
	<TD>���@�Z��</TD><TD><INPUT TYPE="text" NAME="txtAddress1" SIZE=50 maxlength=30 style="ime-mode:active"><BR>
       						<INPUT TYPE="text" NAME="txtAddress2" SIZE=50 maxlength=30 style="ime-mode:active"></TD>
</TR>
<TR>
	<TD>���@���l</TD><TD><TEXTAREA ROWS="5" COLS="35" NAME="txtBikou" maxlength=50 style="ime-mode:active"></TEXTAREA></TD>
</TR>
</TABLE>

	<br>

<table align="center" width=20%>
	<tr>
		<td align=center>
			<INPUT TYPE="submit" VALUE="�X �V" name="Submit"></FORM>
		</td>
		<td align=center>
			<form action="INitiran.asp" target="Right"><INPUT TYPE="submit" VALUE="�� ��"></form>
		</td>
	</tr>
</table>

</BODY>

</HTML>