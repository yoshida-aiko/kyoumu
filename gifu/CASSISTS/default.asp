
<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: 
' ��۸���ID : default.asp
' �@      �\: ���O�C��ID�E�p�X���[�h�̏ƍ����s��
'-------------------------------------------------------------------------
' ��      ��:���O�C��ID�A�p�X���[�h
' ��      ��:�Ȃ�
' ��      �n:
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/06/15 ���u �m��
' ��      �X: 2001/06/15 �≺ �K��Y
'*************************************************************************/
%>
<!--#include file="Common/com_All.asp"-->
<%


'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bErrMsg           '�װү����
	Public  m_SchoolName		'�w�Z��
	
'///////////////////////////���C������/////////////////////////////


    '�o�[�W�����̕\��
    'Response.Write "[ ORACLE Ver:" & OraSession.OIPVersionNumber & " ]"
    'Ҳ�ٰ�ݎ��s
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

Sub Main()
'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="�g�b�v"
	w_sMsg=""
	w_sRetURL= C_RetURL     
	w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do

        '// �ް��ް��ڑ�
        If gf_OpenDatabase() <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_bErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

	'// �Z�b�V�����N���A�[
	Call s_SessionClear

        '//�p�����[�^�Z�b�g
        If SetPara() = false Then
        	m_bErrFlg = True
        	Exit Do
        End If
        
        
        '//�w�Z�����擾
        If not f_GetSchoolName() Then
        	m_bErrFlg = True
        	Exit Do
        End If
        
	    '// �y�[�W��\��
'Call ErrPage("�e�X�g��")
	    Call showPage()
	    Exit Do
	Loop

	'// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
Call ErrPage(w_sMsg)
'		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
    
    '// �I������
    Call gs_CloseDatabase()
End Sub

Sub s_SessionClear()
'********************************************************************************
'*  [�@�\]  �Z�b�V�������N���A�[����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
'/** Dim Item
Dim w_OraDatabase 
Dim w_Qurs

	'// �K�v�ȃZ�b�V�����͕ϐ��ɓ����
	w_User_ID =	Session("USER_ID")
	w_PASS    =	Session("PASS")
	w_CONNECT =	Session("CONNECT")
	w_TYUGAKU_TIZU_PATH = Session("TYUGAKU_TIZU_PATH")

    SET w_OraDatabase = Session("OraDatabase")
    SET w_Qurs = Session("qurs")

	'�Z�b�V�����N���A�[
	for Each name in Session.Contents
'/** Response.Write name & " " & session(name) & "***"
                session(name) = ""
	next

	'// �Z�b�V�����ɖ߂�
	Session("USER_ID") = w_User_ID
	Session("PASS")    = w_PASS
	Session("CONNECT") = w_CONNECT
	Session("TYUGAKU_TIZU_PATH") = w_TYUGAKU_TIZU_PATH

    SET Session("OraDatabase") = w_OraDatabase
    SET Session("qurs") = w_Qurs

End Sub

Function SetPara() 
'********************************************************************************
'*  [�@�\]  �ϐ����Z�b�g
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	Dim w_sSQL,w_Rs,w_iRet
	Dim w_nendo
	
	SetPara = false

  Do
	w_sSQL = ""
	w_sSQL = w_sSQL & "SELECT "
	w_sSQL = w_sSQL & "M00_KANRI AS NENDO "
	w_sSQL = w_sSQL & "FROM "
	w_sSQL = w_sSQL & "M00_KANRI "
	w_sSQL = w_sSQL & "WHERE "
	w_sSQL = w_sSQL & "M00_NENDO = 9999 AND "
	w_sSQL = w_sSQL & "M00_NO = 0 AND "
	w_sSQL = w_sSQL & "M00_SYUBETU = 0 "

	Set w_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

	If w_iRet <> 0 Then
	'ں��޾�Ă̎擾���s
		m_bErrFlg = True
		Exit Do 'GOTO LABEL_MAIN_END
	End If

	'//�@�����N�x������
	Session("NENDO") = w_Rs("NENDO")
	Exit Do
  Loop	

	w_Rs = ""

  Do
	w_sSQL = ""
	w_sSQL = w_sSQL & "SELECT "
	w_sSQL = w_sSQL & "M00_KANRI AS GAKKI "
	w_sSQL = w_sSQL & "FROM "
	w_sSQL = w_sSQL & "M00_KANRI "
	w_sSQL = w_sSQL & "WHERE "
	w_sSQL = w_sSQL & "M00_NENDO = " & Session("NENDO") & " AND "
	w_sSQL = w_sSQL & "M00_NO = 11 AND "
	w_sSQL = w_sSQL & "M00_SYUBETU = 0 "

	Set w_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

	If w_iRet <> 0 Then
	'ں��޾�Ă̎擾���s
		m_bErrFlg = True
		Exit Do 'GOTO LABEL_MAIN_END
	End If

	'//�@�w�����Z�b�V�����ɓ����B
	If w_Rs("GAKKI") > gf_YYYY_MM_DD(date(),"/") Then 
		Session("GAKKI") = C_GAKKI_ZENKI
	Else 
		Session("GAKKI") = C_GAKKI_KOUKI
	End If
	SetPara = true
	Exit Do
  Loop	

	'// ��׳�ް���擾
	wBrauza = request.servervariables("HTTP_USER_AGENT")
	if InStr(wBrauza,"IE") <> 0 then
		session("browser") = "IE"
	Else
		session("browser") = "NN"
	End if

	call gf_closeObject(m_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �w�Z�����擾����B
'*  [����]  
'*  [����]  
'********************************************************************************
Function f_GetSchoolName()
	
	Dim w_sSQL
	Dim w_Rs
	Dim w_FieldName
	Dim w_Table,w_TableName,w_KamokuName
	
	On Error Resume Next
	Err.Clear
	
	f_GetSchoolName = false
	m_SchoolName = ""
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M19_NAME "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M19_GAKKO "
	
	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function
	
	if not w_Rs.EOF then
		m_SchoolName = w_Rs("M19_NAME") & "�@��U��"
		Call gf_closeObject(w_Rs)
	end if
	
	f_GetSchoolName = true
	
End Function

Sub ErrPage(p_sMsg)
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
<html>
<head>
<title>�L�����p�X�A�V�X�g</title>
</head>
<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" bgcolor="#F6F7FC" >
<center>
<table width="100%" height="100%" cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td width="100%" height="40%" background="image/back.gif" style="background-repeat: repeat-y "><img src="image/title.gif" width="504" height="214"><br><br></td>
	</tr>
	<tr>
		<td align="center" background="image/back.gif" style="background-repeat: repeat-y;">
		<%=p_sMsg%><BR><BR>
		<font color="#FF0000">
		�f�[�^�x�[�X�ڑ��Ɏ��s���܂���<BR>
		�Ǘ��҂ɘA�����Ă��������B
		</font>
		</td>
	</tr>
</table>
</center>
</body>
</html>
<%
End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear

	'// �t�H�[���T�C�Y
	if session("browser") = "NN" then
		wformSize = "15"
	Else
		wformSize = "20"
	End if

%>
<html>

<head>
<title>Campus Assist</title>
<!-- <link rel=stylesheet href="common/style.css" type=text/css> -->
<LINK REL="SHORTCUT ICON" href="image/CAtitle.ico">
<script language="javascript">
<!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		document.frm.txtLogin.focus();
    }

    //************************************************************
    //  [�@�\]  ���Z�b�g�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //  [�쐬��] 
    //************************************************************
	function f_clear() {
		document.frm.reset();
		return false;
	}
//-->
</script>
<style type="text/css">
<!--
   input { font-size:12px;}
   A {	 text-decoration:none; 
   		font-size:9pt;
   		text-align:center;
   	 }

   a:link {color:#222268;}
   a:visited {color:#222268;}
   a:active {color:#222268;}
   a:hover {color:#682222; text-decoration:underline; }
//-->
</style>
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" bgcolor="#ffffff" onLoad="window_onload();">

<table width="100%" height="100%" cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td nowrap width="25%" valign="top" rowspan="3"><%= "[ ORACLE Ver:" & OraSession.OIPVersionNumber & " ]" %></td>
		<td width="504" height="40%" background="image/back.gif"><img src="image/title.gif" width="504" height="214"><br><br></td>
		<td nowrap width="25%" rowspan="3">&nbsp;</td>
	</tr>
	
	<tr>
		<td height="50%" width="504" align="center" background="image/back.gif">
			
			<table cellspacing="0" cellpadding="0" width="244" height="140" border="0">
				<tr><td align="center" colspan="3"><font size="-1" color="#222268"><%=m_SchoolName%></font></td></tr>
				<tr><td colspan="3">&nbsp;</td></tr>
				
				<tr>
					<td height="5" width="5"><img src="image/table1.gif"></td>
					<td height="5" width="230" background="image/table2.gif"><img src="image/sp.gif"></td>
					<td height="5" width="9"><img src="image/table3.gif"></td>
				</tr>
				
				<tr>
					<td height="139" width="5" background="image/table4.gif"><img src="image/sp.gif"></td>
					<td height="139" width="230" bgcolor="#ffffff" align="center" background="image/sp.gif">
						
						<img src="image/sp.gif" height="1"><br>
						<form action="login/default.asp" name="frm" method="post">
						<table border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td><img src="image/login.gif" border="0"></td>
								<td><input type="text" size="<%=wformSize%>" name="txtLogin" value="<%= DC_USERADMIN %>"></td>
							</tr>
							<tr>
								<td colspan="2"><img src="image/sp.gif" height="5"></td>
							</tr>
							<tr>
								<td><img src="image/pass.gif" border="0"></td>
								<td><input type="password" size="<%=wformSize%>" name="txtPass" value="<%= DC_USERADMIN %>"></td>
							</tr>
							<tr>
								<td colspan="2"><img src="image/sp.gif" height="10"></td>
							</tr>
							<tr>
								<td colspan="2" align="center" valign="bottom"><input type="image" border="0" src="image/login_b.gif"><img src="image/sp.gif" width="35"><input type="image" border="0" src="image/clear.gif" onclick="return f_clear()"></td>
							</tr>
						</table>
		<% if gf_empPasChg() then %>
				<a href="web/web0400/default.asp">- �p�X���[�h�ύX�͂����� -</a>
		<% End if %>
					</td>
					<td height="139" width="9" background="image/table5.gif"><img src="image/sp.gif"></td>
				</tr>
				<tr>
					<td height="5" width="5"><img src="image/table6.gif"></td>
					<td height="5" width="230" background="image/table7.gif"><img src="image/sp.gif"></td>
					<td height="5" width="9"><img src="image/table8.gif"></td>
				</tr>
				<tr><td colspan="3">&nbsp;</td></tr>
			</table>
			
		</td>
	</tr>
	<tr>
		<td height="10%" width="504" valign="bottom" align="center" background="image/back.gif"><img src="image/info_logo.gif"></td>
	</tr>
</table>


<input type="hidden" name="hidLoginFlg" value="<%= C_LOGIN_FLG %>">

</form>
</body>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>
