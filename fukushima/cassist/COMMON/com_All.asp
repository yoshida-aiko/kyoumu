<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ʏ���
' ��۸���ID : COM_LOGGET
' �@      �\: �G���[���O�擾���ʏ���
'-------------------------------------------------------------------------
' ��      ��: 2001.01.01 ���u
' ��      �X: 2001/06/16 ���u     Gf_IIF�ǉ�
' �@      �@: 2001/07/23 �ɓ�     gf_GetRecordset_OpenStatic�ǉ�
' �@      �@: 2001/07/24 �J�e     �֐��ꗗ�ǉ�
'           : 2002/05/06 ��O     
'             �@�@�@�@�@�@�@�@    �O���[�o���ϐ��̒ǉ�
'                                 gf_GetRecordset_oo4o�̒ǉ�
'*************************************************************************/

'//////////////////////////////////////////////////////////////////////////////////////////
'
'	�֐��ꗗ
'
'//////////////////////////////////////////////////////////////////////////////////////////
'�c�a�I�[�v��			 			gf_OpenDatabase()
'�n�c�a�b�c�a�I�[�v��	 			gf_openODBCDatabase(dbOpen)	
'�I���N���c�a�I�[�v��				gf_openORAADODatabase(dbOpen)
'�c�a�I�[�v��						AutoOpen(m_oConObj,p_sDbName,p_sConnect)
'�c�a�N���[�Y						gs_CloseDatabase()
'�I�u�W�F�N�g�N���[�Y				gf_closeObject(oClose)
'�E���ɃX�y�[�X�����Ă��낦��B	gf_Keta(p_Str,p_iSpace)
'�g�����U�N�V�����J�n				gs_BeginTrans()
'�g�����U�N�V���������[���o�b�N		gs_RollbackTrans()
'�g�����U�N�V�������R�~�b�g			gs_CommitTrans()
'���R�[�h�Z�b�g�̎擾 ADO��			gf_GetRecordset(p_Rs, p_sSQL)
'���R�[�h�Z�b�g�̎擾(pagesize�L)	gf_GetRecordsetExt(p_Rs, p_sSQL, p_iPageSize)
'���R�[�h�Z�b�g�̒��o				gf_GetRecordset_OpenStatic(p_Rs,p_sSQL)
'�y�[�W���̌v�Z						gf_PageCount(p_Rs, p_iDsp)
'���R�[�h�J�E���g�擾				gf_GetRsCount(p_Rs)
'��Βl�y�[�W�̐ݒ�					gs_AbsolutePage(p_Rs,p_iPage, p_iDsp)
'�󂾂�����S�p�X�y�[�X��Ԃ�		gf_HTMLTableSTR(p_Data)
'�t�B�[���h�l�̎擾					gf_GetFieldValue(p_Rs, p_sField)
'�r�p�k�̎��s						gf_ExecuteSQL(p_sSQL)
'�G���[���y�[�W�\��				gs_showMsgPage(p_sWinTitle, p_sMsgTitle, p_sMsgString, p_sRetURL, p_sTarget)
'�G���[���b�Z�[�W��ݒ�				gs_SetErrMsg(p_sErrMsg)
'�װү�������ق��擾����			gf_GetErrMsgTitle(p_sURL,p_sMsgTitle)
'�G���[���b�Z�[�W���擾����			gf_GetErrMsg()
'�G���[���b�Z�[�W���N���A����		gs_ClearErrMsg()
'�e�L�X�g�t�@�C����OPEN				gf_OpenFile(p_File,p_sPath)
'NULL���󔒂��n�C�t���ɕϊ�����		gf_SetNull2Haifun(p_sStr)
'null�`�F�b�N�iIsNull���ǂ��j		gf_IsNull(p_null)
'NULL���󕶎��ɕϊ�					gf_SetNull2String(p_sStr)
'NULL��0�ɕϊ�						gf_SetNull2Zero(p_iNum)
'YYYY/MM/DD�Ƀt�H�[�}�b�g			gf_YYYY_MM_DD(p_sDate,p_sDelimit)
'�a��t�H�[�}�b�g					gf_fmtWareki(pDate)
'�N�������`							gf_FormatDate(p_Date,p_Delimiter)
'���t�^�t�H�[�}�b�g�֐�(����)		FormatTime(datTime,strFormat)
'������̃o�C�g����Ԃ�				f_LenB(p_str)
'IIF�֐��̎���						gf_IIF(p_Judge, p_tStr, p_fStr)
'���l�����킹(�O�t�H�[�}�b�g)		gf_fmtZero(w_str,w_kazu)
'�敪�}�X�^����e��f�[�^���擾		gf_GetKubunName(pDAIBUNRUI,pSYOBUNRUI,pNendo,pKubunName)
'�敪�}�X�^����e��f�[�^���擾(����)gf_GetKubunName_R(pDAIBUNRUI,pSYOBUNRUI,pNendo,pKubunName)
'�S�p�𔼊p�ɕϊ�					gf_Zen2Han(pStr)


'** �萔��` **
'** �L���b�V���I�t **
Response.Expires = 0
Response.AddHeader "Pragma", "No-Cache"

'** �\����` **
'** �ϐ��錾 ** 
'** �O����ۼ��ެ��` **
%>
<!--#include file="adovbs.inc"-->
<!--#include file="CACommon.asp"-->
<!--#include file="common_combo.asp"-->
<!--#include file="com_const.asp"-->
<!--#include file="com_const_web.asp"-->
<%

Const WebRootPath = "C:\Inetpub\wwwroot/cassist"

'**EXCEĻ�ي֘A�̺ݽ�**
Const C_KK_PATH         = "C:\Inetpub\wwwroot"  '�H���Ǘ���ۼު�Ẵp�X


'DB����
Const C_DB_PATH             = ""
Const C_BAK_PATH            = ""        'C_CSV_PATH�Ƒ�����ׂ�
Const C_HOME_DIR            = ""
Const C_ROOT_URL            = ""
Const C_CSV_PATH            = ""        'C_BAK_PATH�Ƒ�����ׂ�
Const C_DB_FILE_NAME        = ""        'MDB�t�@�C����

CONST C_DB_NAME             = ""

'// �װ����
Const C_ERR_DATA_EXIST      = -2147217900       '// ���ᔽ�̏ꍇ�ɔ�������װ����
Const C_ERR_DATA_EXIST2     = -2147467259       '// ���ᔽ�̏ꍇ�ɔ�������װ����
Const C_CommandTimeout      = 600               '// �ڑ����m������܂ő҂b��
Const C_ConnectionTimeout   = 60                '// �ڑ��m���܂ł̑҂�����

Public m_sGrpKey                    '//��ٰ�߷�
Public m_objDB
Public m_sErrMsg

'// ̧�ٵ�޼ު��
Public m_oFile
Const C_CSV_TANTO = "TANTO.CSV"     '// �S���҈ꗗ


Set m_objDB = Server.CreateObject("ADODB.Connection")

'////////////////////////////////////////////////////////////////////////
'// �f�[�^�x�[�X�̃I�[�v��
'//
'// ���@���F
'// �߂�l�F����I��    : 0
'//         �ُ�I��    : -1
'////////////////////////////////////////////////////////////////////////
Function gf_OpenDatabase()

    Dim w_bRetCode              '// Boolean�߂�l
    Dim w_bErrFlg               '// �װ�׸�
    Dim w_bErrMsg
    
    On Error Resume Next
    Err.Clear
'Response.Write "OpenDatabase1"
    gf_OpenDatabase = -1
    w_bErrFlg = True
    
'Response.Write "OpenDatabase2"
    '// �ް��ް��ڑ�(�I���N��) 
    If gf_openORAADODatabase(m_objDB)=False Then
'Response.Write "OpenDatabase3"
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        m_sErrMsg = Err.description & vbCrLf 
        w_bErrFlg = True
    else
'Response.Write "OpenDatabase4"
        '����I��
        gf_OpenDatabase = 0
    End If
    
    Err.Clear
    
End Function

'////////////////////////////////////////////////////////////////////////
'// �n�c�a�b�f�[�^�x�[�X�̃I�[�v��
'//
'// ���@���FOUT dbOpen      : �I�[�v������c�a
'// �߂�l�F����I��    : True
'//         �ُ�I��    : False
'////////////////////////////////////////////////////////////////////////
Function gf_openODBCDatabase(dbOpen)

    On Error Resume Next
    gf_openODBCDatabase = False
    If Err <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        Response.Write "OpenODBCDataBase�֐��O�ɃG���["
    End If
    '// �ް��ް��ڑ�
    Set dbOpen = Server.CreateObject("ADODB.Connection")
    If Err <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        Exit Function
    End If

    dbOpen.CommandTimeout = C_CommandTimeout        '// �ڑ����m������܂ő҂b��
    dbOpen.ConnectionTimeout = C_ConnectionTimeout  '// �ڑ��m���܂ł̑҂�����
    dbOpen.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & server.MapPath (C_HOME_DIR) & C_DB_PATH

    
    If Err <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        Exit Function
    End If
    gf_openODBCDatabase = True
    
End Function

'////////////////////////////////////////////////////////////////////////
'// �I���N���f�[�^�x�[�X�iADO)�̃I�[�v��
'//
'// ���@���FOUT dbOpen      : �I�[�v������c�a
'// �߂�l�F����I��    : True
'//         �ُ�I��    : False
'////////////////////////////////////////////////////////////////////////
Function gf_openORAADODatabase(dbOpen)
    On Error Resume Next

    gf_openORAADODatabase = False
    
    If Err <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        Response.Write "OpenODBCDataBase�֐��O�ɃG���["
    End If
    
    '// �ް��ް��ڑ�
    Set dbOpen = Server.CreateObject("ADODB.Connection")
    If Err <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        Exit Function
    End If
    
    If NOT AutoOpen(dbopen,Session("CONNECT"),Session("USER_ID") & "/" & Session("PASS")) Then
        Exit Function
    End If
    
    If Err <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        Exit Function
    End If
    gf_openORAADODatabase = True
    
End Function


Public Function AutoOpen(m_oConObj,ByVal p_sDbName, ByVal p_sConnect)
'*************************************************************************************
' �@    �\: �f�[�^�x�[�X�ڑ� For Oracle
' ��    �l: True - ���� / False - ���s
' ��    ��: p_sDbName - �f�[�^�x�[�X��
'           p_sConnect - �ڑ������� (���[�U�� "/" �p�X���[�h)
' �@�\�ڍ�:
' ��    �l:
'*************************************************************************************
    On Error Resume Next 
    AutoOpen = False
    Dim w_vSplit            '�I���N���N���X�Ɠ��������ɂ���B
    
    w_vSplit = Split(p_sConnect, "/")

    '/* ADO �źȸ�
    ' 20220601 kiyomoto Edit ---------------------------------------------ST
    'm_oConObj.ConnectionString = "Provider=MSDAORA;"
    m_oConObj.ConnectionString = "Provider=OraOLEDB.Oracle;"
    ' 20220601 kiyomoto Edit ---------------------------------------------ED

    m_oConObj.ConnectionString = m_oConObj.ConnectionString & "Data Source=" & p_sDbName & ";"

    If Err <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        Exit Function
    End If

    m_oConObj.Open , w_vSplit(0), w_vSplit(1)
    If Err <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        Exit Function
    End If
 
   AutoOpen = True

End Function


'////////////////////////////////////////////////////////////////////////
'// �f�[�^�x�[�X�̃N���[�Y
'//
'// ���@���F
'// �߂�l�F
'////////////////////////////////////////////////////////////////////////
Sub gs_CloseDatabase()

    '�ް��ް���۰�ނ���
    gf_closeObject(m_objDB)
    
End Sub

'////////////////////////////////////////////////////////////////////////
'// �I�u�W�F�N�g�̃N���[�Y�i�f�[�^�x�[�X�A���R�[�h�Z�b�g�j
'//
'// ���@���FOUT oClose      : �N���[�Y����I�u�W�F�N�g
'// �߂�l�F����I��    : True
'//         �ُ�I��    : False
'////////////////////////////////////////////////////////////////////////
Function gf_closeObject(oClose)

    On Error Resume Next

   'ADO�֘A
    gf_closeObject = False
    oClose.Close
    set oClose = Nothing

    gf_closeObject = True

    On Error Goto 0
    Err.Clear

End Function

'********************************************************************************
'*  [�@�\]  ������̌������E���ɽ�߰������đ�����
'*  [����]  p_EigyoCD�F�c�Ə��R�[�h
'*  [�ߒl]  �c�Ə��R�[�h
'*  [����]  
'********************************************************************************
Function gf_Keta(p_Str,p_iSpace)

    Dim i
    Dim w_sCd
    
    On Error Resume Next
    Err.Clear

    For i = 0 To p_iSpace - f_LenB(p_Str)
        w_sCd = w_sCd & "&nbsp;"
    Next

    gf_Keta = p_Str & w_sCd
    
End Function

'////////////////////////////////////////////////////////////////////////
'// �G���[���y�[�W�\��
'//
'// ���@���FIN  
'// �߂�l�F����I��    : True
'//         �ُ�I��    : False
'////////////////////////////////////////////////////////////////////////
Sub gs_showMsgPage(p_sWinTitle, p_sMsgTitle, p_sMsgString, p_sRetURL, p_sTarget)
    Dim i

	'// �װү���޺��ް�
	wErrMsg = Replace(p_sMsgString, Chr(13), "\n")
	wErrMsg = Replace(wErrMsg, Chr(10), "\n")

	'===========================================================
	'=[����]  ���ϐ�"URL"���擾���āA
	'=		  ���̒��̃t�H���_�[������װү�������ق��擾����
	'===========================================================
	w_sURL = request.servervariables("URL")
	Call gf_GetErrMsgTitle(w_sURL,w_sMsgTitle)

	'===========================================================
	'=[����]  ���ϐ�"URL"�̒���
	'=		  "login/"�������Ă���Cassist/default.asp�ɂ��ǂ��B
	'=		  �����ĂȂ�������Alogin/top.asp�ɂ��ǂ��B
	'===========================================================
	'// �װ���߂�� & �^�[�Q�b�g�擾
	if InStr(w_sURL,"login/") <> 0 then 
		w_sRetURL= C_RetURL & "default.asp"
	Else
	    w_sRetURL= C_RetURL & C_ERR_RETURL
	End if
	w_sTarget="_top"

	'// �s���A�N�Z�X��
	if gf_IsNull(Session("LOGIN_ID")) then
	    w_sMsgTitle="���O�C���G���["
	    w_sRetURL = C_RetURL & "default.asp"
	End if

	%>
	<HTML>
	<HEAD>
	<TITLE><%=p_sWinTitle%></TITLE>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//  [�@�\]  �y�[�W���[�h������
	//  [����]
	//  [�ߒl]
	//  [����]
	//************************************************************
	function window_onload() {

		window.alert("<%=w_sMsgTitle%>\n\n<%= wErrMsg %>");

		document.frm.target = "<%=w_sTarget%>";
		document.frm.action = "<%=w_sRetURL%>";
		document.frm.submit();

	}
	    
	//-->
	</SCRIPT>
	</HEAD>
	<BODY bgcolor="#ffffff" LANGUAGE=javascript onload="return window_onload()">
	<FORM method="POST" name="frm">
	<!--
	<FORM method="POST" name="frm" action="<%=p_sRetURL%>" target="<%=p_sTarget %>">
	<TABLE border="0" cellpadding="0" cellspacing="0" width="100%">
	    <TR>
	        <TD nowrap align="center" valign="middle" height="20">
	            <FONT size="4" color="#cc0000" face="�l�r �S�V�b�N">
	            <B><%=p_sMsgTitle%></B></FONT>
	        </TD>
	    </TR>
	    <TR>
	        <TD nowrap align="center" valign="middle" height="10"></TD>
	    </TR>
	    <TR>
	        <TD nowrap align="center" valign="middle"><FONT size="2" face="�l�r �S�V�b�N">
	        <%
	            For i = 1 To Len(p_sMsgString)
	                If Mid(p_sMsgString, i, 1) <> Chr(10) Then
	                    Response.Write Mid(p_sMsgString, i, 1)
	                Else
	                    Response.Write "<BR>"
	                End If
	            Next
	        %>
	        </FONT></TD>
	    </TR>
	    <TR>
	        <TD nowrap align="center" valign="middle" height="10"></TD>
	    </TR>
	    <TR>
	        <TD nowrap align="center" valign="middle" height="20">
				<INPUT type="submit" value="�߁@��" name="Back" tabindex="15">
	        </TD>
	    </TR>
	</TABLE>
	//-->
	</FORM>
	</BODY>
	</HTML>

<%
End Sub

'********************************************************************************
'*  [�@�\]  �װү�������ق��擾����
'*  [����]  p_sURL      : ���ϐ�"URL"
'*  [�ߒl]  p_sMsgTitle : �װү��������
'*  [����]  
'********************************************************************************
Function gf_GetErrMsgTitle(p_sURL,p_sMsgTitle)
    
    On Error Resume Next
    Err.Clear

	if InStr(p_sURL,"kks0110/") <> 0 then p_sMsgTitle = "���Əo������"
	if InStr(p_sURL,"kks0140/") <> 0 then p_sMsgTitle = "�s���o������"
	if InStr(p_sURL,"kks0170/") <> 0 then p_sMsgTitle = "�����o������"
	if InStr(p_sURL,"skn0130/") <> 0 then p_sMsgTitle = "�������{�Ȗړo�^"
	if InStr(p_sURL,"skn0120/") <> 0 then p_sMsgTitle = "�����ēƏ��\���o�^"
	if InStr(p_sURL,"sei0100/") <> 0 then p_sMsgTitle = "���ѓo�^"
	if InStr(p_sURL,"sei0200/") <> 0 then p_sMsgTitle = "���шꗗ"
	if InStr(p_sURL,"sei0300/") <> 0 then p_sMsgTitle = "�l���шꗗ"
	if InStr(p_sURL,"skn0170/") <> 0 then p_sMsgTitle = "�������Ԋ�(�N���X�ʁj"
	if InStr(p_sURL,"skn0180/") <> 0 then p_sMsgTitle = "�������ԋ����\��ꗗ"
	if InStr(p_sURL,"han0121/") <> 0 then p_sMsgTitle = "���N�Y���҈ꗗ"
	if InStr(p_sURL,"gyo0200/") <> 0 then p_sMsgTitle = "�s�������ꗗ"
	if InStr(p_sURL,"jik0210/") <> 0 then p_sMsgTitle = "�N���X�ʎ��Ǝ��Ԉꗗ"
	if InStr(p_sURL,"jik0200/") <> 0 then p_sMsgTitle = "�����ʎ��Ǝ��Ԉꗗ"
	if InStr(p_sURL,"web0310/") <> 0 then p_sMsgTitle = "���Ԋ������A��"
	if InStr(p_sURL,"mst0144/") <> 0 then p_sMsgTitle = "�i�H����o�^"
	if InStr(p_sURL,"web0320/") <> 0 then p_sMsgTitle = "�g�p���ȏ��o�^"
	if InStr(p_sURL,"gak0460/") <> 0 then p_sMsgTitle = "�w���v�^�������o�^"
	if InStr(p_sURL,"gak0461/") <> 0 then p_sMsgTitle = "�������������o�^"
	if InStr(p_sURL,"gak0470/") <> 0 then p_sMsgTitle = "�e��ψ��o�^"
	if InStr(p_sURL,"web0340/") <> 0 then p_sMsgTitle = "�l���C�I���Ȗڌ���"
	if InStr(p_sURL,"gak0300/") <> 0 then p_sMsgTitle = "�w����񌟍�"
	if InStr(p_sURL,"mst0113/") <> 0 then p_sMsgTitle = "���w�Z��񌟍�"
	if InStr(p_sURL,"mst0123/") <> 0 then p_sMsgTitle = "�����w�Z��񌟍�"
	if InStr(p_sURL,"mst0133/") <> 0 then p_sMsgTitle = "�i�H���񌟍�"
	if InStr(p_sURL,"web0300/") <> 0 then p_sMsgTitle = "���ʋ����\��"
	if InStr(p_sURL,"web0330/") <> 0 then p_sMsgTitle = "�A���f����"
	if InStr(p_sURL,"web0350/") <> 0 then p_sMsgTitle = "�󂫎��ԏ�񌟍�"
	if InStr(p_sURL,"web0360/") <> 0 then p_sMsgTitle = "�����������ꗗ"
	if InStr(p_sURL,"sei0400/") <> 0 then p_sMsgTitle = "���і������o�^"
	if InStr(p_sURL,"sei0500/") <> 0 then p_sMsgTitle = "���͎������ѓo�^"

End Function

'********************************************************************************
'*  [�@�\]  ��ݻ޸��݊J�n
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub gs_BeginTrans()

   On Error Resume Next
   m_objDB.BeginTrans
   On Error Goto 0
   Err.Clear
   
End Sub

'********************************************************************************
'*  [�@�\]  ��ݻ޸��݂�۰��ޯ�����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub gs_RollbackTrans()

   On Error Resume Next
   
   m_objDB.RollbackTrans
   
   On Error Goto 0
   Err.Clear
   
End Sub

'********************************************************************************
'*  [�@�\]  ��ݻ޸��݂�ЯĂ���
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub gs_CommitTrans()

   On Error Resume Next
   
   m_objDB.CommitTrans
   
   On Error Goto 0
   Err.Clear
   
End Sub

'********************************************************************************
'*  [�@�\]  ں��޾�Ă��擾����
'*  [����]  p_Rs            :�擾����ں��޾��
'*          p_sSQL          :ں��޾�Ď擾�̂��߂�SQL
'*  [�ߒl]  0:����I���A���̑�:�װ
'*  [����]  
'********************************************************************************
Function gf_GetRecordset(p_Rs, p_sSQL)


    On Error Resume Next
    Err.Clear
    '����I����ݒ�
    gf_GetRecordset = 0
    Do
        'ں��޾�Ă̎擾
        set p_Rs = m_objDB.Execute(p_sSQL)
'	Set p_Rs = Server.CreateObject("ADODB.Recordset")
'	p_Rs.Open p_sSQL,m_objDB,adOpenStatic
        If Err <> 0 Then
            Call gs_SetErrMsg("gf_GetRecordset:" & Replace(Err.description,vbCrLf," "))
'response.write Err.description
            gf_GetRecordset = Err.number
            Exit Do
        End If
        
        '����I��
        Exit Do
    Loop
    
    Err.Clear
    
End Function

'********************************************************************************
'*  [�@�\]  ں��޾�Ă��擾����
'*  [����]  p_Rs            :�擾����ں��޾��
'*          p_sSQL          :ں��޾�Ď擾�̂��߂�SQL
'*          p_iPageSize     :�߰�޻���
'*  [�ߒl]  0:����I���A���̑�:�װ
'*  [����]  
'********************************************************************************
Function gf_GetRecordsetExt(p_Rs, p_sSQL, p_iPageSize)

    On Error Resume Next
    Err.Clear
    
    '����I����ݒ�
    gf_GetRecordsetExt = 0

    Do
        'ں��޾�Ă̎擾
        p_Rs.ActiveConnection = m_objDB
        p_Rs.CursorType = adOpenKeyset
        'p_Rs.CursorType = adOpenForwardOnly
        p_Rs.Source = p_sSQL
        p_Rs.LockType = adLockOptimistic
        p_Rs.Pagesize = p_iPageSize
        p_Rs.Open
        If Err <> 0 Then
            Call gs_SetErrMsg("gf_GetRecordsetExt:" & Replace(Err.description,vbCrLf," "))
            gf_GetRecordsetExt = Err.number
'response.write Err.description
            Exit Do
        End If
        
        '����I��
        Exit Do
    Loop
    
    Err.Clear
    
End Function

'*****************************************************
'*	[�T�v]	�f�[�^�̒��o�i���R�[�h�Z�b�g���擾�j���o�p
'*	[����]	p_sSQL		= SQL
'*	[�ߒl]	p_Rs		= ���R�[�h�Z�b�g
'*  [����]	p_Rs.MovePrevious�g�p��
'*@***************************************************
Function gf_GetRecordset_OpenStatic(p_Rs,p_sSQL)

	gf_GetRecordset_OpenStatic=0
	
	On Error Resume Next
	Err.Clear 

	Set p_Rs = Server.CreateObject("ADODB.Recordset")
	p_Rs.Open p_sSQL,m_objDB,adOpenStatic

	If Err.number <>0 Then
		Call gs_SetErrMsg("gf_GetRecordset_OpenStatic:" & Replace(Err.description,vbCrLf," "))
		gf_GetRecordset_OpenStatic = Err.number
		Exit Function
	End If

End Function

'********************************************************************************
'*  [�@�\]  �y�[�W���̌v�Z
'*  [����]  p_Rs            :�擾����ں��޾��
'*          p_sField        :�P�y�[�W���Ƃ̕\������
'*  [�ߒl]  �y�[�W��
'*  [����]  �i�h���R�[�h�Z�b�g.PageCount�h���g�p�ł��Ȃ��ꍇ�̂ݎg�p���Ă��������j
'********************************************************************************
Function gf_PageCount(p_Rs, p_iDsp)

    dim w_iRecCount
    On Error Resume Next
    Err.Clear
    gf_PageCount=0

    w_iRecCount=0
    p_Rs.MoveFirst
    Do Until p_Rs.EOF
        p_Rs.MoveNext
        w_iRecCount=w_iRecCount+1
    Loop
    p_Rs.MoveFirst

    gf_PageCount = INT(w_iRecCount/m_iDsp) + gf_IIF(w_iRecCount mod m_iDsp = 0,0,1)
    'gf_PageCount = INT(w_iRecCount/m_iDsp) + gf_IIF(m_Rs.RecordCount mod m_iDsp = 0,0,1)

    Err.Clear
   
End Function

'********************************************************************************
'*  [�@�\]  ں��ރJ�E���g�擾
'*  [����]  p_Rs
'*  [�ߒl]  gf_GetRsCount:ں��ސ�
'*  [����]  p_Rs.RecordCount���g���Ȃ��ꍇ
'********************************************************************************
Function gf_GetRsCount(p_Rs)
Dim w_iRecCount

    On Error Resume Next
    Err.Clear

    w_iRecCount= 0

    If p_Rs.EOF = False Then
        p_Rs.MoveFirst
        Do Until p_Rs.EOF
            p_Rs.MoveNext
            w_iRecCount=w_iRecCount+1
        Loop
        p_Rs.MoveFirst
    End If

    gf_GetRsCount = w_iRecCount
    Err.Clear

End Function

'********************************************************************************
'*  [�@�\]  ��Βl�y�[�W�̐ݒ�
'*  [����]  p_Rs            :�擾����ں��޾��
'*          p_iPage         :�w�肵�����y�[�W
'*          p_sField        :�P�y�[�W���Ƃ̕\������
'*  [�ߒl]  �Ȃ�
'*  [����]  �i�h���R�[�h�Z�b�g.AbsolutePage�h���g�p�ł��Ȃ��ꍇ�̂ݎg�p���Ă��������j
'********************************************************************************
Sub gs_AbsolutePage(p_Rs,p_iPage, p_iDsp)
    dim w_iRecCount
    On Error Resume Next
    Err.Clear

    '��Βl�y�[�W�̐ݒ�
    p_Rs.MoveFirst
    for w_iRecCount=1 to p_iDsp*(p_iPage-1)
        p_Rs.MoveNext
    Next    

    Err.Clear
    
End Sub

'********************************************************************************
'*  [�@�\]  �󂾂�����S�p�X�y�[�X��Ԃ�
'*  [����]  p_Data          :�\���������f�[�^
'*  [�ߒl]  �ϊ�������
'*  [����]  
'********************************************************************************
Function gf_HTMLTableSTR(p_Data)
    On Error Resume Next
    Err.Clear
    
    gf_HTMLTableSTR=gf_IIF(ISNULL(p_Data) OR p_DATA="","�@",p_DATA)

    Err.Clear
    
End Function



'********************************************************************************
'*  [�@�\]  �t�B�[���h�̒l�擾
'*  [����]  p_Rs            :�擾����ں��޾��
'*          p_sField        :field name
'*  [�ߒl]  �擾������
'*  [����]  
'********************************************************************************
Function gf_GetFieldValue(p_Rs, p_sField)

    On Error Resume Next
    Err.Clear
    
    '����I����ݒ�
    gf_GetFieldValue = p_Rs(p_sField)

    Err.Clear
    
End Function


'********************************************************************************
'*  [�@�\]  SQL�����s����
'*  [����]  p_sSQL          :ں��޾�Ď擾�̂��߂�SQL
'*  [�ߒl]  0:����I���A���̑�:�װ
'*  [����]  
'********************************************************************************
Function gf_ExecuteSQL(p_sSQL)
    On Error Resume Next
    Err.Clear
    
    '����I����ݒ�
    gf_ExecuteSQL = 0

    Do
        'ں��޾�Ă̎擾
        set p_Rs = m_objDB.Execute(p_sSQL)
        If Err <> 0 Then
            Call gs_SetErrMsg("gf_ExecuteSQL:" & Replace(Err.description,vbCrLf," "))
            gf_ExecuteSQL = Err.number
            Exit Do
        End If
        
        '����I��
        Exit Do
    Loop
    
    Err.Clear
    
End Function

'********************************************************************************
'*  [�@�\]  �װү���ނ�ݒ肷��
'*  [����]  p_sErrMsg       :�װү����
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub gs_SetErrMsg(p_sErrMsg)

    m_sErrMsg = p_sErrMsg

End Sub

'********************************************************************************
'*  [�@�\]  �װү���ނ��擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �װү����
'*  [����]  
'********************************************************************************
Function gf_GetErrMsg()

    gf_GetErrMsg = m_sErrMsg

End Function

'********************************************************************************
'*  [�@�\]  �װү���ނ�ر����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub gs_ClearErrMsg()

    m_sErrMsg = ""

End Sub

'********************************************************************************
'*  [�@�\]  ÷��̧�ق�Open
'*  [����]  p_File  �F̧�ٵ�޼ު��
'*  �@�@�@  p_sPath �F�߽
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function gf_OpenFile(p_File,p_sPath)

    On Error Resume Next
    gf_OpenFile = False

    Set m_oFile = Server.CreateObject("Scripting.FileSystemObject")
    If Err <> 0 Then
        '̧�ٵ�޼ު�Đ������s
        Call gs_SetErrMsg("gf_OpenFile:" & Replace(Err.description,vbCrLf," "))
        Exit Function
    End If

    Set p_File = m_oFile.CreateTextFile(p_sPath,True,False) 
    If Err <> 0 Then
        '̧�ٵ���ݎ��s
        Call gs_SetErrMsg("gf_OpenFile:" & Replace(Err.description,vbCrLf," "))
        Exit Function
    End If

    gf_OpenFile = True
    
End Function


'********************************************************************************
'*  [�@�\]  NULL���󔒂��n�C�t���ɕϊ�����
'*  [����]  �m�F���镶����
'*  [�ߒl]  NULL����:�n�C�t���A���̑�:���̂܂�
'*  [����]  
'********************************************************************************
Function gf_SetNull2Haifun(p_sStr)

    If isNull(p_sStr) Then
        gf_SetNull2Haifun = "-"
    Else
        If Trim(p_sStr) = "" Then
            gf_SetNull2Haifun = "-"
        Else
            gf_SetNull2Haifun = p_sStr
        End If          
    End If

End Function

'********************************************************************************
'*  [�@�\]  NULL���󕶎��ɕϊ�
'*  [����]  p_sStr      :������
'*  [�ߒl]  �ϊ���̕�����
'*  [����]  �w�肳�ꂽ������NULL�̏ꍇ�󕶎���ݒ肵�ANULL�łȂ��ꍇ���̂܂܂�ݒ�
'********************************************************************************
Function gf_SetNull2String(p_sStr)

    If IsNull(p_sStr) Then
        gf_SetNull2String = ""
    Else
        gf_SetNull2String = p_sStr
    End If

End Function

'********************************************************************************
'*  [�@�\]  NULL��0�ɕϊ�
'*  [����]  p_iNum      :���l
'*  [�ߒl]  �ϊ���̐��l
'*  [����]  �w�肳�ꂽ������NULL�̏ꍇ0��ݒ肵�ANULL�łȂ��ꍇ���̂܂܂�ݒ�
'********************************************************************************
Function gf_SetNull2Zero(p_iNum)

    If IsNull(p_iNum) or p_iNum="" Then
        gf_SetNull2Zero = 0
    Else
        gf_SetNull2Zero = p_iNum
    End If

End Function

'************************************
'*  null�`�F�b�N�iIsNull���ǂ��j
'*  [����]
'*        ����Ώ�
'*  [�߂�l]
'*        true, false
'************************************
function gf_IsNull(p_null)
	w_jud = false
	if IsNull(p_null) = true then
		w_jud = true
	elseif IsEmpty(p_null) = true then
		w_jud = true
	elseif p_null = "" then
		w_jud = true
	end if
	gf_IsNull = w_jud
end function


'********************************************************************************
'*  [�@�\]  ���t�������YYYY/MM/DD��̫�ϯ�
'*  [����]  p_sDate     :���t������
'*          p_sDelimit  :��؂蕶��
'*  [�ߒl]  YYYY/MM/DD������i�ׯ�������͋�؂蕶���ɂ��j
'*  [����]  ���t�łȂ��ꍇ�A�N���S���ŗ^�����Ă��Ȃ��ꍇ�͂��̂܂܂�Ԃ�
'********************************************************************************
Function gf_YYYY_MM_DD(p_sDate,p_sDelimit)

    Dim w_sStr
    
    gf_YYYY_MM_DD = ""
    
    If IsNull(p_sDate)  Then
        gf_YYYY_MM_DD = ""
        Exit Function
    End If
    
    '�󔒁A���t����Ȃ��A�N���S���ł͂Ȃ��ꍇ�͋󔒂�Ԃ�
    If p_sDate = "" Then 
        gf_YYYY_MM_DD = p_sDate
        Exit Function
    End If
    If IsDate(CDate(p_sDate)) <> True Then
        gf_YYYY_MM_DD = p_sDate
        Exit Function
    End If
    
    w_sMM = DatePart("m",CDate(p_sDate))
    w_sDD = DatePart("d",CDate(p_sDate))
    
    If Len(w_sMM) = 1 Then w_sMM = "0" & w_sMM
    If Len(w_sDD) = 1 Then w_sDD = "0" & w_sDD
	
    w_sStr = DatePart("yyyy",CDate(p_sDate)) & p_sDelimit
    w_sStr = w_sStr & w_sMM & p_sDelimit
    w_sStr = w_sStr & w_sDD
    
    gf_YYYY_MM_DD = w_sStr
    
End Function

'****************************************************
'[�@�\]	�a��t�H�[�}�b�g	:MM��DD���i�j���j
'[����]	pDate : �Ώۓ��t(YYYY/MM/DD)
'[�ߒl]	
'****************************************************
Function gf_fmtWareki(pDate)

	gf_fmtWareki = ""

	'// Null�Ȃ甲����
	if gf_IsNull(trim(pDate)) then	Exit Function

	'// MM��DD���쐬
	w_MM = Mid(FormatYYYYMMDD(pDate),6,2) & "��"
	w_DD = Right(FormatYYYYMMDD(pDate),2) & "��"

	'// �j�����擾
	w_Youbi = WeekdayName(Weekday(FormatYYYYMMDD(pDate))) & "<BR>"
	w_Youbi = "�i" & Left(w_Youbi,1) & "�j"

	gf_fmtWareki = w_MM & w_DD & w_Youbi

End Function

'********************************************************************************
'*  [�@�\]  �N�������`
'*  [����]  ���t
'*  [�ߒl]  ��؂蕶���t�����t(�G���[���A���������̂܂�)
'*  [����]  �N������YYYY/MM/DD�̌`�ɕϊ��B�����݂̂̔N��������؂蕶���ŕ�����B
'********************************************************************************
Function FormatYYYYMMDD( target )
	dim yyyy, mm, dd

	yyyy = year(target)
	mm = month(target)
	dd = day(target)

	FormatYYYYMMDD = yyyy & "/" & right( "00" & mm, 2 ) & "/" & right( "00" & dd, 2)

End Function

'********************************************************************************
'*  [�@�\]  �N�������`
'*  [����]  �����̂ݓ��t(YYYYMMDD�`���̂��̂̂�)
'*  [�ߒl]  ��؂蕶���t�����t(�G���[���A���������̂܂�)
'*  [����]  �����݂̂̔N��������؂蕶���ŕ�����B
'*          gf_YYYY_MM_DD��YYYYMMDD���󂯕t���Ȃ��悤�Ȃ̂ŁB
'********************************************************************************
Function gf_FormatDate(p_Date,p_Delimiter)
    Dim w_sDate 
    Dim w_sYear
    Dim w_sMonth
    Dim w_sDay

    '�󔒂Ȃ�G���[
    If IsNull(p_Date)  Then
        gf_FormatDate = p_Date
        Exit Function
    End If
    If p_Date = "" Then 
        gf_FormatDate = p_Date
        Exit Function
    End If

    '�����łȂ��Ȃ�G���[
    If Not IsNumeric(p_Date) Then 
        gf_FormatDate = p_Date
        Exit Function
    End If

    '8���łȂ��Ȃ�G���[
    If Len(p_Date) <> 8 Then
        gf_FormatDate = p_Date
        Exit Function
    End If

    w_sYear  = Mid(p_Date,1,4)
    w_sMonth = Mid(p_Date,5,2)
    w_sDay   = Mid(p_Date,7,2)

    w_sDate = w_sYear & p_Delimiter 
    w_sDate = w_sDate & w_sMonth & p_Delimiter
    w_sDate = w_sDate & w_sDay

    '�ŏI�I�ɓ��t�łȂ��Ȃ�G���[
    If Not IsDate(w_sDate) Then 
        gf_FormatDate = p_Date
        Exit Function
    End If

    gf_FormatDate = w_sDate

End Function

'*********************************************************************
'�@�@���t�^�t�H�[�}�b�g�֐��@�@�@�@�@�@�@�@�@�@�@�@�@ver 1.0  00.10.19
'
'�@�@����(1)�F[Date]   �t�H�[�}�b�g���������t�^
'�@�@�@�@(2)�F[String] �t�H�[�}�b�g�^�i�y�[�W����ɋL�ځj
'�@�@�ߒl   �F[String] �t�H�[�}�b�g���ꂽ������
'*********************************************************************

Function FormatTime(datTime,strFormat)

	Dim tmpFormat
	Dim cntType
	Dim FormatType

	FormatType = Split("YYYY/YY/MM/M/DD/D/HH24/H24/HH/H/II/I/SS/S/XX/ZZ","/")

	tmpFormat = Cstr(strFormat)

	For cntType = 0 To Ubound(FormatType)

		If InStr(tmpFormat,FormatType(cntType)) > 0 Then

			Select Case FormatType(cntType)
			Case "HH24"
				tmpFormat = Replace(tmpFormat,"HH24",Right(CStr(Hour(datTime) + 100),2))
			Case "H24"
				tmpFormat = Replace(tmpFormat,"H24",CStr(Hour(datTime)))
			Case "HH"
				tmpFormat = Replace(tmpFormat,"HH",Right(CStr((Hour(datTime) Mod 12) + 100),2))
			Case "H"
				tmpFormat = Replace(tmpFormat,"H",CStr(Hour(datTime) Mod 12))		
			Case "II"
				tmpFormat = Replace(tmpFormat,"II",Right(CStr(Minute(datTime) + 100),2))
			Case "I"
				tmpFormat = Replace(tmpFormat,"I",CStr(Minute(datTime)))
			Case "SS"
				tmpFormat = Replace(tmpFormat,"SS",Right(CStr(Second(datTime) + 100),2))
			Case "S"
				tmpFormat = Replace(tmpFormat,"S", CStr(Second(datTime)))
			Case "YYYY"
				If Len(CStr(Year(datTime))) = 2 Then
					If Year(datTime) > 30 Then
						tmpFormat = Replace(tmpFormat,"YYYY","19" & CStr(Year(datTime)))
					Else
						tmpFormat = Replace(tmpFormat,"YYYY","20" & CStr(Year(datTime)))
					End If
				Else
					tmpFormat = Replace(tmpFormat,"YYYY",CStr(Year(datTime)))
				End If
			Case "YY"
				tmpFormat = Replace(tmpFormat,"YY",Right(CStr(Year(datTime)),2))
			Case "MM"
				tmpFormat = Replace(tmpFormat,"MM",Right(CStr(Month(datTime) + 100),2))
			Case "M"
				tmpFormat = Replace(tmpFormat,"M",CStr(Month(datTime)))
			Case "DD"
				tmpFormat = Replace(tmpFormat,"DD",Right(CStr(Day(datTime) + 100),2))
			Case "D"
				tmpFormat = Replace(tmpFormat,"D",CStr(Day(datTime)))
			Case "XX"
				If Hour(datTime) < 12 Then
					tmpFormat = Replace(tmpFormat,"XX","�ߑO")
				Else
					tmpFormat = Replace(tmpFormat,"XX","�ߌ�")
				End If
			Case "ZZ"
				If Hour(datTime) < 12 Then
					tmpFormat = Replace(tmpFormat,"ZZ","AM")
				Else
					tmpFormat = Replace(tmpFormat,"ZZ","PM")
				End If
			End Select
		
		End If

	Next

	FormatTime = CStr(tmpFormat)

End Function

'*********************************************************************
'�@�t�H�[�}�b�g�w��ł���^�ɂ��āi���t�^����̕ϊ��j
'�@�@YYYY	����S��
'�@�@YY		����Q��
'�@�@MM		���Q��
'�@�@M		���P��
'�@�@DD		���Q��
'�@�@D		���P��
'�@�@HH24	���Q���i�Q�S���ԁj
'�@�@H24	���P���i�Q�S���ԁj
'�@�@HH		���Q���i�P�Q���ԁj
'�@�@H		���P���i�P�Q���ԁj
'�@�@II		���Q��
'�@�@I		���P��
'�@�@SS		�b�Q��
'�@�@S		�b�P��
'�@�@XX		�ߑO/�ߌ�
'�@�@ZZ		AM/PM
'*********************************************************************


'********************************************************************
'*  ������̃o�C�g����Ԃ�
'*  [����]
'*          p_str   :   ���ׂ镶����
'*  [�߂�l] 
'*          ������̃o�C�g��
'********************************************************************
Function f_LenB(p_str)

    Dim w_sbyte, w_dbyte, w_len, w_idx

    w_len = Len(p_str & "")

    For w_idx = 1 To w_len
        If Len(Hex(Asc(Mid(p_str, w_idx, 1)))) > 2 Then
            w_dbyte = w_dbyte + 1
        End If
    Next

    w_sbyte = w_len - w_dbyte

    f_LenB = w_sbyte + (w_dbyte * 2)

End Function

function gf_IIF(p_Judge, p_tStr, p_fStr)
'************************************
'*  VB��IIF�֐���ASP�Ŏ���
'*  [����]
'*        VB��IIF�ɓ���
'*  [�߂�l]
'*        VB��IIF�ɓ���
'************************************
    if p_Judge = true then
        gf_iif = p_tStr
    else
        gf_iif = p_fStr
    end if
end function

'***************************************************
'*  Format�֐� ���̫�ϯ�
'*  ����:
'*      �Ώۂ̐��l:w_str
'*      ����:w_kazu
'*      ��)fmtZero(125,7) ----> 0000125
'*          
'***************************************************
Function gf_fmtZero(w_str,w_kazu)
    gf_fmtZero = Right((String(w_kazu,"0") & w_str),w_kazu)
End Function

'********************************************************************
'*  �s���A�N�Z�X��h��
'*  [����]
'*      p_LoginURL : ���O�C�����url
'*
'********************************************************************
Function gf_UseValidRoute(p_LoginURL)

    'հ�ް��������ݏ�������ĂȂ��ꍇ۸޲݉�ʂ�
    If Len(Session("LOGIN_ID")) = 0 Then
        Response.Redirect(p_LoginURL)
    End If

End Function


'********************************************************************************
'*  [�@�\]  �敪�}�X�^����e��f�[�^���擾
'*  [����]  pDAIBUNRUI	= �啪��CD
'*			pSYOBUNRUI  = ������CD
'*			pNendo		= �N�x
'*
'*  [�ߒl]  True:����I��	False:�G���[�i�Y���Ȃ��j
'*			pKubunName  = �擾�����l
'*  [����]  
'********************************************************************************
Function gf_GetKubunName(pDAIBUNRUI,pSYOBUNRUI,pNendo,pKubunName)
	Dim w_iRet
	Dim w_sSQL
	Dim wKubunRs

	On Error Resume Next
	Err.Clear

	gf_GetKubunName = False

	'// �����ނɒl�������ĂȂ������甲����
	if gf_IsNull(pSYOBUNRUI) then
		pKubunName = ""
		gf_GetKubunName = True
		Exit Function
	End if

	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	A.M01_SYOBUNRUIMEI "
	w_sSql = w_sSql & " FROM  "
	w_sSql = w_sSql & " 	M01_KUBUN A "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	 A.M01_NENDO = " & pNendo
	w_sSql = w_sSql & "  AND A.M01_DAIBUNRUI_CD = " & pDAIBUNRUI
	w_sSql = w_sSql & "  AND A.M01_SYOBUNRUI_CD = " & pSYOBUNRUI

	iRet = gf_GetRecordset(wKubunRs, w_sSql)
	If iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit Function
	End If

	if wKubunRs.Eof then
		gf_GetKubunName = True
		Exit Function
	End if

	pKubunName = wKubunRs("M01_SYOBUNRUIMEI")

    If Not IsNull(wKubunRs) Then gf_closeObject(wKubunRs)

	'//����I��
	gf_GetKubunName = True

End Function

'********************************************************************************
'*  [�@�\]  �敪�}�X�^����e��f�[�^���擾(���̂��擾)
'*  [����]  pDAIBUNRUI	= �啪��CD
'*			pSYOBUNRUI  = ������CD
'*			pNendo		= �N�x
'*
'*  [�ߒl]  True:����I��	False:�G���[�i�Y���Ȃ��j
'*			pKubunName  = �擾�����l
'*  [����]  
'********************************************************************************
Function gf_GetKubunName_R(pDAIBUNRUI,pSYOBUNRUI,pNendo,pKubunName)
	Dim w_iRet
	Dim w_sSQL
	Dim wKubunRs

	On Error Resume Next
	Err.Clear

	gf_GetKubunName_R = False

	'// �����ނɒl�������ĂȂ������甲����
	if gf_IsNull(pSYOBUNRUI) then
		pKubunName = ""
		gf_GetKubunName_R = True
		Exit Function
	End if

	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	A.M01_SYOBUNRUIMEI_R "
	w_sSql = w_sSql & " FROM  "
	w_sSql = w_sSql & " 	M01_KUBUN A "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	 A.M01_NENDO = " & pNendo
	w_sSql = w_sSql & "  AND A.M01_DAIBUNRUI_CD = " & pDAIBUNRUI
	w_sSql = w_sSql & "  AND A.M01_SYOBUNRUI_CD = " & pSYOBUNRUI

	iRet = gf_GetRecordset(wKubunRs, w_sSql)
	If iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit Function
	End If

	if wKubunRs.Eof then
		gf_GetKubunName_R = True
		Exit Function
	End if

	pKubunName = wKubunRs("M01_SYOBUNRUIMEI_R")

    If Not IsNull(wKubunRs) Then gf_closeObject(wKubunRs)

	'//����I��
	gf_GetKubunName_R = True

End Function

'********************************************************************************
'*  [�@�\]  �S�p�𔼊p��
'*  [����]  pStr = �ϊ�������������
'*  [�ߒl]  �Ȃ� 
'*  [����]  
'********************************************************************************
function gf_Zen2Han(pStr)

	zenStr = "�O,�P,�Q,�R,�S,�T,�U,�V,�W,�X,"
	zenStr = zenStr & "�A,�C,�E,�G,�I,�J,�L,�N,�P,�R,�T,�V,�X,�Z,�\,�^,�`,�c,�e,�g,�i,�j,�k,�l,�m,�n,�q,�t,�w,�z,�},�~,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
	zenStr = zenStr & "�K,�M,�O,�Q,�S,�U,�W,�Y,�[,�],�_,�a,�d,�f,�h,�o,�r,�u,�x,�{,�p,�s,�v,�y,�|,��,"
	zenStr = zenStr & "�@,�B,�D,�F,�H,�b,�[,�|,�@"
	hanStr = "0,1,2,3,4,5,6,7,8,9,"
	hanStr = hanStr & "�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,"
	hanStr = hanStr & "��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
	hanStr = hanStr & "�,�,�,�,�,�,�,-, "
	wZen = split(zenStr,",")	
	wHan = split(hanStr,",")

	wStr = pStr
	wLen = len(wStr)
	for pf_iCnt = 0 to 89
		'response.write pf_iCnt & "---" & wZen(pf_iCnt) & "----" & wHan(pf_iCnt) & "<br>"
		bChg = false
		while not bChg
			wCnt = instr(1,wStr,wZen(pf_iCnt))
			if wCnt <> 0 then
				'response.write "wCnt=" & wCnt & "  wLen=" & wLen & "   wLen-wCnt=" & wLen-wCnt & "   wStr=" & wStr & "<br>" 
				if len(wHan(pf_iCnt)) = 2 then
					wLen = wLen + 1
					wStr = left(wStr,wCnt-1) & wHan(pf_iCnt) & right(wStr,wLen-wCnt-1)
				else 
					wStr = left(wStr,wCnt-1) & wHan(pf_iCnt) & right(wStr,wLen-wCnt)
				end if
			else 
				bChg = true
			end if
		wend
	next
	gf_Zen2Han = wStr
end function

%>