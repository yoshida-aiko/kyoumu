<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���O�C���I�������
' ��۸���ID : login/header.asp
' �@      �\: ���O�C���I�����̃w�b�_�[���
'-------------------------------------------------------------------------
' ��      ��    
'               
'           
' ��      ��
' ��      �n
'           
'           
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 
' ��      �X: 2001/07/26    ���`�i�K
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

    Dim m_LoginDay      '// ۸޲ݓ�
    Dim m_WaNengappi    '// �a��N����
    Dim m_SchoolName    '// �w�Z��
	Dim m_bErrFlg

'///////////////////////////���C������/////////////////////////////

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

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�w�b�_�[�f�[�^"
    w_sMsg=""
    w_sRetURL="../default.asp"
    w_sTarget="_parent"

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '//�\���f�[�^�擾
        if Not f_GetViewRs() then Exit Do

        '//�����\��
        Call showPage()

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// �I������
    Call gs_CloseDatabase()

End Sub


'********************************************************************************
'*  [�@�\]  �\���f�[�^���擾
'*  [����]  
'*          
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetViewRs()
    Dim w_sSql
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_GetViewRs = False

    '// ���̐��N����
    w_NowDay = gf_YYYY_MM_DD(date,"/")

    '// �����擾
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "     M09_GENGOMEI, "
    w_sSQL = w_sSQL & "     M09_KAISIBI "
    w_sSQL = w_sSQL & "FROM M09_GENGO "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "     M09_KAISIBI <= '" & w_NowDay & "' "
    w_sSQL = w_sSQL & "Order By M09_KAISIBI desc"

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
    'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

    if w_Rs.Eof then
        m_bErrFlg = True
        m_sErrMsg = "�������擾���ɃG���[���ł܂���"
        Exit Function
    End if

    '// ������
'    w_GENGOMEI = w_Rs("M09_GENGOMEI")

    '// �N�i�a��j
'    w_iDiff = DateDiff("yyyy", w_Rs("M09_KAISIBI"), w_NowDay)
'    w_iWaNen = w_iDiff + 1

    '// �w�����擾
    if Not gf_GetKubunName(C_GAKKI,Session("GAKKI"),Session("NENDO"),w_GAKKI) then Exit Function

    '// ����N����
    m_WaNengappi = Session("NENDO") & "�N�x�@" & w_GAKKI

	'// �j��
	w_Youbi = left(WeekDayName( Weekday(w_NowDay) ),1)

    '// ���t��\��
    m_LoginDay = year(date) & "�N�@" & Month(date) & "��" & Day(date) & "��" & "(" & w_Youbi & ")"

    '// �w�Z���擾
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "     M19_NAME "
    w_sSQL = w_sSQL & "FROM M19_GAKKO "
    'w_sSQL = w_sSQL & "Where "
    'w_sSQL = w_sSQL & "     M19_NO = " & C_School_CD

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
    'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

    if w_Rs.Eof then
        m_bErrFlg = True
        m_sErrMsg = "�w�Z���擾���ɃG���[���ł܂���"
        Exit Function
    End if

    '// �w�Z��
    m_SchoolName = w_Rs("M19_NAME")

    f_GetViewRs = True

    call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear

    %>

    <html>
    <head>
    <title>header</title>

    <STYLE TYPE="text/css">
    <!--
        body,tr,td { font-size:11pt;  color:#ffffff;
        font-family: "�l�r �o�S�V�b�N", Osaka, "�l�r �S�V�b�N", Gothic, sans-serif;
        }

        Span.gakoumei { font-size:12pt; color:#ffffff; }

        b { font-weight: bold; }
        hr { border-style:solid;  border-color:#886688; }

        /* A�@�A���J�[ ��{*/
        a:link { color:#ffffff; font-size:10pt; text-decoration:none; }
        a:visited { color:#ffffff; font-size:10pt; text-decoration:none; }
        a:active { color:#FF8364; font-size:10pt; text-decoration:none; }
        a:hover { color:#FF8364; font-size:10pt; text-decoration:underline; }
    //-->
    </style>
    </head>

<body rightmargin="0" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table border="0" cellspacing="0" cellpadding="0" width="100%" height="61">
<tr>
<td width="156" height="61" align="left" valign="top">
<img src="images/title.gif">
</td>

<td bgcolor="#56567F" height="61" width="100%" background="images/back_gla.gif">

	<table cellspacing="0" cellpadding="0" border="0" height="61" width="100%">
	<tr>
	<td height="21" width="100%" align="right" valign="middle" nowrap>
	<font color="#ffffff" size="2">
	<%= m_LoginDay %>
	</font><img src="images/sp.gif" height="1" width="20">
	</td>

	<td height="21" align="right" valign="top">
		<table height="21" width="82" cellspacing="0" cellpadding="0" border="0">
		<tr>
		<td height="21">
		<a href="../manual/default.asp" target="_blank">
		<img src="images/help.gif" border="0">
		</a>
		</td>
		<td height="21">
		<a href="../default.asp" target="_top">
		<img src="images/logout.gif" border="0">
		</a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<tr>
	<td height="40" width="100%" align="left" valign="top">
		<table cellspacing="0" cellpadding="0" border="0" height="41" width="100%">
		<tr>
		<td>
		<img src="images/sp.gif" height="1" width="20"><Span class="gakoumei"><%= m_SchoolName %></Span>
		</td>
		<td align="right">
			<table cellspacing="0" cellpadding="0" border="0">
			<tr>
			<td>
			<font color="#ffffff">���[�U�[��</font>
			</td>
			<td>
			<font color="#ffffff">:</font>
			</td>
			<td>
			<font color="#ffffff"><%= Session("USER_NM") %></font>
			</td>
			<td>
			<img src="../image/sp.gif" width="15">
			</td>
			</tr>
			</table>
		</td>
		</tr>
		</table>
	</td>
	<td height="40" align="center">
	<font color="#ffffff"><%= m_WaNengappi %></font>
	</td>
	</tr>
	</table>

</td>
</tr>
</table>

</body>
</html>

<% End Sub %>