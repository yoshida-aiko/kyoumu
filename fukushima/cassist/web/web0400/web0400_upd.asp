<%@ Language=VBScript %>
<%Response.Expires = 0%>
<%Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �p�X���[�h�ύX
' ��۸���ID : web/web0400/default.asp
' �@      �\: ���O�C���p�X���[�h��ύX���܂��B
'-------------------------------------------------------------------------
' ��      ��:SESSION(""):�����R�[�h     ��      SESSION���
' ��      ��:�Ȃ�
' ��      �n:SESSION(""):�����R�[�h     ��      SESSION���
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/10/04 �J�e
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public m_bErrFlg           '�װ�׸� �V�X�e���n
    Public m_bErrFlgPara           '�װ�׸� �p�����[�^�`�F�b�N�n
    Public m_iNendo
    Public m_sUser
    Public m_sPass
    Public m_sPassN1
    Public m_sPassN2

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()
response.end
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
    Dim w_iLevel

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�p�X���[�h�ύX"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
	m_bErrFlgPara = False
    Do

		'// �ϐ�������
		call f_paraSet()

			'// �ް��ް��ڑ�
			w_iRet = gf_OpenDatabase()
			If w_iRet <> 0 Then
				'�ް��ް��Ƃ̐ڑ��Ɏ��s
				m_bErrFlg = True
				m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
				Exit Do
			End If

			'// �����`�F�b�N�Ɏg�p
	'		session("PRJ_No") = "WEB0400"

			'// �s���A�N�Z�X�`�F�b�N
			'Call gf_userChk(session("PRJ_No"))
			
			'// �s���A�N�Z�X�`�F�b�N
			w_iRet = f_login(m_sUser,m_sPass,w_iLevel)
			If w_iRet = false Then
				'�ް��ް��Ƃ̐ڑ��Ɏ��s
				If m_bErrFlgPara = true then 
					w_sMsg = "���O�C��ID�ƃp�X���[�h����v���܂���ł����B<BR>������x�A���O�C��ID�A�p�X���[�h���m�F�̏�A�ύX�{�^���������Ă��������B"
				else
					m_bErrFlg = True
					m_sErrMsg = "���O�C���f�[�^�̎擾���ł��܂���ł����B"
				End If
				Exit Do
			End If

			'// �����`�F�b�N
			w_iRet = f_TT51(w_iLevel)
			If w_iRet = false Then
				'�ް��ް��Ƃ̐ڑ��Ɏ��s
				If m_bErrFlgPara = true then 
					w_sMsg = "�p�X���[�h��ύX���錠��������܂���B"
				else
					m_bErrFlg = True
					m_sErrMsg = "���O�C���f�[�^�̎擾���ł��܂���ł����B"
				End If
				Exit Do
			End If

			'// �X�V����
			w_iRet = f_Update()
			If w_iRet = false Then
				'�ް��ް��Ƃ̐ڑ��Ɏ��s
					m_bErrFlg = True
					m_sErrMsg = "�f�[�^�̍X�V���ł��܂���ł����B"
				Exit Do
			End If

        '// �ύX�y�[�W��\��
        Call showPage()
'        Call showErrPage("�听��")
        Exit Do
    Loop

    '// �p�����[�^�̴װ�̏ꍇ�̓p�����[�^�װ�߰�ނ�\��
    If m_bErrFlgPara = True Then
        Call showErrPage(w_sMsg)
    End If
    
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Sub

Sub f_paraSet()
'*******************************************************************************
' �@�@�@�\�F�ϐ��̏������Ƒ��
' ���@�@���F�Ȃ�
' �@�\�ڍׁF
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
m_iNendo = session("NENDO")
m_sMode = Request("txtMode")
m_sUser = Request("txtUser")
m_sPass = Request("txtPass")
m_sPassN1 = Request("txtPassN1")
m_sPassN2 = Request("txtPassN2")
'm_iNendo = 2001
	
End Sub


Function f_login(p_id,p_pass,p_level)
'********************************************************************************
'*  [�@�\]  ���O�C������
'*  [����]  p_id   = հ�ް�����͂���հ�ްID
'*          p_pass = հ�ް�����͂����߽ܰ��
'*  [�ߒl]  p_level= հ�ް�̌���
'*  [����]  
'********************************************************************************
    Dim w_sSql
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_login = false

  Do
    '// Null�Ȃ甲����
    if trim(p_id) = "" then
        Exit Function
    Elseif trim(p_pass) = "" then
        Exit Function
    End if

    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & "     M10_USER_ID, "      '0
    w_sSql = w_sSql & "     M10_KYOKAN_CD, "    '1
    w_sSql = w_sSql & "     M10_USER_NAME, "    '2
    w_sSql = w_sSql & "     M10_USER_KBN, "     '3
    w_sSql = w_sSql & "     M10_LEVEL "         '4
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & "     M10_USER  "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & "     M10_NENDO    =  " & m_iNendo & " AND "
    w_sSql = w_sSql & "     M10_USER_ID  = '" & p_id & "' AND "
    w_sSql = w_sSql & "     M10_PASSWORD = '" & p_pass  & "' "

    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit do 'GOTO LABEL_MAIN_END
    End If

	'// ں��޾�Ă��Ȃ������甲����
	If m_Rs.Eof then
        m_bErrFlgPara = True
        Exit do 'GOTO LABEL_MAIN_END
	End if

	'// �����擾
	p_level = m_Rs("M10_LEVEL")

    f_login = true
    exit do
  Loop

    call gf_closeObject(m_Rs)

End Function

Function f_TT51(p_level)
'********************************************************************************
'*  [�@�\]  �����`�F�b�N
'*  [����]  p_level = հ�ް�̌���
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    Dim w_sSql
    Dim w_sLevel
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_TT51 = false

  Do
    '// Null�Ȃ甲����
    if trim(p_level) = "" then
        Exit Function
    End if

    w_sLevel = "T51_LEVEL" & trim(p_level)

    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & w_sLevel
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & "     TT51_SYORI_LEVEL  "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & "     T51_ID         = 'WEB0400' AND "
    w_sSql = w_sSql & "     T51_SYORI_KBN  = 12 "

    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit do
    End If

    '// ں��޾�Ă��Ȃ������甲����
    If m_Rs.Eof then
        m_bErrFlgPara = True
        Exit do
    End if

    '// 1�FOK
    If m_Rs(w_sLevel) <> "1" then
        m_bErrFlgPara = True
        Exit do
    End If

    f_TT51 = true
    exit do
  Loop

    call gf_closeObject(m_Rs)

End Function

Function f_Update()
'********************************************************************************
'*  [�@�\]  �p�X���[�h�ύX
'*  [����]  �Ȃ�
'*  [�ߒl]  true:���� false:���s
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    
    f_Update = false

    Do 

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

            '//T11_GAKUSEKI��UPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE M10_USER SET "
            w_sSQL = w_sSQL & vbCrLf & "   M10_PASSWORD = '"  & Trim(m_sPassN1) & "' ,"
            w_sSQL = w_sSQL & vbCrLf & "   M10_UPD_DATE = '"  & gf_YYYY_MM_DD(date(),"/") & "', "
            w_sSQL = w_sSQL & vbCrLf & "   M10_UPD_USER = '"  & Session("LOGIN_ID") & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        M10_USER_ID = '" & Trim(m_sUser) & "' AND "
            w_sSQL = w_sSQL & vbCrLf & "        M10_NENDO = " & m_iNendo & " "

            iRet = gf_ExecuteSQL(w_sSQL)
            If iRet <> 0 Then
                '//۰��ޯ�
                Call gs_RollbackTrans()
                msMsg = Err.description
                f_Update = 99
                Exit Do
            End If

        '//�Я�
        Call gs_CommitTrans()

        '//����I��
        f_Update = true
        Exit Do
    Loop

End Function

Sub showErrPage(p_msg)
'********************************************************************************
'*  [�@�\]  �G���[HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
<html>
<head>
    <title>�p�X���[�h�ύX</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
</head>
<body>
<center>
    <%call gs_title("���O�C���p�X���[�h�ύX","�X�@�V")%>

<BR>
	<table class="hyo" border="0" width="70%">
		<FORM action="default.asp" name="frm" method="post" target="_self">
			<tr><td colspan="2" height="30" class="detail"></td></tr>
			<tr><td colspan="2" height="15" class="detail" align="center"><%=p_msg%></td></tr>
			<tr><td colspan="2" height="30" class="detail"></td></tr>
			<tr><td colspan="2" height="30" class="detail" align="center">
				<input type="submit" name="submit" value=" �� �� " maxlength="16"> <!-- 2023.10.25 Upd Kiyomoto �p�X���[�h��10����16���ɕύX -->
	            </td>
	        </tr>
<input type="hidden" name="txtUser" value="<%=m_sUser%>">
<input type="hidden" name="txtPass" value="<%=m_sPass%>">
<input type="hidden" name="txtPassN1" value="<%=m_sPassN1%>">
<input type="hidden" name="txtPassN2" value="<%=m_sPassN2%>">
		</FORM>
	</table>
</center>
</body>
</head>
</html>
<%
End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  �G���[HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
<html>
<head>
    <title>�p�X���[�h�ύX</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
		function f_load(){
			alert("�p�X���[�h��ύX���܂����B");
			top.location.href="../../default.asp";
		}
	//-->
	</SCRIPT>
</head>
<body onload="f_load();">
<center>
    <%call gs_title("���O�C���p�X���[�h�ύX","�X�@�V")%>
</center>
</body>
</head>
</html>
<%
End Sub
%>
