<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���O�C�����
' ��۸���ID : login/default.asp
' �@      �\: ���O�C�����
'-------------------------------------------------------------------------
' ��      ��    Request("txtLogin")   = ۸޲�ID
'               Request("txtPass")    = �߽ܰ��
'               Request("hidLoginFlg) = ۸޲݉�ʂ��痈�����邵
'           
' ��      ��
' ��      �n    Session("NENDO")            '�N�x
'               Session("LOGIN_ID")         '���O�C���b�c
'               Session("USER_NM")          '���[�U�l�[��
'               Session("LEVEL")            '����
'               Session("USER_KBN")         '���[�U�敪
'               Session("KYOKAN_CD")        '����CD

' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 
' ��      �X: 2001/07/26    ���`�i�K
'*************************************************************************/
%>
<!--#include file="../common/com_All.asp"-->
<%

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_Rs                'ں��޾�ĵ�޼ު��
    Public  m_bErrFlg           '�װ�׸�
    Public  m_iMiKengenFlg		'�������������Ȃ��׸�


'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////


'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sRetURL           '// �װү���ޗp�߂��URL
    Dim w_sTarget           '// �װү���ޗp�߂���ڰ�
    Dim w_sWinTitle         '// �װү���ޗp����
    Dim w_sMsgTitle         '// �װү���ޗp����
    
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�w�Ѓf�[�^��������"
    w_sMsg=""
    w_sRetURL="../default.asp"
    w_sTarget="_parent"

     Do

		'// ���O�C����ʂ��痈���ꍇ�́A���O�C���`�F�b�N������
		if Cint(Request("hidLoginFlg")) = Cint(C_LOGIN_FLG) then

	        '// ���Ұ��擾
	        w_LoginID  = Request("txtLogin")
	        w_PassWord = Request("txtPass")

			'// �ް��ް��ڑ�
			w_iRet = gf_OpenDatabase()
			If w_iRet <> 0 Then
				'�ް��ް��Ƃ̐ڑ��Ɏ��s
				m_bErrFlg = True
				m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
				Exit Do
			End If

			'// �N�x�擾
			If Not f_GetNendo() then Exit Do End if

			'// ۸޲�����
			If Not f_login(w_LoginID,w_PassWord) Then
				Call ErrPage()          '// �G���[�y�[�W��\��
			End if

		End if

		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		Call showPage()         '// �y�[�W��\��

        '// ����I��
        Exit Do
    LOOP

   '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, m_sErrMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �N�x���o��
'*  [����]  
'*          
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetNendo()
    Dim w_sSql
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_GetNendo = False

    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & "     A.M00_KANRI "
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & "     M00_KANRI A "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & "     A.M00_NENDO  = " & C_M00_NENDO & " AND "
    w_sSql = w_sSql & "     A.M00_NO     = 0 "

    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_NendoRs, w_sSQL)

    If w_iRet <> 0 Then
    'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

    '// �e������Z�b�V�����Ɋi�[
    session("NENDO") = w_NendoRs("M00_KANRI")           '�N�x

    f_GetNendo = True

    call gf_closeObject(w_NendoRs)

End Function

'********************************************************************************
'*  [�@�\]  ���O�C������
'*  [����]  p_id   = հ�ް�����͂���հ�ްID
'*          p_pass = հ�ް�����͂����߽ܰ��
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_login(p_id,p_pass)
    Dim w_sSql
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_login = false

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
    w_sSql = w_sSql & "     M10_NENDO    =  " & session("NENDO") & " AND "
    w_sSql = w_sSql & "     M10_USER_ID  = '" & p_id & "' AND "
    w_sSql = w_sSql & "     M10_PASSWORD = '" & p_pass  & "' "

    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

	'// ں��޾�Ă��Ȃ������甲����
	If m_Rs.Eof then
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
	End if

    '// �P�����̒P�ʐ����擾
    w_iRet = f_GetJigen_Tani(w_iTani)
    If w_iRet <> 0 Then
        m_bErrFlg = True
        Exit Function
    End If

    '// �e������Z�b�V�����Ɋi�[
    Session("LOGIN_ID")  = m_RS("M10_USER_ID")          '���O�C���b�c
    Session("USER_NM")   = m_Rs("M10_USER_NAME")        '���[�U�l�[��
    Session("LEVEL")     = m_Rs("M10_LEVEL")            '����
    Session("USER_KBN")  = m_Rs("M10_USER_KBN")         '���[�U�敪
    Session("KYOKAN_CD") = m_Rs("M10_KYOKAN_CD")        '����CD
	
	application("KYOKAN_CD") = m_Rs("M10_KYOKAN_CD")        '����CD
	
	Session("JIKAN_TANI") = cint(w_iTani)				'�P�����̒P��(����)��
    f_login = true


    call gf_closeObject(m_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �P�����̒P�ʐ����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_iTani : �P�����̎��Ԑ����擾(��{�́A�P�΂P)
'*  [����]  
'********************************************************************************
Function f_GetJigen_Tani(p_iTani)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetJigen_Tani = 1
	p_iTani = ""

    Do 

		'//�����}�X�^���P�ʐ����擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(M07_TANISU) AS TANISU "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M07_JIGEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M07_NENDO=" & session("NENDO")

'response.write w_sSQL  & "<BR>"

         iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSikenKbn = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_iTani = cint(w_Rs("TANISU"))
		End If

		'//�f�[�^���擾�ł��Ȃ��Ƃ��́A�P�ɂ���B
		If p_iTani = "" Then
			p_iTani = 1
		End If

        f_GetJigen_Tani = 0

        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  ���O�C���G���[�\���y�[�W
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub ErrPage()
%>
 <HTML>
   <HEAD>
    <meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
    <TITLE></TITLE>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_loginErr(){
        alert("���O�C���G���[\n���O�C��ID�ƃp�X���[�h�����m���߂̏�A\n�ēx���O�C�����Ă��������B");
        location.href="<%=C_RetURL%>default.asp"
        return true;
    }

            </SCRIPT>
        </HEAD>
        <BODY onLoad="f_loginErr();">
            <br>
        </BODY>
    </HTML>
<%
    End Sub

'********************************************************************************
'*  [�@�\]  HTML�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
%>
<html>

<head>
<title>Campus Assist</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>

<frameset rows=61,*,24 frameborder="0" FRAMESPACING="0" BORDER="0">
    <frame src="header.asp" scrolling="no" NAME="topHead">
    <frameset cols=166,* frameborder="0" FRAMESPACING=0 frameborder="0">
        <frame src="menu.asp" scrolling="auto" noresize name="menu">
        <frame src="top.asp" scrolling="auto" noresize name="<%=C_MAIN_FRAME%>">
    </frameset>
        <frame src="foot.asp" scrolling="auto" noresize name="foot">
</frameset>

</html>
<%
End Sub
%>