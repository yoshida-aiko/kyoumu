<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �}�C�y�[�W�̂��m�点
' ��۸���ID : login/top_lwr.asp
' �@      �\: ���y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/23 �O�c
' ��      �X: 2001/08/07 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'           : 2001/08/10 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp          '// �ꗗ�\���s��
    Public  m_rs
    Dim     m_iNendo
    Dim     m_sKyokanCd
    Dim     m_sUserId
    Dim     m_sName

    '�G���[�n
    Public  m_bErrFlg       '�װ�׸�
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

    Dim w_iRet              '// �߂�l

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�������o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

        '// �����`�F�b�N�Ɏg�p
        session("PRJ_No") = C_LEVEL_NOCHK

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '// ���Ұ�SET
        Call s_SetParam()

        '//���X�g�̈ꗗ�f�[�^�̏ڍ׎擾
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ں��޾��CLOSE
    Call gf_closeObject(m_rs)
    '// �I������
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
'response.write "(" & session("KYOKAN_CD") & ")kyoukan <br>"
	m_sUserId = session("LOGIN_ID")
'response.write "(" & session("LOGIN_ID") & ")kyoukan <br>"
    m_iDsp = C_PAGE_LINE

End Sub

Function f_GetData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
	Dim w_user

    On Error Resume Next
    Err.Clear
    f_GetData = 1

	w_user = m_sUserId
	'���[�U�������ł���΁A����CD����
	'If m_sKyokanCd <> "" then w_user = m_sKyokanCd

	'T46�ɂ́A���[�U�[ID�œo�^����Ă���̂ŁA�����R�[�h�ł͊Y�����Ȃ�
	'���[�U�[ID�Œ��o�A�X�V����悤�ɕύX�@2001/12/11 �ɓ�	
	w_user = m_sUserId

    Do
        '//���X�g�̕\��
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT * "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T46_RENRAK "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T46_KYOKAN_CD = '" & w_user & "' "
        m_sSQL = m_sSQL & " AND T46_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "'"
        m_sSQL = m_sSQL & " AND T46_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"
        m_sSQL = m_sSQL & " ORDER BY T46_KAKNIN"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_rCnt=gf_GetRsCount(m_rs)

    f_GetData = 0

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Function f_Sosin()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_sUserCD
Dim w_rs

    On Error Resume Next
    Err.Clear

    Do
        If m_rs("T46_UPD_USER") <> "" Then
            w_sUserCD = m_rs("T46_UPD_USER")
        Else
            w_sUserCD = m_rs("T46_INS_USER")
        End If

        '//���M�҂̐����̎擾
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT "
        m_sSQL = m_sSQL & "     M10_USER_NAME "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M10_USER "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     M10_NENDO = " & m_iNendo & " AND "
        m_sSQL = m_sSQL & "     M10_USER_ID = '" & w_sUserCD & "' "

        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

        f_Sosin = w_rs("M10_USER_NAME")

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub S_syousai()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear

%>
    <table class="disp" width="100%" border="1">
        <colgroup width="10%">
        <colgroup width="50%">
        <colgroup width="15%">
        <colgroup width="25%">
        <tr>
            <td class="disph" align="center" nowrap>�@<br></td>
            <td class="disph" align="center" nowrap>�p�@��</td>
            <td class="disph" align="center" nowrap>���@�t</td>
            <td class="disph" align="center" nowrap>�� �M ��</td>
        </tr>

        <% m_rs.Movefirst
            Do Until m_rs.EOF 
            call gs_cellPtn(w_cell)%>
        <%Call f_Sosin
            m_sName = f_Sosin()
        %>
        <TR>
            <%If cInt(m_rs("T46_KAKNIN")) = C_KAKU_SUMI Then%>
                    <TD CLASS="<%=w_cell%>" nowrap>�@<br></TD>
            <%Else%>
                    <TD CLASS="<%=w_cell%>" align=center nowrap>��</TD>
            <%End If%>
            <TD CLASS="<%=w_cell%>" nowrap>�E<a href="#" onclick="NewWin(<%=m_rs("T46_NO")%>,'<%=m_sName%>');"><%=m_rs("T46_KENMEI")%></a></TD>

            <%If m_rs("T46_UPD_USER") <> "" Then%>
                    <TD CLASS="<%=w_cell%>" nowrap><%=m_rs("T46_UPD_DATE")%></TD>
            <%Else%>
                    <TD CLASS="<%=w_cell%>" nowrap><%=m_rs("T46_INS_DATE")%></TD>
            <%End If%>
            <TD CLASS="<%=w_cell%>" nowrap><%=m_sName%></TD>
        </TR>
        <% m_rs.MoveNext : Loop %>
    </table>
    <br>
    <Div align="center"><span class="CAUTION">�� �p�����N���b�N����Ƒ��t���e���m�F�ł��܂��B<br>
	<div align="center"><span class=CAUTION>�� ���b�Z�[�W�́A�\�����Ԃ��߂���Ǝ����I�ɍ폜����܂��B<br>
	</span></div>


<%
End Sub

Function s_Jikanwari(p_hyoji)
'********************************************************************************
'*  [�@�\]  ���Ԋ��ύX�f�[�^�̗L���̊m�F
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	Dim w_user

	s_Jikanwari = false
	p_hyoji = ""
	

	w_user = m_sUserId
	'���[�U�������ł���΁A����CD����
	If m_sKyokanCd <> "" then w_user = m_sKyokanCd

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT * "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T52_JYUGYO_HENKO "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & "     T52_KYOKAN_CD = '" & w_user & "' "
    w_sSQL = w_sSQL & " AND T52_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "'"
    w_sSQL = w_sSQL & " AND T52_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"

    Set m_Rds = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(m_Rds, w_sSQL,m_iDsp)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    End If

    If m_Rds.EOF Then
        Exit Function
    End If
	
	p_hyoji = ""
'	p_hyoji = p_hyoji & "<HR>"
	p_hyoji = p_hyoji & "<CENTER>"
	p_hyoji = p_hyoji & "<a href='#' onclick=NewWinJik()> �����Ԋ��̕ύX�A�����m�F����ꍇ�͂������N���b�N���ĉ������B</a>  "
	p_hyoji = p_hyoji & "</CENTER>"
	p_hyoji = p_hyoji & "<BR>"
	s_Jikanwari = true

End Function

Sub showPage_NoData()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
    <center>
        �A�������͂���܂���B
    </center>

<%
    '---------- HTML END   ----------
End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    Dim i			'��Ɨp
    Dim w_bDataF	'�\���f�[�^�t���O �i�A���������̘A��������ꍇ�ɗ��Ă�j
    w_bDataF = false
    i=1

%>
<HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link rel="stylesheet" href="../common/style.css" type="text/css">
    <title>���������V�X�e���FCanpus Assist �g�b�v�y�[�W</title>

    <!--#include file="../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [�@�\]  �\�����e�\���p�E�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function NewWin(p_Int,p_sSEIMEI) {
        URL = "view.asp?txtNo="+p_Int+"&txtSEIMEI="+escape(p_sSEIMEI)+"";
        <%if session("browser") = "NN" Then%>
        nWin=window.open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,outerwidth=400,outerheight=300,top=0,left=0");
        <%else%>
        nWin=window.open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=400,height=300,top=0,left=0");
        <%end if%>
        nWin.focus();
        return false;   
    }

    //************************************************************
    //  [�@�\]  �\�����e�\���p�E�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function NewWinJik() {
        URL = "j_view.asp";
        <%if session("browser") = "NN" Then%>
	        nWin=window.open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,outerwidth=400,outerheight=300,top=0,left=0");
        <%else%>
	        nWin=window.open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=450,height=450,top=0,left=0");
        <%end if%>
        nWin.focus();
        return false;   
    }
    //-->
    </SCRIPT>
</head>
<BODY>
<center>
<!--
<hr width="80%" size="1">
-->
<br>
<font size="3">���@�m�@��@��</font>
<br><br>
<FORM NAME="frm" ACTION="post">
<input type="hidden" name="txtNo">
<input type="hidden" name="txtSEIMEI">
<table width="90%"><tr><td>
<%
	'//���Ԋ��ύX�A���̕\��
    If s_Jikanwari(w_hyoji) = true Then
		response.write w_hyoji
        w_bDataF = true
	End If
	'//�A�������̕\��
    If m_rs.EOF = false Then
        Call S_syousai()
        w_bDataF = true
	End If 


	'//��L�̓�Ƃ��Ȃ��ꍇ
	If w_bDataF = false then 
        Call showPage_NoData()
    End If

%>
</td></tr></table>
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
