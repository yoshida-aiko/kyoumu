<%@Language=VBScript %>
<%
'******************************************************************
'�V�X�e����     �F���������V�X�e��
'���@���@��     �F�e��ψ��o�^
'�v���O����ID   �Fgak/gak0470/default.asp
'�@�@�@�@�\     �F�t���[���y�[�W �w�Јψ������͂̕\�����s��
'------------------------------------------------------------------
'���@�@�@��     �F
'�ρ@�@�@��     �F
'���@�@�@�n     �F
'���@�@�@��     �F
'------------------------------------------------------------------
'��@�@�@��     �F2001.07.02    �O�c�@�q�j
'��      �X     : 2001/08/09 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'
'******************************************************************
'*******************�@ASP���ʃ��W���[���錾�@**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******�@�� �W �� �[ �� �� ���@********
    '�y�[�W�֌W
Public  m_iMax          ':�ő�y�[�W
Public  m_iDsp                      '// �ꗗ�\���s��
Public  m_bErrFlg       '//�G���[�t���O�iDB�ڑ��G���[���̏ꍇ�ɃG���[�y�[�W��\�����邽�߂̃t���O�j
Public  m_sDebugStr     '//�ȉ��f�o�b�N�p
Public  m_iNendo
Public  m_sKyokanCd
Public  m_rs            '//���R�[�h�Z�b�g
Public  m_Irs           '//���R�[�h�Z�b�g�i�ψ��p�j
Public  m_Grs           '//���R�[�h�Z�b�g�i�w�Дԍ��p�j
Public  m_sDaiNm()
Public  m_iDai()
Public  m_iSyo()
Public  m_iIinNm()

Public  m_IrCnt           '//���R�[�h�J�E���g
Public  m_GrCnt           '//���R�[�h�J�E���g
Public  m_iGAKKIKBN '�w���敪

'******�@���C�������@********

    'Ҳ�ٰ�ݎ��s
    Call Main()

'******�@�d�@�m�@�c�@********

Sub Main()
'******************************************************************
'�@�@�@�\�F�{ASP��Ҳ�ٰ��
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    '******���ʊ֐�******
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�e��ψ��o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")

   if Request("cboGakki") = "" then 'cboGakki

	m_iGAKKIKBN = session("GAKKI")
   
   else

	m_iGAKKIKBN = Request("cboGakki")

   End if

Response.Write m_iGAKKIKBN

    m_iDsp = C_PAGE_LINE

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
        session("PRJ_No") = "GAK0470"

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '// �S�C�`�F�b�N
	  If gf_Tannin(m_iNendo,m_sKyokanCd,1) <> 0 Then
	            m_bErrFlg = True
	            m_sErrMsg = "�S�C�ȊO�̓��͂͂ł��܂���B"
	            Exit Do
	  End If

        w_iRet = f_getData()
        If w_iRet <> 0 Then
            '�G���[����
            m_bErrFlg = True
            Exit Do
        End If
		If m_IrCnt = 0 then
	        '// �y�[�W��\��
	        Call showPage_NO()
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

    '// �I������
    Call gs_CloseDatabase()
End Sub

Function f_getData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

Dim i
i = 1

    On Error Resume Next
    Err.Clear
    f_getData = 1

    Do
        '//�w�N��N���X�̃f�[�^
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT"
        m_sSQL = m_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M05_CLASS "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     M05_NENDO = " & Cint(m_iNendo) & " "
        m_sSQL = m_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_rs, m_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

        '//���X�g(�ψ���ʁC�ψ�����)�̃f�[�^
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT"
        m_sSQL = m_sSQL & "     M34_DAIBUN_CD,M34_SYOBUN_CD,M34_IIN_NAME "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M34_IIN "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     M34_NENDO = " & Cint(m_iNendo) & "  "
        m_sSQL = m_sSQL & " AND M34_IIN_KBN <> " & C_IIN_GAKKO & " "
        m_sSQL = m_sSQL & " UNION"
        m_sSQL = m_sSQL & " SELECT"
        m_sSQL = m_sSQL & "     M34_DAIBUN_CD,M34_SYOBUN_CD,M34_IIN_NAME "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M34_IIN "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     M34_NENDO = " & Cint(m_iNendo) & "  "
        m_sSQL = m_sSQL & " AND M34_SYOBUN_CD = " & C_M34_SYOBUN_CD & " "

        Set m_Irs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Irs, m_sSQL)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_IrCnt=gf_GetRsCount(m_Irs)
       '//���X�g(����)�̃f�[�^
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT "
        m_sSQL = m_sSQL & "     A.T06_GAKUSEI_NO,A.T06_DAIBUN_CD,A.T06_SYOBUN_CD,B.T11_SIMEI "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T06_GAKU_IIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     A.T06_NENDO = " & Cint(m_iNendo) & " "
        m_sSQL = m_sSQL & " AND C.T13_GAKUNEN  = " & Cint(m_rs("M05_GAKUNEN")) & " "
        m_sSQL = m_sSQL & " AND C.T13_CLASS = " & Cint(m_rs("M05_CLASSNO")) & " "
        m_sSQL = m_sSQL & " AND A.T06_NENDO = C.T13_NENDO "
        m_sSQL = m_sSQL & " AND A.T06_GAKUSEI_NO = B.T11_GAKUSEI_NO "
        m_sSQL = m_sSQL & " AND B.T11_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		m_sSQL = m_sSQL & " AND A.T06_GAKKI_KBN = " & m_iGAKKIKBN

        Set m_Grs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Grs, m_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_GrCnt=gf_GetRsCount(m_Grs)

        m_Irs.Movefirst

        Do Until m_Irs.EOF
            If Cint(m_Irs("M34_SYOBUN_CD")) = 0 Then
                ReDim Preserve m_sDaiNm(m_Irs("M34_DAIBUN_CD"))
                m_sDaiNm(m_Irs("M34_DAIBUN_CD")) = m_Irs("M34_IIN_NAME")
            Else
                ReDim Preserve m_iDai(i)
                ReDim Preserve m_iSyo(i)
                ReDim Preserve m_iIinNm(i)
                m_iDai(i) = m_Irs("M34_DAIBUN_CD")
                m_iSyo(i) = m_Irs("M34_SYOBUN_CD")
                m_iIinNm(i) = m_Irs("M34_IIN_NAME")
                i = i + 1
            End If
            
            m_Irs.MoveNext
            
        Loop

    f_getData = 0

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    On Error Resume Next
    Err.Clear

    Dim i
    i = 0 

    '---------- HTML START ----------
    %>
    <html>
    <head>
    <title>�e��ψ��o�^</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <script language="javascript">
    <!--
    //************************************************************
    //  [�@�\]  �\�����e�\���p�E�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function NewWin(p_Int,p_Str,p_GInt,p_CInt,p_sNm) {

		var obj=eval("document.frm."+p_sNm)
        URL = "select.asp?i="+p_Int+ "&IINNM="+escape(p_Str)+"&GAKUNEN="+p_GInt+"&CLASS="+p_CInt+"&GName="+escape(obj.value)+"";
        //URL = "select.asp?i="+p_Int+ "&IINNM="+escape(p_Str)+"&GAKUNEN="+p_GInt+"&CLASS="+p_CInt+"&GName="+escape(p_sNm)+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,toolbar=no,width=700,height=600,top=0,left=0");
        return true;    
    }

    //************************************************************
    //  [�@�\] �N���A�{�^���������ꂽ�Ƃ�
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function jf_Clear(p_Name,p_Cd){
        eval("document.frm."+p_Name).value = "";
        eval("document.frm."+p_Cd).value = "";
    }

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Touroku(){

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        //���X�g����submit
        document.frm.action="gak0470_edt.asp";
        document.frm.submit();

    }
	//************************************************************
    //  [�@�\]  �w�����ύX���ꂽ�Ƃ��A�{��ʂ��ĕ\��
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./default.asp";
        //document.frm.target="fTopMain";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //-->
    </script>

    </head>
    <body>
    <center>
    <form action="" name="frm" method="post">
<table border="0" width="90%">
<tr>
<td align="center">
<%call gs_title("�e��ψ��o�^","��@��")%>

    <table border="0" cellpadding="1" cellspacing="1" width="100%">
    <tr>
    <td align="center">
        <table border="0" width="500" class=hyo align="center">
        <tr>
        <th width="65" class="header">�w�N</th>
        <td width="50" align="center" class="detail"><%=m_rs("M05_GAKUNEN")%>�N</td>
        <th width="65" class="header">�N���X</th>
        <td width="220" align="center" class="detail"><%=m_rs("M05_CLASSMEI")%></td>
		<th width="50" class="header">�w��</th>
		<td width="50" class="detail">
		<select name="cboGakki" onchange = 'javascript:f_ReLoadMyPage()' ><Option Value="1"�@Selected>�O��
		<Option Value="2" Selected>���</select></td>
        </tr>
        </table></td>
    </tr>
    <tr>
    <td align="center">
 <!--    <img src="../../image/sp.gif" height="30"> -->
<span class="msg">�e�ψ���>>�{�^���������ƁA�w���I����ʂ��o�Ă��܂��B<BR>�w����I�����A�o�^�{�^���������ĉ������B</span>
    </td>
    </tr>
    <tr>
    <td align="center">

        <table width="100%" border="1" class="hyo">
        <tr>
        <th width="25%" class="header">�ψ����</th>
        <th width="25%" class="header">�ψ�����</th>
        <th width="40%" class="header">���@��</th>
        <th width="5%" class="header">�I��</th>
        <th width="5%" class="header">�@</th>
        </tr>

        <tr>

        <%
        For i = 1 to UBound(m_iDai)
            call gs_cellPtn(w_cell)%>

            <td  class="<%=w_cell%>" align="center"><font color="#000000"><%= m_sDaiNm(m_iDai(i)) %></font></td>
            <td  class="<%=w_cell%>" align="center"><font color="#000000"><%= m_iIinNm(i) %><br></font></td>

                <%
                If m_Grs.EOF = False Then
                    w_Name = ""
                    w_Gakusei_No = ""
                    m_Grs.MoveFirst
                    Do Until m_Grs.EOF
                        If Cint(m_iDai(i)) = Cint(m_Grs("T06_DAIBUN_CD")) and Cint(m_iSyo(i)) = Cint(m_Grs("T06_SYOBUN_CD")) Then
                            w_Name = m_Grs("T11_SIMEI")
                            w_Gakusei_No = m_Grs("T06_GAKUSEI_NO")
                            Exit Do
                        End If
                        m_Grs.MoveNext
                    Loop
                    m_Grs.MoveFirst
                End If %>
            <td  class="<%=w_cell%>" align="center">
                <input type="text" class="<%=w_cell%>" name="gakuNm<%=i%>" value="<%= w_Name %>" readonly><br>
                <input type="hidden" name="gakuNo<%=i%>" value="<%= w_Gakusei_No %>">
                <input type="hidden" name="iinDai<%=i%>" value="<%= Cint(m_iDai(i)) %>">
                <input type="hidden" name="iinSyo<%=i%>" value="<%= Cint(m_iSyo(i)) %>">
                <input type="hidden" name="Before<%=i%>" value="<%= w_Gakusei_No %>"></td>
            <!--<td  class="<%=w_cell%>" align="center"><input type="button" class="button" value=">>" onclick="NewWin(<%=i%>,'<%= m_iIinNm(i) %>',<%=m_rs("M05_GAKUNEN") %>,<%=m_rs("M05_CLASSNO") %>,'<%= w_Name %>')"></td>-->
            <td  class="<%=w_cell%>" align="center"><input type="button" class="button" value=">>" onclick="NewWin(<%=i%>,'<%= m_iIinNm(i) %>',<%=m_rs("M05_GAKUNEN") %>,<%=m_rs("M05_CLASSNO") %>,'gakuNm<%=i%>')"></td>
            <td  class="<%=w_cell%>"><input type="button" class="button" value="�N���A" onclick="jf_Clear('gakuNm<%=i%>','gakuNo<%=i%>')"></td>
        </tr>
        <%Next%>
        </table>

    </td>
    </tr>
    </table><br><br>
        <input type="button" value="�o�@�^" class="button" onclick="javascript:f_Touroku()">

    <INPUT TYPE="HIDDEN" NAME="HIDMAX" VALUE="<%= i-1 %>">
    <INPUT TYPE="HIDDEN" NAME="CLASS" VALUE="<%= m_rs("M05_CLASSNO") %>">
    <INPUT TYPE="HIDDEN" NAME="GAKUNEN" VALUE="<%= m_rs("M05_GAKUNEN")%>">
	<INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
	<INPUT TYPE="HIDDEN" NAME="GAKKI"	  value = "<% m_iGAKKI %>">

</td>
</tr>
</table>
    </form>
    </center>
    </body>
    </html>

<%
    '---------- HTML END   ----------
End Sub

Sub showPage_NO()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
    <html>
    <head>
    <title>�e��ψ��o�^</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <script language="javascript">
    </script>
    </head>
    <body>
    <center>
    <form action="" name="frm" method="post">
<table border="0" width="90%">
<tr>
<td align="center">
<%call gs_title("�e��ψ��o�^","��@��")%>
<br><br><br><br><br>
        <span class="msg">�w�Јψ����̃f�[�^������܂���B</span>


    </center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub
%>