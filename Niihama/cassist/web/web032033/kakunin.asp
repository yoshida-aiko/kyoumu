<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �g�p���ȏ��o�^�m�F
' ��۸���ID : mst/mst0144/kakunin.asp
' �@      �\: ���y�[�W �A�E��}�X�^�̏ڍוύX���s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroCD      :�i�H�R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSyusyokuName     :�i�H���́i�ꕔ�j
'           txtPageSinro        :�\���ϕ\���Ő��i�������g����󂯎������j
'           Sinro_syuseiCD      :�I�����ꂽ�i�H�R�[�h
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroCD      :�i�H�R�[�h�i�߂�Ƃ��j
'           txtSingakuCd        :�i�w�R�[�h�i�߂�Ƃ��j
'           txtSyusyokuName     :�i�H���́i�߂�Ƃ��j
'           txtPageSinro        :�\���ϕ\���Ő��i�߂�Ƃ��j
' ��      ��:
'           �������\��
'               �w�肳�ꂽ�i�w��E�A�E��̏ڍ׃f�[�^��\��
'           ���n�}�摜�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ��i�w��E�A�E���\������i�ʃE�B���h�E�j
'-------------------------------------------------------------------------
' ��      ��: 2001/07/12 �≺ �K��Y
' ��      �X: 2001/08/22 �ɓ� ���q ������I���ł���悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_sDBMode           'DB��Ӱ�ނ̐ݒ�

    '�擾�����f�[�^�����ϐ�
    Public  m_Rs            'recordset
    Public  m_sNendo
    Public  m_sGakkiCD
    Public  m_sNo
    Public  m_sGakunenCD
    Public  m_sGakkaCD
    Public  m_sCourseCD
    Public  m_sKamokuCD
    Public  m_sKyokanMei
    Public  m_sKyokasyoName
    Public  m_sSyuppansya
    Public  m_sTyosya
    Public  m_sKyokanyo
    Public  m_sSidousyo
    Public  m_sBiko
    Public  m_sKyokan_CD

    ''����
    Public  m_sSYOBUNRUI_CD
    Public  m_sSYOBUNRUIMEI
    Public  m_sGAKKAMEI
    Public  m_sCOURSEMEI
    Public  m_sKAMOKUMEI

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
    Dim w_sSQL              '// SQL��
    Dim w_sWHERE            '// WHERE��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�g�p���ȏ��o�^�m�F"
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
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '// ���Ұ�SET
        Call s_SetParam()

        '// ��ʂɕ\�����閼�̂��擾
        if f_Get_Name = False then
            exit do
        end if

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


'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    Dim strErrMsg

    strErrMsg = ""
    m_sDBMode    = Request("txtMode")
    m_sNendo     = Request("txtNendo")      '�N�x�̎擾
    m_sGakkiCD   = Request("txtGakkiCD")    '�w���̎擾
    m_sNo = Request("txtUpdNo")             ''�X�V�pNo�i�[
    m_sGakunenCD     = Request("txtGakunenCD")  '�w�N�̎擾
    m_sGakkaCD   = Request("txtGakkaCD")    '�w�Ȃ̎擾
    m_sCourseCD  = Request("txtCourseCD")   '�R�[�X�̎擾
    m_sKamokuCD  = Request("txtKamokuCD")   '�Ȗڂ̎擾
    m_sKyokanMei     = Request("txtKyokanMei")  '�������̎擾
    m_sKyokasyoName  = Request("txtKyokasyoName")   '���ȏ����̎擾
    m_sSyuppansya    = Request("txtSyuppansya") '�o�ŎЂ̎擾
    m_sTyosya    = Request("txtTyosya")     '���҂̎擾
    m_sKyokanyo  = Request("txtKyokanyo")   '�����p�̎擾
    m_sSidousyo  = Request("txtSidousyo")   '�w�����̎擾
    m_sBiko      = Request("txtBiko")       '�����p�̎擾

    m_sKyokan_CD = Request("SKyokanCd1")

    m_sSYOBUNRUI_CD = ""
    m_sSYOBUNRUIMEI = ""
    m_sGAKKAMEI = ""
    m_sCOURSEMEI = ""
    m_sKAMOKUMEI = ""

	If m_sKyokanyo = "" Then
	  m_sKyokanyo = 0
	End If

	If m_sSidousyo = "" Then
	  m_sSidousyo = 0
	End If

    If strErrmsg <> "" Then
        ' �G���[��\������t�@���N�V����
        Call err_page(strErrMsg)
        response.end
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
Exit Sub

	response.write("<BR>m_sDBMode = " & m_sDBMode)
	response.write("<BR>m_sNendo = " & m_sNendo)
	response.write("<BR>m_sGakkiCD = " & m_sGakkiCD)
	response.write("<BR>m_sNo = " & m_sNo)
	response.write("<BR>m_sGakunenCD = " & m_sGakunenCD)
	response.write("<BR>m_sGakkaCD = " & m_sGakkaCD)
	response.write("<BR>m_sCourseCD = " & m_sCourseCD)
	response.write("<BR>m_sKamokuCD = " & m_sKamokuCD)
	response.write("<BR>m_sKyokanMei = " & m_sKyokanMei)
	response.write("<BR>m_sKyokasyoName = " & m_sKyokasyoName)
	response.write("<BR>m_sSyuppansya = " & m_sSyuppansya)
	response.write("<BR>m_sTyosya = " & m_sTyosya)
	response.write("<BR>m_sKyokanyo = " & m_sKyokanyo)
	response.write("<BR>m_sSidousyo = " & m_sSidousyo)
	response.write("<BR>m_sBiko = " & m_sBiko)

End Sub

'********************************************************************************
'*  [�@�\]  ���̂��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
function f_Get_Name
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_Rs

    f_Get_Name = False

	'==============
    ''�w�����̎擾
	'==============
    w_sSQL = ""
    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " M01.M01_SYOBUNRUIMEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_SYOBUNRUI_CD "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    M01_KUBUN M01 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    'w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD  =  " & 51 & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD  =  " & C_KAISETUKI & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M01.M01_SYOBUNRUI_CD  =  " & m_sGakkiCD & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M01.M01_NENDO         =  " & m_sNendo

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    End If

	If w_Rs.EOF = false Then
	    m_sSYOBUNRUI_CD = gf_HTMLTableSTR(w_Rs("M01_SYOBUNRUI_CD"))
	    m_sSYOBUNRUIMEI = gf_HTMLTableSTR(w_Rs("M01_SYOBUNRUIMEI"))
	End If

    w_Rs.close
    set w_Rs = nothing

	'=================
    ''�w�ȏ����擾
	'=================
    If cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) Then
        m_sGAKKAMEI = "�S�w��"
    else
	    w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M02.M02_GAKKAMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKA M02 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M02.M02_NENDO         =  " & m_sNendo & " AND "
        If cstr(m_sGakkaCD) <> cstr(C_CLASS_ALL) Then
                w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKA_CD          = '" & m_sGakkaCD & "'"
        End If

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Function
        End If

		If w_Rs.EOF = false Then
	        m_sGAKKAMEI = gf_HTMLTableSTR(w_Rs("M02_GAKKAMEI"))
		End If

        w_Rs.close
        set w_Rs = nothing
    end if

	'=================
    ''��������擾
	'=================
    If cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) Then
        m_sCOURSEMEI = ""
    else

        If m_sCOURSECD <> "@@@" AND m_sCOURSECD <> "" Then
		    w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " SELECT "
            w_sSQL = w_sSQL & vbCrLf & " M20.M20_COURSEMEI "
            w_sSQL = w_sSQL & vbCrLf & " FROM "
            w_sSQL = w_sSQL & vbCrLf & "    M20_COURSE M20 "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "    M20.M20_NENDO         =  " & m_sNendo & " AND "
            w_sSQL = w_sSQL & vbCrLf & "    M20_GAKKA_CD          = '" & m_sGakkaCD & "' AND "
            w_sSQL = w_sSQL & vbCrLf & "    M20_GAKUNEN           =  " & m_sGakunenCD & " AND "
            w_sSQL = w_sSQL & vbCrLf & "    M20_COURSE_CD         = '" & m_sCOURSECD & "'"

            w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
            If w_iRet <> 0 Then
                'ں��޾�Ă̎擾���s
                m_bErrFlg = True
                Exit Function
            End If

			If w_Rs.EOF = false Then
	            m_sCOURSEMEI = gf_HTMLTableSTR(w_Rs("M20_COURSEMEI"))
			End If

            w_Rs.close
            set w_Rs = nothing
        else
            m_sCOURSEMEI = ""
        end if
    end if

	'=================
    ''�Ȗڏ����擾
	'=================
    w_sSQL = ""
    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " T15.T15_KAMOKUMEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    T15_RISYU T15 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    T15.T15_NYUNENDO      =  " & (m_sNendo - m_sGakunenCD + 1) & " AND "
    If cstr(m_sGakkaCD) <> cstr(C_CLASS_ALL) Then
        w_sSQL = w_sSQL & vbCrLf & "    T15_GAKKA_CD          = '" & m_sGakkaCD & "' AND "
    else
        w_sSQL = w_sSQL & vbCrLf & "    T15_KAMOKU_KBN          = 0 AND "
    End If
    w_sSQL = w_sSQL & vbCrLf & "    T15_KAMOKU_CD         = '" & m_sKAMOKUCD & "'"

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    End If

	If w_Rs.EOF = false Then
	    m_sKAMOKUMEI = gf_HTMLTableSTR(w_Rs("T15_KAMOKUMEI"))
	End If

    w_Rs.close
    set w_Rs = nothing

    f_Get_Name = True

end function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()

%>

<html>

    <head>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_BackClick(){

	    document.frm.txtMode.value = "Disp";
	    document.frm.action="./touroku.asp";
	    document.frm.target="_self";
	    document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_SinkiClick(){

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

	    document.frm.txtMode.value = "<%=m_sDBMode%>";
	    document.frm.action='./db.asp';
	    document.frm.target="_self";
	    document.frm.submit();
    }

    //-->
    </SCRIPT>

    <link rel=stylesheet href="../../common/style.css" type=text/css>

    </head>

	<body>

	<center>

	<form name="frm" action="" target="_self">

	<br>

	<% call gs_title("�g�p���ȏ��o�^",Request("txtTitle")) %>

	<br>

	�o�@�^�@���@�e
	<br><br>

	    <table border="0" class=form width=75%>
	    <tr>
	    <td class=form align="left" width="100">�N�@�x</td>
	    <td class=form align="left">
	    <%= m_sNendo %>�N
	    <input type="hidden" name="txtNendo" value="<%= m_sNendo %>">
	    <BR></td>
	    </tr>

	    <tr>
	    <td class=form align="left">�w�@��</td>
	    <td class=form align="left">
	    <%= m_sSYOBUNRUIMEI %>
	    <input type="hidden" name="txtGakkiCD" value="<%= m_sSYOBUNRUI_CD %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">�w�@�N</td>
	    <td class=form align="left">
	    <%= m_sGakunenCD %>�N
	    <input type="hidden" name="txtGakunenCD" value="<%= m_sGakunenCD %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">�w�@��</td>
	    <td class=form align="left">
	<%
	If cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) then
	    response.write "�S�w��"
	  Else
	    response.write m_sGAKKAMEI
	End If
	%>
	    <input type="hidden" name="txtGakkaCD" value="<%= m_sGakkaCD %>">
	    <BR></td>
	    </tr>

	    <tr>
	    <td class=form align="left">�R�[�X</td>
	    <td class=form align="left">
	<%
	If m_sCOURSECD <> "@@@" AND m_sCOURSECD <> "" Then
	response.write m_sCOURSEMEI
	End If
	%>
	    <input type="hidden" name="txtCourseCD" value="<%If m_sCourseCD = "@@@" Then
	response.write ""
	Else
	response.write m_sCourseCD 
	End If%>">
	    <BR></td>
	    </tr>

	    <tr>
	    <td class=form align="left">�ȁ@��</td>
	    <td class=form align="left">
	    <%= m_sKAMOKUMEI %>
	    <input type="hidden" name="txtKamokuCD" value="<%= m_sKamokuCD %>">
	    <BR></td>
	    </tr>

	    <tr>
	    <td class=form align="left">���@��</td>
	    <td class=form align="left">
	    <%= m_sKyokanMei %>
	    <input type="hidden" name="txtKyokanMei" value="<%= m_sKyokanMei %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">���ȏ���</td>
	    <td class=form align="left">
	    <%= m_sKyokasyoName %>
	    <input type="hidden" name="txtKyokasyoName" value="<%= m_sKyokasyoName %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">�o�Ŏ�</td>
	    <td class=form align="left">
	    <%= m_sSyuppansya %>
	    <input type="hidden" name="txtSyuppansya" value="<%= m_sSyuppansya %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">���Җ�</td>
	    <td class=form align="left">
	    <%= m_sTyosya %>
	    <input type="hidden" name="txtTyosya" value="<%= m_sTyosya %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">�����p</td>
	    <td class=form align="left"><%= m_sKyokanyo %>��
	    <input type="hidden" name="txtKyokanyo" value="<%= m_sKyokanyo %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">�w����</td>
	    <td class=form align="left"><%= m_sSidousyo %>��
	    <input type="hidden" name="txtSidousyo" value="<%= m_sSidousyo %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">���l</td>
	    <td class=form align="left">
	    <%= m_sBiko %>
	    <input type="hidden" name="txtBiko" value="<%= m_sBiko %>">

	    <BR></td>
	    </tr>
	    </table>
	<br>
	�ȏ�̓��e�œo�^���܂��B
	<br><br>
	<table border="0">
	<tr>
	<td valign="top">
	<input type="button" class=button value="�@�o�@�^�@" Onclick="f_SinkiClick()">
	<input type="hidden" name="txtTitle" value="<%= Request("txtTitle") %>">
	<input type="hidden" name="txtUpdNo" value="<%= Request("txtUpdNo") %>">
	<img src="../../image/sp.gif" width="20" height="1">
	<input type="button" class=button value="�L�����Z��" Onclick="f_BackClick()">

	<!--�l�n���p-->
	<input type="hidden" name="txtMode" value="">
    <input type="hidden" name="SKyokanCd1" value="<%=m_sKyokan_CD%>">

    <input type="hidden" name="KeyNendo" value="<%=Request("KeyNendo")%>">

	</form>
	</td>
	</tr>
	</table>
	</center>
	</body>
	</html>

<%
    '---------- HTML END   ----------
End Sub


'**********  �G���[��\������t�@���N�V����  *********
Function err_page(myErrMsg)
%>
    <html>
    <head>
    <title>���ڃG���[</title>
    <link rel=stylesheet href=bar.css type=text/css>
    </head>

    <body bgcolor="#ffffff">
    <center>
    <form>
    <font size="2">
    Error:���ڃG���[<br><br>
    �ȉ��̍��ڂ̃G���[���łĂ��܂��B<br><br>

    <%=myErrMsg%>

    <br><br>
    �ȏ�̍��ڂ���͂��čēx���M���Ă��������B<p>
    <input class=button type="button" class=button value="�L�����Z��" onclick="JavaScript:history.back();">

    </font>

    </form>
    </center>
    </body>
    </html>
<%
End Function
%>