<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �i�H����o�^
' ��۸���ID : mst/mst0144/kakunin.asp
' �@      �\: ���y�[�W �i�H��}�X�^�̓o�^�m�F���s��
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
' ��      ��: 2001/06/22 �≺ �K��Y
' ��      �X: 2001/07/12 �J�e �ǖ�
' �@      �@: 2001/07/24 ���{ ����(DB�ύX�ɔ����C��)
' �@�@�@�@�@: 2001/08/22 �ɓ��@���q�@�Ǝ�敪�ǉ��Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�

    Public  m_sSINROMEI         ':�i�H��
    'Public m_sSINROMEI_EIGO    ':�i�H���p��
    Public  m_sSINROMEI_KANA    ':�i�H���J�i
    Public  m_sSINRORYAKSYO     ':�i�H����
    'Public m_sJUSYO            ':�Z��
    Public  m_sJUSYO1           ':�Z��1
    Public  m_sJUSYO2           ':�Z��2
    Public  m_sJUSYO3           ':�Z��3
    Public  m_iKenCd            ':���R�[�h
    Public  m_iSityoCd          ':�s�����R�[�h
    Public  m_sSityoson         ':�s������
    Public  m_sDENWABANGO       ':�d�b�ԍ�
    Public  m_iYUBINBANGO       ':�X�֔ԍ�
    Public  m_sSINRO_URL        ':URL
    Public  m_sSinrokubun       ':�f�[�^�x�[�X����擾�����i�H�敪
    Public  m_sSingakukubun     ':�f�[�^�x�[�X����擾�����i�w�敪
    Public  m_Rs                ':recordset
    Public  m_iNendo            ':�N�x
    Public  m_sYubin            ':�X�֔ԍ�
    Public  m_iGyosyu_Kbn       ':�Ǝ�敪
    Public  m_iSihonkin         ':���{��
    Public  m_iJyugyoin_Suu     ':�]�ƈ�
    Public  m_iSyoninkyu        ':���C��
    Public  m_sBiko             ':���l

    Public  m_sRenrakusakiCD    ':�A����R�[�h
    Public  m_sSinroCD      ':�i�H�R�[�h
    Public  m_sSingakuCD        ':�i�w�R�[�h
    Public  m_sSinroCD2     ':Main����擾�����i�H�R�[�h
    Public  m_sSingakuCD2       ':Main����擾�����i�w�R�[�h
    Public  m_sSyusyokuName     ':�i�H���́i�ꕔ�j
    Public  m_sPageCD       ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_sMode
    Public  m_bReFlg
    Public  m_sMsgFlg       ':�G���[�t���O
    Public  m_sMsg

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
    w_sMsgTitle="�i�H����o�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_sMsgFlg = False

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

        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI_KANA "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_KEN_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SITYOSON_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo & " AND "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO_CD = '" & m_sRenrakusakiCD & "' "

'Response.Write w_sSQL & "<br>"

        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If

		'//�V�K���A�f�[�^�̏d���`�F�b�N
		If m_sMode = "Sinki" Then
	        If m_Rs.EOF = False Then
	            m_sMsgFlg = True
	            m_sMsg = "���͂��ꂽ�i�H��R�[�h�͂��łɎg�p����Ă��܂�"
	        End If
		End If

        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    call gf_closeObject(m_Rs)
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

    m_sMode          = Request("txtMode")               '���[�h�̐ݒ�
    m_bReFlg         = Request("txtReFlg")              '�����[�h�L���̐ݒ�

    m_sRenrakusakiCD = Request("txtRenrakusakiCD")      ':�A����R�[�h
    m_sSINROMEI      = Request("txtSINROMEI")           ':�i�H��
    m_sSINROMEI_KANA = Request("txtSINROMEI_KANA")      ':�i�H���J�i
    'm_sSINROMEI_EIGO = Request("txtSINROMEI_EIGO")     ':�i�H���p��
    'If m_sSINROMEI_EIGO="" Then m_sSINROMEI_EIGO="�@"
    m_sSINRORYAKSYO  = Request("txtSINRORYAKSYO")       ':�i�H����
    If m_sSINRORYAKSYO="" Then m_sSINRORYAKSYO="�@"
    'm_sJUSYO         = Request("txtJUSYO")             ':�Z��
    m_iKenCd         = Request("txtKenCd")              ':���R�[�h
    m_iSityoCd       = Request("txtSityoCd")            ':�s�����R�[�h�i�Z��1�j
    m_sJUSYO1        = Request("txtJUSYO1")             ':�Z��1
    m_sJUSYO2        = Request("txtJUSYO2")             ':�Z��2
    m_sJUSYO3        = Request("txtJUSYO3")             ':�Z��3
    m_iKenCd         = Request("txtKenCd")              ':���R�[�h
    m_iSityoCd       = Request("txtSityoCd")            ':�s�����R�[�h
    m_sDENWABANGO    = Request("txtDENWABANGO")         ':�d�b�ԍ�
    m_sSinroCD       = Request("txtSinroCD")            ':�i�H�敪
    m_sSingakuCD     = Request("txtSingakuCd")          ':�i�w�敪
    m_sSINRO_URL     = Request("txtSINRO_URL")          'URL
    if Instr(m_sSINRO_URL,"http://") = 0 then m_sSINRO_URL = "http://" & m_sSINRO_URL
    if m_sSINRO_URL  = "http://" then m_sSINRO_URL = ""
    m_sSinroCD2      = Request("txtSinroCD2")           ':�߂�p�̐i�H�敪
    m_sSingakuCD2    = Request("txtSingakuCd2")         ':�߂�p�̐i�w�敪
    m_sSyusyokuName  = Request("txtSyusyokuName")       '�߂�p��:�A�E�於�́i�ꕔ�j
    m_sPageCD        = INT(Request("txtPageCD"))        '�߂�p�̕\����

    m_iNendo        = Request("txtNendo")               ':�N�x
    m_sYubin        = Request("txtYUBINBANGO")          ':�X�֔ԍ�
    'm_iGyosyu_Kbn  = Request("txtGYOSYU_KBN")          ':�Ǝ�敪
    m_iSihonkin = Request("txtSIHONKIN")                ':���{��
    m_iJyugyoin_Suu = Request("txtJYUGYOIN_SUU")        ':�]�ƈ���
    m_iSyoninkyu    = Request("txtSYONINKYU")           ':���C��
    m_sBiko         = Request("txtBIKO")                ':���l

    
    If strErrmsg <> "" Then
        ' �G���[��\������t�@���N�V����
        Call err_page(strErrMsg)
        response.end
    End If
'   call s_viewForm(request.form)   '�f�o�b�O�p�@�����̓��e������
End Sub

'********************************************************************************
'*  [�@�\]  �s���������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SityosonMei()

        '// �i�H�敪�����擾
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M12_SITYOSON "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "        M12_KEN_CD = '" & m_iKenCd & "'"
        w_sSQL = w_sSQL & vbCrLf & "    AND M12_SITYOSON_CD = " & m_iSityoCd & " "
        w_sSQL = w_sSQL & vbCrLf & "    GROUP BY M12_SITYOSONMEI "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Sub
        End If

    m_sSityoson = w_Rs("M12_SITYOSONMEI")


End Sub

Sub S_syousai()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_slink
Dim w_iCnt

	w_iCnt = 0

	Do While not m_Rs.EOF

	w_slink = "�@"

	if m_Rs("M32_SINRO_URL") <> "" Then 
	    w_sLink= "<a href='" & gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "'>" 
	    w_sLink= w_sLink &  gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "</a>"
	End if

	        %>
	        <%=w_slink%>
	        <%
	            m_Rs.MoveNext
    Loop

End sub


Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

%>

<html>

    <head>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageSinro.value = p_iPage;
        document.frm.submit();
    
    }

    function f_OpenWindow(p_Url){
    //************************************************************
    //  [�@�\]  �q�E�B���h�E���I�[�v������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
        var window_location;
        window_location=window.open(p_Url,"window","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,scrolling=no,Width=500,Height=500");
        window_location.focus();
    }

    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_GoSyosai(p_sSinroKBN){

        document.frm.action="./syousai.asp";
        document.frm.target="";
        document.frm.txtMode.value = "Syosai";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_BackClick(){

        document.frm.action="./syusei.asp";
        document.frm.target="_self";
        document.frm.txtReFlg.value = "<%=m_bReFlg%>";
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

	    document.frm.action="update.asp";
	    document.frm.target="_self";
	    document.frm.submit();
    }

    //************************************************************
    //  [�@�\]  �E�C���h�E�I�[�v����
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function Window_onload(){
    <%
    If m_sMsgFlg = True Then
    %>
    alert("<%=m_sMsg%>")

      document.frm.action="./syusei.asp";
      document.frm.target="_self";
      document.frm.txtMode.value = "Sinki";
      document.frm.submit();

    <%
    End If
    %>
    }
    //-->
    </SCRIPT>

    <link rel=stylesheet href="../../common/style.css" type=text/css>

    </head>

<body onload="Window_onload()">

<center>

<form name="frm" action="update.asp" target="_self" method=post>

<%
If m_sMode = "Sinki" Then
  m_sSubtitle = "�V�K�o�^"
else
  m_sSubtitle = "�C�@��"
End If

call gs_title("�i�H����o�^",m_sSubtitle)
%>
<br>
�i�@�H�@��@��@��
<br><br>

    <table border="0" class=form width=75%>
    <tr>
    <td class=form align="left" width="100">�i�H��R�[�h</td>
    <td class=form align="left">
    <%= m_sRenrakusakiCD %>
    <input type="hidden" name="txtRenrakusakiCD" value="<%= m_sRenrakusakiCD %>">
    </td>
    </tr>

    <tr>
    <td class=form align="left">���@��</td>
    <td class=form align="left">
    <%= m_sSINROMEI %>
    <input type="hidden" name="txtSINROMEI" value="<%= m_sSINROMEI %>">
    </td>
    </tr>
    <!--
    <tr>
    <td class=form align="left">���@�́i�p��j</td>
    <td class=form align="left">
    <%= m_sSINROMEI_EIGO %>
    <input type="hidden" name="txtSINROMEI_EIGO" value="<%= m_sSINROMEI_EIGO %>">
    </td>
    </tr>
    //-->
    <tr>
    <td class=form align="left">���@�́i�J�i�j</td>
    <td class=form align="left">
    <%= m_sSINROMEI_KANA %>
    <input type="hidden" name="txtSINROMEI_KANA" value="<%= m_sSINROMEI_KANA %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">���@��</td>
    <td class=form align="left">
    <%= m_sSINRORYAKSYO %>
    <input type="hidden" name="txtSINRORYAKSYO" value="<%= m_sSINRORYAKSYO %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">�X�֔ԍ�</td>
    <td class=form align="left">
    <%= m_sYubin %>
    <input type="hidden" name="txtYUBINBANGO" value="<%= m_sYubin %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">�Z�@���i�P�j</td>
    <td class=form align="left">
    <%= m_sJUSYO1 %>
    <input type="hidden" name="txtJUSYO1" value="<%= m_sJUSYO1 %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">�Z�@���i�Q�j</td>
    <td class=form align="left">
    <%= m_sJUSYO2 %>
    <input type="hidden" name="txtJUSYO2" value="<%= m_sJUSYO2 %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">�Z�@���i�R�j</td>
    <td class=form align="left">
    <%= m_sJUSYO3 %>
    <input type="hidden" name="txtJUSYO3" value="<%= m_sJUSYO3 %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">�d�b�ԍ�</td>
    <td class=form align="left">
    <%= m_sDENWABANGO %>
    <input type="hidden" name="txtDENWABANGO" value="<%= m_sDENWABANGO %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">�i�H�敪</td>
    <td class=form align="left">

    <%
	'// �i�H�敪���̊m��
	 Call gf_GetKubunName(C_SINRO,m_sSinroCD,m_iNendo,m_sSinrokubun)
	 response.write m_sSinrokubun 
	%>

    <input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD %>">
    </td>
    </tr>

	<tr>
	<%
	'=================================
	'//�i�H�敪�ɂ��\����ς���
	'=================================
	w_sKbnName = ""
	Select case cint(gf_SetNull2Zero(m_sSinroCD))
		Case C_SINRO_SINGAKU	'//�i�H�敪���i�w�̏ꍇ

			'//�i�w�敪���̂��擾
			Call gf_GetKubunName(C_SINGAKU,m_sSingakuCD,m_iNendo,w_sKbnName)

		Case C_SINRO_SYUSYOKU	'//�i�H�敪���A�E�̏ꍇ

			'//�Ǝ�敪���̂��擾
			Call gf_GetKubunName(C_GYOSYU_KBN,m_sSingakuCD,m_iNendo,w_sKbnName)

		Case C_SINRO_SONOTA	'//�i�H�敪�����̑��̏ꍇ
	End Select
	%>

    <td class=form align="left">��ʋ敪</td>
    <td class=form align="left"><%=w_sKbnName%>
    <input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">���{��</td>
    <td class=form align="left">
    <%= m_iSihonkin %>
    <input type="hidden" name="txtSIHONKIN" value="<%= m_iSihonkin %>">���~
    </td>
    </tr>
    <tr>
    <td class=form align="left">�]�ƈ���</td>
    <td class=form align="left">
    <%= m_iJyugyoin_Suu %>
    <input type="hidden" name="txtJYUGYOIN_SUU" value="<%= m_iJyugyoin_Suu %>">�l
    </td>
    </tr>
    <tr>
    <td class=form align="left">���C��</td>
    <td class=form align="left">
    <%= m_iSyoninkyu %>
    <input type="hidden" name="txtSYONINKYU" value="<%= m_iSyoninkyu %>">�~
    </td>
    </tr>
    <tr>
    <td class=form align="left">�t�@�q�@�k</td>
    <td class=form align="left">
    <%= m_sSINRO_URL %>
    <input type="hidden" name="txtSINRO_URL" value="<%= m_sSINRO_URL %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">���@�l</td>
    <td class=form align="left">
    <%= m_sBiko %>
    <input type="hidden" name="txtBIKO" value="<%= m_sBiko %>">
    </td>
    </tr>
    </table>
<br>
�ȏ�̓��e�œo�^���܂��B
<br><br>
<table border="0">
<tr>
<td valign="top">

<input type="button" class=button value="�@�o�@�^�@" Onclick="f_SinkiClick()">
<img src="../../image/sp.gif" width="20" height="1">
<input type="button" class=button value="�L�����Z��" Onclick="f_BackClick()">

</td>
</tr>
</table>

<input type="hidden" name="txtMode" value="<%= m_sMode %>">
<input type="hidden" name="txtReFlg" value="<%= m_bReFlg %>">
<input type="hidden" name="txtSinroCD2" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD2" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<!--<input type="hidden" name="txtNendo" value="<%= Session("SYORI_NENDO") %>">-->
<input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
<input type="hidden" name="txtKenCd" value="<%= m_iKenCd %>">
<input type="hidden" name="txtSityoCd" value="<%= m_iSityoCd %>">
</form>


</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
<%
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