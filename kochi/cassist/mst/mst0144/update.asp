<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A�E��}�X�^
' ��۸���ID : mst/mst0144/update.asp
' �@      �\: ���y�[�W �A�E��}�X�^�̏ڍוύX���s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H�R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSyusyokuName     :�i�H���́i�ꕔ�j
'           txtPageSinro        :�\���ϕ\���Ő��i�������g����󂯎������j
'           Sinro_syuseiCD      :�I�����ꂽ�i�H�R�[�h
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H�R�[�h�i�߂�Ƃ��j
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
' ��      �X: 2001/07/13 �J�e�@�ǖ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�

    Public  m_sRenrakusakiCD        ':�A����R�[�h
    Public  m_sSINROMEI             ':�i�H��
    'Public m_sSINROMEI_EIGO        ':�i�H���p��
    Public  m_sSINROMEI_KANA        ':�i�H���J�i
    Public  m_sSINRORYAKSYO         ':�i�H����
    'Public m_sJUSYO                ':�Z��
    Public  m_sJUSYO1               ':�Z��1
    Public  m_sJUSYO2               ':�Z��2
    Public  m_sJUSYO3               ':�Z��3
    Public  m_iKenCd                ':���R�[�h
    Public  m_iSityoCd              ':�s�����R�[�h
    Public  m_sDENWABANGO           ':�d�b�ԍ�
    Public  m_sSinro_syuseiCD       ':�i�H�敪
    Public  m_sSINRO_URL            ':URL
    Public  m_iNendo        ':�N�x
    Public  m_sDATE
    Public  m_sKyokanCD
    Public  m_sMode
    Public  m_sYubin            ':�X�֔ԍ�
    Public  m_iGyosyu_Kbn       ':�Ǝ�敪
    Public  m_iSihonkin         ':���{���i�P�ʁF���~�j
    Public  m_iSihonkinY         ':���{���i�P�ʁF�~�j
    Public  m_iJyugyoin_Suu     ':�]�ƈ�
    Public  m_iSyoninkyu        ':���C��
    Public  m_sBiko             ':���l

    Public  m_sSinroCD      ':�i�H�R�[�h
    Public  m_sSingakuCD        ':�i�w�R�[�h
    Public  m_sSinroCD2     ':�i�H�R�[�h
    Public  m_sSingakuCD2       ':�i�w�R�[�h
    Public  m_sSyusyokuName     ':�i�H���́i�ꕔ�j
    Public  m_iPageCD       ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_Rs            'recordset


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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E��}�X�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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
        
        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()
        
        '// �����̐U�蕪��
        If m_sMode = "Sinki" then
            call s_ins(w_sSQL)
        Else
            call s_update(w_sSQL)
        End If

'Response.Write w_sSQL & "<br>"

		w_iRet = gf_ExecuteSQL(w_sSQL)
        If w_iRet <> 0 Then
            '���s
            '//۰��ޯ�
            Call gs_RollbackTrans()
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If
        
        '//�Я�
        Call gs_CommitTrans()

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
    call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'/* �V�K�o�^
Sub s_ins(p_sSQL)

        p_sSQL = p_sSQL & vbCrLf & " Insert Into "
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO"
        p_sSQL = p_sSQL & "(M32_NENDO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_CD,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI,"
        'p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI_EIGO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI_KANA,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRORYAKSYO,"
        'p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO1,"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO2,"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO3,"
        p_sSQL = p_sSQL & vbCrLf & " M32_KEN_CD,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SITYOSON_CD,"
        p_sSQL = p_sSQL & vbCrLf & " M32_DENWABANGO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_YUBIN_BANGO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_KBN,"

		If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
			'//�i�HCD��1(�i�w)�̏ꍇ
	        p_sSQL = p_sSQL & vbCrLf & " M32_SINGAKU_KBN,"
		ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
			'//�i�HCD��2(�A�E)�̏ꍇ
	        p_sSQL = p_sSQL & vbCrLf & " M32_GYOSYU_KBN,"
		End If

        p_sSQL = p_sSQL & vbCrLf & " M32_SIHONKIN,"
        p_sSQL = p_sSQL & vbCrLf & " M32_JYUGYOIN_SUU,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SYONINKYU,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_URL,"
        p_sSQL = p_sSQL & vbCrLf & " M32_BIKO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_INS_DATE,"
        p_sSQL = p_sSQL & vbCrLf & " M32_INS_USER,"
        p_sSQL = p_sSQL & vbCrLf & " M32_UPD_DATE,"
        p_sSQL = p_sSQL & vbCrLf & " M32_UPD_USER)"
        p_sSQL = p_sSQL & vbCrLf & " Values"
        p_sSQL = p_sSQL & "(" & m_iNendo & ","
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sRenrakusakiCD) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sSINROMEI) & "',"
        'p_sSQL = p_sSQL & vbCrLf & "'" & m_sSINROMEI_EIGO & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sSINROMEI_KANA) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sSINRORYAKSYO) & "',"
        'p_sSQL = p_sSQL & vbCrLf & "'" & m_sJUSYO & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sJUSYO1) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sJUSYO2) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sJUSYO3) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_iKenCd) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_iSityoCd) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sDENWABANGO) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sYubin) & "',"
        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_sSinroCD) & ","

		If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
			'//�i�HCD��1(�i�w)�̏ꍇ
	        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_sSingakuCD) & ","
		ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
			'//�i�HCD��2(�A�E)�̏ꍇ
			If trim(replace(m_sSingakuCD,"@@@","")) <> "" Then
		        p_sSQL = p_sSQL & vbCrLf & "" & trim(m_sSingakuCD) & ","
			Else
		        p_sSQL = p_sSQL & vbCrLf & " NULL,"
			End If

		End If

        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_iSihonkinY) & ","
        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_iJyugyoin_Suu) & ","
        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_iSyoninkyu) & ","
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sSINRO_URL) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sBiko) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sDATE) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "',"
        p_sSQL = p_sSQL & vbCrLf & "'',"
        p_sSQL = p_sSQL & vbCrLf & "'')"
End Sub

Sub s_update(p_sSQL)
        p_sSQL = ""
        p_sSQL = p_sSQL & vbCrLf & " Update "
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO M32"
        p_sSQL = p_sSQL & vbCrLf & " Set "
        p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI         = '" & trim(m_sSINROMEI) & "',"
        'p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI_EIGO   = '" & m_sSINROMEI_EIGO & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI_KANA    = '" & trim(m_sSINROMEI_KANA) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRORYAKSYO     = '" & trim(m_sSINRORYAKSYO) & "',"
        'p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO           = '" & m_sJUSYO & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO1           = '" & trim(m_sJUSYO1) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO2           = '" & trim(m_sJUSYO2) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO3           = '" & trim(m_sJUSYO3) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_KEN_CD           = '" & trim(m_iKenCd) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_SITYOSON_CD      = '" & trim(m_iSityoCd) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_DENWABANGO       = '" & trim(m_sDENWABANGO) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_YUBIN_BANGO      = '" & trim(m_sYubin) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_KBN        =  " & trim(m_sSinroCD) & ","

		If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
			'//�i�HCD��1(�i�w)�̏ꍇ
	        p_sSQL = p_sSQL & vbCrLf & " M32_SINGAKU_KBN     =  " & trim(m_sSingakuCD) & ","
		ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
			'//�i�HCD��2(�A�E)�̏ꍇ
			If trim(replace(m_sSingakuCD,"@@@","")) <> "" Then
		        p_sSQL = p_sSQL & vbCrLf & " M32_GYOSYU_KBN      = " & trim(m_sSingakuCD) & ","
			Else
		        p_sSQL = p_sSQL & vbCrLf & " M32_GYOSYU_KBN      = NULL,"
			End If
		End If

        p_sSQL = p_sSQL & vbCrLf & " M32_SIHONKIN      =  " & trim(m_iSihonkinY)    & ","
        p_sSQL = p_sSQL & vbCrLf & " M32_JYUGYOIN_SUU  =  " & trim(m_iJyugyoin_Suu) & ","
        p_sSQL = p_sSQL & vbCrLf & " M32_SYONINKYU     =  " & trim(m_iSyoninkyu)    & ","
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_URL     = '" & trim(m_sSINRO_URL)    & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_BIKO          = '" & trim(m_sBiko)         & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_UPD_DATE      = '" & trim(m_sDATE)         & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_UPD_USER      = '" & Session("LOGIN_ID")   & "'"
        p_sSQL = p_sSQL & vbCrLf & " WHERE "
        p_sSQL = p_sSQL & vbCrLf & "    M32_NENDO      =  " & m_iNendo              & " AND "
        p_sSQL = p_sSQL & vbCrLf & " M32.M32_SINRO_CD  = '" & m_sRenrakusakiCD      &"' "

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_sMode = Request("txtMode")

    m_sRenrakusakiCD = Request("txtRenrakusakiCD")          ':�A����R�[�h

    m_sSINROMEI      = Request("txtSINROMEI")       ':�i�H��
    'm_sSINROMEI_EIGO = Request("txtSINROMEI_EIGO")     ':�i�H���p��
    'If m_sSINROMEI_EIGO="�@" Then m_sSINROMEI_EIGO=""
    m_sSINROMEI_KANA = Request("txtSINROMEI_KANA")      ':�i�H���p��
    m_sSINRORYAKSYO  = Request("txtSINRORYAKSYO")       ':�i�H����
    If m_sSINRORYAKSYO="�@" Then m_sSINRORYAKSYO=""
    'm_sJUSYO         = Request("txtJUSYO")         ':�Z��
    m_sJUSYO1         = Request("txtJUSYO1")            ':�Z��1
    m_sJUSYO2         = Request("txtJUSYO2")            ':�Z��2
    m_sJUSYO3         = Request("txtJUSYO3")            ':�Z��3
    m_iKenCd          = Request("txtKenCd")             ':���R�[�h
    m_iSityoCd        = Request("txtSityoCd")           ':�s�����R�[�h
    m_sDENWABANGO    = Request("txtDENWABANGO")     ':�d�b�ԍ�
    m_sSINRO_URL     = Request("txtSINRO_URL")      ':URL
    If m_sSINRO_URL="�@" Then m_sSINRO_URL=""
    m_sSinroCD = Request("txtSinroCD")          ':�i�H�敪
    m_sSingakuCD = Request("txtSingakuCD")          ':�i�w�敪

    If m_sSingakuCD ="" Then m_sSingakuCD = 0   '�R���{���I����

    m_sDate = gf_YYYY_MM_DD(date(),"/")
    m_sKyokanCD = Session("KYOKAN_CD")          ':���[�U�[ID
    m_iNendo = Request("txtNendo")              ':�N�x
    m_sYubin = Request("txtYUBINBANGO")              ':�X�֔ԍ�
    m_iGyosyu_Kbn = Request("txtGYOSYU_KBN")              ':�Ǝ�敪
    m_iSihonkin = gf_SetNull2Zero(Request("txtSIHONKIN"))              ':���{���i�P�ʁF���~�j
    if m_iSihonkin <> "" Then
        m_iSihonkinY = m_iSihonkin & "0000"              ':���{���i�P�ʁF�~�j
    end if
    m_iJyugyoin_Suu = gf_SetNull2Zero(Request("txtJYUGYOIN_SUU"))              ':�]�ƈ���
    m_iSyoninkyu = gf_SetNull2Zero(Request("txtSYONINKYU"))              ':���C��
    m_sBiko = Request("txtBIKO")              ':���l

    m_sSinroCD2 = Request("txtSinroCD2")        ':�i�H�R�[�h�i�u�߂�v���Ɏg�p�j
    m_sSingakuCD2 = Request("txtSingakuCD2")    ':�i�w�R�[�h�i�u�߂�v���Ɏg�p�j

    m_sSyusyokuName = Request("txtSyusyokuName")            ':�A�E�於�́i�ꕔ�j

End Sub

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
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    function gonext() {
    
<%
If m_sMode = "Syusei" Then
    response.write "window.alert('" & C_TOUROKU_OK_MSG & "');"
Else
    response.write "window.alert('" & C_TOUROKU_OK_MSG & "');"
End If
%>
            document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>

<body bgcolor="#ffffff" onLoad="gonext()">
<center>
<form name="frm" action="./default.asp" target="<%=C_MAIN_FRAME%>" method=post>
<input type="hidden" name="txtMode" value="search">
<input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtPageCD" value="<%= m_iPageCD %>">
</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>