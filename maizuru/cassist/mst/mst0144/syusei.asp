<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A�E��}�X�^
' ��۸���ID : mst/mst0144/syusei.asp
' �@      �\: ���y�[�W �i�H�}�X�^�̏ڍוύX���s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtSinroCD      :�i�H�R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSyusyokuName     :�i�H���́i�ꕔ�j
'           txtPageSinro        :�\���ϕ\���Ő��i�������g����󂯎������j
'           RenrakusakiCD       :�I�����ꂽ�i�H�R�[�h
'           txtSchMode          :����Ӱ��
'                                   + JyusyoSch = �Z������
'                                   + ZipSch    = �X�֔ԍ�����
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
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
' ��      ��: 2001/06/26 �≺ �K��Y
' ��      �X: 2001/07/13 �J�e�@�ǖ�
' �@      �@: 2001/07/24 ���{�@�����iDB�ύX�ɔ����C���j
' �@�@�@�@�@: 2001/08/22 �ɓ��@���q�@�Ǝ�敪�ǉ��Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public m_sSinroCD           ':�f�[�^�x�[�X����擾�����i�H�敪
    Public m_sSinroCD3          ':���y�[�W����擾�����i�H�敪
    Public m_sSingakuCD         ':�f�[�^�x�[�X����擾�����i�w�敪
    Public m_sSinromei          ':�f�[�^�x�[�X����擾�����i�H����
    'Public m_sSinromei_Eig     ':�f�[�^�x�[�X����擾�����i�H�p�ꖼ��
    Public m_sSinromei_Kan      ':�f�[�^�x�[�X����擾�����i�H���̃J�i
    Public m_sSinromei_Rya      ':�f�[�^�x�[�X����擾�����i�H����
    Public m_sJusyo             ':�f�[�^�x�[�X����擾�����i�H�Z��
    Public m_sJusyo1            ':�f�[�^�x�[�X����擾�����i�H�Z��1
    Public m_sJusyo2            ':�f�[�^�x�[�X����擾�����i�H�Z��2
    Public m_sJusyo3            ':�f�[�^�x�[�X����擾�����i�H�Z��3
    Public m_sTel               ':�f�[�^�x�[�X����擾�����i�H�d�b�ԍ�
    Public m_sSinro_URL         ':�f�[�^�x�[�X����擾�����i�HURL
    Public m_Rs                 ':recordset
    Public m_sMode              ':���[�h
    Public m_sSchMode           ':����Ӱ��
    Public m_bReFlg             ':�����[�h���ꂽ���ǂ���
    Public m_sRenrakusakiCD     ':Main����擾�����A����CD
    Public m_sPageCD            ':�y�[�W��
    Public m_sSinroCD2          ':Main����擾�����i�H�敪
    Public m_sSingakuCD2        ':Main����擾�����i�w�敪
    Public m_sSyusyokuName      ':Main����擾�����������́i�ꕔ)
    Public m_iNendo             ':�N�x
    Public m_sSubtitle          ':�T�u�^�C�g��
    Public m_iKenCd             ':���R�[�h
    Public m_iSityoCd           ':�s�����R�[�h
    Public m_sYubin             ':�X�֔ԍ�
    'Public m_iGyosyu_Kbn        ':�Ǝ�敪
    Public m_iSihonkin          ':���{���i�P�ʁF���~�j
    Public m_iSihonkinY         ':���{���i�P�ʁF�~�j
    Public m_iJyugyoin_Suu      ':�]�ƈ�
    Public m_iSyoninkyu         ':���C��
    Public m_sBiko              ':���l

    '�i�H��Where����
    Public m_sSinroWhere        ':�i�H�̏���
    Public m_sSingakuWhere      ':�i�H�̏���
    Public m_sSelected1         ':�i�H�R���{�̏���
    Public m_sSelected2         ':�i�w�R���{�̏���
    Public m_sSingakuOption     ':�i�w�R���{�̃I�v�V����

    Public m_sKenWhere          ':���̏���
    Public m_sSityoWhere        ':�s�����R���{�̏���
    Public m_sSityoOption       ':�s�����R���{�̃I�v�V����
    Public m_sKenSentakuWhere
    Public m_sSityoSentakuWhere

    'Public Const C_SYORYAKU_KETA=4'//�\�����ɏȗ����錅���i���{���j


'///////////////////////////���C������/////////////////////////////
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()

    '�l�̏�����
    Call s_Syokika

    '// �ް��ް��ڑ�
    w_iRet = gf_OpenDatabase()
    If w_iRet <> 0 Then
        '�ް��ް��Ƃ̐ڑ��Ɏ��s
        m_bErrFlg = True
        m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
        Exit sub
    End If  

    '// �s���A�N�Z�X�`�F�b�N
    Call gf_userChk(session("PRJ_No"))

    '�p�����[�^�Z�b�g
    Call s_ParaSet()

    '// �����{�^���������ꂽ�Ƃ��́ADB����f�[�^�擾
    If m_sMode = "Syusei" and m_bReFlg = false Then
        call db_get()
    End If

    If m_sSchMode = "JyusyoSch" then
        Call f_SchJyusyo()
    End if

    '���Ɋւ���WHRE���쐬����
    Call f_MakeKenWhere()   

    '�s�����Ɋւ���WHRE���쐬����
    Call f_MakeSityoWhere() 

    '// �R���{�pwhere���쐬
    Call f_MakeCommbo()

    'HTML���쐬����
    Call showPage()

    '// �I������
    call gf_closeObject(m_Rs)
    call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  DB�Ńf�[�^�̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub db_get()

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�i�H�}�X�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    Do

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
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_YUBIN_BANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINGAKU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_GYOSYU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SIHONKIN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JYUGYOIN_SUU "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SYONINKYU "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_BIKO "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & vbCrLf & "    AND M32_SINRO_CD = '" & m_sRenrakusakiCD & "' "

        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If

        '//���R�[�h�Z�b�g��ϐ��ɓ����
        Call s_Dataset()

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
        response.end
    End If
    

End Sub

'********************************************************************************
'*  [�@�\]  DB�Ńf�[�^�̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub f_SchJyusyo()

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�i�H�}�X�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & "     M12_SITYOSONMEI,  "
        w_sSQL = w_sSQL & "     M12_TYOIKIMEI "
        w_sSQL = w_sSQL & "FROM  "
        w_sSQL = w_sSQL & "     M12_SITYOSON "
        w_sSQL = w_sSQL & "WHERE "
        w_sSQL = w_sSQL & "     M12_YUBIN_BANGO = '" & Request("txtYUBINBANGO") & "'"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If

        Exit Do
    Loop

    if Not w_Rs.Eof then
        m_sJusyo1 = w_Rs("M12_SITYOSONMEI")
        m_sJusyo2 = w_Rs("M12_TYOIKIMEI")
        m_sJusyo3 = ""
    End if

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
        response.end
    End If
    
End Sub

'///////////////////////////�l�̏�����/////////////////////////////
Sub s_Syokika()

    m_sSinroCD          = ""
    m_sSingakuCD        = ""
    m_sSinromei         = ""
    'm_sSinromei_Eig    = ""
    m_sSinromei_Kan     = ""
    m_sSinromei_Rya     = ""
    'm_sJusyo           = ""
    m_sJusyo1           = ""
    m_sJusyo2           = ""
    m_sJusyo3           = ""
    m_sTel              = ""
    m_sSinro_URL        = ""
    m_sPageCD           = ""
    m_Rs                = ""
    m_sMode             = ""
    m_sSinroWhere       = ""
    m_sSingakuWhere     = ""
    m_sSelected1        = ""
    m_sSelected2        = ""
    m_sSingakuOption    = ""
    m_sSinroCD2         = ""
    m_sSingakuCD2       = ""
    m_iKenCd            = ""
    m_iSityoCd          = ""
    m_sYubin            = ""
    'm_iGyosyu_Kbn       = ""
    m_iSihonkin         = ""
    m_iSihonkinY        = ""
    m_iJyugyoin_Suu     = ""
    m_iSyoninkyu        = ""
    m_sBiko             = ""

End Sub


'********************************************************************************
'*  [�@�\]  ������ϐ��ɑ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ParaSet()

    m_sMode     = Request("txtMode")
    If m_sMode  = "" then m_sMode = "Sinki"

    m_sSchMode  = Request("txtSchMode")                     ':����Ӱ��

    m_bReFlg    = Request("txtReFlg")
    If m_bReFlg = "" then m_bReFlg = false

    m_sRenrakusakiCD = Request("txtRenrakusakiCD")          ':�A����R�[�h

    '/*�d�l�ύX�ɂ��g�p���Ȃ�
    'If m_sMode = "Sinki" Then
        'Call f_Max()
    'End If

    'm_sSinroCD3    = Request("txtSinroCD")                 ':�i�H�R�[�h
    m_sSinroCD      = gf_cboNull(Request("txtSinroCD"))     ':�i�H�R�[�h
    m_sSingakuCD    = gf_cboNull(Request("txtSingakuCD"))   ':�i�w�R�[�h
    m_sSyusyokuName = Request("txtSyusyokuName")            ':�A�E�於�́i�ꕔ�j

    m_sSinroCD2     = Request("txtSinroCD2")                ':�u�߂�v�p�̐i�H�R�[�h
    if m_sSinroCD   = "" then Request("txtSinroCD")

    m_sSingakuCD2   = Request("txtSingakuCD2")              ':�u�߂�v�p�̐i�w�R�[�h
    if m_sSingakuCD2= "" then Request("txtSingakuCD")

    m_sSINROMEI     = Request("txtSINROMEI")                ':�i�H��
    'm_sSinromei_Eig= Request("txtSINROMEI_EIGO")           ':�i�H���p��
    m_sSinromei_Kan = Request("txtSINROMEI_KANA")           ':�i�H���J�i
    m_sSinromei_Rya = Request("txtSINRORYAKSYO")            ':�i�H����
    'm_sJusyo       = Request("txtJUSYO")                   ':�Z��

    If m_sSchMode = "" then
        m_sJusyo1       = Request("txtJUSYO1")              ':�Z��1
        m_sJusyo2       = Request("txtJUSYO2")              ':�Z��2
        m_sJusyo3       = Request("txtJUSYO3")              ':�Z��3
    End if

    m_sTel          = Request("txtDENWABANGO")              ':�d�b�ԍ�
    m_iNendo        = Session("NENDO")                      ':�N�x

    m_iKenCd        = Request("txtKenCd")                   ':���R�[�h
    m_iSityoCd      = Request("txtSityoCd")                 ':�s�����R�[�h
    m_sYubin        = Request("txtYUBINBANGO")              ':�X�֔ԍ�
    'm_iGyosyu_Kbn   = Request("txtGYOSYU_KBN")              ':�Ǝ�敪
    m_iSihonkin     = Request("txtSIHONKIN")                ':���{��
    m_iJyugyoin_Suu = Request("txtJYUGYOIN_SUU")            ':�]�ƈ���
    m_iSyoninkyu    = Request("txtSYONINKYU")               ':���C��
    m_sBiko         = Request("txtBIKO")                    ':���l

    m_sSinro_URL    = Request("txtSINRO_URL")               ':URL
    if m_sSinro_URL = "" Then m_sSinro_URL = "http://"

    '//BLANK�̏ꍇ�͍s���ر
    If m_sMode = "Sinki" Then
        m_sPageCD = 1
    Else
        m_sPageCD = INT(Request("txtPageCD"))               ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If

End Sub

'********************************************************************************
'*  [�@�\]  DB�̒l��ϐ��ɑ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_Dataset()

Dim w_iSihonkin

    m_sSinroCD      = m_Rs("M32_SINRO_KBN")

	If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
		'//�i�H��CD��1(�i�w)�̏ꍇ�A�i�w�敪���擾
    	m_sSingakuCD    = gf_SetNull2String(m_Rs("M32_SINGAKU_KBN"))
	ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
		'//�i�H��CD��2(�A�E)�̏ꍇ�A�Ǝ�敪���擾
    	m_sSingakuCD    = gf_SetNull2String(m_Rs("M32_GYOSYU_KBN"))
	End If

    m_sSinromei     = m_Rs("M32_SINROMEI")
    'm_sSinromei_Eig = m_Rs("M32_SINROMEI_EIGO")
    m_sSinromei_Kan = m_Rs("M32_SINROMEI_KANA")
    m_sSinromei_Rya = m_Rs("M32_SINRORYAKSYO")
    'm_sJusyo        = m_Rs("M32_JUSYO")
    m_sJusyo1        = m_Rs("M32_JUSYO1")
    m_sJusyo2        = m_Rs("M32_JUSYO2")
    if IsNull(m_Rs("M32_JUSYO3")) = False Then
        m_sJusyo3        = m_Rs("M32_JUSYO3")
    end if
    m_sTel          = m_Rs("M32_DENWABANGO")
    m_iKenCd        = m_Rs("M32_KEN_CD")
    m_iSityoCd        = m_Rs("M32_SITYOSON_CD")
    m_sYubin = m_Rs("M32_YUBIN_BANGO")

'    if IsNull(m_Rs("M32_GYOSYU_KBN")) = False Then
'        m_iGyosyu_Kbn = m_Rs("M32_GYOSYU_KBN")
'    end if

    if IsNull(m_Rs("M32_SIHONKIN")) = False Then
        m_iSihonkinY = m_Rs("M32_SIHONKIN")
        w_iSihonkin = CInt(Len(m_iSihonkinY)) - C_SYORYAKU_KETA
        m_iSihonkin = Mid(m_iSihonkinY,1,w_iSihonkin)
    end if
    if IsNull(m_Rs("M32_JYUGYOIN_SUU")) = False Then
        m_iJyugyoin_Suu = m_Rs("M32_JYUGYOIN_SUU")
    end if
    if IsNull(m_Rs("M32_SYONINKYU")) = False Then
        m_iSyoninkyu = m_Rs("M32_SYONINKYU")
    end if
    if IsNull(m_Rs("M32_BIKO")) = False Then
        m_sBiko = m_Rs("M32_BIKO")
    end if

    if IsNull(m_Rs("M32_SINRO_URL")) = False Then
        m_sSinro_URL = m_Rs("M32_SINRO_URL")
    else
        m_sSinro_URL = "http://"
    end if
    
    'if m_sSinro_URL = "" Then m_sSinro_URL = "http://"
    
End Sub


Sub f_MakeKenWhere()
'********************************************************************************
'*  [�@�\]  ���R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sKenWhere=""
    m_sKenSentakuWhere=""
        m_sKenWhere = " M16_NENDO = '" & Session("NENDO") & "' "
        'm_sKenSentakuWhere = Request("txtKenCd")
        m_sKenSentakuWhere = m_iKenCd
End Sub

Sub f_MakeSityoWhere()
'********************************************************************************
'*  [�@�\]  �s�����R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sSityoWhere=""
    m_sSityoSentakuWhere = ""
    m_sSityoOption=""

    'If Request("txtKenCd") <> "" Then
    If m_iKenCd <> "" Then
        'm_sSityoWhere = "     M12_KEN_CD = '" & Request("txtKenCd") & "' "
        m_sSityoWhere = "     M12_KEN_CD = '" & m_iKenCd & "' "
        m_sSityoWhere = m_sSityoWhere & " GROUP BY M12_SITYOSON_CD,M12_SITYOSONMEI "
        m_sSityoSentakuWhere = m_iSityoCd
    Else
        m_sSityoOption = " DISABLED "
        m_sSityoWhere  = " M12_Ken_CD = '0' "
    End IF

End Sub

'********************************************************************************
'*  [�@�\]  �i�H�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub f_MakeCommbo()


    m_sSinroWhere = " M01_DAIBUNRUI_CD = "&C_SINRO&"  AND "
    m_sSinroWhere = m_sSinroWhere & " M01_NENDO = " & m_iNendo & ""

    If m_sMode = "Insert" Then
        m_sSingakuOption = "DISABLED"
        m_sSelected1      = ""
        m_sSelected2      = ""

    Else 

        m_sSelected1 = m_sSinroCD

		'// �i�w
	    If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
	        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_SINGAKU & "  AND "
	        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
            m_sSelected2 = m_sSingakuCD

		'// �A�E
		ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
	        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_GYOSYU_KBN & "  AND "
	        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
            m_sSelected2 = m_sSingakuCD

		'// ���̑�
	    Else
	        m_sSingakuWhere= " M01_DAIBUNRUI_CD = 0 "
	        m_sSingakuOption = " DISABLED "
	    End IF


    End If

End Sub

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
    //  [�@�\]  �i�H���C�����ꂽ�Ƃ��A�ĕ\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./syusei.asp";
        document.frm.target="";

        document.frm.txtReFlg.value ='true';
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �Z�������{�^�������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  true:����OK�Afalse:�����װ
    //  [����]  
    //************************************************************
    function jf_JyusyoSch(){

        if( f_Trim(document.frm.txtYUBINBANGO.value) == "" ){
            window.alert("�X�֔ԍ������͂���Ă��܂���");
            document.frm.txtYUBINBANGO.focus();
            return false;
        }

        // ����������������������
        // ���i�H��X�֔ԍ�
        var str = new String(document.frm.txtYUBINBANGO.value);
        if( getLengthB(str) != "8" ){
            window.alert("�X�֔ԍ���8�������͂��Ă�������");
            document.frm.txtYUBINBANGO.focus();
            return false;
        }

        var str = new String(document.frm.txtYUBINBANGO.value);
        if( f_Trim(str) != "" ){
          if( IsHankakuSujiHyphen(str) == false ){
              window.alert("�X�֔ԍ��͔��p�����ƃn�C�t���̂ݓ��͂��Ă�������");
              document.frm.txtYUBINBANGO.focus();
              return false;
          }
        }

        document.frm.txtSchMode.value = "JyusyoSch";
        document.frm.action="./syusei.asp";
        document.frm.target="fTopMain";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �X�֔ԍ������{�^�������ꂽ�Ƃ�
    //  [����]  pMode = Ӱ��
    //                  'SEARCH' = ����
    //                  'DISPLAY'= �Q��
    //  [�ߒl]  true:����OK�Afalse:�����װ
    //  [����]  
    //************************************************************
    function jf_ZipCodeSch(pMode){
        var w_JUSYO1 = ""
        var w_JUSYO2 = ""


        // ����Ӱ�ނ̏ꍇ�A�Z�����K�v
        if( pMode == 'SEARCH' ){
            if( f_Trim(document.frm.txtJUSYO1.value) == "" && f_Trim(document.frm.txtJUSYO2.value) == "" ){
                window.alert("�Z������͂��Ă�������");
                document.frm.txtJUSYO1.focus();
                return false;
            }
            w_JUSYO1 = document.frm.txtJUSYO1.value;
            w_JUSYO2 = document.frm.txtJUSYO2.value;
        }

        // �T�u�E�B���h�[���J��
        w   = 520;
        h   = 520;
        url = "../../Common/com_select/SEL_JYUSYO/default.asp";
        wn  = "SubWindow";
        opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=yes";
        if (w > 0)
            opt = opt + ",width=" + w;
        if (h > 0)
            opt = opt + ",height=" + h;
        newWin = window.open(url, wn, opt);

		// �����̏ꍇ�́A�Z���𑗂�i���������j
		if ( pMode == 'SEARCH' ){
	        document.frm.action="../../Common/com_select/SEL_JYUSYO/default.asp";
	        document.frm.target="SubWindow";
	        document.frm.submit();
		}

        // window�ړ�
        x   = (screen.availWidth - w) / 2;
        y   = (screen.availHeight - h) / 2;
        newWin.moveTo(x, y);

    }


    //************************************************************
    //  [�@�\]  ���͒l������
    //  [����]  �Ȃ�
    //  [�ߒl]  0:����OK�A1:�����װ
    //  [����]  ���͒l��NULL�����A�p���������A�����������s��
    //          ���n�ް��p���ް������H����K�v������ꍇ�ɂ͉��H���s��
    //************************************************************
    function f_CheckData() {
    
        // ������NULL����������
        // ���A����R�[�h
        if( f_Trim(document.frm.txtRenrakusakiCD.value) == "" ){
            window.alert("�A����R�[�h�����͂���Ă��܂���");
            document.frm.txtRenrakusakiCD.focus();
            return false;
        }

        // ���������l�Ó�������������
        // ���A����R�[�h���l
        if( isNaN(document.frm.txtRenrakusakiCD.value) ){
            window.alert("�A����R�[�h�ɂ͐��l����͂��Ă�������");
            document.frm.txtRenrakusakiCD.focus();
            return false;
        }

        // ������NULL����������
        // ������
        if( f_Trim(document.frm.txtSINROMEI.value) == "" ){
            window.alert("���̂����͂���Ă��܂���");
            document.frm.txtSINROMEI.focus();
            return false;
        }

        // ����������������������
        // ������
        var str = new String(document.frm.txtSINROMEI.value);
        if( getLengthB(str) > "60" ){
            window.alert("���̂͑S�p30�����œ��͂��Ă�������");
            document.frm.txtSINROMEI.focus();
            return false;
        }

        // ����������������������
        // ������
        var str = new String(document.frm.txtSINROMEI_KANA.value);
        if( getLengthB(str) > "60" ){
            window.alert("���̂͑S�p30�����ȓ��œ��͂��Ă�������");
            document.frm.txtSINROMEI_KANA.focus();
            return false;
        }

        // ����������������������
        // ������
        var str = new String(document.frm.txtSINRORYAKSYO.value);
        if( getLengthB(str) > "10" ){
            window.alert("���̂͑S�p5�����ȓ��œ��͂��Ă�������");
            document.frm.txtSINRORYAKSYO.focus();
            return false;
        }

        // ������NULL����������
        // ���i�H�敪
        if( f_Trim(document.frm.txtSinroCD.value) == "@@@" ){
            window.alert("�i�H�敪�����͂���Ă��܂���");
            document.frm.txtSinroCD.focus();
            return false;
        }

        if(document.frm.txtSinroCD.value == '1') {

                // ������NULL����������
                // ���i�w�敪
                if( f_Trim(document.frm.txtSingakuCD.value) == "@@@" ){
                    window.alert("�i�w�敪�����͂���Ă��܂���");
                    document.frm.txtSingakuCD.focus();
                    return false;
                }
        };


        // ����������������������
        // ���i�H��X�֔ԍ�
        var str = new String(document.frm.txtYUBINBANGO.value);
        if( getLengthB(str) != "8" ){
            window.alert("�X�֔ԍ���8�������͂��Ă�������");
            document.frm.txtYUBINBANGO.focus();
            return false;
        }

        var str = new String(document.frm.txtYUBINBANGO.value);
        if( f_Trim(str) != "" ){
          if( IsHankakuSujiHyphen(str) == false ){
              window.alert("�X�֔ԍ��͔��p�����ƃn�C�t���̂ݓ��͂��Ă�������");
              document.frm.txtYUBINBANGO.focus();
              return false;
          }
        }
        // ������NULL����������
        // �����R�[�h
//        if( f_Trim(document.frm.txtKenCd.value) == "@@@" ){
//          window.alert("�s���{�����I������Ă��܂���");
//          document.frm.txtKenCd.focus();
//          return false;
//        }

        // ������NULL����������
        // ���Z���i�P�j
        if( f_Trim(document.frm.txtJUSYO1.value) == "" ){
            window.alert("�Z���i�P�j�����͂���Ă��܂���");
            document.frm.txtJUSYO1.focus();
            return false;
        }

        // ����������������������
        // ���Z���i�P�j
        var str = new String(document.frm.txtJUSYO1.value);
        if( getLengthB(str) > "40" ){
            window.alert("�Z���͑S�p20�����ȓ��œ��͂��Ă�������");
            document.frm.txtJUSYO1.focus();
            return false;
        }

        // ������NULL����������
        // ���Z���i�Q�j
        if( f_Trim(document.frm.txtJUSYO2.value) == "" ){
            window.alert("�Z���i�Q�j�����͂���Ă��܂���");
            document.frm.txtJUSYO2.focus();
            return false;
        }
        // ����������������������
        // ���Z���i�Q�j
        var str = new String(document.frm.txtJUSYO2.value);
        if( getLengthB(str) > "40" ){
            window.alert("�Z���͑S�p20�����ȓ��œ��͂��Ă�������");
            document.frm.txtJUSYO2.focus();
            return false;
        }

        // ����������������������
        // ���Z���i�R�j
        var str = new String(document.frm.txtJUSYO3.value);
        if( getLengthB(str) > "40" ){
            window.alert("�Z���͑S�p20�����ȓ��œ��͂��Ă�������");
            document.frm.txtJUSYO3.focus();
            return false;
        }

        // ������NULL����������
        // ���i�H��d�b�ԍ�
        if( f_Trim(document.frm.txtDENWABANGO.value) == "" ){
            window.alert("�d�b�ԍ������͂���Ă��܂���");
            document.frm.txtDENWABANGO.focus();
            return false;
        }
        // ����������������������
        // ���i�H��d�b�ԍ�
        var str = new String(document.frm.txtDENWABANGO.value);
        if( getLengthB(str) > "15" ){
            window.alert("�d�b�ԍ���15�����ȓ��œ��͂��Ă�������");
            document.frm.txtDENWABANGO.focus();
            return false;
        }
        // ��������������������
        var str = new String(document.frm.txtDENWABANGO.value);
        if( f_Trim(str) != "" ){
          if( IsHankakuSujiHyphen(str) == false ){
              window.alert("�d�b�ԍ��͔��p�����ƃn�C�t���̂ݓ��͂��Ă�������");
              document.frm.txtDENWABANGO.focus();
              return false;
          }
        }
    
        // ����������������������
        // ��URL
        var str = new String(document.frm.txtSINRO_URL.value);
        if( f_Trim(str) != "" ){
            if( getLengthB(str) > "40" ){
                window.alert("URL��40�����ȓ��œ��͂��Ă�������");
                document.frm.txtSINRO_URL.focus();
                return false;
            }
        }

<%
'//        // ���������l�Ó�������������
'//        // ���Ǝ�敪
'//        if( f_Trim(document.frm.txtGYOSYU_KBN.value) != "" ){
'//            if( isNaN(document.frm.txtGYOSYU_KBN.value) ){
'//                window.alert("�Ǝ�敪�͐��l�𔼊p�œ��͂��Ă�������");
'//                document.frm.txtGYOSYU_KBN.focus();
'//                return false;
'//            }
'//        }
'//
'//        // ������������������
'//        // ���Ǝ�敪
'//        if( f_Trim(document.frm.txtGYOSYU_KBN.value) != "" ){
'//            var str = new String(document.frm.txtGYOSYU_KBN.value);
'//            if( str.length > 2 ){
'//                window.alert("�Ǝ�敪�̓��͒l��2���ȓ��ɂ��Ă�������");
'//                document.frm.txtGYOSYU_KBN.focus();
'//                return false;
'//            }
'//        }
%>
        // ���������l�Ó�������������
        // �����{��
        if( f_Trim(document.frm.txtSIHONKIN.value) != "" ){
            if( isNaN(document.frm.txtSIHONKIN.value) ){
                window.alert("���{���͐��l�𔼊p�œ��͂��Ă�������");
                document.frm.txtSIHONKIN.focus();
                return false;
            }
        }
        
        // ����������������������
        // �����{��
        var str = new String(document.frm.txtSIHONKIN.value);
        if( getLengthB(str) > "7" ){
            window.alert("���{���͔��p7���ȓ��œ��͂��Ă�������");
            document.frm.txtSIHONKIN.focus();
            return false;
        }
        
        // ���������l�Ó�������������
        // ���]�ƈ���
        if( f_Trim(document.frm.txtJYUGYOIN_SUU.value) != "" ){
            if( isNaN(document.frm.txtJYUGYOIN_SUU.value) ){
                window.alert("�]�ƈ����͐��l�𔼊p�œ��͂��Ă�������");
                document.frm.txtJYUGYOIN_SUU.focus();
                return false;
            }
        }
        
        // ������������������
        // ���]�ƈ���
        var str = new String(document.frm.txtJYUGYOIN_SUU.value);
        if( str.length > 7 ){
            window.alert("�]�ƈ����̓��͒l��7���ȓ��ɂ��Ă�������");
            document.frm.txtJYUGYOIN_SUU.focus();
            return false;
        }
        
        // ���������l�Ó�������������
        // �����C��
        if( f_Trim(document.frm.txtSYONINKYU.value) != "" ){
            if( isNaN(document.frm.txtSYONINKYU.value) ){
                window.alert("���C���͐��l�𔼊p�œ��͂��Ă�������");
                document.frm.txtSYONINKYU.focus();
                return false;
            }
        }
        
        // ������������������
        // �����C��
        var str = new String(document.frm.txtSYONINKYU.value);
        if( str.length > 7 ){
            window.alert("���C���̓��͒l��7���ȓ��ɂ��Ă�������");
            document.frm.txtSYONINKYU.focus();
            return false;
        }
        
        // ������������������
        // �����l
        if( getLengthB(document.frm.txtBIKO.value) > "100" ){
            window.alert("���l�̗��͑S�p50�����ȓ��œ��͂��Ă�������");
            document.frm.txtBIKO.focus();
            return false;
        }

        document.frm.action="./kakunin.asp";
        document.frm.target="_self";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    </head>

    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <div align="center">
<!--    <form name="frm" action="kakunin.asp" target="_self" Method="POST">-->
    <form name="frm" Method="POST">
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
<table border="0" cellpadding="1" cellspacing="1">
    <tr>
        <td align="left">
            <table width="100%" border=1 CLASS="hyo">
	        <colgroup valign="top">
	        <colgroup valign="top">
		        <tr>
		            <th class=header width="100">�i�H��R�[�h</th>
		            <td nowrap class=detail align="left">
		            <% If m_sMode <> "Sinki" Then %>
		                <%=m_sRenrakusakiCD%>
		                <input type="hidden" name="txtRenrakusakiCD" value="<%= m_sRenrakusakiCD %>">
		            <% Else %>
		                <input type="text" name="txtRenrakusakiCD" value="<%= m_sRenrakusakiCD %>" MAXLENGTH=6 size=8>
		                <span class=hissu>*</span>�i���p����6���ȓ��j
		            <% End If %>
		            </td>
		        </tr>
		        <tr>
		            <th class=header>���@��</th>
		            <td nowrap class=detail><input type="text" size="64" name="txtSINROMEI" value="<%= m_sSinromei %>" MAXLENGTH=60 size=30><span class=hissu>*</span>�i�S�p30�����ȓ��j</td>
		        </tr>
		        <tr>
		            <th class=header>���@�́i�J�i�j</th>
		            <td nowrap class=detail><input type="text" size="64" name="txtSINROMEI_KANA" value="<%= m_sSinromei_Kan %>" MAXLENGTH=60 size=30>�i�S�p30�����ȓ��j</td>
		        </tr>
		        <tr>
		            <th class=header>���@��</th>
		            <td nowrap class=detail><input type="text" size="20" name="txtSINRORYAKSYO" value="<%= m_sSinromei_Rya %>" MAXLENGTH=10 size=10>�i�S�p5�����ȓ��j</td>
		        </tr>
		        <tr>
		            <th class=header>�i�H�敪</th>
		            <td nowrap class=detail>
		                <%  '���ʊ֐�����i�H�Ɋւ���R���{�{�b�N�X���o�͂���
		                    call gf_ComboSet("txtSinroCD",C_CBO_M01_KUBUN,m_sSinroWhere,"onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sSelected1)
		                %><span class=hissu>*</span>
		            </td>
		        </tr>
		        <tr>
		            <th class=header>��ʋ敪</th>
		            <td nowrap class=detail>
		                <% '���ʊ֐�����i�w�Ɋւ���R���{�{�b�N�X���o�͂���i�i�H�敪�������j�i�i�H�敪��1�ł͂Ȃ��Ƃ��́ADISABLED�ƂȂ�j
		                    call gf_ComboSet("txtSingakuCD",C_CBO_M01_KUBUN,m_sSingakuWhere,"style='width=100px' "&m_sSingakuOption,True,m_sSelected2)
		                %>

		                <% If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then %>
		                    <span class=hissu>*</span>
		                <% End If %>
		            </td>
		        </tr>
		        <tr>
		            <th class=header>�X�֔ԍ�</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtYUBINBANGO" value="<%= m_sYubin %>" MAXLENGTH=8><span class=hissu>*</span>�i��:000-0000�j
		                <img src="../../image/sp.gif" width="100" height="1">
		                <input type="button" class="button" name="btnJyusyoSch" value="�� �� �Z��" onClick="javascript:return jf_JyusyoSch();">
		                <input type="button" class="button" name="btnZipSch"    value="�Z�� �� ��" onClick="javascript:return jf_ZipCodeSch('SEARCH');">
		            </td>
		        </tr>
		        <tr>
		            <th class=header>�Z�@���i�P�j</th>
		            <td nowrap class=detail><input type="text" size="44" name="txtJUSYO1" value="<%= m_sJusyo1 %>" MAXLENGTH=40 size=40><span class=hissu>*</span>�i�S�p20�����ȓ��j
		                <img src="../../image/sp.gif" width="10" height="1"><input type="button" class="button" name="btnJyusyoDsp" value="�Q��" onClick="javascript:return jf_ZipCodeSch('DISPLAY');"></td>
		        </tr>
		        <tr>
		            <th class=header>�Z�@���i�Q�j</th>
		            <td nowrap class=detail><input type="text" size="44" name="txtJUSYO2" value="<%= m_sJusyo2 %>" MAXLENGTH=40 size=40><span class=hissu>*</span>�i�S�p20�����ȓ��j</td>
		        </tr>
		        <tr>
		            <th class=header>�Z�@���i�R�j</th>
		            <td nowrap class=detail><input type="text" size="44" name="txtJUSYO3" value="<%= m_sJusyo3 %>" MAXLENGTH=40 size=40>�i�����������L���E�S�p20�����ȓ��j</td>
		        </tr>
		        <tr>
		            <th class=header>�d�b�ԍ�</th>
		            <td nowrap class=detail><input type="text" size="20" name="txtDENWABANGO" value="<%= m_sTel %>" MAXLENGTH=15 size=15><span class=hissu>*</span>�i��:000-000-0000�j</td>
		        </tr>
		        <tr>
		            <th class=header>�t�q�k</th>
		            <td nowrap class=detail><input type="text" size="50" name="txtSINRO_URL"  value="<%= m_sSinro_URL %>" MAXLENGTH=40></td>
		        </tr>
<!--
		        <tr>
		            <th class=header>�Ǝ�敪</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtGYOSYU_KBN"  value="<%= m_iGyosyu_Kbn %>" MAXLENGTH=2>�i���p����6���ȓ��j</td>
		        </tr>
-->
		        <tr>
		            <th class=header>���{��</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtSIHONKIN"  value="<%= m_iSihonkin %>" MAXLENGTH="7">���~</td>
		        </tr>
		        <tr>
		            <th class=header>�]�ƈ���</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtJYUGYOIN_SUU"  value="<%= m_iJyugyoin_Suu %>" MAXLENGTH=7>�l</td>
		        </tr>
		        <tr>
		            <th class=header>���C��</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtSYONINKYU"  value="<%= m_iSyoninkyu %>" MAXLENGTH=7>�~</td>
		        </tr>
		        <tr>
		            <th class=header>���@�l</th>
		            <td nowrap class=detail><textarea rows=3 cols=40 name="txtBIKO"><%= m_sBiko %></textarea>�i�S�p50�����ȓ��j</td>
                </TR>
            </TABLE>
		    <table width=100%><tr><td align=right><span class=hissu>*��͕K�{���ڂł��B</span></td></tr></table>
		    <br>
        </td>
    </TR>
	</TABLE>
    <table border="0" >
        <tr>
            <td valign="top" align=left>
                <input type="button" class="button" value="�@�o�@�^�@" Onclick="return f_CheckData()">
                <input type="hidden" name="txtKenCd"   value="<%= m_iKenCd %>">
                <input type="hidden" name="txtSityoCd" value="<%= m_iSityoCd %>">
                <input type="hidden" name="txtRenban"  value="">
                <input type="hidden" name="txtMode" value="<%= m_sMode %>">
                <input type="hidden" name="txtReFlg" value="<%= m_bReFlg %>">
                <input type="hidden" name="txtSinroCD2" value="<%= m_sSinroCD2 %>">
                <input type="hidden" name="txtSingakuCD2" value="<%= m_sSingakuCD2 %>">
                <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
                <!--<input type="hidden" name="txtNendo" value="<%= Session("SYORI_NENDO") %>">-->
                <input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
                <input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
                <input type="hidden" NAME="ButtonClick" value="">
                <input type="hidden" NAME="txtSchMode">
                </form>
            </td>
			<td><img src="../../image/sp.gif" width="20" height="1"></td>
            <td valign="top" align=right>
                <form action=default.asp name="cansel" method=post target="<%=C_MAIN_FRAME%>">
                    <input type="hidden" name="txtMode" value="search">
                    <input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD2 %>">
                    <input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD2 %>">
                    <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
                    <input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
                    <input class=button type='submit' value='�L�����Z��'>
                </form>
            </td>
        </tr>
    </table>

</div>
</body>
</html>


<%
    '---------- HTML END   ----------
End Sub
%>