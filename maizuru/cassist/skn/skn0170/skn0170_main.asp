<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������Ԋ�(�N���X��)
' ��۸���ID : skn/skn0170/skn0170_main.asp
' �@      �\: MAIN�y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:   NENDO           '//�N�x
'               KYOKAN_CD       '//����CD
'               cboGakunenCd    '//�w�N
'               cboClassCd      '//�N���X
'               cboSikenKbn     '//�����敪
'               cboSikenCd      '//����CD
' ��      �n:
' ��      ��:
'           �������\��
'               �󔒃y�[�W��\��
'           ���\���{�^���������ꂽ�ꍇ
'               ���������ɂ��Ȃ����������Ԋ���\��
'-------------------------------------------------------------------------
' ��      ��: 2001/07/19 �ɓ����q
' ��      �X: 2001/08/10 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////

    Public Const C_TIMES_1COL = 5   '//1COLSPAN������̎���(��)
    Public Const C_WIDTH_1COL = 9   '//1COLSPAN�������TD��WIDTH
    Public Const C_TD_PADDING = 5   '//TD�̗]�� '2001/08/10 �ǉ�

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '//�N�x
    'Public m_iKyokanCd         '//��������
    Public m_iGakunen           '//�w�N
    Public m_iClassNo           '//�N���XNO
    Public m_iSikenKbn          '//�����敪
    Public m_sSikenCd           '//����CD
    Public m_sClassName         '//�N���X����
    Public m_sGakkaCd           '//�w��CD

    Public m_sSikenName         '//��������
    Public m_sJiWari_Syuryo_Max '//�����I�������̍ő厞��
    Public m_sJiGen_Syuryo_Max  '//�����I�������̍ő厞��
    Public m_sJiGen_Kaisi_Min   '//�����J�n�����̍ŏ�����

    'ں��ރZ�b�g
    Public m_Rs_Jigen           '//����ں��޾��
    Public m_Rs_Jiwari          '//���Ԋ�ں��޾��

    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
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
    w_sMsgTitle="�������Ԋ�(�N���X��)"
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

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '//�l�̏�����
        Call s_ClearParam()

        '//�ϐ��Z�b�g
        Call s_SetParam()

'//�f�o�b�O
'Call s_DebugPrint()

        '//�\������(�N���X)���擾
        w_iRet = f_GetDisp_Data_Class()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//�\������(����)���擾
        w_iRet = f_GetDisp_Data_Siken()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//�������̎擾
        w_iRet = f_GetJigen()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//�������̂����A�ł��x���I��鎞�Ԃ��擾
        w_iRet = f_GetJigen_Max()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '// �������Ԋ��̎擾 
        w_iRet = f_GetSikenJkanwari()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//�������Ԋ��f�[�^�̂����A�ł��x���I��鎎�����Ԃ��擾
        w_iRet = f_GetSiken_Max()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
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

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Rs_Jigen)
    Call gf_closeObject(m_Rs_Jiwari)

    '// �I������
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    'm_iKyokanCd = ""
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_sClassMei = ""
    m_iSikenKbn = ""
    m_sSikenCd  = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")
    'm_iKyokanCd = Session("KYOKAN_CD")
    m_iGakunen  = Request("cboGakunenCd")   '//�w�N
    m_iClassNo  = Request("cboClassCd")     '//�N���X
    m_iSikenKbn = Request("cboSikenKbn")    '//�����敪
    m_sSikenCd  = Request("cboSikenCd")     '//����CD

    If trim(m_sSikenCd) = "" Or trim(m_sSikenCd) = "@@@" Then
        m_sSikenCd = "0"
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    'response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_iSikenKbn = " & m_iSikenKbn & "<br>"
    response.write "m_sSikenCd =  " & m_sSikenCd & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDisp_Data_Class()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Class = 1

    Do
        '�N���X�}�X�^���f�[�^���擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_CLASSMEI"
        w_sSql = w_sSql & vbCrLf & "  ,M05_CLASS.M05_GAKKA_CD"
        w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN= " & m_iGakunen
        w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_CLASSNO= "   & m_iClassNo

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetDisp_Data_Class = 99
            Exit Do
        End If

        m_sClassName = ""
        If rs.EOF = False Then
            m_sClassName = rs("M05_CLASSMEI")
            m_sGakkaCd = rs("M05_GAKKA_CD")
        End If

        '//����I��
        f_GetDisp_Data_Class = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �\������(����)���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDisp_Data_Siken()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Siken = 1

    Do
        '�����}�X�^���f�[�^���擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_iSikenKbn

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetDisp_Data_Siken = 99
            Exit Do
        End If

        m_sSikenName = ""
        If rs.EOF = False Then
            m_sSikenName = rs("M01_SYOBUNRUIMEI")

            '//���͎����܂��́A�ǎ����I�����ꂽ�ꍇ�����ڍז����ǉ��\��
            If cint(m_sSikenCd) <> 0  Then
                m_sSikenName = m_sSikenName & " (" 
                m_sSikenName = m_sSikenName & rs("M27_SIKENMEI")
                m_sSikenName = m_sSikenName & " )" 
            End If

        End If

        '//����I��
        f_GetDisp_Data_Siken = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �������̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetJigen()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen = 1

    Do
        '���������}�X�^���{�N�x�̎����������擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M26_JIGEN,"
        w_sSql = w_sSql & vbCrLf & "  M26_KAISI_JIKOKU,"
        w_sSql = w_sSql & vbCrLf & "  M26_SYURYO_JIKOKU"
        w_sSql = w_sSql & vbCrLf & " FROM M26_SIKEN_JIGEN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M26_NENDO = " & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & " ORDER BY "
        w_sSql = w_sSql & vbCrLf & "  M26_JIGEN "

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs_Jigen, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetJigen = 99
            Exit Do
        End If

        '//����I��
        f_GetJigen = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �{�N�x�̎��������̍ŏI���Ԃƍŏ����Ԃ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetJigen_Max()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen_Max = 1

    Do
        '���������}�X�^���{�N�x�̎����������擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  MIN(M26_KAISI_JIKOKU) AS MIN_KAISI_JIKOKU,"
        w_sSql = w_sSql & vbCrLf & "  MAX(M26_SYURYO_JIKOKU) AS MAX_SYURYO_JIKOKU"
        w_sSql = w_sSql & vbCrLf & " FROM M26_SIKEN_JIGEN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M26_NENDO = " & m_iSyoriNen

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetJigen_Max = 99
            Exit Do
        End If

        m_sJiGen_Syuryo_Max = ""

        If rs.EOF = False Then
            m_sJiGen_Kaisi_Min =  rs("MIN_KAISI_JIKOKU")
            m_sJiGen_Syuryo_Max = rs("MAX_SYURYO_JIKOKU")
        End If

        '//����I��
        f_GetJigen_Max = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �������Ԋ��̎擾 
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSikenJkanwari()
    Dim w_iRet
    Dim w_sSQL

    On Error Resume Next
    Err.Clear

    f_GetSikenJkanwari = 1

    Do

        '�������Ԋ��e�[�u����莞�Ԋ������擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  A.T26_SIKENBI, "
        w_sSql = w_sSql & vbCrLf & "  A.T26_KAMOKU, "
        w_sSql = w_sSql & vbCrLf & "  B.M04_KYOKANMEI_SEI AS JISSI_KYOKAN, "
        w_sSql = w_sSql & vbCrLf & "  C.M04_KYOKANMEI_SEI AS KANTOKU_KYOKAN," 
        'w_sSql = w_sSql & vbCrLf & "  D.M03_KAMOKUMEI, "
        w_sSql = w_sSql & vbCrLf & "  E.M06_KYOSITUMEI, "
        w_sSql = w_sSql & vbCrLf & "  A.T26_SIKEN_JIKAN, "
        w_sSql = w_sSql & vbCrLf & "  A.T26_KAISI_JIKOKU, "
        w_sSql = w_sSql & vbCrLf & "  A.T26_SYURYO_JIKOKU"
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "   T26_SIKEN_JIKANWARI A"
        w_sSql = w_sSql & vbCrLf & "  ,M04_KYOKAN B"
        w_sSql = w_sSql & vbCrLf & "  ,M04_KYOKAN C"
        'w_sSql = w_sSql & vbCrLf & "  ,M03_KAMOKU D"
        w_sSql = w_sSql & vbCrLf & "  ,M06_KYOSITU E"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      A.T26_JISSI_KYOKAN = B.M04_KYOKAN_CD(+) "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_KANTOKU_KYOKAN = C.M04_KYOKAN_CD(+)"
        'w_sSql = w_sSql & vbCrLf & "  AND A.T26_KAMOKU = D.M03_KAMOKU_CD(+)"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_KYOSITU = E.M06_KYOSITU_CD(+)"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=B.M04_NENDO(+) "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=C.M04_NENDO(+) "
        'w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=D.M03_NENDO(+) "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=E.M06_NENDO(+) "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_SIKEN_KBN=" & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_SIKEN_CD='" & m_sSikenCd & "' "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_GAKUNEN=" & m_iGakunen
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_CLASS=" & m_iClassNo
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_JISSI_FLG=" & C_SIKEN_KBN_JISSI
        '//�f�[�^���s���S�Ȃ��͎̂擾���Ȃ�(���{���t�E���{���ԁE�J�n���ԁE���{�����E�ē����̂ǂꂩ�ЂƂł������ĂȂ����͕̂\�����Ȃ�)
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_SIKENBI IS NOT NULL"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_KAISI_JIKOKU IS NOT NULL"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_SYURYO_JIKOKU IS NOT NULL"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_JISSI_KYOKAN IS NOT NULL"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_KANTOKU_KYOKAN IS NOT NULL "
        w_sSql = w_sSql & vbCrLf & " ORDER BY "
        w_sSql = w_sSql & vbCrLf & "  T26_SIKENBI,T26_KAISI_JIKOKU "

'response.write w_sSQL & "<br>"
'response.end

        iRet = gf_GetRecordset_OpenStatic(m_Rs_Jiwari,w_sSQL)
        If iRet <> 0  Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSikenJkanwari = 99
            Exit Do
        End If
        '//����I��
        f_GetSikenJkanwari = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �������Ԋ��f�[�^�̂����A�ł��x���I��鎎�����Ԃ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSiken_Max()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetSiken_Max = 1

    Do

        '//�ł��x���I��鎎�����Ԃ��擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  Max(T26_SYURYO_JIKOKU) AS MAX_SYURYO_JIKOKU"
        w_sSql = w_sSql & vbCrLf & " FROM T26_SIKEN_JIKANWARI"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      T26_SIKEN_JIKANWARI.T26_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_KBN=" & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_CD='" & m_sSikenCd & "' "
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_GAKUNEN=" & m_iGakunen
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_CLASS=" & m_iClassNo
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_JISSI_FLG=" & C_SIKEN_KBN_JISSI

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSiken_Max = 99
            Exit Do
        End If

        m_sJiWari_Syuryo_Max = ""

        If rs.EOF = False Then
            m_sJiWari_Syuryo_Max = rs("MAX_SYURYO_JIKOKU")
        End If

        '//����I��
        f_GetSiken_Max = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �Ȗږ����擾
'*  [����]  p_sKamokuCd
'*  [�ߒl]  f_GetKamokName
'*  [����]  
'********************************************************************************
Function f_GetKamokuName(p_sKamokuCd)
    Dim w_iRet
    Dim w_sSQL
    Dim rs
    Dim w_sKamokuName

    On Error Resume Next
    Err.Clear

    w_sKamokuName = ""

    Do

        '//�Ȗږ����擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI"
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  T15_RISYU"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & m_iSyoriNen - m_iGakunen + 1
        w_sSql = w_sSql & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "' "
        w_sSql = w_sSql & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_sKamokuCd & "'"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            Exit Do
        End If

        If rs.EOF = False Then
            w_sKamokuName = rs("T15_KAMOKUMEI")
        End If

        '//�ߒl���
        f_GetKamokuName = w_sKamokuName

        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  ���Ԃ��COLSPAN���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  1Colspan/5���Ƃ���
'********************************************************************************
Function f_Get_Colspan(p_sStartTime,p_sEndTime)
    Dim w_iTime
    Dim w_iColSpan
    On Error Resume Next
    Err.Clear

    w_iTime = 0
    w_iColSpan = 0

    Do
        w_iTime = DateDiff("n", p_sStartTime, p_sEndTime)
        w_iColSpan = w_iTime\C_TIMES_1COL   '//C_TIMES_1COL = 5 (1colspan/5��)5���P��

        Exit Do
    Loop

    Err.Clear
    f_Get_Colspan = w_iColSpan

End Function

'********************************************************************************
'*  [�@�\]  ���Ԋ����e���Z�b�g
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_SetNaiyo(p_Naiyo,p_Add)
    Dim w_sNaiyo

    w_sNaiyo = ""
    If Trim(gf_SetNull2String(p_Naiyo)) <> "" Then
        w_sNaiyo = "<br>" & p_Naiyo & p_Add
    End If

    f_SetNaiyo = w_sNaiyo

End Function

'********************************************************************************
'*  [�@�\]  ���t��"M��D��(�j��)"�̌`�ɂ���
'*  [����]  p_Date
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
Function f_fmtDate(p_Date)
    Dim w_sDate

    w_sDate = ""

    If gf_SetNull2String(p_Date) <> "" Then
        w_sDate = month(p_Date) & "��"
        w_sDate = w_sDate & day(p_Date) & "��"
        w_sDate = w_sDate & "("
        w_sDate = w_sDate & gf_GetYoubi(Weekday(p_Date))
        w_sDate = w_sDate & ")"
    End If

    f_fmtDate = w_sDate

End Function

'********************************************************************************
'*  [�@�\]  ��TD��\������
'*  [����]  p_STime:����(��)
'*          p_BTime:����(��)
'*          p_Class:TD��class
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetBrankTD(p_STime,p_BTime,p_Class)
Dim w_iColSpan

    '//Colspan���擾
    w_iColSpan = f_Get_Colspan(p_STime, p_BTime)
    If w_iColSpan > 0 Then
        %>
        <!--<td class="<%=p_Class%>" align="center" width="<%=w_iColSpan*C_WIDTH_1COL%>" colspan="<%'=w_iColSpan%>"><font ><br></font></td>-->
        <td class="<%=p_Class%>" align="center" colspan="<%=w_iColSpan%>" ><font ><br></font></td>
        <%
    End If
End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    Dim w_sNaiyo    '//�\�����e
    Dim w_MaxTime   '//�������ԍŏI����
    Dim w_iColSpan  '//COLSPAN
    Dim w_sEndTime  '//�����I������
    Dim w_sDate     '//������
    Dim w_sKaisi    '//�����J�n����

%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>�������Ԋ�(�N���X��)</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">

<%
'//�f�o�b�O
'Call s_DebugPrint()
%>
<center>
<br>
    <%Do%>
        <%
        '//�����f�[�^���Ȃ� �܂��͎��Ԋ���f�[�^���Ȃ��ꍇ
        If m_Rs_Jiwari.EOF = True or m_Rs_Jigen.EOF = True Then 
        %>
        <br><br><span class="msg">�������Ԋ���񂪂���܂���</span>
        <%
            Exit Do
        End If
        %>


        <table class="hyo" border="1" width="400">
            <tr>
                <th class="header" width="80"  align="center" nowrap><font size="2">�N���X</font></th>
                <td class="detail" width="120" align="center" nowrap><font size="2"><%=m_iGakunen & "�N�@" & m_sClassName%></font></td>
                <th class="header" width="80"  align="center" nowrap><font size="2">����</font></th>
                <td class="detail" width="180" align="center" nowrap><font size="2"><%=m_sSikenName%></font></td>
            </tr>
        </table>
        <br>

        <table border="0">
        <tr><td width="10"><br></td></tr>
        <tr><td align="center">

        <!--�w�b�_��-->
        <table class="hyo" border="1" width="100%">
        <%If m_Rs_Jigen.EOF=False Then%>
            <%
            '//===============================
            '//�������Ԃ̍ő�̍ŏI���Ԃ��擾
            '//===============================
            If m_sJiWari_Syuryo_Max >= m_sJiGen_Syuryo_Max Then
                w_MaxTime = m_sJiWari_Syuryo_Max
            Else
                w_MaxTime = m_sJiGen_Syuryo_Max
            End If

            %>
            <tr>
            <td class="header" align="center" colspan="1" nowrap><font color="#ffffff" size="2">���@��</font></td>
            <%

            '//=============
            '//������\��
            '//=============
            Do Until m_Rs_Jigen.EOF
                '===����===
                '//Colspan���擾
                w_iColSpan = f_Get_Colspan(m_Rs_Jigen("M26_KAISI_JIKOKU"), m_Rs_Jigen("M26_SYURYO_JIKOKU"))
                w_sEndTime = m_Rs_Jigen("M26_SYURYO_JIKOKU")
                %>
                <td class="header2" align="center" width="<%=w_iColSpan*C_WIDTH_1COL%>" colspan="<%=w_iColSpan%>" nowrap><img src="../../image/sp.gif" width="<%=w_iColSpan*C_WIDTH_1COL-C_TD_PADDING*2%>" height="1"><br><font color="#ffffff" size="2"><%=m_Rs_Jigen("M26_JIGEN")%></font></td>
                <%
                m_Rs_Jigen.MoveNext
                If m_Rs_Jigen.EOF = False Then
                    '//��TD���Z�b�g
                    Call s_SetBrankTD(w_sEndTime, m_Rs_Jigen("M26_KAISI_JIKOKU"),"header2")
                Else
                    '//��TD���Z�b�g
                    Call s_SetBrankTD(w_sEndTime, w_MaxTime,"header2")
                End If%>

            <%Loop%>
            </tr>
            <tr>
            <td class="header" align="center" colspan="1" nowrap><font size="2" color="#ffffff">���@��</font></td>
            <%

            '//=================
            '//�������Ԃ�\��
            '//=================
            m_Rs_Jigen.MoveFirst
            Do Until m_Rs_Jigen.EOF
                '//Colspan���擾
                w_iColSpan = f_Get_Colspan(m_Rs_Jigen("M26_KAISI_JIKOKU"), m_Rs_Jigen("M26_SYURYO_JIKOKU"))
                w_sEndTime = ""
                w_sEndTime = m_Rs_Jigen("M26_SYURYO_JIKOKU")
                '===����===
                %>
                <td class="header2" align="center" width="<%=w_iColSpan*C_WIDTH_1COL%>" colspan="<%=w_iColSpan%>" nowrap>
                <font size="2" color="#ffffff"><%=gf_SetNull2String(m_Rs_Jigen("M26_KAISI_JIKOKU"))%>�`<%=gf_SetNull2String(m_Rs_Jigen("M26_SYURYO_JIKOKU"))%></font></td>
                <%

                m_Rs_Jigen.MoveNext
                If m_Rs_Jigen.EOF = False Then
                    '//��TD���Z�b�g
                    Call s_SetBrankTD(w_sEndTime, m_Rs_Jigen("M26_KAISI_JIKOKU"),"header2")
                Else
                    '//��TD���Z�b�g
                    Call s_SetBrankTD(w_sEndTime, w_MaxTime,"header2")
                End If
            Loop%>
            <!--</tr>-->
        <%End If%>


        <!--���ו�-->
        <%If m_Rs_Jiwari.EOF = False Then%>

            <%
            Do Until m_Rs_Jiwari.EOF

                '//=================
                '//��������\��
                '//=================
                If w_sDate <> m_Rs_Jiwari("T26_SIKENBI") Then
                    w_sDate = m_Rs_Jiwari("T26_SIKENBI")%>
                    </tr>
                    <tr>
                        <td class="header" align="center" height="35" colspan="1" nowrap><font size="2" color="#ffffff"><%=f_fmtDate(m_Rs_Jiwari("T26_SIKENBI"))%></font></td>
                    <%
                    '//�������Ԃ̍ŏ����Ԃ��A�������Ԃ��x���ꍇ
                    If m_sJiGen_Kaisi_Min < m_Rs_Jiwari("T26_KAISI_JIKOKU") Then
                        '//��TD���Z�b�g
                        Call s_SetBrankTD(m_sJiGen_Kaisi_Min, m_Rs_Jiwari("T26_KAISI_JIKOKU"),"CELL2")
                    End If
                End If

                '//=================
                '//�������e��\��
                '//=================
                '//�\��������e���擾
                w_sNaiyo = f_GetKamokuName(m_Rs_Jiwari("T26_KAMOKU"))
                w_sNaiyo = w_sNaiyo & "(" & m_Rs_Jiwari("T26_SIKEN_JIKAN") & ")"
                w_sNaiyo = w_sNaiyo & f_SetNaiyo(m_Rs_Jiwari("JISSI_KYOKAN"),"(��)") & f_SetNaiyo(m_Rs_Jiwari("KANTOKU_KYOKAN"),"(��)") & f_SetNaiyo(m_Rs_Jiwari("M06_KYOSITUMEI"),"")

                '===============================================
                '//���������ɕʂ̃e�X�g�Ȗڂ������Ă����ꍇ�̍l��
                w_sKaisi = m_Rs_Jiwari("T26_KAISI_JIKOKU")
                w_iMax_Time = m_Rs_Jiwari("T26_SIKEN_JIKAN")
                Do Until m_Rs_Jiwari.EOF
                    m_Rs_Jiwari.MoveNext
                    '//���̃��R�[�h��EOF�łȂ��ꍇ
                    If m_Rs_Jiwari.EOF = False Then
                        '//���t���ς���ĂȂ����ǂ���
                        If w_sDate <> m_Rs_Jiwari("T26_SIKENBI") Then
                            m_Rs_Jiwari.MovePrevious
                            Exit Do
                        Else
                            '//�O�̃��R�[�h�̊J�n���ԂƁA����ں��ނ̊J�n���Ԃ������ꍇ�͓��������ɕʂ̃e�X�g�������Ă���
                            If w_sKaisi = m_Rs_Jiwari("T26_KAISI_JIKOKU") Then

                                '//�ő厞�Ԃ��擾
                                If cint(w_iMax_Time) < cint(m_Rs_Jiwari("T26_SIKEN_JIKAN")) Then
                                    w_iMax_Time = m_Rs_Jiwari("T26_SIKEN_JIKAN")
                                End If

                                '//���e���
                                w_sNaiyo = w_sNaiyo & "<br>-------<br>"
                                w_sNaiyo = w_sNaiyo & f_GetKamokuName(m_Rs_Jiwari("T26_KAMOKU")) & "(" & m_Rs_Jiwari("T26_SIKEN_JIKAN") & ")"
                                w_sNaiyo = w_sNaiyo & f_SetNaiyo(m_Rs_Jiwari("JISSI_KYOKAN"),"(��)") & f_SetNaiyo(m_Rs_Jiwari("KANTOKU_KYOKAN"),"(��)") & f_SetNaiyo(m_Rs_Jiwari("M06_KYOSITUMEI"),"")
                            Else
                                m_Rs_Jiwari.MovePrevious
                                Exit Do
                            End If
                        End If
                    Else
                        m_Rs_Jiwari.MovePrevious
                        Exit Do
                    End If
                Loop
                '===============================================

                '//COLSPAN���擾
                'w_iColSpan = f_Get_Colspan(m_Rs_Jiwari("T26_KAISI_JIKOKU"), m_Rs_Jiwari("T26_SYURYO_JIKOKU"))
                w_iColSpan = cint(w_iMax_Time)\C_TIMES_1COL

                '//�����I���������擾(��TD�ɕK�v)
                w_sEndTime = ""
                w_sEndTime = m_Rs_Jiwari("T26_SYURYO_JIKOKU")
                %>
                <td class="CELL1" width="<%=w_iColSpan*C_WIDTH_1COL%>" colspan="<%=w_iColSpan%>" valign="top"><font size="2"><%=w_sNaiyo%></font></td>

                <%m_Rs_Jiwari.MoveNext
                If m_Rs_Jiwari.EOF = False Then
                    '//���̃��R�[�h�̎��{�����ς�����ꍇ�A�c���TD��ǉ�����
                    If w_sDate <> m_Rs_Jiwari("T26_SIKENBI") Then
                        '//��TD���Z�b�g
                        Call s_SetBrankTD(w_sEndTime, w_MaxTime,"CELL2")
                    Else
                        '//��TD���Z�b�g
                        Call s_SetBrankTD(w_sEndTime, m_Rs_Jiwari("T26_KAISI_JIKOKU"),"CELL2")
                    End If
                Else
                    '//��TD���Z�b�g
                    Call s_SetBrankTD(w_sEndTime, w_MaxTime,"CELL2")
                End If
            Loop
        End If%>

                </tr>
    </table>
    </td></tr>
    </table>

    <%
        Exit Do
    Loop%>

</center>
</body>
</html>
<%
End Sub
%>