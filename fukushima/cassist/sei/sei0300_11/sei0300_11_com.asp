<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �l�ʐ��шꗗ�@��p���ʊ֐�
' ��۸���ID : sei/sei0300_11/sei0300_11_com.asp
' �@      �\: ���Ɠ����擾�֐�
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/09/03 �ɓ����q
' ��      �X: 2006/01/27 �����@�ʘa�q�@��������p�ɍ쐬
'*************************************************************************/

'////////////////////////////////////////////////////////////////
'//�R���X�g��`
Private Const C_ZENKIKAISI = 10     '�O���I����
Private Const C_KOUKIKAISI = 11     '����J�n��
Private Const C_KOUKISYURYO = 12    '����I����

Private Const C_NULLGYOJI = 0       '�s���Ȃ�
Private Const C_ALLNEN = 0          '�S�w�N
Private Const C_ALLCLASS = 99       '�S�N���X

Private Const C_NULLJIGEN = 0       '�󎞌�
Private Const C_NULLJIKAN = 0       '�󑍎���
Private Const C_NOZENJITU = 0       '��O���t���O
Private Const C_FLG_TYOKIKYUKA = 1  '�����x�Ƀt���O

Private Const C_SOU_NISSU = True        '�����Ɠ���
Private Const C_JUGYO_NISSU = False     '�����Ɠ���

'////////////////////////////////////////////////////////////////

Public Function gf_SouJugyo(p_lJikan, p_sKCode, p_iNen, p_iClass, p_sKaisibi, p_sSyuryobibi, p_iNendo)
'*******************************************************************************
' �@�@�@�\�F�����ƃf�[�^�̎擾
' �ԁ@�@�l�Ftrue: �����@false: ���s
' ���@�@���Fp_lJikan - �擾�������Ԑ�
' �@�@�@�@�@p_sKCode - �ȖڃR�[�h
' �@�@�@�@�@p_iNen - �w�N
' �@�@�@�@�@p_iClass - �N���X
' �@�@�@�@�@p_sKaisibi - �J�n��
' �@�@�@�@�@p_sSyuryobibi - �I����
' �@�\�ڍׁF�����Ǝ��Ԃ̎擾
' ���@�@�l�F�Ȃ�
'*******************************************************************************
    Dim w_bRtn               '�߂�l
    Dim w_sTyokiWhere         '�����x�ɗp��Where
    Dim w_lJikan                '���Ԑ�
    
    Dim w_sStartDay           '�J�n��
    Dim w_sEndDay             '�I����
    
    'On Error GoTo Err_Func
	On Error Resume Next
	Err.Clear
    
    '== �ϐ��̏����� ==
    gf_SouJugyo = False
    
    w_sTyokiWhere = ""
    w_lJikan = 0
    
    w_sStartDay = ""
    w_sEndDay = ""
    
    '�O���J�n������̏ꍇ�G���[
    If p_sKaisibi < f_GetGakkiDay(C_ZENKIKAISI, p_iNendo) Then
        'Call gf_iMsg(2129)
        Exit Function
    End If

    '����I��������̏ꍇ�G���[
    If p_sSyuryobibi > f_GetGakkiDay(C_KOUKISYURYO, p_iNendo) Then
        'Call gf_iMsg(2129)
        Exit Function
    End If

    '�����x�ɂ̎Z�o
    w_bRtn = f_GetTyokikyuka(w_sTyokiWhere, p_iNendo)
    If w_bRtn <> True Then: Exit Function

    '== �J�n��������J�n�����O(�O��)�̏ꍇ ==
    If p_sKaisibi < f_GetGakkiDay(C_KOUKIKAISI, p_iNendo) Then
        '== �ϐ��̐ݒ� ==
        w_sStartDay = p_sKaisibi
        w_sEndDay = gf_YYYY_MM_DD(DateAdd("d", -1, f_GetGakkiDay(C_KOUKIKAISI, p_iNendo)),"/")

        '== �I�������O���I�������O�̏ꍇ�A�I�����ɕύX ==
        If p_sSyuryobibi < w_sEndDay Then
            w_sEndDay = p_sSyuryobibi
        End If

        '== �O���̎��Ԋ�����擾 ==
        w_bRtn = f_GetJikanWari(C_GAKKI_ZENKI, p_sKCode, p_iNen, p_iClass, w_sStartDay, w_sEndDay, w_sTyokiWhere, C_SOU_NISSU, w_lJikan, p_iNendo)
        If w_bRtn <> True Then
            Exit Function
        End If

        '== �Ȃ����A�I����������J�n������(�w���܂���)�̏ꍇ ==
        If p_sSyuryobibi > f_GetGakkiDay(C_KOUKIKAISI, p_iNendo) Then
            '== �ϐ��̐ݒ� ==
            w_sStartDay = f_GetGakkiDay(C_KOUKIKAISI, p_iNendo)
            w_sEndDay = p_sSyuryobibi

            '== ����̎��Ԋ�����擾 ==
            w_bRtn = f_GetJikanWari(C_GAKKI_KOUKI, p_sKCode, p_iNen, p_iClass, w_sStartDay, w_sEndDay, w_sTyokiWhere, C_SOU_NISSU, w_lJikan, p_iNendo)
            If w_bRtn <> True Then
                Exit Function
            End If

        End If

    '== �J�n��������J�n������(���)�̏ꍇ ==
    Else
        '== �ϐ��̐ݒ� ==
        w_sStartDay = f_GetGakkiDay(C_KOUKIKAISI, p_iNendo)
        w_sEndDay = p_sSyuryobibi

        '== ����̎��Ԋ�����擾 ==
        w_bRtn = f_GetJikanWari(C_GAKKI_KOUKI, p_sKCode, p_iNen, p_iClass, w_sStartDay, w_sEndDay, w_sTyokiWhere, C_SOU_NISSU, w_lJikan, p_iNendo)
        If w_bRtn <> True Then
            Exit Function
        End If

    End If

    '== ���Ԑ��̃Z�b�g ==
    p_lJikan = w_lJikan

    gf_SouJugyo = True

End Function

Private Function f_GetGakkiDay(p_iNo, p_iNendo)
'*******************************************************************************
' �@�@�@�\�F���t�Q�b�g
' �ԁ@�@�l�F�敪�ɓK��������t
' ���@�@���F�擾��̃R�[�h
' �@�\�ڍׁF�}�X�^����f�[�^���擾����@�w���敪
' ���@�@�l�F�Ȃ�
'*******************************************************************************
Dim w_sSql
Dim w_oRecord
Dim w_bRtn

	On Error Resume Next
	Err.Clear

    f_GetGakkiDay = ""

    '�d���`�F�b�N
    w_sSql = ""
    w_sSql = "Select M00_KANRI "
    w_sSql = w_sSql & "From "
    w_sSql = w_sSql & "M00_KANRI "
    w_sSql = w_sSql & "Where "
    w_sSql = w_sSql & "M00_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "M00_NO = " & p_iNo & " "

    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        '�擾�Ɏ��s
        Exit Function
    End If
    f_GetGakkiDay = cstr(w_oRecord("M00_KANRI"))

    gf_closeObject(w_oRecord)
    
    Exit Function

'f_GetGakkiDay_E:


End Function

Private Function f_GetTyokikyuka(p_oTyokiWhere,p_iNendo)
'*******************************************************************************
' �@�@�@�\�F�����x�ɂ̓��t�̎擾
' �ԁ@�@�l�Ftrue: �����@false: ���s
' ���@�@���F�����x�ɏ���
' �@�@�@�@�@p_iNendo - �N�x
' �@�\�ڍׁF�����x�ɂ̓��t�̎擾
' ���@�@�l�F�Ȃ�
'*******************************************************************************
Dim w_bRtn               '�߂�l
Dim w_sSql                'SQL
Dim w_oRecKyuka
Dim w_sWhereTyoki

	'On Error GoTo Err_Func
	On Error Resume Next
	Err.Clear

    '== ������ ==
    f_GetTyokikyuka = False

    '== SQL�̍쐬 ==
    w_sSql = ""
    w_sSql = w_sSql & "Select "
    w_sSql = w_sSql & "T31_KAISI_BI, "
    w_sSql = w_sSql & "T31_SYURYO_BI "
    w_sSql = w_sSql & "From "
    w_sSql = w_sSql & "T31_GYOJI_H "
    w_sSql = w_sSql & "Where "
    w_sSql = w_sSql & "T31_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "And "
    w_sSql = w_sSql & "T31_KYUKA_FLG = '" & C_FLG_TYOKIKYUKA & "' "
    w_sSql = w_sSql & "order by T31_KAISI_BI "

    '== �f�[�^�̎擾 ==
    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecKyuka, w_sSql)
    If w_bRtn <> 0 Then
        Exit Function
    End If

    '== �����x�ɂ̓��t���������߂�Where���쐬���� ==
    Do Until w_oRecKyuka.EOF

        w_sWhereTyoki = w_sWhereTyoki & "And Not ("
        w_sWhereTyoki = w_sWhereTyoki & "T32_HIDUKE "
        w_sWhereTyoki = w_sWhereTyoki & "Between "
        w_sWhereTyoki = w_sWhereTyoki & "'" & cstr(w_oRecKyuka("T31_KAISI_BI")) & "' "
        w_sWhereTyoki = w_sWhereTyoki & "And "
        w_sWhereTyoki = w_sWhereTyoki & "'" & cstr(w_oRecKyuka("T31_SYURYO_BI")) & "') "
        w_oRecKyuka.MoveNext

    Loop

    '== ���R�[�h�Z�b�g����� ==
    Call gf_closeObject(w_oRecKyuka)

    p_oTyokiWhere = w_sWhereTyoki

    '�f�[�^�����������Ƃ��Ă��G���[�ł͂Ȃ�
    f_GetTyokikyuka = True

    Exit Function

End Function

Private Function f_GetJikanWari(p_sGakki,p_sKCode,p_iNen, p_iClass, p_sStart, p_sEnd, p_sTyoki, p_bFlg, p_lJikan,p_iNendo)
'*******************************************************************************
' �@�@�@�\�F���Ԋ��f�[�^�̎擾
' �ԁ@�@�l�Ftrue: �����@false: ���s
' ���@�@���Fp_sGakki - �w���t���O
' �@�@�@�@�@p_sKCode - �ȖڃR�[�h
' �@�@�@�@�@p_iNen - �w�N
' �@�@�@�@�@p_iClass - �N���X
' �@�@�@�@�@p_sStart - �J�n��
' �@�@�@�@�@p_sEnd - �I����
' �@�@�@�@�@p_sTyoki - �����x�ɂ�Where
' �@�@�@�@�@p_bFlg - �����t���O�itrue�F�����Ɓ@false�F�����Ɓj
' �@�@�@�@�@p_lJikan - ���ʊi�[�ϐ�
' �@�\�ڍׁF���Ԋ��̎擾
' ���@�@�l�F�Ȃ�
'*******************************************************************************
    Dim w_sSql
    Dim w_bRtn
    Dim w_oRecord
    Dim w_lSojikan              '�����Ԑ�
    Dim w_lGyojiJikan           '�s�����Ԑ�
    
    'On Error GoTo f_GetJikanWari_Err
	On Error Resume Next
	Err.Clear

    f_GetJikanWari = False
    w_lSojikan = 0

    '�j���@�����J�E���g
    w_sSql = ""
    w_sSql = w_sSql & "SELECT DISTINCT "
    w_sSql = w_sSql & "T20_YOUBI_CD, "
    w_sSql = w_sSql & "T20_JIGEN "
    w_sSql = w_sSql & "FROM "
    w_sSql = w_sSql & "T20_JIKANWARI "
    w_sSql = w_sSql & "WHERE "
    w_sSql = w_sSql & "T20_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T20_GAKKI_KBN = " & p_sGakki & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T20_KAMOKU = '" & p_sKCode & "' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T20_GAKUNEN = " & p_iNen & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T20_CLASS = " & p_iClass & " "

    '== �f�[�^���擾���� ==
    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        Exit Function
    End If

    w_oRecord.MoveFirst

    '== �f�[�^�̊i�[ ==
    Do Until w_oRecord.EOF = True
        '== �����Ɠ����̎擾 ==
        w_bRtn = f_SouJugyoCnt(p_iNen, p_iClass, p_sStart, p_sEnd, w_oRecord("T20_YOUBI_CD"), w_oRecord("T20_JIGEN"), p_sTyoki, w_lSojikan, p_iNendo)
        If w_bRtn <> True Then
            '== ���� ==
            Call gf_closeObject(w_oRecord)

            Exit Function
        End If

        '== ���Ԃ̗݌v ==
        p_lJikan = p_lJikan + w_lSojikan

        '== �����Ǝ��Ԑ������߂�ꍇ ==
        If p_bFlg = C_JUGYO_NISSU Then
            '== �s�����Ԑ��̎擾 ==
            w_bRtn = f_GyojiJugyoCnt(p_iNen, p_iClass, p_sStart, p_sEnd, w_oRecord("T20_YOUBI_CD"), w_oRecord("T20_JIGEN"), p_sTyoki, w_lGyojiJikan, p_iNendo)
            If w_bRtn <> True Then
                Exit Function
            End If

            '== �����Ԑ�����s�����Ԑ������� ==
            p_lJikan = p_lJikan - w_lGyojiJikan * f_GetJigenTani(w_oRecord("T20_JIGEN"), p_iNendo)

        End If

        w_oRecord.MoveNext

    Loop

    '== ���� ==
    Call gf_closeObject(w_oRecord)

    f_GetJikanWari = True
    
    Exit Function

End Function

Public Function f_GetJigenTani(p_iJigen, p_iNendo)
'*******************************************************************************
' �@�@�@�\�F�����P�ʐ��̎擾
' �ԁ@�@�l�F�����P��
' ���@�@���F�����A�N�x
' �@�\�ڍׁF�����P�ʐ��̎擾
' ���@�@�l�F�Ȃ�
'*******************************************************************************
Dim w_sSql
Dim w_oRecord
Dim w_bRtn

	On Error Resume Next
	Err.Clear

    f_GetJigenTani = 1

    '�d���`�F�b�N
    w_sSql = ""
    w_sSql = w_sSql & "SELECT "
    w_sSql = w_sSql & "M07_TANISU "

    w_sSql = w_sSql & "FROM "
    w_sSql = w_sSql & "M07_JIGEN "

    w_sSql = w_sSql & "WHERE "
    w_sSql = w_sSql & "M07_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "M07_JIKAN = " & p_iJigen & " "

    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        '�擾�Ɏ��s
        Exit Function
    End If

    If w_oRecord.EOF = True Then: Exit Function

    f_GetJigenTani = w_oRecord("M07_TANISU")

    Call gf_closeObject(w_oRecord)

    Exit Function

End Function

Private Function f_SouJugyoCnt(p_iNen, p_iClass, p_sStart, _
                               p_sEnd, p_iYoubi, p_iJigen, _
                               p_sTyoki, p_lJikan, p_iNendo)
'*******************************************************************************
' �@�@�@�\�F�j�������ԃf�[�^�̎擾
' �ԁ@�@�l�Ftrue: �����@false: ���s
' ���@�@���F�w�N�A�N���X�A�J�n���A�I�����A�j���A�����A�����x�ɁA���ʊi�[�ϐ�
' �@�\�ڍׁF�j�������ԃf�[�^�̎擾
' ���@�@�l�F�Ȃ�
'*******************************************************************************
Dim w_sSql
Dim w_bRtn
Dim w_oRecord

	On Error Resume Next
	Err.Clear

    f_SouJugyoCnt = False
    p_lJikan = 0

    '�j���@�����J�E���g
    w_sSql = ""

    w_sSql = w_sSql & "SELECT DISTINCT "
    w_sSql = w_sSql & "T32_HIDUKE, "
    w_sSql = w_sSql & "T32_JIGEN "
    w_sSql = w_sSql & "FROM "
    w_sSql = w_sSql & "T32_GYOJI_M "

    w_sSql = w_sSql & "WHERE "
    w_sSql = w_sSql & "T32_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_KYUJITU_FLG = '0' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_GYOJI_CD = 0 "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_YOUBI_CD = " & p_iYoubi & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_JIGEN = " & p_iJigen & " "

    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_HIDUKE >= '" & p_sStart & "' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_HIDUKE <= '" & p_sEnd & "' "

    w_sSql = w_sSql & p_sTyoki

    '== �f�[�^���擾���� ==
    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        Exit Function
    End If

    '== �f�[�^�̊i�[ ==
    If w_oRecord.EOF = False Then

        w_oRecord.MoveLast

        '���̗j�����Ƃ̓������J�E���g����
		p_lJikan = gf_GetRsCount(w_oRecord)

    End If

    Call gf_closeObject(w_oRecord)

    f_SouJugyoCnt = True

    Exit Function

End Function

Private Function f_GyojiJugyoCnt( p_iNen,  p_iClass,  p_sStart, _
                                p_sEnd,  p_iYoubi,  p_iJigen, _
                                p_sTyoki, p_lJikan,  p_iNendo)
'*******************************************************************************
' �@�@�@�\�F�s�����ԃf�[�^�̎擾
' �ԁ@�@�l�Ftrue: �����@false: ���s
' ���@�@���F�w�N�A�N���X�A�J�n���A�I�����A�j���A�����A�����x�ɁA���ʊi�[�ϐ�
' �@�\�ڍׁF�s�����ԃf�[�^�̎擾
' ���@�@�l�F�Ȃ�
'*******************************************************************************
Dim w_sSql
Dim w_bRtn
Dim w_oRecord

	On Error Resume Next
	Err.Clear

    f_GyojiJugyoCnt = False

    '�j���@�����J�E���g
    w_sSql = ""

    w_sSql = w_sSql & "SELECT DISTINCT "
    w_sSql = w_sSql & "T32_HIDUKE, "
    w_sSql = w_sSql & "T32_JIGEN "
    
    w_sSql = w_sSql & "FROM "
    w_sSql = w_sSql & "T32_GYOJI_M "

    w_sSql = w_sSql & "WHERE "
    w_sSql = w_sSql & "T32_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_KYUJITU_FLG = '0' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_GYOJI_CD <> 0 "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_COUNT_KBN <> " & C_COUNT_KBN_JUGYO & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_YOUBI_CD = " & p_iYoubi & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_JIGEN = " & p_iJigen & " "

    If p_iNen <> C_ALLNEN Then

        w_sSql = w_sSql & "AND "
        w_sSql = w_sSql & "T32_GAKUNEN = " & p_iNen & " "

        If p_iClass <> C_ALLCLASS Then

            w_sSql = w_sSql & "AND "
            w_sSql = w_sSql & "T32_CLASS = " & C_ALLCLASS & " "

        End If

    End If

    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_HIDUKE >= '" & p_sStart & "' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_HIDUKE <= '" & p_sEnd & "' "

    w_sSql = w_sSql & p_sTyoki

    '== �f�[�^���擾���� ==
    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        Exit Function
    End If

    '== �f�[�^�̊i�[ ==
    If w_oRecord.EOF = False Then

        w_oRecord.MoveLast
        '���̗j�����Ƃ̓������J�E���g����
		p_lJikan = gf_GetRsCount(w_oRecord)

    End If

    Call gf_closeObject(w_oRecord)

    f_GyojiJugyoCnt = True

    Exit Function

End Function
%>
