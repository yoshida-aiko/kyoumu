<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A���f����
' ��۸���ID : web/web0330/web0330_edt.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
'            ���[�h         ��      txtMode
'                                   �V�K = NEW
'                                   �X�V = UPDATE
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 2001/09/01 �ɓ����q �����ȊO�����p�ł���悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iMax           ':�ő�y�[�W
    Public m_iDsp           '// �ꗗ�\���s��
    Public m_sNendo         '�N�x
    Public m_sKyokanCd      '��������
    Public m_stxtMode       '���[�h
    Public m_sKenmei        '����
    Public m_sNaiyou        '���e
    Public m_sKaisibi       '�J�n��
    Public m_sSyuryoubi     '������
    Public m_sJoukin        '��΋敪
    Public m_sGakka         '�w�ȋ敪
    Public m_sKkanKBN       '�����敪
    Public m_sKkeiKBN       '���Ȍn��敪
    Public m_stxtNo         '�����ԍ�
    Public m_rs
    Public m_sListCd
    Dim    m_rCnt           '//���R�[�h����

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
    w_sMsgTitle="�A���f����"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_stxtMode = request("txtMode")

    m_sKenmei   = request("Kenmei")
    m_sNaiyou   = request("Naiyou")
    m_sKaisibi  = request("Kaisibi")
    m_sSyuryoubi= request("Syuryoubi")
    m_sNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_stxtNo    = request("txtNo")
If m_stxtMode = "UPD" Then
    m_sListCd   = request("KCD")
Else
    m_sListCd   = request("txtListCd")
End If
    m_iDsp = C_PAGE_LINE

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

        Select Case m_stxtMode
            Case "NEW2"
            '//�f�[�^�̎擾
            w_iRet = f_insertData()
            If w_iRet <> 0 Then
                '�ް��ް��Ƃ̐ڑ��Ɏ��s
                m_bErrFlg = True
                Exit Do
            End If
            Call showPage()
            Exit Do
            
            Case "UPD","UPD2"
            '//�f�[�^�̎擾�A�\��
            w_iRet = f_updateData()
            If w_iRet <> 0 Then
                '�ް��ް��Ƃ̐ڑ��Ɏ��s
                m_bErrFlg = True
                Exit Do
            End If
            Call showPage()
            Exit Do

        End Select
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
    Call gf_closeObject(m_Rs)
    '// �I������
    Call gs_CloseDatabase()
End Sub

Function f_insertData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_sSQL
Dim w_rs
Dim w_sKyokanList
Dim w_sListCd
Dim w_sKyokanCd
Dim w_iMaxNo
Dim i

    On Error Resume Next
    Err.Clear
    f_insertData = 1

    Do

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

        '//No�̍ő�l���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & "  MAX(T46_NO) AS MAXNO "
        w_sSQL = w_sSQL & "FROM "
        w_sSQL = w_sSQL & "  T46_RENRAK "

        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

        If IsNull(w_rs("MAXNO")) Then
            w_iMaxNo = 1
        Else
            w_iMaxNo = cInt(w_rs("MAXNO")) + 1
        End If

        '//���M��I����ʂŃ`�F�b�N���ꂽ�f�[�^��z��Ŏ擾
        w_sKyokanList = split(replace(m_sListCd," ",""),",")

        iMax = UBound(w_sKyokanList)

'---------20010901 ito
'        m_sSQL = ""
'        m_sSQL = m_sSQL & "SELECT "
'        m_sSQL = m_sSQL & "  M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
'        m_sSQL = m_sSQL & "FROM "
'        m_sSQL = m_sSQL & "  M04_KYOKAN "
'        m_sSQL = m_sSQL & "WHERE "
'        m_sSQL = m_sSQL & "  M04_KYOKAN_CD IN (" & Trim(m_sListCd) & ") "
'
'        Set m_rs = Server.CreateObject("ADODB.Recordset")
'        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
'        If w_iRet <> 0 Then
'            'ں��޾�Ă̎擾���s
'            m_bErrFlg = True
'            Exit Do 
'        End If

    For i=0 to iMax
        w_sKyokanCd = w_sKyokanList(i)

        '//�w�N��N���X�̃f�[�^
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & "INSERT INTO T46_RENRAK " 
        m_sSQL = m_sSQL & vbCrLf & " ( " 
        m_sSQL = m_sSQL & vbCrLf & "  T46_NO,T46_KYOKAN_CD,T46_KENMEI,T46_NAIYO,T46_KAISI,T46_SYURYO,T46_KAKNIN, " 
        m_sSQL = m_sSQL & vbCrLf & "  T46_INS_DATE,T46_INS_USER " 
        m_sSQL = m_sSQL & vbCrLf & ") " 
        m_sSQL = m_sSQL & vbCrLf & " VALUES " 
        m_sSQL = m_sSQL & vbCrLf & "( " 
        m_sSQL = m_sSQL & vbCrLf & " '" & cInt(w_iMaxNo) & "', " 
        m_sSQL = m_sSQL & vbCrLf & "'" & Trim(w_sKyokanCd) & "', " 
        m_sSQL = m_sSQL & vbCrLf & "'" & Trim(m_sKenmei) & "', " 
        m_sSQL = m_sSQL & vbCrLf & "'" & Trim(m_sNaiyou) & "', " 
        m_sSQL = m_sSQL & vbCrLf & "'" & gf_YYYY_MM_DD(Trim(m_sKaisibi),"/") & "', " 
        m_sSQL = m_sSQL & vbCrLf & "'" & gf_YYYY_MM_DD(Trim(m_sSyuryoubi),"/") & "', " 
        m_sSQL = m_sSQL & vbCrLf & " 0 , " 
        m_sSQL = m_sSQL & vbCrLf & "'" & gf_YYYY_MM_DD(date(),"/") & "', " 
        m_sSQL = m_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "' " 
        m_sSQL = m_sSQL & vbCrLf & "   )"

        iRet = gf_ExecuteSQL(m_sSQL)
        If iRet <> 0 Then
            '//۰��ޯ�
            Call gs_RollbackTrans()
            msMsg = Err.description
            f_insertData = 99
            Exit Do
        End If
    Next

    '//�Я�
    Call gs_CommitTrans()

    f_insertData = 0

    Exit Do

    Loop

End Function

Function f_updateData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_sSQL
Dim w_Srs           '�폜�p�̃��R�[�h�Z�b�g
Dim w_Brs           '�ȑO�̃��R�[�h�Z�b�g
Dim w_Nrs           '���݂̃��R�[�h�Z�b�g
Dim w_sKyokanList
Dim w_sKyokanCd
Dim w_sUpdFlg
Dim i

    On Error Resume Next
    Err.Clear
    f_updateData = 1

    Do

        Call gs_BeginTrans()

        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & "  T46_NO,T46_KYOKAN_CD "
        w_sSQL = w_sSQL & "FROM "
        w_sSQL = w_sSQL & "  T46_RENRAK "
        w_sSQL = w_sSQL & "WHERE "
        w_sSQL = w_sSQL & "  T46_NO = '" & cInt(m_stxtNo) & "' "

        Set w_Brs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_Brs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

        '//���M��I����ʂŃ`�F�b�N���ꂽ�f�[�^��z��Ŏ擾
        w_sKyokanList = split(replace(m_sListCd," ",""),",")

        iMax = UBound(w_sKyokanList)

        '//�e�[�u���ɏ�������
        For i=0 to iMax
            w_sKyokanCd = w_sKyokanList(i)

            w_Brs.MoveFirst
            Do Until w_Brs.EOF

                UpdFlg = False
                If w_Brs("T46_KYOKAN_CD") = Trim(w_sKyokanCd) Then

                    '//T46_RENRAK�ɂ��łɐ��k��񂪂���ꍇ��UPDATE
                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " UPDATE T46_RENRAK SET "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_KENMEI = '"  & Trim(m_sKenmei) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "   T46_NAIYO = '"  & Trim(m_sNaiyou) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "   T46_KAISI = '"  & gf_YYYY_MM_DD(Trim(m_sKaisibi),"/") & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "   T46_SYURYO = '"  & gf_YYYY_MM_DD(Trim(m_sSyuryoubi),"/") & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "   T46_KAKNIN = '0' ,"
                    w_sSQL = w_sSQL & vbCrLf & "   T46_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")            & "',"
                    w_sSQL = w_sSQL & vbCrLf & "   T46_UPD_USER = '"    & Session("LOGIN_ID") & "'"
                    w_sSQL = w_sSQL & vbCrLf & " WHERE "
                    w_sSQL = w_sSQL & vbCrLf & "        T46_NO = " & Cint(m_stxtNo) & "  "
                    w_sSQL = w_sSQL & vbCrLf & "    AND T46_KYOKAN_CD = '" & Trim(w_sKyokanList(i)) & "' "

                    iRet = gf_ExecuteSQL(w_sSQL)
                    If iRet <> 0 Then
                        '//۰��ޯ�
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_updateData = 99
                        Exit Do
                    End If
                UpdFlg = True
                Exit Do
                End If 
                w_Brs.MoveNext
            Loop

                If UpdFlg = False Then

                    '//T06_GAKU_IIN�ɐ��k��񂪂Ȃ��ꍇINSERT
                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T46_RENRAK  "
                    w_sSQL = w_sSQL & vbCrLf & "   ("
                    w_sSQL = w_sSQL & vbCrLf & "   T46_NO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_KYOKAN_CD, "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_KENMEI, "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_NAIYO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_KAISI, "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_SYURYO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_KAKNIN, "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_INS_DATE, "
                    w_sSQL = w_sSQL & vbCrLf & "   T46_INS_USER "
                    w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
                    w_sSQL = w_sSQL & vbCrLf & "    '" & cInt(m_stxtNo) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & Trim(w_sKyokanList(i)) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & Trim(m_sKenmei) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & Trim(m_sNaiyou) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & gf_YYYY_MM_DD(Trim(m_sKaisibi),"/") & "',"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & gf_YYYY_MM_DD(Trim(m_sSyuryoubi),"/") & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '0' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & gf_YYYY_MM_DD(date(),"/") & "',"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & Session("LOGIN_ID") & "' "
                    w_sSQL = w_sSQL & vbCrLf & "   )"

                    iRet = gf_ExecuteSQL(w_sSQL)
                    If iRet <> 0 Then
                        '//۰��ޯ�
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_updateData = 99
                        Exit For
                    End If
                End If
        Next

    '//�Я�
    Call gs_CommitTrans()

    '//�폜����
    Call gs_BeginTrans()

            w_sSQL = ""
            w_sSQL = w_sSQL & "SELECT "
            w_sSQL = w_sSQL & "  T46_NO,T46_KYOKAN_CD "
            w_sSQL = w_sSQL & "FROM "
            w_sSQL = w_sSQL & "  T46_RENRAK "
            w_sSQL = w_sSQL & "WHERE "
            w_sSQL = w_sSQL & "  T46_NO = '" & cInt(m_stxtNo) & "' "

            Set w_Srs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(w_Srs, w_sSQL,m_iDsp)
            If w_iRet <> 0 Then
                'ں��޾�Ă̎擾���s
                m_bErrFlg = True
                Exit Do 
            End If
    
        w_Srs.MoveFirst
        Do Until w_Srs.EOF
    
            For i=0 to iMax
                UpdFlg = False
                w_sKyokanCd = w_sKyokanList(i)
    
                If w_Srs("T46_KYOKAN_CD") = w_sKyokanList(i) Then
                    UpdFlg = True
                    Exit For
                End If
            Next
            If UpdFlg = False Then
    
                w_sSQL = ""
                w_sSQL = w_sSQL & vbCrLf & " DELETE FROM T46_RENRAK  "
                w_sSQL = w_sSQL & vbCrLf & " WHERE "
                w_sSQL = w_sSQL & vbCrLf & "     T46_NO = '" & cInt(m_stxtNo) & "' "
                w_sSQL = w_sSQL & vbCrLf & " AND T46_KYOKAN_CD = '" & w_Srs("T46_KYOKAN_CD") & "' "

                iRet = gf_ExecuteSQL(w_sSQL)
                If iRet <> 0 Then
                    '//۰��ޯ�
                    Call gs_RollbackTrans()
                    msMsg = Err.description
                    f_updateData = 99
                    Exit Do
                End If
            End If
            w_Srs.MoveNext
        Loop

    '//�Я�
    Call gs_CommitTrans()

    f_updateData = 0

    Exit Do

    Loop

End Function

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
    <title>�s���p�o������</title>
    <link rel=stylesheet href=../../font.css type=text/css>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

        location.href = "./default.asp"
        return;
    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>