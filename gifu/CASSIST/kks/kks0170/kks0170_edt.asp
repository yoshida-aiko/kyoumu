<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����o������
' ��۸���ID : kks/kks0170/kks0170_edt.asp
' �@      �\: ���y�[�W�����o�����͂̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: NENDO        '//�����N
'             KYOKAN_CD    '//����CD
'             GAKUNEN      '//�w�N
'             CLASSNO      '//�׽No
'             cboDate      '//���t
' ��      ��:
' ��      �n: cboDate      '//���t
' ��      ��:
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2001/07/24 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public m_iSyoriNen
    Public m_iKyokanCd
    Public m_iGakunen
    Public m_iClassNo
    Public m_sDate

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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�����o������"
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

        '//�ϐ�������
        Call s_ClearParam()

        '// Main���Ұ�SET
        Call s_SetParam()

'//�f�o�b�O
'Call s_DebugPrint()

        '// �����o���o�^
        w_iRet = f_AbsUpdate()
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
    m_iKyokanCd = ""
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_sDate     = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = trim(Request("NENDO"))
    m_iKyokanCd = trim(Request("KYOKAN_CD"))
    m_iGakunen  = trim(Request("GAKUNEN"))
    m_iClassNo  = trim(Request("CLASSNO"))
    m_sDate     = trim(Request("cboDate"))

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
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_sDate     = " & m_sDate     & "<br>"

End Sub

''********************************************************************************
''*  [�@�\]  ���O�C���ҏ����擾
''*  [����]  �Ȃ�
''*  [�ߒl]  0:���擾���� 99:���s
''*  [����]  
''********************************************************************************
'Function f_Get_UserInfo(p_UserName)
'Dim rs
'Dim w_sSQL
'
'    f_Get_UserInfo = 1
'    p_UserName = ""
'
'    Do
'        w_sSQL = ""
'        w_sSQL = w_sSQL & vbCrLf & " SELECT "
'        w_sSQL = w_sSQL & vbCrLf & "   M04_KYOKANMEI_SEI"
'        w_sSQL = w_sSQL & vbCrLf & " FROM M04_KYOKAN "
'        w_sSQL = w_sSQL & vbCrLf & " WHERE"
'        w_sSQL = w_sSQL & vbCrLf & "       M04_NENDO = " & cInt(m_iSyoriNen)
'        w_sSQL = w_sSQL & vbCrLf & "   AND M04_KYOKAN_CD = '" & m_iKyokanCd & "'"
'
'        iRet = gf_GetRecordset(rs, w_sSQL)
'        If iRet <> 0 Then
'            'ں��޾�Ă̎擾���s
'            msMsg = Err.description
'            f_Get_UserInfo = 99
'            Exit Do
'        End If
'
'        If rs.EOF = False Then
'            p_UserName = rs("M04_KYOKANMEI_SEI")
'        End If
'
'        f_Get_UserInfo = 0
'        Exit Do
'    Loop
'
'    Call gf_closeObject(rs)
'
'End Function

'********************************************************************************
'*  [�@�\]  ���ȕʏo���o�^
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_AbsUpdate()

    Dim w_sSQL
    Dim w_Rs
    Dim w_sUserId
    Dim w_iKekka

    On Error Resume Next
    Err.Clear
    
    f_AbsUpdate = 1

    Do 

		'//հ�ްID���擾
		w_sUserId = Session("LOGIN_ID")

        '//�w��No���擾
        m_sGakuseiNo = split(replace(Request("GAKUSEKI_NO")," ",""),",")
        m_iGakusekiCnt = UBound(m_sGakuseiNo)

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

        '//�N���X�̐l�������������s
        For i=0 To m_iGakusekiCnt

            '//�o��CD���擾
            w_iKekka = trim(Request("hidKBN" & m_sGakuseiNo(i)))

            If w_iKekka = "---" Then
                '//�o�����͕s�̂Ƃ��͍X�V���������Ȃ�
            Else

                w_sSQL = ""
                w_sSQL = w_sSQL & vbCrLf & " SELECT "
                w_sSQL = w_sSQL & vbCrLf & "  T30_NENDO, "
                w_sSQL = w_sSQL & vbCrLf & "  T30_HIDUKE, "
                w_sSQL = w_sSQL & vbCrLf & "  T30_GAKUNEN, "
                w_sSQL = w_sSQL & vbCrLf & "  T30_CLASS, "
                w_sSQL = w_sSQL & vbCrLf & "  T30_GAKUSEKI_NO"
                w_sSQL = w_sSQL & vbCrLf & " FROM T30_KESSEKI"
                w_sSQL = w_sSQL & vbCrLf & " WHERE "
                w_sSQL = w_sSQL & vbCrLf & "      T30_NENDO=" & m_iSyoriNen
                w_sSQL = w_sSQL & vbCrLf & "  AND T30_HIDUKE='" & m_sDate & "' "
                w_sSQL = w_sSQL & vbCrLf & "  AND T30_GAKUNEN= " & m_iGakunen
                w_sSQL = w_sSQL & vbCrLf & "  AND T30_CLASS= " & m_iClassNo
                w_sSQL = w_sSQL & vbCrLf & "  AND T30_GAKUSEKI_NO='" & m_sGakuseiNo(i) & "'"

                iRet = gf_GetRecordset(rs, w_sSQL)
                If iRet <> 0 Then
                    'ں��޾�Ă̎擾���s
                    msMsg = Err.description
                    f_AbsUpdate = 99
                    Exit Do
                End If

                If rs.EOF Then

                    If w_iKekka <> "" and cstr(w_iKekka)<>"0" Then

                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T30_KESSEKI"
                        w_sSQL = w_sSQL & vbCrLf & "  ("
                        w_sSQL = w_sSQL & vbCrLf & "  T30_NENDO, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_HIDUKE, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_YOUBI_CD, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_GAKUNEN, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_CLASS, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_GAKUSEKI_NO, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_SYUKKETU_KBN, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_INS_DATE, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_INS_USER"
                        w_sSQL = w_sSQL & vbCrLf & "  )VALUES("
                        w_sSQL = w_sSQL & vbCrLf & "   "  & cInt(m_iSyoriNen) & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & Trim(m_sDate)     & "',"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & Weekday(m_sDate)  & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   "  & cInt(m_iGakunen)  & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "   "  & cInt(m_iClassNo)  & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & m_sGakuseiNo(i)   & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   "  & Trim(w_iKekka)    & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & Date()            & "',"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & w_sUserId       & "'"
                        w_sSQL = w_sSQL & vbCrLf & "  )"

                        iRet = gf_ExecuteSQL(w_sSQL)
'response.write w_sSQL & "<br>"
'response.write "INSERT iRet = " & iRet & "<br>"
                        If iRet <> 0 Then
                            '//۰��ޯ�
                            Call gs_RollbackTrans()
                            msMsg = Err.description
                            f_AbsUpdate = 99
                            Exit Do
                        End If

                    End If

                Else

                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " UPDATE T30_KESSEKI SET "
                    w_sSQL = w_sSQL & vbCrLf & "  T30_SYUKKETU_KBN = "  & Trim(w_iKekka)    & " ," 
                    w_sSQL = w_sSQL & vbCrLf & "  T30_UPD_DATE =    '"  & Date()            & "',"
                    w_sSQL = w_sSQL & vbCrLf & "  T30_UPD_USER =    '"  & w_sUserId         & "'"
                    w_sSQL = w_sSQL & vbCrLf & " WHERE "
                    w_sSQL = w_sSQL & vbCrLf & "      T30_NENDO="    & m_iSyoriNen
                    w_sSQL = w_sSQL & vbCrLf & "  AND T30_HIDUKE='"  & m_sDate & "' "
                    w_sSQL = w_sSQL & vbCrLf & "  AND T30_GAKUNEN= " & m_iGakunen
                    w_sSQL = w_sSQL & vbCrLf & "  AND T30_CLASS= "   & m_iClassNo
                    w_sSQL = w_sSQL & vbCrLf & "  AND T30_GAKUSEKI_NO='" & m_sGakuseiNo(i) & "'"

                    iRet = gf_ExecuteSQL(w_sSQL)

'response.write w_sSQL & "<br>"
'response.write "UPDATE iRet = " & iRet & "<br>"

                    If iRet <> 0 Then
                        '//۰��ޯ�
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_AbsUpdate = 99
                        Exit Do
                    End If

                    '//ں��޾��CLOSE
                    Call gf_closeObject(rs)
                End If
            End If

        Next

        '//�Я�
        Call gs_CommitTrans()

        '//����I��
        f_AbsUpdate = 0
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
    <title>�����o������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
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
		alert("<%= C_TOUROKU_OK_MSG %>");

        //���X�g����submit
        document.frm.target = "main";
        document.frm.action = "./kks0170_bottom.asp"
        document.frm.submit();
        return;

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">

    <form name="frm" method="post">
    <input type="hidden" name="cboDate"   value="<%=Request("cboDate")%>">
    </form>

    </center>
    </body>
    </html>
<%
End Sub
%>