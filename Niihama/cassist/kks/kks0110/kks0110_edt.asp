<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0110/kks0110_edt.asp
' �@      �\: ���y�[�W���Əo�����͂̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
'             TUKI           '//��
' ��      ��:
' ��      �n: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
'             TUKI           '//��
' ��      ��:
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 �ɓ����q
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
    Public m_sGakunen
    Public m_sClassNo
    Public m_sTuki
    Public m_sKamokuCd

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
    w_sMsgTitle="���Əo������"
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

        '// ���ȕʏo���o�^
        w_iRet = f_AbsUpdate()
        If w_iRet <> 0 Then
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
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sTuki     = ""
    m_sKamokuCd = ""

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
    m_sGakunen  = trim(Request("GAKUNEN"))
    m_sClassNo  = trim(Request("CLASSNO"))
    m_sTuki     = trim(Request("TUKI"))
    m_sKamokuCd = trim(Request("KAMOKU_CD"))

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
Exit Sub
    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  ���ȕʏo���o�^
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_AbsUpdate()

    Dim w_sSQL
    Dim w_Rs
    Dim w_iKekka

    On Error Resume Next
    Err.Clear
    
    f_AbsUpdate = 1

    Do 

		'//հ�ްID���擾
		w_sUserId = Session("LOGIN_ID")

        '//���Ԋ������擾
        m_Date_Jigen = split(replace(Request("JIKANWARI")," ",""),",")
        m_iJikanCnt = UBound(m_Date_Jigen)

        '//�w��No���擾
        m_sGakuseiNo = split(replace(Request("GAKUSEI")," ",""),",")
        m_iGakusekiCnt = UBound(m_sGakuseiNo)

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

        '//�N���X�̐l�������������s
        For i=0 To m_iGakusekiCnt

            '//���Ԑ������������s
            For j=0 to m_iJikanCnt

                '//�o��CD���擾
                w_iKekka = trim(Request("hidKBN" & m_sGakuseiNo(i) & "_" & replace(m_Date_Jigen(j),"/","")))

                If w_iKekka = "---" Or Trim(m_sGakuseiNo(i)) = "" Then
                    '//�o�����͕s�̂Ƃ��͍X�V���������Ȃ�
                Else

                    w_DJ = split(m_Date_Jigen(j),"_")
                    w_Date  = w_DJ(0)
                    w_Jigen = replace(w_DJ(1),"$",".")

                    '//�w�N�A�N���XNO���擾
                    iRet = f_Get_NenClass(m_sGakuseiNo(i),w_Gakunen,w_Class,w_Gakuseki)
                    If iRet <> 0 Then
                        Exit Do
                    End If

                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " SELECT "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_NENDO, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_HIDUKE, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_GAKUSEKI_NO, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIGEN, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_KAMOKU, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_KYOKAN, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_SYUKKETU_KBN, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIMU_FLG"
                    w_sSQL = w_sSQL & vbCrLf & " FROM T21_SYUKKETU"
                    w_sSQL = w_sSQL & vbCrLf & " WHERE "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_NENDO=" & cInt(m_iSyoriNen) & " AND "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_HIDUKE='" & w_Date & "' AND "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_GAKUSEKI_NO='" & w_Gakuseki & "' AND "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIGEN=" & w_Jigen

                    iRet = gf_GetRecordset(rs, w_sSQL)
                    If iRet <> 0 Then
                        'ں��޾�Ă̎擾���s
                        msMsg = Err.description
                        f_AbsUpdate = 99
                        Exit Do
                    End If

                    If rs.EOF Then

                        If w_iKekka <> "" and cstr(w_iKekka)<>"0" Then

                            '//T22_GYOJI_SYUKKETU�ɐ��k��񂪂Ȃ��ꍇ�ŁA���Ȑ������͂���Ă���ꍇ��INSERT
                            w_sSQL = ""
                            w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T21_SYUKKETU  "
                            w_sSQL = w_sSQL & vbCrLf & "   ("
                            w_sSQL = w_sSQL & vbCrLf & "  T21_NENDO, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_HIDUKE, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_YOUBI_CD, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_GAKUNEN, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_CLASS, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_GAKUSEKI_NO, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_JIGEN, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_KAMOKU, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_KYOKAN, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU_KBN, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_JIMU_FLG, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_INS_DATE, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_INS_USER"
                            w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
                            w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_iSyoriNen) & " ,"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(w_Date)      & "',"
                            w_sSQL = w_sSQL & vbCrLf & "    "  & cint(Weekday(w_Date))   & ","
                            w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(w_Gakunen)   & " ,"
                            w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(w_Class)     & " ,"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(w_Gakuseki)  & "',"
                            w_sSQL = w_sSQL & vbCrLf & "    "  & w_Jigen     & " ,"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(m_sKamokuCd) & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(m_iKyokanCd) & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(w_iKekka)    & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & cstr(C_JIMU_FLG_NOTJIMU) & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_YYYY_MM_DD(date(),"/")            & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & w_sUserId         & "' "
                            w_sSQL = w_sSQL & vbCrLf & "   )"

'response.write w_sSQL & "<br>"
'response.write "INSERT iRet = " & iRet & "<br>"

                            iRet = gf_ExecuteSQL(w_sSQL)
                            If iRet <> 0 Then
                                '//۰��ޯ�
                                Call gs_RollbackTrans()
                                msMsg = Err.description
                                f_AbsUpdate = 99
                                Exit Do
                            End If

                        End If

                    Else

                        '//T21_SYUKKETU�ɂ��łɐ��k��񂪂���ꍇ��UPDATE
                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " UPDATE T21_SYUKKETU SET "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_SYUKKETU_KBN ='" & Trim(w_iKekka)    & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   T21_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")            & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   T21_UPD_USER = '"    & w_sUserId         & "' "
                        w_sSQL = w_sSQL & vbCrLf & " WHERE "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_NENDO="          & cInt(m_iSyoriNen) & "  AND "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_HIDUKE='"        & Trim(w_Date)      & "' AND "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_GAKUSEKI_NO='"   & Trim(w_Gakuseki)  & "' AND "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_JIGEN="          & w_Jigen

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

                    End If

                    '//ں��޾��CLOSE
                    Call gf_closeObject(rs)

                End If

            Next
        Next

        '//�Я�
        Call gs_CommitTrans()

        '//����I��
        f_AbsUpdate = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �w�N�A�w��NO���擾
'*  [����]  p_Gakuseki�F�w��NO
'*  [�ߒl]  p_Gakunen�F�w�N
'*          p_Class�F�N���X
'*          p_Gakuseki:�w��NO
'*  [����]  
'********************************************************************************
Function f_Get_NenClass(p_Gakusei,p_Gakunen,p_Class,p_Gakuseki)

    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear
    
    f_Get_NenClass = 1

    p_Gakunen = ""
    p_Class = ""
    p_Gakuseki = ""

    Do 

        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLASS,"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO"
        w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & " T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO='" & trim(p_Gakusei) & "'"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_NenClass = 99
            Exit Do
        End If

        If rs.EOF = false Then
            p_Gakunen  = rs("T13_GAKUNEN")
            p_Class    = rs("T13_CLASS")
            p_Gakuseki = rs("T13_GAKUSEKI_NO")
        End If

        f_Get_NenClass = 0
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
    <title>���Əo������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>

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

		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("�ĕ\�����Ă��܂��@���΂炭���҂���������")%>"

	    parent.main.document.frm.target = "main";
        //parent.main.document.frm.action="./WaitAction.asp";
	    parent.main.document.frm.action = "./kks0110_bottom.asp"
	    parent.main.document.frm.submit();
	    return;


    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    <input type="hidden" name="Tuki_Zenki_Start" value="<%=Request("Tuki_Zenki_Start")%>">
    <input type="hidden" name="Tuki_Kouki_Start" value="<%=Request("Tuki_Kouki_Start")%>">
    <input type="hidden" name="Tuki_Kouki_End"   value="<%=Request("Tuki_Kouki_End")%>">
    <INPUT TYPE=HIDDEN NAME="NENDO"     value="<%=Request("NENDO")%>">
    <INPUT TYPE=HIDDEN NAME="KYOKAN_CD" value="<%=Request("KYOKAN_CD")%>">
    <INPUT TYPE=HIDDEN NAME="TUKI"      value="<%=Request("TUKI")%>">
    <INPUT TYPE=HIDDEN NAME="GAKKI"     value="<%=Request("GAKKI")%>">
    <INPUT TYPE=HIDDEN NAME="GAKUNEN"   value="<%=Request("GAKUNEN")%>">
    <INPUT TYPE=HIDDEN NAME="CLASSNO"   value="<%=Request("CLASSNO")%>">
    <INPUT TYPE=HIDDEN NAME="KAMOKU_CD" value="<%=Request("KAMOKU_CD")%>">
    <INPUT TYPE=HIDDEN NAME="SYUBETU"   value="<%=Request("SYUBETU")%>">

    <INPUT TYPE=HIDDEN NAME="KAMOKU_NAME" value="<%=Request("KAMOKU_NAME")%>">
    <INPUT TYPE=HIDDEN NAME="CLASS_NAME"  value="<%=Request("CLASS_NAME")%>">

    <input TYPE="HIDDEN" NAME="txtURL" VALUE="kks0110_bottom.asp">
    <input TYPE="HIDDEN" NAME="txtMsg" VALUE="<%=Server.HTMLEncode("�ĕ\�����Ă��܂��@���΂炭���҂���������")%>">

    </form>
    </body>
    </html>
<%
End Sub
%>