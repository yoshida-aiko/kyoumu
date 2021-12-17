<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �s���o������
' ��۸���ID : kks/kks0140/kks0140_edt.asp
' �@      �\: ���y�[�W �s���o�����͂̓o�^�A�X�V
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
    Const DebugPrint = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public m_iSyoriNen
    Public m_iKyokanCd
    Public m_sGakunen
    Public m_sClassNo
    Public m_sTuki

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
    w_sMsgTitle="�s���o������"
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

        '// �s���o���o�^
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
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sTuki     = ""

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

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_AbsUpdate()

    Dim w_sSQL
    Dim w_Rs
    Dim w_sUserName
    Dim w_iKekka

    On Error Resume Next
    Err.Clear
    
    f_AbsUpdate = 1

    Do 

		'//հ�ްID���擾
		w_sUserId = Session("LOGIN_ID")

        '//�s��CD���擾
        m_sGyojiCD = Request("GYOJI_CD")

        '//�w��No���擾
        m_sGakusekiNo = split(replace(Request("GAKUSEKI_NO")," ",""),",")
        m_iGakusekiCnt = UBound(m_sGakusekiNo)

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

        '//�N���X�̐l�������������s
        For i=0 To m_iGakusekiCnt

            '//���Ȑ����擾
            w_iKekka = replace(trim(Request("SU_" & m_sGakusekiNo(i))),"+","")

            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " SELECT "
            w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_KEKKA"
            w_sSQL = w_sSQL & vbCrLf & " FROM T22_GYOJI_SYUKKETU"
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "   T22_NENDO=" & cInt(m_iSyoriNen) & " AND "
            w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN=" & cInt(m_sGakunen) & " AND "
            w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS=" & cInt(m_sClassNo) & " AND "
            w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUSEKI_NO='" & Trim(m_sGakusekiNo(i)) & "' AND "
            w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD='" & Trim(m_sGyojiCD) & "'"

            iRet = gf_GetRecordset(rs, w_sSQL)
            If iRet <> 0 Then
                'ں��޾�Ă̎擾���s
                msMsg = Err.description
                f_AbsUpdate = 99
                Exit Do
            End If

            If rs.EOF Then

                If w_iKekka <> "" Then

                    '//T22_GYOJI_SYUKKETU�ɐ��k��񂪂Ȃ��ꍇ�ŁA���Ȑ������͂���Ă���ꍇ��INSERT
                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T22_GYOJI_SYUKKETU  "
                    w_sSQL = w_sSQL & vbCrLf & "   ("
                    w_sSQL = w_sSQL & vbCrLf & "   T22_NENDO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUSEKI_NO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_KEKKA, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_INS_DATE,"
                    w_sSQL = w_sSQL & vbCrLf & "   T22_INS_USER"
                    w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
                    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_iSyoriNen)      & " ,"
                    w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(m_sGyojiCD)       & "',"
                    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_sGakunen)       & " ,"
                    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_sClassNo)       & " ,"
                    w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_SetNull2Zero(Trim(m_sGakusekiNo(i))) & "',"
                    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(gf_SetNull2Zero(trim(w_iKekka))) & " ,"
                    w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_YYYY_MM_DD(date(),"/")                 & "',"
                    w_sSQL = w_sSQL & vbCrLf & "   '"  & w_sUserId              & "' "
                    w_sSQL = w_sSQL & vbCrLf & "   )"

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

                '//T22_GYOJI_SYUKKETU�ɂ��łɐ��k��񂪂���ꍇ��UPDATE
                w_sSQL = ""
                w_sSQL = w_sSQL & vbCrLf & " UPDATE T22_GYOJI_SYUKKETU SET "
                w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_KEKKA =" & cInt(gf_SetNull2Zero(trim(w_iKekka))) & " ,"
                w_sSQL = w_sSQL & vbCrLf & "   T_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")              & "',"
                w_sSQL = w_sSQL & vbCrLf & "   T_UPD_USER = '"    & w_sUserId           & "'"
                w_sSQL = w_sSQL & vbCrLf & " WHERE "
                w_sSQL = w_sSQL & vbCrLf & "   T22_NENDO="        & cInt(m_iSyoriNen) & " AND "
                w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN="      & cInt(m_sGakunen)  & " AND "
                w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS="        & cInt(m_sClassNo)  & " AND "
                w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUSEKI_NO='" & Trim(m_sGakusekiNo(i)) & "' AND "
                w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD='"    & Trim(m_sGyojiCD)  & "'"

                iRet = gf_ExecuteSQL(w_sSQL)
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
    <title>�s���o������</title>
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

        //���X�g����submit
        document.frm.target = "main";
        document.frm.action = "./kks0140_bottom.asp"
        document.frm.submit();
        return;


    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    <input type="hidden" name="NENDO"     value="<%=Request("NENDO")%>">
    <input type="hidden" name="KYOKAN_CD" value="<%=Request("KYOKAN_CD")%>">
    <input type="hidden" name="GAKUNEN"   value="<%=Request("GAKUNEN")%>">
    <input type="hidden" name="CLASSNO"   value="<%=Request("CLASSNO")%>">
    <input type="hidden" name="TUKI"      value="<%=Request("TUKI")%>">
    <INPUT TYPE=HIDDEN NAME="GYOJI_CD"  value = "<%=Request("GYOJI_CD")%>">
    <INPUT TYPE=HIDDEN NAME="GYOJI_MEI" value = "<%=Request("GYOJI_MEI")%>">
    <INPUT TYPE=HIDDEN NAME="KAISI_BI"  value = "<%=Request("KAISI_BI")%>">
    <INPUT TYPE=HIDDEN NAME="SYURYO_BI" value = "<%=Request("SYURYO_BI")%>">
    <INPUT TYPE=HIDDEN NAME="SOJIKANSU" value = "<%=Request("SOJIKANSU")%>">

	</form>
    </body>
    </html>
<%
End Sub
%>