<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �e��ψ��o�^
' ��۸���ID : gah/gak0470/gaku0470_edt.asp
' �@      �\: ���y�[�W �e��ψ��o�^�̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
' ��      ��:
' ��      �n: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
' ��      ��:
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 �O�c
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
    Public  m_sGakunen
    Public  m_sClassNo
    Dim     m_iNendo
    Dim     m_sKyokanCd

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
    w_sMsgTitle="�e��ψ��o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp = C_PAGE_LINE

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

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()

End Sub

Function f_AbsUpdate()
'********************************************************************************
'*  [�@�\]  �w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************

    Dim w_sUserName
    Dim w_sSQL
    Dim w_Rs
    Dim i
    i = 1

    On Error Resume Next
    Err.Clear
    
    f_AbsUpdate = 1

    w_iMax = request("HIDMAX")
'	w_iGAKKI = 

    Do 

'        '// ���O�C���Җ��̂��擾
'        w_iRet = f_Get_UserInfo(w_sUserName)
'        If w_iRet <> 0 Then
'            Exit Do
'        End If

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

            '//�s���������������s
            For i=1 to w_iMax

                If request("gakuNo" & i ) <> "" Then

                    If request("Before" & i ) = "" Then

                        '//T06_GAKU_IIN�ɐ��k��񂪂Ȃ��ꍇINSERT
                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T06_GAKU_IIN  "
                        w_sSQL = w_sSQL & vbCrLf & "   ("
                        w_sSQL = w_sSQL & vbCrLf & "   T06_NENDO, "
						w_sSQL = w_sSQL & vbCrLf & "   T06_GAKKI_KBN, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_GAKUSEI_NO, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_DAIBUN_CD, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_SYOBUN_CD, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_INS_DATE, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_INS_USER "
                        w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
                        w_sSQL = w_sSQL & vbCrLf & "    '"  & cInt(m_iNendo) & "' ,"
						w_sSQL = w_sSQL & vbCrLf & "   " & request("GAKKI") & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(request("gakuNo" & i )) & "',"
                        w_sSQL = w_sSQL & vbCrLf & "    '"  & cInt(request("iinDai" & i )) & "' ,"
                        w_sSQL = w_sSQL & vbCrLf & "    '"  & cInt(request("iinSyo" & i )) & "' ,"
                        w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_YYYY_MM_DD(date(),"/") & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   '"  & Session("LOGIN_ID") & "' "
                        w_sSQL = w_sSQL & vbCrLf & "   )"

                        iRet = gf_ExecuteSQL(w_sSQL)
                        If iRet <> 0 Then
                            '//۰��ޯ�
                            Call gs_RollbackTrans()
                            msMsg = Err.description
                            f_AbsUpdate = 99
                            Exit Do
                        End If

                    ElseIf request("gakuNo" & i ) <> request("Before" & i ) Then

                        '//T06_GAKU_IIN�ɂ��łɐ��k��񂪂���ꍇ��UPDATE
                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " UPDATE T06_GAKU_IIN SET "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_GAKUSEI_NO = '"  & Trim(request("gakuNo" & i))    & "' ,"
                        w_sSQL = w_sSQL & vbCrLf & "   T06_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")            & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   T06_UPD_USER = '"    & Session("LOGIN_ID")       & "'"
                        w_sSQL = w_sSQL & vbCrLf & " WHERE "
                        w_sSQL = w_sSQL & vbCrLf & "        T06_NENDO = '" & m_iNendo & "'  "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_GAKUSEI_NO = '" & Trim(request("Before" & i )) & "' "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_DAIBUN_CD = '" & cInt(request("iinDai" & i )) & "' "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_SYOBUN_CD = '" & cInt(request("iinSyo" & i )) & "' "
						w_sSQL = w_sSQL & vbCrLf & "	AND T06_GAKKI_KBN = " & request("GAKKI") & " "

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
                    If request("Before" & i ) <> "" Then

                        '//T06_GAKU_IIN�ɂ��łɐ��k��񂪂���A���ݕ\����ŋ󔒂̏ꍇDELETE
                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " DELETE FROM T06_GAKU_IIN "
                        w_sSQL = w_sSQL & vbCrLf & " WHERE "
                        w_sSQL = w_sSQL & vbCrLf & "        T06_NENDO = '" & m_iNendo & "'  "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_GAKUSEI_NO = '" & Trim(request("Before" & i )) & "' "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_DAIBUN_CD = '" & cInt(request("iinDai" & i )) & "' "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_SYOBUN_CD = '" & cInt(request("iinSyo" & i )) & "' "
						w_sSQL = w_sSQL & vbCrLf & "	AND T06_GAKKI_KBN = " & request("GAKKI") & " "

                        iRet = gf_ExecuteSQL(w_sSQL)
                        If iRet <> 0 Then
                            '//۰��ޯ�
                            Call gs_RollbackTrans()
                            msMsg = Err.description
                            f_AbsUpdate = 99
                            Exit Do
                        End If
                    End If
                End If

                    '//ں��޾��CLOSE
                    Call gf_closeObject(rs)
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
    <title>�e��ψ��o�^</title>
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

		alert("<%= C_TOUROKU_OK_MSG %>");

        location.href = "./default.asp"
        return;
    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    <INPUT TYPE=HIDDEN NAME=CLASS   VALUE="<%=Request("CLASS")%>">
    <INPUT TYPE=HIDDEN NAME=GAKUNEN VALUE="<%=Request("GAKUNEN")%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>