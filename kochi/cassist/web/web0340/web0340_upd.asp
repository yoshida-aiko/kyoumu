<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �l���C�I���Ȗڌ���
' ��۸���ID : web/web0340/web0340_edt.asp
' �@      �\: ���y�[�W �l���C�I���Ȗڌ���̍X�V
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
' ��      ��: 2001/07/24 �O�c �q�j
' ��      �X: 2001/08/28 �ɓ����q �w�b�_���؂藣���Ή�
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
    Dim     m_iNendo        '//�N�x
    Dim     m_sKyokanCd     '//����CD
    Dim     n_Max           '//�ő吔
    Dim     k_Max           '//�ő吔

    Public  m_iMax          '�ő�y�[�W
    Public  m_iDsp          '�ꗗ�\���s��

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

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
'--------2001/08/28 ito 
'    m_sgakuNo   = Trim(request("gakuNo"))
'    m_skamokuCd = Trim(request("kamokuCd"))
    n_Max       = request("n_Max")
    k_Max       = request("k_Max")
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

        '// �l���C�I���Ȗڌ���
        w_iRet = f_Update()
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

Function f_Update()
'********************************************************************************
'*  [�@�\]  �w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Dim w_sGakuNo
Dim w_sKamokuCd
Dim n
Dim k



    On Error Resume Next
    Err.Clear
    
    f_Update = 1

    Do 

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

        For k=1 To k_Max                '//�w��No����
            For n=1 To n_Max            '//�Ȗ�CD����

                w_sGakuNo = request("GakuNo"& k)
                w_sKamokuCd = request("KamokuCd"& n)

                If request("MAE"& k &"_"& n) <> "��" AND request("ATO"& k &"_"& n) = "��" Then

                    w_sSQL = ""
                    w_sSQL = w_sSQL & " UPDATE T16_RISYU_KOJIN SET "
                    w_sSQL = w_sSQL & "   T16_SELECT_FLG = " & C_SENTAKU_YES & ", "
                    w_sSQL = w_sSQL & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "',"
                    w_sSQL = w_sSQL & "   T16_UPD_USER = '" & Session("LOGIN_ID") & "' "
                    w_sSQL = w_sSQL & " WHERE "
                    w_sSQL = w_sSQL & "   T16_NENDO = " & m_iNendo & " "
                    w_sSQL = w_sSQL & " AND T16_GAKUSEI_NO = '" & w_sGakuNo & "' "
                    w_sSQL = w_sSQL & " AND T16_KAMOKU_CD = '" & w_sKamokuCd & "' "

                    iRet = gf_ExecuteSQL(w_sSQL)
                    If iRet <> 0 Then
                        '//۰��ޯ�
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_Update = 99
                        Exit For
                    End If

                ElseIf request("MAE"& k &"_"& n) = "��" AND request("ATO"& k &"_"& n) <> "��" Then

                    w_sSQL = ""
                    w_sSQL = w_sSQL & " UPDATE T16_RISYU_KOJIN SET "
                    w_sSQL = w_sSQL & "   T16_SELECT_FLG = " & C_SENTAKU_NO & ", "
                    w_sSQL = w_sSQL & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "',"
                    w_sSQL = w_sSQL & "   T16_UPD_USER = '" & Session("LOGIN_ID") & "' "
                    w_sSQL = w_sSQL & " WHERE "
                    w_sSQL = w_sSQL & "   T16_NENDO = " & m_iNendo & " "
                    w_sSQL = w_sSQL & " AND T16_GAKUSEI_NO = '" & w_sGakuNo & "' "
                    w_sSQL = w_sSQL & " AND T16_KAMOKU_CD = '" & w_sKamokuCd & "' "

                    iRet = gf_ExecuteSQL(w_sSQL)
                    If iRet <> 0 Then
                        '//۰��ޯ�
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_Update = 99
                        Exit For
                    End If
                End If
            Next
        Next

        '//�Я�
        Call gs_CommitTrans()

        '//����I��
        f_Update = 0
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
    <title>�l���C�I���Ȗڌ���</title>
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

		alert("<%=C_TOUROKU_OK_MSG%>");

//        document.frm.action="default2.asp";
//        document.frm.target="main";
//        document.frm.submit();

	    document.frm.target = "main";
	    document.frm.action = "./web0340_main.asp"
	    document.frm.submit();
	    return;
    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type="hidden" name="txtNendo"    value="<%=Request("txtNendo")%>">
	<input type="hidden" name="txtKyokanCd" value="<%=Request("txtKyokanCd")%>">
	<input type="hidden" name="txtGakunen"  value="<%=Request("txtGakunen")%>">
	<input type="hidden" name="txtClass"    value="<%=Request("txtClass")%>">
	<input type="hidden" name="txtKBN"      value="<%=Request("txtKBN")%>">
	<input type="hidden" name="txtGRP"      value="<%=Request("txtGRP")%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

