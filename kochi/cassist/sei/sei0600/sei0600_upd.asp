<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ȓ����o�^
' ��۸���ID : gak/sei0600/sei0600_upd.asp
' �@      �\: ���y�[�W ���ȓ����̓o�^�A�X�V
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
' ��      ��: 2001/09/26 �J�e �ǖ�
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
    Dim     m_iSikenKBN

    Public  m_iMax          '�ő�y�[�W
    Public  m_iDsp          '�ꗗ�\���s��
	Public  m_rCnt

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
    w_sMsgTitle="���ȓ����o�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_iSikenKBN = Cint(request("txtSikenKBN"))
	m_rCnt = cint(Request("txtCnt"))

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
        w_iRet = f_Update()
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

Function f_Update()
'********************************************************************************
'*  [�@�\]  �w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
    dim w_sSikenKBN_KE,w_sSikenKBN_KI
	dim w_Gno,w_Kss,w_kbk

    On Error Resume Next
    Err.Clear
    
    f_Update = 1

	select case cint(m_iSikenKBN)
		case C_SIKEN_ZEN_TYU '�O������
			w_sSikenKBN_KE = "T13_KESSEKI_TYUKAN_Z"
			w_sSikenKBN_KI = "T13_KIBIKI_TYUKAN_Z"
		case C_SIKEN_ZEN_KIM '�O������
			w_sSikenKBN_KE = "T13_KESSEKI_KIMATU_Z"
			w_sSikenKBN_KI = "T13_KIBIKI_KIMATU_Z"
		case C_SIKEN_KOU_TYU '�������
			w_sSikenKBN_KE = "T13_KESSEKI_TYUKAN_K"
			w_sSikenKBN_KI = "T13_KIBIKI_TYUKAN_K"
		case C_SIKEN_KOU_KIM '��������i�w�N���j
			w_sSikenKBN_KE = "T13_SUMKESSEKI"
			w_sSikenKBN_KI = "T13_SUMKIBTEI"
	End select

    Do 

	For i = 1 to m_rCnt 
	
        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()
		w_Gno = "txtGAKUSEINO_"&i
		w_Kss = "txtKESSEKI_"&i
		w_kbk = "txtKIBIKI_"&i

            '//T11_GAKUSEKI��UPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN SET "
            w_sSQL = w_sSQL & vbCrLf & "   " & w_sSikenKBN_KE & "= '"  & gf_SetNull2Zero(Request(w_Kss)) & "',"
            w_sSQL = w_sSQL & vbCrLf & "   " & w_sSikenKBN_KI & "= '"  & gf_SetNull2Zero(Request(w_kbk)) & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/") & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_USER = '"    & Session("LOGIN_ID")       & "'"
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T13_GAKUSEI_NO = '" & Request(w_Gno) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "   AND  T13_NENDO = " & session("NENDO") & "  "
            
            iRet = gf_ExecuteSQL(w_sSQL)
            If iRet <> 0 Then
                '//۰��ޯ�
                Call gs_RollbackTrans()
                msMsg = Err.description
                f_Update = 99
                Exit Do
            End If

        '//�Я�
        Call gs_CommitTrans()
	Next
	
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
    <title>���ȓ����o�^</title>
    <link rel=stylesheet href="../../common/style.css" type=text/css>

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
		parent.topFrame.location.href = "white.htm";

        document.frm.action="default.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

		<input type="hidden" name="txtSikenKBN" value="<%=m_iSikenKBN%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

