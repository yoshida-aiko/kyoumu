<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����������o�^
' ��۸���ID : gak/sei0400/sei0400_upd.asp
' �@      �\: ���y�[�W �����������o�^�̓o�^�A�X�V
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
' ��      ��: 2001/07/19 �O�c �q�j
' ��      �X�F2001/08/30 �ɓ� ���q     ����������2�d�ɕ\�����Ȃ��悤�ɕύX
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
    Dim     m_sKyokanCd     '//����CD
    Dim     m_sGakuNo       '//�w���ԍ�
    Dim     m_iSikenKBN
    Dim     m_sSyoken 
    Dim     m_sBikou 
    Dim     m_sSinroCd 
    Dim     m_sSRondai 
    Dim     m_sSKyokanCd1 
    Dim     m_sSKyokanCd2 
    Dim     m_sSKyokanCd3
    Dim     m_sGakunen
    Dim     m_sClass
    Dim     m_sClassNm
    Dim     m_sGakusei

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
    w_sMsgTitle="�����������o�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sKyokanCd     = session("KYOKAN_CD")
    m_sGakuNo       = request("txtGakuNo")
    m_sSyoken     = request("Syoken")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass  = Cint(request("txtClass"))
	m_sGakusei  = request("GakuseiNo")
	m_iSikenKBN = Cint(request("txtSikenKBN"))
	
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
    dim w_sSikenKBN

    On Error Resume Next
    Err.Clear
    
    f_Update = 1

	select case m_iSikenKBN
		case C_SIKEN_ZEN_TYU
			w_sSikenKBN = "T13_SYOKEN_TYUKAN_Z"
		case C_SIKEN_ZEN_KIM
			w_sSikenKBN = "T13_SYOKEN_KIMATU_Z"
		case C_SIKEN_KOU_TYU
			w_sSikenKBN = "T13_SYOKEN_TYUKAN_K"
		case C_SIKEN_KOU_KIM
			w_sSikenKBN = "T13_SYOKEN_KIMATU_K"
	End select

    Do 

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

            '//T11_GAKUSEKI��UPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN SET "
            w_sSQL = w_sSQL & vbCrLf & "   " & w_sSikenKBN & "= '"  & Trim(m_sSyoken) & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/") & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_USER = '"    & Session("LOGIN_ID")       & "'"
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T13_GAKUSEI_NO = '" & m_sGakuNo & "'  "
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
    <title>�����������o�^</title>
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

		<%
		'//�o�^�{�^���������A������ʂɖ߂�
		If trim(Request("GakuseiNo")) = "" Then%>
	        document.frm.action="default.asp";
			document.frm.target="<%=C_MAIN_FRAME%>";
		<%
		'//�O��OR���փ{�^���������A�����������͉�ʂ�\������
		Else %>
    	    document.frm.action="sei0400_main.asp";
        	document.frm.target="main";
		<%End If %>
        document.frm.submit();

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

		<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
		<input type="hidden" name="GakuseiNo" value="<%=m_sGakusei%>">
		<input type="hidden" name="txtClass" value="<%=m_sClass%>">
		<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
		<input type="hidden" name="txtSikenKBN" value="<%=m_iSikenKBN%>">

		<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

