<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������������o�^
' ��۸���ID : gak/gak0460/gak0460_upd.asp
' �@      �\: ���y�[�W �������������o�^�̓o�^�A�X�V
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
' ��      �X�F2002/10/08 �A�c �k��Y   �S�C�����A���i���̍��ڂ�ǉ�
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
    Dim     m_sSGSyoken 
    Dim     m_sBikou 
    Dim     m_sTanninSyoken  '�S�C����		'2002.10.08 Add hirota
    Dim     m_sTanninBikou   '���i��  		'2002.10.08 Add hirota
	Dim     m_sNendo         '�����N�x�@�@�@'2002.10.08 Add hirota
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
    w_sMsgTitle="�������������o�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sKyokanCd     = session("KYOKAN_CD")
    m_sGakuNo       = request("txtGakuNo")
    m_sSGSyoken     = request("SGSyoken")
    m_sBikou        = request("Bikou")
    m_sSGSyoken     = request("SGSyoken")
    m_sTanninSyoken = request("TanninSyoken")		'2002.10.08 Add hirota
    m_sTanninBikou  = request("TanninBikou")		'2002.10.08 Add hirota
	m_sNendo        = request("txtNendo")			'2002.10.08 Add hirota
    m_sSRondai      = request("SRondai")
    m_sSKyokanCd1   = request("SKyokanCd1")
    m_sSKyokanCd2   = request("SKyokanCd2")
    m_sSKyokanCd3   = request("SKyokanCd3")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass  = Cint(request("txtClass"))
	m_sClassNm  = request("txtClassNm")
	m_sGakusei  = request("GakuseiNo")
    m_iDsp = C_PAGE_LINE

    Do
        '// �ް��ް��ڑ�
        If gf_OpenDatabase() <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '//�w���v�^�����X�V
		Call gs_BeginTrans()		'�g�����U�N�V�����J�n         	'2002.10.08 hirota

        If f_Update() <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		if Not gf_GetGakkoNO(w_sGakkoNO) then
            m_bErrFlg = True
            Exit Do
		end if

		if w_sGakkoNO = cstr(C_NCT_KUMAMOTO) then

	        If f_Update_T13() <> 0 Then
	            m_bErrFlg = True
	            Exit Do
	        End If

		end if

		Call gs_CommitTrans() 		'�g�����U�N�V�������R�~�b�g   	'2002.10.08 hirota

        '// �y�[�W��\��
        Call showPage()

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
		Call gs_RollbackTrans() 	'�g�����U�N�V���������[���o�b�N '2002.10.08 hirota
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Update()
	
	On Error Resume Next
	Err.Clear
	
	f_Update = 1
	
	Do
		'//T11_GAKUSEKI��UPDATE
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " UPDATE T11_GAKUSEKI SET "
		w_sSQL = w_sSQL & vbCrLf & "   T11_SOGOSYOKEN = '"  & Trim(m_sSGSyoken) & "' ,"
		w_sSQL = w_sSQL & vbCrLf & "   T11_KOJIN_BIK = '"  & Trim(m_sBikou) & "' ,"
		
		If m_sGakunen = 5 Then
			w_sSQL = w_sSQL & vbCrLf & "   T11_SINRO = '"  & Trim(m_sSinroCd) & "' ,"
			w_sSQL = w_sSQL & vbCrLf & "   T11_SOTUKEN_DAI = '"  & Trim(m_sSRondai) & "' ,"
			w_sSQL = w_sSQL & vbCrLf & "   T11_SOTU_KYOKAN_CD1 = '"  & Trim(m_sSKyokanCd1) & "' ,"
			w_sSQL = w_sSQL & vbCrLf & "   T11_SOTU_KYOKAN_CD2 = '"  & Trim(m_sSKyokanCd2) & "' ,"
			w_sSQL = w_sSQL & vbCrLf & "   T11_SOTU_KYOKAN_CD3 = '"  & Trim(m_sSKyokanCd3) & "', "
		End If
		
		w_sSQL = w_sSQL & vbCrLf & "   T11_UPD_DATE = '"  & gf_YYYY_MM_DD(date(),"/") & "', "
		w_sSQL = w_sSQL & vbCrLf & "   T11_UPD_USER = '"  & Session("LOGIN_ID") & "' "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "        T11_GAKUSEI_NO = '" & m_sGakuNo & "'  "
		
		If gf_ExecuteSQL(w_sSQL) <> 0 Then
			msMsg = Err.description
			f_Update = 99
			Exit Do
		End If
		
		'//����I��
		f_Update = 0
		Exit Do
	Loop
	
End Function

'********************************************************************************
'*  [�@�\]  T13�X�V����(�S������,���i��)
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'*  [�쐬]  �A�c : 2002.10.08
'********************************************************************************
Function f_Update_T13()
	
	On Error Resume Next
	Err.Clear
	
	f_Update_T13 = 1
	
	Do
		'//T13_GAKU_NEN��UPDATE
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN SET "
		w_sSQL = w_sSQL & vbCrLf & "   T13_TANNINSYOKEN = '"  & Trim(m_sTanninSyoken) & "' ,"
		w_sSQL = w_sSQL & vbCrLf & "   T13_TANNIN_BIK = '"  & Trim(m_sTanninBikou) & "' ,"

		w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_DATE = '"  & gf_YYYY_MM_DD(date(),"/") & "', "
		w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_USER = '"  & Session("LOGIN_ID") & "' "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKUSEI_NO = '" & m_sGakuNo & "'  "
		w_sSQL = w_sSQL & vbCrLf & "   AND T13_NENDO = " & m_sNendo

		If gf_ExecuteSQL(w_sSQL) <> 0 Then
			msMsg = Err.description
			f_Update_T13 = 99
			Exit Do
		End If
		
		'//����I��
		f_Update_T13 = 0
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
    <title>�������������o�^</title>
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
    	    document.frm.action="gak0460_main.asp";
        	document.frm.target="main";
		<%End If %>
        document.frm.submit();

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">
		<input type="hidden" name="txtNendo" value="<%=request("txtNendo")%>">
		<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
		<input type="hidden" name="GakuseiNo" value="<%=m_sGakusei%>">
		<input type="hidden" name="txtClass" value="<%=m_sClass%>">
		<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
		
		<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

