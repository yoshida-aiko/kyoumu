<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ������w���ѓo�^
' ��۸���ID : sei/sei0100/sei0160_upd.asp
' �@      �\: ���y�[�W ������w���ѓo�^�̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
' ��      ��:
' ��      �n:
' ��      ��:
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Dim m_bErrFlg		'//�װ�׸�
	
    '�擾�����f�[�^�����ϐ�

    Dim m_iNendo				'//�N�x
    Dim m_sKyokanCd				'//�����R�[�h
    Dim m_sGakunen				'//�w�N
    Dim m_sClass				'//�N���X

    Dim m_sBunruiCD		 		'//���ރR�[�h
    Dim m_sBunruiNM		 		'//���ޖ���
    Dim m_sTani		 			'//�P��
    Dim i_max

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()
	Dim w_iRet              '// �߂�l
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="������w���ѓo�^"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = False
	
	Do
		'//�ް��ް��ڑ�
		if gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		end If
		
		'�f�[�^�擾
		Call s_SetParam()
		
		'//�s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'//�g�����U�N�V�����J�n
		Call gs_BeginTrans()
		
		'//���C�F��e�[�u���폜����
		if f_Delete() <> 0 Then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//���C�F��e�[�u���X�V����
		If f_Update() <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'// �y�[�W��\��
		Call showPage()
		
		Exit Do
	Loop
	
    '//�װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        '//���[���o�b�N
        Call gs_RollbackTrans()
        
        w_sMsg = gf_GetErrMsg()
        'response.write "w_sMsg =" & w_sMsg & "<BR>"
        'response.end
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    else
    	'//�R�~�b�g
    	Call gs_CommitTrans()
    End If
    
    '// �I������
    Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()

    	m_iNendo    	= request("txtNendo")		'�N�x
    	m_sKyokanCd 	= request("txtKyokanCd")	'�����R�[�h


	m_sGakunen  	= request("txtGakunen")         '�w�N
	m_sClass    	= request("txtClass")           '�N���X
	m_sBunruiCD 	= request("txtBunruiCd")	'���ރR�[�h
	m_sBunruiNm 	= request("txtBunruiNm")	'���ޖ���
	m_sTani     	= request("txtTani")		'�P��

	i_max           = request("i_Max")

End Sub

'********************************************************************************
'*  [�@�\]  ���C�F��e�[�u��(T100_RISYU_NINTEI)�폜�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Delete()
	Dim i
	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Delete = 99
	
	Do

		For i=1 to i_max

			w_sSQL = ""
			w_sSQL = w_sSQL & " DELETE"
			w_sSQL = w_sSQL & " FROM"
			w_sSQL = w_sSQL & " T100_RISYU_NINTEI"
			w_sSQL = w_sSQL & " WHERE"
			w_sSQL = w_sSQL & "      T100_GAKUSEI_NO = '" & request("txtGseiNo"&i)  & "'"	
			w_sSQL = w_sSQL & " AND  T100_BUNRUI_CD = '" & m_sBunruiCD & "'"
		
			'�폜
			if gf_ExecuteSQL(w_sSQL) <> 0 then Exit Do
 'response.write w_sSQL
		Next
		
		'//����I��
		f_Delete = 0
		
		Exit Do
	Loop
	
End Function


'********************************************************************************
'*  [�@�\]  ���C�F��e�[�u��(T100_RISYU_NINTEI)�X�V�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Update()
	Dim i
	Dim w_sSQL
	Dim w_Seiseki
	Dim w_Hyoka
	Dim w_sHyokaFuka
	Dim w_SumiTani

	
	On Error Resume Next
	Err.Clear
	
	f_Update = 99
	
	Do

		For i=1 to i_max

			w_Seiseki = gf_SetNull2String(request("Seiseki"&i))
			w_Hyoka   = gf_SetNull2String(request("hidHyoka"&i))

			'���сE�]�������͂��ꂽ�Ƃ��̂ݓo�^
			if w_Seiseki <> "" or w_Hyoka <> "" then 

				'�C���N�x�����͂���Ă���΁A�C���Ƃ݂Ȃ� 
				If request("SyuNendo"&i) <> "" Then
					w_sHyokaFuka = "0"	'���i
					w_SumiTani = m_sTani	'�ϒP��
				Else
					w_sHyokaFuka = "1"	'�s���i
					w_SumiTani = "0"	
				End If

				w_sSQL = ""
				w_sSQL = w_sSQL & " INSERT INTO T100_RISYU_NINTEI "
				w_sSQL = w_sSQL & " ("
				w_sSQL = w_sSQL & " T100_GAKUSEI_NO"
				w_sSQL = w_sSQL & ", T100_BUNRUI_CD"
				'w_sSql = w_sSql & ", T100_KYU_CD"
				w_sSQL = w_sSQL & ", T100_SYUTOKU_NENDO"
				w_sSQL = w_sSQL & ", T100_SYUTOKU_GAKUNEN"
				w_sSQL = w_sSQL & ", T100_GAKUSEKI_NO"
				w_sSQL = w_sSQL & ", T100_GAKKA_CD"
				w_sSQL = w_sSQL & ", T100_CLASS"
				w_sSQL = w_sSQL & ", T100_COURSE_CD"
				w_sSQL = w_sSQL & ", T100_BUNRUI_MEISYO"
				'w_sSql = w_sSql & ", T100_KYU_MEI"
				w_sSQL = w_sSQL & ", T100_HAITOTANI"
				w_sSQL = w_sSQL & ", T100_TANI_SUMI"
				'w_sSql = w_sSql & ", T100_HYOTEI"
				w_sSQL = w_sSQL & ", T100_HYOKA"
				w_sSQL = w_sSQL & ", T100_HYOKA_FUKA_KBN"
				'w_sSql = w_sSql & ", T100_NINTEIBI"
				w_sSQL = w_sSQL & ", T100_SEISEKI"
				w_sSQL = w_sSQL & ", T100_INS_DATE"
				w_sSQL = w_sSQL & ", T100_INS_USER"

				w_sSQL = w_sSQL & " )VALUES("

				w_sSQL = w_sSQL & " '" & request("txtGseiNo"&i)  & "'"		'T100_GAKUSEI_NO"
				w_sSQL = w_sSQL & ",'" & m_sBunruiCD & "'"			'T100_BUNRUI_CD"
				'w_sSql = w_sSql & ", T100_KYU_CD"
				w_sSQL = w_sSQL & ", " & f_CnvNumNull(request("SyuNendo"&i)) 	'T100_SYUTOKU_NENDO"
				w_sSQL = w_sSQL & ", " & m_sGakunen				'T100_SYUTOKU_GAKUNEN"
				w_sSQL = w_sSQL & ",'" & request("txtGsekiNo"&i) & "'"		'T100_GAKUSEKI_NO"
				w_sSQL = w_sSQL & ",'" & request("txtGakkaCD"&i) & "'"		'T100_GAKKA_CD"
				w_sSQL = w_sSQL & ",'" & request("txtClass"&i) 	& "'"		'T100_CLASS"
				w_sSQL = w_sSQL & ",'" & request("txtCorceCD"&i) & "'"		'T100_COURSE_CD"
				w_sSQL = w_sSQL & ",'" & m_sBunruiNm & "'"			'T100_BUNRUI_MEISYO"
				'w_sSql = w_sSql & ", T100_KYU_MEI"
				w_sSQL = w_sSQL & ", " & m_sTani				'T100_HAITOTANI"
				w_sSQL = w_sSQL & ", " & w_SumiTani 				'T100_TANI_SUMI"
				'w_sSql = w_sSql & ", T100_HYOTEI"
				w_sSQL = w_sSQL & ",'" & w_Hyoka & "'"				'T100_HYOKA"
				w_sSQL = w_sSQL & ", " & w_sHyokaFuka 				'T100_HYOKA_FUKA_KBN"
				'w_sSql = w_sSql & ", T100_NINTEIBI"
				w_sSQL = w_sSQL & ", " & f_CnvNumNull(w_Seiseki)		'T100_SEISEKI
				w_sSQL = w_sSQL & ",'" & gf_YYYY_MM_DD(date(),"/") & "'" 	'T100_INS_DATE"
				w_sSQL = w_sSQL & ",'" & Trim(Session("LOGIN_ID")) & "'" 	'T100_INS_USER"
				w_sSQL = w_sSQL & " )"
'response.write w_sSQL

				'���s
				if gf_ExecuteSQL(w_sSQL) <> 0 then Exit Do
			End If
		Next
		'//����I��
		f_Update = 0
 		
		Exit Do
	Loop
	
End Function

'********************************************************************************
'*  [�@�\]  ���l�^���ڂ̍X�V���̐ݒ�
'*  [����]  �l
'*  [�ߒl]  �Ȃ�
'*  [����]  ���l�������Ă���ꍇ��[�l]�A�����ꍇ��"NULL"��Ԃ�
'********************************************************************************
Function f_CnvNumNull(p_vAtai)

	If Trim(p_vAtai) = "" Then
		f_CnvNumNull = "NULL"
	Else
		f_CnvNumNull = cInt(p_vAtai)
    End If

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'********************************************************************************
%>
    <html>
    <head>
    <title>������w���ѓo�^</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
	
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
	
    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //************************************************************
    function window_onload() {
	alert("<%=C_TOUROKU_OK_MSG%>");
	document.frm.action="sei0160_top.asp";
	document.frm.target="topFrame";
	document.frm.submit();
	document.frm.target = "main";
	document.frm.action = "sei0160_bottom.asp"
	document.frm.submit();
	}
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE="javascript" onload="window_onload();">
    <form name="frm" method="post">
	
	<input type="hidden" name="txtNendo"     value="<%=trim(Request("txtNendo"))%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=trim(Request("txtKyokanCd"))%>">
	<input type="hidden" name="txtGakunen"   value="<%=trim(Request("txtGakunen"))%>">
	<input type="hidden" name="txtClass"     value="<%=trim(Request("txtClass"))%>">
	<input type="hidden" name="txtBunruiCd"  value="<%=trim(Request("txtBunruiCd"))%>">
	<input type="hidden" name="txtBunruiNm"  value="<%=trim(Request("txtBunruiNm"))%>">
	<input type="hidden" name="txtTani"      value="<%=trim(Request("txtTani"))%>">
	<input type="hidden" name="i_Max"        value="<%=request("i_Max")%>">
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>