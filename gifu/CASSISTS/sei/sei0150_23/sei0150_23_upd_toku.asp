<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0150_upd_tuku.asp
' �@      �\: ���y�[�W ���ѓo�^�̓o�^�A�X�V
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
    Public  m_bErrFlg           '�װ�׸�
	
    '�擾�����f�[�^�����ϐ�
    Dim m_sKyokanCd     '//����CD
    Dim m_iNendo 
    Dim m_sSikenKBN
    Dim m_sKamokuCd
    Dim i_max 
    Dim m_sGakuNo	'//�w�N
    Dim m_sGakkaCd	'//�w��
	
	Dim m_SchoolFlg
	Dim m_KekkaGaiDispFlg
	Dim m_iSeisekiInpType
	
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


    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���ѓo�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
	
    Do
        '// �ް��ް��ڑ�
        If gf_OpenDatabase() <> 0 Then
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		Call s_SetParam()
		
		'//�g�����U�N�V�����J�n
		Call gs_BeginTrans()
		
		If f_Update(m_sSikenKBN) <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If
		
        '// �y�[�W��\��
        Call showPage()
		
        Exit Do
    Loop
	
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
	    '//���[���o�b�N
        Call gs_RollbackTrans()
        
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
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
	
	m_sKyokanCd     = request("txtKyokanCd")
	m_iNendo        = request("txtNendo")
	m_sSikenKBN     = Cint(request("sltShikenKbn"))
	m_sKamokuCd     = request("KamokuCd")
	i_max           = request("i_Max")
	m_sGakuNo	= Cint(request("txtGakuNo"))	'//�w�N
	m_sGakkaCd	= request("txtGakkaCd")			'//�w��
	
	m_iSeisekiInpType = cint(request("hidSeisekiInpType"))
	m_SchoolFlg = cbool(request("hidSchoolFlg"))
	m_KekkaGaiDispFlg = cbool(request("hidKekkaGaiDispFlg"))
	
End Sub


'********************************************************************************
'*  [�@�\]  �w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Update(p_sSikenKBN)
	Dim i,w_Today
	Dim w_DataKbnFlg
	Dim w_DataKbn
	
    On Error Resume Next
    Err.Clear
    
    f_Update = 1
	w_DataKbnFlg = false
	w_DataKbn = 0
	
    Do 
		w_Today = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")
		
		'// ���Z�敪�擾(sei0150_upd_func.asp���֐�)
		'If Not Incf_SelGenzanKbn() Then Exit Function

		'// ���ہE���Ȑݒ�擾(sei0150_upd_func.asp���֐�)
		'If Not Incf_SelM15_KEKKA_KESSEKI() then Exit Function

		'// �ݐϋ敪�擾(sei0150_upd_func.asp���֐�)
		'If Not Incf_SelKanriMst(m_iNendo,C_K_KEKKA_RUISEKI) then Exit Function

		For i=1 to i_max

			'// �����Ǝ��Ԏ擾(sei0150_upd_func.asp���֐�)
			'Call Incs_GetJituJyugyou(i)

			'// �w�����̏ꍇ�A�Œ᎞�Ԃ��擾����
			'if Cint(m_sSikenKBN) = C_SIKEN_KOU_KIM then
			'	'// �Œ᎞�Ԏ擾(sei0150_upd_func.asp���֐�)
			'	If Not Incf_GetSaiteiJikan(i) then Exit Function
			'End if
			
			'//�]���s�\�`�F�b�N(�F�{�d�g�̂�)
			if m_SchoolFlg = true then
				w_DataKbn = 0
				w_DataKbnFlg = false
				
				'//���]���A�]���s�\�̐ݒ�
				if cint(gf_SetNull2Zero(request("hidMihyoka"))) <> 0 then
					w_DataKbn = cint(gf_SetNull2Zero(request("hidMihyoka")))
					w_DataKbnFlg = true
				else
					w_DataKbn = cint(gf_SetNull2Zero(request("chkHyokaFuno" & i)))
					
					if w_DataKbn = cint(C_HYOKA_FUNO) then
						w_DataKbnFlg = true
					end if
				end if
			end if
			
			
			'//T16_RISYU_KOJIN��UPDATE
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " UPDATE T34_RISYU_TOKU SET "
			
			Select Case p_sSikenKBN
				Case C_SIKEN_ZEN_TYU
					if not request("hidNoChange" & i ) Then
						
						'//���l����
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & "	T34_HYOKA_TYUKAN_Z = '', "
							w_sSQL = w_sSQL & "	T34_SEI_TYUKAN_Z   = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							
						'//��������
						elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & "	T34_HYOKA_TYUKAN_Z ='" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & "	T34_SEI_TYUKAN_Z   = NULL, "
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_TYUKAN_Z		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						
						'//���ۑΏۊO�\���̂Ƃ�
						if m_KekkaGaiDispFlg then
							w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_NASI_TYUKAN_Z		= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T34_CHIKAI_TYUKAN_Z		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
					End if
					
					w_sSQL = w_sSQL & vbCrLf & " 	T34_SOJIKAN_TYUKAN_Z    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T34_JUNJIKAN_TYUKAN_Z   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T34_J_JUNJIKAN_TYUKAN_Z = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					
					w_sSQL = w_sSQL & vbCrLf & " 	T34_KOUSINBI_TYUKAN_Z = '" & w_Today & "',"
					
					'//�]���敪�̕\��������Ă���Ƃ�
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & " T34_DATAKBN_TYUKAN_Z = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
				Case C_SIKEN_ZEN_KIM
					if not request("hidNoChange" & i ) Then
						
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T34_HYOKA_KIMATU_Z = '', "
							w_sSQL = w_sSQL & " T34_SEI_KIMATU_Z   = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
						elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & " T34_HYOKA_KIMATU_Z ='" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T34_SEI_KIMATU_Z   = NULL, "
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_KIMATU_Z		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						
						if m_KekkaGaiDispFlg then
							w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_NASI_KIMATU_Z		= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T34_CHIKAI_KIMATU_Z		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
					End if
					w_sSQL = w_sSQL & vbCrLf & " 	T34_SOJIKAN_KIMATU_Z    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T34_JUNJIKAN_KIMATU_Z   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T34_J_JUNJIKAN_KIMATU_Z = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					
					w_sSQL = w_sSQL & vbCrLf & " 	T34_KOUSINBI_KIMATU_Z = '" & w_Today & "',"
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & " T34_DATAKBN_KIMATU_Z = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
				Case C_SIKEN_KOU_TYU
					if not request("hidNoChange" & i ) Then
						
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T34_HYOKA_TYUKAN_K = '', "
							w_sSQL = w_sSQL & " T34_SEI_TYUKAN_K   =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
						elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & " T34_HYOKA_TYUKAN_K = '" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T34_SEI_TYUKAN_K   =  NULL, "
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_TYUKAN_K		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						
						if m_KekkaGaiDispFlg then
							w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_NASI_TYUKAN_K		= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T34_CHIKAI_TYUKAN_K		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
					End if
					w_sSQL = w_sSQL & vbCrLf & " 	T34_SOJIKAN_TYUKAN_K    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T34_JUNJIKAN_TYUKAN_K   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T34_J_JUNJIKAN_TYUKAN_K = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					
					w_sSQL = w_sSQL & vbCrLf & " 	T34_KOUSINBI_TYUKAN_K = '" & w_Today & "',"
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & " T34_DATAKBN_TYUKAN_K = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
				Case C_SIKEN_KOU_KIM
					if not request("hidNoChange" & i ) Then
						
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T34_HYOKA_KIMATU_K = '', "
							w_sSQL = w_sSQL & " T34_SEI_KIMATU_K   =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
						elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & " T34_HYOKA_KIMATU_K = '" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T34_SEI_KIMATU_K   = NULL, "
							w_sSQL = w_sSQL & " T34_HYOKA_FUKA_KBN = " & gf_SetNull2Zero(request("hidHyokaFukaKbn" & i)) & ", "
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_KIMATU_K		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						
						if m_KekkaGaiDispFlg then
							w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_NASI_KIMATU_K		= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T34_CHIKAI_KIMATU_K		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
					End if
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & " T34_DATAKBN_KIMATU_K = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & vbCrLf & " 	T34_SOJIKAN_KIMATU_K    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T34_JUNJIKAN_KIMATU_K   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T34_J_JUNJIKAN_KIMATU_K = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					'w_sSQL = w_sSQL & vbCrLf & " 	T34_SAITEI_JIKAN        = " & f_CnvNumNull(m_iSaiteiJikan) & ","
					
					w_sSQL = w_sSQL & vbCrLf & " 	T34_KOUSINBI_KIMATU_K = '" & w_Today & "',"
					
					'if Not gf_IsNull(m_iKyuSaiteiJikan) Then
					'	w_sSQL = w_sSQL & vbCrLf & " 	T34_KYUSAITEI_JIKAN = " & f_CnvNumNull(m_iKyuSaiteiJikan) & ","
					'End if
					
					
					
			End Select
			
			w_sSQL = w_sSQL & vbCrLf & "   T34_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
			w_sSQL = w_sSQL & vbCrLf & "   T34_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T34_NENDO = " & Cint(m_iNendo) & " "
            w_sSQL = w_sSQL & vbCrLf & "    AND T34_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "    AND T34_TOKUKATU_CD = '" & Trim(m_sKamokuCd) & "'  "
			
            If gf_ExecuteSQL(w_sSQL) <> 0 Then
                '//۰��ޯ�
                msMsg = Err.description
                f_Update = 99
                Exit Do
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
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
    <html>
    <head>
    <title>���ѓo�^</title>
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

		alert("<%=C_TOUROKU_OK_MSG%>");

		document.frm.target = "main";
		document.frm.action = "sei0150_23_print.asp" //UPP 2005/10/03 ����
	    document.frm.submit();
	    return;

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type=hidden name=txtNendo    value="<%=trim(Request("txtNendo"))%>">
	<input type=hidden name=txtKyokanCd value="<%=trim(Request("txtKyokanCd"))%>">
	<input type=hidden name=sltShikenKbn value="<%=trim(Request("sltShikenKbn"))%>">
	<input type=hidden name=txtGakuNo   value="<%=trim(Request("txtGakuNo"))%>">
	<input type=hidden name=txtClassNo  value="<%=trim(Request("txtClassNo"))%>">
	<input type=hidden name=txtKamokuCd value="<%=trim(Request("txtKamokuCd"))%>">
	<input type=hidden name=txtGakkaCd  value="<%=trim(Request("txtGakkaCd"))%>">
	<input type="hidden" name="hidKamokuKbn" value="<%=request("hidKamokuKbn")%>">
	
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>
