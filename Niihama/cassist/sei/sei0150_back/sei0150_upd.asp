<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0150_upd.asp
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
    Dim m_bErrFlg		'//�װ�׸�
	
    '�擾�����f�[�^�����ϐ�
    Dim m_sKyokanCd		'//����CD
    Dim m_iNendo
    Dim m_sSikenKBN
    Dim m_sKamokuCd
    Dim i_max
    Dim m_sGakuNo		'//�w�N
    Dim m_sGakkaCd		'//�w��
	Dim m_iSeisekiInpType
	Dim m_bZenkiOnly	'�O���J�݃t���O(True���O���J�݁AFalse���O���J�݂łȂ�)
	Dim m_SchoolFlg
	Dim m_bSeiInpFlg
	Dim m_HyokaDispFlg
	Dim m_KekkaGaiDispFlg
	
	Dim m_bNiteiFlg
	
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
	w_sMsgTitle="���ѓo�^"
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
		
		'//�F��O����擾
		if not gf_GetNintei(m_iNendo,m_bNiteiFlg) then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//���ѓo�^
		If f_Update(m_sSikenKBN) <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//�����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
		'//�O���݂̂̏ꍇ�́A�擾�����f�[�^��������������ɂ��o�^����
		If cint(m_sSikenKBN) = cint(C_SIKEN_ZEN_KIM) Then
			If m_bZenkiOnly = True Then
				'//���ѓo�^(�O���݂̂̎����Ȗڂ̏ꍇ)
				If f_Update(C_SIKEN_KOU_KIM) <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If
			End If
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
	
	m_sKyokanCd     = request("txtKyokanCd")
	m_iNendo        = request("txtNendo")
	m_sSikenKBN     = Cint(request("sltShikenKbn"))
	m_sKamokuCd     = request("KamokuCd")
	i_max           = request("i_Max")
	m_sGakuNo	    = Cint(request("txtGakuNo"))	'//�w�N
	m_sGakkaCd	    = request("txtGakkaCd")			'//�w��
	
	m_iSeisekiInpType = cint(request("hidSeisekiInpType"))
	m_bZenkiOnly = cbool(request("hidZenkiOnly"))
	m_SchoolFlg = cbool(request("hidSchoolFlg"))
	m_HyokaDispFlg = cbool(request("hidHyokaDispFlg"))
	m_KekkaGaiDispFlg = cbool(request("hidKekkaGaiDispFlg"))
	
	m_bSeiInpFlg = cbool(request("hidKikan"))
	
End Sub

'********************************************************************************
'*  [�@�\]  �w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Update(p_sSikenKBN)
	Dim i
	Dim w_Today
	Dim w_FieldName
	Dim w_DataKbnFlg
	Dim w_DataKbn
	
	On Error Resume Next
	Err.Clear
	
	f_Update = 99
	w_DataKbnFlg = false
	w_DataKbn = 0
	
	Do
		w_Today = gf_YYYY_MM_DD(m_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		
		'//���Z�敪�擾(sei0150_upd_func.asp���֐�)
		'If Not Incf_SelGenzanKbn() Then Exit Function
		
		'//���ہE���Ȑݒ�擾(sei0150_upd_func.asp���֐�)
		'If Not Incf_SelM15_KEKKA_KESSEKI() then Exit Function
		
		'//�ݐϋ敪�擾(sei0150_upd_func.asp���֐�)
		'If Not Incf_SelKanriMst(m_iNendo,C_K_KEKKA_RUISEKI) then Exit Function
		
		For i=1 to i_max
			'//�����Ǝ��Ԏ擾(sei0150_upd_func.asp���֐�)
			'Call Incs_GetJituJyugyou(i)
			
			'//�w�����̏ꍇ�A�Œ᎞�Ԃ��擾����
			'if Cint(m_sSikenKBN) = C_SIKEN_KOU_KIM then
			'	'//�Œ᎞�Ԏ擾(sei0150_upd_func.asp���֐�)
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
			
			w_sSQL = ""
			w_sSQL = w_sSQL & " UPDATE T16_RISYU_KOJIN SET "
			
			Select Case p_sSikenKBN
				
				Case C_SIKEN_ZEN_TYU
					
					if not request("hidNoChange" & i ) Then
						
						'//���l����
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & "	T16_HYOKA_TYUKAN_Z = '', "
							
							if not m_bNiteiFlg then
								w_sSQL = w_sSQL & "	T16_SEI_TYUKAN_Z   = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							end if
							
							w_sSQL = w_sSQL & "	T16_HTEN_TYUKAN_Z  = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
						'//��������
						elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & "	T16_HYOKA_TYUKAN_Z ='" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & "	T16_SEI_TYUKAN_Z   = NULL, "
							w_sSQL = w_sSQL & "	T16_HTEN_TYUKAN_Z  = NULL, "
						end if
						
						'//�]���\��\���̂Ƃ��A���ѓo�^���ԓ�
						if m_HyokaDispFlg and m_bSeiInpFlg then
							w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z	= '" & request("Hyoka"&i) & "', "
						end if
						
						w_sSQL = w_sSQL & " T16_KEKA_TYUKAN_Z		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						
						'//���ۑΏۊO�\���̂Ƃ�
						if m_KekkaGaiDispFlg then
							w_sSQL = w_sSQL & " T16_KEKA_NASI_TYUKAN_Z	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						end if
						
						w_sSQL = w_sSQL & " T16_CHIKAI_TYUKAN_Z		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
					End if
					
					'//�]���敪�̕\��������Ă���Ƃ�
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & " T16_DATAKBN_TYUKAN_Z = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & " T16_SOJIKAN_TYUKAN_Z    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
					w_sSQL = w_sSQL & " T16_JUNJIKAN_TYUKAN_Z   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					w_sSQL = w_sSQL & " T16_J_JUNJIKAN_TYUKAN_Z = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					
					w_sSQL = w_sSQL & " T16_KOUSINBI_TYUKAN_Z = '" & w_Today & "',"
					
				Case C_SIKEN_ZEN_KIM
					
					if not request("hidNoChange" & i ) Then
						
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T16_HYOKA_KIMATU_Z = '', "
							
							if not m_bNiteiFlg then
								w_sSQL = w_sSQL & " T16_SEI_KIMATU_Z   = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							end if
							
							w_sSQL = w_sSQL & " T16_HTEN_KIMATU_Z  = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
						elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & " T16_HYOKA_KIMATU_Z ='" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T16_SEI_KIMATU_Z   = NULL, "
							w_sSQL = w_sSQL & " T16_HTEN_KIMATU_Z  = NULL, "
						end if
						
						w_sSQL = w_sSQL & " T16_KEKA_KIMATU_Z		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						
						if m_KekkaGaiDispFlg then
							w_sSQL = w_sSQL & " T16_KEKA_NASI_KIMATU_Z	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						end if
						
						w_sSQL = w_sSQL & " T16_CHIKAI_KIMATU_Z		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
					End if
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & " T16_DATAKBN_KIMATU_Z = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & " T16_SOJIKAN_KIMATU_Z    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
					w_sSQL = w_sSQL & " T16_JUNJIKAN_KIMATU_Z   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					w_sSQL = w_sSQL & " T16_J_JUNJIKAN_KIMATU_Z = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					
					w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_Z = '" & w_Today & "',"
					
				Case C_SIKEN_KOU_TYU
					
					if not request("hidNoChange" & i ) Then
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T16_HYOKA_TYUKAN_K = '', "
							
							if not m_bNiteiFlg then
								w_sSQL = w_sSQL & " T16_SEI_TYUKAN_K   =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							end if
							
							w_sSQL = w_sSQL & " T16_HTEN_TYUKAN_K  =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							
						elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							
							w_sSQL = w_sSQL & " T16_HYOKA_TYUKAN_K = '" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T16_SEI_TYUKAN_K   =  NULL, "
							w_sSQL = w_sSQL & " T16_HTEN_TYUKAN_K  =  NULL, "
						end if
						
						if m_HyokaDispFlg and m_bSeiInpFlg then
							w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K	= '" & request("Hyoka"&i) & "', "
						end if
						
						w_sSQL = w_sSQL & " T16_KEKA_TYUKAN_K		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						
						if m_KekkaGaiDispFlg then
							w_sSQL = w_sSQL & " T16_KEKA_NASI_TYUKAN_K	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						end if
						
						w_sSQL = w_sSQL & " T16_CHIKAI_TYUKAN_K		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
					End if
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & " T16_DATAKBN_TYUKAN_K = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & " T16_SOJIKAN_TYUKAN_K    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
					w_sSQL = w_sSQL & " T16_JUNJIKAN_TYUKAN_K   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					w_sSQL = w_sSQL & " T16_J_JUNJIKAN_TYUKAN_K = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					
					w_sSQL = w_sSQL & " T16_KOUSINBI_TYUKAN_K = '" & w_Today & "',"
					
				Case C_SIKEN_KOU_KIM
					
					if not request("hidNoChange" & i ) Then
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T16_HYOKA_KIMATU_K = '', "
							
							if not m_bNiteiFlg then
								w_sSQL = w_sSQL & " T16_SEI_KIMATU_K   =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							end if
							
							w_sSQL = w_sSQL & " T16_HTEN_KIMATU_K  =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							
						elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							
							w_sSQL = w_sSQL & " T16_HYOKA_KIMATU_K = '" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T16_SEI_KIMATU_K   = NULL, "
							w_sSQL = w_sSQL & " T16_HTEN_KIMATU_K  = NULL, "
							w_sSQL = w_sSQL & " T16_HYOKA_FUKA_KBN = " & gf_SetNull2Zero(request("hidHyokaFukaKbn" & i)) & ", "
							
						end if
						
						w_sSQL = w_sSQL & " T16_KEKA_KIMATU_K		=  " & f_CnvNumNull(request("Kekka"&i)) & ", "
						
						if m_KekkaGaiDispFlg then
							w_sSQL = w_sSQL & " T16_KEKA_NASI_KIMATU_K	=  " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						end if
						
						w_sSQL = w_sSQL & " T16_CHIKAI_KIMATU_K		=  " & f_CnvNumNull(request("Chikai"&i)) & ", "
					End if
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & " T16_DATAKBN_KIMATU_K = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & " T16_SOJIKAN_KIMATU_K    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
					w_sSQL = w_sSQL & " T16_JUNJIKAN_KIMATU_K   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					w_sSQL = w_sSQL & " T16_J_JUNJIKAN_KIMATU_K = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					'w_sSQL = w_sSQL & " T16_SAITEI_JIKAN        = " & f_CnvNumNull(m_iSaiteiJikan) & ","
					
					w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_K = '" & w_Today & "',"
					
					'if Not gf_IsNull(m_iKyuSaiteiJikan) Then
					'	w_sSQL = w_sSQL & " T16_KYUSAITEI_JIKAN = " & f_CnvNumNull(m_iKyuSaiteiJikan) & ","
					'End if
			End Select
			
			w_sSQL = w_sSQL & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
			w_sSQL = w_sSQL & "   T16_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
			w_sSQL = w_sSQL & " WHERE "
			w_sSQL = w_sSQL & "        T16_NENDO = " & Cint(m_iNendo) & " "
			w_sSQL = w_sSQL & "    AND T16_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
			w_sSQL = w_sSQL & "    AND T16_KAMOKU_CD = '" & Trim(m_sKamokuCd) & "'  "
			
			if gf_ExecuteSQL(w_sSQL) <> 0 then Exit Do
			
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
    <title>���ѓo�^</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
	
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
	
    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //************************************************************
    function window_onload() {
		alert("<%=C_TOUROKU_OK_MSG%>");
		document.frm.target = "main";
	    document.frm.action = "sei0150_bottom.asp"
	    document.frm.submit();
	}
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE="javascript" onload="window_onload();">
    <form name="frm" method="post">
	
	<input type="hidden" name="txtNendo"     value="<%=trim(Request("txtNendo"))%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=trim(Request("txtKyokanCd"))%>">
	<input type="hidden" name="sltShikenKbn" value="<%=trim(Request("sltShikenKbn"))%>">
	<input type="hidden" name="txtGakuNo"    value="<%=trim(Request("txtGakuNo"))%>">
	<input type="hidden" name="txtClassNo"   value="<%=trim(Request("txtClassNo"))%>">
	<input type="hidden" name="txtKamokuCd"  value="<%=trim(Request("txtKamokuCd"))%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=trim(Request("txtGakkaCd"))%>">
	<input type="hidden" name="hidKamokuKbn" value="<%=request("hidKamokuKbn")%>">
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>