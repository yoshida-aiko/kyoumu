<%

'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0800/sei0800_upd_func.asp
' �@      �\: ���y�[�W ���ѓo�^�̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: 1.�����Ǝ��ԂƏ����Ǝ��Ԃ�������Ǝ��Ԃ��Z�o
'             2.�Œ᎞�Ԃ̌v�Z
'-------------------------------------------------------------------------
' ��      ��: 2002/03/27 ���`�i�K
' ��      �X: 
' �f�o�b�O  : ���Z�敪�擾��WHERE�����̃R���X�g���킩��Ȃ�
'*************************************************************************/

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
Dim m_iSouJyugyou			'������
Dim m_iJunJyugyou			'������
Dim m_iJituJyugyou			'������
Dim m_iSouJituJyugyou		'��������
Dim m_iGenzan				'���Z�敪(0:���Ȃ�,1:����)
Dim m_iRuisekiKbn			'�ݐϋ敪(0:������,1:�ݐ�)
Dim m_iSaiteiJikan			'�Œ᎞��
Dim m_iKyuSaiteiJikan		'�x�w�Œ᎞��
Dim mKojinRs				'���k���ں��޾��
Dim mM15Rs					'���ہE���Ȑݒ�ں��޾��
Dim m_sUpdMode				'�����ް�Ӱ�ށiTUJO:�ʏ�, TOKU:���ʁj

const C_UPDMODE_TUJO = "TUJO"
const C_UPDMODE_TOKU = "TOKU"

Dim m_bDebugFlg				'۰�����ޯ���׸�

	m_bDebugFlg = false

	'// ���Ұ��擾
	m_iSouJyugyou = request("hidSouJyugyou")		'// ������
	m_iJunJyugyou = request("hidJunJyugyou")		'// ������
	m_sUpdMode    = request("hidUpdMode")			'// �����ް�Ӱ��

'********************************************************************************
'*  [�@�\]  �f�o�b�O�\��
'********************************************************************************
Sub Incs_ShowDebug()

	On Error Resume Next
	Err.Clear

	mM15Rs.MoveFirst

	response.write "<BR>*********** ���ޯ��Ӱ�� **************<br>"
	response.write "������ = " & m_iSouJyugyou & "<BR>"
	response.write "������ = " & m_iJunJyugyou & "<BR>"
	response.write "���Z�敪 = " & m_iGenzan & "<BR>"
	response.write "������ = " & m_iJituJyugyou & "<BR>"

	response.write "<BR>--- �Œ᎞�Ԍv�Z ----<br>"
	response.write "M15 �敪 = " & mM15Rs("M15_KEKKA_KBN") & "<BR>"
	response.write "�ݐϋ敪 = " & m_iRuisekiKbn & "<BR>"
	response.write "<BR>"

	Select Case Cint(mM15Rs("M15_KEKKA_KBN"))
		Case 0
			response.write "// �Œ᎞�Œ�<br>"
			response.write "�Œ᎞�� = " & m_iSaiteiJikan & "<BR>"

		Case 1
			response.write "<BR>// �Œ᎞�ʏ�<br>"
			response.write "�������� = " & m_iSouJituJyugyou & "<BR>"
			response.write "���q = " & mM15Rs("M15_BUNSHI") & "<BR>"
			response.write "���� = " & mM15Rs("M15_BUNBO") & "<BR>"
			response.write "�[���敪 = " & mM15Rs("M15_HASUU_KBN") & "<br>"
			response.write "�Œ᎞�� = " & m_iSaiteiJikan & "<BR>"

			mM15Rs.MoveNext
			if Not mM15Rs.Eof Then
				response.write "<BR>// �Œ᎞�x�w<br>"
				response.write "M15 �敪 = " & mM15Rs("M15_KEKKA_KBN") & "<BR>"
				response.write "�ݐϋ敪 = " & m_iRuisekiKbn & "<BR>"
				response.write "<BR>"
				response.write "�������� = " & m_iSouJituJyugyou & "<BR>"
				response.write "���q = " & mM15Rs("M15_BUNSHI") & "<BR>"
				response.write "���� = " & mM15Rs("M15_BUNBO") & "<BR>"
				response.write "�[���敪 = " & mM15Rs("M15_HASUU_KBN") & "<br>"
				response.write "�Œ᎞�� = " & m_iKyuSaiteiJikan & "<BR>"
			End if

	End Select

End Sub

'********************************************************************************
'*  [�@�\]  ���Z�敪�擾
'********************************************************************************
Function Incf_SelGenzanKbn()

	On Error Resume Next
	Err.Clear

	Incf_SelGenzanKbn = False

	wSql = ""
	wSql = wSql & " SELECT "
	wSql = wSql & " 	M15_KEKKA_KESSEKI.M15_GENZAN_KBN "
	wSql = wSql & " FROM M15_KEKKA_KESSEKI "
	wSql = wSql & " WHERE "
	wSql = wSql & " 	M15_KEKKA_KESSEKI.M15_NENDO     = " & m_iNendo
	wSql = wSql & " AND M15_KEKKA_KESSEKI.M15_KEKKA_CD  = 4 "				'�H�H�H�v�f�o�b�O
	wSql = wSql & " AND M15_KEKKA_KESSEKI.M15_KEKKA_KBN = " & C_K_KEKKA_NASI
	wSql = wSql & " AND M15_KEKKA_KESSEKI.M15_TANI      = " & C_K_KEKKA_TANI_NASI

	iRet = gf_GetRecordset(wRs, wSql)
	If iRet <> 0 Then
		m_sErrMsg = Err.description
	    Call gf_closeObject(wRs)
		Exit Function
	End If

	m_iGenzan = wRs("M15_GENZAN_KBN")

    Call gf_closeObject(wRs)
	Incf_SelGenzanKbn = True

End Function

'********************************************************************************
'*  [�@�\]  �����Ǝ��Ԏ擾
'********************************************************************************
Sub Incs_GetJituJyugyou(pNo)

	On Error Resume Next
	Err.Clear

	if Cint(m_iGenzan) = 0 then
		m_iJituJyugyou = m_iJunJyugyou
	Else
		m_iJituJyugyou = m_iJunJyugyou - gf_SetNull2Zero(request("KekkaGai" & pNo ))
	End if

End Sub

'********************************************************************************
'*  [�@�\]  �Œ᎞�Ԏ擾
'********************************************************************************
Function Incf_GetSaiteiJikan(pNo)

	On Error Resume Next
	Err.Clear

	Incf_GetSaiteiJikan = False

	mM15Rs.MoveFirst

	if gf_IsNull(request("txtGseiNo" & pNo)) then
		m_iSaiteiJikan    = ""
		m_iKyuSaiteiJikan = ""
		Incf_GetSaiteiJikan = True
		Exit Function
	End if

	if gf_IsNull(m_iJunJyugyou) then
		m_iSaiteiJikan    = ""
		m_iKyuSaiteiJikan = ""
		Incf_GetSaiteiJikan = True
		Exit Function
	End if

	'// ���k���擾
	If Not f_SelKojinJyouhou(request("txtGseiNo" & pNo)) then Exit Function

	'// �Œ�̏ꍇ
	Select Case Cint(mM15Rs("M15_KEKKA_KBN"))
		Case C_K_KEKKA_NASI

			'// �O���̍��v
			w_iZenkiKei = Cint(gf_SetNull2Zero(mKojinRs("JUNJIKAN_TYUKAN_Z"))) + Cint(gf_SetNull2Zero(mKojinRs("JUNJIKAN_KIMATU_Z")))
			'// ����̍��v
			w_iKoukiKei = Cint(gf_SetNull2Zero(mKojinRs("JUNJIKAN_TYUKAN_K"))) + Cint(gf_SetNull2Zero(m_iJunJyugyou))

			'// �ݐς̏ꍇ
			if Cint(m_iRuisekiKbn) = C_K_KEKKA_RUISEKI_KEI then
				if (w_iZenkiKei + w_iKoukiKei) = 0 then
					m_iSaiteiJikan = 0
				Elseif Cint(gf_SetNull2Zero(mKojinRs("JUNJIKAN_KIMATU_Z"))) = Cint(gf_SetNull2Zero(m_iJunJyugyou)) Then
					m_iSaiteiJikan = mM15Rs("M15_BUNSHI")
				Elseif w_iZenkiKei = 0 then
					m_iSaiteiJikan = mM15Rs("M15_BUNBO")
				Else
					m_iSaiteiJikan = Cint(gf_SetNull2Zero(mM15Rs("M15_BUNSHI"))) + Cint(gf_SetNull2Zero(mM15Rs("M15_BUNBO")))
				End if

			'// �������̏ꍇ
			Else

				if (w_iZenkiKei + w_iKoukiKei) = 0 then
					m_iSaiteiJikan = 0
				Elseif w_iKoukiKei = 0 then
					m_iSaiteiJikan = mM15Rs("M15_BUNSHI")
				Elseif w_iZenkiKei = 0 Then
					m_iSaiteiJikan = mM15Rs("M15_BUNBO")
				Else
					m_iSaiteiJikan = Cint(gf_SetNull2Zero(mM15Rs("M15_BUNSHI"))) + Cint(gf_SetNull2Zero(mM15Rs("M15_BUNBO")))
				End if

			End if

	'// �ʏ�
		Case 1

			'// �ݐς̏ꍇ
			if Cint(m_iRuisekiKbn) = C_K_KEKKA_RUISEKI_KEI then
				'������̎�����
				m_iSouJituJyugyou = gf_SetNull2Zero(m_iJituJyugyou)
			Else
			'// �������̏ꍇ
				'�����ԍ��v���擾
				m_iSouJituJyugyou = Cint(gf_SetNull2Zero(mKojinRs("J_JUNJIKAN_TYUKAN_Z")))
				m_iSouJituJyugyou = Cint(m_iSouJituJyugyou) + Cint(gf_SetNull2Zero(mKojinRs("J_JUNJIKAN_KIMATU_Z")))
				m_iSouJituJyugyou = Cint(m_iSouJituJyugyou) + Cint(gf_SetNull2Zero(mKojinRs("J_JUNJIKAN_TYUKAN_K")))
				m_iSouJituJyugyou = Cint(m_iSouJituJyugyou) + Cint(gf_SetNull2Zero(m_iJituJyugyou))
			End if

			'// �Œ᎞�Ԍv�Z
			If Not f_SaiteiJikanKeisan(m_iSouJituJyugyou,m_iSaiteiJikan) Then Exit Function

			mM15Rs.MoveNext

			'// �x�w�̍Œ᎞��
			if Not mM15Rs.Eof Then
				if Cint(mM15Rs("M15_KEKKA_KBN")) = C_K_KEKKA_KYUGAKU then
					'// �Œ᎞�Ԍv�Z
					If Not f_SaiteiJikanKeisan(m_iSouJituJyugyou,m_iKyuSaiteiJikan) Then Exit Function
				End if
			End if

	End Select

	'// ���ޯ��Ӱ��
	if m_bDebugFlg Then Call Incs_ShowDebug()

	'// ���k���۰��
    Call gf_closeObject(mKojinRs)

	Incf_GetSaiteiJikan = True

End Function


'********************************************************************************
'*  [�@�\]  �Ǘ����擾
'********************************************************************************
Function Incf_SelKanriMst(pNendo,pNo)

	On Error Resume Next
	Err.Clear

	Incf_SelKanriMst = False

	'// SQL
	wSql = ""
	wSql = wSql & " SELECT * FROM M00_KANRI "
	wSql = wSql & " WHERE "
	wSql = wSql & " 	M00_KANRI.M00_NENDO = " & pNendo
	wSql = wSql & " AND M00_KANRI.M00_NO    = " & pNo

	iRet = gf_GetRecordset(wRs, wSql)
	If iRet <> 0 Then
		m_sErrMsg = Err.description
	    Call gf_closeObject(wRs)
		Exit Function
	End If

	if wRs.Eof Then
		m_sErrMsg = "�K�v�ȃf�[�^���擾�ł��Ȃ��������߁A�G���[���������܂����B"
	    Call gf_closeObject(wRs)
		Exit Function
	End If

	m_iRuisekiKbn = wRs("M00_SYUBETU")

    Call gf_closeObject(wRs)
	Incf_SelKanriMst = True

End Function


'********************************************************************************
'*  [�@�\]  ���k���擾
'********************************************************************************
Function f_SelKojinJyouhou(pGakusekiNo)

	On Error Resume Next
	Err.Clear

	f_SelKojinJyouhou = False

	if m_sUpdMode = C_UPDMODE_TUJO then
		wSql = ""
		wSql = wSql & " SELECT "
		wSql = wSql & " 	T16_JUNJIKAN_TYUKAN_Z   AS JUNJIKAN_TYUKAN_Z,"
		wSql = wSql & " 	T16_JUNJIKAN_KIMATU_Z   AS JUNJIKAN_KIMATU_Z,"
		wSql = wSql & " 	T16_JUNJIKAN_TYUKAN_K   AS JUNJIKAN_TYUKAN_K,"
		wSql = wSql & " 	T16_J_JUNJIKAN_TYUKAN_Z AS J_JUNJIKAN_TYUKAN_Z,"
		wSql = wSql & " 	T16_J_JUNJIKAN_KIMATU_Z AS J_JUNJIKAN_KIMATU_Z,"
		wSql = wSql & " 	T16_J_JUNJIKAN_TYUKAN_K AS J_JUNJIKAN_TYUKAN_K "
		wSql = wSql & " FROM T16_RISYU_KOJIN "
		wSql = wSql & " WHERE "
		wSql = wSql & "     T16_RISYU_KOJIN.T16_NENDO      =  " & m_iNendo
		wSql = wSql & " AND T16_RISYU_KOJIN.T16_GAKUSEI_NO = '" & pGakusekiNo & "' "
		wSql = wSql & " AND T16_RISYU_KOJIN.T16_KAMOKU_CD  = '" & m_sKamokuCd & "' "
	Else
		wSql = ""
		wSql = wSql & " SELECT "
		wSql = wSql & " 	T34_JUNJIKAN_TYUKAN_Z   AS JUNJIKAN_TYUKAN_Z,"
		wSql = wSql & " 	T34_JUNJIKAN_KIMATU_Z   AS JUNJIKAN_KIMATU_Z,"
		wSql = wSql & " 	T34_JUNJIKAN_TYUKAN_K   AS JUNJIKAN_TYUKAN_K,"
		wSql = wSql & " 	T34_J_JUNJIKAN_TYUKAN_Z AS J_JUNJIKAN_TYUKAN_Z,"
		wSql = wSql & " 	T34_J_JUNJIKAN_KIMATU_Z AS J_JUNJIKAN_KIMATU_Z,"
		wSql = wSql & " 	T34_J_JUNJIKAN_TYUKAN_K AS J_JUNJIKAN_TYUKAN_K "
		wSql = wSql & " FROM T34_RISYU_TOKU "
		wSql = wSql & " WHERE "
		wSql = wSql & "     T34_RISYU_TOKU.T34_NENDO      =  " & m_iNendo
		wSql = wSql & " AND T34_RISYU_TOKU.T34_GAKUSEI_NO = '" & pGakusekiNo & "' "
		wSql = wSql & " AND T34_RISYU_TOKU.T34_TOKUKATU_CD  = '" & m_sKamokuCd & "' "
	End if

	iRet = gf_GetRecordset(mKojinRs, wSql)
	If iRet <> 0 Then
		m_sErrMsg = Err.description
	    Call gf_closeObject(mKojinRs)
		Exit Function
	End If

	if mKojinRs.Eof Then
		m_sErrMsg = "�K�v�ȃf�[�^���擾�ł��Ȃ��������߁A�G���[���������܂����B"
	    Call gf_closeObject(mKojinRs)
		Exit Function
	End If

	f_SelKojinJyouhou = True

End Function

'********************************************************************************
'*  [�@�\]  ���ہE���Ȑݒ�擾
'********************************************************************************
Function Incf_SelM15_KEKKA_KESSEKI()

	On Error Resume Next
	Err.Clear

	Incf_SelM15_KEKKA_KESSEKI = False

	'// SQL
	wSql = ""
	wSql = wSql & " SELECT * FROM M15_KEKKA_KESSEKI "
	wSql = wSql & " WHERE "
	wSql = wSql & " 	M15_KEKKA_KESSEKI.M15_NENDO    = " & m_iNendo
	wSql = wSql & " AND M15_KEKKA_KESSEKI.M15_KEKKA_CD = 1 "				'�H�H�H�v�f�o�b�O
	wSql = wSql & " AND M15_KEKKA_KESSEKI.M15_TANI     = " & C_K_KEKKA_TANI_NASI
	wSql = wSql & " ORDER BY M15_KEKKA_KBN "

	iRet = gf_GetRecordset(mM15Rs, wSql)
	If iRet <> 0 Then
		m_sErrMsg = Err.description
	    Call gf_closeObject(mM15Rs)
		Exit Function
	End If

	if mM15Rs.Eof Then
		m_sErrMsg = "�K�v�ȃf�[�^���擾�ł��Ȃ��������߁A�G���[���������܂����B"
	    Call gf_closeObject(mM15Rs)
		Exit Function
	End If

	Incf_SelM15_KEKKA_KESSEKI = True

End Function


'********************************************************************************
'*  [�@�\]  �Œ᎞�Ԍv�Z
'*  [����]  pSouJituJyugyou = "�������Ǝ���"
'********************************************************************************
Function f_SaiteiJikanKeisan(pSouJituJyugyou,pSaiteiJikan)

	On Error Resume Next
	Err.Clear

	f_SaiteiJikanKeisan = False

	'// �ϐ�������
	pSaiteiJikan = ""

	if pSouJituJyugyou = 0 then
		pSaiteiJikan = 0
		f_SaiteiJikanKeisan = True
		Exit Function
	End if

	'// �؎́E�؏�E�l�̌ܓ�
	Select Case Cint(mM15Rs("M15_HASUU_KBN"))
		Case C_HASU_SYORI_KIRISUTE   : wPlus = 0
		Case C_HASU_SYORI_KIRIAGE    : wPlus = 0.9
		Case C_HASU_SYORI_SISYAGONYU : wPlus = 0.5
	End Select

	pSaiteiJikan = (pSouJituJyugyou * (Cint(mM15Rs("M15_BUNSHI")) / Cint(mM15Rs("M15_BUNBO")))) + wPlus
	pSaiteiJikan = Int(pSaiteiJikan)

	'// � = �v�Z�l���܂ޏꍇ�́A���莞�ɒ����Ȃ��悤��"-1"����
	if Cint(mM15Rs("M15_KIJYUN_KBN")) = C_SUUCHI_KIJYUN_KBN_INC Then
		pSaiteiJikan = pSaiteiJikan - 1
	End if

	'// ���ޯ��Ӱ��
	If m_bDebugFlg Then
		response.write (pSouJituJyugyou * (Cint(mM15Rs("M15_BUNSHI")) / Cint(mM15Rs("M15_BUNBO")))) & "<BR>"
		response.write "(" & pSouJituJyugyou & " * (" & mM15Rs("M15_BUNSHI") & "/" & mM15Rs("M15_BUNBO") & ")) + " & wPlus & " = int(" & pSaiteiJikan & ")<br>"
	End if

	f_SaiteiJikanKeisan = True

End Function

%>
