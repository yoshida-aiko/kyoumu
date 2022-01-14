<%

'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0800/sei0800_upd_func.asp
' 機      能: 下ページ 成績登録の登録、更新
'-------------------------------------------------------------------------
' 説      明: 1.総授業時間と純授業時間から実授業時間を算出
'             2.最低時間の計算
'-------------------------------------------------------------------------
' 作      成: 2002/03/27 モチナガ
' 変      更: 
' デバッグ  : 減算区分取得でWHERE条件のコンストがわからない
'*************************************************************************/

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
Dim m_iSouJyugyou			'総時間
Dim m_iJunJyugyou			'純時間
Dim m_iJituJyugyou			'実時間
Dim m_iSouJituJyugyou		'総実時間
Dim m_iGenzan				'減算区分(0:しない,1:する)
Dim m_iRuisekiKbn			'累積区分(0:試験毎,1:累積)
Dim m_iSaiteiJikan			'最低時間
Dim m_iKyuSaiteiJikan		'休学最低時間
Dim mKojinRs				'生徒情報ﾚｺｰﾄﾞｾｯﾄ
Dim mM15Rs					'欠課・欠席設定ﾚｺｰﾄﾞｾｯﾄ
Dim m_sUpdMode				'ｱｯﾌﾟﾃﾞｰﾄﾓｰﾄﾞ（TUJO:通常, TOKU:特別）

const C_UPDMODE_TUJO = "TUJO"
const C_UPDMODE_TOKU = "TOKU"

Dim m_bDebugFlg				'ﾛｰｶﾙﾃﾞﾊﾞｯｸﾞﾌﾗｸﾞ

	m_bDebugFlg = false

	'// ﾊﾟﾗﾒｰﾀ取得
	m_iSouJyugyou = request("hidSouJyugyou")		'// 総時間
	m_iJunJyugyou = request("hidJunJyugyou")		'// 純時間
	m_sUpdMode    = request("hidUpdMode")			'// ｱｯﾌﾟﾃﾞｰﾄﾓｰﾄﾞ

'********************************************************************************
'*  [機能]  デバッグ表示
'********************************************************************************
Sub Incs_ShowDebug()

	On Error Resume Next
	Err.Clear

	mM15Rs.MoveFirst

	response.write "<BR>*********** ﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞ **************<br>"
	response.write "総時間 = " & m_iSouJyugyou & "<BR>"
	response.write "純時間 = " & m_iJunJyugyou & "<BR>"
	response.write "減算区分 = " & m_iGenzan & "<BR>"
	response.write "実時間 = " & m_iJituJyugyou & "<BR>"

	response.write "<BR>--- 最低時間計算 ----<br>"
	response.write "M15 区分 = " & mM15Rs("M15_KEKKA_KBN") & "<BR>"
	response.write "累積区分 = " & m_iRuisekiKbn & "<BR>"
	response.write "<BR>"

	Select Case Cint(mM15Rs("M15_KEKKA_KBN"))
		Case 0
			response.write "// 最低時固定<br>"
			response.write "最低時間 = " & m_iSaiteiJikan & "<BR>"

		Case 1
			response.write "<BR>// 最低時通常<br>"
			response.write "総実時間 = " & m_iSouJituJyugyou & "<BR>"
			response.write "分子 = " & mM15Rs("M15_BUNSHI") & "<BR>"
			response.write "分母 = " & mM15Rs("M15_BUNBO") & "<BR>"
			response.write "端数区分 = " & mM15Rs("M15_HASUU_KBN") & "<br>"
			response.write "最低時間 = " & m_iSaiteiJikan & "<BR>"

			mM15Rs.MoveNext
			if Not mM15Rs.Eof Then
				response.write "<BR>// 最低時休学<br>"
				response.write "M15 区分 = " & mM15Rs("M15_KEKKA_KBN") & "<BR>"
				response.write "累積区分 = " & m_iRuisekiKbn & "<BR>"
				response.write "<BR>"
				response.write "総実時間 = " & m_iSouJituJyugyou & "<BR>"
				response.write "分子 = " & mM15Rs("M15_BUNSHI") & "<BR>"
				response.write "分母 = " & mM15Rs("M15_BUNBO") & "<BR>"
				response.write "端数区分 = " & mM15Rs("M15_HASUU_KBN") & "<br>"
				response.write "最低時間 = " & m_iKyuSaiteiJikan & "<BR>"
			End if

	End Select

End Sub

'********************************************************************************
'*  [機能]  減算区分取得
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
	wSql = wSql & " AND M15_KEKKA_KESSEKI.M15_KEKKA_CD  = 4 "				'？？？要デバッグ
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
'*  [機能]  実授業時間取得
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
'*  [機能]  最低時間取得
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

	'// 生徒情報取得
	If Not f_SelKojinJyouhou(request("txtGseiNo" & pNo)) then Exit Function

	'// 固定の場合
	Select Case Cint(mM15Rs("M15_KEKKA_KBN"))
		Case C_K_KEKKA_NASI

			'// 前期の合計
			w_iZenkiKei = Cint(gf_SetNull2Zero(mKojinRs("JUNJIKAN_TYUKAN_Z"))) + Cint(gf_SetNull2Zero(mKojinRs("JUNJIKAN_KIMATU_Z")))
			'// 後期の合計
			w_iKoukiKei = Cint(gf_SetNull2Zero(mKojinRs("JUNJIKAN_TYUKAN_K"))) + Cint(gf_SetNull2Zero(m_iJunJyugyou))

			'// 累積の場合
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

			'// 試験毎の場合
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

	'// 通常
		Case 1

			'// 累積の場合
			if Cint(m_iRuisekiKbn) = C_K_KEKKA_RUISEKI_KEI then
				'後期末の実時間
				m_iSouJituJyugyou = gf_SetNull2Zero(m_iJituJyugyou)
			Else
			'// 試験毎の場合
				'実時間合計を取得
				m_iSouJituJyugyou = Cint(gf_SetNull2Zero(mKojinRs("J_JUNJIKAN_TYUKAN_Z")))
				m_iSouJituJyugyou = Cint(m_iSouJituJyugyou) + Cint(gf_SetNull2Zero(mKojinRs("J_JUNJIKAN_KIMATU_Z")))
				m_iSouJituJyugyou = Cint(m_iSouJituJyugyou) + Cint(gf_SetNull2Zero(mKojinRs("J_JUNJIKAN_TYUKAN_K")))
				m_iSouJituJyugyou = Cint(m_iSouJituJyugyou) + Cint(gf_SetNull2Zero(m_iJituJyugyou))
			End if

			'// 最低時間計算
			If Not f_SaiteiJikanKeisan(m_iSouJituJyugyou,m_iSaiteiJikan) Then Exit Function

			mM15Rs.MoveNext

			'// 休学の最低時間
			if Not mM15Rs.Eof Then
				if Cint(mM15Rs("M15_KEKKA_KBN")) = C_K_KEKKA_KYUGAKU then
					'// 最低時間計算
					If Not f_SaiteiJikanKeisan(m_iSouJituJyugyou,m_iKyuSaiteiJikan) Then Exit Function
				End if
			End if

	End Select

	'// ﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞ
	if m_bDebugFlg Then Call Incs_ShowDebug()

	'// 生徒情報ｸﾛｰｽﾞ
    Call gf_closeObject(mKojinRs)

	Incf_GetSaiteiJikan = True

End Function


'********************************************************************************
'*  [機能]  管理情報取得
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
		m_sErrMsg = "必要なデータが取得できなかったため、エラーが発生しました。"
	    Call gf_closeObject(wRs)
		Exit Function
	End If

	m_iRuisekiKbn = wRs("M00_SYUBETU")

    Call gf_closeObject(wRs)
	Incf_SelKanriMst = True

End Function


'********************************************************************************
'*  [機能]  生徒情報取得
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
		m_sErrMsg = "必要なデータが取得できなかったため、エラーが発生しました。"
	    Call gf_closeObject(mKojinRs)
		Exit Function
	End If

	f_SelKojinJyouhou = True

End Function

'********************************************************************************
'*  [機能]  欠課・欠席設定取得
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
	wSql = wSql & " AND M15_KEKKA_KESSEKI.M15_KEKKA_CD = 1 "				'？？？要デバッグ
	wSql = wSql & " AND M15_KEKKA_KESSEKI.M15_TANI     = " & C_K_KEKKA_TANI_NASI
	wSql = wSql & " ORDER BY M15_KEKKA_KBN "

	iRet = gf_GetRecordset(mM15Rs, wSql)
	If iRet <> 0 Then
		m_sErrMsg = Err.description
	    Call gf_closeObject(mM15Rs)
		Exit Function
	End If

	if mM15Rs.Eof Then
		m_sErrMsg = "必要なデータが取得できなかったため、エラーが発生しました。"
	    Call gf_closeObject(mM15Rs)
		Exit Function
	End If

	Incf_SelM15_KEKKA_KESSEKI = True

End Function


'********************************************************************************
'*  [機能]  最低時間計算
'*  [引数]  pSouJituJyugyou = "総実授業時間"
'********************************************************************************
Function f_SaiteiJikanKeisan(pSouJituJyugyou,pSaiteiJikan)

	On Error Resume Next
	Err.Clear

	f_SaiteiJikanKeisan = False

	'// 変数初期化
	pSaiteiJikan = ""

	if pSouJituJyugyou = 0 then
		pSaiteiJikan = 0
		f_SaiteiJikanKeisan = True
		Exit Function
	End if

	'// 切捨・切上・四捨五入
	Select Case Cint(mM15Rs("M15_HASUU_KBN"))
		Case C_HASU_SYORI_KIRISUTE   : wPlus = 0
		Case C_HASU_SYORI_KIRIAGE    : wPlus = 0.9
		Case C_HASU_SYORI_SISYAGONYU : wPlus = 0.5
	End Select

	pSaiteiJikan = (pSouJituJyugyou * (Cint(mM15Rs("M15_BUNSHI")) / Cint(mM15Rs("M15_BUNBO")))) + wPlus
	pSaiteiJikan = Int(pSaiteiJikan)

	'// 基準 = 計算値を含む場合は、判定時に超えないように"-1"する
	if Cint(mM15Rs("M15_KIJYUN_KBN")) = C_SUUCHI_KIJYUN_KBN_INC Then
		pSaiteiJikan = pSaiteiJikan - 1
	End if

	'// ﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞ
	If m_bDebugFlg Then
		response.write (pSouJituJyugyou * (Cint(mM15Rs("M15_BUNSHI")) / Cint(mM15Rs("M15_BUNBO")))) & "<BR>"
		response.write "(" & pSouJituJyugyou & " * (" & mM15Rs("M15_BUNSHI") & "/" & mM15Rs("M15_BUNBO") & ")) + " & wPlus & " = int(" & pSaiteiJikan & ")<br>"
	End if

	f_SaiteiJikanKeisan = True

End Function

%>
