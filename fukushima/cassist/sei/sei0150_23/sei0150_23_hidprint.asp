<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0150/sei0150_23_print.asp
' 機      能: 印刷処理を行なう
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 説      明:
'	(パターン)
'	・通常授業、特別活動
'	・科目区分(0:一般科目,1:専門科目)
'	・必修選択区分(1:必修,2:選択)
'	・レベル別区分(0:一般科目,1:レベル別科目)を調べる
'-------------------------------------------------------------------------
' 作      成: 2003/05/08 hirota
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Dim m_bErrFlg				'//ｴﾗｰﾌﾗｸﾞ

    Const C_ERR_GETDATA = "データの取得に失敗しました"

    '氏名選択用のWhere条件
    Dim m_iNendo				'//年度
    Dim m_sKyokanCd				'//教官コード
    Dim m_sSikenKBN				'//試験区分
    Dim m_iGakunen				'//学年m_sGakuNo
    Dim m_sClassNo				'//学科
    Dim m_sKamokuCd				'//科目コード
    Dim m_sSikenNm				'//試験名
    Dim m_sGakkaCd
    Dim m_iKamoku_Kbn
    Dim m_iHissen_Kbn
	Dim m_ilevelFlg
	Dim m_Rs
    Dim m_rCnt					'//レコードカウント
	Dim m_SRs
	Dim m_bSeiInpFlg			'//入力期間フラグ
	Dim m_bKekkaNyuryokuFlg		'//欠課入力可能ﾌﾗｸﾞ(True:入力可 / False:入力不可)
	Dim m_iShikenInsertType
	Dim m_sSyubetu
	Dim m_iKamokuKbn			'//科目区分( 0:通常授業、 1:特別科目)
	Dim m_sKamokuBunrui			'//科目分類(01:通常授業、02:認定科目、03:特別科目)
	Dim m_iSeisekiInpType
	Dim m_Date
	Dim m_bZenkiOnly
	Dim m_bNiteiFlg
	Dim m_sGakkoNO				'学校番号
	Dim m_sUpdDate

'///////////////////////////メイン処理/////////////////////////////
	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()
	Dim w_iRet
	Dim w_sSQL
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
	'Message用の変数の初期化
	w_sWinTitle = "キャンパスアシスト"
	w_sMsgTitle = "成績登録"
	w_sMsg = ""
	w_sRetURL = C_RetURL & C_ERR_RETURL
	w_sTarget = ""
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = false

	Do
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()

		'学校番号の取得
		if Not gf_GetGakkoNO(m_sGakkoNO) then Exit Do

		'//成績入力方法の取得(0:点数[C_SEISEKI_INP_TYPE_NUM]、1:文字[C_SEISEKI_INP_TYPE_STRING]、2:欠課、遅刻[C_SEISEKI_INP_TYPE_KEKKA])
		if not gf_GetKamokuSeisekiInp(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iSeisekiInpType) then 
			m_bErrFlg = True
			Exit Do
		end if

		'//前期のみ開設か通年か調べる
		if not f_SikenInfo(m_bZenkiOnly) then
			m_bErrFlg = True
			Exit Do
		end if

		'//認定前後情報取得
		if not gf_GetGakunenNintei(m_iNendo,cint(m_iGakunen),m_bNiteiFlg) then
			m_bErrFlg = True
			Exit Do
		end if

		If m_iKamokuKbn = C_JIK_JUGYO then  '通常授業の場合
			'//科目情報を取得
			'//科目区分(0:一般科目,1:専門科目)、及び、必修選択区分(1:必修,2:選択)を調べる
			'//レベル別区分(0:一般科目,1:レベル別科目)を調べる
			If not f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) Then m_bErrFlg = True : Exit Do
		end if

		'//成績、学生データ取得
		If not f_GetStudent() Then m_bErrFlg = True : Exit Do

		If m_Rs.EOF Then
			Call gs_showWhitePage("個人履修データが存在しません。","成績登録")
			Exit Do
		End If

		'// ページを表示
		Call showPage()
		Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		if w_sMsg = "" then w_sMsg = C_ERR_GETDATA
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If

	'// 終了処理
	Call gf_closeObject(m_Rs)
	Call gf_closeObject(m_SRs)
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()
	
	m_iNendo	 = request("txtNendo")
	m_sKyokanCd	 = request("txtKyokanCd")
	m_sSikenKBN	 = cint(request("sltShikenKbn"))
	m_iGakunen	 = Cint(request("txtGakuNo"))
	m_sClassNo	 = cint(request("txtClassNo"))
	m_sKamokuCd	 = request("txtKamokuCd")
	m_sGakkaCd	 = request("txtGakkaCd")
	m_sSyubetu	 = trim(Request("hidSyubetu"))
	m_iShikenInsertType = 0

	m_iKamokuKbn = cint(Request("hidKamokuKbn"))

	if m_iKamokuKbn = C_JIK_JUGYO then
		'通常科目
		m_sKamokuBunrui = C_KAMOKUBUNRUI_TUJYO
	else
		'特別科目
		m_sKamokuBunrui = C_KAMOKUBUNRUI_TOKUBETU
	end if
	
	m_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")
	
End Sub

'********************************************************************************
'*  [機能]  前期開設かどうか調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_SikenInfo(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs
    
    On Error Resume Next
    Err.Clear
    
    f_SikenInfo = false
	p_bZenkiOnly = false
	
	'//試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T15_KAMOKU_CD "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T15_RISYU "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	T15_NYUNENDO = " & Cint(m_iNendo)-cint(m_iGakunen)+1
	w_sSQL = w_sSQL & " AND T15_GAKKA_CD = '" & m_sGakkaCd & "'"
	w_sSQL = w_sSQL & " AND T15_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'" 
	w_sSQL = w_sSQL & " AND T15_KAISETU" & m_iGakunen & "=" & C_KAI_ZENKI
	
	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function
	
	'Response.Write "0"
	
	'//戻り値ｾｯﾄ
	If w_Rs.EOF = False Then
		p_bZenkiOnly = True
	End If
	
	f_SikenInfo = true
	
	Call gf_closeObject(w_Rs)
	
End Function

'********************************************************************************
'*  [機能]  コンボで選択された科目の科目区分及び、必修選択区分を調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKamokuInfo(p_iKamoku_Kbn,p_iHissen_Kbn,p_ilevelFlg)
	Dim w_sSQL
	Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_GetKamokuInfo = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T15_RISYU.T15_KAMOKU_KBN"
	w_sSQL = w_sSQL & " 	,T15_RISYU.T15_HISSEN_KBN"
	w_sSQL = w_sSQL & " 	,T15_RISYU.T15_LEVEL_FLG"
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T15_RISYU"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(m_iGakunen) + 1
	w_sSQL = w_sSQL & " AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
	w_sSQL = w_sSQL & " AND T15_RISYU.T15_KAMOKU_CD='" & m_sKamokuCd & "' "

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

	'//戻り値ｾｯﾄ
	If w_Rs.EOF = False Then
		p_iKamoku_Kbn = w_Rs("T15_KAMOKU_KBN")
		p_iHissen_Kbn = w_Rs("T15_HISSEN_KBN")
		p_ilevelFlg = w_Rs("T15_LEVEL_FLG")
	End If

	f_GetKamokuInfo = true

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[機能]	データの取得
'********************************************************************************
Function f_GetStudent()
	
	Dim w_sSQL
	Dim w_FieldName
	Dim w_Table
	Dim w_TableName
	Dim w_KamokuName

	On Error Resume Next
	Err.Clear
	
	f_GetStudent = false

	'科目区分
	if m_iKamokuKbn = C_JIK_JUGYO then  '通常授業の場合
		w_Table      = "T16"
		w_TableName  = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	else								'特活などの場合
		w_Table      = "T34"
		w_TableName  = "T34_RISYU_TOKU"
		w_KamokuName = "T34_TOKUKATU_CD"
	end if

	'//文字、数値入力により、取ってくるフィールドを変える
	if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
		if m_bNiteiFlg and m_iKamokuKbn = C_JIK_JUGYO then
			w_FieldName = "HTEN"
		else
			w_FieldName = "SEI"
		end if
	else
		w_FieldName = "HYOKA"
	end if

	'//検索結果の値より一覧を表示
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI1, "
	w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI2, "
	w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI3, "
	w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI4, "
	w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z       AS KEKA_ZT, "			'欠課（前期中間）
	w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z       AS KEKA_ZK, "			'欠課（前期期末）
	w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K       AS KEKA_KT, "			'欠課（後期中間）
	w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_K       AS KEKA_KK, "			'欠課（後期期末）
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z  AS TEISI_ZT,"			'停止（前期中間）
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z  AS TEISI_ZK,"			'停止（前期期末）
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K  AS TEISI_KT,"			'停止（期末中間）
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_K  AS TEISI_KK,"			'停止（後期期末）
	w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_Z       AS KIBI_ZT, "			'忌引（前期中間）
	w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_Z       AS KIBI_ZK, "			'忌引（前期期末）
	w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_K       AS KIBI_KT, "			'忌引（後期期末）
	w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_K       AS KIBI_KK, "			'忌引（後期期末）
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_Z   AS HAKEN_ZT,"			'派遣（前期中間）
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_Z   AS HAKEN_ZK,"			'派遣（前期期末）
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_K   AS HAKEN_KT,"			'派遣（後期中間）
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_K   AS HAKEN_KK,"			'派遣（後期期末）
	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z    AS SOUJI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z    AS SOUJI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K    AS SOUJI_KT,"
	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K    AS SOUJI_KK,"
	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z   AS JUNJI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z   AS JUNJI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K   AS JUNJI_KT,"
	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K   AS JUNJI_KK,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_TYUKAN_Z AS J_JUNJI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_KIMATU_Z AS J_JUNJI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_TYUKAN_K AS J_JUNJI_KT,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_KIMATU_K AS J_JUNJI_KK,"
	w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_Z      AS HYOKA_ZT,  "
	w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_Z      AS HYOKA_ZK,  "
	w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_K      AS HYOKA_KT,  "
	w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_K      AS HYOKA_KK,  "
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_Z   AS KOUSINBI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_Z   AS KOUSINBI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_K   AS KOUSINBI_KT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_K   AS KOUSINBI_KK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_TYUKAN_Z AS KOUSINTIME_ZT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_KIMATU_Z AS KOUSINTIME_ZK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_TYUKAN_K AS KOUSINTIME_KT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_KIMATU_K AS KOUSINTIME_KK,"
	w_sSQL = w_sSQL & w_Table & "_HYOKA_FUKA_KBN      AS HYOKA_FUKA, "
	w_sSQL = w_sSQL & w_Table & "_HAITOTANI           AS HAITOTANI, "

	if m_iKamokuKbn = C_JIK_JUGYO then
		w_sSQL = w_sSQL & " 	T16_SELECT_FLG, "
		w_sSQL = w_sSQL & " 	T16_LEVEL_KYOUKAN, "
		w_sSQL = w_sSQL & " 	T16_OKIKAE_FLG, "
	end if

	w_sSQL = w_sSQL & " 	T13_GAKUSEI_NO  AS GAKUSEI_NO, "
	w_sSQL = w_sSQL & " 	T13_GAKUSEKI_NO AS GAKUSEKI_NO,"
	w_sSQL = w_sSQL & " 	T11_SIMEI       AS SIMEI       "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & 		w_TableName & ","
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI,   "
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN    "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & 				w_Table & "_NENDO      = " & Cint(m_iNendo)
	w_sSQL = w_sSQL & " 	AND	" & w_Table & "_GAKUSEI_NO = T11_GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND	" & w_Table & "_GAKUSEI_NO = T13_GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND	T13_GAKUNEN = " & cint(m_iGakunen)
	w_sSQL = w_sSQL & " 	AND	T13_CLASS   = " & cint(m_sClassNo)
	w_sSQL = w_sSQL & " 	AND	" & w_KamokuName & "  = '" & m_sKamokuCd & "' "
	w_sSQL = w_sSQL & " 	AND	" & w_Table & "_NENDO = T13_NENDO "

	if m_iKamokuKbn = C_JIK_JUGYO then
		'//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
		w_sSQL = w_sSQL & " AND	T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
	end if

	w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "

	'レコード取得
	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit function

	'表示する更新日付 & 時間
	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_ZT"))) & "　" & gf_SetNull2String(m_Rs("KOUSINTIME_ZT"))
		Case C_SIKEN_ZEN_KIM : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_ZK"))) & "　" & gf_SetNull2String(m_Rs("KOUSINTIME_ZK"))
		Case C_SIKEN_KOU_TYU : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_KT"))) & "　" & gf_SetNull2String(m_Rs("KOUSINTIME_KT"))
		Case C_SIKEN_KOU_KIM : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_KK"))) & "　" & gf_SetNull2String(m_Rs("KOUSINTIME_KK"))
	End Select

	'//ﾚｺｰﾄﾞカウント取得
	m_rCnt = gf_GetRsCount(m_Rs)

	f_GetStudent = true

End Function

'********************************************************************************
'*	[機能]	データの取得
'********************************************************************************
Function f_Syukketu2New(p_gaku,p_kbn)
	Dim w_GAKUSEI_NO
	Dim w_SYUKKETU_KBN
	
	f_Syukketu2New = ""
	w_GAKUSEI_NO = ""
	w_SYUKKETU_KBN = ""
	w_SKAISU = ""
	
	If m_SRs.EOF Then
		Exit Function
	Else
		Do Until m_SRs.EOF
			w_GAKUSEI_NO = m_SRs("T21_GAKUSEKI_NO")
			w_SYUKKETU_KBN = m_SRs("T21_SYUKKETU_KBN")
			w_SKAISU = gf_SetNull2String(m_SRs("KAISU"))
			
			If Cstr(w_GAKUSEI_NO) = Cstr(p_gaku) AND cstr(w_SYUKKETU_KBN) = cstr(p_kbn) Then
				f_Syukketu2New = w_SKAISU
				Exit Do
			End If
			
			m_SRs.MoveNext
		Loop
		
		m_SRs.MoveFirst
	End If
	
End Function

'********************************************************************************
'*  [機能] 異動チェック
'********************************************************************************
Sub s_IdouCheck(p_GakusekiNo,p_IdouKbn,p_IdouName,p_bNoChangeZK,p_bNoChangeKT,p_bNoChangeKK)
	Dim w_IdoutypeName	'異動状況名
	
	w_IdoutypeName = ""
	p_IdouName = ""
	
	p_IdouKbn = gf_Get_IdouChk(p_GakusekiNo,m_Date,m_iNendo,w_IdoutypeName)
	
	if Cstr(p_IdouKbn) <> "" and Cstr(p_IdouKbn) <> CStr(C_IDO_FUKUGAKU) AND _
		Cstr(p_IdouKbn) <> Cstr(C_IDO_TEI_KAIJO) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKO) AND _
		Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKA) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_KOKUHI) Then
					
		p_IdouName = "[" & w_IdoutypeName & "]"
		p_bNoChangeZK = True
		p_bNoChangeKT = True
		p_bNoChangeKK = True
	end if
	
end Sub

'********************************************************************************
'*  [機能] 成績のセット
'********************************************************************************
Sub s_SetGrades(p_sSeiseki_ZK,  p_sSeiseki_KT,  p_sSeiseki_KK, _
				p_sHyoka_ZK,    p_sHyoka_KT,    p_sHyoka_KK, _
				p_bNoChange_ZK, p_bNoChange_KT, p_bNoChange_KK)

	Dim w_UpdDateZK
    Dim w_UpdDateKK

	'/試験区分によって取ってくる、フィールドを変える。
	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sSeiseki_ZK = gf_SetNull2String(m_Rs("SEI1"))
			p_sHyoka_ZK   = gf_SetNull2String(m_Rs("HYOKA_ZT"))
		Case Else
			p_sSeiseki_ZK = gf_SetNull2String(m_Rs("SEI2"))
			p_sHyoka_ZK   = gf_SetNull2String(m_Rs("HYOKA_ZK"))
	End Select

	p_sSeiseki_KT = gf_SetNull2String(m_Rs("SEI3"))
	p_sSeiseki_KK = gf_SetNull2String(m_Rs("SEI4"))
	p_sHyoka_KT   = gf_SetNull2String(m_Rs("HYOKA_KT"))
	p_sHyoka_KK   = gf_SetNull2String(m_Rs("HYOKA_KK"))

	'学年末試験の場合のみ
'	If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
'		w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
'		w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))
'		if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
'			p_sSeiseki_KK = gf_SetNull2String(m_Rs("SEI2"))
'			p_sHyoka_KK   = gf_SetNull2String(m_Rs("HYOKA_ZK"))
'		End If
'	End If

	'//通常授業のとき
	if m_iKamokuKbn = C_JIK_JUGYO then

		p_bNoChange_ZK = False
		p_bNoChange_KT = False
		p_bNoChange_KK = False

		'//科目が選択科目の場合は、生徒が選択しているかどうかを判別する。選択しいない生徒は入力不可とする。
		if cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then 
			if cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then
				p_bNoChange_ZK = true
				p_bNoChange_KT = true
				p_bNoChange_KK = true
			end if
		else
			if Cstr(m_iLevelFlg) = "1" then
				if isNull(m_Rs("T16_LEVEL_KYOUKAN")) = true then
					p_bNoChange_ZK = true
					p_bNoChange_KT = true
					p_bNoChange_KK = true
				else
					if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
						p_bNoChange_ZK = true
						p_bNoChange_KT = true
						p_bNoChange_KK = true
					End if
				End if
			End if
		end if
	end if

end Sub

'********************************************************************************
'*  [機能]  欠課、遅刻数のセット
'********************************************************************************
Sub s_SetKekka(p_sKekka_ZK, p_sKekka_KT, p_sKekka_KK, _
			   p_sKibi_ZK , p_sKibi_KT , p_sKibi_KK , _
			   p_sTeisi_ZK, p_sTeisi_KT, p_sTeisi_KK, _
			   p_sHaken_ZK, p_sHaken_KT, p_sHaken_KK)

	Dim w_UpdDateZK
	Dim w_UpdDateKK

	'/試験区分によって取ってくる、フィールドを変える。
	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sKekka_ZK  = gf_SetNull2String(m_Rs("KEKA_ZT"))
			p_sKibi_ZK   = gf_SetNull2String(m_Rs("KIBI_ZT"))
			p_sTeisi_ZK  = gf_SetNull2String(m_Rs("TEISI_ZT"))
			p_sHaken_ZK  = gf_SetNull2String(m_Rs("HAKEN_ZT"))
		Case Else
			p_sKekka_ZK  = gf_SetNull2String(m_Rs("KEKA_ZK"))
			p_sKibi_ZK   = gf_SetNull2String(m_Rs("KIBI_ZK"))
			p_sTeisi_ZK  = gf_SetNull2String(m_Rs("TEISI_ZK"))
			p_sHaken_ZK  = gf_SetNull2String(m_Rs("HAKEN_ZK"))
	End Select

	p_sKekka_KT  = gf_SetNull2String(m_Rs("KEKA_KT"))
	p_sKibi_KT   = gf_SetNull2String(m_Rs("KIBI_KT"))
	p_sTeisi_KT  = gf_SetNull2String(m_Rs("TEISI_KT"))
	p_sHaken_KT  = gf_SetNull2String(m_Rs("HAKEN_KT"))
	p_sKekka_KK  = gf_SetNull2String(m_Rs("KEKA_KK"))
	p_sKibi_KK   = gf_SetNull2String(m_Rs("KIBI_KK"))
	p_sTeisi_KK  = gf_SetNull2String(m_Rs("TEISI_KK"))
	p_sHaken_KK  = gf_SetNull2String(m_Rs("HAKEN_KK"))

End Sub

'********************************************************************************
'*	[機能]	試験名取得
'********************************************************************************
Function f_ShikenMei()
	Dim w_Rs
	
	On Error Resume Next
	Err.Clear
	
	f_ShikenMei = ""
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD = " & cint(m_sSikenKBN)
	w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
	w_sSQL = w_sSQL & " AND M01_NENDO = " & cint(m_iNendo)
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function
	
	If not w_Rs.EOF Then
		f_ShikenMei = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))
	End If

    call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  学校名を取得
'********************************************************************************
Function f_GetSchoolName()

	Dim w_Rs
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetSchoolName = ""

    '// 学校名取得
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "     M19_NAME "
    w_sSQL = w_sSQL & "FROM M19_GAKKO "

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

    '// 学校名
    f_GetSchoolName = w_Rs("M19_NAME")

    call gf_closeObject(w_Rs)

End Function

'****************************************************
'[機能]	和暦フォーマット	:MM月DD日（曜日）
'[引数]	pDate : 対象日付(YYYY/MM/DD)
'[戻値]	
'****************************************************
Function f_fmtWareki(pDate)

	f_fmtWareki = ""

	'// Nullなら抜ける
	if gf_IsNull(trim(pDate)) then	Exit Function

	'// MM月DD日作成
	w_YY = Left(FormatYYYYMMDD(pDate),4) & "年"
	w_MM = Mid(FormatYYYYMMDD(pDate),6,2) & "月"
	w_DD = Right(FormatYYYYMMDD(pDate),2) & "日"

	'// 曜日を取得
	w_Youbi = WeekdayName(Weekday(FormatYYYYMMDD(pDate))) & "<BR>"
	w_Youbi = "（" & Left(w_Youbi,1) & "）"

	f_fmtWareki = w_YY & w_MM & w_DD

End Function

'***********************************************************
' 機　　能：西暦年度から和暦年度を求める
' 戻　　値：変換結果
'           (成功):和暦、(失敗):""
' 引　　数：p_sNendo - 西暦の年度
' 詳細機能：西暦年度から和暦年度を求める
' 備　　考：和暦年度を返す。元号はつかない。
'***********************************************************
Function f_Nendo2Wareki(p_iNendo)
    Dim w_sSql
    Dim w_Rs

	On Error Resume Next
	Err.Clear

    '== 初期化 ==
    f_Nendo2Wareki = ""

    '== 和暦の取得 ==
    w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	M00_KANRI "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & " 	M00_KANRI "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & " 		M00_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & " 	AND M00_NO    = " & C_K_WAREKI_NENDO

    '== データ取得 ==
    If gf_GetRecordset(w_Rs,w_sSql) <> 0 Then Exit function

    f_Nendo2Wareki = gf_SetNull2String(w_Rs("M00_KANRI"))

    '== 閉じる ==
    call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  必選区分名称を取得
'********************************************************************************
Function f_GetHissenNM(p_iHissen)
    Dim w_sSQL
    Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_GetHissenNM = ""

    '== 和暦の取得 ==
    w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & " 		M01_NENDO        = " & m_iNendo
    w_sSQL = w_sSQL & " 	AND M01_SYOBUNRUI_CD = " & p_iHissen
	w_sSQL = w_sSQL & " 	AND M01_DAIBUNRUI_CD = " & C_HISSEN

    '== データ取得 ==
    If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

	f_GetHissenNM = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))

    '== 閉じる ==
    call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  授業時間数をセット
'********************************************************************************
Sub s_GetJikan(p_sJ_JunJikan)

	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sJ_JunJikan = m_Rs("J_JUNJI_ZT")
		Case Else
			p_sJ_JunJikan = m_Rs("J_JUNJI_ZK")
	End Select

End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
Sub showPage()

	Dim w_sSeiseki
	Dim w_sHyoka
	Dim w_sKekka_ZK
	Dim w_sKekka_KT
	Dim w_sKekka_KK
	Dim w_sKibi_ZK
	Dim w_sKibi_KT
	Dim w_sKibi_KK
	Dim w_sTeisi_ZK
	Dim w_sTeisi_KT
	Dim w_sTeisi_KK
	Dim w_sHaken_ZK
	Dim w_sHaken_KT
	Dim w_sHaken_KK
	Dim w_sSeiseki_ZK
	Dim w_sSeiseki_KT
	Dim w_sSeiseki_KK
	Dim w_sHyoka_ZK
	Dim w_sHyoka_KT
	Dim w_sHyoka_KK
	Dim w_bNoChange_ZK
	Dim w_bNoChange_KT
	Dim w_bNoChange_KK
	Dim i
	Dim w_IdouKbn									'異動タイプ
	Dim w_IdouName
	Dim w_sInputClass
	Dim w_Padding
	Dim w_cell
	Dim w_sJ_JunJikan_Z

	w_Padding   = "style='padding:2px 0px;font-size:10px;text-align:center'"
	w_Padding2  = "style='padding:2px 0px;font-size:10px;writing-mode:tb-rl'"
	w_Padding3  = "style='padding:2px 0px;font-size:10px'"

	i = 1

	'//前期中間 or 前期期末（試験区分によって分岐）データセット
	Call s_GetJikan(w_sJ_JunJikan_Z)

	'//NN対応
	If session("browser") = "IE" Then
		w_sInputClass  = "class='num'"
	Else
		w_sInputClass = ""
	End If

%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<!--OBJECT ID="thebrowser" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" -->
<!--/OBJECT -->
<SCRIPT language="javascript">
<!--
	function window_onload(){
		alert("<%=C_TOUROKU_OK_MSG%>");
		window.focus();
		window.print();
//		thebrowser.Execwb(6,2);
		document.frm.target = "main";
		document.frm.action = "sei0150_23_bottom.asp"
		document.frm.submit();
	}
//-->
</SCRIPT>
<style TYPE="text/css">
table.hyo1 {
	border-layout : fixed;
	border-collapse:collapse;
	border-style:solid;
	border-width:1px;
	padding:0px;
	margin:0px;
}
td.head1 { 
	font-size:8pt;
	padding:2px 5px;
}
td.head2 { 
	font-size:8pt;
	padding:2px 5px;
	writing-mode:tb-rl;
}
td.head3 { 
	font-size:10pt;
	padding:2px 5px;
}
p.margin1 {
	margin: 0px 0px 0px 0px
}
<!--
	@media screen,print{
		BODY {
			margin: 0;  ?*ブロック領域とブロック枠の余白幅ゼロ指定*?
			padding: 0; ?*ブロック枠とブロック文書の余白幅ゼロ指定*?
		}
	}
//-->
</style>
</head>
<body LANGUAGE="javascript" onload="window_onload();">
<center>
<form name="frm" method="post">
	<p class="margin1"></p>
	<table aling="center">
		<tr>
			<td class="head1" colspan="3" height="10"></td>
			<th width="350" align="center" rowspan="2"><font size="5">成　績　評　価　表</font></th>
			<td class="head1"></td>
		</tr>
		<tr>
			<td class="head3" aling="right">平成</td>
			<td class="head3" aling="center"><%=f_Nendo2Wareki(m_iNendo)%></td>
			<td class="head3">年度</td>
			<td class="head3"><%=f_GetSchoolName%></td>
		</tr>
	</table>
	<table aling="center" cellpadding="0" cellspacing="0">
		<tr height="5">
			<td></td>
		</tr>
	</table>
	<table aling="center" cellpadding="0" cellspacing="0">
		<tr>
			<td class="head3" width="140" align="center"><%=gf_GetClassName(m_iNendo,m_iGakunen,m_sClassNo)%></td>
			<td class="head3" width="140" align="center">第 <%=m_iGakunen%> 学年</td>
			<td class="head3" width="140" align="center"><%=f_ShikenMei%></td>
			<td class="head3" width="230" align="right"><%=m_sUpdDate%>　登録</td>
		</tr>
	</table>
	<table>
		<tr>
			<td>
				<table class="hyo1" align="center" border="1">
					<tr>
						<td class="head1" colspan="3"  align="center" nowrap>教科目</td>
						<td class="head1" colspan="2"  align="center" nowrap>単位数</td>
						<td class="head1" colspan="19" align="center" nowrap>担当教官氏名</td>
					</tr>
					<tr>
						<td class="head1" colspan="3" rowspan="2"  align="center" nowrap><%=gf_GetKamokuMei(m_iNendo,m_sKamokuCd,m_iKamoku_Kbn)%></td>
						<td class="head1" colspan="2" align="center" nowrap><%=f_GetHissenNM(m_iHissen_Kbn)%></td>
						<td class="head2" rowspan="2" align="center" nowrap>前期</td>
						<td class="head1" colspan="5" rowspan="2"  align="center" nowrap>
							<table>
								<tr>
									<td class="head1" width="85" align="center" nowrap><%=Session("USER_NM")%></td>
									<td class="head1" align="right" nowrap>印</td>
								</tr>
							</table>
						</td>
						<td class="head2" rowspan="2" align="center" nowrap>後期</td>
						<td class="head1" colspan="5" rowspan="2"  align="center" nowrap>
							<table>
								<tr>
									<td class="head1" width="85" align="center" nowrap><%=Session("USER_NM")%></td>
									<td class="head1" align="right" nowrap>印</td>
								</tr>
							</table>
						</td>
						<td class="head2" rowspan="2" align="center" nowrap>学年</td>
						<td class="head1" colspan="5" rowspan="2"  align="center" nowrap>
							<table>
								<tr>
									<td class="head1" width="85" align="center" nowrap><%=Session("USER_NM")%></td>
									<td class="head1" align="right" nowrap>印</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class="head1" colspan="2" align="center" nowrap><%=gf_SetNull2String(m_Rs("HAITOTANI"))%></td>
					</tr>
					<tr>
						<td class="head2" rowspan="4" align="center" width="15"  nowrap>名列番号</td>
						<td class="head1" colspan="2" align="center" nowrap>期</td>
						<td class="head1" colspan="4" align="center" nowrap>前　期</td>
						<td class="head1" colspan="4" align="center" nowrap>後　期</td>
						<td class="head1" colspan="4" align="center" nowrap>学　年</td>
						<td class="head1" colspan="3" rowspan="2" align="center" nowrap>成績評価</td>
						<td class="head1" colspan="5" rowspan="4" align="center" nowrap>備　考</td>
					</tr>
					<tr>
						<td class="head1" colspan="2" align="center" nowrap>授業時数</td>
						<td class="head1" colspan="4" align="center" nowrap><%=gf_SetNull2String(w_sJ_JunJikan_Z)%> 時間</td>
						<td class="head1" colspan="4" align="center" nowrap><%=gf_SetNull2String(m_Rs("J_JUNJI_KT"))%> 時間</td>
						<td class="head1" colspan="4" align="center" nowrap><%=gf_SetNull2String(m_Rs("J_JUNJI_KK"))%> 時間</td>
					</tr>
					<tr>
						<td class="head1" colspan="2" align="center" nowrap>欠課時数</td>
						<td class="head2" align="center" rowspan="2" nowrap>欠課</td>
						<td class="head2" align="center" rowspan="2" nowrap>忌引</td>
						<td class="head2" align="center" rowspan="2" nowrap>停止</td>
						<td class="head2" align="center" rowspan="2" nowrap>派遣</td>
						<td class="head2" align="center" rowspan="2" nowrap>欠課</td>
						<td class="head2" align="center" rowspan="2" nowrap>忌引</td>
						<td class="head2" align="center" rowspan="2" nowrap>停止</td>
						<td class="head2" align="center" rowspan="2" nowrap>派遣</td>
						<td class="head2" align="center" rowspan="2" nowrap>欠課</td>
						<td class="head2" align="center" rowspan="2" nowrap>忌引</td>
						<td class="head2" align="center" rowspan="2" nowrap>停止</td>
						<td class="head2" align="center" rowspan="2" nowrap>派遣</td>
						<td class="head2" align="center" rowspan="2" nowrap>前期</td>
						<td class="head2" align="center" rowspan="2" nowrap>後期</td>
						<td class="head2" align="center" rowspan="2" nowrap>学年</td>
					</tr>
					<tr>
						<td class="head1" width="55" align="center" nowrap>学生番号</td>
						<td class="head1" width="90" align="center" nowrap>学　生　氏　名</td>
					</tr>

				<%
					m_Rs.MoveFirst
					Do Until m_Rs.EOF
						j = j + 1 

						w_sKekka_ZK = ""
						w_sKekka_KT = ""
						w_sKekka_KK = ""
						w_sKibi_ZK  = ""
						w_sKibi_KT  = ""
						w_sKibi_KK  = ""
						w_sTeisi_ZK = ""
						w_sTeisi_KT = ""
						w_sTeisi_KK = ""
						w_sHaken_ZK = ""
						w_sHaken_KT = ""
						w_sHaken_KK = ""
						w_sSeiseki  = ""
						w_sHyoka    = ""
						w_bNoChange = false

						Call gs_cellPtn(w_cell)

						'//欠課、遅刻数のセット
						Call s_SetKekka(w_sKekka_ZK, w_sKekka_KT, w_sKekka_KK, _
										w_sKibi_ZK , w_sKibi_KT , w_sKibi_KK, _
										w_sTeisi_ZK, w_sTeisi_KT, w_sTeisi_KK, _
										w_sHaken_ZK, w_sHaken_KT, w_sHaken_KK)

						'//成績データセット
						Call s_SetGrades(w_sSeiseki_ZK, w_sSeiseki_KT, w_sSeiseki_KK, _
										 w_sHyoka_ZK, w_sHyoka_KT, w_sHyoka_KK, _
										 w_bNoChange_ZK, w_bNoChange_KT, w_bNoChange_KK)

						'//異動チェック
						Call s_IdouCheck(m_Rs("GAKUSEKI_NO"),w_IdouKbn,w_IdouName,w_bNoChange_ZK, w_bNoChange_KT, w_bNoChange_KK)
				%>
					<tr>
						<td class="<%=w_cell%>" align="center" nowrap <%=w_Padding3%>><%=i%></td>
						<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding3%>><%=m_Rs("GAKUSEKI_NO")%></td>
						<td class="<%=w_cell%>" align="left"   width="90" nowrap <%=w_Padding3%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>

						<!-- 欠課 -->
						<!-- 前期期末 -->
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKekka_ZK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKibi_ZK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sTeisi_ZK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sHaken_ZK%></td>
						<!-- 後期期末 -->
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKekka_KT%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKibi_KT%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sTeisi_KT%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sHaken_KT%></td>
						<!-- 学年末 -->
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKekka_KK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKibi_KK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sTeisi_KK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sHaken_KK%></td>

						<!--選択科目の時に未選択の場合、入力不可。また、休学など-->
						<% If w_bNoChange_ZK = True Then %>
							<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
							<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
							<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>

						<!-- 成績 (数値入力、文字入力、成績なし入力により処理を分ける) -->
						<% Else %>
							<!-- 数値入力 -->
							<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_ZK%></td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_KT%></td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_KK%></td>

							<!-- 文字入力 -->
							<% elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_ZK%></td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_KT%></td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_KK%></td>

							<!-- 以外 -->
							<% else %>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
							<% end if
						End If %>
						<td class="<%=w_cell%>" width="125" colspan="5" nowrap <%=w_Padding%>></td>
					</tr>
					<%

							if (i Mod 5) = 0 then
								Response.write "<tr>"
									Response.write "<td colspan='23'>"
									Response.write "</td>"
								Response.write "</tr>"
							end if

							m_Rs.MoveNext
							i = i + 1
						Loop
					%>

				</table>
			</td>
		</tr>
	</table>

	<input type="hidden" name="txtNendo"     value="<%=trim(Request("txtNendo"))%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=trim(Request("txtKyokanCd"))%>">
	<input type="hidden" name="sltShikenKbn" value="<%=trim(Request("sltShikenKbn"))%>">
	<input type="hidden" name="txtGakuNo"    value="<%=trim(Request("txtGakuNo"))%>">
	<input type="hidden" name="txtClassNo"   value="<%=trim(Request("txtClassNo"))%>">
	<input type="hidden" name="txtKamokuCd"  value="<%=trim(Request("txtKamokuCd"))%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=trim(Request("txtGakkaCd"))%>">
	<input type="hidden" name="hidKamokuKbn" value="<%=request("hidKamokuKbn")%>">
	<input type="hidden" name="hidSyubetu"   value="<%=request("hidSyubetu")%>">

</form>
</center>
</body>
</html>
<%
End sub
%>