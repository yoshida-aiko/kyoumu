<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0150/sei0150_23_hidprint.asp
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
' 作      成: 2005/10/04
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

    Dim m_iIdouEnd        '//異動対象期間終了日

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

		'//成績、欠課入力期間チェック
		If not f_Nyuryokudate() Then
			m_bErrFlg = True
			Exit Do
		End If

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
		'Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
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
'	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z  AS TEISI_ZT,"			'停止（前期中間）
'	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z  AS TEISI_ZK,"			'停止（前期期末）
'	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K  AS TEISI_KT,"			'停止（期末中間）
'	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_K  AS TEISI_KK,"			'停止（後期期末）
'	w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_Z       AS KIBI_ZT, "			'忌引（前期中間）
'	w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_Z       AS KIBI_ZK, "			'忌引（前期期末）
'	w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_K       AS KIBI_KT, "			'忌引（後期期末）
'	w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_K       AS KIBI_KK, "			'忌引（後期期末）
'	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_Z   AS HAKEN_ZT,"			'派遣（前期中間）
'	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_Z   AS HAKEN_ZK,"			'派遣（前期期末）
'	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_K   AS HAKEN_KT,"			'派遣（後期中間）
'	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_K   AS HAKEN_KK,"			'派遣（後期期末）
'	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z    AS SOUJI_ZT,"
'	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z    AS SOUJI_ZK,"
'	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K    AS SOUJI_KT,"
'	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K    AS SOUJI_KK,"
'	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z   AS JUNJI_ZT,"
'	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z   AS JUNJI_ZK,"
'	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K   AS JUNJI_KT,"
'	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K   AS JUNJI_KK,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_TYUKAN_Z AS J_JUNJI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_KIMATU_Z AS J_JUNJI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_TYUKAN_K AS J_JUNJI_KT,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_KIMATU_K AS J_JUNJI_KK,"
'	w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_Z      AS HYOKA_ZT,  "
'	w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_Z      AS HYOKA_ZK,  "
'	w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_K      AS HYOKA_KT,  "
'	w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_K      AS HYOKA_KK,  "
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_Z   AS KOUSINBI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_Z   AS KOUSINBI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_K   AS KOUSINBI_KT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_K   AS KOUSINBI_KK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_TYUKAN_Z AS KOUSINTIME_ZT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_KIMATU_Z AS KOUSINTIME_ZK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_TYUKAN_K AS KOUSINTIME_KT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_KIMATU_K AS KOUSINTIME_KK,"
'	w_sSQL = w_sSQL & w_Table & "_HYOKA_FUKA_KBN      AS HYOKA_FUKA, "
'	w_sSQL = w_sSQL & w_Table & "_HAITOTANI           AS HAITOTANI, "

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


''response.write w_sSQL
''response.end

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
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]
'********************************************************************************
Function f_Nyuryokudate()

	Dim w_sSysDate
	Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_Nyuryokudate = false

	m_bKekkaNyuryokuFlg = false		'欠課入力ﾌﾗｸﾞ
	m_bSeiInpFlg = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "     T24_IDOU_SYURYO "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T24_SIKEN_NITTEI "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "     T24_NENDO=" & Cint(m_iNendo)
	w_sSQL = w_sSQL & " AND T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
	w_sSQL = w_sSQL & " AND T24_SIKEN_CD='0'"
	w_sSQL = w_sSQL & " AND T24_GAKUNEN=" & m_iGakunen

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function

	If w_Rs.EOF Then
		exit function
	Else
		m_iIdouEnd = gf_SetNull2String(w_Rs("T24_IDOU_SYURYO"))  '異動対象終了
	End If

	'入力期間内なら正常
	If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
		m_bSeiInpFlg = true
	End If

	'欠課入力可能ﾌﾗｸﾞ
	If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
		m_bKekkaNyuryokuFlg = True
	End If

	f_Nyuryokudate = true

End Function


'********************************************************************************
'*  [機能] 異動チェック
'********************************************************************************
Sub s_IdouCheck(p_GakusekiNo,p_IdouKbn,p_IdouName,p_bNoChangeZK,p_bNoChangeKT,p_bNoChangeKK,p_IdouDate)
	Dim w_IdoutypeName	'異動状況名
	Dim w_IdouDate
	w_IdoutypeName = ""
	w_IdouDate = ""

	p_IdouName = ""
	p_IdouDate = ""

	m_Date = m_iIdouEnd

	Call f_Get_IdouChk(p_GakusekiNo,m_Date,m_iNendo,w_IdoutypeName,p_IdouKbn,w_IdouDate)

	if Cstr(p_IdouKbn) <> "" and Cstr(p_IdouKbn) <> CStr(C_IDO_FUKUGAKU) AND _
		Cstr(p_IdouKbn) <> Cstr(C_IDO_TEI_KAIJO) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKO) AND _
		Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKA) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_KOKUHI) Then

		p_IdouName = "[" & w_IdoutypeName & "]"
		p_IdouDate = w_IdouDate

		p_bNoChangeZK = True
		p_bNoChangeKT = True
		p_bNoChangeKK = True
	end if

end Sub

'********************************************************************************
'*	[機能]	異動ありの場合移動状況の取得
'*	[引数]	p_Gakusei_No:学績NO
'*			p_Date		:授業実施日
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	2001.12.19 版：岡田
'********************************************************************************
Function f_Get_IdouChk(p_Gakuseki_No,p_Date,p_iNendo,ByRef p_sKubunName,ByRef p_sKubunCD,ByRef p_sIdouDate)

	Dim w_sSQL
	Dim w_Rs
	Dim w_IdoFlg
	Dim w_sKubunName

	On Error Resume Next
	Err.Clear

	f_Get_IdouChk = False
	w_IdoFlg = False

	Do

		'// 明細データ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_8, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_8"
		w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cint(p_iNendo) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO='" & p_Gakuseki_No & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM>0"

'response.write w_sSQL

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			Exit Do
		End If

		If w_Rs.EOF = 0 Then
			i = 1
			'//8…最大移動回数
			Do Until Cint(i) > cint(8)    '//C_IDO_MAX_CNT = 8…最大移動回数
				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) = "" Then
					Exit Do
				End If
'Response.Write "[" & gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) & " > " & p_Date & "]"
				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) > p_Date  Then
					'//1件目の異動が対象日付より未来の場合の処理
					If i = 1 then
						i = 0
					End if

					Exit Do
				End If
				i = i + 1
			Loop

'response.write "学生ＮＯ" & p_Gakuseki_No & " : i = " & i
			w_sKubunName = ""

			If i = 1 then
				'//最初の移動日が授業日より未来の場合、授業日に移動状態ではない
				'w_IdoFlg = False
				'w_sKubunName = ""

				w_sKubunName = gf_SetNull2String(w_Rs("T13_IDOU_KBN_" & i))
				p_sIdouDate = gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i))
				w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i)),p_iNendo,p_sKubunName)
			Elseif i = 0 then '//1件目の異動が対象日付より未来の場合

				w_bRet = False
				w_sKubunName = ""
				p_sIdouDate = ""
			Else

   				w_sKubunName = gf_SetNull2String(w_Rs("T13_IDOU_KBN_" & i-1))
				p_sIdouDate = gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i-1))

				 w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i-1)),p_iNendo,p_sKubunName)

			End If
'response.write "結果:" & w_sKubunName & "異動事由：" & p_sKubunName  & p_sIdouDate
		End If

		Exit Do
	Loop

	p_sKubunCD = w_sKubunName

	Call gf_closeObject(w_Rs)

	Err.Clear

	f_Get_IdouChk = True

End Function



'********************************************************************************
'*  [機能] 成績のセット
'********************************************************************************
Sub s_SetGrades(p_sSeiseki, _
				p_bNoChange_ZK, p_bNoChange_KT, p_bNoChange_KK)

'	Dim w_UpdDateZK
'   Dim w_UpdDateKK


	'/試験区分によって取ってくる、フィールドを変える。
	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI1"))
		Case C_SIKEN_ZEN_KIM
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI2"))
		Case C_SIKEN_KOU_TYU
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI3"))
		Case C_SIKEN_KOU_KIM
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI4"))
		Case Else
			p_sSeiseki = ""
	End Select

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
'*  [機能]  欠課数のセット
'********************************************************************************
Sub s_SetKekka(p_sKekka)


	'/試験区分によって取ってくる、フィールドを変える。
	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sKekka = gf_SetNull2String(m_Rs("KEKA_ZT"))
		Case C_SIKEN_ZEN_KIM
			p_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))
		Case C_SIKEN_KOU_TYU
			p_sKekka = gf_SetNull2String(m_Rs("KEKA_KT"))
		Case C_SIKEN_KOU_KIM
			p_sKekka = gf_SetNull2String(m_Rs("KEKA_KK"))
		Case Else
			p_sKekka = ""
	End Select

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
'*  [機能]  授業時間数をセット
'********************************************************************************
Sub s_GetJikan(p_sJ_JunJikan)


	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sJ_JunJikan = m_Rs("J_JUNJI_ZT")
		Case C_SIKEN_ZEN_KIM
			p_sJ_JunJikan = m_Rs("J_JUNJI_ZK")
		Case C_SIKEN_KOU_TYU
			p_sJ_JunJikan = m_Rs("J_JUNJI_KT")
		Case C_SIKEN_KOU_KIM
			p_sJ_JunJikan = m_Rs("J_JUNJI_KK")
		Case Else
			p_sJ_JunJikan = ""
	End Select

End Sub


'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
Sub showPage()

	Dim w_sSeiseki
	Dim w_bNoChange
	Dim w_sKekka

	Dim w_bNoChange_ZK
	Dim w_bNoChange_KT
	Dim w_bNoChange_KK
	Dim i
	Dim w_IdouKbn									'異動タイプ
	Dim w_IdouName
	Dim w_IdouDate
	Dim w_sInputClass
	Dim w_Padding
	Dim w_cell
	Dim w_sJ_JunJikan

	w_Padding   = "style='padding:2px 0px;font-size:10px;text-align:center'"
	w_Padding2  = "style='padding:2px 0px;font-size:10px;writing-mode:tb-rl'"
	w_Padding3  = "style='padding:2px 0px;font-size:10px'"

	i = 1

	'//授業時数（試験区分によって分岐）データセット
	Call s_GetJikan(w_sJ_JunJikan)

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
//		alert("<%=C_TOUROKU_OK_MSG%>");
		window.focus();
		window.print();
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
table.hyo2 {
	border-layout : fixed;
	border-collapse:collapse;
	border-style:solid;
	border-width:2px;
	padding:0px;
	margin:0px;
}
td.head1 {
	font-size:10pt;
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
    <BR><BR>
	<table aling="center">
		<tr>
			<td class="head1" aling="right">平成<%=f_Nendo2Wareki(m_iNendo)%>年度</td>
			<td class="head1">　専攻科成績評価票</td>
			<td class="head1"></td>
		</tr>
	</table>
	<table aling="center" cellpadding="0" cellspacing="0" border="0">
		<tr height="5">
			<td colspan="3"></td>
		</tr>
		<tr>
			<td class="head1" width="40%" align="right" ><%=f_ShikenMei%></td>
			<td class="head1" align="center" colspan="2">登録日時　<%=m_sUpdDate%></td>
		</tr>

		<tr height="10">
			<td colspan="3"></td>
		</tr>
	</table>
    <BR>

	<table aling="center" cellpadding="0" cellspacing="0" border="0" width="500">
		<tr >
			<td class="head1" width="5%" align="right">&nbsp;</td>
			<td class="head1" width="15%" align="right">&nbsp;</td>
			<td class="head1" width="25%" align="center">専攻名・学年<hr size=1></td>
			<td class="head1" colspan="2" align="left" nowrap><%=gf_GetClassName(m_iNendo,m_iGakunen,m_sClassNo)%>第<%=m_iGakunen%>学年<hr size=1></td>
			<td class="head1" width="20%" align="right">&nbsp;</td>
		</tr>
		<tr >
			<td class="head1" align="right">&nbsp;</td>
			<td class="head1" align="right">&nbsp;</td>
			<td class="head1" align="center">科　　目　　名<hr size=1></td>
			<td class="head1" colspan="2" align="left"><%=gf_GetKamokuMei(m_iNendo,m_sKamokuCd,m_iKamokuKbn)%><hr size=1></td>
			<td class="head1" align="right">&nbsp;</td>
		</tr>
		<tr>
			<td class="head1" align="right">&nbsp;</td>
			<td class="head1" align="right">&nbsp;</td>
			<td class="head1" align="center" nowrap>成績責任教員<hr size=1></td>
			<td class="head1" colspan="2" align="left"><%=Session("USER_NM")%><hr size=1></td>
			<td class="head1" align="right">&nbsp;</td>
		</tr>
	</table>

	<table aling="center" cellpadding="0" cellspacing="0" border="0" width="500">
		<tr height="30">
			<td class="head1" width="5%" align="right">&nbsp;</td>
			<td class="head1" width="15%" align="right">&nbsp;</td>
			<td class="head1" colspan="3" align="center">

				<table  class="hyo1" aling="center" cellspacing=0 cellpadding=0 border="1" width="100%">
					<tr height="30">
						<td class="head1" width="46%" align="center">授業時数</td>
						<td class="head1" width="54%" align="center"><%=gf_SetNull2String(w_sJ_JunJikan)%> 時間</td>
					</tr>
				</table>

			</td>
			<td class="head3" width="25%" align="right">&nbsp;</td>
		</tr>
		<tr height="5">
			<td colspan="3"></td>
		</tr>
	</table>

	<table width="500" >
		<tr>
			<td>
				<table class="hyo2" align="center" border="1" cellspacing=0 cellpadding=1 width="100%" >

					<tr height="30">
						<td class="head2" width="5%" align="center" nowrap></td>
						<td class="head1" width="15%" align="center" nowrap>学籍番号</td>
						<td class="head1" width="25%" align="center" nowrap>氏　　名</td>
						<td class="head1" width="17%" align="center" nowrap>成績評価</td>
						<td class="head1" align="center" nowrap>欠課時数</td>
						<td class="head1" width="25%" align="center" nowrap>備考</td>
					</tr>
					<tr><td colspan='6'>
					</tr>

				<%
					m_Rs.MoveFirst
					Do Until m_Rs.EOF
						j = j + 1

						w_sSeiseki  = ""
						w_sKekka  = ""
						w_bNoChange_ZK = false
						w_bNoChange_KT = false
						w_bNoChange_KK = false

						Call gs_cellPtn(w_cell)

						'//欠課数のセット
						Call s_SetKekka(w_sKekka)

						'//成績データセット
						Call s_SetGrades(w_sSeiseki, _
										 w_bNoChange_ZK, w_bNoChange_KT, w_bNoChange_KK)

						'//異動チェック
						Call s_IdouCheck(m_Rs("GAKUSEKI_NO"),w_IdouKbn,w_IdouName,w_bNoChange_ZK, w_bNoChange_KT, w_bNoChange_KK,w_IdouDate)
				%>
					<tr height="22">
						<td class="<%=w_cell%>" align="center" nowrap <%=w_Padding3%>><font size="2"><%=i%></font></td>
						<td class="<%=w_cell%>" align="center"  nowrap <%=w_Padding3%>><font size="2"><%=m_Rs("GAKUSEI_NO")%></font></td>
						<td class="<%=w_cell%>" align="center"  nowrap <%=w_Padding3%>><font size="2"><%=trim(m_Rs("SIMEI"))%></font></td>


						<!--選択科目の時に未選択の場合、入力不可。また、休学など-->
						<% If w_bNoChange_ZK = True Then %>
							<td class="<%=w_cell%>" align="center"  nowrap <%=w_Padding%>>-</td>


						<!-- 成績 (数値入力、文字入力、成績なし入力により処理を分ける) -->
						<% Else %>
							<!-- 数値入力 -->
							<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
								<td class="<%=w_cell%>" align="center"  nowrap <%=w_Padding%>><font size="2"><%=w_sSeiseki%></font></td>

							<!-- 文字入力 -->
							<% elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
								<td class="<%=w_cell%>" align="center" nowrap <%=w_Padding%>><font size="2"><%=w_sSeiseki%></font></td>

							<!-- 以外 -->
							<% else %>
								<td class="<%=w_cell%>" align="center"  nowrap <%=w_Padding%>>-</td>
							<% end if
						End If %>

						<!-- 欠課 -->
						<td class="<%=w_cell%>" align="center"  nowrap <%=w_Padding%>><font size="2"><%=w_sKekka%></font></td>

						<!-- 備考 -->
						<td class="<%=w_cell%>" nowrap <%=w_Padding%>><%=w_IdouName%><font size="2"><%=w_IdouDate%></font></td>
					</tr>
					<%

							if (i Mod 5) = 0 then
								Response.write "<tr>"
									Response.write "<td colspan='6'>"
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