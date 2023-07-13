<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0150_23_bottom.asp
' 機      能: 下ページ 成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 説      明:
'	(パターン)
'	・通常授業、特別活動
'	・数値入力、文字入力(成績)
'	・評価不能処理(熊本電波のみ)
'	・科目区分(0:一般科目,1:専門科目)
'	・必修選択区分(1:必修,2:選択)
'	・レベル別区分(0:一般科目,1:レベル別科目)を調べる
'-------------------------------------------------------------------------
' 作      成: 2002/06/21 shin
' 変      更: 2003/04/11 hirota 認定状況を学年別にみるように変更
' 変      更: 2009/10/02 iwata  後期開設フラグの取得を追加
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Dim m_bErrFlg					'//ｴﾗｰﾌﾗｸﾞ

	Const C_ERR_GETDATA = "データの取得に失敗しました"

	'氏名選択用のWhere条件
	Dim m_iNendo					'//年度
	Dim m_sKyokanCd				'//教官コード
	Dim m_sSikenKBN				'//試験区分
	Dim m_sGakuNo				'//学年
	Dim m_sClassNo				'//学科
	Dim m_sKamokuCd				'//科目コード
	Dim m_sSikenNm				'//試験名
	Dim m_rCnt						'//レコードカウント
	Dim m_sGakkaCd
	Dim m_iSyubetu				'//出欠値集計方法
	Dim m_iNKaishi
	Dim m_iNSyuryo
	Dim m_iKekkaKaishi
	Dim m_iKekkaSyuryo
	Dim m_iKamoku_Kbn
	Dim m_iHissen_Kbn
	Dim m_ilevelFlg
	Dim m_Rs
	Dim m_SRs
	Dim m_iSouJyugyou			'//総授業時間
	DIm m_iJunJyugyou			'//純授業時間
	Dim m_iSouJyugyou_KK		'//総授業時間（前期期末 + 後期期末）
	Dim m_iJunJyugyou_KK        	'//純授業時間（前期期末 + 後期期末）
	Dim m_bSeiInpFlg				'//入力期間フラグ
	Dim m_bKekkaNyuryokuFlg		'//欠課入力可能ﾌﾗｸﾞ(True:入力可 / False:入力不可)
	Dim m_iShikenInsertType
	Dim m_sSyubetu
	Dim m_iKamokuKbn				'//科目区分(0:通常授業、1:特別科目)
	Dim m_sKamokuBunrui			'//科目分類(01:通常授業、02:認定科目、03:特別科目)
	Dim m_iSeisekiInpType
	Dim m_Date
	Dim m_bZenkiOnly
	Dim m_bKokiOnly				'//後期開設フラグ　(True:後期開設、False:後期開設でない) 2009.10.02 ins
	Dim m_SchoolFlg,m_KekkaGaiDispFlg,m_HyokaDispFlg
	Dim m_MiHyokaFlg
	Dim m_bNiteiFlg
	Dim m_sGakkoNO				'//学校番号

	Dim m_MenjoFlg				'//免除フラグ

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
	m_MiHyokaFlg = false

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

		'2002.12.25 Ins
		'学校番号の取得
		if Not gf_GetGakkoNO(m_sGakkoNO) then Exit Do

'Response.Write "[1]"

		'//評価不能を表示するかチェック
		if not gf_ChkDisp(C_DATAKBN_DISP,m_SchoolFlg) then
			m_bErrFlg = True
			Exit Do
		End If
		m_SchoolFlg = false

'Response.Write "[2]"

		'//評価不能チェックの処理が必要なら
		'//未評価フラグを調べる
		if m_SchoolFlg then
			if not f_GetMihyoka(m_MiHyokaFlg) then
				m_bErrFlg = True
				Exit Do
			end if
		end if

'Response.Write "[3]"

		'//欠課外を表示するかチェック
		if not gf_ChkDisp(C_KEKKAGAI_DISP,m_KekkaGaiDispFlg) then
			m_bErrFlg = True
			Exit Do
		End If

'Response.Write "[4]"

		'//評価予定を表示するかチェック
		if not gf_ChkDisp(C_HYOKAYOTEI_DISP,m_HyokaDispFlg) then
			m_bErrFlg = True
			Exit Do
		End If
		m_HyokaDispFlg = true

'Response.Write "[5]"

		'//成績入力方法の取得(0:点数[C_SEISEKI_INP_TYPE_NUM]、1:文字[C_SEISEKI_INP_TYPE_STRING]、2:欠課、遅刻[C_SEISEKI_INP_TYPE_KEKKA])
		if not gf_GetKamokuSeisekiInp(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iSeisekiInpType) then
			m_bErrFlg = True
			Exit Do
		end if

'Response.Write "[6]"

		'//前期のみ開設か通年か調べる
		if not f_SikenInfo(m_bZenkiOnly) then
			m_bErrFlg = True
			Exit Do
		end if

'2009.10.02 ins str iwata
		'//後期開設か調べる
		if not f_SikenInfo2(m_bKokiOnly) then
			m_bErrFlg = True
			Exit Do
		end if
'2009.10.02 ins end iwata

'Response.Write "[7]"

		'//成績、欠課入力期間チェック
		If not f_Nyuryokudate() Then
			m_bErrFlg = True
			Exit Do
		End If

'Response.Write "[8]"

		'//出欠欠課の取り方を取得
		'//科目区分(0:試験毎,1:累積)
		If gf_GetKanriInfo(m_iNendo,m_iSyubetu) <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

'Response.Write "[9]"

		'//認定前後情報取得
'		if not gf_GetNintei(m_iNendo,m_bNiteiFlg) then
		if not gf_GetGakunenNintei(m_iNendo,cint(m_sGakuNo),m_bNiteiFlg) then	'2003.04.11 hirota
			m_bErrFlg = True
			Exit Do
		end if

'Response.Write "[10]"

		If m_iKamokuKbn = C_JIK_JUGYO then  '通常授業の場合
			'//科目情報を取得
			'//科目区分(0:一般科目,1:専門科目)、及び、必修選択区分(1:必修,2:選択)を調べる
			'//レベル別区分(0:一般科目,1:レベル別科目)を調べる
			If not f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) Then m_bErrFlg = True : Exit Do
		end if

'Response.Write "[11]"

		'//成績、学生データ取得
		If not f_GetStudent() Then m_bErrFlg = True : Exit Do

		If m_Rs.EOF Then
			Call gs_showWhitePage("個人履修データが存在しません。","成績登録")
			Exit Do
		End If

'Response.Write "[12]"

		'//欠課数の取得
		if not gf_GetSyukketuData2(m_SRs,m_sSikenKBN,m_sGakuNo,m_sClassNo,m_sKamokuCd,m_iNendo,m_iShikenInsertType,m_sSyubetu) then
			m_bErrFlg = True
			Exit Do
		end if

'Response.Write "[13]"
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
	m_sGakuNo	 = cint(request("txtGakuNo"))
	m_sClassNo	 = cint(request("txtClassNo"))
	m_sKamokuCd	 = request("txtKamokuCd")
	m_sGakkaCd	 = request("txtGakkaCd")
	m_sSyubetu	 = trim(Request("SYUBETU"))
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
	w_sSQL = w_sSQL & " 	T15_NYUNENDO = " & Cint(m_iNendo)-cint(m_sGakuNo)+1
	w_sSQL = w_sSQL & " AND T15_GAKKA_CD = '" & m_sGakkaCd & "'"
	w_sSQL = w_sSQL & " AND T15_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'"
	w_sSQL = w_sSQL & " AND T15_KAISETU" & m_sGakuNo & "=" & C_KAI_ZENKI

	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

	'Response.Write "0"

	'//戻り値ｾｯﾄ
	If w_Rs.EOF = False Then
		p_bZenkiOnly = True
	End If

	f_SikenInfo = true

	Call gf_closeObject(w_Rs)

End Function

'2009.10.02 ins str iwata
'********************************************************************************
'*  [機能]  後期開設かどうか調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Function f_SikenInfo2(p_bKokiOnly)
    Dim w_sSQL
    Dim w_Rs

    On Error Resume Next
    Err.Clear

    f_SikenInfo2 = false
	p_bKokiOnly = false

	'//科目が後期のみか調べる
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T15_KAMOKU_CD "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T15_RISYU "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	T15_NYUNENDO = " & Cint(m_iNendo)-cint(m_sGakuNo)+1
	w_sSQL = w_sSQL & " AND T15_GAKKA_CD = '" & m_sGakkaCd & "'"
	w_sSQL = w_sSQL & " AND T15_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'"
	w_sSQL = w_sSQL & " AND T15_KAISETU" & m_sGakuNo & "=" & C_KAI_KOUKI

	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

	'Response.Write "0"

	'//戻り値ｾｯﾄ
	If w_Rs.EOF = False Then
		p_bKokiOnly = True
	End If

	f_SikenInfo2 = true

	Call gf_closeObject(w_Rs)

End Function

'2009.10.02 ins end iwata

'********************************************************************************
'*	[機能]	未評価フラグがたっているか調べる
'********************************************************************************
function f_GetMihyoka(p_MiHyokaFlg)

	Dim w_sSQL,w_Rs
	Dim w_Table,w_FieldName,w_FromTable,w_KamokuCd

	On Error Resume Next
	Err.Clear

	f_GetMihyoka = false
	p_MiHyokaFlg = false

	if m_iKamokuKbn = C_JIK_JUGYO then
		w_Table = "T16"
		w_FromTable = "T16_RISYU_KOJIN"
		w_KamokuCd = "T16_KAMOKU_CD"
	else
		w_Table = "T34"
		w_FromTable = "T34_RISYU_TOKU"
		w_KamokuCd = "T34_TOKUKATU_CD"
	end if

	select case m_sSikenKBN
		case C_SIKEN_ZEN_TYU : w_FieldName = w_Table & "_DATAKBN_TYUKAN_Z"
		case C_SIKEN_ZEN_KIM : w_FieldName = w_Table & "_DATAKBN_KIMATU_Z"
		case C_SIKEN_KOU_TYU : w_FieldName = w_Table & "_DATAKBN_TYUKAN_K"
		case C_SIKEN_KOU_KIM : w_FieldName = w_Table & "_DATAKBN_KIMATU_K"
	end select

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	" & w_FieldName & " as MIHYOKA "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & 			w_FromTable
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	" & w_Table & "_NENDO = " & Cint(m_iNendo) & " and "
	w_sSQL = w_sSQL & " 	" & w_KamokuCd & " = '" & m_sKamokuCd & "' and "
	w_sSQL = w_sSQL & " 	" & w_Table & "_HAITOGAKUNEN = " & Cint(m_sGakuNo) & " and "
	w_sSQL = w_sSQL & " 	" & w_Table & "_GAKKA_CD     = '" & m_sGakkaCd & "' and "
	w_sSQL = w_sSQL & 		w_FieldName & "= 4 "

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function

	if not w_Rs.EOF then p_MiHyokaFlg = true

	f_GetMihyoka = true

end function


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
	w_sSQL = w_sSQL & " 	T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(m_sGakuNo) + 1
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

	w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z AS KEKA1,"			'欠課（前期中間）
	w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA2,"			'欠課（前期期末）
	w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K AS KEKA3,"			'欠課（後期中間）
	w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_K AS KEKA4,"			'欠課（後期期末）
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS TEISI1,"		'停止（前期中間）
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS TEISI2,"		'停止（前期期末）
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS TEISI3,"		'停止（後期中間）
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_K AS TEISI4,"		'停止（後期期末）
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_Z AS HAKEN1,"	'派遣（前期中間）
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_Z AS HAKEN2,"	'派遣（前期期末）
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_K AS HAKEN3,"	'派遣（後期中間）
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_K AS HAKEN4,"	'派遣（後期期末）
	w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_Z AS KIBI1,"				'忌引（前期中間）
	w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_Z AS KIBI2,"				'忌引（前期期末）
	w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_K AS KIBI3,"				'忌引（後期中間）
	w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_K AS KIBI4,"				'忌引（後期期末）

	Select Case m_sSikenKBN
		Case C_SIKEN_ZEN_TYU	'前期中間
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI,"
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_TYUKAN_Z   AS DataKbn ,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z      AS KEKA,"
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z   AS SOUJI,"
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z  AS JYUNJI, "
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS TEISI,"
			w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_Z  AS HAKEN,"
			w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_Z      AS KIBI, "
			w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_Z     AS HYOKA, "

			'通常授業
			if m_iKamokuKbn = C_JIK_JUGYO then
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI, "
			end if

		Case C_SIKEN_ZEN_KIM	'前期期末
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI,"
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_Z   AS DataKbn,"
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z   AS SOUJI, "
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z  AS JYUNJI, "
			w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z      AS KEKA_ZK,"			'欠課（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z      AS KEKA,"				'欠課
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS TEISI_ZK,"		'停止（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS TEISI,"			'停止
			w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_Z  AS HAKEN_ZK,"		'派遣（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_Z  AS HAKEN,"		'派遣
			w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_Z      AS KIBI_ZK, "			'忌引（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_Z      AS KIBI, "				'忌引
			w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_Z     AS HYOKA, "

			'通常授業
			if m_iKamokuKbn = C_JIK_JUGYO then
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
			end if

		Case C_SIKEN_KOU_TYU	'後期中間
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI,"
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K   AS SOUJI, "
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K  AS JYUNJI, "
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_TYUKAN_K   AS DataKbn,"
			w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_Z  AS KOUSINBI_ZK, "		'前期末
			w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_K  AS KOUSINBI_TK, "		'後期中間
			w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z      AS KEKA_ZK,"				'欠課（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K      AS KEKA,"				'欠課
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS TEISI_ZK,"			'停止（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS TEISI,"				'停止
			w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_Z  AS HAKEN_ZK,"			'派遣（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_K  AS HAKEN,"			'派遣
			w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_Z      AS KIBI_ZK, "				'忌引（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_K      AS KIBI, "					'忌引
			w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_K     AS HYOKA, "

			'通常授業
			if m_iKamokuKbn = C_JIK_JUGYO then
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
			end if

		Case C_SIKEN_KOU_KIM

			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z      AS KEKA_ZT,"			'欠課（前期中間）
			w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z      AS KEKA_ZK,"			'欠課（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K      AS KEKA_KT,"			'欠課（期末中間）
			w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_K      AS KEKA,"				'欠課
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS TEISI_ZT,"			'停止（前期中間）
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS TEISI_ZK,"			'停止（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS TEISI_KT,"			'停止（期末中間）
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_K AS TEISI,"				'停止
			w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_Z  AS HAKEN_ZK,"			'派遣（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_K  AS HAKEN_KT,"			'派遣（後期期末）
			w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_K  AS HAKEN,"				'派遣
			w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_Z      AS KIBI_ZK, "			'忌引（前期期末）
			w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_K      AS KIBI_KT, "			'忌引（後期期末）
			w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_K      AS KIBI, "				'忌引
			w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_Z     AS HYOKA_ZT, "
			w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_Z     AS HYOKA_ZK, "
			w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_K     AS HYOKA_KT, "
			w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_K     AS HYOKA, "
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z  AS JYUNJI_ZK, "
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z   AS SOUJI_ZK, "
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K  AS JYUNJI_KT, "
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K   AS SOUJI_KT, "
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K   AS SOUJI, "
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K  AS JYUNJI, "
			w_sSQL = w_sSQL & w_Table & "_SAITEI_JIKAN, "
			w_sSQL = w_sSQL & w_Table & "_KYUSAITEI_JIKAN, "
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_K   AS DataKbn,"
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_Z   AS DataKbn_ZK,"

			'通常授業
			if m_iKamokuKbn = C_JIK_JUGYO then
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI_ZT, "
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI_ZK, "
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI_KT, "
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_Z   AS KOUSINBI_ZK, "
				w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_K   AS KOUSINBI_KK, "
			end if
	End Select

	'//INS 2004-12-17 AMANO
	'//通常授業の場合は免除フラグを取得する
	if m_iKamokuKbn = C_JIK_JUGYO then
		w_sSQL = w_sSQL & " T16_MENJYO_FLG  AS Menjo,"
	End if

	w_sSQL = w_sSQL & " T13_GAKUSEI_NO AS GAKUSEI_NO,"
	w_sSQL = w_sSQL & " T13_GAKUSEKI_NO AS GAKUSEKI_NO,"
	w_sSQL = w_sSQL & " T11_SIMEI AS SIMEI, "

	if m_iKamokuKbn = C_JIK_JUGYO then
		w_sSQL = w_sSQL & " 	T16_SELECT_FLG, "
		w_sSQL = w_sSQL & " 	T16_LEVEL_KYOUKAN, "
		w_sSQL = w_sSQL & " 	T16_OKIKAE_FLG, "
	end if

	w_sSQL = w_sSQL & w_Table & "_HYOKA_FUKA_KBN as HYOKA_FUKA "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & 		w_TableName & ","
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI,"
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & 			w_Table & "_NENDO = " & Cint(m_iNendo)
	w_sSQL = w_sSQL & " AND	" & w_Table & "_GAKUSEI_NO = T11_GAKUSEI_NO "
	w_sSQL = w_sSQL & " AND	" & w_Table & "_GAKUSEI_NO = T13_GAKUSEI_NO "
	w_sSQL = w_sSQL & " AND	T13_GAKUNEN = " & cint(m_sGakuNo)
	w_sSQL = w_sSQL & " AND	T13_CLASS = " & cint(m_sClassNo)
	w_sSQL = w_sSQL & " AND	" & w_KamokuName & " = '" & m_sKamokuCd & "' "
	w_sSQL = w_sSQL & " AND	" & w_Table & "_NENDO = T13_NENDO "

	if m_iKamokuKbn = C_JIK_JUGYO then
		'//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
		w_sSQL = w_sSQL & " AND	T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
	end if

	w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "

	'レコード取得
	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit function


	m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
	m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))

	'学年末時のみ前期期末 + 後期期末の総時間・純授業時間を表示する
	If m_sSikenKBN = C_SIKEN_KOU_KIM then
		if gf_SetNull2String(m_Rs("SOUJI_ZK")) <> "" or gf_SetNull2String(m_Rs("SOUJI_KT")) <> "" then
			m_iSouJyugyou_KK = cint(gf_SetNull2Zero(m_Rs("SOUJI_ZK"))) + cint(gf_SetNull2Zero(m_Rs("SOUJI_KT")))
		end if
		if gf_SetNull2String(m_Rs("JYUNJI_ZK")) <> "" or gf_SetNull2String(m_Rs("JYUNJI_KT")) <> "" then
			m_iJunJyugyou_KK = cint(gf_SetNull2Zero(m_Rs("JYUNJI_ZK"))) + cint(gf_SetNull2Zero(m_Rs("JYUNJI_KT")))
		end if
	End if

'2009/06/15 ins str iwata
	'1人目の学生が免除対象の学生だった場合
	If m_iKamokuKbn = C_JIK_JUGYO then
		If ( gf_SetNull2String(m_Rs("Menjo")) = "1" ) OR ( gf_SetNull2Zero(m_Rs("DataKbn")) <> 0 ) Then
			'履修している学生の検索
			If Not m_Rs.EOF Then
				m_Rs.MoveNext
				Do Until m_Rs.EOF
					'免除科目のﾁｪｯｸ
					If ( gf_SetNull2String(m_Rs("Menjo")) <> "1" ) AND ( gf_SetNull2Zero(m_Rs("DataKbn")) = 0 ) Then

						m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
						m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))

						'学年末時のみ前期期末 + 後期期末の総時間・純授業時間を表示する
						If m_sSikenKBN = C_SIKEN_KOU_KIM then
							if gf_SetNull2String(m_Rs("SOUJI_ZK")) <> "" or gf_SetNull2String(m_Rs("SOUJI_KT")) <> "" then
								m_iSouJyugyou_KK = cint(gf_SetNull2Zero(m_Rs("SOUJI_ZK"))) + cint(gf_SetNull2Zero(m_Rs("SOUJI_KT")))
							end if
							if gf_SetNull2String(m_Rs("JYUNJI_ZK")) <> "" or gf_SetNull2String(m_Rs("JYUNJI_KT")) <> "" then
								m_iJunJyugyou_KK = cint(gf_SetNull2Zero(m_Rs("JYUNJI_ZK"))) + cint(gf_SetNull2Zero(m_Rs("JYUNJI_KT")))
							end if
						End if

						Exit Do
					End If
					m_Rs.MoveNext
				loop
			End If
			m_Rs.MoveFirst
		End If
	End If
'2009/06/15 ins str iwata

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
	w_sSQL = w_sSQL & " 	T24_SEISEKI_KAISI, "
	w_sSQL = w_sSQL & " 	T24_SEISEKI_SYURYO, "
	w_sSQL = w_sSQL & " 	T24_KEKKA_KAISI, "
	w_sSQL = w_sSQL & " 	T24_KEKKA_SYURYO, "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI, "
	w_sSQL = w_sSQL & " 	SYSDATE "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T24_SIKEN_NITTEI, "
	w_sSQL = w_sSQL & " 	M01_KUBUN"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD = T24_SIKEN_KBN"
	w_sSQL = w_sSQL & " AND M01_NENDO = T24_NENDO"
	w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
	w_sSQL = w_sSQL & " AND T24_NENDO=" & Cint(m_iNendo)
	w_sSQL = w_sSQL & " AND T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
	w_sSQL = w_sSQL & " AND T24_SIKEN_CD='0'"
	w_sSQL = w_sSQL & " AND T24_GAKUNEN=" & Cint(m_sGakuNo)

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function

	If w_Rs.EOF Then
		exit function
	Else
		m_sSikenNm = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))		'試験名称
		m_iNKaishi = gf_SetNull2String(w_Rs("T24_SEISEKI_KAISI"))		'成績入力開始日
		m_iNSyuryo = gf_SetNull2String(w_Rs("T24_SEISEKI_SYURYO"))		'成績入力終了日
		m_iKekkaKaishi = gf_SetNull2String(w_Rs("T24_KEKKA_KAISI"))		'欠課入力開始
		m_iKekkaSyuryo = gf_SetNull2String(w_Rs("T24_KEKKA_SYURYO"))	'欠課入力終了
		w_sSysDate = gf_SetNull2String(w_Rs("SYSDATE"))					'システム日付
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
'*  [機能]  確定欠課数、遅刻数を取得。
'*  [引数]  p_iNendo　 　：　処理年度
'*          p_iSikenKBN　：　試験区分
'*          p_sKamokuCD　：　科目コード
'*          p_sGakusei 　：　５年間番号
'*  [戻値]  p_iKekka   　：　欠課数
'*          p_ichikoku 　：　遅刻回数
'*          0：正常修了
'*  [説明]  試験区分に入っている、欠課数、遅刻数を取得する。
'*			2002.03.20
'*			NULLを0に変換しないために、関数をモジュール内で作成（CACommon.aspからコピー）
'********************************************************************************
Function f_GetKekaChi(p_iNendo,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku)
	Dim w_sSQL
	Dim w_Rs
	Dim w_sKek,w_sChi
	Dim w_Table,w_TableName
	Dim w_Kamoku

	On Error Resume Next
	Err.Clear

	f_GetKekaChi = false

	p_iKekka = ""
	p_iChikoku = ""

	'特別授業、その他(通常など)の切り分け
	if trim(m_sSyubetu) = "TOKU" then
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_Kamoku = "T34_TOKUKATU_CD"
	else
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_Kamoku = "T16_KAMOKU_CD"
	end if

	'/試験区分によって取ってくる、フィールドを変える。
	Select Case p_iSikenKBN
		Case C_SIKEN_ZEN_TYU
			w_sKek   = w_Table & "_KEKA_TYUKAN_Z"
			w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_Z"
			w_sChi   = w_Table & "_CHIKAI_TYUKAN_Z"
		Case C_SIKEN_ZEN_KIM
			w_sKek   = w_Table & "_KEKA_KIMATU_Z"
			w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_Z"
			w_sChi   = w_Table & "_CHIKAI_KIMATU_Z"
		Case C_SIKEN_KOU_TYU
			w_sKek   = w_Table & "_KEKA_TYUKAN_K"
			w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_K"
			w_sChi   = w_Table & "_CHIKAI_TYUKAN_K"
		Case C_SIKEN_KOU_KIM
			w_sKek   = w_Table & "_KEKA_KIMATU_K"
			w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_K"
			w_sChi   = w_Table & "_CHIKAI_KIMATU_K"
	End Select

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & 	w_sKek   & " as KEKA, "
	w_sSQL = w_sSQL & 	w_sKekG  & " as KEKA_NASI, "
	w_sSQL = w_sSQL & 	w_sChi   & " as CHIKAI "
	w_sSQL = w_sSQL & " FROM "   & w_TableName
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "      " & w_Table & "_NENDO =" & p_iNendo
	w_sSQL = w_sSQL & "  AND " & w_Table & "_GAKUSEI_NO= '" & p_sGakusei & "'"
	w_sSQL = w_sSQL & "  AND " & w_Kamoku & "= '" & p_sKamokuCD & "'"

	If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then exit function

	'//戻り値ｾｯﾄ
	If w_Rs.EOF = False Then
		p_iKekka = gf_SetNull2String(w_Rs("KEKA"))
		p_iChikoku = gf_SetNull2String(w_Rs("CHIKAI"))
	End If

	f_GetKekaChi = true

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能] 異動チェック
'********************************************************************************
Sub s_IdouCheck(p_GakusekiNo,p_IdouKbn,p_IdouName,p_bNoChange)
	Dim w_IdoutypeName	'異動状況名

	w_IdoutypeName = ""
	p_IdouName = ""

	p_IdouKbn = gf_Get_IdouChk(p_GakusekiNo,m_Date,m_iNendo,w_IdoutypeName)

	if Cstr(p_IdouKbn) <> "" and Cstr(p_IdouKbn) <> CStr(C_IDO_FUKUGAKU) AND _
		Cstr(p_IdouKbn) <> Cstr(C_IDO_TEI_KAIJO) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKO) AND _
		Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKA) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_KOKUHI) Then

		p_IdouName = "[" & w_IdoutypeName & "]"
		p_bNoChange = True
	end if

end Sub

'********************************************************************************
'*  [機能] 成績のセット
'********************************************************************************
Sub s_SetGrades(p_sSeiseki,p_sHyoka,p_bNoChange)

	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI1"))
		Case C_SIKEN_ZEN_KIM
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI2"))
		Case C_SIKEN_KOU_TYU
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI3"))
		Case C_SIKEN_KOU_KIM
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI4"))
	End Select

	p_sHyoka   = gf_SetNull2String(m_Rs("HYOKA"))

	'学年末試験の場合のみ
'	If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
'		w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
'		w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))
'		if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
'			p_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))
'			p_sHyoka = gf_SetNull2String(m_Rs("HYOKA_ZK"))
'		End If
'	End If

	'//通常授業のとき
	if m_iKamokuKbn = C_JIK_JUGYO then

		p_bNoChange = False

		'//科目が選択科目の場合は、生徒が選択しているかどうかを判別する。選択しいない生徒は入力不可とする。
		If cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then
			if cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then p_bNoChange = True
		Else
			If Cstr(m_iLevelFlg) = "1" then
				If isNull(m_Rs("T16_LEVEL_KYOUKAN")) = true then
					p_bNoChange = True
				Else
					if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
						p_bNoChange = True
					End if
				End if
			End if
		End if
	End if

end Sub

'********************************************************************************
'*  [機能] 欠課、遅刻の日々計の取得
'********************************************************************************
Sub s_SetKekkaTotal(p_sKekkasu,p_sChikaisu)
	Dim w_sData
	Dim w_iKekka_rui,w_iChikoku_rui

	'//欠課
	p_sKekkasu = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_KEKKA))

	'//1欠課
	w_sData = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_KEKKA_1))

	if p_sKekkasu = "" and w_sData = "" then
		p_sKekkasu = ""
	else
		p_sKekkasu = cint(gf_SetNull2Zero(p_sKekkasu)) + cint(gf_SetNull2Zero(w_sData))
	end if

	'//遅刻数
	p_sChikaisu = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_TIKOKU))

	'//早退数
	w_sData = f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_SOTAI)

	if p_sChikaisu = "" and w_sData = "" then
		p_sChikaisu = ""
	else
		p_sChikaisu = cint(gf_SetNull2Zero(p_sChikaisu)) + cint(gf_SetNull2Zero(w_sData))
	end if


'DEL 2005/02/07 西村
'    累積区分によって前の試験情報を取得する判定はWEBでは行わないように修正

	'「出欠欠課が累積」で「前期中間でない」の場合
'	if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and m_sSikenKBN <> C_SIKEN_ZEN_TYU then
'		'以前の試験で登録されているデータを取得
'		call f_GetKekaChi(m_iNendo,m_iShikenInsertType,m_sKamokuCd,m_Rs("GAKUSEI_NO"),w_iKekka_rui,w_iChikoku_rui)
'
'		'どちらも""の時は""
'		if p_sKekkasu = "" and w_iKekka_rui = "" then
'			p_sKekkasu = ""
'		else
'			p_sKekkasu = cint(gf_SetNull2Zero(p_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))
'		end if
'
'		'どちらも""の時は""
'		if p_sChikaisu = "" and w_iChikoku_rui = "" then
'			p_sChikaisu = ""
'		else
'			p_sChikaisu = cint(gf_SetNull2Zero(p_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
'		end if
'	end if

End Sub

'********************************************************************************
'*  [機能]  欠課、遅刻数のセット
'********************************************************************************
Sub s_SetKekka(p_sKekka,p_sKibi,p_sTeisi,p_sHaken, _
			   p_sKekka_ZK,p_sKibi_ZK,p_sTeisi_ZK,p_sHaken_ZK, _
			   p_sKekka_KK,p_sKibi_KK,p_sTeisi_KK,p_sHaken_KK)

	p_sKekka    = gf_SetNull2String(m_Rs("KEKA"))
	p_sKibi     = gf_SetNull2String(m_Rs("KIBI"))
	p_sTeisi    = gf_SetNull2String(m_Rs("TEISI"))
	p_sHaken    = gf_SetNull2String(m_Rs("HAKEN"))
	p_sKekka_ZK = gf_SetNull2String(m_Rs("KEKA2"))			'前期期末欠課
	p_sKekka_KK = gf_SetNull2String(m_Rs("KEKA3"))			'後期期末欠課
	p_sKibi_ZK  = gf_SetNull2String(m_Rs("KIBI2"))			'前期期末忌引
	p_sKibi_KK  = gf_SetNull2String(m_Rs("KIBI3"))			'後期期末忌引
	p_sTeisi_ZK = gf_SetNull2String(m_Rs("TEISI2"))			'前期期末停止
	p_sTeisi_KK = gf_SetNull2String(m_Rs("TEISI3"))			'後期期末停止
	p_sHaken_ZK = gf_SetNull2String(m_Rs("HAKEN2"))			'前期期末派遣
	p_sHaken_KK = gf_SetNull2String(m_Rs("HAKEN3"))			'後期期末派遣

	'前期末->後期末にセット
	'//学年末試験の場合のみ
'	If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
'	If m_sSikenKBN <> C_SIKEN_KOU_KIM Then
'		Exit Sub
'	End If

	'以下 - 学年末の場合の処理
'	if m_bZenkiOnly = True then
'		w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
'		w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

'		If w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
'			p_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))               '欠課（前期期末）
'			p_sKibi  = gf_SetNull2String(m_Rs("KIBI_ZK"))               '忌引（前期期末）
'			p_sTeisi = gf_SetNull2String(m_Rs("TEISI_ZK"))              '停止（前期期末）
'			p_sHaken = gf_SetNull2String(m_Rs("HAKEN_ZK"))              '派遣（前期期末）
'		End If
'		Exit Sub
'	End if

'	p_sKekka = ""
'	p_sKibi  = ""
'	p_sTeisi = ""
'	p_sHaken = ""

'	w_sKekka   = gf_SetNull2String(m_Rs("KEKA_ZK"))              '欠課（前期期末）
'	w_sKibi    = gf_SetNull2String(m_Rs("KIBI_ZK"))              '忌引（前期期末）
'	w_sTeisi   = gf_SetNull2String(m_Rs("TEISI_ZK"))             '停止（前期期末）
'	w_sHaken   = gf_SetNull2String(m_Rs("HAKEN_ZK"))             '派遣（前期期末）
'	w_sKekka2  = gf_SetNull2String(m_Rs("KEKA_KT"))              '欠課（後期期末）
'	w_sKibi2   = gf_SetNull2String(m_Rs("KIBI_KT"))              '忌引（後期期末）
'	w_sTeisi2  = gf_SetNull2String(m_Rs("TEISI_KT"))             '停止（後期期末）
'	w_sHaken2  = gf_SetNull2String(m_Rs("HAKEN_KT"))             '派遣（後期期末）

'	if w_sKekka <> "" or w_sKekka2 <> "" then                   '欠課計（前期期末 + 後期期末）
'		p_sKekka = cint(gf_SetNull2Zero(w_sKekka)) + cint(gf_SetNull2Zero(w_sKekka2))
'	end if
'	if w_sKibi <> "" or w_sKibi2 <> "" then                     '忌引計（前期期末 + 後期期末）
'		p_sKibi = cint(gf_SetNull2Zero(w_sKibi)) + cint(gf_SetNull2Zero(w_sKibi2))
'	end if
'	if w_sTeisi <> "" or w_sTeisi2 <> "" then                   '停止計（前期期末 + 後期期末）
'		p_sTeisi = cint(gf_SetNull2Zero(w_sTeisi)) + cint(gf_SetNull2Zero(w_sTeisi2))
'	end if
'	if w_sHaken <> "" or w_sHaken2 <> "" then                   '派遣計（前期期末 + 後期期末）
'		p_sHaken = cint(gf_SetNull2Zero(w_sHaken)) + cint(gf_SetNull2Zero(w_sHaken2))
'	End If
End Sub

'********************************************************************************
'*  [機能]  評価不能処理(熊本電波のみ)
'********************************************************************************
Sub s_SetHyoka(p_IdouKbn,p_DataKbn,p_Checked,p_Disabled2,p_Disabled)

	'//評価不能データ設定
	p_DataKbn = 0
	p_Checked = ""
	p_Disabled2 = ""

	p_DataKbn = cint(gf_SetNull2Zero(m_Rs("DataKbn")))

	If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
		w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
		w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

		if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then

			p_DataKbn = cint(gf_SetNull2Zero(m_Rs("DataKbn_ZK")))

		end if
	end if

	if p_Disabled <> "" then p_Disabled2 = "disabled"

	if p_DataKbn = cint(C_HYOKA_FUNO) then
		p_Checked = "checked"
		p_Disabled2 = "disabled"

	elseif p_DataKbn = cint(C_MIHYOKA) then
		p_Disabled2 = "disabled"

	end if

	if not m_bSeiInpFlg Then p_Disabled2 = ""

	select case Cstr(p_IdouKbn)
		case Cstr(C_IDO_KYU_BYOKI),Cstr(C_IDO_KYU_HOKA)
			p_DataKbn = C_KYUGAKU

		case Cstr(C_IDO_TAI_2NEN),Cstr(C_IDO_TAI_HOKA),Cstr(C_IDO_TAI_SYURYO)
			p_DataKbn = C_TAIGAKU
	end select

End Sub

'********************************************************************************
'*  [機能]  テーブルサイズのセット
'********************************************************************************
Sub s_SetTableWidth(p_TableWidth)

	p_TableWidth = 670

	'//評価予定表示フラグオン、または、通常授業のとき
	if m_HyokaDispFlg and Cstr(m_iKamokuKbn) = Cstr(C_TUKU_FLG_TUJO) then
		p_TableWidth = p_TableWidth
	end if

End Sub

'********************************************************************************
'*  [機能]  免除フラグのセット
'********************************************************************************
Sub s_SetMenjo(p_Menjo)

	p_Menjo = 0
	'//特活の場合は0を返す
	If m_iKamokuKbn = C_JIK_JUGYO Then
		p_Menjo = CInt(gf_SetNull2Zero(m_Rs("Menjo")))
	End If

End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
Sub showPage()
	Dim w_sSeiseki
	Dim w_sHyoka
	Dim w_sChikaisu
	Dim w_sKekka,w_sKibi,w_sTeisi,w_sHaken			'欠課、忌引、停止、派遣
	Dim w_sKekka2,w_sKibi2,w_sTeisi2,w_sHaken2		'合計
	Dim w_sKekkaZ,w_sKibiZ,w_sTeisiZ,w_sHakenZ		'前期
	Dim w_sKekkaGai
	Dim w_sKekkasu
	Dim i
	Dim w_lSeiTotal									'成績合計
	Dim w_lGakTotal									'学生人数
	Dim w_IdouKbn										'異動タイプ
	Dim w_IdouName
	Dim w_sInputClass
	Dim w_sInputClass1
	Dim w_sInputClass2
	Dim w_Padding
	Dim w_Padding2
	Dim w_Disabled
	Dim w_Disabled2
	Dim w_TableWidth
	Dim w_sKekka_ZK,w_sKibi_ZK,w_sTeisi_ZK,w_sHaken_ZK
	Dim w_sKekka_KK,w_sKibi_KK,w_sTeisi_KK,w_sHaken_KK

	Dim w_Menjo

	w_Padding = "style='padding:2px 0px;'"
	w_Padding2 = "style='padding:2px 0px;font-size:13px;'"

	w_lSeiTotal = 0
	w_lGakTotal = 0
	i = 1

	'//NN対応
	If session("browser") = "IE" Then
		w_sInputClass  = "class='num'"
		w_sInputClass1 = "class='num'"
		w_sInputClass2 = "class='num'"
	Else
		w_sInputClass = ""
		w_sInputClass1 = ""
		w_sInputClass2 = ""
	End If

	'//テーブルサイズのセット
	Call s_SetTableWidth(w_TableWidth)

	if m_SchoolFlg then
		if m_MiHyokaFlg or (not m_bSeiInpFlg) then
			w_Disabled = "disabled"
		end if
	end if
%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--
	//************************************************************
	//  [機能]  ページロード時処理
	//************************************************************
	function window_onload() {

		document.body.style.cursor = "default";

		//スクロール同期制御
		parent.init();

		//数値入力のときのみ
		<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
			//成績合計値の取得
			f_GetTotalAvg();
		<% end if %>

		//総時間と純時間をhiddenにセット
		document.frm.hidSouJyugyou.value = "<%= m_iSouJyugyou %>";
		document.frm.hidJunJyugyou.value = "<%= m_iJunJyugyou %>";

		//総時間と純時間をhiddenにセット
		document.frm.hidSouJyugyou_KK.value = "<%= m_iSouJyugyou_KK %>";
		document.frm.hidJunJyugyou_KK.value = "<%= m_iJunJyugyou_KK %>";

		document.frm.target = "topFrame";
		document.frm.action = "sei0150_23_middle.asp";
		document.frm.submit();
	}

	//************************************************************
	//  [機能]  評価ボタンが押されたとき
	//************************************************************
	function f_change(p_iS){
		w_sButton = eval("document.frm.button"+p_iS);
		w_sHyouka = eval("document.frm.Hyoka"+p_iS);

		<%If m_sSikenKBN = C_SIKEN_ZEN_TYU Then%>
			if(w_sButton.value == "・") {
				w_sButton.value = "○";
				w_sHyouka.value = "○";
				return true;
			}
			if(w_sButton.value == "○") {
				w_sButton.value = "・";
				w_sHyouka.value = "";
				return true;
			}

		<%Else%>

			if(w_sButton.value == "・") {
				w_sButton.value = "○";
				w_sHyouka.value = "○";
				return true;
			}
			if(w_sButton.value == "○") {
				w_sButton.value = "◎";
				w_sHyouka.value = "◎";
				return true;
			}
			if(w_sButton.value == "◎") {
				w_sButton.value = "・";
				w_sHyouka.value = "";
				return true;
			}
		<%End If%>
	}

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //************************************************************
    function f_Touroku(){
		if(!f_InpCheck()){
			alert("入力値が不正です");
			return false;
		}

		if(!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return false;}

		if(parent.topFrame.document.frm.txtSouJyugyou){
			document.frm.hidSouJyugyou.value = parent.topFrame.document.frm.txtSouJyugyou.value;
		}
		if(parent.topFrame.document.frm.txtJunJyugyou){
			document.frm.hidJunJyugyou.value = parent.topFrame.document.frm.txtJunJyugyou.value;
		}

		//ヘッダ部空白表示
		parent.topFrame.document.location.href="white.asp";

		//登録処理
		<% if m_iKamokuKbn = C_JIK_JUGYO then %>
			document.frm.hidUpdMode.value = "TUJO";
			document.frm.action="sei0150_23_upd.asp";
		<% Else %>
			document.frm.hidUpdMode.value = "TOKU";
			document.frm.action="sei0150_23_upd_toku.asp";
		<% End if %>

		document.frm.target="main";
		document.frm.submit();
	}

	//************************************************************
	//	[機能]	キャンセルボタンが押されたとき
	//************************************************************
	function f_Cancel(){
		parent.document.location.href="default.asp";
	}

	//************************************************************
	//	[機能]	成績の合計と平均を求める
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]	成績入力期間外、期間内によって計算の仕方を変える
	//	[備考]
	//************************************************************
	function f_GetTotalAvg(){
		var i;
		var total;
		var avg;
		var cnt;

		total = 0;
		cnt = 0;
		avg = 0;

		<% If m_bSeiInpFlg Then %>
			//学生数でのループ
			for(i=0;i<<%=m_rCnt%>;i++) {
				//存在するかどうか
				textbox = eval("document.frm.Seiseki" + (i+1));
				if(textbox){
					//未入力チェック
					if (textbox.value != "") {
						//数字でないのは無視する
						if(!isNaN(textbox.value)){
							total = total + parseInt(textbox.value);
						}
					}
					cnt = cnt + 1;
				}
			}

		<% Else %>
			total = document.frm.hidTotal.value;
			cnt   = document.frm.hidGakTotal.value;
		<% End If%>

		document.frm.txtTotal.value = total;

		//四捨五入
		if (cnt!=0){
			avg = total/cnt;
			avg = avg * 10;
			avg = Math.round(avg);
			avg = avg / 10;
		}

		document.frm.txtAvg.value=avg;
	}

	//************************************************************
	//	[機能]	欠課の合計を求める
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]	成績入力期間外、期間内によって計算の仕方を変える
	//	[備考]
	//************************************************************
	function f_GetTotalKekka(p_Obj,p_Val,p_Obj2){
		<% if m_sSikenKBN = C_SIKEN_ZEN_TYU then %>
			return;
		<% end if %>
		var w_sTotalVal

		if(p_Obj.value == ""){
			return;
		}
		if(!isNaN(p_Val)){
			p_Val = 0;
		}
		with(document.frm){
			if(!isNaN(p_Obj.value)){
				eval(p_Obj2).value = parseInt(p_Obj.value) + parseInt(p_Val);
			}
		}
	}

    //************************************************************
    //  [機能]  数値型チェック
    //************************************************************
	function f_CheckNum(pFromName){
		var wFromName,w_len;

		wFromName = eval(pFromName);

		if(isNaN(wFromName.value)){
			wFromName.focus();
			wFromName.select();
			return false;
		}else{
			//桁チェック
			if(wFromName.name.indexOf("Seiseki") != -1){
				if(wFromName.value > 100){
					wFromName.focus();
					wFromName.select();
					return false;
				}
			}

			//欠課・停止は3桁　忌引・派遣は2桁
			if(wFromName.name.indexOf("txtKekka") != -1){
				w_len = 3;
			}else if(wFromName.name.indexOf("txtKibi") != -1){
				w_len = 2;
			}else if(wFromName.name.indexOf("txtTeisi") != -1){
				w_len = 3;
			}else if(wFromName.name.indexOf("txtHaken") != -1){
				w_len = 2;
			}

			if(wFromName.value.length > w_len){
				wFromName.focus();
				wFromName.select();
				return false;
			}

			//マイナスをチェック
			var wStr = new String(wFromName.value)
			if (wStr.match("-")!=null){
				wFromName.focus();
				wFromName.select();
				return false;
			}

			if(wFromName.name.indexOf("txtAvg") == -1){
				//小数点チェック
				w_decimal = new Array();
				w_decimal = wStr.split(".")

				if(w_decimal.length>1){
					wFromName.focus();
					wFromName.select();
					return false;
				}
			}
		}

		return true;
	}

    //************************************************************
    //  [機能]  大小チェック
    //************************************************************
	function f_CheckDaisyou(){
		wObj1 = eval("parent.topFrame.document.frm.txtSouJyugyou");
		wObj2 = eval("parent.topFrame.document.frm.txtJunJyugyou");

		if(wObj1.value != "" && wObj2.value != ""){
			if(wObj1.value < wObj2.value){
				wObj1.focus();
				return false;
			}
		}
		return true;
	}

	//************************************************
	//Enter キーで下の入力フォームに動くようになる
	//引数：p_inpNm	対象入力フォーム名
	//    ：p_frm	対象フォーム
	//　　：i		現在の番号
	//戻値：なし
	//入力フォーム名が、xxxx1,xxxx2,xxxx3,…,xxxxn
	//の名前のときに利用できます。
	//************************************************
	function f_MoveCur(p_inpNm,p_frm,i){
		if (event.keyCode == 13){		//押されたキーがEnter(13)の時に動く。
			<% if w_bNoChange or Not m_bKekkaNyuryokuFlg or m_sSikenKBN = C_SIKEN_KOU_KIM then %>
				i++;
			<% end if %>

			//入力可能のテキストボックスを探す。見つかったらフォーカスを移して処理を抜ける。
	        for (w_li = 1; w_li <= 99; w_li++) {

				<% if Not w_bNoChange And m_bKekkaNyuryokuFlg and m_sSikenKBN <> C_SIKEN_KOU_KIM then %>
					if (p_inpNm == "Seiseki"){
						p_inpNm = "txtKekka";
					}
					else if (p_inpNm == "txtKekka"){
						p_inpNm = "txtKibi";
					}
					else if (p_inpNm == "txtKibi"){
						p_inpNm = "txtTeisi";
					}
					else if (p_inpNm == "txtTeisi"){
						p_inpNm = "txtHaken";
					}
					else if (p_inpNm == "txtHaken"){
						if(document.frm.txtHaken){
							p_inpNm = "Seiseki";
						}else{
							p_inpNm = "txtKekka";
						}
						i++;
						if (i > <%=m_rCnt%>) i = 1; //iが最大値を超えると、はじめに戻る。
					}
				<% else %>
					if (i > <%=m_rCnt%>) i = 1;     //iが最大値を超えると、はじめに戻る。
				<% end if %>

				inpForm = eval("p_frm."+p_inpNm+i);

				//入力可能領域ならフォーカスを移す。
				if (typeof(inpForm) != "undefined") {
					inpForm.focus();			//フォーカスを移す。
					inpForm.select();			//移ったテキストボックス内を選択状態にする。
					break;
				//入力付加なら次の項目へ
				} else{
					i++
				}
	        }
		}else{
			return false;
		}
		return true;
	}

	//************************************************
	//	文字入力時の成績処理
	//
	//************************************************
	function f_SetSeiseki(w_num){
		var ob = new Array();

		ob[0] = eval("parent.topFrame.document.frm.sltHyoka");
		ob[1] = eval("document.frm.Seiseki" + w_num);
		ob[2] = eval("document.frm.hidSeiseki" + w_num);
		ob[3] = eval("document.frm.hidHyokaFukaKbn" + w_num);

		if(ob[0].value.length == 0){
			ob[1].value = "";
			ob[2].value = "";
			ob[3].value = 0;
		}else{
			var vl = ob[0].value.split('#@#');

			ob[1].value = vl[0];
			ob[2].value = vl[0];
			ob[3].value = vl[1];
		}
	}

	//************************************************
	//	入力チェック
	//************************************************
	function f_InpCheck(){
		var w_length;
		var ob;

		//総時間・純時間入力チェック
		if(parent.topFrame.document.frm.txtSouJyugyou){
			if(!f_CheckNum("parent.topFrame.document.frm.txtSouJyugyou")){ return false; }
		}
		if(parent.topFrame.document.frm.txtJunJyugyou){
			if(!f_CheckNum("parent.topFrame.document.frm.txtJunJyugyou")){ return false; }
		}
		// 大小チェック不要 2003.02.20
		// if(!f_CheckDaisyou()){ return false; }

		w_length = document.frm.elements.length;

		for(i=0;i<w_length;i++){
			ob = eval("document.frm.elements[" + i + "]")

			//if(ob.type=="text" && ob.name != "txtAvg"  && ob.name != "txtTotal"){
			if(ob.type=="text" && ob.name != "txtAvg"  && ob.name != "txtTotal" && ob.name != "KekkaSum"){
				ob = eval("document.frm." + ob.name);
				if(!f_CheckNum(ob)){return false;}
			}
		}
		return true;
	}

	//************************************************
	//評価不能がクリックされたときの処理
	//************************************************
	function f_InpDisabled(p_num){

		<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
			var ob = new Array();

			ob[0] = eval("document.frm.chkHyokaFuno" + p_num);
			ob[1] = eval("document.frm.Seiseki" + p_num);

			if(ob[0].checked){
				ob[1].value = "";
				ob[1].disabled = true;

			}else{
				ob[1].disabled = false;
			}
		<% end if %>

		//数値入力のときのみ
		<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
			f_GetTotalAvg();
		<% end if %>
	}

	//************************************************************
	//	[機能]	遅刻、欠課、欠課外の合計の計算
	//************************************************************
	function f_CalcSum(p_Name){
		var w_num;
		var total = 0;
		var cnt = 0;

		var ob_kei = eval("document.frm." + p_Name + "Sum")

		//学生数でのループ
		for(w_num=0;w_num<<%=m_rCnt%>;w_num++) {
			//存在するかどうか
			textbox = eval("document.frm." + p_Name + (w_num+1));
			if(textbox){
				//未入力チェック
				if (textbox.value != "") {
					//数字でないのは無視する
					if(!isNaN(textbox.value)){
						total = total + parseInt(textbox.value);
					}
				}
				cnt = cnt + 1;
			}
		}

		ob_kei.value = total;

	}

	function jf_Print(){
		with(document.frm){
			target="main";
			action = "sei0150_23_print.asp";
			submit();
		}
	}
	//-->
	</SCRIPT>
	</head>
	<body LANGUAGE="javascript" onload="window_onload();">
	<center>
	<form name="frm" method="post">

	<table width="<%=w_TableWidth%>">
	<tr>
	<td>

	<table class="hyo" align="center" width="<%=w_TableWidth%>" border="1">
	<%

		m_Rs.MoveFirst
		Do Until m_Rs.EOF
			j = j + 1

			w_sKibi     = ""
			w_sTeisi    = ""
			w_sHaken    = ""
			w_sSeiseki  = ""
			w_sHyoka    = ""
			w_sChikaisu = ""
			w_sKekka    = ""
			w_sKekkaGai = ""
			w_sKekkasu  = ""
			w_sKekka_ZK = ""
			w_sKibi_ZK  = ""
			w_sTeisi_ZK = ""
			w_sHaken_ZK = ""
			w_sKekka_KK = ""
			w_sKibi_KK  = ""
			w_sTeisi_KK = ""
			w_sHaken_KK = ""
			w_bNoChange = false

			Call gs_cellPtn(w_cell)

			'スタイルシート設定
			if not m_bSeiInpFlg Then
				w_sInputClass1 = "class='" & w_cell & "' style='text-align:right;font-size:13px;' readonly tabindex='-1'"
				w_Disabled = "disabled"
			End if

			if m_sSikenKBN = C_SIKEN_KOU_KIM or Not m_bKekkaNyuryokuFlg Then
				w_sInputClass2 = "class='" & w_cell & "' style='text-align:right;font-size:13px;' readonly tabindex='-1'"
			End if

			'//欠課、遅刻数のセット
			Call s_SetKekka(w_sKekka,w_sKibi,w_sTeisi,w_sHaken, _
							w_sKekka_ZK,w_sKibi_ZK,w_sTeisi_ZK,w_sHaken_ZK, _
							w_sKekka_KK,w_sKibi_KK,w_sTeisi_KK,w_sHaken_KK)

			'//成績データセット
			Call s_SetGrades(w_sSeiseki,w_sHyoka,w_bNoChange)

			'//免除フラグのセット
			Call s_SetMenjo(w_Menjo)

			'//異動チェック
			Call s_IdouCheck(m_Rs("GAKUSEKI_NO"),w_IdouKbn,w_IdouName,w_bNoChange)

			'//欠課、遅刻の日々計の取得
			Call s_SetKekkaTotal(w_sKekkasu,w_sChikaisu)

			'//欠入が0で,欠計が0より大きい場合
			if cint(gf_SetNull2Zero(w_sKekka)) = 0 and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then
				w_sKekka = cint(gf_SetNull2Zero(w_sKekkasu))
			end if

			%>
			<tr>
				<td class="<%=w_cell%>" align="center" width="65"  nowrap <%=w_Padding%>><%=m_Rs("GAKUSEKI_NO")%></td>
				<input type="hidden" name="txtGseiNo<%=i%>"   value="<%=m_Rs("GAKUSEI_NO")%>">
				<input type="hidden" name="hidNoChange<%=i%>" value="<%=w_bNoChange%>">

				<td class="<%=w_cell%>" align="left" width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>

				<!--成績履歴-->
				<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI1")))%></td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI2")))%></td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI3")))%></td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI4")))%></td>
				<% else %>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>>-</td>
				<% end if %>

				<!--選択科目の時に未選択の場合、入力不可。また、休学など-->
				<% If w_bNoChange = True Then %>
					<td class="<%=w_cell%>" align="right" width="50" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>-</td>

				<!-- 成績 (数値入力、文字入力、成績なし入力により処理を分ける) -->
				<% Else %>
					<!--免除フラグが立っていれば、文字として表示-->
					<% If w_Menjo = 1 Then %>
						<% If m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_NUM And m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_STRING Then %>
							<td class="<%=w_cell%>" align="right" width="50" nowrap <%=w_Padding%>>-</td>
						<% Else %>
							<td class="<%=w_cell%>" align="right" width="50" nowrap <%=w_Padding%>><font size="2"><%=w_sSeiseki%></font></td>
						<% End If %>
					<% Else %>
						<!-- 数値入力 -->
						<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
							<td class="<%=w_cell%>" align="right" width="50" nowrap <%=w_Padding2%>>
								<input type="text" <%=w_sInputClass1%> name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size="2" maxlength="3" onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);" onChange="f_GetTotalAvg();">
								<input type="hidden" name="hidSei_ZK<%=i%>" value="<%=w_sSei_ZK%>">
								<input type="hidden" name="hidSei_KK<%=i%>" value="<%=w_sSei_KK%>">
							</td>

						<!-- 文字入力 -->
						<% elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
							<td class="<%=w_cell%>" align="right" width="50" nowrap <%=w_Padding2%>>
								<% if not m_bSeiInpFlg Then %>
									<%=w_sSeiseki%>
								<% else %>
									<input type="button" class="<%=w_cell%>" style="text-align:center;" name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=2 onClick="f_SetSeiseki(<%=i%>);" <%=w_Disabled%>>
								<% end if %>
							</td>
							<input type="hidden" name="hidSeiseki<%=i%>" value="<%=w_sSeiseki%>">
							<input type="hidden" name="hidHyokaFukaKbn<%=i%>" value="<%=m_Rs("HYOKA_FUKA")%>">

						<!-- 以外 -->
						<% else %>
							<td class="<%=w_cell%>" align="right" width="50" nowrap <%=w_Padding%>>-</td>
						<% end if %>
					<% End If %>

					<!-- 欠課 -->
					<!-- 入力期間内ならテキスト表示 -->
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>>
						<% If w_Menjo = 1 Then %>
							<%=w_sKekka%>
						<% Else %>
							<input type="text" <%=w_sInputClass2%> name="txtKekka<%=i%>" value="<%=w_sKekka%>" onKeyDown="f_MoveCur('txtKekka',this.form,<%=i%>);" size="3" maxlength="3">
						<% End If %>
						<input type="hidden" name="hidKeka_ZK<%=i%>" value="<%=w_sKekka_ZK%>">
						<input type="hidden" name="hidKeka_KK<%=i%>" value="<%=w_sKekka_KK%>">

						<!--INS 2004/12/17 Amano 免除フラグ保持フィールド-->
						<INPUT TYPE = "HIDDEN" NAME = "chkMenjo_Flg<%=i%>" VALUE = "<%=w_Menjo%>">
					</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>>
						<% If w_Menjo = 1 Then %>
							<%=w_sKibi%>
						<% Else %>
							<input type="text" <%=w_sInputClass2%> name="txtKibi<%=i%>"  value="<%=w_sKibi%>" onKeyDown="f_MoveCur('txtKibi',this.form,<%=i%>);" size="3" maxlength="2">
						<% End If %>
						<input type="hidden" name="hidKibi_ZK<%=i%>" value="<%=w_sKibi_ZK%>">
						<input type="hidden" name="hidKibi_KK<%=i%>" value="<%=w_sKibi_KK%>">
					</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>>
						<% If w_Menjo = 1 Then %>
							<%=w_sTeisi%>
						<% Else %>
							<input type="text" <%=w_sInputClass2%> name="txtTeisi<%=i%>" value="<%=w_sTeisi%>" onKeyDown="f_MoveCur('txtTeisi',this.form,<%=i%>);" size="3" maxlength="3">
						<% End If %>
						<input type="hidden" name="hidTeisi_ZK<%=i%>" value="<%=w_sTeisi_ZK%>">
						<input type="hidden" name="hidTeisi_KK<%=i%>" value="<%=w_sTeisi_KK%>">
					</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>>
						<% If w_Menjo = 1 Then %>
							<%=w_sHaken%>
						<% Else %>
							<input type="text" <%=w_sInputClass2%> name="txtHaken<%=i%>" value="<%=w_sHaken%>" onKeyDown="f_MoveCur('txtHaken',this.form,<%=i%>);" size="3" maxlength="2">
						<% End If %>
						<input type="hidden" name="hidHaken_ZK<%=i%>" value="<%=w_sHaken_ZK%>">
						<input type="hidden" name="hidHaken_KK<%=i%>" value="<%=w_sHaken_KK%>">
					</td>
<!--
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>><input type="text" <%=w_sInputClass2%> name="txtKekka<%=i%>" value="<%=w_sKekka%>" onKeyDown="f_MoveCur('txtKekka',this.form,<%=i%>);" onChange="f_GetTotalKekka(this,'<%=w_sKekkaZ%>','txtInvisibleKekka<%=i%>');" size=3 maxlength=3></td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>><input type="text" <%=w_sInputClass2%> name="txtKibi<%=i%>"  value="<%=w_sKibi%>" onKeyDown="f_MoveCur('txtKibi',this.form,<%=i%>);" onChange="f_GetTotalKekka(this,'<%=w_sKibiZ%>','txtInvisibleKibi<%=i%>');" size=3 maxlength=3></td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>><input type="text" <%=w_sInputClass2%> name="txtTeisi<%=i%>" value="<%=w_sTeisi%>" onKeyDown="f_MoveCur('txtTeisi',this.form,<%=i%>);" onChange="f_GetTotalKekka(this,'<%=w_sTeisiZ%>','txtInvisibleTeisi<%=i%>');" size=3 maxlength=3></td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding2%>><input type="text" <%=w_sInputClass2%> name="txtHaken<%=i%>" value="<%=w_sHaken%>" onKeyDown="f_MoveCur('txtHaken',this.form,<%=i%>);" onChange="f_GetTotalKekka(this,'<%=w_sHakenZ%>','txtInvisibleHaken<%=i%>');" size=3 maxlength=3></td>
//-->
					<%
						if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							'表示のみの場合の合計・平均値を求める
							If IsNull(w_sSeiseki) = False and IsNumeric(CStr(w_sSeiseki)) = True Then
								w_lSeiTotal = w_lSeiTotal + CLng(w_sSeiseki)
								w_lGakTotal = w_lGakTotal + 1
							End If
						end if
					%>
				<%End If%>
			</tr>
			<%
				m_Rs.MoveNext
				i = i + 1
			Loop
			%>

			<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
				<tr>
					<td class="header" align="right" colspan="7" nowrap>
						<FONT COLOR="#FFFFFF"><B>合計</B></FONT>
						<input type="text" name="txtTotal" size="5" <%=w_sInputClass%> readonly>
					</td>
					<td class="header" align="center" colspan="9" nowrap>&nbsp;</td>
				</tr>

				<tr>
					<td class="header" align="right" colspan="7" nowrap>
						<FONT COLOR="#FFFFFF"><B>　平均点</B></FONT>
						<input type="text" name="txtAvg" size="5" <%=w_sInputClass%> readonly>
					</td>
					<td class="header" align="center" colspan="9" nowrap>&nbsp;</td>
				</tr>
			<% else %>
				<tr>
					<td class="header" align="center" colspan="7" nowrap><FONT COLOR="#FFFFFF"><B>合計</B></FONT></td>
					<td class="header" align="center" colspan="9" nowrap>&nbsp;</td>
				</tr>
			<% end if %>

		</table>

		</td>
		</tr>

		<tr>
		<td align="center">
		<table>
			<tr>
				<td align="center" align="center" colspan="13">
					<%If m_bSeiInpFlg or m_bKekkaNyuryokuFlg Then%>
						<% if m_sSikenKBN = C_SIKEN_KOU_KIM and Not m_bSeiInpFlg then %>
						<% else %>
							<input type="button" class="button" value="　登　録　" onClick="f_Touroku();">
						<% end if %>
					<%End If%>
						<input type="button" class="button" value="キャンセル" onClick="f_Cancel();">
				</td>
			</tr>
		</table>
		</td>
		</tr>
	</table>

	<input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
	<input type="hidden" name="KamokuCd"     value="<%=m_sKamokuCd%>">
	<input type="hidden" name="i_Max"        value="<%=i%>">
	<input type="hidden" name="sltShikenKbn" value="<%=m_sSikenKBN%>">
	<input type="hidden" name="txtGakuNo"    value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=m_sGakkaCd%>">
	<input type="hidden" name="txtClassNo"   value="<%=m_sClassNo%>">
	<input type="hidden" name="txtKamokuCd"  value="<%=m_sKamokuCd%>">
	<input type="hidden" name="PasteType"    value="">
	<input type="hidden" name="hidKikan"     value="<%=m_bSeiInpFlg%>">
	<input type="hidden" name="hidTotal"     value="<%=w_lSeiTotal%>">
	<input type="hidden" name="hidGakTotal"  value="<%=w_lGakTotal%>">
	<input type="hidden" name="txtUpdDate"   value="<%=request("txtUpdDate")%>">
	<input type="hidden" name="hidZenkiOnly" value="<%=m_bZenkiOnly%>">
	<input type="hidden" name="hidKokiOnly" value="<%=m_bKokiOnly%>">	<!-- 2009.10.02 ins -->
	<input type="hidden" name="hidMihyoka"   value ="<%=w_DataKbn%>">
	<input type="hidden" name="hidSchoolFlg" value ="<%=m_SchoolFlg%>">
	<input type="hidden" name="hidKamokuKbn"        value="<%=m_iKamokuKbn%>">
	<input type="hidden" name="hidKamokuBunrui"     value="<%=m_sKamokuBunrui%>">
	<input type="hidden" name="hidSeisekiInpType"   value="<%=m_iSeisekiInpType%>">
	<input type="hidden" name="hidKekkaNyuryokuFlg" value="<%=m_bKekkaNyuryokuFlg%>">
	<input type="hidden" name="hidKekkaGaiDispFlg"  value ="<%=m_KekkaGaiDispFlg%>">
	<input type="hidden" name="hidHyokaDispFlg"     value ="<%=m_HyokaDispFlg%>">
	<input type="hidden" name="hidTableWidth"       value ="<%=w_TableWidth%>">
	<input type="hidden" name="hidFromSei"          value ="<%=m_iNKaishi%>">
	<input type="hidden" name="hidToSei"            value ="<%=m_iNSyuryo%>">
	<input type="hidden" name="hidFromKekka"        value ="<%=m_iKekkaKaishi%>">
	<input type="hidden" name="hidToKekka"          value ="<%=m_iKekkaSyuryo%>">
	<input type="hidden" name="hidSyubetu"          value ="<%=m_sSyubetu%>">
	<input type="hidden" name="hidSouJyugyou">
	<input type="hidden" name="hidJunJyugyou">
	<input type="hidden" name="hidSouJyugyou_KK">
	<input type="hidden" name="hidJunJyugyou_KK">
	<input type="hidden" name="hidUpdMode">

	</form>
	</center>
	</body>
	</html>
<%
End sub
%>