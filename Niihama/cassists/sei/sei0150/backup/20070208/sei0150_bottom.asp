<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0150_bottom.asp
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
' 変      更: 
' 修　    正: 2005/09/30 西村 岐阜高専の場合、遅刻・欠課（日々計）は表示しない
' 修　    正: 2006/10/12 新谷 遅刻・欠課の入力可不可判定で
'                             T16の済単位、再試フラグ、再履修フラグを見るよう修正
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Dim m_bErrFlg				'//ｴﾗｰﾌﾗｸﾞ
    
    Const C_ERR_GETDATA = "データの取得に失敗しました"
    
    '氏名選択用のWhere条件
    Dim m_iNendo				'//科目の年度
    Dim g_iNendo				'//処理年度
    Dim m_sKyokanCd				'//教官コード
    Dim m_sSikenKBN				'//試験区分
    Dim m_sGakuNo				'//学年
    Dim m_sClassNo				'//学科
    Dim m_sKamokuCd				'//科目コード
    Dim m_sSikenNm				'//試験名
    Dim m_rCnt					'//レコードカウント
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
	
	Dim m_bSeiInpFlg			'//入力期間フラグ
	Dim m_bKekkaNyuryokuFlg		'//欠課入力可能ﾌﾗｸﾞ(True:入力可 / False:入力不可)
	
	Dim m_iShikenInsertType
	
	Dim m_sSyubetu
	
	'2002/06/21
	Dim m_iKamokuKbn				'//科目区分(0:通常授業、1:特別科目)
	Dim m_sKamokuBunrui				'//科目分類(01:通常授業、02:認定科目、03:特別科目)
	
	Dim m_iSeisekiInpType
	Dim m_Date
	Dim m_bZenkiOnly
	Dim m_SchoolFlg,m_KekkaGaiDispFlg,m_HyokaDispFlg
	
	Dim m_MiHyokaFlg
	
	Dim m_bNiteiFlg
	Dim m_sGakkoNO	'学校番号 INS 2005/09/30 西村
	
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

		'学校番号を取得 ins 2005/09/30 西村
		if Not gf_GetGakkoNO(m_sGakkoNO) then
	        m_bErrFlg = True
			Exit Do
		end if

		'//評価不能を表示するかチェック
		if not gf_ChkDisp(C_DATAKBN_DISP,m_SchoolFlg) then
			m_bErrFlg = True
			Exit Do
		End If

		'//評価不能チェックの処理が必要なら
		'//未評価フラグを調べる
		if m_SchoolFlg then
			if not f_GetMihyoka(m_MiHyokaFlg) then
				m_bErrFlg = True
				Exit Do
			end if
		end if

		'//欠課外を表示するかチェック
		if not gf_ChkDisp(C_KEKKAGAI_DISP,m_KekkaGaiDispFlg) then
			m_bErrFlg = True
			Exit Do
		End If

		'//評価予定を表示するかチェック
		if not gf_ChkDisp(C_HYOKAYOTEI_DISP,m_HyokaDispFlg) then
			m_bErrFlg = True
			Exit Do
		End If

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

		'//成績、欠課入力期間チェック
		If not f_Nyuryokudate() Then
			m_bErrFlg = True
			Exit Do
		End If

		'//出欠欠課の取り方を取得
		'//科目区分(0:試験毎,1:累積)
		If gf_GetKanriInfo(m_iNendo,m_iSyubetu) <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If

		'//認定前後情報取得
		if not gf_GetNintei(m_iNendo,m_bNiteiFlg) then
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

		'//欠課数の取得
		if not gf_GetSyukketuData2(m_SRs,m_sSikenKBN,m_sGakuNo,m_sClassNo,m_sKamokuCd,m_iNendo,m_iShikenInsertType,m_sSyubetu) then
			m_bErrFlg = True
			Exit Do
		end if

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
	g_iNendo	 = request("txtSyoriNendo")
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
	
	'//戻り値ｾｯﾄ
	If w_Rs.EOF = False Then
		p_bZenkiOnly = True
	End If
	
	f_SikenInfo = true
	
	Call gf_closeObject(w_Rs)
	
End Function

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
	
	if m_iKamokuKbn = C_JIK_JUGYO then	'INS2005/08/08 西村 T16_CURRI_GAKKA_CDの方をみる
		w_sSQL = w_sSQL & " 	" & w_Table & "_CURRI_GAKKA_CD     = '" & m_sGakkaCd & "' and "
	else
		w_sSQL = w_sSQL & " 	" & w_Table & "_GAKKA_CD     = '" & m_sGakkaCd & "' and "
	end if

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
	
	if m_iKamokuKbn = C_JIK_JUGYO then  '通常授業の場合
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	else
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
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
	
	Select Case m_sSikenKBN
		Case C_SIKEN_ZEN_TYU
			
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI,"
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_TYUKAN_Z as DataKbn ,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z AS KEKA,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,"
			w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_Z AS CHIKAI,"
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z as SOUJI,"
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z as JYUNJI, "
			
			if m_iKamokuKbn = C_JIK_JUGYO then
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI, "
			end if
			
		Case C_SIKEN_ZEN_KIM
			
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI,"
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_Z as DataKbn,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS KEKA_NASI,"
			w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_Z AS CHIKAI,"
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z as SOUJI, "
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z as JYUNJI, "
			
			if m_iKamokuKbn = C_JIK_JUGYO then
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
			end if
			
		Case C_SIKEN_KOU_TYU
			
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K AS KEKA,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS KEKA_NASI,"
			w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_K AS CHIKAI,"
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K as SOUJI, "
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K as JYUNJI, "
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_TYUKAN_K as DataKbn,"
			
			if m_iKamokuKbn = C_JIK_JUGYO then
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
			end if
			
		Case C_SIKEN_KOU_KIM
			
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI_ZT,"
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI_ZK,"
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI_KT,"
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI_KK,"
			
			w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z AS KEKA_ZT,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA_ZK,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K AS KEKA_KT,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_K AS KEKA,"
			
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,"
			w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_K AS KEKA_NASI,"
			
			w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_Z AS CHIKAI_ZT,"
			w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_Z AS CHIKAI_ZK,"
			w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_K AS CHIKAI_KT,"
			w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_K AS CHIKAI,"
			
			w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI,"
			
			w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K as SOUJI, "
			w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K as JYUNJI, "
			w_sSQL = w_sSQL & w_Table & "_SAITEI_JIKAN, "
			w_sSQL = w_sSQL & w_Table & "_KYUSAITEI_JIKAN, "
			
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_K as DataKbn,"
			w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_Z as DataKbn_ZK,"
			
			if m_iKamokuKbn = C_JIK_JUGYO then
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI_ZT, "
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI_ZK, "
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI_KT, "
				w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
				
				w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_Z AS KOUSINBI_ZK, "
				w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_K AS KOUSINBI_KK, "
			end if
			
	End Select
	
	w_sSQL = w_sSQL & " T13_GAKUSEI_NO AS GAKUSEI_NO,"
	w_sSQL = w_sSQL & " T13_GAKUSEKI_NO AS GAKUSEKI_NO,"
	w_sSQL = w_sSQL & " T11_SIMEI AS SIMEI, "
	
	if m_iKamokuKbn = C_JIK_JUGYO then
		w_sSQL = w_sSQL & " 	T16_SELECT_FLG, "
		w_sSQL = w_sSQL & " 	T16_LEVEL_KYOUKAN, "
		w_sSQL = w_sSQL & " 	T16_OKIKAE_FLG, "
		'ADD START 2006/10/12 新谷 遅刻・欠課入力の入力可・不可切り替え用
		w_sSQL = w_sSQL & " 	T16_TANI_SUMI, "
		w_sSQL = w_sSQL & " 	T16_SAISI_FLG, "
		w_sSQL = w_sSQL & " 	T16_SAIRISYU_FLG, "
		'ADD END 2006/10/12 新谷 遅刻・欠課入力の入力可・不可切り替え用
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

	'DEL2005/08/08 西村 T13_CLASSを条件にすると他学科から選択している学生が表示されないため
	'w_sSQL = w_sSQL & " AND	T13_CLASS = " & cint(m_sClassNo) 
	w_sSQL = w_sSQL & " AND	" & w_Table & "_CURRI_GAKKA_CD ='" & m_sGakkaCd & "' "	'INS2005/08/08西村

	w_sSQL = w_sSQL & " AND	" & w_KamokuName & " = '" & m_sKamokuCd & "' "
	w_sSQL = w_sSQL & " AND	" & w_Table & "_NENDO = T13_NENDO "
	
	if m_iKamokuKbn = C_JIK_JUGYO then
		'//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
		w_sSQL = w_sSQL & " AND	T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
	end if
	
	w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "
	
	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit function
	
	m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
	m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))
	
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
	w_sSQL = w_sSQL & " AND T24_NENDO=" & Cint(g_iNendo)
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
	
	p_sSeiseki = gf_SetNull2String(m_Rs("SEI"))
	
	'学年末試験の場合のみ
	If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
		w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
		w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))
		
		if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
		'If gf_SetNull2String(m_Rs("SEI")) = "" Then 
			p_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))
		End If
	End If
	
	'//通常授業のとき
	if m_iKamokuKbn = C_JIK_JUGYO then
		
		if m_HyokaDispFlg then
			p_sHyoka = gf_SetNull2String(m_Rs("HYOKAYOTEI"))
			if p_sHyoka = "" then p_sHyoka = "・"
		end if
		
		p_bNoChange = False
		
		'//科目が選択科目の場合は、生徒が選択しているかどうかを判別する。選択しいない生徒は入力不可とする。
		if cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then 
			
			if cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then p_bNoChange = True
			
		else
			if Cstr(m_iLevelFlg) = "1" then
				if isNull(m_Rs("T16_LEVEL_KYOUKAN")) = true then
					p_bNoChange = True
				else
					if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
						p_bNoChange = True
					End if
				End if
			End if
		end if
		
	end if
	
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
	
	'「出欠欠課が累積」で「前期中間でない」の場合
	if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and m_sSikenKBN <> C_SIKEN_ZEN_TYU then 
		'以前の試験で登録されているデータを取得
		call f_GetKekaChi(m_iNendo,m_iShikenInsertType,m_sKamokuCd,m_Rs("GAKUSEI_NO"),w_iKekka_rui,w_iChikoku_rui)
		
		'どちらも""の時は""
		if p_sKekkasu = "" and w_iKekka_rui = "" then
			p_sKekkasu = ""
		else
			p_sKekkasu = cint(gf_SetNull2Zero(p_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))
		end if
		
		'どちらも""の時は""
		if p_sChikaisu = "" and w_iChikoku_rui = "" then
			p_sChikaisu = ""
		else
			p_sChikaisu = cint(gf_SetNull2Zero(p_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
		end if
	end if
	
End Sub

'********************************************************************************
'*  [機能]  欠課、遅刻数のセット
'********************************************************************************
Sub s_SetKekka(p_sKekka,p_sKekkaGai,p_sChikai)

	p_sKekka = gf_SetNull2String(m_Rs("KEKA"))
	p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI"))
	p_sChikai = gf_SetNull2String(m_Rs("CHIKAI"))
	
	'//学年末試験の場合のみ
	If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
		w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
		w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))
		
		'If gf_SetNull2String(m_Rs("KEKA")) = "" Then
		if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
			p_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))			'欠課数
			p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK"))	'欠課対象外
			p_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))		'遅刻回数
		End If
	End If
	
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
	
	p_TableWidth = 610
	
	'//評価不能処理がある(熊本電波のみ)
	if m_SchoolFlg then
		p_TableWidth = 660
	end if
	
	'//評価予定表示フラグオン、または、通常授業のとき
	if m_HyokaDispFlg and Cstr(m_iKamokuKbn) = Cstr(C_TUKU_FLG_TUJO) then
		p_TableWidth = p_TableWidth + 50
	end if
	
	'//欠課外表示フラグオン
	if m_KekkaGaiDispFlg then
		p_TableWidth = p_TableWidth + 55
	end if
	
End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
Sub showPage()
	Dim w_sSeiseki
	Dim w_sHyoka
	
	Dim w_sChikai
	Dim w_sChikaisu
	
	Dim w_sKekka
	Dim w_sKekkaGai
	Dim w_sKekkasu
	
	Dim i
	
	Dim w_lSeiTotal	'成績合計
	Dim w_lGakTotal	'学生人数
	
	Dim w_IdouKbn	'異動タイプ
	Dim w_IdouName
	
	Dim w_sInputClass
	Dim w_sInputClass1
	Dim w_sInputClass2
	
	Dim w_Padding
	Dim w_Padding2
	
	Dim w_Disabled
	Dim w_Disabled2
	Dim w_TableWidth

	w_Padding = "style='padding:2px 0px;'"
	w_Padding2 = "style='padding:2px 0px;font-size:10px;'"
	
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
		
		document.frm.target = "topFrame";
		document.frm.action = "sei0150_middle.asp";
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
		
		document.frm.hidSouJyugyou.value = parent.topFrame.document.frm.txtSouJyugyou.value;
		document.frm.hidJunJyugyou.value = parent.topFrame.document.frm.txtJunJyugyou.value;
		
		//ヘッダ部空白表示
		parent.topFrame.document.location.href="white.asp";
		
		//登録処理
		<% if m_iKamokuKbn = C_JIK_JUGYO then %>
			document.frm.hidUpdMode.value = "TUJO";
			document.frm.action="sei0150_upd.asp";
		<% Else %>
			document.frm.hidUpdMode.value = "TOKU";
			document.frm.action="sei0150_upd_toku.asp";
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
			
			//遅刻は、2桁まで
			if(wFromName.name.indexOf("Chikai") != -1){
				w_len = 2;
			}else{
				w_len = 3;
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
			i++;
			
			//入力可能のテキストボックスを探す。見つかったらフォーカスを移して処理を抜ける。
	        for (w_li = 1; w_li <= 99; w_li++) {
				
				if (i > <%=m_rCnt%>) i = 1; //iが最大値を超えると、はじめに戻る。
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
		if(!f_CheckNum("parent.topFrame.document.frm.txtSouJyugyou")){ return false; }
		if(!f_CheckNum("parent.topFrame.document.frm.txtJunJyugyou")){ return false; }
		if(!f_CheckDaisyou()){ return false; }
		
		w_length = document.frm.elements.length;
		
		for(i=0;i<w_length;i++){
			ob = eval("document.frm.elements[" + i + "]")
			
			if(ob.type=="text" && ob.name != "txtAvg"  && ob.name != "txtTotal"){
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
			
			w_sSeiseki  = ""
			w_sHyoka    = ""
			w_sChikai   = ""
			w_sChikaisu = ""
			w_sKekka    = ""
			w_sKekkaGai = ""
			w_sKekkasu  = ""
			w_bNoChange = false
			
			Call gs_cellPtn(w_cell)
			
			'スタイルシート設定
			if not m_bSeiInpFlg Then
				w_sInputClass1 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
				w_Disabled = "disabled"
			End if
			
			if Not m_bKekkaNyuryokuFlg Then
				w_sInputClass2 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
			End if
			
			'//欠課、遅刻数のセット
			Call s_SetKekka(w_sKekka,w_sKekkaGai,w_sChikai)
			
			'//成績データセット
			Call s_SetGrades(w_sSeiseki,w_sHyoka,w_bNoChange)
			
			'//異動チェック
			Call s_IdouCheck(m_Rs("GAKUSEKI_NO"),w_IdouKbn,w_IdouName,w_bNoChange)
			
			'//欠課、遅刻の日々計の取得
			Call s_SetKekkaTotal(w_sKekkasu,w_sChikaisu)
			
			'//欠入が0で,欠計が0より大きい場合
			if cint(gf_SetNull2Zero(w_sKekka)) = 0 and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then
				w_sKekka = cint(gf_SetNull2Zero(w_sKekkasu))
			end if
			
			'//遅入が0で,遅計が0より大きい場合
			if cint(gf_SetNull2Zero(w_sChikai)) = 0 AND cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then
				w_sChikai = cint(gf_SetNull2Zero(w_sChikaisu))
			end if
			
			'//評価不能処理(熊本電波のみ)
			if m_SchoolFlg  then
				Call s_SetHyoka(w_IdouKbn,w_DataKbn,w_Checked,w_Disabled2,w_Disabled)
			end if

			%>
			
			<tr>
				<td class="<%=w_cell%>" align="center" width="65"  nowrap <%=w_Padding%>><%=m_Rs("GAKUSEKI_NO")%></td>
				<input type="hidden" name="txtGseiNo<%=i%>"   value="<%=m_Rs("GAKUSEI_NO")%>">
				<input type="hidden" name="hidNoChange<%=i%>" value="<%=w_bNoChange%>">

				<% 'ADD START 2006/11/21 新谷
				   '遅刻の表示・非表示、日々計の表示・非表示がある学校は氏名で幅調整をする 
				   Select Case m_sGakkoNO 
					   '北九州
					   Case C_NCT_KITAKYU %>
					<td class="<%=w_cell%>" align="left"   width="<%=gf_IIF(m_KekkaGaiDispFlg,260,315)%>" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>
				<%    '岐阜
					   Case C_NCT_GIFU %>
					<td class="<%=w_cell%>" align="left"   width="<%=gf_IIF(m_KekkaGaiDispFlg,315,370)%>" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>
				<%     'その他
					   Case Else %>
					<td class="<%=w_cell%>" align="left"   width="<%=gf_IIF(m_KekkaGaiDispFlg,150,260)%>" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>
				<% End Select 
				   'ADD END 2006/11/21 新谷
				%>
				
				<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>

					<% '北九州は中間非表示 ADD 2006/11/21 新谷
					   If m_sGakkoNO = C_NCT_KITAKYU Then %>
						<td class="<%=w_cell%>" align="center" width="60"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI2")))%></td>
						<td class="<%=w_cell%>" align="center" width="60"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI4")))%></td>
					<% Else %>
						<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI1")))%></td>
						<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI2")))%></td>
						<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI3")))%></td>
						<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI4")))%></td>
					<% End If %>

				<% else %>

					<% '北九州は中間非表示 ADD 2006/11/21 新谷
					   If m_sGakkoNO = C_NCT_KITAKYU Then %>
						<td class="<%=w_cell%>" align="center" width="60"  nowrap <%=w_Padding2%>>-</td>
						<td class="<%=w_cell%>" align="center" width="60"  nowrap <%=w_Padding2%>>-</td>
					<% Else %>
						<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
						<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
						<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
						<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
					<% End If %>

				<% end if %>
				
				<!--選択科目の時に未選択の場合、入力不可。また、休学など-->
				<% If w_bNoChange = True Then %>
					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
					
					<% if m_HyokaDispFlg and m_iKamokuKbn = C_JIK_JUGYO then %>
						<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
					<% end if %>

					<% 'ADD START 2006/11/21 新谷
					   '遅刻の表示・非表示、日々計の表示・非表示がある学校は列数を調整をする 
					   Select Case m_sGakkoNO 
						   '北九州
						   Case C_NCT_KITAKYU %>
						<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
					<%    '岐阜
						   Case C_NCT_GIFU %>
					<%     'その他
						   Case Else %>
						<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
						<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
						<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
					<% End Select 
					   'ADD END 2006/11/21 新谷
					%>
					
					<% if m_KekkaGaiDispFlg then %>
						<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
					<% end if %>
					
					<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
					
					<% if m_SchoolFlg then %>
						<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
						<input type="hidden" name="chkHyokaFuno<%=i%>" value="<%=w_DataKbn%>">
					<% end if %>
					
				<% Else %>
					
					<%
						'ADD START 2006/10/12 新谷 遅刻・欠課入力の入力可・不可切り替え用
						Dim w_sIptClassWrk
						Dim w_sIptClassWrk1
						'ワーク用変数に退避
						w_sIptClassWrk1 = w_sInputClass1
						w_sIptClassWrk = w_sInputClass2

						'通常授業で欠課入力フラグがONの場合
						If m_iKamokuKbn = C_JIK_JUGYO And m_bKekkaNyuryokuFlg = True Then

							'再試フラグがONの場合
							If gf_SetNull2String(m_Rs("T16_SAISI_FLG")) = "1" Then
								w_sIptClassWrk = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
'response.write "再試フラグがONの場合"
							End if 
							'再履修フラグがONの場合
							If gf_SetNull2String(m_Rs("T16_SAIRISYU_FLG")) = "1" Then
'								w_sIptClassWrk = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
'response.write "再履修フラグがONの場合"
							End If
							'済単位が存在する場合
							If gf_SetNull2String(m_Rs("T16_TANI_SUMI")) <> "" Then
								if gf_SetNull2String(m_Rs("T16_TANI_SUMI")) > "0" Then
'response.write "済単位が存在する場合"
									w_sIptClassWrk = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
									w_sIptClassWrk1 = "class='" & w_cell & "' style='text-align:center;' readonly tabindex='-1'"
								End If
							End If
						End If
						'ADD END 2006/10/12 新谷 遅刻・欠課入力の入力可・不可切り替え用
					%>

					<!-- 成績 (数値入力、文字入力、成績なし入力により処理を分ける) -->
					<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
						<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><input type="text" <%=w_sIptClassWrk1%>  name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);" onChange="f_GetTotalAvg();"></td>
						
					<% elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
						<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>
							<% if not m_bSeiInpFlg Then %>
								<%=w_sSeiseki%>
							<% else %>
								<input type="button" class="<%=w_cell%>" style="text-align:center;" name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=2 onClick="f_SetSeiseki(<%=i%>);" <%=w_Disabled%>>
							<% end if %>
						</td>
						<input type="hidden" name="hidSeiseki<%=i%>" value="<%=w_sSeiseki%>">
						<input type="hidden" name="hidHyokaFukaKbn<%=i%>" value="<%=m_Rs("HYOKA_FUKA")%>">
					<% else %>
						<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
					<% end if %>
					
					<!-- 評価予定 -->
					<% If m_HyokaDispFlg and m_iKamokuKbn = C_JIK_JUGYO then %>
						<% if m_bSeiInpFlg and (m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU) then %>
							<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>
								<input type="button" name="button<%=i%>" value="<%=w_sHyoka%>" size="2" onClick="f_change(<%=i%>);" class="<%=w_cell%>" style="text-align:center">
								<input type="hidden" name="Hyoka<%=i%>"  value="<%=trim(w_sHyoka)%>">
							</td>
						<% else %>
							<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sHyoka)%></td>
						<% end if %>
					<% end if %>

					<%if m_sGakkoNO = C_NCT_GIFU then
							'遅刻 岐阜は非表示 INS2005/09/30   西村
					%>
					<% ELSE %>
<!--						<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td> -->
						<% '北九州 日々計非表示 2006.11.21 新谷
						   If m_sGakkoNO = C_NCT_KITAKYU Then
						%>
							<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sIptClassWrk%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<% Else %>
							<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sIptClassWrk%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" align="right"  width="55" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sChikaisu)%></td>
						<% End If %>
					<% END IF %>
					

<!--					<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td> -->
					<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sIptClassWrk%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
					
					<% if m_KekkaGaiDispFlg then %>
<!--						<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td> -->
						<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sIptClassWrk%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
					<% end if %>
					
					<%if m_sGakkoNO = C_NCT_GIFU then
							'遅刻 岐阜は非表示 INS2005/09/30   西村
					%>
					<% ELSE %>
						<% '北九州 日々計非表示 2006.11.21 新谷
						   If m_sGakkoNO = C_NCT_KITAKYU Then
						%>
						<% Else %>
							<!-- 日々計 -->
							<td class="<%=w_cell%>" align="right"  width="55" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sKekkasu)%></td>
						<% End If %>
					<% END IF %>

					
					<!-- 評価不能処理 -->
					<% if m_SchoolFlg then %>
						<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>>
							<% if w_DataKbn = C_HYOKA_FUNO or w_DataKbn = C_MIHYOKA or w_DataKbn = 0 then %>
								<input type="checkbox" name="chkHyokaFuno<%=i%>" <%=w_Disabled%> value="3"  <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);">
							<% else %>
								<input type="hidden" name="chkHyokaFuno<%=i%>" value="<%=w_DataKbn%>">
							<% end if %>
						</td>
					<% end if %>
					
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
					<td class="header" align="right" colspan="<%=gf_IIF(m_sGakkoNO = C_NCT_KITAKYU,"5","7")%>" nowrap>
						<FONT COLOR="#FFFFFF"><B>成績合計</B></FONT>
						<input type="text" name="txtTotal" size="5" <%=w_sInputClass%> readonly>
					</td>
					<td class="header" align="center" colspan="6" nowrap>&nbsp;</td>
				</tr>
				
				<tr>
					<td class="header" align="right" colspan="<%=gf_IIF(m_sGakkoNO = C_NCT_KITAKYU,"5","7")%>" nowrap>
						<FONT COLOR="#FFFFFF"><B>　平均点</B></FONT>
						<input type="text" name="txtAvg" size="5" <%=w_sInputClass%> readonly>
					</td>
					<td class="header" align="center" colspan="6" nowrap>&nbsp;</td>
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
						<input type="button" class="button" value="　登　録　" onClick="f_Touroku();">　
					<%End If%>
						<input type="button" class="button" value="キャンセル" onClick="f_Cancel();">
				</td>
			</tr>
		</table>
		</td>
		</tr>
	</table>
	
	
	<input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
	<input type="hidden" name="txtSyoriNendo"     value="<%=g_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
	<input type="hidden" name="KamokuCd"     value="<%=m_sKamokuCd%>">
	<input type="hidden" name="i_Max"        value="<%=i%>">
	<input type="hidden" name="sltShikenKbn" value="<%=m_sSikenKBN%>">
	<input type="hidden" name="txtGakuNo"    value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=m_sGakkaCd%>">
	<input type="hidden" name="txtClassNo"   value="<%=m_sClassNo%>">
	<input type="hidden" name="txtKamokuCd"  value="<%=m_sKamokuCd%>">
	<input type="hidden" name="PasteType"    value="">
	
	<input type="hidden" name="hidSouJyugyou">
	<input type="hidden" name="hidJunJyugyou">
	<input type="hidden" name="hidUpdMode">
	
	<input type="hidden" name="hidKamokuKbn" value="<%=m_iKamokuKbn%>">
	<input type="hidden" name="hidKamokuBunrui" value="<%=m_sKamokuBunrui%>">
	<input type="hidden" name="hidSeisekiInpType" value="<%=m_iSeisekiInpType%>">
	
	<input type="hidden" name="hidKikan" value="<%=m_bSeiInpFlg%>">
	<input type="hidden" name="hidKekkaNyuryokuFlg" value="<%=m_bKekkaNyuryokuFlg%>">
	
	<input type="hidden" name="hidTotal" value="<%=w_lSeiTotal%>">
	<input type="hidden" name="hidGakTotal" value="<%=w_lGakTotal%>">
	<input type="hidden" name="txtUpdDate" value="<%=request("txtUpdDate")%>">
	
	<input type="hidden" name="hidZenkiOnly" value="<%=m_bZenkiOnly%>">
	
	<input type="hidden" name="hidMihyoka" value ="<%=w_DataKbn%>">
	<input type="hidden" name="hidSchoolFlg" value ="<%=m_SchoolFlg%>">
	<input type="hidden" name="hidKekkaGaiDispFlg" value ="<%=m_KekkaGaiDispFlg%>">
	<input type="hidden" name="hidHyokaDispFlg" value ="<%=m_HyokaDispFlg%>">
	
	<input type="hidden" name="hidTableWidth" value ="<%=w_TableWidth%>">
	
	
	<input type="hidden" name="hidFromSei"   value ="<%=m_iNKaishi%>">
	<input type="hidden" name="hidToSei"     value ="<%=m_iNSyuryo%>">
	<input type="hidden" name="hidFromKekka" value ="<%=m_iKekkaKaishi%>">
	<input type="hidden" name="hidToKekka"   value ="<%=m_iKekkaSyuryo%>">
	
	</form>
	</center>
	</body>
	</html>
<%
End sub
%>