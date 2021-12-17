<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0150/sei0150_top.asp
' 機      能: 上ページ 成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:
'           :
' 変      数:
' 引      渡:
'           :
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2002/06/20 shin
' 修      正: 2005/11/22 西村
'             2007.01.29 Takaoka
'             2010/02/17 iwata
'             他クラスの科目を選択している場合、コンボボックスに科目が表示されないので
'             履修科目の開講しているクラスの学科で取得する
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Dim  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	
	Dim m_iNendo             '年度（科目の年度））
	Dim g_iNendo             '年度（処理年度）
	Dim m_sKyokanCd          '教官コード
	Dim m_iSikenKbn			'試験区分
	Dim m_bNoData			'入力科目がないときTrue
	Dim gRs
	
'///////////////////////////メイン処理/////////////////////////////
	
	Call Main()
	
'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()
	Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="成績登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"
	
    On Error Resume Next
    Err.Clear
	
    m_bErrFlg = false
	
    Do
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If
		
		'//値を取得
		call s_SetParam()
		
		'//試験区分  INS 2005/02/10
		if request("sltShikenKbn")  = "" then
			'//最初
			w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_SEISEKI_KIKAN,0)
			
			if w_iRet <> 0 then m_bErrFlg = true : exit do
		else
		    '//リロード時
		    m_iSikenKbn = cint(Request("sltShikenKbn"))

		end if

		'ここまで INS 2005/02/10

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//ログイン教官の担当科目の取得
		if not f_GetSubject() then Exit Do
		
'' 修正 2007.01.29 Takaoka 科目がなくても表示するよう変更 コメントアウト
		If gRs.EOF Then
			m_bNoData = True
		Else
			m_bNoData = False
		End If

'		'科目データなし
'		if gRs.EOF Then
'			Call showWhitePage("担当科目データがありません")
'			response.end
'		End If
'' 修正 2007.01.29 Takaoka 科目がなくても表示するよう変更 コメントアウト　ここまで

		'// ページを表示
		Call showPage()
		
		m_bErrFlg = true
		Exit Do
	Loop
	
	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If not m_bErrFlg Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// 終了処理
	Call gf_closeObject(gRs)
	Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
    m_iNendo    = session("NENDO")
    g_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
	
	'//試験区分
	'If Request("sltShikenKbn")  = "" Then
	'	m_iSikenKbn = C_SIKEN_ZEN_TYU
	'Else
	'    m_iSikenKbn = cint(Request("sltShikenKbn"))
	'End If
	
End Sub


'********************************************************************************
'*  [機能]  ログイン教官の受持教科を取得(年度、教官CD、学期より)
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSubject()
	Dim w_sSQL
    Dim w_sJiki
    Dim w_Rs
    Dim w_sMinNendo 
    
    On Error Resume Next
    Err.Clear
	
    f_GetSubject = false

	'ADD START 2006/10/13 新谷 
	'T11に存在するレコードの入学年度の最小値を取得
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	MIN(T11_NYUNENDO) as NENDO "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & "     T11_GAKUSEKI "
	
	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function
	
	if w_Rs.EOF then exit function
	
	w_sMinNendo = gf_SetNull2String(w_Rs("NENDO"))
	
	Call gf_closeObject(w_Rs)

'response.write "w_sMinNendo = " & w_sMinNendo
	'ADD END 2006/10/13 新谷 

	'選んだ試験によって、開設期間を変える
	Select Case cint(m_iSikenKbn)
		Case cint(C_SIKEN_ZEN_TYU),cint(C_SIKEN_ZEN_KIM) : w_sJiki = C_KAI_ZENKI	'前期中間、前期期末
		Case cint(C_SIKEN_KOU_TYU),cint(C_SIKEN_KOU_KIM) : w_sJiki = C_KAI_KOUKI	'後期中間、後期期末
	End Select

	'通常、留学生代替科目取得
	w_sSQL = ""
	w_sSQL = w_sSQL & " select distinct "
	w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'ADD START 2006/10/13 新谷 
	w_sSQL = w_sSQL & "		,T27_NENDO as NENDO "
	'ADD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & " from"
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & "		,M03_KAMOKU "
	w_sSQL = w_sSQL & "		,T15_RISYU "
	w_sSQL = w_sSQL & "		,M05_CLASS "
	w_sSQL = w_sSQL & " where "
	'UPD START 2006/10/13 新谷 
	'w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "		T27_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M03_KAMOKU_CD "
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI   = " & C_JIK_JUGYO
	w_sSQL = w_sSQL & "	and	T27_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = T15_KAMOKU_CD(+) "
	w_sSQL = w_sSQL & "	and	T15_NYUNENDO = T27_NENDO - T27_GAKUNEN + 1 "
	w_sSQL = w_sSQL & " and T27_NENDO = M05_NENDO "
	w_sSQL = w_sSQL & " and T27_GAKUNEN = M05_GAKUNEN "
	w_sSQL = w_sSQL & " and T27_CLASS = M05_CLASSNO "
	'UPD START 2006/10/13 新谷 
	'w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "	and	M03_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 新谷 
	
	'//後期期末試験じゃないとき
'	if m_iSikenKbn <> C_SIKEN_KOU_KIM then
'		w_sSQL = w_sSQL & " and	(("
'		w_sSQL = w_sSQL & "			T15_KAISETU1 =" & w_sJiki & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU2 =" & w_sJiki & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU3 =" & w_sJiki & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU4 =" & w_sJiki & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU5 =" & w_sJiki
'		w_sSQL = w_sSQL & "		 ) "
'		w_sSQL = w_sSQL & "		 or "
'		w_sSQL = w_sSQL & "		 ("
'		w_sSQL = w_sSQL & "			T15_KAISETU1 =" & C_KAI_TUNEN & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU2 =" & C_KAI_TUNEN & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU3 =" & C_KAI_TUNEN & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU4 =" & C_KAI_TUNEN & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU5 =" & C_KAI_TUNEN
'		w_sSQL = w_sSQL & "		 ) )"
'	else
'		w_sSQL = w_sSQL & " and	("
'		w_sSQL = w_sSQL & "			T15_KAISETU1 <" & C_KAI_NASI & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU2 <" & C_KAI_NASI & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU3 <" & C_KAI_NASI & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU4 <" & C_KAI_NASI & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU5 <" & C_KAI_NASI
'		w_sSQL = w_sSQL & "		 )"
'	end if

	w_sSQL = w_sSQL & " and	(("
	w_sSQL = w_sSQL & "			T15_KAISETU1 =" & w_sJiki & " or "
	w_sSQL = w_sSQL & "			T15_KAISETU2 =" & w_sJiki & " or "
	w_sSQL = w_sSQL & "			T15_KAISETU3 =" & w_sJiki & " or "
	w_sSQL = w_sSQL & "			T15_KAISETU4 =" & w_sJiki & " or "
	w_sSQL = w_sSQL & "			T15_KAISETU5 =" & w_sJiki
	w_sSQL = w_sSQL & "		 ) "

	'前期指定の場合は、通年は条件に入れない
'2006.11.20 前期指定の場合は、通年を入れるようにする
'2007.01.29 やっぱり前期指定のときは通年を入れない'
	If w_sJiki <> C_KAI_ZENKI Then
		w_sSQL = w_sSQL & "		 or "
		w_sSQL = w_sSQL & "		 ("
		w_sSQL = w_sSQL & "			T15_KAISETU1 =" & C_KAI_TUNEN & " or "
		w_sSQL = w_sSQL & "			T15_KAISETU2 =" & C_KAI_TUNEN & " or "
		w_sSQL = w_sSQL & "			T15_KAISETU3 =" & C_KAI_TUNEN & " or "
		w_sSQL = w_sSQL & "			T15_KAISETU4 =" & C_KAI_TUNEN & " or "
		w_sSQL = w_sSQL & "			T15_KAISETU5 =" & C_KAI_TUNEN
		'w_sSQL = w_sSQL & "		 ) )"
		w_sSQL = w_sSQL & "		 )"
	End If
'2007.01.29 やっぱり前期指定のときは通年を入れない'
	w_sSQL = w_sSQL & "		)"
	
	w_sSQL = w_sSQL & vbCrLf & "  UNION ALL "
			
	w_sSQL = w_sSQL & " SELECT DISTINCT "
	w_sSQL = w_sSQL & " 	T27_GAKUNEN AS GAKUNEN"
	w_sSQL = w_sSQL & " 	,T27_CLASS AS CLASS"
	w_sSQL = w_sSQL & " 	,T27_KAMOKU_CD AS KAMOKU_CD "
	w_sSQL = w_sSQL & " 	,T16_KAMOKUMEI AS KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & " 	,M05_CLASSMEI AS CLASS_NAME "
	w_sSQL = w_sSQL & " 	,M05_GAKKA_CD AS GAKKA_CD "
	'ADD START 2006/10/13 新谷 
	w_sSQL = w_sSQL & "		,T27_NENDO as NENDO "
	'ADD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & " FROM"
	w_sSQL = w_sSQL & " 	T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & " 	,T16_RISYU_KOJIN "
	w_sSQL = w_sSQL & " 	,M05_CLASS "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 		T27_NENDO = M05_NENDO "
	w_sSQL = w_sSQL & "    AND T27_GAKUNEN = M05_GAKUNEN "
	w_sSQL = w_sSQL & "    AND T27_CLASS = M05_CLASSNO	"
	w_sSQL = w_sSQL & "    AND T27_KAMOKU_CD = T16_KAMOKU_CD(+)"

'2010/02/17 upd iwata
'他クラスの科目を選択している場合、コンボボックスに科目が表示されないので
'履修科目の開講しているクラスの学科で取得する
'	w_sSQL = w_sSQL & "    AND M05_GAKKA_CD(+) = T16_GAKKA_CD "
	w_sSQL = w_sSQL & "    AND M05_GAKKA_CD(+) = T16_CURRI_GAKKA_CD "

	w_sSQL = w_sSQL & "    AND T16_NENDO(+) = T27_NENDO "
	'UPD START 2006/10/13 新谷 
	'w_sSQL = w_sSQL & "    AND T27_NENDO = " & m_iNendo
	w_sSQL = w_sSQL & "    AND T27_NENDO >= " & w_sMinNendo
	'UPD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
	w_sSQL = w_sSQL & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "
	w_sSQL = w_sSQL & "    AND T16_OKIKAE_FLG >= " & C_TIKAN_KAMOKU_SAKI 

'2010.02.12個人履修に追加した科目の開設時期を見るように修正 伊藤
	w_sSQL = w_sSQL & " and	(("
	w_sSQL = w_sSQL & "			T16_KAISETU =" & w_sJiki
	w_sSQL = w_sSQL & "		 ) "
'前期指定のときは通年を入れない'
	If w_sJiki <> C_KAI_ZENKI Then
		w_sSQL = w_sSQL & "		 or "
		w_sSQL = w_sSQL & "		 ("
		w_sSQL = w_sSQL & "			T16_KAISETU =" & C_KAI_TUNEN
		w_sSQL = w_sSQL & "		 )"
	End If
'前期指定のときは通年を入れない'
	w_sSQL = w_sSQL & "		)"
'2010.02.12個人履修に追加した科目の開設時期を見るように修正 伊藤

	w_sSQL = w_sSQL & " union "
	
	'T27でT27_SEISEKI_INP_FLG=1の人が他の先生に成績登録を許可した際、
	'T26にデータが入るため、以下のSQLを実行する必要がある
	w_sSQL = w_sSQL & " SELECT distinct "
	w_sSQL = w_sSQL & "		T26_GAKUNEN AS GAKUNEN "
	w_sSQL = w_sSQL & "		,T26_CLASS AS CLASS "
	w_sSQL = w_sSQL & "		,T26_KAMOKU AS KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,0 as KAMOKU_KBN "
	w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'ADD START 2006/10/13 新谷 
	w_sSQL = w_sSQL & "		,T26_NENDO as NENDO "
	'ADD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & "		T26_SIKEN_JIKANWARI "
	w_sSQL = w_sSQL & "		,M03_KAMOKU "
	w_sSQL = w_sSQL & "		,T15_RISYU "
	w_sSQL = w_sSQL & "		,M05_CLASS "
	w_sSQL = w_sSQL & " WHERE "
	'UPD START 2006/10/13 新谷 
	'w_sSQL = w_sSQL & "		 T26_NENDO = " & cint(m_iNendo)
	w_sSQL = w_sSQL & "		 T26_NENDO >= " & w_sMinNendo
	'UPD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & "	and ("
	w_sSQL = w_sSQL & "		T26_JISSI_KYOKAN    ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN1 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN2 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN3 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN4 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN5 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		)"
	w_sSQL = w_sSQL & "	and	T26_KAMOKU = M03_KAMOKU_CD "	
	w_sSQL = w_sSQL & "	and T26_KAMOKU = T15_KAMOKU_CD(+) "
	w_sSQL = w_sSQL & "	and T15_NYUNENDO(+) = T26_NENDO - T26_GAKUNEN + 1 "
	w_sSQL = w_sSQL & "	and T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "' "
	w_sSQL = w_sSQL & "	and	T26_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'UPD START 2006/10/13 新谷 
	'w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "	and	M03_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & " and T26_NENDO = M05_NENDO "
	w_sSQL = w_sSQL & " and T26_GAKUNEN = M05_GAKUNEN "
	w_sSQL = w_sSQL & " and T26_CLASS = M05_CLASSNO "
	
	'//後期期末試験じゃないとき
	if cint(m_iSikenKbn) <> cint(C_SIKEN_KOU_KIM) then
		w_sSQL = w_sSQL & "	and T26_SIKEN_KBN =" & m_iSikenKbn
	end if
	
	w_sSQL = w_sSQL & " union "
	
	'特別活動取得
	w_sSQL = w_sSQL & " select distinct "
	w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M41_MEISYO as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'ADD START 2006/10/13 新谷 
	w_sSQL = w_sSQL & "		,T27_NENDO as NENDO "
	'ADD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & "		,M41_TOKUKATU "
	w_sSQL = w_sSQL & "		,M05_CLASS "
	w_sSQL = w_sSQL & " where "
	'UPD START 2006/10/13 新谷 
	'w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "		T27_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M41_TOKUKATU_CD "
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI = " & C_JIK_TOKUBETU
	w_sSQL = w_sSQL & "	and	T27_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'UPD START 2006/10/13 新谷 
	'w_sSQL = w_sSQL & "	and	M41_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "	and	M41_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 新谷 
	w_sSQL = w_sSQL & " and T27_NENDO = M05_NENDO "
	w_sSQL = w_sSQL & " and T27_GAKUNEN = M05_GAKUNEN "
	w_sSQL = w_sSQL & " and T27_CLASS = M05_CLASSNO "
	'UPD START 2006/10/13 新谷 
	'w_sSQL = w_sSQL & " order by GAKUNEN,CLASS,KAMOKU_KBN "
	w_sSQL = w_sSQL & " order by NENDO DESC,GAKUNEN,CLASS,KAMOKU_KBN "
	'UPD END 2006/10/13 新谷 
	
	'response.write w_sSQL & "<BR>"
	
'response.write w_sSQL
'response.end		
	If gf_GetRecordset(gRs,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit function
	End If
	
	f_GetSubject = true
    
End Function

'********************************************************************************
'*  [機能]  履修データから更新日を取得する。
'*  [引数]  
'*			p_iNendo - 処理年度
'*			p_iGakunen - 学年
'*			p_sGakkaCd - 学科コード
'*			p_sKamokuCd - 科目コード
'*  [戻値]  更新日付
'*  [説明]  
'********************************************************************************
Function f_GetUpdDate(p_iNendo,p_iGakunen,p_sGakkaCd,p_sKamokuCd,p_KamokuKbn)
	
	Dim w_sSQL
	Dim w_Rs
	Dim w_FieldName
	Dim w_Table,w_TableName,w_KamokuName
	
	On Error Resume Next
	Err.Clear
	
	f_GetUpdDate = ""
	
	if p_KamokuKbn = C_JIK_JUGYO then
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	else
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_KamokuName = "T34_TOKUKATU_CD"
	end if
	
	'UPP 2005/11/22西村 CINT()関数をつける
	select case CINT(m_iSikenKbn)
		case CINT(C_SIKEN_ZEN_TYU) : w_FieldName = w_Table & "_KOUSINBI_TYUKAN_Z"
		case CINT(C_SIKEN_ZEN_KIM) : w_FieldName = w_Table & "_KOUSINBI_KIMATU_Z"
		case CINT(C_SIKEN_KOU_TYU) : w_FieldName = w_Table & "_KOUSINBI_TYUKAN_K"
		case CINT(C_SIKEN_KOU_KIM) : w_FieldName = w_Table & "_KOUSINBI_KIMATU_K"
	end select
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	Max(" & w_FieldName & ") as UPD_DATE "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & 		w_TableName
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	" & w_Table & "_NENDO        =  " & p_iNendo
	w_sSQL = w_sSQL & " And " & w_Table & "_HAITOGAKUNEN =  " & p_iGakunen
	w_sSQL = w_sSQL & " And " & w_Table & "_GAKKA_CD     = '" & p_sGakkaCd & "'"
	w_sSQL = w_sSQL & " And " & w_KamokuName & "    = '" & p_sKamokuCd & "'"
	w_sSQL = w_sSQL & " And " & w_FieldName & " is not NULL "
	
	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function
	
	if w_Rs.EOF then exit function
	
	f_GetUpdDate = gf_SetNull2String(w_Rs("UPD_DATE"))
	
	Call gf_closeObject(w_Rs)
	
End Function

'********************************************************************************
'*  HTMLを出力
'********************************************************************************
Sub showPage()
	Dim w_TukuName
	Dim w_SubjectDisp
	Dim w_SubjectValue
	Dim w_sWhere
	
	Dim w_iGakunen_s
	Dim w_sGakkaCd_s
	Dim w_sKamokuCd_s
	
	On Error Resume Next
    Err.Clear
	
%>
	<html>
	<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//  [機能]  試験が変更されたとき、再表示する
	//************************************************************
	function f_ReLoadMyPage(){
		document.frm.action="sei0150_top.asp";
		document.frm.target="topFrame";
		document.frm.submit();
	}
	
	//************************************************************
	//  [機能]  表示ボタンクリック時の処理
	//************************************************************
	function f_Search(){
		// 選択されたコンボの値をｾｯﾄ
		f_SetData();
		
	    document.frm.action="sei0150_bottom.asp";
	    document.frm.target="main";
	    document.frm.submit();
	}
	
	//************************************************************
	//  [機能]  表示ボタンクリック時に選択されたデータをｾｯﾄ
	//************************************************************
	function f_SetData(){
		//データ取得
		var vl = document.frm.sltSubject.value.split('#@#');
		
		//選択されたデータをｾｯﾄ(学年、クラス、科目CDを取得)
<%		'UPD START 2006/10/13 新谷 パラメータに年を追加した為
'		document.frm.txtGakuNo.value=vl[0];
'		document.frm.txtClassNo.value=vl[1];
'		document.frm.txtKamokuCd.value=vl[2];
'		document.frm.txtGakkaCd.value=vl[3];
'		document.frm.txtUpdDate.value=vl[4];
'		document.frm.SYUBETU.value=vl[5];
'		document.frm.hidKamokuKbn.value=vl[6];
%>
		document.frm.txtNendo.value=vl[0];
		document.frm.txtGakuNo.value=vl[1];
		document.frm.txtClassNo.value=vl[2];
		document.frm.txtKamokuCd.value=vl[3];
		document.frm.txtGakkaCd.value=vl[4];
		document.frm.txtUpdDate.value=vl[5];
		document.frm.SYUBETU.value=vl[6];
		document.frm.hidKamokuKbn.value=vl[7];
<%		'UPD END 2006/10/13 新谷 パラメータに年を追加した為 %>
	}
	
	//************************************************************
	//  [機能]  更新日のセット
	//************************************************************
	function f_SetUpdDate(){
		var vl = document.frm.sltSubject.value.split('#@#');
<%		'UPD 2006/10/13 新谷 パラメータに年を追加した為
'		document.frm.txtUpdDate.value=vl[4]; %>
		document.frm.txtUpdDate.value=vl[5]; 
	}
	
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>
	
    <body LANGUAGE="javascript" onload="f_SetUpdDate();">
	
	<center>
	<form name="frm" METHOD="post">
	
	<% call gs_title(" 成績登録 "," 登　録 ") %>
	<br>
	
	<table border="0">
		<tr><td valign="bottom">
			
			<table border="0" width="100%">
				<tr><td class="search">
					
					<table border="0">
						<tr valign="middle">
							<td align="left" nowrap>試験区分</td>
							<td align="left" colspan="3">
							<% 
								w_sWhere = " M01_NENDO = " & m_iNendo
								w_sWhere = w_sWhere & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
								w_sWhere = w_sWhere & " AND M01_SYOBUNRUI_CD < " & cint(C_SIKEN_JITURYOKU)
								
								Call gf_ComboSet("sltShikenKbn",C_CBO_M01_KUBUN,w_sWhere," onchange = 'f_ReLoadMyPage();' style='width:140px;'",false,m_iSikenKbn)
							%>
							</td>
							<td>&nbsp;</td>
							
							<td align="left" nowrap>科目</td>
							<td align="left">
								<% if not gRs.EOF then %>
									<select name="sltSubject" onChange="f_SetUpdDate();">
									<% 

										do until gRs.EOF
											
											'科目コンボ表示部分生成
											w_SubjectDisp =""
											'ADD START 2006/10/13 新谷
											w_SubjectDisp = w_SubjectDisp & gRs("NENDO") & "　"
											'ADD END 2006/10/13 新谷
											w_SubjectDisp = w_SubjectDisp & gRs("GAKUNEN") & "年　"
											w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "　"
											w_SubjectDisp = w_SubjectDisp & gRs("KAMOKU_NAME") & "　"
											
											w_TukuName = ""
											
											if cint(gf_SetNull2Zero(gRs("KAMOKU_KBN"))) = 1 then
												w_TukuName = "TOKU"
											else
												w_TukuName = "TUJO"
											end if
											
											'科目コンボVALUE部分生成
											w_SubjectValue = ""
											'ADD START 2006/10/13 新谷
											w_SubjectValue = w_SubjectValue & gRs("NENDO")   & "#@#"
											'ADD END 2006/10/13 新谷
											w_SubjectValue = w_SubjectValue & gRs("GAKUNEN")   & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("CLASS")     & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("KAMOKU_CD") & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("GAKKA_CD")  & "#@#"
											'UPD START 2006/10/13 新谷
											'w_SubjectValue = w_SubjectValue & f_GetUpdDate(m_iNendo,gRs("GAKUNEN"),gRs("GAKKA_CD"),gRs("KAMOKU_CD"),cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))) & "#@#"
											w_SubjectValue = w_SubjectValue & f_GetUpdDate(gRs("NENDO"),gRs("GAKUNEN"),gRs("GAKKA_CD"),gRs("KAMOKU_CD"),cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))) & "#@#"
											'UPD END 2006/10/13 新谷
											w_SubjectValue = w_SubjectValue & w_TukuName  & "#@#"
											w_SubjectValue = w_SubjectValue & cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))
											
									%>
										<option value="<%=w_SubjectValue%>"><%=w_SubjectDisp%>
									<% 
											gRs.movenext
										loop 
									%>
									</select>
								<% 	'' UPD 2007/01/29 EOF のときのメッセージを追加
									else %>
									：この試験区分の入力科目はありません
								<% end if %>
							</td>
	                    </tr>
						
						<tr>
							<td align="left" nowrap>最終更新日</td>
							<td align="left" colspan="3" nowrap>
								<input type="text" name="txtUpdDate" value="" onFocus="blur();" readonly style="BACKGROUND-COLOR: #E4E4ED">
							</td>
							
							<td colspan="7" align="right">
							<% ' UPD 2007/01/29 EOF のときの処理を追加
								If Not m_bNoData Then %>
								<input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
							<% 	Else %>
							<% 	end if %>
							</td>
						</tr>
					</table>
					
				</td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
	
	<input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
	<input type="hidden" name="txtSyoriNendo"     value="<%=g_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtGakuNo"    value="<%=w_iGakunen_s%>">
	<input type="hidden" name="txtClassNo"   value="">
	<input type="hidden" name="txtKamokuCd"  value="<%=w_sKamokuCd_s%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=w_sGakkaCd_s%>">
	<input type="hidden" name="SYUBETU"      value="">
	<input type="hidden" name="hidKamokuKbn" value="">	
	</form>
	</center>
	</body>
	</html>
<%
End Sub

'********************************************************************************
'*	空白HTMLを出力
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
	<html>
	<head>
	<title>成績登録</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>
	
	<body LANGUAGE="javascript">
	<form name="frm" mothod="post">
	
	<center>
	<br><br><br>
		<span class="msg"><%=Server.HTMLEncode(p_Msg)%></span>
	</center>
	
	<input type="hidden" name="txtMsg" value="<%=Server.HTMLEncode(p_Msg)%>">
	</form>
	</body>
	</html>
<%
End Sub
%>