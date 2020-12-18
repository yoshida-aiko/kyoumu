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
' 変      更: 2010/05/20 岩田　高知高専　混合クラス、一般、指定科目を学科表示にする
' 変      更: 2016/06/14 清本　高知高専　専門科目をクラス表示にする
' 変      更: 2017/12/12 西村　高知高専　混合クラス、一般、指定科目を学科表示にするで140050生物は除外する
' 変      更: 2018/05/08 清本　高知高専　1学科複数コース対応
' 変      更: 2018/06/18 清本　高知高専　個人履修追加科目も開設時期によって科目を制御する
' 変      更: 2018/06/24 藤林　高知高専　科目コンボは、クラス、学科、科目ごとに表示するように変更。VIEWを使用するように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Dim  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

	Dim m_iNendo             '年度
	Dim m_sKyokanCd          '教官コード
	Dim m_iSikenKbn			'試験区分

	Dim gDisabled

	Dim gRs

	Public m_sGakkoNO       '学校番号

	Dim m_iGakunen
	Dim m_iClass
	Dim m_iKongo

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

		'//試験区分
		if request("sltShikenKbn")  = "" then
			'//最初
			w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_SEISEKI_KIKAN,0)

			if w_iRet <> 0 then m_bErrFlg = true : exit do
		else
		    '//リロード時
		    m_iSikenKbn = cint(Request("sltShikenKbn"))
		end if

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'学校番号の取得
		if Not gf_GetGakkoNO(m_sGakkoNO) then Exit Do

		'//ログイン教官の担当科目の取得
		if not f_GetSubject() then Exit Do

		'科目データなし
		if gRs.EOF Then
			gDisabled = "disabled"
		'	Call showWhitePage("担当科目データがありません")
		'	response.end
		End If

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

	gDisabled = ""

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")

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
    Dim w_iCnt

    On Error Resume Next
    Err.Clear

    f_GetSubject = false

	'選んだ試験によって、開設期間を変える
	Select Case cint(m_iSikenKbn)
		Case C_SIKEN_ZEN_TYU : w_sJiki = C_KAI_ZENKI	'前期中間
		case C_SIKEN_ZEN_KIM : w_sJiki = C_KAI_ZENKI	'前期期末
		Case C_SIKEN_KOU_TYU : w_sJiki = C_KAI_KOUKI	'後期中間
        case C_SIKEN_KOU_KIM : w_sJiki = C_KAI_KOUKI	'後期期末
	End Select

	'--2019/06/24 Upd Fujibayashi(他の画面との整合性を保つため、VIEWを使用する)
	''通常、留学生代替科目取得
	'w_sSQL = ""
	'w_sSQL = w_sSQL & " select distinct "
	'w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	'w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	'w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	'w_sSQL = w_sSQL & "		,M03_KAMOKU_KBN as KAMOKU_KBN_IS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	'w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	'w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'w_sSQL = w_sSQL & " from"
	'w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	'w_sSQL = w_sSQL & "		,M03_KAMOKU "
	'w_sSQL = w_sSQL & "		,T15_RISYU "
	'w_sSQL = w_sSQL & "		,M05_CLASS "
	'w_sSQL = w_sSQL & " where "
	'w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iNendo)
	'w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M03_KAMOKU_CD "
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI   = " & C_JIK_JUGYO
	'w_sSQL = w_sSQL & "	and	T27_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = T15_KAMOKU_CD(+) "
	'w_sSQL = w_sSQL & "	and	T15_NYUNENDO = T27_NENDO - T27_GAKUNEN + 1 "
	'w_sSQL = w_sSQL & "	and	T15_GAKKA_CD = M05_GAKKA_CD"	'2003.01.29
	'w_sSQL = w_sSQL & "	and T15_COURSE_CD IN ('0', CASE WHEN M05_COURSE_CD IS NOT NULL THEN M05_COURSE_CD ELSE T15_COURSE_CD END) "	'2018.05.08 Add Kiyomoto
	'w_sSQL = w_sSQL & " and T27_NENDO = M05_NENDO "
	'w_sSQL = w_sSQL & " and T27_GAKUNEN = M05_GAKUNEN "
	'w_sSQL = w_sSQL & " and T27_CLASS = M05_CLASSNO "
	'w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iNendo)
    '
	''八代高専の場合、後期中間の時前期開設の科目も表示する
    'if m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then
    '
	'	'//後期中間、後期期末試験じゃないとき
	'	if m_iSikenKbn = C_SIKEN_KOU_KIM Or m_iSikenKbn = C_SIKEN_KOU_TYU then
	'		'w_sSQL = w_sSQL & " and	("
	'		'w_sSQL = w_sSQL & "			T15_KAISETU1 <" & C_KAI_NASI & " or "
	'		'w_sSQL = w_sSQL & "			T15_KAISETU2 <" & C_KAI_NASI & " or "
	'		'w_sSQL = w_sSQL & "			T15_KAISETU3 <" & C_KAI_NASI & " or "
	'		'w_sSQL = w_sSQL & "			T15_KAISETU4 <" & C_KAI_NASI & " or "
	'		'w_sSQL = w_sSQL & "			T15_KAISETU5 <" & C_KAI_NASI
	'		'w_sSQL = w_sSQL & "		 )"
    '
	'		'学年に合わせた開設時期を条件にする 2003.01.29
	'		w_sSQL = w_sSQL & " and DECODE(T27_GAKUNEN "
	'		w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'		w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'		w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'		w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'		w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'		w_sSQL = w_sSQL & " ) < " & C_KAI_NASI
    '
	'	else
	'		w_sSQL = w_sSQL & " and	("
	'		for  w_iCnt = 1 to 5
	'		     w_sSQL = w_sSQL & "((T15_KAISETU" & w_iCnt & " = " & w_sJiki & "  "
	'             w_sSQL = w_sSQL & "or  T15_KAISETU" & w_iCnt & " = " &  C_KAI_TUNEN & ")  "
	'             w_sSQL = w_sSQL & " and  T27_GAKUNEN = " & w_iCnt  & ")  "
    '
	'		     if w_iCnt <> 5 then
	'		     	  w_sSQL = w_sSQL & " or  "
	'                     end if
	'		next
	'		     w_sSQL = w_sSQL & " ) "
    '
	'	end if
    '
	''その他の学校は学年末試験の時だけ前期開設の科目を表示する
	'else
	'	'新居浜高専の場合、後期期末は前期科目は表示しない
	'    if m_sGakkoNO = cstr(C_NCT_NIIHAMA) or m_sGakkoNO = cstr(C_NCT_KURUME) then
	'			w_sSQL = w_sSQL & " and	("
	'			for  w_iCnt = 1 to 5
	'			     w_sSQL = w_sSQL & "((T15_KAISETU" & w_iCnt & " = " & w_sJiki & "  "
	'                 w_sSQL = w_sSQL & "or  T15_KAISETU" & w_iCnt & " = " &  C_KAI_TUNEN & ")  "
	'                 w_sSQL = w_sSQL & " and  T27_GAKUNEN = " & w_iCnt  & ")  "
    '
	'			     if w_iCnt <> 5 then
	'			     	  w_sSQL = w_sSQL & " or  "
	'	              end if
	'			next
    '
	'			w_sSQL = w_sSQL & " ) "
    '    else
	'		'//後期期末試験じゃないとき
	'		if m_iSikenKbn <> C_SIKEN_KOU_KIM then
	'			w_sSQL = w_sSQL & " and	("
	'			for  w_iCnt = 1 to 5
	'			     w_sSQL = w_sSQL & "((T15_KAISETU" & w_iCnt & " = " & w_sJiki & "  "
	'                 w_sSQL = w_sSQL & "or  T15_KAISETU" & w_iCnt & " = " &  C_KAI_TUNEN & ")  "
	'                 w_sSQL = w_sSQL & " and  T27_GAKUNEN = " & w_iCnt  & ")  "
    '
	'			     if w_iCnt <> 5 then
	'			     	  w_sSQL = w_sSQL & " or  "
	'	              end if
	'			next
    '
	'			w_sSQL = w_sSQL & " ) "
    '
	'		else
    '
	'			'学年に合わせた開設時期を条件にする 2003.01.29
	'			w_sSQL = w_sSQL & " and DECODE(T27_GAKUNEN "
	'			w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'			w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'			w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'			w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'			w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'			w_sSQL = w_sSQL & " ) < " & C_KAI_NASI
    '
	'		end if
	'    end if
	'end if
    '
    '
	'w_sSQL = w_sSQL & vbCrLf & "  UNION ALL "
    '
	'w_sSQL = w_sSQL & " SELECT DISTINCT "
	'w_sSQL = w_sSQL & " 	T27_GAKUNEN AS GAKUNEN"
	'w_sSQL = w_sSQL & " 	,T27_CLASS AS CLASS"
	'w_sSQL = w_sSQL & " 	,T27_KAMOKU_CD AS KAMOKU_CD "
	'w_sSQL = w_sSQL & " 	,T16_KAMOKUMEI AS KAMOKU_NAME "
	'w_sSQL = w_sSQL & "		,T16_KAMOKU_KBN as KAMOKU_KBN_IS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	'w_sSQL = w_sSQL & " 	,M05_CLASSMEI AS CLASS_NAME "
	'w_sSQL = w_sSQL & " 	,M05_GAKKA_CD AS GAKKA_CD "
	'w_sSQL = w_sSQL & " FROM"
	'w_sSQL = w_sSQL & " 	T27_TANTO_KYOKAN "
	'w_sSQL = w_sSQL & " 	,T16_RISYU_KOJIN "
	'w_sSQL = w_sSQL & " 	,M05_CLASS "
	'w_sSQL = w_sSQL & " WHERE "
	'w_sSQL = w_sSQL & " 		T27_NENDO = M05_NENDO "
	'w_sSQL = w_sSQL & "    AND T27_GAKUNEN = M05_GAKUNEN "
	'w_sSQL = w_sSQL & "    AND T27_GAKUNEN = T16_HAITOGAKUNEN "
	'w_sSQL = w_sSQL & "    AND T27_CLASS = M05_CLASSNO	"
	'w_sSQL = w_sSQL & "    AND T27_KAMOKU_CD = T16_KAMOKU_CD(+)"
	'w_sSQL = w_sSQL & "    AND M05_GAKKA_CD(+) = T16_GAKKA_CD "
	'w_sSQL = w_sSQL & "    AND T16_NENDO(+) = T27_NENDO "
	'w_sSQL = w_sSQL & "    AND T27_NENDO = " & m_iNendo
	'w_sSQL = w_sSQL & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
	'w_sSQL = w_sSQL & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "
	'w_sSQL = w_sSQL & "    AND T16_OKIKAE_FLG >= " & C_TIKAN_KAMOKU_SAKI
    '
	''INS 2008/09/11
	''新居浜高専の場合、前期は「前期・通年」後期は「後期・通年」とする
	''2018.06.18 Add 高知高専の場合も同様
    'if m_sGakkoNO = cstr(C_NCT_NIIHAMA) or m_sGakkoNO = cstr(C_NCT_KOCHI) then
	'	w_sSQL = w_sSQL & "    AND ((T16_KAISETU = " & w_sJiki & "  "
	'	w_sSQL = w_sSQL & "    or  T16_KAISETU = " &  C_KAI_TUNEN & "))  "
	'end if
	''INS END 2008/09/11
    '
    '
    '
	'w_sSQL = w_sSQL & " union "
    '
	''T27でT27_SEISEKI_INP_FLG=1の人が他の先生に成績登録を許可した際、
	''T26にデータが入るため、以下のSQLを実行する必要がある
	'w_sSQL = w_sSQL & " SELECT distinct "
	'w_sSQL = w_sSQL & "		T26_GAKUNEN AS GAKUNEN "
	'w_sSQL = w_sSQL & "		,T26_CLASS AS CLASS "
	'w_sSQL = w_sSQL & "		,T26_KAMOKU AS KAMOKU_CD "
	'w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	'w_sSQL = w_sSQL & "		,M03_KAMOKU_KBN as KAMOKU_KBN_IS "
	'w_sSQL = w_sSQL & "		,0 as KAMOKU_KBN "
	'w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	'w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'w_sSQL = w_sSQL & " FROM "
	'w_sSQL = w_sSQL & "		T26_SIKEN_JIKANWARI "
	'w_sSQL = w_sSQL & "		,M03_KAMOKU "
	'w_sSQL = w_sSQL & "		,T15_RISYU "
	'w_sSQL = w_sSQL & "		,M05_CLASS "
	'w_sSQL = w_sSQL & " WHERE "
	'w_sSQL = w_sSQL & "		 T26_NENDO = " & cint(m_iNendo)
	'w_sSQL = w_sSQL & "	and ("
	'w_sSQL = w_sSQL & "		T26_JISSI_KYOKAN    ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN1 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN2 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN3 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN4 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN5 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		)"
	'w_sSQL = w_sSQL & "	and	T26_KAMOKU = M03_KAMOKU_CD "
	'w_sSQL = w_sSQL & "	and T26_KAMOKU = T15_KAMOKU_CD(+) "
	'w_sSQL = w_sSQL & "	and T15_NYUNENDO(+) = T26_NENDO - T26_GAKUNEN + 1 "
	'w_sSQL = w_sSQL & "	and	T15_GAKKA_CD = M05_GAKKA_CD"	'2003.01.29
	'w_sSQL = w_sSQL & "	and T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "' "
	'w_sSQL = w_sSQL & "	and	T26_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iNendo)
	'w_sSQL = w_sSQL & " and T26_NENDO = M05_NENDO "
	'w_sSQL = w_sSQL & " and T26_GAKUNEN = M05_GAKUNEN "
	'w_sSQL = w_sSQL & " and T26_CLASS = M05_CLASSNO "
    '
	''八代高専の場合、後期中間の時前期開設の科目も表示する
    'if m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then
	'	'//後期中間、後期期末試験じゃないとき
	'	if m_iSikenKbn = C_SIKEN_KOU_KIM Or m_iSikenKbn = C_SIKEN_KOU_TYU then
	'	else
	'		w_sSQL = w_sSQL & "	and T26_SIKEN_KBN =" & m_iSikenKbn
	'	end if
	'else
	''新居浜高専の場合、後期期末は前期科目は表示しない
	'    if m_sGakkoNO = cstr(C_NCT_NIIHAMA) or m_sGakkoNO = cstr(C_NCT_KURUME) then
	'			w_sSQL = w_sSQL & "	and T26_SIKEN_KBN =" & m_iSikenKbn
	'    else
	'		'//後期期末試験じゃないとき
	'		if m_iSikenKbn <> C_SIKEN_KOU_KIM then
	'			w_sSQL = w_sSQL & "	and T26_SIKEN_KBN =" & m_iSikenKbn
	'		end if
	'	end if
    'end if
    '
	''開設時期条件の追加 2003.01.29
	''八代高専の場合、後期中間の時前期開設の科目も表示する
    'if m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then
    '
	'	'//後期中間、後期期末試験じゃないとき
	'	if m_iSikenKbn = C_SIKEN_KOU_KIM Or m_iSikenKbn = C_SIKEN_KOU_TYU then
    '
	'		'学年に合わせた開設時期を条件にする
	'		w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'		w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'		w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'		w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'		w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'		w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'		w_sSQL = w_sSQL & " ) < " & C_KAI_NASI
    '
	'	else
    '
	'		'学年に合わせた開設時期を条件にする
	'		w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'		w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'		w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'		w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'		w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'		w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'		w_sSQL = w_sSQL & ") IN (" & w_sJiki & "," & C_KAI_TUNEN & ")"
    '
    '
	'	end if
    '
	''その他の学校は学年末試験の時だけ前期開設の科目を表示する
	'else
	''新居浜高専の場合、後期期末は前期科目は表示しない
	'    if m_sGakkoNO = cstr(C_NCT_NIIHAMA) or m_sGakkoNO = cstr(C_NCT_KURUME) then
    '
	'			'学年に合わせた開設時期を条件にする
	'			w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'			w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'			w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'			w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'			w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'			w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'			w_sSQL = w_sSQL & ") IN (" & w_sJiki & "," & C_KAI_TUNEN & ")"
    '    else
	'		'//後期期末試験じゃないとき
	'		if m_iSikenKbn <> C_SIKEN_KOU_KIM then
    '
	'			'学年に合わせた開設時期を条件にする
	'			w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'			w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'			w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'			w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'			w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'			w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'			w_sSQL = w_sSQL & ") IN (" & w_sJiki & "," & C_KAI_TUNEN & ")"
    '
	'		else
    '
	'			'学年に合わせた開設時期を条件にする
	'			w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'			w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'			w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'			w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'			w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'			w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'			w_sSQL = w_sSQL & ") < " & C_KAI_NASI
	'		end if
    '    end if
	'end if
    '
    '
    '
    '
	'w_sSQL = w_sSQL & " union "
    '
	''特別活動取得
	'w_sSQL = w_sSQL & " select distinct "
	'w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	'w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	'w_sSQL = w_sSQL & "		,M41_MEISYO as KAMOKU_NAME "
	'w_sSQL = w_sSQL & "		,0 as KAMOKU_KBN_IS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	'w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	'w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'w_sSQL = w_sSQL & " from "
	'w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	'w_sSQL = w_sSQL & "		,M41_TOKUKATU "
	'w_sSQL = w_sSQL & "		,M05_CLASS "
	'w_sSQL = w_sSQL & " where "
	'w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iNendo)
	'w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M41_TOKUKATU_CD "
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI = " & C_JIK_TOKUBETU
	'w_sSQL = w_sSQL & "	and	T27_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'w_sSQL = w_sSQL & "	and	M41_NENDO =" & cint(m_iNendo)
	'w_sSQL = w_sSQL & " and T27_NENDO = M05_NENDO "
	'w_sSQL = w_sSQL & " and T27_GAKUNEN = M05_GAKUNEN "
	'w_sSQL = w_sSQL & " and T27_CLASS = M05_CLASSNO "
    '
	'w_sSQL = w_sSQL & " order by GAKUNEN,CLASS,KAMOKU_KBN "


	w_sSQL = "SELECT GAKUNEN"
	w_sSQL = w_sSQL & " ,CLASS"
	w_sSQL = w_sSQL & " ,KAMOKU_CD"
	w_sSQL = w_sSQL & " ,KAMOKUMEI AS KAMOKU_NAME"
	w_sSQL = w_sSQL & " ,KAMOKU_KBN_IS"
	w_sSQL = w_sSQL & " ,KAMOKU_KBN"
	w_sSQL = w_sSQL & " ,DECODE(MAIN_GAKKA , 1 , M05_CLASSMEI , M02_GAKKARYAKSYO) AS CLASS_NAME"
	w_sSQL = w_sSQL & " ,GAKKA_CD"
	w_sSQL = w_sSQL & " FROM VWEB_RISYU"
	w_sSQL = w_sSQL & " WHERE NENDO = " & cint(m_iNendo)
	w_sSQL = w_sSQL & " AND (KAISETU IN (" & w_sJiki & " ," & C_KAI_TUNEN & ")"
	w_sSQL = w_sSQL & "     OR SIKEN_KBN = " & cint(m_iSikenKbn)
	w_sSQL = w_sSQL & "     )"
	w_sSQL = w_sSQL & " AND KYOKAN_CD = '" & m_sKyokanCd & "' "
	w_sSQL = w_sSQL & " ORDER BY GAKUNEN"
	w_sSQL = w_sSQL & "         ,CLASS"
	w_sSQL = w_sSQL & "         ,KAMOKU_KBN"
	w_sSQL = w_sSQL & "         ,KAMOKU_CD"
	w_sSQL = w_sSQL & "         ,MAIN_GAKKA DESC"
	'--2019/06/24 Upd End

'response.write "w_sSQL = " & w_sSQL
'response.end

	If gf_GetRecordset(gRs,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit function
	End If

	f_GetSubject = true

End Function

'********************************************************************************
'*  [機能]  混合クラスかどうかを調べる
'*  [説明]
'********************************************************************************
Function f_GetKongoClass(p_iGakunen,p_iClass,p_iKongo)
	Dim w_sSQL
	Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_GetKongoClass = false

    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M05_SYUBETU "
    w_sSQL = w_sSQL & "FROM M05_CLASS "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_GAKUNEN = " & p_iGakunen & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_CLASSNO = " & p_iClass & " "

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

	'Response.Write w_sSQL

	'//戻り値ｾｯﾄ
	If w_Rs.EOF = False Then
		p_iKongo = Cint(w_Rs("M05_SYUBETU"))
	End If

	f_GetKongoClass = true

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  学科略名を取得(表示用)
'*  [引数]  なし
'*  [戻値]  gf_GetGakkaNm:学科名
'*  [説明]
'********************************************************************************
Function f_GetGakkaNm(p_iNendo,p_sCD)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetGakkaNm = ""
	w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKARYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKA "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M02_GAKKA_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M02_NENDO = " & p_iNendo & " "

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M02_GAKKARYAKSYO")
        End If

        Exit Do
    Loop

	'//戻り値ｾｯﾄ
    f_GetGakkaNm = w_sName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

    Err.Clear

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

	select case m_iSikenKbn
		case C_SIKEN_ZEN_TYU : w_FieldName = w_Table & "_KOUSINBI_TYUKAN_Z"
		case C_SIKEN_ZEN_KIM : w_FieldName = w_Table & "_KOUSINBI_KIMATU_Z"
		case C_SIKEN_KOU_TYU : w_FieldName = w_Table & "_KOUSINBI_TYUKAN_K"
		case C_SIKEN_KOU_KIM : w_FieldName = w_Table & "_KOUSINBI_KIMATU_K"
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
		document.frm.txtGakuNo.value=vl[0];
		document.frm.txtClassNo.value=vl[1];
		document.frm.txtKamokuCd.value=vl[2];
		document.frm.txtGakkaCd.value=vl[3];
		document.frm.txtUpdDate.value=vl[4];
		document.frm.SYUBETU.value=vl[5];
		document.frm.hidKamokuKbn.value=vl[6];
		document.frm.txtClassMei.value=vl[7];		//2019/06/24 Add Fujibayashi
		document.frm.txtKamokuMei.value=vl[8];		//2019/06/24 Add Fujibayashi
	}

	//************************************************************
	//  [機能]  更新日のセット
	//************************************************************
	function f_SetUpdDate(){
		<% if gDisabled = "" then %>
			var vl = document.frm.sltSubject.value.split('#@#');
			document.frm.txtUpdDate.value=vl[4];
		<% end if %>
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

								<select name="sltSubject" onChange="f_SetUpdDate();" <%=gDisabled%>>
								<%
									if gRs.EOF Then
										'//科目データなし
										response.write "<option value=''>担当科目データがありません"
									else
										'//科目データあり
										do until gRs.EOF

											'科目コンボ表示部分生成
											w_SubjectDisp =""
											w_SubjectDisp = w_SubjectDisp & gRs("GAKUNEN") & "年　"

											'混合クラスの専門科目の場合学科名称を表示する。

											If Not f_GetKongoClass(gRs("GAKUNEN"),gRs("CLASS"),m_iKongou) Then
													m_iKongou = C_CLASS_GAKKA
											End If

											If cint(m_iKongou) = cint(C_CLASS_KONGO) Then
												If cint(gRs("KAMOKU_KBN_IS")) = cint(C_KAMOKU_SENMON) Then

													If m_sGakkoNO <> cstr(C_NCT_OKINAWA) Then
														'舞鶴高専は１年混合学級の専門だけクラス毎に行う
														If m_sGakkoNO = cstr(C_NCT_MAIZURU) Then
															'INS 2007/06/11
															If (gRs("KAMOKU_CD") = 80001 ) OR (gRs("KAMOKU_CD") = 80002 ) Then
																w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "　"
															else
																w_SubjectDisp = w_SubjectDisp & f_GetGakkaNm(m_iNendo,gRs("GAKKA_CD")) & "　"
															End If
															'DEL 2007/06/11 w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "　"
															'INS END 2007/06/11
														'INS 2016/06/14 kiyomoto -->
														elseIf m_sGakkoNO = cstr(C_NCT_KOCHI) Then
																w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "　"
														'INS 2016/06/14 kiyomoto <--
														else
															w_SubjectDisp = w_SubjectDisp & f_GetGakkaNm(m_iNendo,gRs("GAKKA_CD")) & "　"
														End If
													else	'沖縄だったら 2004.09.13 suitoh
														If gRs("KAMOKU_CD") >= 900000 then
															w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "　"
														Else
															w_SubjectDisp = w_SubjectDisp & f_GetGakkaNm(m_iNendo,gRs("GAKKA_CD")) & "  "
														End If
													End If
												Else
													'INS STR 2010/05/20 iwata 高知 混合クラス、一般、指定科目は　学科表示にする
													'UPP Nishimura  (gRs("KAMOKU_CD") = 140050 )を削除
													If m_sGakkoNO = cstr(C_NCT_KOCHI) Then
														If (gRs("KAMOKU_CD") = 140046 ) OR (gRs("KAMOKU_CD") = 140047 ) OR (gRs("KAMOKU_CD") = 140048 ) OR (gRs("KAMOKU_CD") = 180011 ) OR (gRs("KAMOKU_CD") = 180012 ) OR (gRs("KAMOKU_CD") = 180013 ) OR (gRs("KAMOKU_CD") = 180014 ) Then
															w_SubjectDisp = w_SubjectDisp & f_GetGakkaNm(m_iNendo,gRs("GAKKA_CD"))  & "　"
														Else
															w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "　"
														End If
													Else
													'INS END 2010/05/20 iwata
														w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "　"
													End If
												End If
											Else

												w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "　"
											End If

											w_SubjectDisp = w_SubjectDisp & gRs("KAMOKU_NAME") & "　"

											w_TukuName = ""

											if cint(gf_SetNull2Zero(gRs("KAMOKU_KBN"))) = 1 then
												w_TukuName = "TOKU"
											else
												w_TukuName = "TUJO"
											end if

											'科目コンボVALUE部分生成
											w_SubjectValue = ""
											w_SubjectValue = w_SubjectValue & gRs("GAKUNEN")   & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("CLASS")     & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("KAMOKU_CD") & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("GAKKA_CD")  & "#@#"
											w_SubjectValue = w_SubjectValue & f_GetUpdDate(m_iNendo,gRs("GAKUNEN"),gRs("GAKKA_CD"),gRs("KAMOKU_CD"),cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))) & "#@#"
											w_SubjectValue = w_SubjectValue & w_TukuName  & "#@#"
											w_SubjectValue = w_SubjectValue & cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))
											w_SubjectValue = w_SubjectValue & "#@#" & gRs("CLASS_NAME")			'2019/06/24 Add Fujibayashi
											w_SubjectValue = w_SubjectValue & "#@#" & gRs("KAMOKU_NAME")		'2019/06/24 Add Fujibayashi

								%>
										<option value="<%=w_SubjectValue%>"><%=w_SubjectDisp%>
								<%
											gRs.movenext
										loop
									end if
								%>
								</select>
							</td>
						</tr>

						<tr>
							<td align="left" nowrap>最終更新日</td>
							<td align="left" colspan="3" nowrap>
								<input type="text" name="txtUpdDate" value="" onFocus="blur();" readonly style="BACKGROUND-COLOR: #E4E4ED">
							</td>

							<td colspan="7" align="right">
								<input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();" <%=gDisabled%>>
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
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtGakuNo"    value="<%=w_iGakunen_s%>">
	<input type="hidden" name="txtClassNo"   value="">
	<input type="hidden" name="txtKamokuCd"  value="<%=w_sKamokuCd_s%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=w_sGakkaCd_s%>">
	<input type="hidden" name="SYUBETU"      value="">
	<input type="hidden" name="hidKamokuKbn" value="">
	<input type="hidden" name="txtClassMei"   value="">		<%	'2019/06/24 Add Fujibayashi	%>
	<input type="hidden" name="txtKamokuMei"   value="">	<%	'2019/06/24 Add Fujibayashi	%>

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