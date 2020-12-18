<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 使用教科書登録
' ﾌﾟﾛｸﾞﾗﾑID : web/WEB0321/default.asp
' 機	  能: 使用教科書の登録を行う
'-------------------------------------------------------------------------
' 引	  数:教官コード 	＞		SESSIONより（保留）
' 変	  数:なし
' 引	  渡:教官コード 	＞		SESSIONより（保留）
' 説	  明:
'			■フレームページ
'-------------------------------------------------------------------------
' 作	  成: 2001/07/05 岩下　幸一郎
' 変	  更: 2001/07/23 本村　文
' 変	  更: 2001/07/31 伊藤　公子
' 変	  更: 2001/08/01 前田　智史
' 変	  更: 2001/08/18 伊藤　公子 次年度の学期情報がない時は次年度の入力が出来ないようにする
' 変	  更: 2001/12/01 田部　雅幸 自分の所属する学科のみを変更できるように修正
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系

	Public m_sGakkiWhere	'学期の条件
	Public m_sGakkaWhere	'学科コンボの条件
	Public m_sKamokuWhere	'科目の条件
	Public m_sKamokuOption	'科目のオプション
	Public m_sCourseWhere	'科目コースの条件
	Public m_sCourseOption	'科目コースのオプション
	Public m_bErrFlg		'ｴﾗｰﾌﾗｸﾞ
	Public m_iNendo 		'年度
	Public m_sKyokan_CD 	'教官CD
	Public m_iMax
	Public m_iDsp
	Public m_sPageCD
	Public m_sTitle 		''新規登録・修正の表示用
	Public m_sDBMode		''DBへの更新ﾓｰﾄﾞ
	Public m_sMode			''画面の表示のﾓｰﾄﾞ
	Public m_sKengen	''権限(FULLorNOMAL)
	
	''ﾃﾞｰﾀ表示用
	Public m_sNo
	Public m_sNendo
	Public m_sGakkiCD
	Public m_sGakunenCD
	Public m_sGakkaCD
	Public m_sKamokuCD
	Public m_sCourseCD
	Public m_sKyokan_NAME		'教官
	Public m_sKyokasyo_NAME 	'教科書
	Public m_sSyuppansya		'出版社
	Public m_sTyosya			'著者名
	Public m_sSidousyo			'指導書
	Public m_sKyokanyo			'教官用
	Public m_sBiko				'備考

	Public m_sNendoOption
	Public m_bJinendoGakki		'//次年度の学期情報があるかどうか

	Public m_sSyozokuGakka		'//2001/12/01 Add ログインした教官の所属する学科

	Public m_sGetSQL			'2001/12/01 Add

'///////////////////////////メイン処理/////////////////////////////

	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

Sub Main()
'********************************************************************************
'*	[機能]	本ASPのﾒｲﾝﾙｰﾁﾝ
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	Dim w_iRet				'// 戻り値

	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="就職先マスタ登録"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False
	m_iDsp = C_PAGE_LINE

	m_bErrFlg = False

	Do
		'// ﾃﾞｰﾀﾍﾞｰｽ接続
		w_iRet = gf_OpenDatabase()
		If w_iRet <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'// 値を変数に入れる
		Call s_SetParam()


		'// 更新用のﾃﾞｰﾀを表示する
		if m_sMode = "Kousin" then
			if f_GetData() = False then
				exit do
			end if
		end if

		'// 教官の名称を取得する
		if f_GetData_Kyokan() = False then
			exit do
		end if

		'//次年度情報があるかチェック
		w_iRet = f_GetJinendoGakki(m_bJinendoGakki)
		If w_iRet  = False Then
			m_bErrFlg = True
				exit do
		End If

		'学期に関するWHREを作成する
		Call f_MakeGakkiWhere() 
		'学科に関するWHREを作成する
		Call f_MakeGakkaWhere()
		'学科コースに関するWHREを作成する
		Call f_MakeCourseWhere()
		'科目に関するWHREを作成する
		Call f_MakeKamokuWhere()

		'//権限を取得
		w_iRet = gf_GetKengen_WEB0320(m_sKengen)
		If w_iRet <> 0 Then
			m_bErrFlg = true
			w_sMsg = "権限がありません。"
			Exit Do
		End If

		'// ページを表示
		Call showPage()

		Exit Do

	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// 終了処理
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	次年度の学期情報があるかどうかチェックする
'*	[引数]	なし
'*	[戻値]	p_bJinendoGakki=true:学期情報あり
'*			p_bJinendoGakki=false:学期情報なし
'*	[説明]	
'********************************************************************************
Function f_GetJinendoGakki(p_bJinendoGakki)
	Dim w_iRet				'// 戻り値
	Dim w_sSQL				'// SQL文
	dim w_Rs

	on error resume next
	err.clear

	f_GetJinendoGakki = False
	p_bJinendoGakki = False

	'//次年度の学期情報があるかどうか
	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI"
	w_sSQL = w_sSQL & vbCrLf & " FROM M01_KUBUN"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_NENDO=" & cint(SESSION("NENDO"))+1
	w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & C_KAISETUKI

	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		m_bErrFlg = True
		Exit Function
	End If

	'//データがあった時
	If w_Rs.EOF = False Then
		p_bJinendoGakki = True
	End If

	Call gf_closeObject(w_Rs)

	f_GetJinendoGakki = True

End Function

'********************************************************************************
'*	[機能]	値を変数に入れる
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_SetParam()

	m_iNendo	 = Session("NENDO")
	m_sMode 	 = Request("txtMode")		':モード

	'//ﾀｲﾄﾙｾｯﾄ
	if m_sMode = "Touroku" then
		m_sTitle = "新規登録"
	elseif m_sMode = "Kousin" then
		m_sTitle = "修正"
	else
		m_sTitle = Request("txtTitle")	'//リロード時
	end if

	''DBの登録時のﾓｰﾄﾞの設定
	if m_sTitle ="新規登録" then
		m_sDBMode="Insert"
		m_sNendoOption = ""
	else
		m_sDBMode="Update"
		m_sNendoOption = "DISABLED"
	end if

	'//一覧表示中ページを保存
	m_sPageCD	 = Request("txtPageCD")

	''ﾃﾞｰﾀ表示用
	if m_sMode = "Touroku" then
		m_sNendo = Session("NENDO")
		m_sNo = ""		''更新用No格納
		m_sKyokan_CD = session("KYOKAN_CD")

	elseif m_sMode = "Kousin" then
		m_sNendo = Session("NENDO")
		m_sNo = Request("txtUpdNo") 	''更新用No格納
		m_sKyokan_CD = Request("SKyokanCd1")
	else	'//リロード時
		m_sNendo  = Request("txtNendo")
		m_sNo = Request("txtUpdNo") 	''更新用No格納
		m_sKyokan_CD = Request("SKyokanCd1")
	end if

	'm_sKyokan_CD = Session("KYOKAN_CD")
	'm_sKyokan_CD = Request("SKyokanCd1")

	m_sGakkiCD	 = Request("txtGakkiCD")
	m_sGakunenCD = Request("txtGakunenCD")


	m_sGakkaCD	 = Request("txtGakkaCD")
	m_sKamokuCD  = Request("txtKamokuCD")
	m_sCourseCD  = Request("txtCourseCD")
	m_sKyokan_NAME	= Request("txtKyokanMei")		'教官
	m_sKyokasyo_NAME  = Request("txtKyokasyoName")	'教科書
	m_sSyuppansya  = Request("txtSyuppansya")		'出版社
	m_sTyosya  = Request("txtTyosya")				'著者名
	m_sSidousyo  = Request("txtSidousyo")			'指導書
	m_sKyokanyo  = Request("txtKyokanyo")			'教官用
	m_sBiko  = trim(Request("txtBiko")) 				  '備考

End Sub

'********************************************************************************
'*	[機能]	デバッグ用
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iNendo		= " & m_iNendo			& "<br>"
	response.write "m_sMode			= " & m_sMode			& "<br>"
	response.write "m_sTitle		= " & m_sTitle			& "<br>"
	response.write "m_sDBMode		= " & m_sDBMode			& "<br>"
	response.write "m_sPageCD		= " & m_sPageCD			& "<br>"
	response.write "m_sNendo		= " & m_sNendo			& "<br>"
	response.write "m_sNo			= " & m_sNo				& "<br>"
	response.write "m_sKyokan_CD	= " & m_sKyokan_CD		& "<br>"
	response.write "m_sGakkiCD		= " & m_sGakkiCD		& "<br>"
	response.write "m_sGakunenCD	= " & m_sGakunenCD		& "<br>"
	response.write "m_sGakkaCD		= " & m_sGakkaCD		& "<br>"
	response.write "m_sKamokuCD		= " & m_sKamokuCD		& "<br>"
	response.write "m_sCourseCD		= " & m_sCourseCD		& "<br>"
	response.write "m_sKyokan_NAME	= " & m_sKyokan_NAME	& "<br>"
	response.write "m_sKyokasyo_NAME= " & m_sKyokasyo_NAME	& "<br>"
	response.write "m_sSyuppansya	= " & m_sSyuppansya		& "<br>"
	response.write "m_sTyosya		= " & m_sTyosya			& "<br>"
	response.write "m_sSidousyo		= " & m_sSidousyo		& "<br>"
	response.write "m_sKyokanyo		= " & m_sKyokanyo		& "<br>"
	response.write "m_sBiko			= " & m_sBiko			& "<br>"

End Sub

'********************************************************************************
'*	[機能]	教官の名称を取得する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
function f_GetData_Kyokan()
	Dim w_iRet				'// 戻り値
	Dim w_sSQL				'// SQL文
	dim w_Rs

	f_GetData_Kyokan = False

	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & " M04.M04_NENDO "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKAN_CD "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN M04 "
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "    M04_NENDO = " &  m_iNendo & " AND "
	w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN_CD = '" & m_sKyokan_CD & "'"

	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		m_bErrFlg = True
		Exit Function
	Else
		'ページ数の取得
		m_iMax = gf_PageCount(w_Rs,m_iDsp)
	End If

	m_sKyokan_NAME = ""
	If w_Rs.EOF = False Then
		m_sKyokan_NAME = w_Rs("M04_KYOKANMEI_SEI") & "  " & w_Rs("M04_KYOKANMEI_MEI")
	End If

	w_Rs.close

	f_GetData_Kyokan = True

end function

'********************************************************************************
'*	[機能]	更新時の表示ﾃﾞｰﾀを取得する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
function f_GetData()
	Dim w_iRet				'// 戻り値
	Dim w_sSQL				'// SQL文
	Dim w_Rs

	f_GetData = False

	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & " T47.T47_NENDO "			''年度
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKI_KBN "		''学期区分
'	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_NO"				''No
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKUNEN " 		''学年
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKA_CD "		''学科
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_COURSE_CD "		''ｺｰｽｺｰﾄﾞ
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KAMOKU "			''科目ｺｰﾄﾞ
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKASYO "		''教科書名
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SYUPPANSYA "		''出版社
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_TYOSYA "			''著者
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKANYOUSU " 	''教官用数
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SIDOSYOSU "		''指導書数
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_BIKOU "			''備考
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKAN "			 ''教官
	w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKAMEI "
	w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "    T47_KYOKASYO T47 "
	w_sSQL = w_sSQL & vbCrLf & "    ,M02_GAKKA M02 "
	w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
	w_sSQL = w_sSQL & vbCrLf & "    ,M04_KYOKAN M04 "
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M02.M02_NENDO(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_GAKKA_CD  = M02.M02_GAKKA_CD(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M03.M03_NENDO(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M04.M04_NENDO(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = M04.M04_KYOKAN_CD(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO = " & Request("KeyNendo") & " AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = '" & m_sKyokan_CD & "' AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NO = " & m_sNo & ""

response.write(w_sSQL & "<BR>")
	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		m_bErrFlg = True
		Exit Function
	Else
		'ページ数の取得
		m_iMax = gf_PageCount(w_Rs,m_iDsp)
	End If
response.write("Set<BR>")

	m_sNendo   = gf_HTMLTableSTR(w_Rs("T47_NENDO"))
	m_sGakkiCD	 = gf_HTMLTableSTR(w_Rs("T47_GAKKI_KBN"))
	m_sGakunenCD = gf_HTMLTableSTR(w_Rs("T47_GAKUNEN"))
	m_sGakkaCD	 = gf_HTMLTableSTR(w_Rs("T47_GAKKA_CD"))
	m_sKamokuCD  = gf_HTMLTableSTR(w_Rs("T47_KAMOKU"))
	m_sCourseCD  = gf_HTMLTableSTR(w_Rs("T47_COURSE_CD"))
	m_sKyokasyo_NAME  = gf_HTMLTableSTR(w_Rs("T47_KYOKASYO"))		'教科書
	m_sSyuppansya  = gf_HTMLTableSTR(w_Rs("T47_SYUPPANSYA"))		'出版社
	m_sTyosya  = gf_HTMLTableSTR(w_Rs("T47_TYOSYA"))				'著者名
	m_sSidousyo  = gf_HTMLTableSTR(w_Rs("T47_SIDOSYOSU"))			'指導書
	m_sKyokanyo  = gf_HTMLTableSTR(w_Rs("T47_KYOKANYOUSU")) 		'教官用
	m_sBiko  = gf_HTMLTableSTR(w_Rs("T47_BIKOU"))					'備考

	m_sKyokan_CD = gf_HTMLTableSTR(w_Rs("T47_KYOKAN"))

	w_Rs.close

	f_GetData = True
response.write("f_GetData<BR>")

end function


'********************************************************************************
'*	[機能]	学期コンボに関するWHREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub f_MakeGakkiWhere()
Dim w_sNendo
	m_sGakkiWhere=""

	'//新規登録時、次年度情報があるときは次年度を使用。ない時は当年度を使用。
'	If m_bJinendoGakki = True Then
'		w_sNendo = cint(m_iNendo) + 1
'	Else
'		w_sNendo = cint(m_iNendo)
'	End If

	w_sNendo = cint(request("txtNendo"))

	'm_sGakkiWhere = " M01_DAIBUNRUI_CD = 51  AND "
	m_sGakkiWhere = " M01_DAIBUNRUI_CD = " & C_KAISETUKI & " AND "
	m_sGakkiWhere = m_sGakkiWhere & " M01_SYOBUNRUI_CD <> 3 AND "	'<--"開設しない"以外
	If m_sMode = "Touroku" Then
	  m_sGakkiWhere = m_sGakkiWhere & " M01_NENDO = " & w_sNendo  & ""
	Else
	  m_sGakkiWhere = m_sGakkiWhere & " M01_NENDO = " & m_sNendo & ""
	End If

'response.write m_sGakkiWhere & "<BR>"

End Sub

'********************************************************************************
'*	[機能]	学科コンボに関するWHREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub f_MakeGakkaWhere()
	Dim w_sNendo

	'2001/12/01 Add ---->
	Dim w_sSQL				'//SQL文
	Dim w_iRet				'//戻り値

	Dim w_oRecord			'//所属学科取得のため

	'//所属学科の取得
	w_sSQL = ""
	w_sSQL = w_sSQL & "SELECT "
	w_sSQL = w_sSQL & "M04_GAKKA_CD "
	w_sSQL = w_sSQL & "From "
	w_sSQL = w_sSQL & "M04_KYOKAN "
	w_sSQL = w_sSQL & "Where "
	w_sSQL = w_sSQL & "M04_NENDO = " & m_iNendo & " "
	w_sSQL = w_sSQL & "And "
	w_sSQL = w_sSQL & "M04_KYOKAN_CD = '" & Session("KYOKAN_CD") & "'"

	w_iRet = gf_GetRecordset(w_oRecord, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit Sub
	End If

	If w_oRecord.EOF <> True Then
		m_sSyozokuGakka = w_oRecord("M04_GAKKA_CD")
	Else
		m_sSyozokuGakka =""
	End If

	'//閉じる
	w_oRecord.Close
	Set w_oRecord = Nothing

	'2001/12/01 Add <----

	'//新規登録時、次年度情報があるときは次年度を使用。ない時は当年度を使用。
'	If m_bJinendoGakki = True Then
'		w_sNendo = cint(m_iNendo) + 1
'	Else
'		w_sNendo = cint(m_iNendo)
'	End If

	w_sNendo = cint(request("txtNendo"))

	m_sGakkaWhere=""

	If m_sMode = "Touroku" Then
		m_sGakkaWhere = " M02_NENDO = " & w_sNendo	& ""
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD <> '00' "
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD = '" & m_sSyozokuGakka & "' "	'2001/12/01 Mod
	Else
		m_sGakkaWhere = " M02_NENDO = " & m_sNendo & ""
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD <> '00' "
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD = '" & m_sSyozokuGakka & "' "	'2001/12/01 Mod
	End If

End Sub

'********************************************************************************
'*	[機能]	学科コースコンボに関するWHREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub f_MakeCourseWhere()
Dim w_sNendo

	'w_sNendo = cint(m_iNendo) + 1

	'//新規登録時、次年度情報があるときは次年度を使用。ない時は当年度を使用。
'	If m_bJinendoGakki = True Then
'		w_sNendo = cint(m_iNendo) + 1
'	Else
'		w_sNendo = cint(m_iNendo)
'	End If

	w_sNendo = cint(request("txtNendo"))

	m_sCourseWhere=""
	m_sCourseOption=""

	If m_sMode = "Touroku" Then
		m_sCourseOption = " DISABLED "
		m_sCourseWhere = " M20_NENDO = " & w_sNendo  & ""
		Exit Sub
	End If

	''学科未選択時は、学科ｺｰｽは未選択
	'if m_sGakkaCD = "@@@" then
	if m_sGakkaCD = "@@@" Or m_sGakkaCD = "" then
		m_sCourseOption = " DISABLED "
		m_sCourseWhere = " M20_NENDO = " & w_sNendo  & ""
		m_sCourseCD = "@@@"
		Exit Sub
	end if

	''全学科の時は、学科ｺｰｽは使用不可
	if cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) then
		m_sCourseOption = " DISABLED "
		m_sCourseWhere = " M20_NENDO = " & w_sNendo  & ""
		m_sCourseCD = "@@@"
		Exit Sub
	end if


''	If m_sGakkaCD = 99 Then
''		m_sCourseWhere= " M20_NENDO = " & m_sNendo & " AND "
''		m_sCourseWhere = m_sCourseWhere & " M20_GAKUNEN =  " & m_sGakunenCD & ""
''		m_sCourseWhere = m_sCourseWhere & " Group By M20_GAKKA_CD , M20_COURSE_CD, M20_COURSEMEI "
''	Else
		m_sCourseWhere = " M20_NENDO = " & m_sNendo & " AND "
		m_sCourseWhere = m_sCourseWhere & " M20_GAKKA_CD = " & m_sGakkaCD & " AND "
		m_sCourseWhere = m_sCourseWhere & " M20_GAKUNEN =  " & m_sGakunenCD & ""
		m_sCourseWhere = m_sCourseWhere & " Group By M20_GAKKA_CD , M20_COURSE_CD, M20_COURSEMEI "
''	End IF

End Sub

'********************************************************************************
'*	[機能]	科目コンボに関するWHREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub f_MakeKamokuWhere()

	m_sKamokuWhere=""
	m_sKamokuOption=""

	'//新規登録時、次年度情報があるときは次年度を使用。ない時は当年度を使用。
'	If m_bJinendoGakki = True Then
'		w_sNendo = cint(m_iNendo) + 1
'	Else
'		w_sNendo = cint(m_iNendo)
'	End If

	w_sNendo = cint(request("txtNendo"))

	m_sGetSQL = ""
	m_sGetSQL = m_sGetSQL & "Select "
	m_sGetSQL = m_sGetSQL & "Distinct "
	m_sGetSQL = m_sGetSQL & "T15_KAMOKU_CD, "
	m_sGetSQL = m_sGetSQL & "T15_KAMOKUMEI "
	m_sGetSQL = m_sGetSQL & "From "
	m_sGetSQL = m_sGetSQL & "T15_RISYU, "
	m_sGetSQL = m_sGetSQL & "T27_TANTO_KYOKAN "
	m_sGetSQL = m_sGetSQL & "Where "

	If m_sGakunenCD <> "" Then
		'//学年が指定されている場合
		m_sGetSQL = m_sGetSQL & "T15_NYUNENDO = " & (cint(w_sNendo) - cint(m_sGakunenCD) + 1) & " "
	Else
		'//学年が指定されていない場合
		m_sGetSQL = m_sGetSQL & "T15_NYUNENDO = " & cint(w_sNendo) & " "
	End If

	If m_sGakkaCD <> "" Then
		'//学科が指定されている場合
		If cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) Then
			'全学科の場合
			m_sGetSQL = m_sGetSQL & " AND T15_KAMOKU_KBN = " & C_KAMOKU_IPPAN

			If m_sCourseCD <> "@@@" AND m_sCourseCD <> "" Then
				m_sGetSQL = m_sGetSQL & " AND T15_COURSE_CD = " & m_sCourseCD &""
			End IF

			m_sGetSQL = m_sGetSQL & " AND T15_KAISETU" & m_sGakunenCD & "="  & m_sGakkiCD

		Else
			'個別の学科
			m_sGetSQL = m_sGetSQL & " AND T15_GAKKA_CD = '" & m_sGakkaCD &"' "
			m_sGetSQL = m_sGetSQL & " AND T15_KAISETU" & m_sGakunenCD & " = " & m_sGakkiCD
			m_sGetSQL = m_sGetSQL & " AND T15_KAMOKU_KBN <> " & C_KAMOKU_IPPAN & " "

			If cstr(gf_SetNull2String(m_sCourseCD)) <> "@@@" AND trim(cstr(gf_SetNull2String(m_sCourseCD))) <> "" Then
				m_sGetSQL = m_sGetSQL & " AND T15_COURSE_CD = " & m_sCourseCD &""
			End If

		End If

	End If
	
	m_sGetSQL = m_sGetSQL & " AND "
	m_sGetSQL = m_sGetSQL & "T27_NENDO = " & w_sNendo & " "
	m_sGetSQL = m_sGetSQL & " AND "
	m_sGetSQL = m_sGetSQL & "T15_KAMOKU_CD = T27_KAMOKU_CD "
	m_sGetSQL = m_sGetSQL & " AND "
	m_sGetSQL = m_sGetSQL & "T27_KYOKAN_CD = " & Session("KYOKAN_CD")

	If m_sGakunenCD <> "" Then
		'//学年が指定されている場合
		m_sGetSQL = m_sGetSQL & " AND "
		m_sGetSQL = m_sGetSQL & "T27_GAKUNEN = " & m_sGakunenCD & " "
	End If

	m_sGetSQL = m_sGetSQL & " Group By T15_NYUNENDO , T15_KAMOKU_CD , T15_KAMOKUMEI "

	''新規登録時
	If m_sMode = "Touroku" Then
		m_sKamokuOption = " DISABLED "
		Exit Sub
	End If

	''学科未選択時は、科目は未選択
	if m_sGakkaCD = "@@@" Or m_sGakkaCD = "" then
		m_sKamokuOption = " DISABLED "
		m_sKamokuCD = "@@@"
		Exit Sub
	end if

End Sub


'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'		(リストダウンボックス選択表示用)
'[引数] pData1 : データ１
'		pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'					
'****************************************************
Function f_Selected(pData1,pData2)

	f_Selected = ""

	If IsNull(pData1) = False And IsNull(pData2) = False Then
		If trim(cStr(pData1)) = trim(cstr(pData2)) Then
			f_Selected = "selected" 
		Else
		End If
	End If

End Function


Sub showPage()
'********************************************************************************
'*	[機能]	HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
%>

<html>

<head>
<!-- <%= m_sGetSQL %> -->


<title>使用教科書登録</title>

	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	//************************************************************
	//	[機能]	使用教科書登録
	//	[引数]	p_iPage :表示頁数
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_touroku(){

		// 入力値のﾁｪｯｸ
		iRet = f_CheckData();
		if( iRet != 0 ){
			return;
		}

		document.frm.txtKyokanMei.value=document.frm.SKyokanNm1.value

		document.frm.action="./kakunin.asp";
		document.frm.target="";
		document.frm.txtMode.value = "<%=m_sDBMode%>";
		document.frm.submit();
	}

	//************************************************************
	//	[機能]	進路が修正されたとき、再表示する
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_ReLoadMyPage(){

		document.frm.action="./touroku.asp";
		document.frm.target="";
		document.frm.txtMode.value = "Reload";
		document.frm.submit();
	
	}

	//************************************************************
	//	[機能]	メインページへ戻る
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Back(){

		document.frm.action="./default.asp";
		document.frm.target="";
		document.frm.txtMode.value = "Back";
		document.frm.submit();
	
	}

	//************************************************************
	//	[機能]	入力値のﾁｪｯｸ
	//	[引数]	なし
	//	[戻値]	0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
	//	[説明]	入力値のNULLﾁｪｯｸ、英数字ﾁｪｯｸ、桁数ﾁｪｯｸを行う
	//			引渡ﾃﾞｰﾀ用にﾃﾞｰﾀを加工する必要がある場合には加工を行う
	//************************************************************
	function f_CheckData() {
	
	
		// ■■■NULLﾁｪｯｸ■■■
		// ■学科コード
		if( f_Trim(document.frm.txtGakkaCD.value) == "@@@" ){
			window.alert("学科が選択されていません");
			if(document.frm.txtGakkaCD.length!=1){
			document.frm.txtGakkaCD.focus();
			}
			return 1;
		}

		// ■■■NULLﾁｪｯｸ■■■
		// ■科目コード
		if( f_Trim(document.frm.txtKamokuCD.value) == "@@@" ){
			window.alert("科目が選択されていません");
			if(document.frm.txtKamokuCD.length!=1){
				document.frm.txtKamokuCD.focus();
			}
			return 1;
		}

		// ■■■NULLﾁｪｯｸ■■■
		// ■教科書名
		if( f_Trim(document.frm.txtKyokasyoName.value) == "" ){
			window.alert("教科書名が入力されていません");
			document.frm.txtKyokasyoName.focus();
			return 1;
		}

		// ■■■教科書名の桁ﾁｪｯｸ■■■
		if( getLengthB(document.frm.txtKyokasyoName.value) > "80" ){
			window.alert("教科書名の欄は全角40文字以内で入力してください");
			document.frm.txtKyokasyoName.focus();
			return 1;
		}

		// ■■■出版社名の桁ﾁｪｯｸ■■■
		if( getLengthB(document.frm.txtSyuppansya.value) > "40" ){
			window.alert("出版社名の欄は全角20文字以内で入力してください");
			document.frm.txtSyuppansya.focus();
			return 1;
		}

		// ■■■著者名の桁ﾁｪｯｸ■■■
		if( getLengthB(document.frm.txtTyosya.value) > "40" ){
			window.alert("著者名の欄は全角20文字以内で入力してください");
			document.frm.txtTyosya.focus();
			return 1;
		}

		// ■■■教官用冊数の値ﾁｪｯｸ■■■
		if(f_Trim(document.frm.txtKyokanyo.value)!=""){
			//数値チェック
			if( isNaN(document.frm.txtKyokanyo.value)){
				window.alert("冊数は半角数字で入力してください");
				document.frm.txtKyokanyo.focus();
				return 1;
			}else{
				//桁チェック
				if( getLengthB(document.frm.txtKyokanyo.value) > "3" ){
					window.alert("冊数は3桁以内で入力してください");
					document.frm.txtKyokanyo.focus();
					return 1;
				}
			}
		}

		// ■■■指導用冊数の値ﾁｪｯｸ■■■
		if(f_Trim(document.frm.txtSidousyo.value)!=""){
			//数値チェック
			if( isNaN(document.frm.txtSidousyo.value)){
				window.alert("冊数は半角数字で入力してください");
				document.frm.txtSidousyo.focus();
				return 1;
			}else{
				//桁チェック
				if( getLengthB(document.frm.txtSidousyo.value) > "3" ){
					window.alert("冊数は3桁以内で入力してください");
					document.frm.txtSidousyo.focus();
					return 1;
				}
			}
		}

		// ■■■備考の桁ﾁｪｯｸ■■■
		if( getLengthB(document.frm.txtBiko.value) > "80" ){
			window.alert("備考の欄は全角40文字以内で入力してください");
			document.frm.txtBiko.focus();
			return 1;
		}

		return 0;
	}

	//************************************************************
	//	[機能]	教官参照選択画面ウィンドウオープン
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function KyokanWin(p_iInt,p_sKNm) {
		var obj=eval("document.frm."+p_sKNm)
		var w_gak = document.frm.txtGakkaCD.value
		URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"&txtGakka="+w_gak+"";
		nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=600,top=0,left=0");
		nWin.focus();
		return true;	
	}
	//************************************************************
	//	[機能]	クリアボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function fj_Clear(){
		//教官欄を空白にする
		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}

	//-->
	</script>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">

	</head>
	<body>
	<form name="frm" action="" target="" method="post">

<%'call s_DebugPrint%>

	<center>
	<% call gs_title("使用教科書登録",m_sTitle) %>
	<br>
<table border="0" cellpadding="1" cellspacing="1" width="540">
	<tr>
		<td align="left">
			<table width="100%" border=1 CLASS="hyo">
				<tr>
				<th height="16" width="75" class=header nowrap>年　度</th>
				<td height="16" width="430" class=detail nowrap>
				<%If m_sDBMode="Update" Then%>
					<%=Request("KeyNendo")%>
					<input type="hidden" name="txtNendo" value="<%=Request("KeyNendo")%>">
				<%Else%>
					<select name="txtNendo" onchange='javascript:f_ReLoadMyPage()'	>

						<%'//新規登録時、次年度情報があるときは次年度を使用。ない時は当年度を使用。
						If m_bJinendoGakki = True Then
							'//新規登録時、次年度情報があるときは次年度を使用。ない時は当年度を使用。
							If m_sMode = "Touroku" Then
								w_sNendo = cint(m_iNendo) + 1
							Else
								w_sNendo = m_sNendo
							End If
						%>
							<option VALUE="<%= m_iNendo + 1 %>" <%= f_Selected(cstr(w_sNendo),cstr(cint(m_iNendo+1)))%>><%= m_iNendo + 1 %>
							<option VALUE="<%= m_iNendo %>" 	<%= f_Selected(cstr(w_sNendo),cstr(m_iNendo))%>><%= m_iNendo %>
						<%Else%>
							<option VALUE="<%= m_iNendo %>" 	<%= f_Selected(cstr(m_sNendo),cstr(m_iNendo))%>><%= m_iNendo %>
						<%End If%>
					</select><span class=hissu>*</span>
				<%End If%>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class="header" nowrap>学　期</th>
				<td height="16" width="430" class="detail" nowrap>

				<%'共通関数から学期に関するコンボボックスを出力する
				If m_sMode = "Touroku" Then
						call gf_ComboSet("txtGakkiCD",C_CBO_M01_KUBUN,m_sGakkiWhere,"onchange = 'javascript:f_ReLoadMyPage()'",False,0)
					Else
						call gf_ComboSet("txtGakkiCD",C_CBO_M01_KUBUN,m_sGakkiWhere,"onchange = 'javascript:f_ReLoadMyPage()'",False,m_sGakkiCD)
				End If
				%><span class=hissu>*</span>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>学　年</th>
				<td height="16" width="430" class=detail nowrap>
					<select name="txtGakunenCD" onchange = 'javascript:f_ReLoadMyPage()'>
						<option Value="1" <%= f_Selected( 1 ,m_sGakunenCD) %>>1年
						<option Value="2" <%= f_Selected( 2 ,m_sGakunenCD) %>>2年
						<option Value="3" <%= f_Selected( 3 ,m_sGakunenCD) %>>3年
						<option Value="4" <%= f_Selected( 4 ,m_sGakunenCD) %>>4年
						<option Value="5" <%= f_Selected( 5 ,m_sGakunenCD) %>>5年
					</select><span class=hissu>*</span>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>学　科</th>
				<td height="16" width="430" class=detail nowrap>
				<%	'共通関数から学科に関するコンボボックスを出力する
					call f_ComboSet_Gakka("txtGakkaCD",C_CBO_M02_GAKKA,m_sGakkaWhere,"style='width:175px;' onchange = 'javascript:f_ReLoadMyPage()'",True,m_sGakkaCD)%>
				<span class=hissu>*</span><img src="../../image/sp.gif" width="10">
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>コース</font></th>
				<td height="16" width="430" class=detail>
				<%	'共通関数から学科コースに関するコンボボックスを出力する
					call gf_ComboSet("txtCourseCD",C_CBO_M20_COURSE,m_sCourseWhere,"style='width:175px;' onchange = 'javascript:f_ReLoadMyPage()'" & m_sCourseOption,True,m_sCourseCD)%>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>科　目</font></th>
				<td height="16" width="430" class=detail>
<!--
<%= m_sGakunenCD %>
<%= m_sGakkaCD %>
-->
				<%	'共通関数から科目に関するコンボボックスを出力する
					'学年が条件 学科が入力されていないときは、DISABLEDとなる
					call f_ComboSet("txtKamokuCD",C_CBO_T15_RISYU,m_sKamokuWhere,"style='width:175px;'" & m_sKamokuOption,True,m_sKamokuCD)%>
				<span class=hissu>*</span>
				</td>
				</tr>
				<tr>

				<th height="16" width="80" class=header nowrap>教官</font></th>
				<td height="16" width="430" class=detail nowrap>
					<input type="text" class="text" name="SKyokanNm1" VALUE='<%=m_sKyokan_NAME%>' readonly size="30">
					<input type="hidden" name="SKyokanCd1" VALUE='<%=m_sKyokan_CD%>'>
					<%
					'//最高権限者のみ利用者の変更を可とする
					If m_sKengen = C_ACCESS_FULL Then%>
						<input type="button" class="button" value="選択" onclick="KyokanWin(1,'SKyokanNm1')">
						<input type="button" class="button" value="クリア" onClick="fj_Clear()">
					<%End If%>
				</td>
				</tr>

				<tr>
				<th height="16" width="80" class=header nowrap>教科書名</font></th>
				<td height="16" width="430" class=detail nowrap>
				<textarea cols="56" rows="3" Name="txtKyokasyoName" Value="<%= m_sKyokasyo_NAME %>"><%= m_sKyokasyo_NAME %></textarea>
				<span class=hissu>*</span><font size=2><BR>（全角40文字以内）</font>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>出版社</font></th>
				<td height="16" width="430"  class=detail nowrap>
				<input type="text" size="56" Name="txtSyuppansya" Value="<%= m_sSyuppansya %>"><BR><font size=2>（全角20文字以内）</font>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>著者名</font></th>
				<td height="16" width="430" class=detail nowrap>
				<input type="text" size="56" Name="txtTyosya" Value="<%= m_sTyosya %>"><BR><font size=2>（全角20文字以内）</font>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>教官用</font>
				</th>
				<td height="16" width="430" class=detail nowrap>
				<input type="text" size="3" Name="txtKyokanyo" Value="<%= m_sKyokanyo %>" maxlength="3">冊
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>指導書</font>
				</th>
				<td height="16" width="430" class=detail nowrap>
				<input type="text" size="3" Name="txtSidousyo" Value="<%= m_sSidousyo %>"  maxlength="3">冊
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>備　考</font></th>
				<td height="16" width="430" class=detail nowrap>
				<textarea cols="56" rows="3" Name="txtBiko"  Value="<%= trim(m_sBiko) %>"><%= trim(m_sBiko)%></textarea><font size=2>（全角40文字以内）</font>
				</td>
				</TR>
			</TABLE>
			<table width=75%><tr><td align=right><span class=hissu>*印は必須項目です。</span></td></tr></table>
		</td>
	</TR>
</TABLE>
		<table border="0" width=300>
		<tr>
		<td valign="top" align=left>
		<input type="button" class=button value="　登　録　" OnClick="f_touroku()">
			<img src="../../image/sp.gif" width="30" height="1">
		</td>
		<td valign="top" align=right>
		<input type="Button" class=button value="キャンセル" OnClick="f_Back()">
		</td>
		</tr>
		</table>

		</center>

		<input type="hidden" name="txtMode" value="Touroku">
		<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
		<input type="hidden" name="txtUpdNo" value="<%= m_sNo %>">
		<input type="hidden" name="txtTitle" value="<%= m_sTitle %>">

		<input type="hidden" name="KeyNendo" value="<%=Request("KeyNendo")%>">
		<input type="hidden" Name="txtKyokanMei" Value="<%= m_sKyokan_NAME %>">

	</form>
	</body>
	</html>

<%
End Sub

Function f_ComboSet_Gakka(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' 機	能:ComboBoxセット
' 返	値:OK=True/NG=False
' 引	数:p_oCombo - ComboBox
'		   p_sTableName - テーブル名
'		   p_sWhere - Where条件(WHERE句は要らない)
'		   p_sSelectOption - <SELECT>タグにつけるオプション( onchange = 'a_change()' )など
'		   p_bWhite - 先頭に空白をつけるか
'		   p_sSelectCD - 標準選択させたいコード(""なら選択なし)
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTしてComboBoxにセットする
' 備	考:所属学科が一般総合学科の場合は全学科がつく
'*************************************************************************************
	Dim w_sId			'IDフィールド名
	Dim w_sName 		'名称フィールド名
	Dim w_sTableName	'名称テーブル名
	Dim w_rst

	f_ComboSet_Gakka = False

	do 
	''マスタ毎にSELECTするフィールド名を取得
	If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
		Exit Do
	End If

	''マスタSELECT
	If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
		Exit Do
	End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
	Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

	'空白のOptionの代入
	If p_bWhite Then
		response.Write " <Option Value="&C_CBO_NULL&">　　　　　 "& Chr(13)
	End If

	''EOFでなければ、データをセット
	If Not w_rst.EOF Then
		Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
	End If

	'// 一般総合学科の場合は全学科を選択可能
	If m_sSyozokuGakka = "00" Then
		response.write(" <Option Value='" & C_CLASS_ALL & "'")
		If CStr(p_sSelectCD) = CStr(C_CLASS_ALL) Then
			response.write " Selected "
		End If
		response.Write(">" & "全学科" & Chr(13))
	End If

	Response.write("</select>" & chr(13))

	If Not w_rst Is Nothing Then
		w_rst.Close
		Set w_rst = Nothing
	End If
   
	f_ComboSet_Gakka = True
	Exit Do
	Loop
End Function

'/*****************************

Public Function f_ComboSet(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
	Dim w_sId			'IDフィールド名
	Dim w_sName 		'名称フィールド名
	Dim w_sTableName	'名称テーブル名
	Dim w_rst

	do
		'//データの取得
		If gf_GetRecordset(w_rst, m_sGetSQL) <> 0 Then
			Exit Function
		End If

		If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
		Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

		'空白のOptionの代入
		If p_bWhite Then
			response.Write " <Option Value="&C_CBO_NULL&">　　　　　 "& Chr(13)
		End If

		''EOFでなければ、データをセット
		If Not w_rst.EOF Then
			Call s_MstDataSet(p_sCombo, w_rst, "T15_KAMOKU_CD", "T15_KAMOKUMEI", p_sSelectCD)
		End If

		Response.write("</select>" & chr(13))

		If Not w_rst Is Nothing Then
			w_rst.Close
			Set w_rst = Nothing
		End If

		f_ComboSetf_ComboSet = True
		Exit Do
	Loop

End Function



%>
