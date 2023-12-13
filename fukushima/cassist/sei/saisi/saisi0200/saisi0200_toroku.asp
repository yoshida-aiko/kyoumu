<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 追試成績登録
' ﾌﾟﾛｸﾞﾗﾑID : saisi/saisi0200/saisi0200_toroku.asp
' 機      能: 追試験の登録画面
'-------------------------------------------------------------------------
' 引      数    
'               
' 変      数
' 引      渡
'           
'           
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2003/02/17 矢野
' 変      更: 2003/03/07 矢野　受講回数を表示しない様に変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
Dim m_Rs		
Dim m_RsHyoka

Const mC_HYOKA_NO = 2
Const mC_HYOKA_KBN = 0

Const mC_KAISETU_Z   = 1
Const mC_HYOKA_CD_OK = 0

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

Sub Main()
'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="ヘッダーデータ"
    w_sMsg=""
    w_sRetURL="../../default.asp"
    w_sTarget="_parent"

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'// 成績登録の場合
		if Request("hidMode") = "update" then
			if wf_UpdateSeiseki() = false then
				m_bErrFlg = True
				m_sErrMsg = "成績の登録に失敗しました。"
				Exit Do
			end if
'Response.write "完了"
'Response.end
			'REDIRECT処理
			response.redirect "saisi0200_show.asp"
			response.end

		else
			'// 再試験該当学生を取得
			if wf_GetStudent() = false then
				m_bErrFlg = True
				m_sErrMsg = "学生情報の取得に失敗しました。"
				Exit Do
			end if
		end if

		Exit Do
	Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

	'//初期表示
    Call showPage()

    '// 終了処理
    Call gs_CloseDatabase()

End Sub

function wf_UpdateSeiseki()
'********************************************************************************
'*  [機能]  成績の判定、登録
'*  [引数]  なし
'*  [戻値]  true , false
'*  [説明]  
'********************************************************************************

	'変数の宣言
	Dim w_sSql
	Dim w_iRet
	Dim i
	Dim w_bNendo	'True:当年度　false:過去
	Dim w_iSeiseki
	Dim w_iGakusei
	Dim w_iNendo

	wf_UpdateSeiseki = false
	
	'// 評価基準を取得
	if wf_GetHyoka() = false then
		m_bErrFlg = True
		Exit Function
	end if

	for i = 1 to Request("hidCount")

		'変数の格納
		w_iGakusei = Request("hidGakusei" & i)
		w_iNendo   = Request("hidNendo" & i)
		'配列の格納
		w_iSeiseki = Request("txtSeiseki" & i)

'デバッグプリント用		
'Response.Write "番号＞" & Request("hidGakusei" & i) & "<br>"
'Response.Write "年度＞" & Request("hidNendo" & i) & "<br>"
'Response.Write "成績1＞" & Request("txtSeiseki" & i) & "<br>"

		'T16 & T120 の更新
		if wf_UpdateT16(w_iSeiseki,w_iGakusei,w_iNendo) = false then
			m_bErrFlg = True
			Exit Function
		end if

'Response.Write "-------------------------------------------------------------------------------------------------------------------<br>"
	next	
'response.end
	wf_UpdateSeiseki = true

end function

function wf_UpdateT16(p_iSeiseki,p_iGakusei,p_iNendo)
'********************************************************************************
'*  [機能]  T16の成績データ更新
'*  [引数]  p_iSeiseki :成績（入力値）
'*          p_iGakusei :学生番号
'*          p_iNendo   :年度
'*  [戻値]  true , false
'*  [説明]  
'********************************************************************************

	Dim i						'インデックス
	Dim w_sSql					'SQL格納エリア
	Dim w_Rs					'ADODBレコードセット
	Dim w_sHyokaMei				'M08_HYOKA_SYOBUNRUI_MEI
	Dim w_iHyotei				'M08_HYOTEI
	Dim w_iHyokaCd				'M08_HYOKA_SYOBUNRUI_RYAKU
	Dim w_sJugyoJikan(5)		'日数関連

	wf_UpdateT16 = false

	'**********************************
	'* 点数が入ってない場合何もしない *
	'**********************************
	if p_iSeiseki = "" then
		wf_UpdateT16 = true
		exit function
	end if

	'******************
	'* 基本データ取得 *
	'******************
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T16_KAISETU, "
	w_sSql = w_sSql & "		T16_HAITOTANI, "
	w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_Z, "
	w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K "
	w_sSql = w_sSql & "	FROM "
	w_sSql = w_sSql & "		T16_RISYU_KOJIN "
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T16_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T16_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T16_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

	Set w_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(w_Rs, w_sSql)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	'************
	'* 評価判定 *
	'************
	m_RsHyoka.MoveFirst
	w_sHyokaMei = ""
	Do until m_RsHyoka.EOF
		if Cint(p_iSeiseki) >= Cint(m_RsHyoka("M08_MIN")) then
			w_sHyokaMei = m_RsHyoka("M08_HYOKA_SYOBUNRUI_MEI")
			w_iHyotei   = m_RsHyoka("M08_HYOTEI")
			w_iHyokaCd  = m_RsHyoka("M08_HYOKA_SYOBUNRUI_RYAKU")
		end if
		m_RsHyoka.Movenext
	loop

	'******************
	'* 日数更新の準備 *
	'******************
	'前期の場合
	if Cint(w_Rs("T16_KAISETU")) = Cint(mC_KAISETU_Z) then
		if not isNull(w_Rs("T16_J_JUNJIKAN_KIMATU_Z")) and Cint(w_Rs("T16_J_JUNJIKAN_KIMATU_Z")) > 0 then
			w_sJugyoJikan(1) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sJugyoJikan(2) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sJugyoJikan(3) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sJugyoJikan(4) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
		else
			w_sJugyoJikan(1) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(2) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(3) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(4) = Cint(w_Rs("T16_HAITOTANI")) * 30
		end if
	'後期の場合
	else
		if not isNull(w_Rs("T16_J_JUNJIKAN_KIMATU_K")) and Cint(w_Rs("T16_J_JUNJIKAN_KIMATU_K")) > 0 then
			w_sJugyoJikan(1) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sJugyoJikan(2) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sJugyoJikan(3) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sJugyoJikan(4) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
		else
			w_sJugyoJikan(1) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(2) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(3) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(4) = Cint(w_Rs("T16_HAITOTANI")) * 30
		end if
	end if

	'***********
	'* T16更新 *
	'***********
	w_sSql = ""
	w_sSql = w_sSql & " UPDATE "
	w_sSql = w_sSql & "		T16_RISYU_KOJIN "
	w_sSql = w_sSql & " SET "

	'前期/後期場合わけ
	if Cint(w_Rs("T16_KAISETU")) = Cint(mC_KAISETU_Z) then
		w_sSql = w_sSql & "		T16_HTEN_KIMATU_Z  = " & p_iSeiseki & ", "
		w_sSql = w_sSql & "		T16_HYOKA_KIMATU_Z = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_Z = '" & w_sHyotei & "', "
		w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_Z = '" & f_GetNowDate() & "', "
		w_sSql = w_sSql & "		T16_HTEN_KIMATU_K  = " & p_iSeiseki & ", "
		w_sSql = w_sSql & "		T16_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "
		'合格の場合
		if Cint(w_iHyokaCd) = Cint(mC_HYOKA_CD_OK) then
			w_sSql = w_sSql & "		T16_HYOKA_FUKA_KBN = 0, "
			w_sSql = w_sSql & "		T16_TANI_SUMI = T16_HAITOTANI, "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_Z = " & w_sJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_Z = " & w_sJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_Z = " & w_sJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_Z = " & w_sJugyoJikan(4) & ", "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_K = " & w_sJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_K = " & w_sJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K = " & w_sJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_K = " & w_sJugyoJikan(4) & ", "
		end if
	else
		w_sSql = w_sSql & "		T16_HTEN_KIMATU_K = " & p_iSeiseki & ", "
		w_sSql = w_sSql & "		T16_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_K = '" & w_sHyotei & "', "
		w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "
		'合格の場合
		if Cint(w_iHyokaCd) = Cint(mC_HYOKA_CD_OK) then
			w_sSql = w_sSql & "		T16_HYOKA_FUKA_KBN = 0, "
			w_sSql = w_sSql & "		T16_TANI_SUMI = T16_HAITOTANI, "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_K = " & w_sJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_K = " & w_sJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K = " & w_sJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_K = " & w_sJugyoJikan(4) & ", "
		end if
	end if
	w_sSql = w_sSql & " 	T16_UPD_DATE = '" & f_GetNowDate() & "', "
	w_sSql = w_sSql & " 	T16_UPD_USER = '" & Session("LOGIN_ID") & "'"
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T16_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T16_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T16_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

'response.write w_sSql & "<br>"

	w_iRet = gf_ExecuteSQL(w_sSql)
'Response.Write w_sSql & "<br>"


	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	'**************************
	'* T120の成績データの更新 *
	'**************************
	if wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,w_sHyokaMei,w_iHyotei,w_iHyokaCd) = false then
		m_bErrFlg = True
		Exit Function
	end if

	w_Rs.close

	wf_UpdateT16 = true

end function

function wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,p_sHyokaMei,p_iHyotei,p_iHyokaCd)
'********************************************************************************
'*  [機能]  T120_SAISIKENの成績データ更新
'*  [引数]  なし
'*  [戻値]  true , false
'*  [説明]  
'********************************************************************************

	Dim w_sSql
	Dim w_iRet

	wf_UpdateT120 = false

	'**************
	'* T120の更新 *
	'**************
	w_sSql = ""
	w_sSql = w_sSql & " UPDATE "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "	SET "
	w_sSql = w_sSql & "		T120_HYOKA = '" & p_sHyokaMei & "', "
	w_sSql = w_sSql & "		T120_HYOTEI = " & p_iHyotei & ", "
	w_sSql = w_sSql & "		T120_SEISEKI = " & p_iSeiseki & ", "
	'合格の場合
	if Cint(p_iHyokaCd) = Cint(mC_HYOKA_CD_OK) then
		w_sSql = w_sSql & " 	T120_SYUTOKU_NENDO = " & Session("NENDO") & ", "
		w_sSql = w_sSql & " 	T120_SYUTOKU_FLG = 1, "
		w_sSql = w_sSql & " 	T120_HYOKA_FUKA_KBN = 0, "
	else
		w_sSql = w_sSql & " 	T120_HYOKA_FUKA_KBN = 1, "
	end if
	w_sSql = w_sSql & " 	T120_UPD_DATE = '" & f_GetNowDate() & "', "
	w_sSql = w_sSql & " 	T120_UPD_USER = '" & Session("LOGIN_ID") & "'"
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T120_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T120_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T120_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

	w_iRet = gf_ExecuteSQL(w_sSql)
'Response.Write w_sSql & "<br>"

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	wf_UpdateT120 = true

end function


function wf_GetHyoka()
'********************************************************************************
'*  [機能]  成績の判定、登録
'*  [引数]  なし
'*  [戻値]  true ,false
'*  [説明]  
'********************************************************************************

	'変数の宣言
	Dim w_sSql
	Dim w_iRet

	wf_GetHyoka = false

	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		M08_HYOKA_SYOBUNRUI_MEI, "
	w_sSql = w_sSql & "		M08_HYOTEI, "
	w_sSql = w_sSql & "		M08_HYOKA_SYOBUNRUI_RYAKU, "
	w_sSql = w_sSql & "		M08_MIN "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		M08_HYOKAKEISIKI "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & "		M08_NENDO = " & Session("NENDO") & " AND "
	w_sSql = w_sSql & "		M08_HYOKA_TAISYO_KBN = " & mC_HYOKA_KBN & " AND "
	w_sSql = w_sSql & "		M08_HYOUKA_NO = " & mC_HYOKA_NO
	w_sSql = w_sSql & " ORDER BY "
	w_sSql = w_sSql & "		M08_HYOKA_SYOBUNRUI_CD DESC "
	
	Set m_RsHyoka = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(m_RsHyoka, w_sSQL)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	wf_GetHyoka = true

end function

function wf_GetStudent()
'********************************************************************************
'*  [機能]  未修得学生取得
'*  [引数]  なし
'*  [戻値]  true ,false
'*  [説明]  
'********************************************************************************

	'変数の宣言
	Dim w_sSql
	Dim w_iRet

	wf_GetStudent = false
	
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	
	'画面に表示する項目
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_MISYU_GAKUNEN,  "
	w_sSql = w_sSql & "		M05_CLASS.M05_CLASSMEI, "
	w_sSql = w_sSql & "		T11_GAKUSEKI.T11_SIMEI, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_NENDO, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_JYUKOKAISU, "

	'Hidden項目
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_GAKUSEI_NO "
	
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN, "
	w_sSql = w_sSql & "		T11_GAKUSEKI, "
	w_sSql = w_sSql & "		T13_GAKU_NEN, "
	w_sSql = w_sSql & "		M05_CLASS "
	w_sSql = w_sSql & " WHERE "
	
	'TABLEの結合条件
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_NENDO = T13_GAKU_NEN.T13_NENDO AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_GAKUSEI_NO = T13_GAKU_NEN.T13_GAKUSEI_NO AND "
	w_sSql = w_sSql & "		T13_GAKU_NEN.T13_NENDO = M05_CLASS.M05_NENDO AND "
	w_sSql = w_sSql & "		T13_GAKU_NEN.T13_GAKUNEN = M05_CLASS.M05_GAKUNEN AND "
	w_sSql = w_sSql & "		T13_GAKU_NEN.T13_CLASS = M05_CLASS.M05_CLASSNO AND "
	
	'その他条件
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_NENDO = " & Session("NENDO") & " AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "' AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_KYOUKAN_CD = '" & Session("KYOKAN_CD") & "' AND "
'時数対応用（後で外す
'	w_sSql = w_sSql & "		T120_SAISIKEN.T120_HYOKA_FUKA_KBN <> 2 AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SEISEKI Is Null "

	w_sSql = w_sSql & "	ORDER BY"
	w_sSql = w_sSql & "		T13_GAKUNEN,"	
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "


	Set m_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

'Response.write gf_GetRsCount(m_Rs) & "<br>"

	wf_GetStudent = true
	
end function

function f_GetNowDate()
'-----------------------------------------------------------------
'	現在の日付を取得	戻り値：YYYY/MM/DD
'-----------------------------------------------------------------
	Dim wResult

	f_GetNowDate = ""

	wResult = gf_fmtZero(Year(Date()),4) & "/" & gf_fmtZero(Month(Date()),2) & "/" & gf_fmtZero(Day(Date()),2)

	f_GetNowDate = wResult

end function

sub showPage()
'********************************************************************************
'*  [機能]  HTMLの表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	'変数の宣言
	Dim w_iCount
	Dim w_sCellClass
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../common/style.css" type="text/css">
<title>新しいページ 1</title>

<script language="JavaScript">
<!--

//================================================
//	送信処理
//================================================
function jf_Submit() {

	if (!jf_CheckValue()) {
		return;
	}

	if (!confirm("追試験の成績を登録します。よろしいですか？")) {
		return;
	}

	document.frm.hidMode.value = "update";
	document.frm.action = "./saisi0200_toroku.asp";
	document.frm.target = "fMain";
	document.frm.submit();

}

//================================================
//	戻る処理
//================================================
function jf_Back() {

	location.href = "saisi0200_show.asp";
	return;

}
//================================================
//	値のチェック
//================================================
function jf_CheckValue() {

	var i				//インデックス
	var w_iValueCnt;	//生徒数格納エリア
	var w_oText;		//TEXTBOXオブジェクト格納エリア
	
	w_iValueCnt = Number(document.frm.hidCount.value);

	for (i=1;i<=w_iValueCnt;i++) {
		
		//オブジェクトの取得
		w_oText = eval("document.frm.txtSeiseki" + i);

		if (w_oText.value != "") {
			//数値型チェック
			if (isNaN(w_oText.value)) {
				alert("成績は整数で入力してください。");
				w_oText.focus();
				return false;
			}
			//100以下チェック
			if (Number(w_oText.value) > 100) {
				alert("成績が100を超えています。");
				w_oText.focus();
				return false;
			}
		}
	}
	
	return true;
}

//-->
</script>

</head>

<body>

<form name="frm" method="post">

<center>
<br>

<table border="1" class="hyo">
	<tr>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="16"><font color="#FFFFFF">履修学年</font></td>
		<td width="70"  class="CELL2"   height="16" align="center"><%=Request("hidMISYU_GAKUNEN")%></td>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="16"><font color="#FFFFFF">科　　目</font></td>
		<td width="200" class="CELL2"   height="16" align="center"><%=Request("hidKAMOKU_MEI")%></td>
	</tr>
</table>

<br>
<br>

<table border="1" class="hyo">

	<!-- TABLEヘッダ部 -->
	<tr>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">学年</font></td>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">クラス</font></td>
		<td width="200" class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">氏　　　名</font></td>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">年度</font></td>
		<!--<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">受験回数</font></td>-->
		<td width="40"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">成績</font></td>
	</tr>


	<!-- TABLEリスト部 -->
<%
	'カウンタの初期化
	w_iCount = 0
	
	'TDのClassの初期化
	w_sCellClass = "CELL2"

	do until m_Rs.EOF
		w_iCount = w_iCount + 1
%>
	<tr>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_MISYU_GAKUNEN")%><br></td>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("M05_CLASSMEI")%><br></td>
		<td width="200" class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T11_SIMEI")%><br></td>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_NENDO")%><br></td>
		<!--<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_JYUKOKAISU")%><br></td>-->
		<td width="40"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="hidden" name="hidGakusei<%=w_iCount%>"   value="<%=m_Rs("T120_GAKUSEI_NO")%>">
			<input type="hidden" name="hidNendo<%=w_iCount%>"     value="<%=m_Rs("T120_NENDO")%>">
			<input type="text"   name="txtSeiseki<%=w_iCount%>" size="3" style="ime-mode:disabled" maxlength="3">
		</td>
	</tr>

<%
		m_Rs.MoveNext
		
		if w_sCellClass = "CELL2" then
			w_sCellClass = "CELL1"
		else
			w_sCellClass = "CELL2"
		end if
		
	loop	
%>
</table>
<br>

<table>
	<tr>
		<td><input type="button" value=" 登　録 " onclick="jf_Submit();"></td>
		<td><input type="button" value=" 戻　る " onclick="jf_Back();"></td>
	</tr>
</table>

</center>

<!-- 引数 -->
<input type="hidden" name="hidCount" value="<%=w_iCount%>">
<input type="hidden" name="hidKAMOKU_CD" value="<%=Request("hidKAMOKU_CD")%>">
<input type="hidden" name="hidMode">

</form>

</body>

</html>
<%
end sub
%>