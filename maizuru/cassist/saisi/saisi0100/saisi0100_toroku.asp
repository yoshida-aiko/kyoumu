<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 再試成績登録
' ﾌﾟﾛｸﾞﾗﾑID : saisi/saisi0100/saisi0100_toroku.asp
' 機      能: 再試験成績の入力画面
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
' 変      更: 2003/03/03 矢野　不合格時にはT16,T17を更新しないように変更
'           : 2003/03/06 矢野　入力された値は全てT120に更新するように変更
' 変　　　更：2003/07/25 廣田　無限ループ & 登録で落ちていた！その他諸々・・・のため修正
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
Const mC_HYOKA_CD_NO = 1

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
			response.redirect "saisi0100_show.asp"
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
	Dim w_sSql			'SQL文格納エリア
	Dim w_iRet			'戻り値格納エリア
	Dim i				'インデックス
	Dim w_bNendo		'True:当年度　false:過去
	Dim w_iSeiseki(5)	'成績格納エリア
	Dim w_iGakusei		'学生番号
	Dim w_iNendo		'年度
	Dim w_iMotoTen		'元の点数

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
		w_iMotoTen = Request("hidSeiseki" & i)
		'配列の格納
		w_iSeiseki(1) = Request("txtSeiseki1_" & i)
		w_iSeiseki(2) = Request("txtSeiseki2_" & i)
		w_iSeiseki(3) = Request("txtSeiseki3_" & i)
		w_iSeiseki(4) = Request("txtSeiseki4_" & i)
		w_iSeiseki(5) = Request("txtSeiseki5_" & i)

'デバッグプリント用		
'Response.Write "番号＞" & Request("hidGakusei" & i) & "<br>"
'Response.Write "年度＞" & Request("hidNendo" & i) & "<br>"
'Response.Write "元点＞" & Request("hidSeiseki" & i) & "<br>"
'Response.Write "成績1＞" & Request("txtSeiseki1_" & i) & "<br>"
'Response.Write "成績2＞" & Request("txtSeiseki2_" & i) & "<br>"
'Response.Write "成績3＞" & Request("txtSeiseki3_" & i) & "<br>"
'Response.Write "成績4＞" & Request("txtSeiseki4_" & i) & "<br>"
'Response.Write "成績5＞" & Request("txtSeiseki5_" & i) & "<br>"
'response.end

		'当年度/過去年度の判定
		if Cint(Request("hidNendo" & i)) = Cint(Session("NENDO")) then
			w_bNendo = true
		else
			w_bNendo = false
		end if
'response.write w_bNendo & "<br>"

		'当年度科目の場合（T16 & T120）
		if w_bNendo = true then
			if wf_UpdateT16(w_iSeiseki,w_iGakusei,w_iNendo,w_iMotoTen) = false then
				m_bErrFlg = True
				Exit Function
			end if
		'過去年度科目の場合（T17 & T120）
		else
			if wf_UpdateT17(w_iSeiseki,w_iGakusei,w_iNendo,w_iMotoTen) = false then
				m_bErrFlg = True
				Exit Function
			end if
		end if

	next	

	wf_UpdateSeiseki = true

end function

function wf_UpdateT16(p_iSeiseki,p_iGakusei,p_iNendo,p_iMotoTen)
'********************************************************************************
'*  [機能]  T16の成績データ更新
'*  [引数]  p_iSeiseki	:点数（配列）
'*          p_iGakusei	:学生番号
'*          p_iNendo	:年度
'*          p_iMotoTen	:元の点数（T16_HTEN）
'*  [戻値]  true , false
'*  [説明]  
'********************************************************************************

	Dim i						'インデックス
	Dim w_sSql					'SQL格納エリア
	Dim w_Rs					'ADODBレコードセット
	Dim w_sHyokaMei				'M08_HYOKA_SYOBUNRUI_MEI
	Dim w_iHyotei				'M08_HYOTEI
	Dim w_iHyokaCd				'M08_HYOKA_SYOBUNRUI_RYAKU
	Dim w_iSeiseki				'成績（点数格納エリア）
	Dim w_sAryJugyoJikan(4)		'日数関連

	Dim w_bAryUpdateFlg(5)		'T120の履歴更新フラグ
	Dim w_sAryHyokaMei(5)		'T120の履歴のHYOKA
	Dim w_iAryHyotei(5)			'T120の履歴のHYOTEI

	wf_UpdateT16 = false

	'変数の初期化
	i = 5

	'*****************************************
	'* 更新対象回数の取得(T16に更新される値) *
	'*****************************************
	do until i = 0
		if i <> 1 then
			if p_iSeiseki(i) <> "" then
				w_iSeiseki = p_iSeiseki(i)
				exit do
			end if
		else
			if p_iSeiseki(i) <> "" then
				w_iSeiseki = p_iSeiseki(i)
				exit do
			else
				'入力がなかったのでココで終了
				wf_UpdateT16 = true
				Exit Function
			end if
		end if
		i = i - 1
	loop

	'**************************************
	'* 入力された値の判定及び、フラグ生成 *
	'**************************************
	'配列の初期化
	for i = 1 to 5
		w_bAryUpdateFlg(i) = false		'T120更新用のフラグ
		w_sAryHyokaMei(i)  = ""			'T120_HYOKAMEI
		w_iAryHyotei(i)    = 0			'T120_HYOTEI
	next

	'カウンタの設定
	i = 5

	do until i = 0

		'入力がある
		if p_iSeiseki(i) <> "" then
			'評価
			m_RsHyoka.MoveFirst
			Do until m_RsHyoka.EOF
				if Cint(p_iSeiseki(i)) >= Cint(m_RsHyoka("M08_MIN")) then
					w_sAryHyokaMei(i) = m_RsHyoka("M08_HYOKA_SYOBUNRUI_MEI")		'HYOKA格納
					w_iAryHyotei(i)   = m_RsHyoka("M08_HYOTEI")						'HYOTEI格納
					exit do
				end if
				m_RsHyoka.Movenext
			loop
			w_bAryUpdateFlg(i) = true
		end if
		i = i - 1
	loop

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
		if Cint(w_iSeiseki) >= Cint(m_RsHyoka("M08_MIN")) then
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
			w_sAryJugyoJikan(1) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sAryJugyoJikan(2) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sAryJugyoJikan(3) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sAryJugyoJikan(4) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
		else
			w_sAryJugyoJikan(1) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(2) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(3) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(4) = Cint(w_Rs("T16_HAITOTANI")) * 30
		end if
	'後期の場合
	else
		if not isNull(w_Rs("T16_J_JUNJIKAN_KIMATU_K")) and Cint(w_Rs("T16_J_JUNJIKAN_KIMATU_K")) > 0 then
			w_sAryJugyoJikan(1) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sAryJugyoJikan(2) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sAryJugyoJikan(3) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sAryJugyoJikan(4) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
		else
			w_sAryJugyoJikan(1) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(2) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(3) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(4) = Cint(w_Rs("T16_HAITOTANI")) * 30
		end if
	end if

	'合格の場合のみT16を更新
	if Cint(w_iHyokaCd) <> Cint(mC_HYOKA_CD_NO) then

		'***********
		'* T16更新 *
		'***********
		w_sSql = ""
		w_sSql = w_sSql & " UPDATE "
		w_sSql = w_sSql & "		T16_RISYU_KOJIN "
		w_sSql = w_sSql & " SET "

		'前期/後期場合わけ
		if Cint(w_Rs("T16_KAISETU")) = Cint(mC_KAISETU_Z) then
			w_sSql = w_sSql & "		T16_HTEN_KIMATU_Z  = " & w_iSeiseki & ", "
			w_sSql = w_sSql & "		T16_HYOKA_KIMATU_Z = '" & w_sHyokaMei & "', "
			w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_Z = " & w_iHyotei & ", "
			w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_Z = '" & f_GetNowDate() & "', "
			w_sSql = w_sSql & "		T16_HYOKA_FUKA_KBN = 0, "
			w_sSql = w_sSql & "		T16_TANI_SUMI = T16_HAITOTANI, "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_Z = " & w_sAryJugyoJikan(4) & ", "
			'後期の項目
			w_sSql = w_sSql & "		T16_HTEN_KIMATU_K = " & w_iSeiseki & ", "
			w_sSql = w_sSql & "		T16_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
			w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_K = " & w_iHyotei & ", "
			w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_K = " & w_sAryJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_K = " & w_sAryJugyoJikan(4) & ", "
		else
			w_sSql = w_sSql & "		T16_HTEN_KIMATU_K = " & w_iSeiseki & ", "
			w_sSql = w_sSql & "		T16_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
			w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_K = " & w_iHyotei & ", "
			w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "
			w_sSql = w_sSql & "		T16_HYOKA_FUKA_KBN = 0, "
			w_sSql = w_sSql & "		T16_TANI_SUMI = T16_HAITOTANI, "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_K = " & w_sAryJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_K = " & w_sAryJugyoJikan(4) & ", "
		end if

		w_sSql = w_sSql & " 	T16_UPD_DATE = '" & f_GetNowDate() & "', "
		w_sSql = w_sSql & " 	T16_UPD_USER = '" & Session("LOGIN_ID") & "'"
		w_sSql = w_sSql & "	WHERE "
		w_sSql = w_sSql & "		T16_NENDO = " & p_iNendo & " AND "
		w_sSql = w_sSql & "		T16_GAKUSEI_NO = '" & p_iGakusei & "' AND "
		w_sSql = w_sSql & "		T16_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"
'Response.Write w_sSql & "<br>"
'response.end
		'元の点数より高い場合のみT16更新
		if Cint(p_iMotoTen) < Cint(w_iSeiseki) then
			w_iRet = gf_ExecuteSQL(w_sSql)
'Response.Write w_sSql & "<br>"
		end if

		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Function
		End If

	end if

	'**************************
	'* T120の成績データの更新 *
	'**************************
	if wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,w_sAryHyokaMei,w_iAryHyotei,w_bAryUpdateFlg,w_iHyokaCd) = false then
		m_bErrFlg = True
		Exit Function
	end if

	w_Rs.close

	wf_UpdateT16 = true

end function

function wf_UpdateT17(p_iSeiseki,p_iGakusei,p_iNendo,p_iMotoTen)
'********************************************************************************
'*  [機能]  T17の成績データ更新
'*  [引数]  p_iSeiseki	:点数（配列）
'*          p_iGakusei	:学生番号
'*          p_iNendo	:年度
'*          p_iMotoTen	:元の点数（T16_HTEN）
'*  [戻値]  true , false
'*  [説明]  
'********************************************************************************
on error resume next

	Dim i						'インデックス
	Dim w_sSql					'SQL格納エリア
	Dim w_Rs					'ADODBレコードセット
	Dim w_sHyokaMei				'M08_HYOKA_SYOBUNRUI_MEI
	Dim w_iHyotei				'M08_HYOTEI
	Dim w_iHyokaCd				'M08_HYOKA_SYOBUNRUI_RYAKU
	Dim w_iSeiseki				'成績（点数格納エリア）
	Dim w_sAryJugyoJikan(4)		'日数関連

	Dim w_bAryUpdateFlg(5)		'T120の履歴更新フラグ
	Dim w_sAryHyokaMei(5)		'T120の履歴のHYOKA
	Dim w_iAryHyotei(5)			'T120の履歴のHYOTEI

	wf_UpdateT17 = false

	'変数の初期化
	i = 5

	'*****************************************
	'* 更新対象回数の取得(T17に更新される値) *
	'*****************************************
	do until i = 0
		if i <> 1 then
			if p_iSeiseki(i) <> "" then
				w_iSeiseki = p_iSeiseki(i)
				exit do
			end if
		else
			if p_iSeiseki(i) <> "" then
				w_iSeiseki = p_iSeiseki(i)
				exit do
			else
				'入力がなかったのでココで終了
				wf_UpdateT17 = true
				Exit Function
			end if
		end if
		i = i - 1
	loop

	'**************************************
	'* 入力された値の判定及び、フラグ生成 *
	'**************************************
	'配列の初期化
	for i = 1 to 5
		w_bAryUpdateFlg(i) = false		'T120更新用のフラグ
		w_sAryHyokaMei(i)  = ""			'T120_HYOKAMEI
		w_iAryHyotei(i)    = 0			'T120_HYOTEI
	next

	'カウンタの設定
	i = 5

	do until i = 0

		'入力がある
		if p_iSeiseki(i) <> "" then
			'評価
			m_RsHyoka.MoveFirst
			Do until m_RsHyoka.EOF
				if Cint(p_iSeiseki(i)) >= Cint(m_RsHyoka("M08_MIN")) then
					w_sAryHyokaMei(i) = m_RsHyoka("M08_HYOKA_SYOBUNRUI_MEI")		'HYOKA格納
					w_iAryHyotei(i)   = m_RsHyoka("M08_HYOTEI")						'HYOTEI格納
					exit do
				end if
				m_RsHyoka.Movenext
			loop

			w_bAryUpdateFlg(i) = true
		end if
		i = i - 1
	loop

	'******************
	'* 基本データ取得 *
	'******************
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T17_KAISETU "
	w_sSql = w_sSql & "		T17_HAITOTANI, "
	w_sSql = w_sSql & "		T17_J_JUNJIKAN_KIMATU_K "
	w_sSql = w_sSql & "	FROM "
	w_sSql = w_sSql & "		T17_RISYUKAKO_KOJIN "
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T17_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T17_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T17_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

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
		if Cint(w_iSeiseki) >= Cint(m_RsHyoka("M08_MIN")) then
			w_sHyokaMei = m_RsHyoka("M08_HYOKA_SYOBUNRUI_MEI")
			w_iHyotei   = m_RsHyoka("M08_HYOTEI")
			w_iHyokaCd  = m_RsHyoka("M08_HYOKA_SYOBUNRUI_RYAKU")
			exit do
		end if
		m_RsHyoka.Movenext
	loop

	'******************
	'* 日数更新の準備 *
	'******************
	'後期のみ
	if not isNull(w_Rs("T17_J_JUNJIKAN_KIMATU_K")) and Cint(w_Rs("T17_J_JUNJIKAN_KIMATU_K")) > 0 then
		w_sAryJugyoJikan(1) = w_Rs("T17_J_JUNJIKAN_KIMATU_K")
		w_sAryJugyoJikan(2) = w_Rs("T17_J_JUNJIKAN_KIMATU_K")
		w_sAryJugyoJikan(3) = w_Rs("T17_J_JUNJIKAN_KIMATU_K")
		w_sAryJugyoJikan(4) = w_Rs("T17_J_JUNJIKAN_KIMATU_K")
	else
		w_sAryJugyoJikan(1) = Cint(w_Rs("T17_HAITOTANI")) * 30
		w_sAryJugyoJikan(2) = Cint(w_Rs("T17_HAITOTANI")) * 30
		w_sAryJugyoJikan(3) = Cint(w_Rs("T17_HAITOTANI")) * 30
		w_sAryJugyoJikan(4) = Cint(w_Rs("T17_HAITOTANI")) * 30
	end if

	'合格の場合
	if Cint(w_iHyokaCd) <> Cint(mC_HYOKA_CD_NO) then

		'***********
		'* T17更新 *
		'***********
		w_sSql = ""
		w_sSql = w_sSql & " UPDATE "
		w_sSql = w_sSql & "		T17_RISYUKAKO_KOJIN "
		w_sSql = w_sSql & " SET "
		'2003.09.02 前期と後期の成績を同時に更新するように修正 ITO
		w_sSql = w_sSql & "		T17_HTEN_KIMATU_Z  = " & w_iSeiseki & ", "
		w_sSql = w_sSql & "		T17_HYOKA_KIMATU_Z = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T17_HYOTEI_KIMATU_Z = " & w_iHyotei & ", "
		w_sSql = w_sSql & "		T17_KOUSINBI_KIMATU_Z = '" & f_GetNowDate() & "', "

		w_sSql = w_sSql & "		T17_HTEN_KIMATU_K  = " & w_iSeiseki & ", "
		w_sSql = w_sSql & "		T17_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T17_HYOTEI_KIMATU_K = " & w_iHyotei & ", "
		w_sSql = w_sSql & "		T17_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "

		w_sSql = w_sSql & "		T17_HYOKA_FUKA_KBN = 0, "
		w_sSql = w_sSql & "		T17_TANI_SUMI = T17_HAITOTANI, "

		'2003.09.02 前期と後期の成績を同時に更新するように修正 ITO
		w_sSql = w_sSql & "		T17_SOJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(1) & ", "
		w_sSql = w_sSql & "		T17_JUNJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(2) & ", "
		w_sSql = w_sSql & "		T17_J_JUNJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(3) & ", "
		w_sSql = w_sSql & "		T17_KEKA_KIMATU_Z = " & w_sAryJugyoJikan(4) & ", "

		w_sSql = w_sSql & "		T17_SOJIKAN_KIMATU_K = " & w_sAryJugyoJikan(1) & ", "
		w_sSql = w_sSql & "		T17_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(2) & ", "
		w_sSql = w_sSql & "		T17_J_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(3) & ", "
		w_sSql = w_sSql & "		T17_KEKA_KIMATU_Z = " & w_sAryJugyoJikan(4) & ", "

		w_sSql = w_sSql & " 	T17_UPD_DATE = '" & f_GetNowDate() & "', "
		w_sSql = w_sSql & " 	T17_UPD_USER = '" & Session("LOGIN_ID") & "'"
		w_sSql = w_sSql & "	WHERE "
		w_sSql = w_sSql & "		T17_NENDO = " & p_iNendo & " AND "
		w_sSql = w_sSql & "		T17_GAKUSEI_NO = '" & p_iGakusei & "' AND "
		w_sSql = w_sSql & "		T17_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

		'もとの点数より高い場合のみ更新
		if Cint(p_iMotoTen) < Cint(w_iSeiseki) then
			w_iRet = gf_ExecuteSQL(w_sSql)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				Exit Function
			End If
		end if

	end if

	'**************************
	'* T120の成績データの更新 *
	'**************************
	if wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,w_sAryHyokaMei,w_iAryHyotei,w_bAryUpdateFlg,w_iHyokaCd) = false then
		m_bErrFlg = True
		Exit Function
	end if

	w_Rs.close

	wf_UpdateT17 = true

end function

function wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,p_sAryHyokaMei,p_iAryHyotei,p_bAryUpdateFlg,p_iHyokaCd)
'********************************************************************************
'*  [機能]  T120_SAISIKENの成績データ更新
'*  [引数]  p_iSeiseki		:点数（配列）
'*          p_iGakusei		:学生番号
'*          p_iNendo		:年度
'*          p_sAryHyokaMei	:評価名称（配列）
'*          p_iAryHyotei	:評定（配列）
'*          p_bAryUpdateFlg	:更新フラグ（配列）	< True で更新
'*          p_iHyokaCd		:評価コード			< 合否判定に使用
'*  [戻値]  true , false
'*  [説明]  
'********************************************************************************

	Dim i			'インデックス
	Dim w_sSql		'SQL文格納エリア
	Dim w_iRet		'戻り値格納エリア

	wf_UpdateT120 = false

	'**************
	'* T120の更新 *
	'**************
	w_sSql = ""
	w_sSql = w_sSql & " UPDATE "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "	SET "
	'更新フラグ
	for i = 1 to 5

'		if p_bAryUpdateFlg(i) = true then
			if i = 1 then
				w_sSql = w_sSql & " 	T120_SAISI_SEISEKI " & f_IntUpdSQL(p_iSeiseki(i)) & ","
				w_sSql = w_sSql & " 	T120_HYOKA         " & f_StrUpdSQL(p_sAryHyokaMei(i)) & ","
				w_sSql = w_sSql & " 	T120_HYOTEI        " & f_NumUpdSQL(p_iAryHyotei(i)) & ","
			else
				w_sSql = w_sSql & " 	T120_SAISI_SEISEKI" & i & f_IntUpdSQL(p_iSeiseki(i)) & ","
				w_sSql = w_sSql & " 	T120_HYOKA"  & i & f_StrUpdSQL(p_sAryHyokaMei(i))  & ","
				w_sSql = w_sSql & " 	T120_HYOTEI"  & i & f_NumUpdSQL(p_iAryHyotei(i)) & ","
			end if
'		end if

	next

	'合格の場合
	if Cint(p_iHyokaCd) = Cint(mC_HYOKA_CD_OK) then
		w_sSql = w_sSql & " 	T120_SYUTOKU_NENDO = " & Session("NENDO") & ", "
		w_sSql = w_sSql & " 	T120_SYUTOKU_FLG = 1, "
		w_sSql = w_sSql & " 	T120_HYOKA_FUKA_KBN = 0, "
	else
		w_sSql = w_sSql & " 	T120_SYUTOKU_NENDO = NULL, "
		w_sSql = w_sSql & " 	T120_SYUTOKU_FLG = 0, "
		w_sSql = w_sSql & " 	T120_HYOKA_FUKA_KBN = 1, "
	end if
	w_sSql = w_sSql & " 	T120_UPD_DATE = '" & f_GetNowDate() & "', "
	w_sSql = w_sSql & " 	T120_UPD_USER = '" & Session("LOGIN_ID") & "'"
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T120_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T120_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T120_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

	w_iRet = gf_ExecuteSQL(w_sSql)

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
	w_sSql = w_sSql & "		M08_HYOKA_SYOBUNRUI_CD ASC "

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
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SEISEKI, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI2, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI3, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI4, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI5, "

	'Hidden項目
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_GAKUSEI_NO "
	
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN, "
	w_sSql = w_sSql & "		T11_GAKUSEKI, "
	w_sSql = w_sSql & "		T13_GAKU_NEN, "
	w_sSql = w_sSql & "		M05_CLASS "
'	w_sSql = w_sSql & "		M08_HYOKAKEISIKI "
	w_sSql = w_sSql & " WHERE "
	
	'TABLEの結合条件
	w_sSql = w_sSql & "			T120_SAISIKEN.T120_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO "
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_NENDO = T13_GAKU_NEN.T13_NENDO "
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_GAKUSEI_NO = T13_GAKU_NEN.T13_GAKUSEI_NO "
	w_sSql = w_sSql & "		AND T13_GAKU_NEN.T13_NENDO = M05_CLASS.M05_NENDO "
	w_sSql = w_sSql & "		AND T13_GAKU_NEN.T13_GAKUNEN = M05_CLASS.M05_GAKUNEN "
	w_sSql = w_sSql & "		AND T13_GAKU_NEN.T13_CLASS = M05_CLASS.M05_CLASSNO "
'	w_sSql = w_sSql & "		AND M08_NENDO = T120_NENDO"			'履修年度の評価
'	w_sSql = w_sSql & " 	AND M08_HYOUKA_NO = 2 "
'	w_sSql = w_sSql & "		AND M08_HYOKA_TAISYO_KBN = 0 "
'	w_sSql = w_sSql & "		AND M08_HYOKA_SYOBUNRUI_CD = 4 "
	'その他条件
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "' "
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_KYOUKAN_CD = '" & Session("KYOKAN_CD") & "' "
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_MISYU_GAKUNEN = " & Request("hidMISYU_GAKUNEN") 
'時数対応用（後で外す
	w_sSql = w_sSql & " 	AND ( T120_SYUTOKU_NENDO Is Null or T120_SYUTOKU_NENDO = " & Session("NENDO") & " ) "
'	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_HYOKA_FUKA_KBN <> 2 "
	w_sSql = w_sSql & " 	AND NOT T120_SEISEKI Is Null "
	w_sSql = w_sSql & " 	AND T120_TAISYO_FLG = 1 "
'	w_sSql = w_sSql & " 	AND T120_SEISEKI <= M08_MAX "
'	w_sSql = w_sSql & " 	AND T120_SEISEKI >= M08_MIN "

	w_sSql = w_sSql & "	ORDER BY"
	w_sSql = w_sSql & "		T13_GAKUNEN,"	
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "

'response.write w_sSql & "<br>"

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
'--------------------------------------------------------------------------------
'-  [機能]  現在の日付を取得
'-  [引数]  なし
'-  [戻値]  YYYY/MM/DD
'-  [説明]  
'--------------------------------------------------------------------------------
	Dim wResult

	f_GetNowDate = ""

	wResult = gf_fmtZero(Year(Date()),4) & "/" & gf_fmtZero(Month(Date()),2) & "/" & gf_fmtZero(Day(Date()),2)

	f_GetNowDate = wResult

end function

'/*******************************************************************************
'/*          UPDATE SQL 数値型データの作成
'/*                  pv_string　　update SQL データ
'/*
'/*               戻り値      =データ
'/********************************************************************************
Function f_IntUpdSQL(pString)

'***変数宣言***
    Dim ls_String
    Dim ls_Length

'***変換処理***
    '文字列中に'があったら、''に変換する
    pString = Replace(pString, "'", "''")
  
    If gf_IsNull(pString) Then              'NULLの時
        f_IntUpdSQL = " = NULL "
    Else
        ls_String = pString
        f_IntUpdSQL = " = " & pString
    End If

End Function

'/*******************************************************************************
'/*          UPDATE SQL 文字列データの作成
'/*                  pv_string　　update SQL データ
'/*
'/*               戻り値      ='文字列データ'
'/********************************************************************************
Function f_StrUpdSQL(pString)

'***変数宣言***
    Dim ls_String
    Dim ls_Length

'***変換処理***
    '文字列中に'があったら、''に変換する
    pString = Replace(pString, "'", "''")
    
    If gf_IsNull(pString) Then              'NULLの時
        f_StrUpdSQL = " = NULL "
    Else
        ls_String = pString
        f_StrUpdSQL = " = '" & pString & "'"
    End If

End Function

'/*******************************************************************************
'/*          UPDATE SQL 文字列データの作成
'/*                  pv_string　　update SQL データ
'/*
'/*               戻り値      ='文字列データ'
'/********************************************************************************
Function f_NumUpdSQL(pString)

'***変数宣言***
    Dim ls_String
    Dim ls_Length

'***変換処理***
    '文字列中に'があったら、''に変換する
    pString = Replace(pString, "'", "''")
    
    If gf_IsNull(pString) Then              'NULLの時
        f_NumUpdSQL = " = NULL "
    ElseIf pString = 0 Then
		f_NumUpdSQL = " = NULL "
	Else
        ls_String = pString
        f_NumUpdSQL = " = " & pString
    End If

End Function

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

	if (!confirm("再試験の成績を登録します。よろしいですか？")) {
		return;
	}

	document.frm.hidMode.value = "update";
	document.frm.action = "./saisi0100_toroku.asp";
	document.frm.target = "fMain";
	document.frm.submit();

}

//================================================
//	戻る処理
//================================================
function jf_Back() {

	location.href = "saisi0100_show.asp";
	return;

}
//================================================
//	値のチェック
//================================================
function jf_CheckValue() {

	var i,j				//インデックス
	var w_iValueCnt;	//生徒数格納エリア
	var w_oText;		//TEXTBOXオブジェクト格納エリア
	
	w_iValueCnt = Number(document.frm.hidCount.value);

	for (i=1;i<=w_iValueCnt;i++) {
		
		for (j=1;j<=5;j++) {

			//オブジェクトの取得
			w_oText = eval("document.frm.txtSeiseki" + j + "_" + i);

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
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">受験回数</font></td>
		<td width="40"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">成績</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">1回</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">2回</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">3回</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">4回</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">5回</font></td>
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
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_JYUKOKAISU")%><br></td>
		<td width="40"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="hidden" name="hidGakusei<%=w_iCount%>"  value="<%=m_Rs("T120_GAKUSEI_NO")%>">
			<input type="hidden" name="hidNendo<%=w_iCount%>"    value="<%=m_Rs("T120_NENDO")%>">
			<input type="hidden" name="hidSeiseki<%=w_iCount%>"    value="<%=m_Rs("T120_SEISEKI")%>">
			<%=m_Rs("T120_SEISEKI")%><br>
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki1_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI")%>"  size="3" style="ime-mode:disabled" maxlength="3">
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki2_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI2")%>" size="3" style="ime-mode:disabled" maxlength="3">
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki3_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI3")%>" size="3" style="ime-mode:disabled" maxlength="3">
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki4_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI4")%>" size="3" style="ime-mode:disabled" maxlength="3">
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki5_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI5")%>" size="3" style="ime-mode:disabled" maxlength="3">
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