<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 部活動部員一覧
' ﾌﾟﾛｸﾞﾗﾑID : web/web0360/web0360_edt.asp
' 機      能: 生徒の部活活動情報の更新
'-------------------------------------------------------------------------
' 引      数:   txtMode			:処理モード
'               txtClubCd		:部活CD
'               GAKUSEI_NO		:学生NO
'               cboGakunenCd	:学年
'               cboClassCd		:クラスNO
'               txtTyuClubCd	:中学校部活CD
'
' 引      渡:	txtClubCd		:部活CD
'               cboGakunenCd	:学年
'               cboClassCd		:クラスNO
'               txtTyuClubCd	:中学校部活CD
' 説      明:
'           ■生徒の部活活動の更新を行う
'-------------------------------------------------------------------------
' 作      成: 2001/08/22 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
	Public m_iSyoriNen			'//年度
	Public m_iKyokanCd			'//教官ｺｰﾄﾞ
	Public m_sClubCd			'//クラブCD
	Public m_iGakunen           '//学年
	Public m_iClassNo           '//クラスNO
	Public m_sTyuClubCd			'//中学校クラブCD
	Public m_sMode				'//処理モード

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

	Dim w_iRet			  '// 戻り値
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="部活動部員一覧"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"

	On Error Resume Next
	Err.Clear

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

		'//変数初期化
		Call s_ClearParam()

		'// MainﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()

'//デバッグ
'Call s_DebugPrint()

		'// 生徒の部活動の更新
		w_iRet = f_ClubUpdate()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'// ページを表示
		Call showPage()

		Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// 終了処理
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

	m_iSyoriNen  = ""
	m_iKyokanCd  = ""
	m_sClubCd    = ""
	m_iGakunen   = ""
	m_iClassNo   = ""
	m_sTyuClubCd = ""
	m_sMode      = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

	m_iSyoriNen  = Session("NENDO")
	m_iKyokanCd  = Session("KYOKAN_CD")
	m_sClubCd    = Request("txtClubCd")
	m_iGakunen   = Request("cboGakunenCd")	'//学年
	m_iClassNo   = Request("cboClassCd")	'//クラス
	m_sTyuClubCd = replace(Request("txtTyuClubCd"),"@@@","")	'//中学校クラブCD
	m_sMode      = Request("txtMode")	'//クラス

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iSyoriNen  = " & m_iSyoriNen  & "<br>"
	response.write "m_iKyokanCd  = " & m_iKyokanCd  & "<br>"
	response.write "m_sClubCd    = " & m_sClubCd	& "<br>"
	response.write "m_iGakunen   = " & m_iGakunen   & "<br>"
	response.write "m_iClassNo   = " & m_iClassNo   & "<br>"
	response.write "m_sTyuClubCd = " & m_sTyuClubCd & "<br>"

End Sub

'********************************************************************************
'*  [機能]  部員全員の入部日更新
'********************************************************************************
Function f_UpdNyububi()

	f_UpdNyububi = False
	On Error Resume Next
	Err.Clear

	Dim i
	Dim wFieldName

	w_sGakuseiNo   = split(replace(Request("hidGakuseiNo")," ",""),",")
	w_iGakusekiCnt = UBound(w_sGakuseiNo)
	wFieldName     = split(replace(Request("hidFieldName")," ",""),",")
	w_sNyububi     = split(replace(Request("txtNyububiC")," ",""),",")
	w_Taibubi      = split(replace(Request("txtTaibubi")," ",""),",")
	w_TaibuFlg     = split(replace(Request("hidTaibuFlg")," ",""),",")


'response.write Request("hidGakuseiNo") & "<BR>"
'response.write Request("hidFieldName") & "<BR>"
'response.write Request("txtNyububiC") & "<BR><BR>"

	i = 0
	Do Until i > w_iGakusekiCnt

		'//部活情報を更新
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN SET"

		if w_sNyububi(i) = "0000000000" then
			w_sSQL = w_sSQL & vbCrLf & "  T13_CLUB_" & wFieldName(i) & " = null"
			w_sSQL = w_sSQL & vbCrLf & " ,T13_CLUB_" & wFieldName(i) & "_NYUBI = null"
			w_sSQL = w_sSQL & vbCrLf & " ,T13_CLUB_" & wFieldName(i) & "_TAIBI = null"
			w_sSQL = w_sSQL & vbCrLf & " ,T13_CLUB_" & wFieldName(i) & "_FLG   = null"
		Else
			w_sSQL = w_sSQL & vbCrLf & " 	T13_CLUB_" & wFieldName(i) & "_NYUBI = '" & gf_YYYY_MM_DD(w_sNyububi(i),"/") & "'"
			if w_TaibuFlg(i) then
				w_sSQL = w_sSQL & vbCrLf & " 	,T13_CLUB_" & wFieldName(i) & "_TAIBI = '" & gf_YYYY_MM_DD(w_Taibubi(i),"/") & "'"
			End if
		End if

		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  	    T13_NENDO      =  " & cInt(m_iSyoriNen)
		w_sSQL = w_sSQL & vbCrLf & "  	AND T13_GAKUSEI_NO = '" & w_sGakuseiNo(i) & "'"
		w_sSQL = w_sSQL & vbCrLf & "  	AND T13_GAKUSEI_NO = '" & w_sGakuseiNo(i) & "'"

'response.write w_sSQL & "<BR>"
'response.write iRet & "<BR>"
		iRet = gf_ExecuteSQL(w_sSQL)
		If iRet <> 0 Then
'response.end
'response.write "rollllllllllllllllllllllllllllllllllllll" & "<BR>"
			'//ﾛｰﾙﾊﾞｯｸ
			Call gs_RollbackTrans()
			Exit Function
		End If

		i = i + 1
	Loop

'response.end

	f_UpdNyububi = True

End Function


'********************************************************************************
'*  [機能]  生徒の部活活動の更新
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_ClubUpdate()

	Dim w_sSQL
	Dim w_Rs
	Dim w_iKekka

	On Error Resume Next
	Err.Clear

	f_ClubUpdate = 1

	Do 

		'================
		'//学籍Noを取得
		'================
		w_sGakuseiNo = split(replace(Request("GAKUSEI_NO")," ",""),",")
		w_iGakusekiCnt = UBound(w_sGakuseiNo)

		'================
		'//入部日を取得
		'================
		w_sNyububi = split(replace(Request("hidNyububi")," ",""),",")

		'================
		'//退部日も取得
		'================
		w_sTaibubi = split(replace(Request("hidTaibubi")," ",""),",")

		'=================
		'//ﾄﾗﾝｻﾞｸｼｮﾝ開始
		'=================
		Call gs_BeginTrans()

		'// 部員全員の入部日更新
		if CStr(m_sMode) = "DELETE" then Call f_UpdNyububi()

		'====================================
		'//選択された生徒の人数分処理を実行
		'====================================
		For i=0 To w_iGakusekiCnt

			'//更新ﾌﾗｸﾞ初期化
			w_bClub1 = False
			w_bClub2 = False

			'=====================================
			'//現在の生徒のクラブ状況を取得
			'=====================================
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " SELECT "
			w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_1 "
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2"
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_TAIBI"
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_TAIBI"
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_FLG"
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_FLG"
			w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
			w_sSQL = w_sSQL & vbCrLf & " WHERE "
			w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen)
			w_sSQL = w_sSQL & vbCrLf & "  AND T13_GAKU_NEN.T13_GAKUSEI_NO='" & w_sGakuseiNo(i) & "'"

			iRet = gf_GetRecordset(rs, w_sSQL)
			If iRet <> 0 Then
				Call gs_RollbackTrans()
				'//ﾛｰﾙﾊﾞｯｸ
				'ﾚｺｰﾄﾞｾｯﾄの取得失敗
				f_ClubUpdate = 99
				Exit Do
			End If

			'//データありの時
			If rs.EOF = False Then

				'=====================================
				'//処理モードにより処理を振り分ける
				'=====================================
				Select Case m_sMode
					Case "INSERT"
						'//登録処理の場合（入部）
						Call f_InsDataSet(rs("T13_CLUB_1"),rs("T13_CLUB_2"),w_bClub1,w_bClub2,w_sClubCd,rs("T13_CLUB_1_FLG"),rs("T13_CLUB_2_FLG"),rs("T13_CLUB_1_TAIBI"),rs("T13_CLUB_2_TAIBI"))

					Case "DELETE"
						'//削除処理の場合（退部）
						Call f_DelDataSet(rs("T13_CLUB_1"),rs("T13_CLUB_2"),w_bClub1,w_bClub2,w_sClubCd)

					Case Else
						'//処理モード取得失敗
						m_sErrMsg = "処理モードがありません(システムエラー)"
						Exit Do

				End Select

			Else
				'//生徒情報がない事はないため、ここは通らない
				'//ﾛｰﾙﾊﾞｯｸ
				Call gs_RollbackTrans()
				m_sErrMsg = "データの更新に失敗しました。"
				Exit Do
			End If

			'================
			'//更新処理実行
			'================
			If w_bClub1 = True Or w_bClub2 = True Then
				'//部活情報を更新
				w_sSQL = ""
				w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN"
				w_sSQL = w_sSQL & vbCrLf & " SET"

				'//クラブ1を更新
				If w_bClub1 = True Then
					
					'入部の場合
					If m_sMode = "INSERT" Then	
						w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_1_FLG=1"	'1=入部
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_1='" & w_sClubCd & "'"
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_1_NYUBI='" & gf_YYYY_MM_DD(w_sNyububi(i),"/") & "'"
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_1_TAIBI=Null"

					'退部の場合
					Else
						w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_1_FLG=2"	'2=退部
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_1_TAIBI='" & gf_YYYY_MM_DD(w_sTaibubi(i),"/") & "'"
					End If

				End If

				'//クラブ2を更新
				If w_bClub2 = True Then
					'入部の場合
					If m_sMode = "INSERT" Then	
						w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_2_FLG=1"	'1=入部
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_2='" & w_sClubCd & "'"
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_2_NYUBI='" & gf_YYYY_MM_DD(w_sNyububi(i),"/") & "'"
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_2_TAIBI=Null"

					'退部の場合
					Else
						w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_2_FLG=2"	'2=退部
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_2_TAIBI='" & gf_YYYY_MM_DD(w_sTaibubi(i),"/") & "'"
					End If
				End If

				w_sSQL = w_sSQL & vbCrLf & " WHERE "
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen)
				w_sSQL = w_sSQL & vbCrLf & "  AND T13_GAKU_NEN.T13_GAKUSEI_NO='" & w_sGakuseiNo(i) & "'"
				'//更新処理実行
				iRet = gf_ExecuteSQL(w_sSQL)
				If iRet <> 0 Then
					'//ﾛｰﾙﾊﾞｯｸ
					Call gs_RollbackTrans()
					f_ClubUpdate = 99
					Exit Do
				End If

			End If

			'//ﾚｺｰﾄﾞｾｯﾄCLOSE
			Call gf_closeObject(rs)
		Next

		'//ｺﾐｯﾄ
		Call gs_CommitTrans()

		'//正常終了
		f_ClubUpdate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  登録時、クラブ1を更新するか、またはクラブ2を更新するかを調査する
'*  [引数]  p_T13_CLUB1:T13_CLUB_1
'*          p_T13_CLUB2:T13_CLUB_2
'*	[戻値]  p_bClub1=True : Club1更新可 p_bClub1=False : Club1更新不可
'*		    p_bClub2=True : Club2更新可 p_bClub2=False : Club2更新不可
'*          p_sClubCd:登録するCDを返す
'*  [説明]  
'********************************************************************************
Function f_InsDataSet(p_T13_CLUB1,p_T13_CLUB2,p_bClub1,p_bClub2,p_sClubCd,p_iFlg1,p_iFlg2,p_sTaiBi1,p_sTaiBi2)

		'//初期化
		p_bClub1 = False
		p_bClub2 = False
		p_sClubCd = ""

		'両方とも入部の場合、登録不可
		If gf_SetNull2String(p_iFlg1) = "1" And gf_SetNull2String(p_iFlg2) = "1" Then
'response.write "両方とも入部の場合、登録不可"
			Exit Function
		End If

		'//同一クラブで退部していた場合、クラブ1に登録
		If gf_SetNull2String(p_T13_CLUB1) = m_sClubCd And gf_SetNull2String(p_iFlg1) = "2" Then
'response.write "同一クラブで退部していた場合、クラブ1に登録"
			p_bClub1 = True
			p_bClub2 = False
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//同一クラブで退部していた場合、クラブ2に登録
		If gf_SetNull2String(p_T13_CLUB2) = m_sClubCd And gf_SetNull2String(p_iFlg2) = "2" Then
'response.write "同一クラブで退部していた場合、クラブ2に登録"
			p_bClub1 = False
			p_bClub2 = True
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//クラブ1が空きの場合、クラブ1に登録
		If gf_SetNull2String(p_T13_CLUB1) = "" Then
'response.write "クラブ1が空きの場合、クラブ1に登録"
			p_bClub1 = True
			p_bClub2 = False
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//クラブ2のみ空きの場合、クラブ2に登録
		If gf_SetNull2String(p_T13_CLUB2) = "" Then
'response.write "クラブ2のみ空きの場合、クラブ2に登録"
			p_bClub1 = False
			p_bClub2 = True
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//両方とも違うクラブでどちらか退部していた場合、クラブ1に登録
		If gf_SetNull2String(p_iFlg1) = "2" And gf_SetNull2String(p_iFlg2) <> "2" Then
'response.write "両方とも違うクラブでどちらか退部していた場合、クラブ1に登録"
			p_bClub1 = True
			p_bClub2 = False
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//両方とも違うクラブでどちらか退部していた場合、クラブ2に登録
		If gf_SetNull2String(p_iFlg2) = "2" And gf_SetNull2String(p_iFlg1) <> "2" Then
'response.write "両方とも違うクラブでどちらか退部していた場合、クラブ2に登録"
			p_bClub1 = False
			p_bClub2 = True
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//両方とも違うクラブで両方とも退部していた場合
		If gf_SetNull2String(p_iFlg1) = "2" And gf_SetNull2String(p_iFlg2) = "2" Then
'response.write "両方とも違うクラブで両方とも退部していた場合"

			'先に退部した方、クラブ1に登録
			If gf_SetNull2String(p_sTaiBi1) < gf_SetNull2String(p_sTaiBi2) Then
'response.write "先に退部した方、クラブ1に登録"
				p_bClub1 = True
				p_bClub2 = False
				p_sClubCd = m_sClubCd
				Exit Function
			End If

			'先に退部した方、クラブ2に登録
			If gf_SetNull2String(p_sTaiBi1) > gf_SetNull2String(p_sTaiBi2) Then
'response.write "先に退部した方、クラブ2に登録"
				p_bClub1 = False
				p_bClub2 = True
				p_sClubCd = m_sClubCd
				Exit Function
			End If
		End If
'response.write "エラー！！！！！！！！！！！！"

End Function

'********************************************************************************
'*  [機能]  削除時、削除対象クラブがクラブ1か、クラブ2かを調査する
'*  [引数]  p_T13_CLUB1:T13_CLUB_1
'*          p_T13_CLUB2:T13_CLUB_2
'*	[戻値]  p_bClub1=True : Club1更新可 p_bClub1=False : Club1更新不可
'*		    p_bClub2=True : Club2更新可 p_bClub2=False : Club2更新不可
'*          p_sClubCd:登録するCDを返す
'*  [説明]  
'********************************************************************************
Function f_DelDataSet(p_T13_CLUB1,p_T13_CLUB2,p_bClub1,p_bClub2,p_sClubCd)

		'//初期化
		p_bClub1 = False
		p_bClub2 = False
		p_sClubCd = ""

		'//同一クラブが退部対象（クラブ1）
		If gf_SetNull2String(p_T13_CLUB1) = m_sClubCd Then
			p_bClub1 = True
			p_bClub2 = False

			'クラブ名を削除せずに退部日と入部退部フラグ=2にする為に、クラブコードを返す。　2001/12/11 伊藤
			'p_sClubCd = ""
			p_sClubCd = m_sClubCd
		Else

			'//、同一クラブが退部対象（クラブ2）
			If gf_SetNull2String(p_T13_CLUB2) = m_sClubCd Then
				p_bClub1 = False
				p_bClub2 = True
				'クラブ名を削除せずに退部日と入部退部フラグ=2にする為に、クラブコードを返す。　2001/12/11 伊藤
				'p_sClubCd = ""
				p_sClubCd = m_sClubCd
			Else
				p_bClub1 = False
				p_bClub2 = False
			End If

		End If

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
	<html>
	<head>
	<title>部活動部員一覧</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>

	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//  [機能]  ページロード時処理
	//  [引数]
	//  [戻値]
	//  [説明]
	//************************************************************
	function window_onload() {

		<%If m_sMode = "DELETE" Then%>

			alert("<%= "登録が終了しました" %>");

			//退部登録時、初期画面に戻る
			//上フレーム再表示
			parent.topFrame.location.href="./web0360_top.asp?txtClubCd=<%=m_sClubCd%>"
			//下フレーム再表示
			parent.main.location.href="./web0360_main.asp?txtClubCd=<%=m_sClubCd%>"

		<%Else%>

			alert("<%= C_TOUROKU_OK_MSG %>");

			//新規登録時、登録画面に戻る
			var wArg
			wArg="?"
			wArg=wArg + "cboGakunenCd=<%=m_iGakunen%>"
			wArg=wArg + "&cboClassCd=<%=m_iClassNo%>"
			wArg=wArg + "&txtTyuClubCd=<%=m_sTyuClubCd%>"
			wArg=wArg + "&txtClubCd=<%=m_sClubCd%>"

			//上フレーム再表示
			parent.topFrame.location.href="./web0360_insTop.asp"+wArg;
			//下フレーム再表示
			parent.main.location.href="./web0360_insMain.asp"+wArg;
		<%End If%>

        return;
}

	//-->
	</SCRIPT>
	</head>
	<body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">

	<input type="hidden" name="txtClubCd" value="<%=m_sClubCd%>">

	</form>
	</body>
	</html>
<%
End Sub
%>