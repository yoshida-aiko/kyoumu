<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 個人履修選択科目決定
' ﾌﾟﾛｸﾞﾗﾑID : web/web0340/web0340_main.asp
' 機	  能: 下ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引	  数:教官コード 	＞		SESSION("KYOKAN_CD")
'			 年度			＞		SESSION("NENDO")
' 変	  数:
' 引	  渡:
' 説	  明:
'-------------------------------------------------------------------------
' 作	  成: 2001/07/25 前田
' 変	  更: 2001/08/28 伊藤公子 ヘッダ部切り離し対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ

	Dim 	m_sKyokanCd 	'//教官コード
	Dim 	m_sGakunen		'//学年
	Dim 	m_sClass		'//クラス
	Dim 	m_sKBN			'//区分
	Dim 	m_sGRP			'//グループ区分

'///////////////////////////メイン処理/////////////////////////////

	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main

'///////////////////////////　ＥＮＤ　/////////////////////////////

Sub Main()
'********************************************************************************
'*	[機能]	本ASPのﾒｲﾝﾙｰﾁﾝ
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]
'********************************************************************************

	Dim w_iRet				'// 戻り値
	Dim w_sSQL				'// SQL文
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle = "キャンパスアシスト"
	w_sMsgTitle = "個人履修選択科目決定"
	w_sMsg = ""
	w_sRetURL = "../../login/default.asp"
	w_sTarget = "_top"

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

		'// 権限チェックに使用
		SESSION("PRJ_No") = "WEB0340"

		'// 不正アクセスチェック
		Call gf_userChk(SESSION("PRJ_No"))
'
'		 '// 担任チェック
'		If gf_Tannin(session("NENDO"),session("KYOKAN_CD"),1) <> 0 Then
'			m_bErrFlg = True
'			m_sErrMsg = "担任以外の入力はできません。"
'			Exit Do
'		End If

		'2001/12/01 Modd ---->
'		'// ページを表示
'		Call showPage()

		Call s_GetParam 		'渡された引数を取得

		'担当しているかどうかをチェック
		If f_chkTantoKyokan = True Then
			'// 担当している場合、詳細ページを表示
			Call showPage
		Else
			'// 担当していない場合、エラーページを表示
			Call showErrPage
		End If
		'2001/12/01 Add <----

		Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// 終了処理
	Call gs_CloseDatabase
End Sub

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
<title>個人履修選択科目決定</title>
</head>

<frameset rows="138px,1px,*" frameborder="no">
	<frame src="white.asp?txtMsg=<%=Request("txtMsg")%>" scrolling="yes" noresize name="middle">
	<frame src="../../common/bar.html" scrolling="no" noresize name="bar">
	<frame src="web0340_bottom.asp?<%=Server.HTMLEncode(request.form.item)%>" scrolling="yes" noresize name="bottom">
</frameset>

</html>
<%
End Sub

'2001/12/01 Add ---->

Sub s_GetParam()
'********************************************************************************
'*	[機能]　パラメータ取得
'*	[引数]　 なし
'*	[戻値]　なし
'*	[説明]
'********************************************************************************

	m_sKyokanCd = SESSION("KYOKAN_CD")
	m_sGakunen = request("txtGakunen")
	m_sClass = request("txtClass")
	m_sKBN = Cint(request("txtKBN"))
	m_sGRP = Cint(request("txtGRP"))

End Sub

Function f_chkTantoKyokan()
'********************************************************************************
'*	[機能]　担当教官かどうかチェック
'*	[引数]　 なし
'*	[戻値]　True:担当教官をしている、False:担当教官をしていない
'*	[説明]
'********************************************************************************
	Dim w_iRet			'戻り値
	Dim w_sSQL			'SQL
	Dim w_oRecord		'レコード

	f_chkTantoKyokan = False

	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "	T27_KYOKAN_CD"
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "	T27_TANTO_KYOKAN,"
	w_sSQL = w_sSQL & vbCrLf & "	T16_RISYU_KOJIN "
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "	T27_NENDO      = T16_NENDO AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_KAMOKU_CD  = T16_KAMOKU_CD AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_GAKUNEN    = T16_HAITOGAKUNEN AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_NENDO      = " & SESSION("NENDO") & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_GAKUNEN    = " & m_sGakunen & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_CLASS      = " & m_sClass & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T27_KYOKAN_CD  = '" & m_sKyokanCd & "' AND "
	w_sSQL = w_sSQL & vbCrLf & "	T16_HISSEN_KBN = " & C_HISSEN_SEN & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T16_SELECT_FLG = " & C_SENTAKU_YES & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T16_KAMOKU_KBN = " & m_sKBN & " AND "
	w_sSQL = w_sSQL & vbCrLf & "	T16_GRP        = " & m_sGRP

	Set w_oRecord = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset_OpenStatic(w_oRecord, w_sSQL)
'response.write(w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit Function
	End If

	'//担当していない場合
	If w_oRecord.EOF = True Then
		Exit Function
	End If

	w_oRecord.Close
	Set w_oRecord = Nothing

	f_chkTantoKyokan = True

End Function

Sub showErrPage()
'********************************************************************************
'*	[機能]	Htmlを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]
'********************************************************************************
%>
<Html>
<head>
<Title>個人履修選択科目決定エラーページ</Title>
<link rel=stylesheet href=../../common/style.css type=text/css>
</head>
<Body>
	<center>
		<br><br><br>
		<span class="msg">担当している科目はありません。</span>
	</center>
</Body>
</Html>
<%
End Sub

'2001/12/01 Add <----

%>

