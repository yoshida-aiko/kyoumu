<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生一覧
' ﾌﾟﾛｸﾞﾗﾑID : saisi/saisi×××/default.asp
' 機      能: 担任がうけもつ生徒の一覧を参照する
'-------------------------------------------------------------------------
' 引      数:
' 変      数:
' 引      渡:
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2003/02/24
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

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

	Dim w_iRet              '// 戻り値
	Dim w_sSQL              '// SQL文
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle = "キャンパスアシスト"
	w_sMsgTitle = "不合格学生一覧"
	w_sMsg = ""
	w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget = "fTopMain"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK '(権限チェックをしない)

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'// 初期ページを表示
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

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
	dim w_item
	'---------- HTML START  ----------
	On Error Resume Next
	Err.Clear
%>

<html>
<head>
<title><%= w_sMsgTitle %></title>
</head>

<frameset rows=120,1,* frameborder="no">
	<frame src="saisi0500_top2.asp" scrolling="auto" name="_TOP" noresize>
	<frame src="../../common/bar.html" scrolling="auto" name="bar" noresize>
	<frame src="saisi0500_lower.asp?mode=new" scrolling="auto" name="_LOWER" noresize>
</frameset>

</html>

<% End Sub %>