<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験期間教官予定一覧
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0180/skn0180_main.asp
' 機      能: MAINページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/23 伊藤
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="試験期間教官予定一覧"
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
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

		'// 権限チェックに使用
		session("PRJ_No") = "SKN0180"

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

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
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()

%>

<html>

<head>
<title>試験期間教官予定一覧</title>
</head>

<frameset rows="200,1,*" frameborder="0">
	<frame src="skn0180_top.asp" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
	<frame src="default2.asp" scrolling="auto" noresize name="main">
</frameset>

</html>
<% END SUB %>