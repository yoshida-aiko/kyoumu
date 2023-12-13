<%@ Language=VBScript %>
<%Response.Expires = 0%>
<%Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 部活動部員一覧
' ﾌﾟﾛｸﾞﾗﾑID : web/web0360/default.asp
' 機      能: フレームページ 
'-------------------------------------------------------------------------
' 引      数:SESSION(""):教官コード     ＞      SESSIONより
' 変      数:なし
' 引      渡:SESSION(""):教官コード     ＞      SESSIONより
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/08/22 伊藤
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

		'// 権限チェックに使用
		session("PRJ_No") = "WEB0360"
		'session("PRJ_No") = "WEB0300"

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
<title></title>
<frameset rows=110,1,* frameborder="no">
    <frame src="web0360_top.asp" scrolling="auto" noresize name="topFrame">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
    <frame src="default2.asp" scrolling="auto" noresize name="main">
    <!--<frame src="web0360_main.asp" scrolling="auto" noresize name="main">-->
</frameset>
</head>
</html>
<%
End Sub
%>