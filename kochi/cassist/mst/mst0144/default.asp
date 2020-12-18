<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ登録
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0144/default.asp
' 機      能: フレームページ 就職先マスタの登録を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/06/22 岩下　幸一郎
' 変      更: 2001/07/13 谷脇　良也
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    Public  m_sSinroCD2
    Public  m_sSingakuCD2
    Public  m_sNendo
    Public  m_sSyusyokuName
    Public  m_sMode
    Public  m_skubun
    Public  m_sPageCD
    Public  sArg
    Public  sArg_top

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
    w_sMsgTitle="進路先情報登録"
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
        session("PRJ_No") = "MST0144"

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

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

<title>就職先マスタ登録</title>

<frameset rows=180,1,* frameborder="0" framespacing="0">

<frame src="top.asp?<%=request.form.item%>" scrolling="auto"  name="top2" noresize>
<frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
<frame src="main.asp?<%=request.form.item%>" scrolling="auto"  name="main" noresize>

</frameset>

</head>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
