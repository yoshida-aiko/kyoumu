<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 連絡掲示板
' ﾌﾟﾛｸﾞﾗﾑID : web/web0330/default.asp
' 機      能: フレームページ 連絡掲示板を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/07/10 前田
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_stxtMode
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
    w_sMsgTitle="連絡掲示板"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_stxtMode = request("txtMode")

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
		session("PRJ_No") = "WEB0330"

		'session("NENDO") = 2001

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        If m_stxtMode = "NEW" or m_sTxtMode = "UPD" Then
            '// ページを表示
            Call TOUROKU_showpage()
            Exit Do
        ElseIf m_stxtMode = "" Then
            '// ページを表示
            Call showPage()
            Exit Do
        End If

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

<title>連絡掲示板</title>

<frameset rows="110,1,*" frameborder="no">
    <frame src="web0330_top.asp" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
    <frame src="web0330_main.asp" scrolling="auto" noresize name="main">
</frameset>

</head>

</html>
<%
End Sub

Sub TOUROKU_showpage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_sKenmei
Dim w_sNaiyou
Dim w_sKaisibi
Dim w_sSyuryoubi
Dim w_stxtMode
Dim w_sNendo
Dim w_sKyokanCd
Dim w_stxtNo

    w_sKenmei    = request("Kenmei")
    w_sNaiyou    = request("Naiyou")
    w_sKaisibi   = request("Kaisibi")
    w_sSyuryoubi = request("Syuryoubi")
    w_stxtMode   = request("txtMode")
    w_sNendo     = request("txtNendo")
    w_sKyokanCd  = request("txtKyokanCd")
    w_stxtNo     = request("txtNo")

        sArg = ""
        sArg = sArg & "txtKenmei=" & Server.URLEncode(w_sKenmei)
        sArg = sArg & "&txtNaiyou=" & Server.URLEncode(w_sNaiyou)
        sArg = sArg & "&txtKaisibi=" & w_sKaisibi 
        sArg = sArg & "&txtSyuryoubi=" & w_sSyuryoubi 
        sArg = sArg & "&txtMode=" & w_stxtMode 
        sArg = sArg & "&txtNendo=" & w_sNendo 
        sArg = sArg & "&txtKyokanCd=" & w_sKyokanCd 
        sArg = sArg & "&txtNo=" & w_stxtNo 

%>
<html>

<head>

<title>連絡掲示板</title>

<frameset rows=250,1,* frameborder="no">
    <frame src="sousin_top.asp?<%=sArg %>" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
    <frame src="default2.asp?<%=sArg %>" scrolling="auto" noresize name="main">
</frameset>

</head>

</html>
<%
End Sub
%>