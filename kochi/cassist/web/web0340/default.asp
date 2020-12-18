<%@Language=VBScript %>
<%
'******************************************************************
'システム名     ：教務事務システム
'処　理　名     ：個人履修選択科目決定
'プログラムID   ：web/web0340/default.asp
'機　　　能     ：フレームページ 個人履修選択科目決定の表示を行う
'------------------------------------------------------------------
'引　　　数     ：
'変　　　数     ：
'引　　　渡     ：
'説　　　明     ：
'------------------------------------------------------------------
'作　　　成     ：2001.07.23    前田　智史
'変　　　更     ：
'
'******************************************************************
'*******************　ASP共通モジュール宣言　**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******　モ ジ ュ ー ル 変 数　********
Public m_iNendo
Public m_sKyokanCd
'******　メイン処理　********

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'******　Ｅ　Ｎ　Ｄ　********

Sub Main()
'******************************************************************
'機　　能：本ASPのﾒｲﾝﾙｰﾁﾝ
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    '******共通関数******
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="就職マスタ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iNendo = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")

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
		session("PRJ_No") = "WEB0340"

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '// 担任チェック
'	  If gf_Tannin(m_iNendo,m_sKyokanCd,1) <> 0 Then
'	            m_bErrFlg = True
'	            m_sErrMsg = "担任以外の入力はできません。"
'	            Exit Do
'	  End If


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
    <title>個人履修選択科目決定</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    </script>
    <frameset rows=170,1,* frameborder="0">
        <frame src="web0340_top.asp" scrolling="auto" noresize name="top">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        <frame src="default2.asp" scrolling="auto" noresize name="main">
    </frameset>
    </head>
</html>
<%
End Sub
%>