<%@Language=VBScript %>
<%
'******************************************************************
'システム名     ：教務事務システム
'処　理　名     ：欠席日数登録
'プログラムID   ：gak/sei0600/default.asp
'機　　　能     ：フレームページ 試験毎所見登録の表示を行う
'------------------------------------------------------------------
'引　　　数     ：
'変　　　数     ：
'引　　　渡     ：
'説　　　明     ：
'------------------------------------------------------------------
'作　　　成     ：2001/09/26 谷脇
'変　　　更     ：
'******************************************************************
Public m_sMode
Public m_iNendo
Public m_sKyokanCd
Public m_iSikenKBN
Public m_sGakuNo
Public m_sGakunen
Public m_sClass
Public m_sClassNm
'*******************　ASP共通モジュール宣言　**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******　モ ジ ュ ー ル 変 数　********
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
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_parent"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sMode = request("txtMode")
    m_iNendo = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iSikenKBN = request("txtSikenKBN")
    m_sGakuNo = request("GakuseiNo")
    m_sGakunen = request("txtGakunen")
    m_sClass = request("txtClass")
    m_sClassNm = request("txtClassNm")

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
		session("PRJ_No") = "SEI0400"

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '// 担任チェック
	  If gf_Tannin(m_iNendo,m_sKyokanCd,1) <> 0 Then
	            m_bErrFlg = True
	            m_sErrMsg = "担任以外の入力はできません。"
	            Exit Do
	  End If

'-------2001/08/30 ito-------
'		If m_sGakuNo <> "" Then
'			Call showPageBack()
'	        Exit Do
'		End If

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
    On Error Resume Next
    Err.Clear

%>
<html>
    <head>
    <title>欠席日数登録</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    </script>
    <frameset rows=190,1,* frameborder="0" FRAMESPACING="0" border="0">
        <frame src="sei0600_top.asp?txtGakuNo=<%=request("txtGakuNo")%>&txtSikenKBN=<%=m_iSikenKBN%>" scrolling="auto" noresize name="topFrame">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        <frame src="default2.asp" scrolling="auto" noresize name="main">
    </frameset>
    </head>
</html>
<%
End Sub

Sub showPageBack()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    On Error Resume Next
    Err.Clear

%>
<html>
    <head>
    <title>欠席日数登録</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    </script>
txtSikenKBN=<%=m_iSikenKBN%>
<% response.end %>
    <frameset rows=180,1,* frameborder="0" FRAMESPACING="0" border="0">
        <frame src="sei0600_top.asp?txtGakuNo=<%=m_sGakuNo%>&txtSikenKBN=<%=m_iSikenKBN%>" scrolling="auto" noresize name="topFrame">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        <frame src="sei0600_main.asp?txtGakuNo=<%=m_sGakuNo%>&txtGakunen=<%=m_sGakunen%>&txtClass=<%=m_sClass%>&txtClassNm=<%=m_sClassNm%>&txtSikenKBN=<%=m_iSikenKBN%>" scrolling="auto" noresize name="main">
    </frameset>
    </head>
</html>
<%
End Sub
%>