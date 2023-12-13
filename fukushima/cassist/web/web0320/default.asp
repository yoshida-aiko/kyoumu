<%@Language=VBScript %>
<%
'******************************************************************
'システム名     ：教務事務システム
'処　理　名     ：使用教科書登録
'プログラムID   ：web/web0320/default.asp
'機　　　能     ：フレームページ 使用教科書登録の表示を行う
'------------------------------------------------------------------
'引　　　数     ：
'変　　　数     ：
'引　　　渡     ：
'説　　　明     ：
'------------------------------------------------------------------
'作　　　成     ：2001.08.01    前田　智史
'変　　　更     ：
'
'******************************************************************
'*******************　ASP共通モジュール宣言　**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******　モ ジ ュ ー ル 変 数　********

	Public m_iNendo
	Public m_iGakunen
	Public m_iClassNo
	Public m_iPage

'******　メイン処理　********

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'******　Ｅ　Ｎ　Ｄ　********

'******************************************************************
'機　　能：本ASPのﾒｲﾝﾙｰﾁﾝ
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Sub Main()
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="使用教科書登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""
	
    On Error Resume Next
    Err.Clear
	
    m_bErrFlg = False
	
    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        If gf_OpenDatabase() <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
		
		'// 権限チェックに使用
		session("PRJ_No") = "WEB0320"
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'm_iNendo = Request("txtNendo")
		m_iNendo = Request("hidYear")
		
		If m_iNendo <> "" Then
	        '// ページを表示
	        Call showPage_Reload()
		Else
	        '// ページを表示
	        Call showPage()
		End If

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
    <title>使用教科書登録</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    </script>
    <frameset rows=140,1,* frameborder="0">
        <frame src="web0320_top.asp" scrolling="auto" noresize name="top">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        <frame src="default2.asp" scrolling="auto" noresize name="main">
    </frameset>
    </head>
</html>
<%
End Sub

Sub showPage_Reload()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	sArg = ""
	sArg = sArg & "?txtNendo=" & m_iNendo
%>
<html>
    <head>
    <title>使用教科書登録</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    <!--
    /*
    window.onload = init;
    
    function init(){
		alert("<%=request("hidYear")%>");
		alert("<%=request("hidGakunen")%>");
		alert("<%=request("hidGakka")%>");
	}
	*/
    //-->
    </script>
    <frameset rows=140,1,* frameborder="0">
        <frame src="web0320_top.asp?<%=Request.Form.Item%>" scrolling="auto" noresize name="top">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        
        <% if request("txtPageCD") <> "" then %>
        	<frame src="web0320_main.asp?<%=Request.Form.Item%>" scrolling="auto" noresize name="main">
        <% else %>
        	<frame src="default2.asp" scrolling="auto" noresize name="main">
        <% end if %>
    </frameset>
    </head>
</html>
<%
End Sub
%>