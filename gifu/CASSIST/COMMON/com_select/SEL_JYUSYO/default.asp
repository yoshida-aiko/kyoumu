<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 市町村検索画面
' ﾌﾟﾛｸﾞﾗﾑID : Common/com_select/SEL_JYUSYO/default.asp
' 機      能: フレーム定義部分
'-------------------------------------------------------------------------
' 引      数:	
' 	           	JUSYO1	= 県市区
'   	        JUSYO2	= 町
' 
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/30 持永
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

	Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	Public  m_JUSYO1			'住所1
	Public  m_JUSYO2			'住所2

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
    w_sMsgTitle="連絡事項登録"
    w_sMsg=""
    w_sRetURL="../../login/top.asp"
    w_sTarget="_parent"

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

		'// ﾊﾟﾗﾒｰﾀ取得
		m_JUSYO1 = request("txtJUSYO1")
		m_JUSYO2 = request("txtJUSYO2")

		'// ｾｯｼｮﾝ格納
'		Session("m_JUSYO1") = m_JUSYO1
'		Session("m_JUSYO2") = m_JUSYO2

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


'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear

	m_JUSYO1 = Server.URLEncode(m_JUSYO1)
	m_JUSYO2 = Server.URLEncode(m_JUSYO2)

%>

<html>

<head>
<title>市町村検索</title>
<link rel=stylesheet href="../../style.css" type=text/css>
</head>

<frameset rows=230,1,* frameborder="no" onload="window.focus();">
	<frame src="Jyusyo_top.asp?JUSYO1=<%=m_JUSYO1%>&JUSYO2=<%=m_JUSYO2%>" scrolling="auto" noresize name="top">
        <frame src="bar.html" scrolling="auto" noresize name="bar">
	<frame src="Jyusyo_dow.asp" scrolling="auto" noresize name="dow">
</frameset>

</html>
<%
End Sub
%>