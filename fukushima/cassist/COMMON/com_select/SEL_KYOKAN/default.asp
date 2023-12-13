<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 教官参照選択画面
' ﾌﾟﾛｸﾞﾗﾑID : Common/com_select/SEL_KYOKAN/default.asp
' 機      能: フレームページ 教官の参照、選択を行う
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/07/19 前田 智史
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bErrMsg           'ｴﾗｰﾒｯｾｰｼﾞ
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
	w_sMsgTitle="教官参照選択画面"
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
            m_bErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

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
Dim w_iI
Dim w_sKNM
Dim w_sGakkaCd
	w_iI	 = request("txtI")
	w_sKNM	 = request("txtKNm")
	w_sGakkaCd = request("txtGakka")
	sArg = ""
	sArg = sArg & "txtI=" & w_iI 
	sArg = sArg & "&txtKNm=" & Server.URLEncode(w_sKNM)
	sArg = sArg & "&txtGakka=" & Server.URLEncode(w_sGakkaCd)

%>
<html>

<head>

<title>教官参照選択画面</title>

<frameset rows=175px,1,* frameborder="no">
	<frame src="SEL_KYOKAN_top.asp?<%=sArg %>" scrolling="auto" noresize name="top">
    <frame src="bar.html" scrolling="auto" noresize name="bar">
	<frame src="default2.asp" scrolling="auto" noresize name="main">
</frameset>

</head>

</html>
<%
End Sub
%>