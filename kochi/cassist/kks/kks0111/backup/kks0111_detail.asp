<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠一覧
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0110/kks0111_detail.asp
' 機      能: フレームページ 授業出欠表示
'-------------------------------------------------------------------------
' 引      数:
' 変      数:
' 引      渡:
' 説      明:
'           
'-------------------------------------------------------------------------
' 作      成: 2002/05/07 shin
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

'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()
	Dim w_iRet              '// 戻り値
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="授業出欠入力"
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
            m_bErrFlg = True
            w_sMsg = "データベースとの接続に失敗しました。"
            'm_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
		
		'// 権限チェックに使用
		session("PRJ_No") = "KKS0110"
		
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
<title>授業出欠入力</title>

<frameset rows="138px,*" border="1" frameborder="no" onBlur="window.focus();">
    <frame src="kks0111_detail_top.asp?<%=Request.QueryString%>" scrolling="yes" noresize name="topFrame">
    <frame src="kks0111_detail_bottom.asp?<%=Request.QueryString%>" scrolling="yes" noresize name="main">
    
</frameset>

</head>
</html>
<%
End Sub
%>