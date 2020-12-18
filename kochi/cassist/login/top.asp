<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: （トップページ）
' ﾌﾟﾛｸﾞﾗﾑID : login/top.asp
' 機      能: フレームページ 今週の時間割とお知らせの参照を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/07/19 根本 直美
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
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
    Dim w_sSQL              '// SQL文
    Dim w_sRetURL           '// ｴﾗｰﾒｯｾｰｼﾞ用戻り先URL
    Dim w_sTarget           '// ｴﾗｰﾒｯｾｰｼﾞ用戻り先ﾌﾚｰﾑ
    Dim w_sWinTitle         '// ｴﾗｰﾒｯｾｰｼﾞ用ﾀｲﾄﾙ
    Dim w_sMsgTitle         '// ｴﾗｰﾒｯｾｰｼﾞ用ﾀｲﾄﾙ
    
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="トップページ"
    w_sMsg=""
    w_sRetURL="../default.asp"
    w_sTarget="_parent"

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
		session("PRJ_No") = C_LEVEL_NOCHK

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		Call showPage()         '// ページを表示

        '// 正常終了
        Exit Do
    LOOP

   '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, m_sErrMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gs_CloseDatabase()

End Sub


'********************************************************************************
'*  [機能]  HTML表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()

'	if gf_IsNull(trim(Session("KYOKAN_CD"))) then
		w_FremeSize = "0,*"
		w_FrameUrl  = "about:blank;"
'	Else
'		w_FremeSize = "300,*"
'		w_FrameUrl  = "jikanwari.asp"
'	End if

%>

<html>

<head>
<title>教務事務システム：Campus Assist トップページ</title>
</head>

<frameset rows="<%=w_FremeSize%>" frameborder="0">
	<frame src="<%=w_FrameUrl%>" scrolling="auto" noresize name="<%=C_MAIN_FRAME%>_up">
	<frame src="top_lwr.asp"   scrolling="auto" noresize name="<%=C_MAIN_FRAME%>_low">
</frameset>

</html>

<% End Sub %>