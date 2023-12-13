<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績参照（教官側）
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0800/default.asp
' 機      能: 
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2003/05/13 廣田
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	Dim     m_sGakuseiNo
	Dim     m_sGakunen
	Dim     m_sClass
	Dim     m_sGakuseiNM

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
	w_sMsgTitle="成績参照"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_parent"

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
'		Session("PRJ_No") = "SEI0800"

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		Call s_SetParam()

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
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()

	m_sGakuseiNo = Request("hidGakuseiNo")
	m_sGakunen   = Request("hidGakunen")
	m_sClass     = Request("hidClass")
	m_sGakuseiNM = Request("hidGakuseiNM")

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
<!--#include file="../../Common/scroll.js"-->
<title>成績参照</title>
</head>

<frameset rows=137px,1,* frameborder="0" framespacing="0">
	<frame src="sei0800_resulttop.asp?hidGakunen=<%=m_sGakunen%>&hidClass=<%=m_sClass%>&hidGakuseiNM=<%=m_sGakuseiNM%>&hidGakuseiNo=<%=m_sGakuseiNo%>"    scrolling="auto" name="topFrame" noresize>
    <frame src="../../common/bar.html"    scrolling="auto" name="bar"      noresize>
	<frame src="sei0800_resultbottom.asp?hidGakuseiNo=<%=m_sGakuseiNo%>" scrolling="auto"  name="main"     noresize>
</frameset>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
