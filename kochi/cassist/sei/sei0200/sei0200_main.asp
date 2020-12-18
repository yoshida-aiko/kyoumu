<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績一覧
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0200/sei0200_main.asp
' 機      能: フレームページ 成績一覧の登録を行う
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/10/22 谷脇　良也
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
	w_sMsgTitle="成績一覧"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
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
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 権限チェックに使用
		session("PRJ_No") = "SEI0200"

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

dim w_sItem

	w_sItem = ""
	w_sItem = w_sItem & "?txtSikenKBN=" & request("txtSikenKBN")
	w_sItem = w_sItem & "&txtHyojiKBN=" & request("txtHyojiKBN")
	w_sItem = w_sItem & "&txtGakuNo=" & request("txtGakuNo")
	w_sItem = w_sItem & "&txtClassNo=" & request("txtClassNo")
	w_sItem = w_sItem & "&txtKBN=" & request("txtKBN")
	w_sItem = w_sItem & "&txtNendo=" & request("txtNendo")
	w_sItem = w_sItem & "&txtKyokanCd=" & request("txtKyokanCd")
	w_sItem = w_sItem & "&txtKengen=" & request("txtKengen")

if request("txtGakkaNo") = C_SEI0200_ACCESS_TANNIN then
	w_sItem = w_sItem & "&txtClassNo=" & request("txtClassNo")
ElseIf request("txtGakkaNo") = C_SEI0200_ACCESS_GAKKA then
	w_sItem = w_sItem & "&txtGakkaNo=" & request("txtGakkaNo")
End if
%>

<html>

<head>
<title>成績一覧</title>
</head>

<frameset rows=310,1,* frameborder="0" framespacing="0">
	<frame src="sei0200_middle.asp<%=w_sItem%>" scrolling="no"  name="stop">
    <frame src="../../common/bar.html" scrolling="auto" name="bar" noresize>
	<frame src="sei0200_bottom.asp<%=w_sItem%>" scrolling="auto"  name="smain" >
</frameset>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
