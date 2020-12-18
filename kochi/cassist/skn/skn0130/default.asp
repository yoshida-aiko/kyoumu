<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験実施科目登録
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0130/default.asp
' 機      能: フレームページ 教官予定マスタの参照を行う
'-------------------------------------------------------------------------
' 引      数:
' 変      数:
' 引      渡:
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/06/18
' 変      更: 2001/06/26
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
    w_sWinTitle = "キャンパスアシスト"
    w_sMsgTitle = "試験実施科目登録"
    w_sMsg = ""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget = ""

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
		session("PRJ_No") = "SKN0130"

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '// 初期ページを表示
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
	dim w_item
    '---------- HTML START  ----------
    On Error Resume Next
    Err.Clear
    
	w_item = ""
    w_item = w_item & "txtMode="&request("txtMode")
    w_item = w_item & "&txtSikenKbn="&request("txtSikenKbn")

%>
<html>
<head>
<% 'タイトルがあるとFireFoxで文字化けするため削除 --2019/06/24 Del Fujibayashi <title>試験実施科目登録</title> %>
</head>

<frameset rows=120,1,* frameborder="no">
	<frame src="SKN0130_top.asp?<%=request.form.item%>" scrolling="auto" noresize>
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
	<frame src="SKN0130_main.asp?<%=w_item%>" scrolling="auto" noresize name=main>
</frameset>

</html>
<% End Sub %>