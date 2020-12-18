<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験監督免除申請登録
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0120/default.asp
' 機      能: フレームページ 教官予定マスタの参照を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/06/18 高丘 知央
' 変      更: 2001/06/26 根本
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
    w_sMsgTitle="試験監督免除登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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
		session("PRJ_No") = "SKN0120"

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
    '---------- HTML START  ----------

's_Arg_top = "?"
's_Arg_top = s_Arg_top & "txtSikenKbn=" & Request("txtSikenKbn")

's_Arg = "?"
's_Arg = s_Arg & "txtSikenKbn=" & Request("txtSikenKbn")
's_Arg = s_Arg & "txtSikenCd" & Request("txtSikenCd")
's_Arg = s_Arg & "txtMode" & Request("txtMode")

%>

<html>

<head>

<title>試験監督免除申請登録</title>

<frameset rows=125,1,* frameborder="no">
<%
'//初期表示
If Request("txtMode")="" or Request("txtMode") = "no" Then%>
    <frame src="top.asp" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
    <frame src="default2.asp?txtMode=<%=Request("txtMode")%>" scrolling="auto" noresize name="main">
<%
'//更新処理後
Else%>
    <frame src="top.asp?<%=request.form.item%>" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
    <frame src="main.asp?<%=request.form.item%>" scrolling="auto" noresize name="main">
<%End If%>
</frameset>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>
