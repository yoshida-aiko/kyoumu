<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: ログイン終了時画面
' ﾌﾟﾛｸﾞﾗﾑID : login/menu.asp
' 機      能: ログイン終了時のメニュー画面
'-------------------------------------------------------------------------
' 引      数    
'               
' 変      数
' 引      渡
'           
'           
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 
' 変      更: 2001/07/26    モチナガ
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
Dim m_MenuMode		'//ﾒﾆｭｰﾓｰﾄﾞ

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

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="フッターデータ"
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

		'//ﾒﾆｭｰﾓｰﾄﾞ
		m_MenuMode = request("hidMenuMode")

        '//初期表示
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

    %>
    <html>
    <head>
    <title>フッター</title>
    <meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<link rel=stylesheet href="../common/style.css" type=text/css>
	    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  トップへ戻る。
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function fj_BackTop() {

		document.frm.action="menu.asp";
		document.frm.target="menu";
		document.frm.submit();
		
    }
    //-->
    </SCRIPT>
    </head>
    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/back.gif" style="background-repeat: repeat-y;">
<table border=0 cellpadding=0 cellspacing=0 width="100%">
		<form action="menu.asp" method="post" name="frm">
            <tr><td width="152" class="info" align="center" valign="center" nowrap><span color="#ffffff"><a class="menu" href="http://www.infogram.co.jp/" target="_blank"><img src="images/logo.gif" border="0"></a></span></td>
				<td>　</td>
				<td width="125" align="right" nowrap><a href="../web/web0380/default.asp" target="<%=C_MAIN_FRAME%>" onClick="">［異動状況一覧］</a></td>
				<td width="120" align="right" nowrap><a href="../web/web0370/default.asp" target="<%=C_MAIN_FRAME%>" onClick="">［学生数一覧］</a></td>
				<td width="125" align="right" nowrap><a href="top.asp" target="<%=C_MAIN_FRAME%>" onClick="javascript:fj_BackTop()">［トップへ戻る］</a></td>
				<td width="120" align="right" nowrap><a href="../default.asp" target="_top">［ログアウト］</a></td>
			</tr>
		</form>
		</table></body>
</html>
<%
End Sub%>
