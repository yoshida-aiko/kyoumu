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
' 変      更: 2005/09/13    西村 岐阜高専のみ欠課一覧を表示する
' 変      更: 2005/10/05    西村 連絡事項を使用可能に修正
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
Dim m_MenuMode		'//ﾒﾆｭｰﾓｰﾄﾞ
Dim m_sGakkoNO       '学校番号　INS2005/09/13 西村

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
    w_sMsgTitle="ヘッダーデータ"
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

	    '2005.09.13 Ins 西村
	    '学校番号の取得 
	    if Not gf_GetGakkoNO(m_sGakkoNO) then 
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "学校番号の失敗しました。"
            Exit Do
		end if
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
    <title>学籍データ検索</title>
    <meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<link rel=stylesheet href="../common/style.css" type=text/css>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  リロードしてメニューの表示をかえる
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function fj_CaseMenu(pMode){

		document.frm.hidMenuMode.value = pMode;
		document.frm.action="menu.asp";
		document.frm.target="menu";
		document.frm.submit();
		
    }
    //-->
    </SCRIPT>

    </head>

    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/back.gif">
	<form name="frm" method="post">
    <table border="0" cellspacing="0" cellpadding="0" width="150" height="100%">
        <tr>
            <td align="center" valign="top">

                <table border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td class=home>

                            <table bordercolor="#222268" border="1" cellspacing="0" cellpadding="0" width="140">
                                <tr>
					<% if m_MenuMode = "" then %>
                                    <td class=home align="center"><font color="#ffff00">Ｔ　Ｏ　Ｐ</font></td>
					<% Else %>
                                    <td class=home align="center"><font color="#ffffff"><a class=menu href="top.asp" target="<%=C_MAIN_FRAME%>" onClick="javascript:fj_CaseMenu('');">Ｔ　Ｏ　Ｐ</a></font></td>
					<% End if %>
                                </tr>
                            </table>

                        </td>
                    </tr>
                    <tr><td><img src="../image/sp.gif"></td></tr>
					<!--
					<% if m_MenuMode = "SYUKETU" then %>
							<tr><td class=category><font color="#ffff00">出欠入力<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('SYUKETU');"><font color="#ffffff">出欠入力<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					
					<% if m_MenuMode = "SYUKETU" then Call s_MenuData("SYUKETU") %>
					<tr><td><img src="../image/sp.gif"></td></tr>
					-->
					<% if m_MenuMode = "SHIKEN" then %>
							<tr><td class=category><font color="#ffff00">試験・成績<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('SHIKEN');"><font color="#ffffff">試験・成績<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					
					<% if m_MenuMode = "SHIKEN" then Call s_MenuData("SHIKEN") %>
					<tr><td><img src="../image/sp.gif"></td></tr>
					
					<!--
					<% if m_MenuMode = "SCHE" then %>
							<tr><td class=category><font color="#ffff00">スケジュール<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('SCHE');"><font color="#ffffff">スケジュール<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "SCHE" then Call s_MenuData("SCHE") %>
					<tr><td><img src="../image/sp.gif"></td></tr>
					
					<% if m_MenuMode = "OTHERS" then %>
							<tr><td class=category><font color="#ffff00">その他入力<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('OTHERS');"><font color="#ffffff">その他入力<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					
					<% if m_MenuMode = "OTHERS" then Call s_MenuData("OTHERS") %>
					<tr><td><img src="../image/sp.gif"></td></tr>
					-->
					
					<% if m_MenuMode = "INFO" then %>
							<tr><td class=category><font color="#ffff00">情報検索<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('INFO');"><font color="#ffffff">情報検索<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					
					<% if m_MenuMode = "INFO" then Call s_MenuData("INFO") %>
					<tr><td><img src="../image/sp.gif"></td></tr>
					
				　
					<% if m_MenuMode = "SUPPORT" then %>
							<tr><td class=category><font color="#ffff00">支援機能<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('SUPPORT');"><font color="#ffffff">支援機能<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					
					<% if m_MenuMode = "SUPPORT" then Call s_MenuData("SUPPORT") %>
					<tr><td><img src="../image/sp.gif"></td></tr>
<!--
                    <tr><td class=info align="center"><font color="#ffffff"><a class=menu href="http://www.infogram.co.jp/" target="_blank"><img src="images/logo.gif" border="0"></a></font></td></tr>
//-->
                </table>

            </td>
        </tr>
    </table>

	<input type="hidden" name="hidMenuMode">
	</form>
    </body>

    </html>
<% End Sub



'********************************************************************************
'*  [機能]  データ登録メニュー
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]   
'********************************************************************************
Sub s_MenuData(p_menu) 
Select Case p_menu
	Case "SYUKETU" %>
		<!--
		<% if gf_empMenu("KKS0112") then %><tr><td><a class=menu href="../kks/kks0112/" target="<%=C_MAIN_FRAME%>">授業出欠入力</a></td></tr><% End if %>
		<% if gf_empMenu("KKS0111") then %><tr><td><a class=menu href="../kks/kks0111/" target="<%=C_MAIN_FRAME%>">授業出欠参照</a></td></tr><% End if %>
		
		<% if gf_empMenu("KKS0140") then %><tr><td><a class=menu href="../kks/kks0140/" target="<%=C_MAIN_FRAME%>">行事出欠入力</a></td></tr><% End if %>
		<% if gf_empMenu("KKS0170") then %><tr><td><a class=menu href="../kks/kks0170/" target="<%=C_MAIN_FRAME%>">日毎出欠入力</a></td></tr><% End if %>
		-->
	<% Case "SHIKEN" %>
		<!--
		<% if gf_empMenu("SKN0130") then %><tr><td><a class=menu href="../skn/skn0130/" target="<%=C_MAIN_FRAME%>">試験実施科目登録</a></td></tr><% End if %>
		<% if gf_empMenu("SKN0120") then %><tr><td><a class=menu href="../skn/skn0120/" target="<%=C_MAIN_FRAME%>">試験監督免除申請登録</a></td></tr><% End if %>
		-->
		<!--2005/10/03 Upd_S 西村 -->
		<% if m_sGakkoNO = cstr(C_NCT_GIFU) Then %>
			<% if gf_empMenu("SEI0100") then %><tr><td><a class=menu href="../sei/sei0150_23/" target="<%=C_MAIN_FRAME%>">成績登録</a></td></tr><% End if %>
		<% ELSE %>
			<% if gf_empMenu("SEI0100") then %><tr><td><a class=menu href="../sei/sei0150/" target="<%=C_MAIN_FRAME%>">成績登録</a></td></tr><% End if %>
		<!-- ins 2007/04/11 岩田 放送大学成績登録追加 -->
			<% if gf_empMenu("SEI0100") then %><tr><td><a class=menu href="../sei/sei0160/" target="<%=C_MAIN_FRAME%>">放送大学成績登録</a></td></tr><% End if %>
		<!-- ins 2007/04/11 岩田 放送大学成績登録追加 -->
		<% END IF %>	
		<!--<% if gf_empMenu("SEI0100") then %><tr><td><a class=menu href="../sei/sei0100/" target="<%=C_MAIN_FRAME%>">成績登録</a></td></tr><% End if %>-->
		<!--<% if gf_empMenu("SEI0500") then %><tr><td><a class=menu href="../sei/sei0500/" target="<%=C_MAIN_FRAME%>">実力試験成績登録</a></td></tr><% End if %>-->
		<% if gf_empMenu("SEI0400") then %><tr><td><a class=menu href="../sei/sei0400/" target="<%=C_MAIN_FRAME%>">試験毎所見登録</a></td></tr><% End if %>
		<!--<% if gf_empMenu("SEI0600") then %><tr><td><a class=menu href="../sei/sei0600/" target="<%=C_MAIN_FRAME%>">欠席日数登録</a></td></tr><% End if %>-->
		<!--<% if gf_empMenu("SKN0170") then %><tr><td><a class=menu href="../skn/skn0170/" target="<%=C_MAIN_FRAME%>">試験時間割(クラス別)</a></td></tr><% End if %>-->
		<!--<% if gf_empMenu("SKN0180") then %><tr><td><a class=menu href="../skn/skn0180/" target="<%=C_MAIN_FRAME%>">試験期間教官予定一覧</a></td></tr><% End if %>-->
		<% if gf_empMenu("SEI0200") then %><tr><td><a class=menu href="../sei/sei0700/default.asp?p_mode=P_HAN0100" target="<%=C_MAIN_FRAME%>">成績一覧</a></td></tr><% End if %>
		<%  'UPD 2006/11/21 新谷 欠課一覧は岐阜と北九州
			if (m_sGakkoNO = cstr(C_NCT_GIFU) OR m_sGakkoNO = cstr(C_NCT_KITAKYU)) AND gf_empMenu("SEI0200") then %><tr><td><a class=menu href="../sei/sei0700/default.asp?p_mode=P_KKS0200" target="<%=C_MAIN_FRAME%>">欠課一覧</a></td></tr><% End if %>
		<%  'UPD 2006/11/21 新谷 遅刻一覧は北九州のみ
			if m_sGakkoNO = cstr(C_NCT_KITAKYU) AND gf_empMenu("SEI0200") then %><tr><td><a class=menu href="../sei/sei0700/default.asp?p_mode=P_KKS0210" target="<%=C_MAIN_FRAME%>">遅刻一覧</a></td></tr><% End if %>
		<!--<% if gf_empMenu("SEI0300") then %><tr><td><a class=menu href="../sei/sei0300/" target="<%=C_MAIN_FRAME%>">個人別成績一覧</a></td></tr><% End if %>-->
		<!--<% if gf_empMenu("HAN0121") then %><tr><td><a class=menu href="../han/han0121/" target="<%=C_MAIN_FRAME%>">留年該当者一覧</a></td></tr><% End if %>-->
		
	<% Case "SCHE" %>
		<!--
		<% if gf_empMenu("GYO0200") then %><tr><td><a class=menu href="../gyo/gyo0200/" target="<%=C_MAIN_FRAME%>">行事日程一覧</a></td></tr><% End if %>
		<% if gf_empMenu("JIK0210") then %><tr><td><a class=menu href="../jik/jik0210/" target="<%=C_MAIN_FRAME%>">クラス別授業時間一覧</a></td></tr><% End if %>
		<% if gf_empMenu("JIK0200") then %><tr><td><a class=menu href="../jik/jik0200/" target="<%=C_MAIN_FRAME%>">教官別授業時間一覧</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0310") then %><tr><td><a class=menu href="../web/web0310/" target="<%=C_MAIN_FRAME%>">時間割交換連絡</a></td></tr><% End if %>
		-->
	<% Case "OTHERS" %>
		<!--
		<% if gf_empMenu("MST0144") then %><tr><td><a class=menu href="../mst/mst0144/" target="<%=C_MAIN_FRAME%>">進路先情報登録</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0320") then %><tr><td><a class=menu href="../web/web0320/" target="<%=C_MAIN_FRAME%>">使用教科書登録</a></td></tr><% End if %>
		<% if gf_empMenu("GAK0460") then %><tr><td><a class=menu href="../gak/gak0460/" target="<%=C_MAIN_FRAME%>">指導要録所見等登録</a></td></tr><% End if %>
		<% if gf_empMenu("GAK0461") then %><tr><td><a class=menu href="../gak/gak0461/" target="<%=C_MAIN_FRAME%>">調査書所見等登録</a></td></tr><% End if %>
		<% if gf_empMenu("GAK0470") then %><tr><td><a class=menu href="../gak/gak0470/" target="<%=C_MAIN_FRAME%>">各種委員登録</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0340") then %><tr><td><a class=menu href="../web/web0340/" target="<%=C_MAIN_FRAME%>">個人履修選択科目決定</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0390") then %><tr><td><a class=menu href="../web/web0390/" target="<%=C_MAIN_FRAME%>">レベル別科目決定</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0360") then %><tr><td><a class=menu href="../web/web0360/" target="<%=C_MAIN_FRAME%>">部活動部員一覧</a></td></tr><% End if %>
		-->
	<% Case "INFO" %>
		<% if gf_empMenu("GAK0300") then %><tr><td><a class=menu href="../gak/gak0310/" target="<%=C_MAIN_FRAME%>">学生情報検索</a></td></tr><% End if %>
		<!--
		<% if gf_empMenu("MST0113") then %><tr><td><a class=menu href="../mst/mst0113/" target="<%=C_MAIN_FRAME%>">中学校情報検索</a></td></tr><% End if %>
		<% if gf_empMenu("MST0123") then %><tr><td><a class=menu href="../mst/mst0123/" target="<%=C_MAIN_FRAME%>">高等学校情報検索</a></td></tr><% End if %>
		<% if gf_empMenu("MST0133") then %><tr><td><a class=menu href="../mst/mst0133/" target="<%=C_MAIN_FRAME%>">進路先情報検索</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0350") then %><tr><td><a class=menu href="../web/web0350/" target="<%=C_MAIN_FRAME%>">空き時間情報検索</a></td></tr><% End if %>
		-->
	<% Case "SUPPORT" %>
		<!--
		<% if gf_empMenu("WEB0300") then %><tr><td><a class=menu href="../web/web0300/" target="<%=C_MAIN_FRAME%>">特別教室予約</a></td></tr><% End if %>
		-->
		<% if gf_empMenu("WEB0330") then %><tr><td><a class=menu href="../web/web0330/" target="<%=C_MAIN_FRAME%>">連絡事項登録</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0330") then %><tr><td><a class=menu href="../login/top.asp" target="<%=C_MAIN_FRAME%>">連絡掲示板</a></td></tr><% End if %>
		
<% End Select %>
		<tr><td> </td></tr>
<%
End Sub
%>