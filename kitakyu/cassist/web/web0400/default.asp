<%@ Language=VBScript %>
<%Response.Expires = 0%>
<%Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: パスワード変更
' ﾌﾟﾛｸﾞﾗﾑID : web/web0400/default.asp
' 機      能: ログインパスワードを変更します。
'-------------------------------------------------------------------------
' 引      数:SESSION(""):教官コード     ＞      SESSIONより
' 変      数:なし
' 引      渡:SESSION(""):教官コード     ＞      SESSIONより
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/10/04 谷脇
' 変      更: 2019/03/18 藤林 パスワードのエラーチェックを半角英数記号チェックに変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public m_iNendo				'処理年度
    Public m_sUser				'ログインユーザＩＤ
    Public m_sPass				'古いパスワード
    Public m_sPassN1			'新しいパスワード１
    Public m_sPassN2			'新しいパスワード２
	Public m_loginF 			'ログインID入力させるかどうか

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()
response.end
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
    w_sMsgTitle="パスワード変更"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do

		'// 変数初期化
		call f_paraSet()

			'// 権限チェックに使用
	'		session("PRJ_No") = "WEB0400"

			'// 不正アクセスチェック
	'		Call gf_userChk(session("PRJ_No"))
			
        '// 変更ページを表示
        Call showPage()
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
'response.write w_sMsg
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
End Sub

Sub f_paraSet()
'*******************************************************************************
' 機　　能：変数の初期化と代入
' 引　　数：なし
' 機能詳細：
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
m_sUser = Request("txtUser")
m_sPass = Request("txtPass")
m_sPassN1 = Request("txtPassN1")
m_sPassN2 = Request("txtPassN2")

if m_sUser = "" then m_sUser = session("LOGIN_ID")
'm_iNendo = 2001

m_loginF = true 'ログインＩＤを入力させるときはtrue
	
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
    <title>パスワード変更</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
		function jf_check(){
		var w_str //作業用
		var w_nm //作業用
		var w_msg //作業用
		var w_err //作業用
		
		w_err = true;
		while(1) {
			//ログインIDの入力チェック
			w_nm = eval(document.frm.txtUser);
			w_str = w_nm.value;
			if (w_str == "") {w_str = "ログインIDを入力してください。";break;}
			if (!IsHankakuEisu(w_str)) {w_str = "ログインIDが、半角英数文字ではありません。";break;}
			if (getLengthB(w_str) > 16 ) {w_str = "ログインIDが、16文字以下ではありません。";break;}

			//パスワードの入力チェック
			w_nm = eval(document.frm.txtPass);
			w_str = w_nm.value;
			if (w_str == "") {w_str = "パスワードを入力してください。";break;}
			//if (!IsHankakuEisu(w_str)) {w_str = "パスワードが、半角英数文字ではありません。";break;}
			if (getLengthB(w_str) > 10 ) {w_str = "パスワードが、10文字以下ではありません。";break;}

			//パスワードの入力チェック
			w_nm = eval(document.frm.txtPassN1);
			w_str = w_nm.value;
			if (w_str == "") {w_str = "新しいパスワード(1)を入力してください。";break;}
			if (!IsHankakuEisuKigo(w_str)) {w_str = "新しいパスワード(1)が、半角英数記号（シングルコーテーションを除く）文字ではありません。";break;}
			if (getLengthB(w_str) > 10 ) {w_str = "新しいパスワード(1)が、10文字以下ではありません。";break;}

			//パスワードの入力チェック
			w_nm = eval(document.frm.txtPassN2);
			w_str = w_nm.value;
			if (w_str == "") {w_str = "新しいパスワード(2)を入力してください。";break;}
			if (!IsHankakuEisuKigo(w_str)) {w_str = "新しいパスワード(2)が、半角英数記号（シングルコーテーションを除く）文字ではありません。";break;}
			if (getLengthB(w_str) > 10 ) {w_str = "新しいパスワード(2)が、10文字以下ではありません。";break;}
			if (document.frm.txtPassN1.value != w_str) {w_str = "新しいパスワード(2)が、新しいパスワード(1)と同じではありません。";break;}

		w_err = false;
		break;
		}

		//エラー有り
		if (w_err) {
			jf_errMsg(w_nm,w_str);
			return false;
		}
		
		//正常終了
		return true;
//		return false; //テスト用

		}

		function jf_errMsg(p_nm,p_str) {
			alert(p_str);
			p_nm.select();
			return;
		}

	//-->
	</script>
</head>
<body>
<center>
    <%call gs_title("ログインパスワード変更","更　新")%>

<BR>
ログインパスワードの変更を行うことができます。<br>
セキュリティ管理のため、定期的にパスワードの変更を行うようにしてください。<BR><BR>
	<table class="hyo" border="0" width="70%">
		<FORM action="web0400_upd.asp" name="frm" method="post" target="_self" onSubmit="return jf_check();">
			<tr><td colspan="2" height="30" class="detail"></td></tr>
	        <tr>
	            <td nowrap class="detail" align="right" width="45%">ログインID：</td>
<% If m_loginF = true then %>
	            <td nowrap class="detail" width="55%">　<input type="text" name="txtUser" value="<%=m_sUser%>" maxlength="16" size="25"></td>
	        <tr>
	            <td nowrap class="detail" align="right"></td>
	            <td nowrap class="detail">　<span class="CAUTION" style="text-align:left;">※半角英数16文字以内</span></td>
	        </tr>
<% else %>
	            <td nowrap class="detail" width="55%">　<%=m_sUser%><input type="hidden" name="txtUser" value="<%=m_sUser%>" maxlength="16" size="25"></td></tr>
			<tr><td colspan="2" height="15" class="detail"></td>
<% End If %>
	        </tr>
			<tr><td colspan="2" height="15" class="detail"></td></tr>
	        <tr>
	            <td nowrap class="detail" align="right">パスワード：</td>
	            <td nowrap class="detail">　<input type="password" name="txtPass" value="<%=m_sPass%>" maxlength="10"></td>
	        </tr>
	        <tr>
	            <td nowrap class="detail" align="right"></td>
	            <td nowrap class="detail">　<span class="CAUTION" style="text-align:left;">※半角英数10文字以内</span></td>
	        </tr>
			<tr><td colspan="2" height="15" class="detail"></td></tr>
	        <tr>
	            <td nowrap class="detail" align="right">新しいパスワード：</td>
	            <td nowrap class="detail">　<input type="password" name="txtPassN1" value="<%=m_sPassN1%>" maxlength="10"></td>
	        </tr>
	        <tr>
	            <td nowrap class="detail" align="right"></td>
	            <td nowrap class="detail">　<span class="CAUTION" style="text-align:left;">※半角英数10文字以内</span></td>
	        </tr>
			<tr><td colspan="2" height="15" class="detail"></td></tr>
	        <tr>
	            <td nowrap class="detail" align="right">新しいパスワード：</td>
	            <td nowrap class="detail">　<input type="password" name="txtPassN2" value="<%=m_sPassN2%>" maxlength="10"></td>
	        </tr>
	        <tr>
	            <td nowrap class="detail" align="right"></td>
	            <td nowrap class="detail">　<span class="CAUTION" style="text-align:left;">※確認のためにもう一度入力してください</span></td>
	        </tr>
			<tr><td colspan="2" height="30" class="detail"></td></tr>
			<tr><td colspan="2" height="30" class="detail" align="center">
				<input type="submit" name="submit" value=" 変 更 " maxlength="10">
	            　<input type="reset" name="can" value="ｷｬﾝｾﾙ" maxlength="10" onclick="history.back()">
	            </td>
	        </tr>
		</FORM>
	</table>
</center>
</body>
</head>
</html>
<%
End Sub
%>