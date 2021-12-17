<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索詳細
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/kojin_sita4.asp
' 機      能: 検索された学生の詳細を表示する(備考・所見)
'-------------------------------------------------------------------------
' 引      数	Session("GAKUSEI_NO")  = 学生番号
'            	Session("HyoujiNendo") = 表示年度
'           
' 変      数
' 引      渡
'           
'           
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public m_bErrFlg        'ｴﾗｰﾌﾗｸﾞ
	Public m_Rs				'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
	Public m_HyoujiFlg		'表示ﾌﾗｸﾞ

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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="学生情報検索結果"
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

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'//表示項目を取得
		w_iRet = f_GetDetailBikou()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//初期表示
        if m_TxtMode = "" then
            Call showPage()
            Exit Do
        end if

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub


'********************************************************************************
'*  [機能]  表示項目を取得
'*  [引数]  なし
'*  [戻値]  0:正常終了	1:任意のエラー  99:システムエラー
'*  [説明]  
'********************************************************************************
Function f_GetDetailBikou()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailBikou = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T11_SYUMITOKUGI,  "
		w_sSql = w_sSql & " 	A.T11_SOGOSYOKEN,  "
		w_sSql = w_sSql & " 	A.T11_KODOSYOKEN,  "
		w_sSql = w_sSql & " 	A.T11_KOJIN_BIK "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T11_GAKUSEKI A "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & "  	A.T11_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "

		iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailBikou = 99
			Exit Do
		End If

		'//正常終了
		f_GetDetailBikou = 0
		Exit Do
	Loop

End Function


'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()

	On Error Resume Next
	Err.Clear

	'// 変数初期化
	m_HyoujiFlg = 0 		'<!-- 表示フラグ（0:なし  1:あり）

	m_SYUMITOKUGI = ""
	m_SOGOSYOKEN  = ""
	m_KODOSYOKEN  = ""
	m_KOJIN_BIK   = ""

	if Not m_Rs.EOF then
		m_SYUMITOKUGI = m_Rs("T11_SYUMITOKUGI")
		m_SOGOSYOKEN  = m_Rs("T11_SOGOSYOKEN")
		m_KODOSYOKEN  = m_Rs("T11_KODOSYOKEN")
		m_KOJIN_BIK   = m_Rs("T11_KOJIN_BIK")
	End if

%>

	<html>
	<head>
	<title>学籍データ参照</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<style type="text/css">
	<!--
		a:link { color:#cc8866; text-decoration:none; }
		a:visited { color:#cc8866; text-decoration:none; }
		a:active { color:#888866; text-decoration:none; }
		a:hover { color:#888866; text-decoration:underline; }
		b { color:#88bbbb; font-weight: bold; font-size:14px}
	//-->
	</style>
	<script language="javascript">
	<!--
		function sbmt(m,i) {
			document.forms[0].mode.value = m;
			document.forms[0].id.value = i;
			document.forms[0].submit();
		}
	//-->
	</script>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">●基本情報</a></td>
			<td nowrap><a href="kojin_sita1.asp">●個人情報</a></td>
			<td nowrap><a href="kojin_sita2.asp">●入学情報</a></td>
			<td nowrap><a href="kojin_sita3.asp">●学年情報</a></td>
			<td nowrap><b>●備考・所見</b></td>
			<td nowrap><a href="kojin_sita5.asp">●異動情報</a></td>
		</tr>
	</table>
	<br>


	<table border="0" cellpadding="1" cellspacing="1">
<!--
		<tr>
			<td valign="top" align="left">

				<br>
				<% if gf_empItem(C_T11_SYUMITOKUGI) then %>
					<table class="disp" border="1" width="220">
						<tr><td class="disph" width="120" height="16">趣味・特技・資格</td></tr>
						<tr><td class="disp" height="100" valign="top"><%= m_SYUMITOKUGI %></td></tr>
					</table>
				<% End if %>

			</td>
			<td valign="top" align="left">

				<br>
				<% if gf_empItem(C_T11_KODOSYOKEN) then %>
					<table class="disp" border="1" width="220">
						<tr><td class="disph" width="100" height="16">行動所見</td></tr>
						<tr><td class="disp" height="100" valign="top"><%= m_KODOSYOKEN %></td></tr>
					</table>
				<% End if %>

			</td>
		</tr>
-->
		<tr>
			<td valign="top" align="left">

				<br>
				<% if gf_empItem(C_T11_SOGOSYOKEN) then %>
					<table class="disp" border="1" width="220">
						<tr><td class="disph" width="100" height="16">総合所見</td></tr>
						<tr><td class="disp" height="100" valign="top"><%= m_SOGOSYOKEN %></td></tr>
					</table>
				<% End if %>

			</td>
			<td valign="top" align="left">

				<BR>
				<% if gf_empItem(C_T11_KOJIN_BIK) then %>
					<table class="disp" border="1" width="220">
						<tr><td class="disph" width="100" height="16">備　考</td></tr>
						<tr><td class="disp" height="100" valign="top"><%= m_KOJIN_BIK %></td></tr>
					</table>
				<% End if %>

			</td>
		</tr>
	</table>

	<% if m_HyoujiFlg = 0 then %>
		<BR>
		表示できるデータがありません<BR>
		<BR>
	<% End if %>

	<BR>
	<input type="button" class="button" value="　閉じる　" onClick="parent.window.close();">

	</div>
	</form>
	</body>
	</html>
<% End Sub %>