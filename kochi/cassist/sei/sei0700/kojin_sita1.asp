<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索詳細
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/kojin_sita1.asp
' 機      能: 検索された学生の詳細を表示する(個人情報)
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
	Public m_SEIBETU		'性別
	Public m_BLOOD			'血液型
	Public m_RH				'RH
	Public m_HOG_ZOKU		'保護者続柄
	Public m_HOS_ZOKU		'保証人続柄

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
		w_iRet = f_GetDetailKojin()
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
Function f_GetDetailKojin()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailKojin = 1

	Do
		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T11_SEIBETU,  "
		w_sSql = w_sSql & " 	A.T11_SEINENBI,  "
		w_sSql = w_sSql & " 	A.T11_KETUEKI,  "
		w_sSql = w_sSql & " 	A.T11_RH,  "
		w_sSql = w_sSql & " 	A.T11_HOG_SIMEI,  "
		w_sSql = w_sSql & " 	A.T11_HOG_SIMEI_K,  "
		w_sSql = w_sSql & " 	A.T11_HOG_ZOKU,  "
		w_sSql = w_sSql & " 	A.T11_HOG_ZIP,  "
		w_sSql = w_sSql & " 	A.T11_HOG_JUSYO,  "
		w_sSql = w_sSql & " 	A.T11_HOGO_TEL,  "
		w_sSql = w_sSql & " 	A.T11_HOS_SIMEI,  "
		w_sSql = w_sSql & " 	A.T11_HOS_SIMEI_K,  "
		w_sSql = w_sSql & " 	A.T11_HOS_ZOKU,  "
		w_sSql = w_sSql & " 	A.T11_HOS_ZIP,  "
		w_sSql = w_sSql & " 	A.T11_HOS_JUSYO,  "
		w_sSql = w_sSql & " 	A.T11_HOS_TEL "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T11_GAKUSEKI A "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & "  	A.T11_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "

		iRet = gf_GetRecordset(m_Rs, w_sSql)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailKojin = 99
			Exit Do
		End If

		'// 性別を取得
		if Not gf_GetKubunName(C_SEIBETU,m_Rs("T11_SEIBETU"),Session("HyoujiNendo"),m_SEIBETU) then Exit Do

		'// 血液型を取得
		if Not gf_GetKubunName(C_BLOOD,m_Rs("T11_KETUEKI"),Session("HyoujiNendo"),m_BLOOD) then Exit Do

		'// RHを取得
		if Not gf_GetKubunName(C_RH,m_Rs("T11_RH"),Session("HyoujiNendo"),m_RH) then Exit Do

		'// 保護者続柄を取得
		if Not gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_HOG_ZOKU"),Session("HyoujiNendo"),m_HOG_ZOKU) then Exit Do

		'// 保証人続柄を取得
		if Not gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_HOS_ZOKU"),Session("HyoujiNendo"),m_HOS_ZOKU) then Exit Do

		'//正常終了
		f_GetDetailKojin = 0
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

	m_HyoujiFlg = 0 		'<!-- 表示フラグ（0:なし  1:あり）

	m_SEINENBI 		= ""
	m_HOG_SIMEI		= ""
	m_HOG_SIMEI_K	= ""
	m_HOG_ZIP 		= ""
	m_HOG_JUSYO		= ""
	m_HOGO_TEL 		= ""
	m_HOS_SIMEI		= ""
	m_HOS_SIMEI_K	= ""
	m_HOS_ZIP		= ""
	m_HOS_JUSYO		= ""
	m_HOS_TEL		= ""

 	if Not m_Rs.EOF then
		m_SEINENBI 		= m_Rs("T11_SEINENBI") 
		m_HOG_SIMEI		= m_Rs("T11_HOG_SIMEI") 
		m_HOG_SIMEI_K	= m_Rs("T11_HOG_SIMEI_K") 
		m_HOG_ZIP 		= m_Rs("T11_HOG_ZIP") 
		m_HOG_JUSYO		= m_Rs("T11_HOG_JUSYO") 
		m_HOGO_TEL 		= m_Rs("T11_HOGO_TEL") 
		m_HOS_SIMEI		= m_Rs("T11_HOS_SIMEI") 
		m_HOS_SIMEI_K	= m_Rs("T11_HOS_SIMEI_K") 
		m_HOS_ZIP		= m_Rs("T11_HOS_ZIP")
		m_HOS_JUSYO		= m_Rs("T11_HOS_JUSYO")
		m_HOS_TEL		= m_Rs("T11_HOS_TEL")
	End if

%>

	<html>
	<head>
	<title>学籍データ参照</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<script language="javascript">
	<!--
		function sbmt(m,i) {
			document.forms[0].mode.value = m;
			document.forms[0].id.value = i;
			document.forms[0].submit();
		}
	//-->
	</script>
	<style type="text/css">
	<!--
		a:link { color:#cc8866; text-decoration:none; }
		a:visited { color:#cc8866; text-decoration:none; }
		a:active { color:#888866; text-decoration:none; }
		a:hover { color:#888866; text-decoration:underline; }
		b { color:#88bbbb; font-weight: bold; font-size:14px}
	//-->
	</style>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">●基本情報</a></td>
			<td nowrap><b>●個人情報</b></td>
			<td nowrap><a href="kojin_sita2.asp">●入学情報</a></td>
			<td nowrap><a href="kojin_sita3.asp">●学年情報</a></td>
			<td nowrap><a href="kojin_sita4.asp">●備考・所見</a></td>
			<td nowrap><a href="kojin_sita5.asp">●異動情報</a></td>
		</tr>
	</table>
	<br>

	
	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td valign="top" align="left">
				<br>

					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_SEIBETU) then %>
							<tr>
								<td width="100" height="16" class="disph">性　　別</td>
								<td class="disp"><%= m_SEIBETU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_SEINENBI) then %>
							<tr>
								<td height="16" class="disph">生年月日</td>
								<td class="disp"><%= m_SEINENBI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KETUEKI) then %>
							<tr>
								<td height="16" class="disph">血 液 型</td>
								<td class="disp"><%= m_BLOOD %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_RH) then %>
							<tr>
								<td height="16" class="disph">Ｒ　　Ｈ</td>
								<td class="disp"><%= m_RH %>&nbsp</td>
							</tr>
						<% End if %>
					</table>

			</td>

			<td valign="top" align="left">
				【 保護者情報 】

					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_HOG_SIMEI) then %>
							<tr>
								<td class="disph" width="100" height="16"><font color="white">氏　　名</font></td>
								<td class="disp"><%= m_HOG_SIMEI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_SIMEI_K) then %>
							<tr>
								<td class="disph" height="16"><font color="white">カ　　ナ</font></td>
								<td class="disp"><%= m_HOG_SIMEI_K %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_ZOKU) then %>
							<tr>
								<td class="disph" height="16"><font color="white">続　　柄</font></td>
								<td class="disp"><%= m_HOG_ZOKU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_ZIP) then %>
							<tr>
								<td class="disph" height="16"><font color="white">〒</font></td>
								<td class="disp"><%= m_HOG_ZIP %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_JUSYO) then %>
							<tr>
								<td class="disph" height="16"><font color="white">住　　所</font></td>
								<td class="disp"><%= m_HOG_JUSYO %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_TEL) then %>
							<tr>
								<td class="disph" height="16"><font color="white">Ｔ Ｅ Ｌ</font></td>
								<td class="disp"><%= m_HOGO_TEL %>&nbsp</td>
							</tr>
						<% End if %>
					</table>

			</td>

			<td valign="top" align="left">
				【 保証人情報 】

					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_HOS_SIMEI) then %>
							<tr>
								<td class="disph" width="100" height="16">氏　　名</td>
								<td class="disp"><%= m_HOS_SIMEI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_SIMEI_K) then %>
							<tr>
								<td class="disph" height="16">カ　　ナ</td>
								<td class="disp"><%= m_HOS_SIMEI_K %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_ZOKU) then %>
							<tr>
								<td class="disph" height="16">続　　柄</td>
								<td class="disp"><%= m_HOS_ZOKU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_ZIP) then %>
							<tr>
								<td class="disph" height="16">〒</td>
								<td class="disp"><%= m_HOS_ZIP %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_JUSYO) then %>
							<tr>
								<td class="disph" height="16">住　　所</td>
								<td class="disp"><%= m_HOS_JUSYO %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_TEL) then %>
							<tr>
								<td class="disph" height="16">Ｔ Ｅ Ｌ</td>
								<td class="disp"><%= m_HOS_TEL %>&nbsp</td>
							</tr>
						<% End if %>
					</table>

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