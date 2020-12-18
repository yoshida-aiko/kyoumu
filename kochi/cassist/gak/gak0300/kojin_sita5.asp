<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索詳細
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/kojin_sita5.asp
' 機      能: 検索された学生の詳細を表示する(異動情報)
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
		w_iRet = f_GetDetailIdo()
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
Function f_GetDetailIdo()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailIdo = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T13_IDOU_NUM, "
		w_sSql = w_sSql & " 	A.T13_NENDO, "
		w_sSql = w_sSql & " 	A.T13_GAKUNEN, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_1, 	A.T13_IDOU_BI_1, 	A.T13_IDOU_BIK_1, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_2, 	A.T13_IDOU_BI_2, 	A.T13_IDOU_BIK_2, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_3, 	A.T13_IDOU_BI_3, 	A.T13_IDOU_BIK_3, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_4, 	A.T13_IDOU_BI_4, 	A.T13_IDOU_BIK_4, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_5, 	A.T13_IDOU_BI_5, 	A.T13_IDOU_BIK_5, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_6, 	A.T13_IDOU_BI_6, 	A.T13_IDOU_BIK_6, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_7, 	A.T13_IDOU_BI_7, 	A.T13_IDOU_BIK_7, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_8, 	A.T13_IDOU_BI_8, 	A.T13_IDOU_BIK_8 "
		w_sSql = w_sSql & " FROM "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		w_sSql = w_sSql & " AND A.T13_NENDO = " & Session("HyoujiNendo")

		iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailIdo = 99
			Exit Do
		End If


		'//正常終了
		f_GetDetailIdo = 0
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

	m_IDOU_NUM		= ""
	m_NENDO			= ""
	m_GAKUNEN		= ""

	if Not m_Rs.Eof then
		m_IDOU_NUM		= m_Rs("T13_IDOU_NUM")
		m_NENDO			= m_Rs("T13_NENDO")
		m_GAKUNEN		= m_Rs("T13_GAKUNEN")
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
			<td nowrap><a href="kojin_sita4.asp">●備考・所見</a></td>
			<td nowrap><b>●異動情報</b></td>
		</tr>
	</table>
	<br>

	<table border="1" class=hyo>
		<tr>
			<% if gf_empItem(C_T13_IDOU_KBN) then %>
				<th width="80"  class="header">事由</th>
			<% End if %>
			<% if gf_empItem(C_T13_NENDO) then %>
				<th width="100" class="header">年度</th>
			<% End if %>
			<% if gf_empItem(C_T13_GAKUNEN) then %>
				<th width="100" class="header">学年</th>
			<% End if %>
			<% if gf_empItem(C_T13_IDOU_BI) then %>
				<th width="160" class="header">日付または期間</th>
			<% End if %>
			<% if gf_empItem(C_T13_IDOU_BIK) then %>
				<th width="140" class="header">備考</th>
			<% End if %>
		</tr>
		<%
			'// 移動回数分,回す
			if Cint(gf_SetNull2Zero(m_IDOU_NUM)) > 0 then
				i_line = 1
				Do Until i_line > Cint(m_IDOU_NUM)

					'// 事由を取得
					m_IDOU_KBN = ""
					w_IDOU_KBN_NO = m_Rs("T13_IDOU_KBN_" & i_line)
					Call gf_GetKubunName(C_IDO,w_IDOU_KBN_NO,Session("HyoujiNendo"),m_IDOU_KBN)
					Call gs_cellPtn(w_cell)
				%>
					<tr>
						<% if gf_empItem(C_T13_IDOU_KBN) then %>
							<td class="<%=w_cell%>"><%= m_IDOU_KBN %></td>
						<% End if %>
						<% if gf_empItem(C_T13_NENDO) then %>
							<td class="<%=w_cell%>"><%= m_NENDO %></td>
						<% End if %>
						<% if gf_empItem(C_T13_GAKUNEN) then %>
							<td class="<%=w_cell%>"><%= m_GAKUNEN %></td>
						<% End if %>
						<% if gf_empItem(C_T13_IDOU_BI) then %>
							<td class="<%=w_cell%>"><%= m_Rs("T13_IDOU_BI_" & i_line ) %></td>
						<% End if %>
						<% if gf_empItem(C_T13_IDOU_BIK) then %>
							<td class="<%=w_cell%>"><%= m_Rs("T13_IDOU_BIK_" & i_line ) %></td>
						<% End if %>
					</tr>
				<%
					i_line = i_line + 1
				Loop
			End if
		%>
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