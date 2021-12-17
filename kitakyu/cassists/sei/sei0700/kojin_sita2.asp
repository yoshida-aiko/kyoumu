<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索詳細
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/kojin_sita2.asp
' 機      能: 検索された学生の詳細を表示する(入学情報)
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
	Public m_NYUGAKU_KBN	'入学区分
	Public m_HyoujiFlg		'表示ﾌﾗｸﾞ
	Public m_TYUGAKKOMEI	'中学校名
	Public m_NYU_GAKKA		'学科名
	Public m_KURABUMEI		'クラブ名


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
		w_iRet = f_GetDetailNyugaku()
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
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
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
Function f_GetDetailNyugaku()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailNyugaku = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql &  vbCrLf &" SELECT "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUNENDO,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKU_KBN, " 
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKUBI,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKUBI,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYU_GAKKA, "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_JUKEN_NO,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYU_SEISEKI, " 
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYUGAKKO_CD, "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYUSOTUGYOBI,  "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYU_CLUB, "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYU_CLUB_SYOSAI "
		w_sSql = w_sSql &  vbCrLf &" FROM  "
		w_sSql = w_sSql &  vbCrLf &" 	T11_GAKUSEKI A "
		w_sSql = w_sSql &  vbCrLf &" WHERE "
		w_sSql = w_sSql &  vbCrLf &" 	A.T11_GAKUSEI_NO  = '" & Session("GAKUSEI_NO") & "' "

'		w_sSql = ""
'		w_sSql = w_sSql &  vbCrLf &" SELECT "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUNENDO,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKU_KBN, " 
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYUGAKUBI,  "
'		w_sSql = w_sSql &  vbCrLf &" 	C.M02_GAKKAMEI,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_JUKEN_NO,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_NYU_SEISEKI, " 
'		w_sSql = w_sSql &  vbCrLf &" 	D.M13_TYUGAKKOMEI,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYUSOTUGYOBI,  "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYU_CLUB, "
'		w_sSql = w_sSql &  vbCrLf &" 	A.T11_TYU_CLUB_SYOSAI "
'		w_sSql = w_sSql &  vbCrLf &" FROM  "
'		w_sSql = w_sSql &  vbCrLf &" 	T11_GAKUSEKI A, "
'		w_sSql = w_sSql &  vbCrLf &" 	M02_GAKKA    C, "
'		w_sSql = w_sSql &  vbCrLf &" 	M13_TYUGAKKO D "
'		w_sSql = w_sSql &  vbCrLf &" WHERE "
'		w_sSql = w_sSql &  vbCrLf &" 		A.T11_TYUGAKKO_CD = D.M13_TYUGAKKO_CD(+) "
'		w_sSql = w_sSql &  vbCrLf &" 	AND A.T11_NYU_GAKKA   = C.M02_GAKKA_CD(+) "
'		w_sSql = w_sSql &  vbCrLf &" 	AND D.M13_NENDO       =  " & Session("HyoujiNendo")
'		w_sSql = w_sSql &  vbCrLf &" 	AND C.M02_NENDO       =  " & Session("HyoujiNendo")
'		w_sSql = w_sSql &  vbCrLf &" 	AND A.T11_GAKUSEI_NO  = '" & Session("GAKUSEI_NO") & "' "

		iRet = gf_GetRecordset(m_Rs, w_sSql)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailNyugaku = 99
			Exit Do
		End If

		'//入学区分を取得
		if Not gf_GetKubunName(C_NYUGAKU,m_Rs("T11_NYUGAKU_KBN"),Session("HyoujiNendo"),m_NYUGAKU_KBN) then Exit Do

		'//中学校名を取得
		if Not f_GetTyugakkoMei(m_Rs("T11_TYUGAKKO_CD"),m_TYUGAKKOMEI) then Exit Do

		'//学科名を取得
		if Not f_GetGakkaMei(m_Rs("T11_NYU_GAKKA"),m_NYU_GAKKA) then Exit Do

		'//クラブ名を取得
		if Not f_GetKurabuMei(m_Rs("T11_TYU_CLUB"),m_KURABUMEI) then Exit Do

		'//正常終了
		f_GetDetailNyugaku = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  中学校名を取得
'*  [引数]  なし
'*  [戻値]  True: False
'*  [説明]  
'********************************************************************************
Function f_GetTyugakkoMei(pKey,pTYUGAKKOMEI)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetTyugakkoMei = False

	'// NULLなら抜ける(False)
	if trim(pKey) = "" then Exit Function

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
    w_sSQL = w_sSQL & " 	M13_TYUGAKKOMEI "
    w_sSQL = w_sSQL & " FROM M13_TYUGAKKO "
    w_sSQL = w_sSQL & " WHERE M13_TYUGAKKO_CD = '" & pKey & "'"
    w_sSQL = w_sSQL & " 	AND M13_NENDO = " & Session("HyoujiNendo")

	iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit Function
	End If

	'// EOFなら抜ける(False)
	if w_Rs.Eof then 
		f_GetTyugakkoMei = True
		Exit Function
	End if

	'// 中学校名
	pTYUGAKKOMEI = w_Rs("M13_TYUGAKKOMEI")

    '// 終了処理
    If Not IsNull(w_Rs) Then gf_closeObject(w_Rs)

	'//正常終了
	f_GetTyugakkoMei = True

End Function


'********************************************************************************
'*  [機能]  学科名を取得
'*  [引数]  なし
'*  [戻値]  True: False
'*  [説明]  
'********************************************************************************
Function f_GetGakkaMei(pKey,pNYU_GAKKA)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetGakkaMei = False

	'// NULLなら抜ける(False)
	if trim(pKey) = "" then Exit Function

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
    w_sSQL = w_sSQL & " 	M02_GAKKAMEI "
    w_sSQL = w_sSQL & " FROM M02_GAKKA "
    w_sSQL = w_sSQL & " WHERE M02_GAKKA_CD = '" & pKey & "'"
    w_sSQL = w_sSQL & " 	AND M02_NENDO = " & Session("HyoujiNendo")

	iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit Function
	End If

	'// EOFなら抜ける(False)
	if w_Rs.Eof then 
		f_GetGakkaMei = True
		Exit Function
	End if

	'// 学科名
	pNYU_GAKKA = w_Rs("M02_GAKKAMEI")

    '// 終了処理
    If Not IsNull(w_Rs) Then gf_closeObject(w_Rs)

	'//正常終了
	f_GetGakkaMei = True

End Function


'********************************************************************************
'*  [機能]  クラブ名を取得
'*  [引数]  なし
'*  [戻値]  True: False
'*  [説明]  
'********************************************************************************
Function f_GetKurabuMei(pKey,pKURABUMEI)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetGakkaMei = False

	'// NULLなら抜ける(False)
	if trim(pKey) = "" then Exit Function

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
    w_sSQL = w_sSQL & " 	M17_BUKATUDOMEI "
    w_sSQL = w_sSQL & " FROM M17_BUKATUDO "
    w_sSQL = w_sSQL & " WHERE M17_BUKATUDO_CD = '" & pKey & "'"
    w_sSQL = w_sSQL & " 	AND M17_NENDO = " & Session("HyoujiNendo")

	iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit Function
	End If

	'// EOFなら抜ける(False)
	if w_Rs.Eof then 
		f_GetKurabuMei = True
		Exit Function
	End if

	'// クラブ名
	pKURABUMEI = w_Rs("M17_BUKATUDOMEI")

    '// 終了処理
    If Not IsNull(w_Rs) Then gf_closeObject(w_Rs)

	'//正常終了
	f_GetKurabuMei = True

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

	m_NYUNENDO 		  = ""
	m_F_NYUGAKUBI 	  = ""
'	m_GAKKAMEI 		  = ""
	m_JUKEN_NO		  = ""
	m_NYU_SEISEKI 	  = ""
'	m_TYUGAKKOMEI 	  = ""
	m_TYUSOTUGYOBI 	  = ""
'	m_BUKATUDOMEI 	  = ""
	m_TYU_CLUB_SYOSAI = ""

	if Not m_Rs.Eof Then
		m_NYUNENDO 		  = m_Rs("T11_NYUNENDO") 
		m_NYUGAKUBI 	  = m_Rs("T11_NYUGAKUBI") 
'		m_GAKKAMEI 		  = m_Rs("M02_GAKKAMEI") 
		m_JUKEN_NO		  = m_Rs("T11_JUKEN_NO") 
		m_NYU_SEISEKI 	  = m_Rs("T11_NYU_SEISEKI") 
'		m_TYUGAKKOMEI 	  = m_Rs("M13_TYUGAKKOMEI") 
		m_TYUSOTUGYOBI 	  = m_Rs("T11_TYUSOTUGYOBI") 
'		m_BUKATUDOMEI 	  = m_Rs("M17_BUKATUDOMEI") 
		m_TYU_CLUB_SYOSAI = m_Rs("T11_TYU_CLUB_SYOSAI") 
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
	-->
	</style>
	<script language="javascript">
	<!--
		function sbmt(m,i) {
			document.forms[0].mode.value = m;
			document.forms[0].id.value = i;
			document.forms[0].submit();
		}
	-->
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
			<td nowrap><b>●入学情報</b></td>
			<td nowrap><a href="kojin_sita3.asp">●学年情報</a></td>
			<td nowrap><a href="kojin_sita4.asp">●備考・所見</a></td>
			<td nowrap><a href="kojin_sita5.asp">●異動情報</a></td>
		</tr>
	</table>
	<br>
	

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td width="60">&nbsp</td>
			<td valign="top" align="left">

				<br>
				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T11_NYUGAKU_KBN) then %>
						<tr>
							<td class="disph" width="100" height="16">入学区分</td>
							<td class="disp"><%= m_NYUGAKU_KBN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_NYUNENDO) then %>
						<tr>
							<td class="disph" height="16">入学年度</td>
							<td class="disp"><%= m_NYUNENDO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_NYUGAKUBI) then %>
						<tr>
							<td class="disph" height="16">入 学 日</td>
							<td class="disp"><%= m_NYUGAKUBI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_NYU_GAKKA) then %>
						<tr>
							<td class="disph" height="16">学　　科</td>
							<td class="disp"><%= m_NYU_GAKKA %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" align="left">

				<br>
				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T11_JUKEN_NO) then %>
						<tr>
							<td class="disph" width="100" height="16">受験番号</td>
							<td class="disp"><%= m_JUKEN_NO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_NYU_SEISEKI) then %>
						<tr>
							<td class="disph" height="16">入学成績</td>
							<td class="disp"><%= m_NYU_SEISEKI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_TYUGAKKO_CD) then %>
						<tr>
							<td class="disph" height="16">中学校名</td>
							<td class="disp"><%= m_TYUGAKKOMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TYUSOTUGYOBI) then %>
						<tr>
							<td class="disph" height="16">卒 業 日</td>
							<td class="disp"><%= m_TYUSOTUGYOBI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_TYU_CLUB) then %>
						<tr>
							<td class="disph" height="16">ク ラ ブ</td>
							<td class="disp"><%= m_KURABUMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_TYU_CLUB_SYOSAI) then %>
						<tr>
							<td class="disph" height="16">クラブ詳細</td>
							<td class="disp"><%= m_TYU_CLUB_SYOSAI %>&nbsp</td>
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