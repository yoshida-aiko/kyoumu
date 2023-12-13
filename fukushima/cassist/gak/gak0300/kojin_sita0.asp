<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索詳細
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/kojin_sita0.asp
' 機      能: 検索された学生の詳細を表示する(基本情報)
'-------------------------------------------------------------------------
' 引      数	Session("GAKUSEI_NO")  = 学生番号
'            	Session("Nendo") = 表示年度
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
		w_iRet = f_GetDetailKihon()
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
Function f_GetDetailKihon()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailKihon = 1

	Do
		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO,  "
		w_sSql = w_sSql & " 	A.T13_GAKUSEKI_NO,  "
		w_sSql = w_sSql & " 	A.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	C.M02_GAKKAMEI,  "
		w_sSql = w_sSql & " 	A.T13_SYUSEKI_NO1, " 
		w_sSql = w_sSql & " 	A.T13_CLASS, "
		w_sSql = w_sSql & " 	A.T13_SYUSEKI_NO2,  "
		w_sSql = w_sSql & " 	E.M01_SYOBUNRUIMEI, "
		w_sSql = w_sSql & " 	B.T11_SIMEI,  "
		w_sSql = w_sSql & " 	B.T11_SIMEI_KD, " 
		w_sSql = w_sSql & " 	B.T11_SIMEI_GAIJI,  "
		w_sSql = w_sSql & " 	B.T11_SIMEI_ROMA,  "
		w_sSql = w_sSql & " 	B.T11_HON_ZIP,  "
		w_sSql = w_sSql & " 	B.T11_HON_JUSYO,  "
		w_sSql = w_sSql & " 	B.T11_GEN_ZIP,  "
		w_sSql = w_sSql & " 	B.T11_GEN_JUSYO,  "
		w_sSql = w_sSql & " 	B.T11_GEN_TEL,  "
		w_sSql = w_sSql & " 	D.T09_IMAGE "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A, "
		w_sSql = w_sSql & " 	T11_GAKUSEKI B, "
		w_sSql = w_sSql & " 	M02_GAKKA    C, "
		w_sSql = w_sSql & " 	T09_GAKU_IMG D, "
		w_sSql = w_sSql & " 	M01_KUBUN E  "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 		A.T13_GAKUSEI_NO   = B.T11_GAKUSEI_NO(+) "
		w_sSql = w_sSql & " 	AND	A.T13_NENDO		   = C.M02_NENDO(+) "
		w_sSql = w_sSql & " 	AND A.T13_GAKKA_CD 	   = C.M02_GAKKA_CD(+) "
		w_sSql = w_sSql & " 	AND A.T13_NENDO		   = E.M01_NENDO "
		w_sSql = w_sSql & " 	AND E.M01_DAIBUNRUI_CD = " & C_ZAISEKI				'在籍区分
		w_sSql = w_sSql & " 	AND A.T13_ZAISEKI_KBN  = E.M01_SYOBUNRUI_CD(+) "
		w_sSql = w_sSql & " 	AND A.T13_GAKUSEI_NO   = D.T09_GAKUSEI_NO(+) "
		w_sSql = w_sSql & " 	AND A.T13_GAKUSEI_NO   = '" & Session("GAKUSEI_NO") & "'"
		w_sSql = w_sSql & " 	AND A.T13_NENDO 	   =  " & Session("HyoujiNendo")

		iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailKihon = 99
			Exit Do
		End If

		'//正常終了
		f_GetDetailKihon = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  写真があるか検索
'*  [引数]  なし
'*  [戻値]  True: False
'*  [説明]  
'********************************************************************************
Function f_Photoimg(pGAKUSEI_NO)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_Photoimg = False

	'// NULLなら抜ける(False)
	if trim(pGAKUSEI_NO) = "" then Exit Function

	Do

	    w_sSQL = ""
	    w_sSQL = w_sSQL & " SELECT "
	    w_sSQL = w_sSQL & " T09_IMAGE "
	    w_sSQL = w_sSQL & " FROM T09_GAKU_IMG "
	    w_sSQL = w_sSQL & " WHERE T09_GAKUSEI_NO = '" & cstr(pGAKUSEI_NO) & "'"

		iRet = gf_GetRecordset(w_ImgRs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			Exit Do
		End If

		'// EOFなら抜ける(False)
		if w_ImgRs.Eof then	Exit Do

		'//正常終了
		f_Photoimg = True
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

	m_GAKUSEI_NO   = ""
	m_GAKUSEKI_NO  = ""
	m_GAKUNEN      = ""
	m_GAKKAMEI     = ""
	m_SYUSEKI_NO1  = ""
	m_CLASS        = ""
	m_SYUSEKI_NO2  = ""
	m_SYOBUNRUIMEI = ""
	m_SIMEI        = ""
	m_SIMEI_KD     = ""
	m_SIMEI_GAIJI  = ""
	m_SIMEI_ROMA   = ""
	m_HON_ZIP      = ""
	m_HON_JUSYO    = ""
	m_GEN_ZIP      = ""
	m_GEN_JUSYO    = ""
	m_GEN_TEL      = ""

 	if Not m_Rs.EOF then
		m_GAKUSEI_NO   = m_Rs("T13_GAKUSEI_NO")
		m_GAKUSEKI_NO  = m_Rs("T13_GAKUSEKI_NO")
		m_GAKUNEN      = m_Rs("T13_GAKUNEN")
		m_GAKKAMEI     = m_Rs("M02_GAKKAMEI")
		m_SYUSEKI_NO1  = m_Rs("T13_SYUSEKI_NO1")
		m_CLASS        = m_Rs("T13_CLASS")
		m_SYUSEKI_NO2  = m_Rs("T13_SYUSEKI_NO2")
		m_SYOBUNRUIMEI = m_Rs("M01_SYOBUNRUIMEI")
		m_SIMEI        = m_Rs("T11_SIMEI")
		m_SIMEI_KD     = m_Rs("T11_SIMEI_KD")
		m_SIMEI_GAIJI  = m_Rs("T11_SIMEI_GAIJI")
		m_SIMEI_ROMA   = m_Rs("T11_SIMEI_ROMA")
		m_HON_ZIP      = m_Rs("T11_HON_ZIP")
		m_HON_JUSYO    = m_Rs("T11_HON_JUSYO")
		m_GEN_ZIP      = m_Rs("T11_GEN_ZIP")
		m_GEN_JUSYO    = m_Rs("T11_GEN_JUSYO")
		m_GEN_TEL      = m_Rs("T11_GEN_TEL")
	End if

%>
	<html>
	<head>
	<title>学籍データ参照</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
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
	<br><br>
	<div align="center">

	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><b>●基本情報</b></td>
			<td nowrap><a href="kojin_sita1.asp">●個人情報</a></td>
			<td nowrap><a href="kojin_sita2.asp">●入学情報</a></td>
			<td nowrap><a href="kojin_sita3.asp">●学年情報</a></td>
			<td nowrap><a href="kojin_sita4.asp">●備考・所見</a></td>
			<td nowrap><a href="kojin_sita5.asp">●異動情報</a></td>
		</tr>
	</table>
	<br>

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td valign="top">

				<br>

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_GAKUSEI_NO) then %>
						<tr>
							<td class="disph" width="100" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_5NEN)%></td>
							<td class="disp"><%= m_GAKUSEI_NO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_GAKUSEKI_NO) then %>
						<tr>
							<td class="disph" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_1NEN)%></td>
							<td class="disp"><%= m_GAKUSEKI_NO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_GAKUNEN) then %>
						<tr>
							<td class="disph" height="16">学　　年</td>
							<td class="disp"><%= m_GAKUNEN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_COURCE_CD) then %>
						<tr>
							<td class="disph" height="16">所属学科</td>
							<td class="disp"><%= m_GAKKAMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO1) then %>
						<tr>
							<td class="disph" height="16">出席番号<BR>(学科)</td>
							<td class="disp"><%= m_SYUSEKI_NO1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLASS) then %>
						<tr>
							<td class="disph" height="16">ク ラ ス</td>
							<td class="disp"><%= m_CLASS %>組&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO2) then %>
						<tr>
							<td class="disph" height="16">出席番号<BR>(クラス)</td>
							<td class="disp"><%= m_SYUSEKI_NO2 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_ZAISEKI_KBN) then %>
						<tr>
							<td class="disph" height="16">在籍区分</td>
							<td class="disp"><%= m_SYOBUNRUIMEI %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" rowspan="2">
	
				<br>
				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T11_SIMEI) then %>
						<tr>
							<td class="disph" width="100" height="16">氏　　名</td>
							<td class="disp"><%= m_SIMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_SIMEI_KD) then %>
						<tr>
							<td class="disph" height="16">氏名カナ</td>
							<td class="disp"><%= m_SIMEI_KD %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_SIMEI_GAIJI) then %>
						<tr>
							<td class="disph" height="16">氏名外字</td>
							<td class="disp"><%= m_SIMEI_GAIJI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_SIMEI_ROMA) then %>
						<tr>
							<td class="disph" height="16">氏名ローマ字</td>
							<td class="disp"><%= m_SIMEI_ROMA %>&nbsp</td>
						</tr>
					<% End if %>
				</table>
				<br>

				<div align="center">
				【 写　　真 】
				<table border="1" class="disp">
					<tr><td class="disp">
						<%
						'// 顔写真があるか先に検索する
						w_bRet = ""
						w_bRet = f_Photoimg(Session("GAKUSEI_NO"))
						if w_bRet = True then
							%><IMG SRC="DispBinary.asp?gakuNo=<%= Session("GAKUSEI_NO") %>" width="100" height="120" border="0"><%
						Else
							%><IMG SRC="images/Img0000000000.gif" width="100" height="120" border="0"><%
						End if
						%><br>
					</td></tr>
				</table>
				</div>

			</td>
		</tr>
		<tr>
			<td valign="top">

				<br>【 本 籍 】
				<table border="1" width="260" class="disp">
					<% if gf_empItem(C_T11_HON_ZIP) then %>
						<tr>
							<td class="disph" width="100" height="16">〒</td>
							<td class="disp"><%= m_HON_ZIP %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_HON_JUSYO) then %>
						<tr>
							<td class="disph" height="16" rowspan="3">住　　所</td>
							<td class="disp"><%= m_HON_JUSYO %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

				<BR>
				【 現 住 所 】
				<table class="disp" border="1" width="260">
					<% if gf_empItem(C_T11_GEN_ZIP) then %>
						<tr>
							<td class="disph" width="100" height="16">〒</td>
							<td class="disp"><%= m_GEN_ZIP %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_GEN_JUSYO) then %>
						<tr>
							<td class="disph" height="16">住　　所</td>
							<td class="disp"><%= m_GEN_JUSYO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T11_GEN_TEL) then %>
						<tr>
							<td class="disph" height="16">Ｔ Ｅ Ｌ</td>
							<td class="disp"><%= m_GEN_TEL %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
		</tr>
	</table>

	<BR>
	<input type="button" class="button" value="　閉じる　" onClick="parent.window.close();">

	</div>
	</form>
	</body>
	</html>
<% End Sub %>