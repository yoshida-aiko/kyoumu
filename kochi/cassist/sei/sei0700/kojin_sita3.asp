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
	Public m_RYOSEI_KBN		'寮生区分
	Public m_RYUNEN_FLG		'進級区分

	Public m_HyoujiFlg		'表示ﾌﾗｸﾞ
	Public m_KakoRs			'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ(過去ｸﾗｽ)
	Public mHyoujiNendo		'表示年度

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
		'//過去のクラスを取得
		w_iRet = f_GetDetailKakoClass()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'//表示項目を取得
		w_iRet = f_GetDetailGakunen()
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
'*  [機能]  過去のクラスを取得
'*  [引数]  なし
'*  [戻値]  0:正常終了	1:任意のエラー  99:システムエラー
'*  [説明]  
'********************************************************************************
Function f_GetDetailKakoClass()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailKakoClass = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	T13.T13_NENDO, "
		w_sSql = w_sSql & " 	T13.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	T13.T13_CLASS "
		w_sSql = w_sSql & " FROM T13_GAKU_NEN T13 "
		w_sSql = w_sSql & " WHERE  "
		w_sSql = w_sSql & " 	T13.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		w_sSql = w_sSql & " 	ORDER BY T13.T13_NENDO DESC "

		iRet = gf_GetRecordset(m_KakoRs, w_sSql)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailKakoClass = 99
			Exit Do
		End If

		if m_KakoRs.Eof then
			msMsg = "学年情報取得時にエラーが発生しました"
			f_GetDetailKakoClass = 99
			Exit Do
		End if

		'//正常終了
		f_GetDetailKakoClass = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  表示項目を取得
'*  [引数]  なし
'*  [戻値]  0:正常終了	1:任意のエラー  99:システムエラー
'*  [説明]  
'********************************************************************************
Function f_GetDetailGakunen()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	'// 表示する年度を決める
	wSelNendo = request("selNendo")
	if gf_IsNull(wSelNendo) then
		mHyoujiNendo = Session("HyoujiNendo")
	Else
		mHyoujiNendo = wSelNendo
	End if

	f_GetDetailGakunen = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T13_NENDO, "
		w_sSql = w_sSql & " 	A.T13_GAKUSEKI_NO,"
		w_sSql = w_sSql & " 	A.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	B.M02_GAKKAMEI, "
		w_sSql = w_sSql & " 	A.T13_SYUSEKI_NO1,  "
		w_sSql = w_sSql & " 	A.T13_CLASS,  "
		w_sSql = w_sSql & " 	A.T13_SYUSEKI_NO2,  "
		w_sSql = w_sSql & " 	A.T13_RYOSEI_KBN,  "
		w_sSql = w_sSql & " 	A.T13_RYUNEN_FLG,  "
		w_sSql = w_sSql & " 	A.T13_SINTYO,  "
		w_sSql = w_sSql & " 	A.T13_TAIJYU,  "
		w_sSql = w_sSql & " 	A.T13_CLUB_1,  "
		w_sSql = w_sSql & " 	A.T13_CLUB_2,  "
		w_sSql = w_sSql & " 	A.T13_TOKUKATU, "
		w_sSql = w_sSql & " 	A.T13_TOKUKATU_DET,  "
		w_sSql = w_sSql & " 	A.T13_NENSYOKEN, "
		w_sSql = w_sSql & " 	A.T13_NENBIKO "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A, "
		w_sSql = w_sSql & " 	M02_GAKKA    B "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 	 A.T13_GAKKA_CD   = B.M02_GAKKA_CD(+) "
		w_sSql = w_sSql & "  AND A.T13_NENDO      = B.M02_NENDO(+) "
		w_sSql = w_sSql & "  AND A.T13_NENDO      = " & mHyoujiNendo
		w_sSql = w_sSql & "  AND A.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "

		iRet = gf_GetRecordset(m_Rs, w_sSql)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetailGakunen = 99
			Exit Do
		End If

		'//寮生区分を取得
		if Not gf_GetKubunName(C_NYURYO,m_Rs("T13_RYOSEI_KBN"),Session("HyoujiNendo"),m_RYOSEI_KBN) then Exit Do

		'//進級区分を取得
		Select Case m_Rs("T13_RYUNEN_FLG")
			Case 0: m_RYUNEN_FLG = ""
			Case 1: m_RYUNEN_FLG = C_SHINKYU_NO
		End Select

		'//正常終了
		f_GetDetailGakunen = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  部活名を取得する
'*  [引数]  p_sClubCd:部活CD
'*  [戻値]  f_GetClubName：部活名
'*  [説明]  
'********************************************************************************
Function f_GetClubName(p_sClubCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubName = ""
	w_sClubName = ""

	Do

		'//部活CDが空の時
		If trim(gf_SetNull2String(p_sClubCd)) = "" Then
			Exit Do
		End If

		'//部活動情報取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & mHyoujiNendo
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD=" & p_sClubCd

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//部活名
			w_sClubName = rs("M17_BUKATUDOMEI")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	f_GetClubName = w_sClubName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

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

	m_NENDO		 	= ""
	m_GAKUSEKI_NO	= ""
	m_GAKUNEN		= ""
	m_GAKKAMEI	 	= ""
	m_SYUSEKI_NO1	= ""
	m_CLASS		 	= ""
	m_SYUSEKI_NO2	= ""
	m_SINTYO		= ""
	m_TAIJYU		= ""
	m_CLUB_1		= ""
	m_CLUB_2		= ""
	m_TOKUKATU	 	= ""
	m_TOKUKATU_DET  = ""
	m_NENSYOKEN	 	= ""
	m_F_NENBIKO	 	= ""

	if Not m_Rs.Eof Then
		m_NENDO		 	= m_Rs("T13_NENDO")
		m_GAKUSEKI_NO	= m_Rs("T13_GAKUSEKI_NO")
		m_GAKUNEN		= m_Rs("T13_GAKUNEN")
		m_GAKKAMEI	 	= m_Rs("M02_GAKKAMEI")
		m_SYUSEKI_NO1	= m_Rs("T13_SYUSEKI_NO1")
		m_CLASS		 	= m_Rs("T13_CLASS")
		m_SYUSEKI_NO2	= m_Rs("T13_SYUSEKI_NO2")
		m_SINTYO		= m_Rs("T13_SINTYO")
		m_TAIJYU		= m_Rs("T13_TAIJYU")
		m_CLUB_1		= f_GetClubName(gf_SetNull2String(m_Rs("T13_CLUB_1")))
		m_CLUB_2		= f_GetClubName(gf_SetNull2String(m_Rs("T13_CLUB_2")))
		m_TOKUKATU	 	= m_Rs("T13_TOKUKATU")
		m_TOKUKATU_DET  = m_Rs("T13_TOKUKATU_DET")
		m_NENSYOKEN	 	= m_Rs("T13_NENSYOKEN")
		m_F_NENBIKO	 	= m_Rs("T13_NENBIKO")
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
		//**************************************
		//*   年度ｾﾚｸﾄﾎﾞｯｸｽが変更されたとき
		//**************************************
		function jf_ChangSelect(){

			document.frm.submit();

		}

	//-->
	</script>
	</head>

	<body>
	<form action="kojin_sita3.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">●基本情報</a></td>
			<td nowrap><a href="kojin_sita1.asp">●個人情報</a></td>
			<td nowrap><a href="kojin_sita2.asp">●入学情報</a></td>
			<td nowrap><b>●学年情報</b></td>
			<td nowrap><a href="kojin_sita4.asp">●備考・所見</a></td>
			<td nowrap><a href="kojin_sita5.asp">●異動情報</a></td>
		</tr>
	</table>
	<br>

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td colspan="3">
				<span class="msg"><font size="2">※ 処理年度を変更すると、過去の学年情報を見ることができます<BR></font></span>
			</td>
		</tr>
		<tr>
			<td valign="top" align="left">

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_NENDO) then %>
						<tr>
							<td class="disph" width="100">処理年度</td>
							<td class="disp"><select name="selNendo" onChange="jf_ChangSelect();">
												<% do until m_KakoRs.Eof 
													wSelected = ""
													if Cint(mHyoujiNendo) = Cint(m_KakoRs("T13_NENDO")) then
														wSelected = "selected"
													End if
													%>
													<option value="<%=m_KakoRs("T13_NENDO")%>" <%=wSelected%>><%=m_KakoRs("T13_NENDO")%>年度
												<% m_KakoRs.MoveNext : Loop %>
											</select></td>
						</tr>
<!--
						<tr>
							<td class="disph" width="100" height="16">処理年度</td>
							<td class="disp"><%= m_NENDO %>&nbsp</td>
						</tr>
-->
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
					<% if gf_empItem(C_T13_CLASS) then %>
						<tr>
							<td class="disph" height="16">クラス</td>
							<td class="disp"><%= m_CLASS %>組&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO1) then %>
						<tr>
							<td class="disph" height="16">出席番号(学科)</td>
							<td class="disp"><%= m_SYUSEKI_NO1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO2) then %>
						<tr>
							<td class="disph" height="16">出席番号(クラス)</td>
							<td class="disp"><%= m_SYUSEKI_NO2 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_RYOSEI_KBN) then %>
						<tr>
							<td class="disph" height="16">寮生区分</td>
							<td class="disp"><%= m_RYOSEI_KBN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_RYUNEN_FLG) then %>
						<tr>
							<td class="disph" height="16">進級区分</td>
							<td class="disp"><%= m_RYUNEN_FLG %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" align="left">

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_SINTYO) then %>
						<tr>
							<td class="disph" width="100" height="16">身　　長</td>
							<td class="disp"><%= m_SINTYO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TAIJYU) then %>
						<tr>
							<td class="disph" height="16">体　　重</td>
							<td class="disp"><%= m_TAIJYU %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_1) then %>
						<tr>
							<td class="disph" height="16">クラブ活動１</td>
							<td class="disp"><%= m_CLUB_1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_2) then %>
						<tr>
							<td class="disph" height="16">クラブ活動２</td>
							<td class="disp"><%= m_CLUB_2 %>&nbsp</td>
						</tr>
					<% End if %>
<!--
					<% if gf_empItem(C_T13_TOKUKATU) then %>
						<tr>
							<td class="disph" height="16">特別活動</td>
							<td class="disp"><%= m_TOKUKATU %>&nbsp</td>
						</tr>
-->
					<% End if %>
					<% if gf_empItem(C_T13_TOKUKATU_DET) then %>
						<tr>
							<td class="disph" height="16">特別活動詳細</td>
							<td class="disp"><%= m_TOKUKATU_DET %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" align="left">

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_NENSYOKEN) then %>
						<tr><td class="disph" width="220" height="16">指導上参考となる諸事項</td></tr>
						<tr><td class="disp" valign="top" height="220"><%= m_NENSYOKEN %><br><br></td></tr>
					<% End if %>
					<% if gf_empItem(C_T13_NENBIKO) then %>
						<tr><td class="disph" width="100" height="16">備 考</td></tr>
						<tr><td class="disp" valign="top" height="100"><%= m_F_NENBIKO %></td></tr>
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