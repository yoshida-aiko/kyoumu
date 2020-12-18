<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索詳細
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0310/kojin_sita1.asp
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
' 作      成: 2001/12/01 岡田
' 変      更: 
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
		w_sSql = w_sSql & " 	A.T11_NYUNENDO,  "	'入学年度
		w_sSql = w_sSql & " 	A.T11_SEIBETU,  "	'性別
		w_sSql = w_sSql & "		A.T11_NYUGAKU_KBN, " '入学区分
		w_sSql = w_sSql & " 	A.T11_SEINENBI,  "	'生年月日
		w_sSql = w_sSql & " 	A.T11_KETUEKI,  "	'血液型
		w_sSql = w_sSql & " 	A.T11_RH,  "		'RH
		w_sSql = w_sSql & "		A.T11_SYUSSINKO,  "	'出身校
		w_sSql = w_sSql & "		A.T11_SYUSSINKOKU,  "	'出身国
		w_sSql = w_sSql & "		A.T11_RYUGAKU_KBN,  "	'留学区分
		w_sSql = w_sSql & " 	A.T11_HOG_SIMEI,  "		'保護者氏名
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
		w_sSql = w_sSql & " 	A.T11_HOS_TEL, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_1, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_1, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_2, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_2, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_3, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_3, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_4, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_4,"
		w_sSql = w_sSql & "		A.T11_KAZOKU_5,	"
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_5, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_6,  "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_6, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_7,  "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_7, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_8,  "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_8,"
		w_sSql = w_sSql & " 	A.T11_HOG_SEINEIBI,"
		w_sSql = w_sSql & " 	A.T11_HOS_SEINEIBI,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_1,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_2,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_3,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_4,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_5,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_6,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_7,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_8 "
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

		'// 留学区分を取得
		if Not gf_GetKubunName(C_RYUGAKU,m_Rs("T11_RYUGAKU_KBN"),Session("HyoujiNendo"),m_RYUGAKU_KBN) then Exit Do

		'//正常終了
		f_GetDetailKojin = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  出身校を取得
'*  [引数]  学校CD
'*  [戻値]  学校名
'********************************************************************************
Function f_GetSyussinko(p_Nendo,p_GakkouCd)

	On Error Resume Next
	Err.Clear

	f_GetSyussinko = ""

	'// 年度・学校CDがNULLだったら抜ける
	if gf_IsNull(p_Nendo) or gf_IsNull(p_GakkouCd) then
		Exit Function
	End if

	w_sSql = "" 
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	M31_GAKKOMEI "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & " 	M31_SYUSSINKO "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	M31_NENDO    = '" & p_Nendo & "'"
	w_sSql = w_sSql & " AND M31_GAKKO_CD = '" & p_GakkouCd & "'"

	iRet = gf_GetRecordset(w_Rs, w_sSql)
	If iRet = 0 Then
		if Not w_Rs.Eof then
			f_GetSyussinko = w_Rs("M31_GAKKOMEI")
		End if
	End If

	p_oRecordset.Close
	Set p_oRecordset = Nothing

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
	m_KAZOKU_1      = ""
	m_KAZOKU_ZOKU_1 = ""
	m_KAZOKU_2      = ""
	m_KAZOKU_ZOKU_2 = ""
	m_KAZOKU_3      = ""
	m_KAZOKU_ZOKU_3 = ""
	m_KAZOKU_4      = ""
	m_KAZOKU_ZOKU_4 = ""
	m_KAZOKU_5      = ""
	m_KAZOKU_ZOKU_5 = ""
	m_KAZOKU_6      = ""
	m_KAZOKU_ZOKU_6 = ""
	m_KAZOKU_7      = ""
	m_KAZOKU_ZOKU_7 = ""
	m_KAZOKU_8      = ""
	m_KAZOKU_ZOKU_8 = ""
	m_SYUSSINKO    = ""
	m_SYUSSINKOKU   = ""
'	m_RYUGAKU_KBN   = ""

	m_Ken = ""
	m_SITYOSONCD = ""
	m_SITYOSONMEI = ""
	m_TYOIKIMEI = ""
	m_Ken2 = ""
	m_SITYOSONCD2 = ""
	m_SITYOSONMEI2 = ""
	m_TYOIKIMEI2 = ""

	m_HOG_SEINEIBI      = ""
	m_HOS_SEINEIBI      = ""
	m_KAZOKU_SEINEIBI_2 = ""
	m_KAZOKU_SEINEIBI_3 = ""
	m_KAZOKU_SEINEIBI_4 = ""
	m_KAZOKU_SEINEIBI_5 = ""
	m_KAZOKU_SEINEIBI_6 = ""
	m_KAZOKU_SEINEIBI_7 = ""
	m_KAZOKU_SEINEIBI_8 = ""

	'// 家族続柄１～８を取得
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_1"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_1)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_2"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_2)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_3"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_3)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_4"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_4)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_5"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_5)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_6"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_6)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_7"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_7)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_8"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_8)

	Call gf_ComvZip(m_Rs("T11_HOG_ZIP"),m_Ken,m_SITYOSONCD,m_SITYOSONMEI,m_TYOIKIMEI,Session("HyoujiNendo"))
	Call gf_ComvZip(m_Rs("T11_HOS_ZIP"),m_Ken2,m_SITYOSONCD2,m_SITYOSONMEI2,m_TYOIKIMEI2,Session("HyoujiNendo"))


 	if Not m_Rs.EOF then
		m_SEINENBI 		= m_Rs("T11_SEINENBI") 
		m_HOG_SIMEI		= m_Rs("T11_HOG_SIMEI") 
		m_HOG_SIMEI_K	= m_Rs("T11_HOG_SIMEI_K")
		m_HOG_ZIP 		= m_Rs("T11_HOG_ZIP")

	'/* 住所に県、市町村名が存在していた場合削除して再度付け直す。*/ Add 2002.04.30 okada
		m_HOG_JUSYO		= m_Rs("T11_HOG_JUSYO")
		m_HOG_JUSYO     = Replace(m_HOG_JUSYO,m_Ken,"")
		m_HOG_JUSYO     = Replace(m_HOG_JUSYO,m_SITYOSONMEI,"")

		m_HOG_JUSYO		= m_Ken & m_SITYOSONMEI & m_HOG_JUSYO 'm_SITYOSONMEI & Replace(m_Rs("T11_HOG_JUSYO"),m_SITYOSONMEI,"")

		
		m_HOGO_TEL 		= m_Rs("T11_HOGO_TEL") 
		m_HOS_SIMEI		= m_Rs("T11_HOS_SIMEI") 
		m_HOS_SIMEI_K	= m_Rs("T11_HOS_SIMEI_K") 
		m_HOS_ZIP		= m_Rs("T11_HOS_ZIP")
		m_HOS_JUSYO		= m_Rs("T11_HOS_JUSYO")
	
	'/* 住所に県、市町村名が存在していた場合削除して再度付け直す。*/ Add 2002.04.30 okada
		m_HOS_JUSYO		= m_Rs("T11_HOS_JUSYO")
		m_HOS_JUSYO		= Replace(m_HOS_JUSYO,m_Ken2,"")
		m_HOS_JUSYO		= Replace(m_HOS_JUSYO,m_SITYOSONMEI2,"")

		m_HOS_JUSYO		= m_Ken2 & m_SITYOSONMEI2 & m_HOS_JUSYO'm_SITYOSONMEI2 & Replace(m_Rs("T11_HOS_JUSYO"),m_SITYOSONMEI2,"")

		'm_HOS_TEL		= m_Rs("T11_HOS_TEL")
		'_KAZOKU_ZOKU_1 = m_Rs("T11_KAZOKU_ZOKU_1")
		'_KAZOKU_ZOKU_2 = m_Rs("T11_KAZOKU_ZOKU_2")
		'_KAZOKU_ZOKU_3 = m_Rs("T11_KAZOKU_ZOKU_3")
		'_KAZOKU_ZOKU_4 = m_Rs("T11_KAZOKU_ZOKU_4")
		'_KAZOKU_ZOKU_5 = m_Rs("T11_KAZOKU_ZOKU_5")
		'_KAZOKU_ZOKU_6 = m_Rs("T11_KAZOKU_ZOKU_6")
		'_KAZOKU_ZOKU_7 = m_Rs("T11_KAZOKU_ZOKU_7")
		'_KAZOKU_ZOKU_8 = m_Rs("T11_KAZOKU_ZOKU_8")
		'm_RYUGAKU_KBN  = m_Rs("T11_RYUGAKU_KBN")
		m_KAZOKU_1      = m_Rs("T11_KAZOKU_1")
		m_KAZOKU_2      = m_Rs("T11_KAZOKU_2")
		m_KAZOKU_3      = m_Rs("T11_KAZOKU_3")
		m_KAZOKU_4      = m_Rs("T11_KAZOKU_4")
		m_KAZOKU_5      = m_Rs("T11_KAZOKU_5")
		m_KAZOKU_6      = m_Rs("T11_KAZOKU_6")
		m_KAZOKU_7      = m_Rs("T11_KAZOKU_7")
		m_KAZOKU_8      = m_Rs("T11_KAZOKU_8")
		m_SYUSSINKO     = f_GetSyussinko(Session("HyoujiNendo"),m_Rs("T11_SYUSSINKO"))
		m_SYUSSINKOKU   = m_Rs("T11_SYUSSINKOKU")

		m_HOG_SEINEIBI      =   m_Rs("T11_HOG_SEINEIBI")
		m_HOS_SEINEIBI      =   m_Rs("T11_HOS_SEINEIBI")
		m_KAZOKU_SEINEIBI_1 =   m_Rs("T11_KAZOKU_SEINEIBI_1")
		m_KAZOKU_SEINEIBI_2 =   m_Rs("T11_KAZOKU_SEINEIBI_2")
		m_KAZOKU_SEINEIBI_3 =   m_Rs("T11_KAZOKU_SEINEIBI_3")
		m_KAZOKU_SEINEIBI_4 =   m_Rs("T11_KAZOKU_SEINEIBI_4")
		m_KAZOKU_SEINEIBI_5 =   m_Rs("T11_KAZOKU_SEINEIBI_5")
		m_KAZOKU_SEINEIBI_6 =   m_Rs("T11_KAZOKU_SEINEIBI_6")
		m_KAZOKU_SEINEIBI_7 =   m_Rs("T11_KAZOKU_SEINEIBI_7")
		m_KAZOKU_SEINEIBI_8 =   m_Rs("T11_KAZOKU_SEINEIBI_8")

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
			<td nowrap><a href="kojin_sita4.asp">●その他予備情報</a></td>
			<td nowrap><a href="kojin_sita5.asp">●異動情報</a></td>
		</tr>
	</table>
	<br>

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td valign="top" align="left">

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
						<% if Cint(gf_SetNull2Zero(m_Rs("T11_NYUGAKU_KBN"))) = C_NYU_RYUGAKU then %>
							<% if gf_empItem(C_T11_SYUSSINKO) then %>
								<tr>
									<td height="16" class="disph">出 身 校</td>
									<td class="disp"><%= m_SYUSSINKO %>&nbsp</td>
								</tr>
							<% End if %>
							<% if gf_empItem(C_T11_SYUSSINKOKU) then %>
								<tr>
									<td height="16" class="disph">出 身 国</td>
									<td class="disp"><%= m_SYUSSINKOKU %>&nbsp</td>
								</tr>
							<% End if %>
							<% if gf_empItem(C_T11_RYUGAKU_KBN) then %>
								<tr>
									<td height="16" class="disph">留 学 区 分</td>
									<td class="disp"><%= m_RYUGAKU_KBN %>&nbsp</td>
								</tr>
							<% End if %>
						<% End if %>
					</table>
			</td>

			<td valign="top" align="left">

					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_HOG_SIMEI) then %>
							<tr>
								<td class="disph" width="110" height="16"><font color="white">保護者氏名</font></td>
								<td class="disp"  width="110"><%= m_HOG_SIMEI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_SIMEI_K) then %>
							<tr>
								<td class="disph" height="16"><font color="white">保護者カナ</font></td>
								<td class="disp"><%= m_HOG_SIMEI_K %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_ZOKU) then %>
							<tr>
								<td class="disph" height="16"><font color="white">保護者続柄</font></td>
								<td class="disp"><%= m_HOG_ZOKU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_SEINEIBI) then %>
							<tr>
								<td class="disph" height="16"><font color="white">保護者生年月日</font></td>
								<td class="disp"><%= m_HOG_SEINEIBI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_ZIP) then %>
							<tr>
								<td class="disph" height="16"><font color="white">保護者〒</font></td>
								<td class="disp"><%= m_HOG_ZIP %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_JUSYO) then %>
							<tr>
								<td class="disph" height="16"><font color="white">保護者住所</font></td>
								<td class="disp"><%= m_HOG_JUSYO %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_TEL) then %>
							<tr>
								<td class="disph" height="16"><font color="white">保護者TEL</font></td>
								<td class="disp"><%= m_HOGO_TEL %>&nbsp</td>
							</tr>
						<% End if %>
					</table>
			<br>

			</td>
			<td valign="top" align="left">
					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_HOS_SIMEI) then %>
							<tr>
								<td class="disph" width="110" height="16">保証人氏名</td>
								<td class="disp"  width="110"><%= m_HOS_SIMEI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_SIMEI_K) then %>
							<tr>
								<td class="disph" height="16">保証人カナ</td>
								<td class="disp"><%= m_HOS_SIMEI_K %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_ZOKU) then %>
							<tr>
								<td class="disph" height="16">保証人続柄</td>
								<td class="disp"><%= m_HOS_ZOKU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_SEINEIBI) then %>
							<tr>
								<td class="disph" height="16">保証人生年月日</td>
								<td class="disp"><%= m_HOS_SEINEIBI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_ZIP) then %>
							<tr>
								<td class="disph" height="16">保証人〒</td>
								<td class="disp"><%= m_HOS_ZIP %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_JUSYO) then %>
							<tr>
								<td class="disph" height="16">保証人住所</td>
								<td class="disp"><%= m_HOS_JUSYO %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_TEL) then %>
							<tr>
								<td class="disph" height="16">保証人TEL</td>
								<td class="disp"><%= m_HOS_TEL %>&nbsp</td>
							</tr>
						<% End if %>
					</table>
					
			</td>

		</tr>
	<tr><td colspan=3>

					<table border="1" class="disp">
						<% if gf_empItem(C_T11_KAZOKU_1) then %>
							<tr>
								<td class="disph" width="100" height="16">家族名称１</td>
								<td class="disp" width="120" ><%= m_KAZOKU_1 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_1) then %>
								<td class="disph" width="60" height="16"> 続 柄 １</td>
								<td class="disp" width="60"><%= m_KAZOKU_ZOKU_1 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_1) then %>
								<td class="disph" width="100" height="16">生年月日１</td>
								<td class="disp" width="100"><%= m_KAZOKU_SEINEIBI_1 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_2) then %>
							<tr>
								<td class="disph" height="16">家族名称２</td>
								<td class="disp"><%= m_KAZOKU_2 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_2) then %>
								<td class="disph" height="16"> 続 柄 ２</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_2 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_2) then %>
								<td class="disph" height="16">生年月日２</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_2 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_3) then %>
							<tr>
								<td class="disph" height="16">家族名称３</td>
								<td class="disp"><%= m_KAZOKU_3 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_3) then %>
								<td class="disph" height="16"> 続 柄 ３</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_3 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_3) then %>
								<td class="disph" height="16">生年月日３</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_3 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_4) then %>
							<tr>
								<td class="disph" height="16">家族名称４</td>
								<td class="disp"><%= m_KAZOKU_4 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_4) then %>
								<td class="disph" height="16"> 続 柄 ４</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_4 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_4) then %>
								<td class="disph" height="16">生年月日４</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_4 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_5) then %>
							<tr>
								<td class="disph" height="16">家族名称５</td>
								<td class="disp"><%= m_KAZOKU_5 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_5) then %>
								<td class="disph" height="16"> 続 柄 ５</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_5 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_5) then %>
								<td class="disph" height="16">生年月日５</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_5 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_6) then %>
							<tr>
								<td class="disph" height="16">家族名称６</td>
								<td class="disp"><%= m_KAZOKU_6 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_6) then %>
								<td class="disph" height="16"> 続 柄 ６</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_6 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_6) then %>
								<td class="disph" height="16">生年月日６</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_6 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_7) then %>
							<tr>
								<td class="disph" height="16">家族名称７</td>
								<td class="disp"><%= m_KAZOKU_7 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_7) then %>
								<td class="disph" height="16"> 続 柄 ７</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_7 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_7) then %>
								<td class="disph" height="16">生年月日７</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_7 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_8) then %>
							<tr>
								<td class="disph" height="16">家族名称８</td>
								<td class="disp"><%= m_KAZOKU_8 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_8) then %>
								<td class="disph" height="16"> 続 柄 ８</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_8 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_8) then %>
								<td class="disph" height="16">生年月日８</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_8 %>&nbsp</td>
							</tr>
						<% End if %>
					</table>

	</td></tr>
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