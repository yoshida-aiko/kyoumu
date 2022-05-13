<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名:
' ﾌﾟﾛｸﾞﾗﾑID :
' 機      能:
'-------------------------------------------------------------------------
' 引      数:
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2003/02/24 hirota
'*************************************************************************/

%>
<!--#include file="../../Common/com_All.asp"-->
<%

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

	Public gRs					'//レコード
	Public msURL
	Public m_bErrFlg
	Public m_sLoad

	Public m_iGakunen			'//学年
	Public m_iClassNo			'//クラスNO
	Public m_iSyoriNen          '//年度
	Public m_iKyokanCd          '//教官ｺｰﾄﾞ
	Public m_iGakka				'//学科
	Public m_sHyoka()			'//評価記号配列
	Public m_sClass				'//クラス
	Public m_sClassNM			'//クラス名

'///////////////////////////メイン処理/////////////////////////////

	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()

	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    On Error Resume Next
    Err.Clear

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="不合格学生一覧"
    w_sMsg=""
    w_sRetURL = C_RetURL & C_ERR_RETURL
    w_sTarget = "fTopMain"

    m_bErrFlg = False

    Do
		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'// パラメータ取得
		Call s_GetParameter()

		'// 表示ボタン押下時
		If m_sLoad = "load" then

			'// ﾃﾞｰﾀﾍﾞｰｽ接続
			If gf_OpenDatabase() <> 0 Then
				'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
				m_sErrMsg = "データベースとの接続に失敗しました。"
				Exit Do
			End If

			'// 評価記号取得
			If Not f_GetHyoka() then
				m_sErrMsg = "評価形式取得に失敗しました。"
				Exit Do
			End If

			'// クラスデータ取得
			If Not f_GetClassData() then
				m_sErrMsg = "クラスデータ取得に失敗しました。"
				Exit Do
			End If

		End If

		'// ページを表示
		Call showPage()

		m_bErrFlg = True
        Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If Not m_bErrFlg Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
    End If

	'// 終了処理
    Call gf_closeObject(gRs)
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	パラメータ取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_GetParameter()

	m_sLoad     = Request("mode")
	m_iSyoriNen = Session("NENDO")
	m_sClass    = Request("hidClass")
	m_iGakunen  = Request("hidGakunen")
	m_sClassNM  = Request("hidClassNM")

End Sub

'********************************************************************************
'*	[機能]	評価記号取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_GetHyoka()

	Dim w_sSQL
	Dim w_lRecCnt
	Dim w_iRet
	Dim wRs

	On Error Resume Next
	Err.Clear

	f_GetHyoka = False

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD, "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI_R "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & " 	AND M01_DAIBUNRUI_CD = " & C_HYOKA_FUKA
	w_sSQL = w_sSQL & " ORDER BY "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD "

	w_iRet = gf_GetRecordset(wRs, w_sSQL)

	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit Function
	End If

	If Not wRs.EOF then
		w_lRecCnt = wRs.RecordCount											'//レコードカウント
		wRs.MoveFirst

		'//評価記号配列セット
		Do While Not wRs.EOF
			Redim Preserve m_sHyoka(wRs("M01_SYOBUNRUI_CD"))							'//評価記号配列定義
			m_sHyoka(wRs("M01_SYOBUNRUI_CD")) = wRs("M01_SYOBUNRUIMEI_R")	'//評価記号
			wRs.MoveNext
		Loop
	End If

	'//レコード解放
    Call gf_closeObject(wRs)

	f_GetHyoka = True

End Function

'********************************************************************************
'*	[機能]	クラスデータ取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_GetClassData()

	Dim w_sSQL
	Dim w_iRet

	On Error Resume Next
	Err.Clear

	f_GetClassData = False

	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T11.T11_SIMEI, "
	w_sSQL = w_sSQL & " 	T11.T11_GAKUSEI_NO, "
	w_sSQL = w_sSQL & " 	T13.T13_SYUSEKI_NO1, "
	w_sSQL = w_sSQL & " 	T13.T13_GAKUNEN , "
	w_sSQL = w_sSQL & " 	T13.T13_CLASS , "
	w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO , "
	w_sSQL = w_sSQL & " 	T120.*, "
	w_sSQL = w_sSQL & " 	M04.M04_KYOKANMEI_SEI, "
	w_sSQL = w_sSQL & " 	M04.M04_KYOKANMEI_MEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI T11, "
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN T13, "
	w_sSQL = w_sSQL & " 	T120_SAISIKEN T120, "
	w_sSQL = w_sSQL & " 	M04_KYOKAN M04 "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	T13.T13_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & " 	AND T13.T13_CLASS        = '" & m_sClass & "'"
	w_sSQL = w_sSQL & " 	AND T13.T13_GAKUNEN      =  " & m_iGakunen
	w_sSQL = w_sSQL & " 	AND T13.T13_GAKUSEI_NO   = T120.T120_GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND T11.T11_GAKUSEI_NO   = T120.T120_GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND T120.T120_NENDO      = M04.M04_NENDO "
	w_sSQL = w_sSQL & " 	AND T120.T120_KYOUKAN_CD = M04.M04_KYOKAN_CD(+) "
	w_sSQL = w_sSQL & " ORDER BY "
	w_sSQL = w_sSQL & " 	T13.T13_GAKUNEN, "
	w_sSQL = w_sSQL & " 	T13.T13_CLASS, "
	w_sSQL = w_sSQL & " 	T13.T13_SYUSEKI_NO1, "
	w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO, "
	w_sSQL = w_sSQL & " 	T120.T120_NENDO DESC, "
	w_sSQL = w_sSQL & " 	T120.T120_KAMOKU_CD "

	w_iRet = gf_GetRecordset(gRs, w_sSQL)

	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit Function
	End If

	f_GetClassData = True

End Function

'********************************************************************************
'*	[機能]	HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear
	'---------- HTML START ----------
%>
<html>
<head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>不合格学生一覧</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
    //************************************************************
    //  [機能]  フォームロード時
    //  [引数]  
    //  [戻値]  
    //  [説明]
    //************************************************************
	function jf_windowload(){
		<% If m_sLoad = "load" then %>
			<% If gRs.EOF then %>
				alert("対象データは存在しません。");
				parent._TOP.document.body.style.cursor = "default";
				return;
			<% End If %>
			with(document.frm){
				target = "_TOP";
				action = "saisi0500_head.asp";
				submit();
			}
		<% End If %>
	}
	window.onload = jf_windowload;
	//-->
	</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<center>

<form name="frm" action="" target="main" Method="POST">

<%
If m_sLoad = "load" then

	Dim w_Class
	Dim w_sSeiseki
	Dim w_sSaisi

	w_Class = ""

	Do While Not gRs.EOF

		'テーブル背景色設定
		Call gs_cellPtn(w_Class)

		'評価記号 + 旧評価
		w_sSeiseki = m_sHyoka(gRs("T120_OLD_HYOKA_FUKA_KBN")) & gRs("T120_SEISEKI")
		'w_sSaisi   = m_sHyoka(gRs("T120_HYOKA_FUKA_KBN")) & gRs("T120_SAISI_SEISEKI")
		w_sSaisi   = gRs("T120_SAISI_SEISEKI")
%>
	<table class="hyo" border="1">
		<tr>
			<td width="60"  class="<%= w_Class %>" nowrap align="right"><%= gRs("T13_SYUSEKI_NO1") %></td>
			<td width="150" class="<%= w_Class %>" nowrap><%= gRs("T11_SIMEI") %></td>
			<td width="150" class="<%= w_Class %>" nowrap><%= gRs("T120_KAMOKUMEI") %></td>
			<td width="50"  class="<%= w_Class %>" nowrap align="center"><%= gRs("T120_NENDO") %></td>
			<td width="70"  class="<%= w_Class %>" nowrap align="right"><%= gRs("T120_KEKASU") & " / " & gRs("T120_JUNJIKAN") %></td>
			<td width="40"  class="<%= w_Class %>" nowrap align="right"><%= w_sSeiseki %></td>
			<td width="40"  class="<%= w_Class %>" nowrap align="right"><%= w_sSaisi %></td>
			<td width="100" class="<%= w_Class %>" nowrap><%= gRs("M04_KYOKANMEI_SEI") & " " & gRs("M04_KYOKANMEI_MEI") %></td>
		</tr>
	</table>
<%
		gRs.MoveNext
	Loop
Else
%>

	<br><br><br>
	<CENTER><span class="msg">※　表示ボタンを押してください </span></CENTER>

<% End If %>
<input type="hidden" name="hidClass" value="<%= m_sClass %>">
<input type="hidden" name="hidGakunen" value="<%= m_iGakunen %>">
<input type="hidden" name="hidClassNM" value="<%= m_sClassNM %>">
</form>

</center>

</body>
</html>
<%
'---------- HTML END   ----------
End Sub
%>