<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績参照（教官側）
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0800/default.asp
' 機      能: 
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2003/05/13 廣田
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

	Public  m_iNendo   			'年度
	Public  m_sKyokanCd			'ログイン教官
	Public  m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ
	Public  m_Rs
	Public  m_RecCnt			'レコードカウント

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

	Dim w_sWinTitle
	Dim w_sMsgTitle
	Dim w_sMsg
	Dim w_sRetURL
	Dim w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="成績参照"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_parent"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If

		'// 権限チェックに使用
		Session("PRJ_No") = "SEI0800"

		'// 不正アクセスチェック
		Call gf_userChk(Session("PRJ_No"))

		'//ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()

		'// 学生一覧取得アクセスチェック
		If Not f_GetStudent() Then m_bErrFlg = True : Exit Do

		'// 該当者がいない場合
		If m_Rs.EOF Then
			Call gs_showWhitePage("個人履修データが存在しません。","成績参照")
			Exit Do
		End If

		'// ページを表示
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
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()

    m_iNendo    = Session("NENDO")
    m_sKyokanCd = Session("KYOKAN_CD")

End Sub

Function f_GetStudent()
'********************************************************************************
'*  [機能]  ログイン教官が担当するクラスの学生一覧を取得する
'*  [引数]  なし
'*  [戻値]  True / False
'*  [説明]  
'********************************************************************************

	On Error Resume Next
	Err.Clear

	Dim w_sSQL

	f_GetStudent = False

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T13_GAKUSEI_NO,  "
	w_sSQL = w_sSQL & " 	T13_GAKUSEKI_NO, "
	w_sSQL = w_sSQL & " 	T13_GAKUNEN, "
	w_sSQL = w_sSQL & " 	T11_SIMEI,   "
	w_sSQL = w_sSQL & " 	M05_CLASSMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI, "
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN, "
	w_sSQL = w_sSQL & " 	M05_CLASS "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M05_NENDO      =  " & m_iNendo & " AND "
	w_sSQL = w_sSQL & " 	M05_TANNIN     = '" & m_sKyokanCd & "' AND"
	w_sSQL = w_sSQL & " 	T13_NENDO      =  M05_NENDO      AND "
	w_sSQL = w_sSQL & " 	T13_GAKUNEN    =  M05_GAKUNEN    AND "
	w_sSQL = w_sSQL & " 	T13_GAKKA_CD   =  M05_GAKKA_CD   AND "
	w_sSQL = w_sSQL & " 	T13_GAKUSEI_NO =  T11_GAKUSEI_NO     "

	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit Function

	'//ﾚｺｰﾄﾞカウント取得
	m_RecCnt = gf_GetRsCount(m_Rs)

	f_GetStudent = True

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	On Error Resume Next
	Err.Clear

	Dim w_sCell

	w_sCell = "CELL1"

%>
<html>

<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>

<body LANGUAGE="javascript">
	<center>
	<form name="frm" METHOD="post">
	<% call gs_title(" 成績参照 "," 参　照 ") %>

	<table  border="1" class="hyo">
		<tr>
			<th class="header2" width="550" align="center"><%=m_Rs("T13_GAKUNEN")%>　年　<%=m_Rs("M05_CLASSMEI")%></th>
		</tr>
	<table>

	<br>

	<!-- TABLEヘッダ部 -->
	<table border="1" class="hyo">
		<tr>
			<th width="100" class="header3" align="center" height="20">学生番号</th>
			<th width="100" class="header3" align="center" height="20">出席番号</th>
			<th width="250" class="header3" align="center" height="20">氏　　名</th>
			<th width="100" class="header3" align="center" height="20">成績表示</th>
		</tr>
	</table>

	<!-- TABLEリスト部 -->
	<table class="hyo" align="center" border="1">

<%
	Do While Not m_Rs.EOF
		w_sCell = gf_IIF(w_sCell="CELL1","CELL2","CELL1")
%>
						<tr>
							<td width="100" class="<%=w_sCell%>" align="center" nowrap><%=m_Rs("T13_GAKUSEI_NO")%></td>
							<td width="100" class="<%=w_sCell%>" align="center" nowrap><%=m_Rs("T13_GAKUSEKI_NO")%></td>
							<td width="250" class="<%=w_sCell%>" align="left"   nowrap>　<%=m_Rs("T11_SIMEI")%></td>
							<td width="100" class="<%=w_sCell%>" align="center" nowrap><input type="button" name="btnDisp" value="表　示"></td>
						</tr>

<%
		m_Rs.MoveNext
	Loop
%>

	</table>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
