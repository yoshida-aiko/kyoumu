<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 連絡掲示板
' ﾌﾟﾛｸﾞﾗﾑID : web/web0330/web0330_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/10 前田
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_sNendo             '年度
    Public m_PgMode             '処理別フラグ
    Public m_sMsgTitle          'ﾀｲﾄﾙ
	
	Public m_Rs
	
	'エラー系
	Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
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
	
	On Error Resume Next
	Err.Clear
	
	m_PgMode=request("p_mode")
	
	Select Case m_PgMode
		Case "P_HAN0100"
		    m_sMsgTitle="成績一覧表"
		Case "P_KKS0200"
		    m_sMsgTitle="欠課一覧表"
		Case "P_KKS0210"
		    m_sMsgTitle="遅刻一覧表"
		Case "P_KKS0220"
		    m_sMsgTitle="行事欠課一覧表"
		Case "P_HAN0111"
		    m_sMsgTitle="評点一覧表"
		Case Else
	End Select
	
	m_bErrFlg = False
	
	m_sNendo    = session("NENDO")
	
	Do
		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'// ページを表示
		Call showPage()
		Exit Do
	Loop
	
End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
%>
	<html>
	<head>
		<link rel="stylesheet" href=../../common/style.css type="text/css">
		<title><%=m_sMsgTitle%></title>
	</head>
	
	<body>
	<form>
		<center><% call gs_title(m_sMsgTitle,"一　覧") %><br></center>
		<INPUT TYPE=HIDDEN NAME=txtNendo    VALUE="<%=m_sNendo%>">
	</form>
	</body>
	</html>
<%
End Sub
%>
