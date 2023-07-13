<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0150_23/default.asp
' 機      能: 
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2003/05/01 廣田　耕一郎
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

	Dim m_sHidVal

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

	m_sHidVal = ""

	For Each I_Name In Request.Form
		m_sHidVal = m_sHidVal & "&" & I_Name & "=" & Request.Form(I_Name)
	Next

	Call showPage()

End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>

<html>

<head>
<title>成績登録</title>
</head>

<frameset rows=*,0 frameborder="0" framespacing="0">
	<frame src="sei0150_23_dispprint.asp" scrolling="yes"  name="Disp" noresize>
	<frame src="sei0150_23_hidprint.asp?a=a<%=m_sHidVal%>" name="Hidden" noresize>
</frameset>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
