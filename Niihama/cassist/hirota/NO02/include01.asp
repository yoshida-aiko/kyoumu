<%
	On Error Resume Next
	Err.Clear
	Dim w_sCD,w_sName,w_sYEAR,w_sMONTH,w_sDAY,w_sTel1,w_sTel2,w_sTel3,w_sTel4,w_sTel5,w_sTel6
	Dim w_sPost1,w_sPost2,w_sAddress1,w_sAddress2,w_sBikou

' テキストに入力されたデータを変数に格納する。
	w_sCD = Request.Form("w_sCD")
	w_sName = Request.Form("w_sName")
	w_sYEAR = Request.Form("w_sYEAR")
	w_sMONTH = Request.Form("w_sMONTH")
	w_sDAY = Request.Form("w_sDAY")
	w_sTel1 = Request.Form("w_sTel1")
	w_sTel2 = Request.Form("w_sTel2")
	w_sTel3 = Request.Form("w_sTel3")
	w_sTel4 = Request.Form("w_sTel4")
	w_sTel5 = Request.Form("w_sTel5")
	w_sTel6 = Request.Form("w_sTel6")
	w_sPost1 = Request.Form("w_sPost1")
	w_sPost2 = Request.Form("w_sPost2")
	w_sAddress1 = Request.Form("w_sAddress1")
	w_sAddress2 = Request.Form("w_sAddress2")
	w_sBikou = Request.Form("w_sBikou")	

Function SHINKI()
	'**************************************************************
	'			新　規
	'**************************************************************
	
		Dim g_cCn,g_rRs,SQL
		
		g_sFLG="1"
		Response.Write "<h3 align=center>★ 新規確認画面 ★</h3>"
		
	' 社員CD、社員名称の入力チェック
		if w_sCD = "" or w_sName = "" then
			w_FLG = "4" '(入力エラーフラグ : 4 )
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
		
	' オブジェクト定義
		Set g_cCn = Server.CreateObject("ADODB.Connection")
		Set g_rRs = Server.CreateObject("ADODB.Recordset")
		g_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
		                    & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
		g_rRs.Open "M_社員",g_cCn,2,2
		
	' 社員CD重複チェック
		SQL="SELECT 社員CD FROM M_社員 WHERE 使用FLG=1 AND 社員CD=" & w_sCD
		Set g_rRs = g_cCn.Execute(SQL)
		
	' SQL実行時のエラー処理
		if Err then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if	
		On Error Goto 0
		
	' 重複チェック
		if g_rRs.EOF=false then
			w_FLG="2" '(重複メッセージフラグ : 2 )
			Session.Contents("w_sCD")=w_sCD
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
End Function
	'**************************************************************
	'			修　正
	'**************************************************************
Function SYUUSEI()
		g_sFLG="2"
		w_sCD=Request.Form("CD")
		Response.Write "<h3 align=center>★ 修正確認画面 ★</h3>"
End Function

%>