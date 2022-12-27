<%
	On Error Resume Next
	Err.Clear
	Dim w_cCn,w_rRs
	Dim w_sSQL

' オブジェクト定義
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

 	w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
 	w_rRs.Open "M_社員",w_cCn,2,2
 	
 e_SELECT = Request.Form("FLG")
 Select case e_SELECT
	'**************************************************************
	'			新規
	'**************************************************************
	 case "1"
	 
	 ' 使用FLG=0の社員データが残っている場合
	 	w_sSQL = "SELECT 社員CD FROM M_社員 WHERE 使用FLG=0 AND 社員CD=" & Request.Form("社員CD")
	 	Set w_rRs = w_cCn.Execute(w_sSQL)
	 	
	 ' SQL実行時のエラー処理
	 	if Err then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if

		On Error Goto 0
		
	' 社員CDの重複チェック(使用FLG=0の場合)
	 	if w_rRs.EOF=false then
			Session.Contents("社員名称")=Request.Form("社員名称")
			Session.Contents("生年月日")=Request.Form("生年月日")
			Session.Contents("電話番号1")=Request.Form("電話番号1")
			Session.Contents("電話番号2")=Request.Form("電話番号2")
			Session.Contents("郵便")=Request.Form("郵便")
			Session.Contents("住所1")=Request.Form("住所1")
			Session.Contents("住所2")=Request.Form("住所2")
			Session.Contents("備考")=Request.Form("備考")
			Session.Contents("社員CD")=Request.Form("社員CD")
			w_FLG="3"
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
		
	' 新規登録のSQL文の作成	
		w_sSQL = "INSERT INTO M_社員 (社員CD,社員名称,生年月日,電話番号1,電話番号2,"
	    w_sSQL = w_sSQL & "郵便,住所1,住所2,備考,使用FLG)"
	    w_sSQL = w_sSQL & " VALUES (" & Request.Form("社員CD") & ",'" & Request.Form("社員名称") & "'"
	    w_sSQL = w_sSQL & "," & Request.Form("生年月日")
	    w_sSQL = w_sSQL & "," & Request.Form("電話番号1")
	    w_sSQL = w_sSQL & "," & Request.Form("電話番号2")
	    w_sSQL = w_sSQL & "," & Request.Form("郵便")
	    w_sSQL = w_sSQL & "," & Request.Form("住所1")
	    w_sSQL = w_sSQL & "," & Request.Form("住所2")
	    w_sSQL = w_sSQL & "," & Request.Form("備考") & ",1)"

		if gf_SQLexe(w_sSQL)=false then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if
		w_sFLG = "1"
		
	'**************************************************************
	'			修正
	'**************************************************************
	case "2"
		
	' 修正SQL
		w_sSQL ="UPDATE M_社員 SET 社員名称='" & Request.Form("社員名称") & "'"
		w_sSQL = w_sSQL & ",生年月日=" & Request.Form("生年月日")
		w_sSQL = w_sSQL & ",電話番号1=" & Request.Form("電話番号1")
		w_sSQL = w_sSQL & ",電話番号2=" & Request.Form("電話番号2")
		w_sSQL = w_sSQL & ",郵便=" & Request.Form("郵便")
		w_sSQL = w_sSQL & ",住所1=" & Request.Form("住所1")
		w_sSQL = w_sSQL & ",住所2=" & Request.Form("住所2")
		w_sSQL = w_sSQL & ",備考=" & Request.Form("備考")
		w_sSQL = w_sSQL & ",使用FLG=1 WHERE 社員CD=" & Request.Form("社員CD")
		
		if gf_SQLexe(w_sSQL)=false then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if
		w_sFLG="2"	

	'**************************************************************
	'			削除
	'**************************************************************
	case "3"
	    
	' 社員データ削除のSQL文
	    w_SQL = "UPDATE M_社員 SET 使用FLG=0 WHERE 社員CD=" & Request.Form("社員CD")

		if gf_SQLexe(w_SQL)=false then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if
		w_sFLG="3"
	
end Select

Response.Redirect "FinishMsg.asp?FLG=" & w_sFLG

w_rRs.Close
w_cCn.Close
Set w_rRs = Nothing
Set w_cCn = Nothing

'**************************************************************
'			SQL実行関数
'**************************************************************
function gf_SQLexe(p_sSQL)
	Set w_rRs = w_cCn.Execute(p_sSQL)
 ' SQL実行時のエラー処理
	if Err then
		gf_SQLexe=false
	end if
	On Error Goto 0
	gf_SQLexe=true
end function

%>