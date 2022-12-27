<%
	On Error Resume Next
	Err.Clear
	Dim w_cCn,w_rRs,SQL

' オブジェクトの定義
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_社員",w_cCn,2,2
	
' 修正SQL
	SQL ="UPDATE M_社員 SET 社員名称='" & Request.Form("社員名称") & "'"
	SQL = SQL & ",生年月日=" & Request.Form("生年月日")
	SQL = SQL & ",電話番号1=" & Request.Form("電話番号1")
	SQL = SQL & ",電話番号2=" & Request.Form("電話番号2")
	SQL = SQL & ",郵便=" & Request.Form("郵便")
	SQL = SQL & ",住所1=" & Request.Form("住所1")
	SQL = SQL & ",住所2=" & Request.Form("住所2")
	SQL = SQL & ",備考=" & Request.Form("備考")
	SQL = SQL & ",使用FLG=1 WHERE 社員CD=" & Request.Form("社員CD")
	
	Set w_rRs = w_cCn.Execute(SQL)
	
' SQL実行時のエラー処理
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	
	On Error Goto 0

	'w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
	
' 終了メッセージ
	Response.Redirect "FinishSYUUSEI.asp"
	
%>
