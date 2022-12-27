<%
	On Error Resume Next
	Err.Clear
	Dim w_cCn,w_rRs,SQL

' オブジェクト定義
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_社員",w_cCn,2,2
    
' 社員データ削除のSQL文
    SQL = "UPDATE M_社員 SET 使用FLG=0 WHERE 社員CD=" & Request.Form("社員CD")
    
    Set w_rRs = w_cCn.Execute(SQL)
    
' SQL実行時のエラー処理
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	
	On Error Goto 0
	
' 成功メッセージ
	Response.Redirect "FinishDel.asp"
	
    w_rRs.Close
	Set w_rRs = Nothing
	w_cCn.Close
	Set w_cCn = Nothing

%>
