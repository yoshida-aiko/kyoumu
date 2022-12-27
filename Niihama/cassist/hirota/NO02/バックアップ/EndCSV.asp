<%
	On Error Resume Next
    Err.Clear
    
' オブジェクト定義   
	Set g_cCn = Server.CreateObject("ADODB.Connection")
	Set g_rRs = Server.CreateObject("ADODB.Recordset")
	Set g_file = Server.CreateObject("Scripting.fileSystemObject")	'ファイルシステムオブジェクトの作成

    g_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    g_rRs.Open "M_社員",g_cCn,2,2

	Set g_rRs = g_cCn.Execute(Session.Contents("g_CSV"))
	
' SQL実行時のエラー処理
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	
	On Error Goto 0
	
'************************************************************************************************
'											CSV出力処理
'************************************************************************************************
	FileName="\\WEBSVR_2\infogram\hirota\No02\Sample.csv"
	'FileName=Server.MapPath("Sample.csv") C:\infogram\hirota\No02\Sample.csv 
	
	Set fs_test = g_file.CreateTextFile(FileName, True)	'ファイルの作成
	w_Index = 0
	
' CSVファイルに書込み
	Do while not g_rRs.EOF
	w_Index = w_Index + 1
		fs_test.WriteLine(g_rRs("社員CD") & "," & g_rRs("社員名称") & "," & g_rRs("生年月日") _
				& "," & g_rRs("電話番号1") & "," & g_rRs("電話番号2")	 & "," & g_rRs("郵便") & "," _
							& g_rRs("住所1") & "," & g_rRs("住所2") & "," & g_rRs("備考"))	'１行ライト
		g_rRs.MoveNext
	Loop
	fs_test.Close						'ファイルのクローズ

' 出力パスとレコードカウントの送信
	Session.Contents("Path")=FileName
	Response.Redirect "FinishCSV.asp?Count=" & w_Index

' オブジェクトの開放
    w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
	
%>