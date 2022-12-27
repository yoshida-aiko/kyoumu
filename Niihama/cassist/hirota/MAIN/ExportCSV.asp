<%	
	On Error Resume Next
    Err.Clear

	Dim w_cCn,w_rRs,w_SQL,w_Index
	Dim w_StartCD,w_EndCD,w_Name,w_CheckDel

	w_StartCD = Request.Form("txtStartCD")
	w_EndCD = Request.Form("txtEndCD")
	w_Name = Request.Form("txtName")
	w_CheckDel = Request.Form("checkDel")
	w_SQL = Request.Form("SQL")
	
'--------------------全角を半角に変換----------------------------------
	Set bobj = Server.CreateObject("basp21")
	w_StartCD = bobj.StrConv(w_StartCD,8)	'全角→半角変換
	
' オブジェクトの定義   
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")
	
    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_社員",w_cCn,2,2
    
    Set w_rRs = w_cCn.Execute(w_SQL)
    
' SQL実行時のエラー処理
	if Err then
		Session.Contents("SQLerror")=Err.description
		Response.Redirect "SQLerror.asp"
	end if

	On Error Goto 0
    
' 該当する社員がいるかどうかの判定
	if w_rRs.EOF=true then
		Response.Redirect "NOexport.asp"
	end if
'************************************************************************************************
'											CSV出力処理
'************************************************************************************************
	FileName="\\WEBSVR_2\infogram\hirota\No02\Sample.csv"
   'FileName=Server.MapPath("Sample.csv") C:\infogram\hirota\No02\Sample.csv 
	Set g_file = Server.CreateObject("Scripting.fileSystemObject")	'ファイルシステムオブジェクトの作成

	
	Set f_ExportFile = g_file.CreateTextFile(FileName, True)	'ファイルの作成
	w_Index = 0
	
' CSVファイルに書込み
	Do while not w_rRs.EOF
	w_Index = w_Index + 1
		f_ExportFile.WriteLine(w_rRs("社員CD") & "," & w_rRs("社員名称") & "," & w_rRs("生年月日") _
				& "," & w_rRs("電話番号1") & "," & w_rRs("電話番号2")	 & "," & w_rRs("郵便") & "," _
							& w_rRs("住所1") & "," & w_rRs("住所2") & "," & w_rRs("備考"))	'１行ライト
		w_rRs.MoveNext
	Loop
	f_ExportFile.Close						'ファイルのクローズ

' 出力パスとレコードカウントの送信
	Session.Contents("Path")=FileName
	Response.Redirect "FinishCSV.asp?Count=" & w_Index

	
    w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
%>