
<%
'----------------------------------------
'ファイルのデータを表示
'----------------------------------------

'--- オブジェクト作成 ---
Set w_oFile = Server.CreateObject("Scripting.FileSystemObject")

'--- ファイルを開く（読み取り専用、ファイルが存在しないときは新規作成） ---
Set w_oCSV = w_oFile.OpenTextFile(Server.MapPath("Sample.csv"),1,True)

'--- ファイルのデータを表示 ---
Do Until w_oCSV.AtEndofStream
    Response.Write "<p>" & w_oCSV.ReadLine
Loop

'--- ファイルを閉じる ---
w_oCSV.Close

'--- オブジェクト解放 ---
Set w_oCSV = Nothing
Set w_oFile = Nothing

%>
