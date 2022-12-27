
<!--#INCLUDE FILE="include02.asp"-->

<%

' SQL文をセッション変数に格納 （GO → EndCSV.asp）
    Session.Contents("g_CSV")=w_sSQL
    
' 該当する社員がいるかどうかの判定
	if g_rRs.EOF=true then
		w_sFLG="3"
	end if

' 条件の指定が無い場合（すべて出力するかどうかのメッセージ）
	if w_sStartCD = "" AND w_sEndCD = "" AND w_sName = "" AND w_checkDel <> 1 then
		w_sFLG="2"
	end if
	
' 確認メッセージ
	Session.Contents("SELECT")="CSV"
	Response.Redirect "CorEKAKUNIN.asp?FLG=" & w_sFLG

    w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing

%>

