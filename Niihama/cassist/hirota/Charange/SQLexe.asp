<%
	On Error Resume Next
	Err.Clear

	Dim w_cCn,w_rRs
	Dim w_sSQL
	
	w_CD=Request.Form("CD")
	w_NAME=Request.Form("NAME")
	w_BIRTHDAY=Request.Form("BIRTHDAY")
	w_TELL1=Request.Form("TELL1")
	w_TELL2=Request.Form("TELL2")
	w_POST=Request.Form("POST")
	w_ADDRESS1=Request.Form("ADDRESS1")
	w_ADDRESS2=Request.Form("ADDRESS2")
	w_BIKOU=Request.Form("BIKOU")

' オブジェクト定義
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")
 	w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
 	w_rRs.Open "M_社員",w_cCn,2,2

 
	if Session.Contents("FLG")="ADDNEW" then
	
	' 社員CD重複チェック
		w_sSQL = "SELECT 社員CD FROM M_社員 WHERE 使用FLG=1 AND 社員CD =" & w_CD
		if gf_SQLexe(w_sSQL)=false then
			Response.Redirect "SQLerror.asp"
		end if
	' 重複チェック
		if w_rRs.EOF=false then
			Session.Contents("ErrorCD")=w_CD
			Response.Redirect "WStaffMsg.asp?WStaff=1"
		end if
	' 使用FLG=0の社員データが残っている場合
		w_sSQL = "SELECT 社員CD FROM M_社員 WHERE 使用FLG=0 AND 社員CD=" & w_CD
		if gf_SQLexe(w_sSQL)=false then
			Response.Redirect "SQLerror.asp"
		end if
	' 社員CDの重複チェック(使用FLG=0の場合)
	 	if w_rRs.EOF=false then
			Session.Contents("社員CD")=w_CD
			Session.Contents("社員名称")=w_NAME
			Session.Contents("生年月日")=w_BIRTHDAY
			Session.Contents("電話番号1")=w_TELL1
			Session.Contents("電話番号2")=w_TELL2
			Session.Contents("郵便")=w_POST
			Session.Contents("住所1")=w_ADDRESS1
			Session.Contents("住所2")=w_ADDRESS2
			Session.Contents("備考")=w_BIKOU
			Response.Redirect "WStaffMsg.asp?WStaff=2"
		end if
		if f_ADDNEW()=false then
			Response.Redirect "SQLerror.asp"
		end if
		
	elseif Session.Contents("FLG")="UPDATE" then
		if f_UPDATE()=false then
			Response.Redirect "SQLerror.asp"
		end if
	else
		if f_DELETE()=false then
			Response.Redirect "SQLerror.asp"
		end if
	end if
	
	Response.Redirect "INitiran.asp"

	w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing



'**************************************************************
'			新規
'**************************************************************
Function f_ADDNEW()
	f_ADDNEW=false
' 新規登録のSQL文の作成	
	w_sSQL = "INSERT INTO M_社員 (社員CD,社員名称,生年月日,電話番号1,電話番号2,"
    w_sSQL = w_sSQL & "郵便,住所1,住所2,備考,使用FLG)"
    w_sSQL = w_sSQL & " VALUES (" & Request.Form("CD")
    w_sSQL = w_sSQL & "," & Request.Form("NAME")
    w_sSQL = w_sSQL & "," & Request.Form("BIRTHDAY")
    w_sSQL = w_sSQL & "," & Request.Form("TELL1")
    w_sSQL = w_sSQL & "," & Request.Form("TELL2")
    w_sSQL = w_sSQL & "," & Request.Form("POST")
    w_sSQL = w_sSQL & "," & Request.Form("ADDRESS1")
    w_sSQL = w_sSQL & "," & Request.Form("ADDRESS2")
    w_sSQL = w_sSQL & "," & Request.Form("BIKOU")
    w_sSQL = w_sSQL & ",1)"

	if gf_SQLexe(w_sSQL)=false then
		Exit Function
	end if
	f_ADDNEW=true
End Function



'**************************************************************
'			修正
'**************************************************************
Function f_UPDATE()
		f_UPDATE=false
	' 修正SQL
		w_sSQL ="UPDATE M_社員 SET 社員名称=" & Request.Form("NAME")
		w_sSQL = w_sSQL & ",生年月日=" & Request.Form("BIRTHDAY")
		w_sSQL = w_sSQL & ",電話番号1=" & Request.Form("TELL1")
		w_sSQL = w_sSQL & ",電話番号2=" & Request.Form("TELL2")
		w_sSQL = w_sSQL & ",郵便=" & Request.Form("POST")
		w_sSQL = w_sSQL & ",住所1=" & Request.Form("ADDRESS1")
		w_sSQL = w_sSQL & ",住所2=" & Request.Form("ADDRESS2")
		w_sSQL = w_sSQL & ",備考=" & Request.Form("BIKOU")
		w_sSQL = w_sSQL & ",使用FLG=1 WHERE 社員CD=" & Request.Form("CD")
		
		if gf_SQLexe(w_sSQL)=false then
			Exit Function
		end if
		f_UPDATE=true
End Function



'**************************************************************
'			削除
'**************************************************************
Function f_DELETE()
	    f_DELETE=false
	' 社員データ削除のSQL文
	    w_sSQL = "UPDATE M_社員 SET 使用FLG=0 WHERE 社員CD=" & Request.Form("CD")

		if gf_SQLexe(w_sSQL)=false then
			Exit Function
		end if
		f_DELETE=true
End Function



'**************************************************************
'			SQL実行関数
'**************************************************************

Function gf_SQLexe(p_sSQL)
	Set w_rRs = w_cCn.Execute(p_sSQL)
 ' SQL実行時のエラー処理
	if Err then
		Session.Contents("SQLerror")=Err.Description
		gf_SQLexe=false
	end if
	On Error Goto 0
	gf_SQLexe=true
End Function

%>