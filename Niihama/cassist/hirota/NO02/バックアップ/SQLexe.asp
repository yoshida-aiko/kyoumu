<%
	On Error Resume Next
	Err.Clear
	Dim w_cCn,w_rRs
	Dim w_sSQL

' IuWFNgè`
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

 	w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
 	w_rRs.Open "M_Ðõ",w_cCn,2,2
 	
 e_SELECT = Request.Form("FLG")
 Select case e_SELECT
	'**************************************************************
	'			VK
	'**************************************************************
	 case "1"
	 
	 ' gpFLG=0ÌÐõf[^ªcÁÄ¢éê
	 	w_sSQL = "SELECT ÐõCD FROM M_Ðõ WHERE gpFLG=0 AND ÐõCD=" & Request.Form("ÐõCD")
	 	Set w_rRs = w_cCn.Execute(w_sSQL)
	 	
	 ' SQLÀsÌG[
	 	if Err then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if

		On Error Goto 0
		
	' ÐõCDÌd¡`FbN(gpFLG=0Ìê)
	 	if w_rRs.EOF=false then
			Session.Contents("Ðõ¼Ì")=Request.Form("Ðõ¼Ì")
			Session.Contents("¶Nú")=Request.Form("¶Nú")
			Session.Contents("dbÔ1")=Request.Form("dbÔ1")
			Session.Contents("dbÔ2")=Request.Form("dbÔ2")
			Session.Contents("XÖ")=Request.Form("XÖ")
			Session.Contents("Z1")=Request.Form("Z1")
			Session.Contents("Z2")=Request.Form("Z2")
			Session.Contents("õl")=Request.Form("õl")
			Session.Contents("ÐõCD")=Request.Form("ÐõCD")
			w_FLG="3"
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
		
	' VKo^ÌSQL¶Ìì¬	
		w_sSQL = "INSERT INTO M_Ðõ (ÐõCD,Ðõ¼Ì,¶Nú,dbÔ1,dbÔ2,"
	    w_sSQL = w_sSQL & "XÖ,Z1,Z2,õl,gpFLG)"
	    w_sSQL = w_sSQL & " VALUES (" & Request.Form("ÐõCD") & ",'" & Request.Form("Ðõ¼Ì") & "'"
	    w_sSQL = w_sSQL & "," & Request.Form("¶Nú")
	    w_sSQL = w_sSQL & "," & Request.Form("dbÔ1")
	    w_sSQL = w_sSQL & "," & Request.Form("dbÔ2")
	    w_sSQL = w_sSQL & "," & Request.Form("XÖ")
	    w_sSQL = w_sSQL & "," & Request.Form("Z1")
	    w_sSQL = w_sSQL & "," & Request.Form("Z2")
	    w_sSQL = w_sSQL & "," & Request.Form("õl") & ",1)"

		if gf_SQLexe(w_sSQL)=false then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if
		w_sFLG = "1"
		
	'**************************************************************
	'			C³
	'**************************************************************
	case "2"
		
	' C³SQL
		w_sSQL ="UPDATE M_Ðõ SET Ðõ¼Ì='" & Request.Form("Ðõ¼Ì") & "'"
		w_sSQL = w_sSQL & ",¶Nú=" & Request.Form("¶Nú")
		w_sSQL = w_sSQL & ",dbÔ1=" & Request.Form("dbÔ1")
		w_sSQL = w_sSQL & ",dbÔ2=" & Request.Form("dbÔ2")
		w_sSQL = w_sSQL & ",XÖ=" & Request.Form("XÖ")
		w_sSQL = w_sSQL & ",Z1=" & Request.Form("Z1")
		w_sSQL = w_sSQL & ",Z2=" & Request.Form("Z2")
		w_sSQL = w_sSQL & ",õl=" & Request.Form("õl")
		w_sSQL = w_sSQL & ",gpFLG=1 WHERE ÐõCD=" & Request.Form("ÐõCD")
		
		if gf_SQLexe(w_sSQL)=false then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if
		w_sFLG="2"	

	'**************************************************************
	'			í
	'**************************************************************
	case "3"
	    
	' Ðõf[^íÌSQL¶
	    w_SQL = "UPDATE M_Ðõ SET gpFLG=0 WHERE ÐõCD=" & Request.Form("ÐõCD")

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
'			SQLÀsÖ
'**************************************************************
function gf_SQLexe(p_sSQL)
	Set w_rRs = w_cCn.Execute(p_sSQL)
 ' SQLÀsÌG[
	if Err then
		gf_SQLexe=false
	end if
	On Error Goto 0
	gf_SQLexe=true
end function

%>