<%@ Language=VBScript %>
<%
	public m_sUser
	public m_sPass
	public m_sConnect	
	
	public m_sPDFPath

	public m_iNendo 
	public m_iSiken
	public m_iGakunen
	public m_sClass

	public m_bRet
	public m_sERR

    'Ҳ�ٰ�ݎ��s
    Call Main()

Sub Main()
	dim w_oReport 

'	On Error Resume Next
'	Err.Clear

	Response.Expires = 0					'������ر / PG   SIDE
	Response.AddHeader "Pragma", "No-Cache"	'������ر / HTML SIDE

	'�p�����[�^
	m_sUser = Request.form("DBUser")
	m_sPass = request("DBPass")
	m_sConnect = request("DBConnection")

'	m_sPDFPath = server.mappath(".") & "\" & request("PDFPath")
	m_sPDFPath = request("PDFPath")

	m_iNendo = request("txtNendo")
	m_iSiken = request("txtSiken")
	m_iGakunen = request("txtGakunen")
	m_sClass = request("txtClass")


	'���
	set w_oReport = CreateObject("HAN0240.clsHan0240")
response.write "CreateObject <BR>"
    
	w_oReport.DBUserName = m_sUser
    w_oReport.DBPassWord = m_sPass
    w_oReport.DBConnection = m_sConnect
response.write "DB CONNECT=" & m_sUser & "/" & m_sPass & "/" & m_sConnect & "<BR>"

	w_oReport.ExportPath = m_sPDFPath
response.write "PDF PATH=" & m_sPDFPath

    m_bRet =  w_oReport.ExportPDF(m_iNendo, m_iSiken, m_iGakunen, m_sClass) 
response.write "m_bRet=" & m_bRet

	m_sERR = w_oReport.ErrorMessage
response.write "ERR=" & m_sERR

	
	set w_oReport = Nothing

	call s_ShowPage()

End Sub



sub s_ShowPage()
%>

<HTML>
<HREAD></HEAD>
<BODY>

<TABLE>
<TR>
<TD>User=<%=m_sUser%></TD>
<TD>Password=<%=m_sPass%></TD>
<TD>Connection=<%=m_sConnect%></TD>
</TR>
</TABLE>

<TABLE>
<TR>
<TD>PDFPath=<%=m_sPDFPath%></TD>
</TR>
</TABLE>

<TABLE>
<TR>
<TD>�N�x=<%=m_iNendo%></TD>
<TD>����=<%=m_iSiken%></TD>
<TD>�w�N=<%=m_iGakunen%></TD>
<TD>�N���X=<%=m_sClass%></TD>
</TR>
</TABLE>

���|�[�gret=<%=m_bRet%><BR>
err:<%=m_sERR%>

</BODY>
</HTML>

<%
end sub 
%>

