
<!--#INCLUDE FILE="ExportEXCELhidden.asp"-->

<SCRIPT LANGUAGE="VBS">

'*****************************************************************************************************
'											EXCEL�̍쐬
'*****************************************************************************************************

<!-- �u���E�U���̃X�N���v�g

' �ϐ��錾
	Dim objExcelApp
	Dim FileName
	Dim i
	Dim j
	
	On Error Resume Next
	Err.Clear
	
	FileName = "<%= w_cboName & w_FileName %>.xls"
	'FileName = "Y:\<%= w_cboName %>\<%= w_sFileName %>.xls"
	
' �u���E�U��EXCEL�𗧂��グ��
	Set objExcelApp = CreateObject("Excel.Application")

' �I�u�W�F�N�g�G���[�Ȃ�΃��b�Z�[�W
	If Err Then
		ErrMESSAGE()
	Else
		On Error goto 0
		' �����e���v���[�g��Open
		' �� �V�K���[�N�V�[�g�̍쐬�̏ꍇ�́A�ς��� objExcelApp.Workbooks.Add
		objExcelApp.Workbooks.Open "<%= GetURLPath() & "�Ј��}�X�^.xls" %>",,True
									'"\\WEBSVR_2\infogram\hirota\No02\�Ј��}�X�^.xls",,True
									'<%= GetURLPath() & "demo.xlt" %>
		Set objExcelBook = objExcelApp.ActiveWorkbook
		Set objExcelSheets = objExcelBook.Worksheets
		Set objExcelSheet = objExcelBook.Sheets(1)
		
		objExcelSheet.Activate
		objExcelApp.Application.Visible = True

'------------------------------------------EXCEL�o��--------------------------------------------------
        i = 7	' 8�s�ڂ���̏����o��

        j = 0	' �J�E���g��

            objExcelSheet.Cells(3, 8).Value = "������F" & Date	' �����̓��t
            
        ' EXCEL�V�[�g�ɏ�����
            <% w_rRs.MoveFirst %>
            <% Do While Not w_rRs.EOF %>
					j = j + 1
                    i = i + 1
                    objExcelSheet.Cells(i, 1).Value = "<%= w_rRs("�Ј�CD") %>"
                    objExcelSheet.Cells(i, 2).Value = "<%= w_rRs("�Ј�����") %>"
                    objExcelSheet.Cells(i, 4).Value = "<%= w_rRs("���N����") %>"
                    objExcelSheet.Cells(i, 5).Value = "<%= w_rRs("�d�b�ԍ�1") %>"
                    objExcelSheet.Cells(i, 7).Value = "<%= w_rRs("�d�b�ԍ�2") %>"
                    i = i + 1
                    objExcelSheet.Cells(i, 2).Value = "<%= w_rRs("�X��") %>"
                    objExcelSheet.Cells(i, 4).Value = "<%= w_rRs("�Z��1") %>"
                    objExcelSheet.Cells(i, 8).Value = "<%= w_rRs("�Z��2") %>"
                <% w_rRs.MoveNext %>
            <% Loop %>
            
       ' �u�b�N�ɏ������񂾃f�[�^��ۑ�
            objExcelBook.SaveAs	"<%= w_cboName & w_FileName %>.xls"
										'objExcelBook.Save	"Y:\�A�c\����p�v���O����\Test.xls"
	   ' �J�����u�b�N�����	
            objExcelBook.close
            
       ' EXCEL�����
            objExcelApp.Quit
            
       ' �I�u�W�F�N�g�̊J��
            Set objExcelSheet = Nothing
            Set objExcelBook = Nothing
            Set objExcelSheets = Nothing
            Set objExcelApp = Nothing
            <% w_rRs.Close %>
			<% w_cCn.Close %>
			<% Set w_rRs = Nothing %>
			<% Set w_cCn = Nothing %>

            FinishMESSAGE()
	End If

'--------------------------------------HTML���b�Z�[�W---------------------------------------------------

' �������b�Z�[�W
	Function FinishMESSAGE()
		document.write"<html>"
		document.write"<head><title>�Ј��Ǘ�</title><base target=Right></head>"
		document.write"<body>"
			document.write"<h3 align=center>�� EXCEL�o�� ��</h3><hr>"
			document.write"<h2 align=center><font color=red>EXCEL�o�͂��������܂����I</font></h2>"
		document.write"<table align=center>"
			document.write"<tr>"
				document.write"<td>�o�͏ꏊ</td>"
				document.write"<td>�F</td>"
				document.write"<td>" & FileName & "</td>"
			document.write"</tr>"
			document.write"<tr>"
				document.write"<td>�o�͌���</td>"
				document.write"<td>�F</td>"
				document.write"<td>" & j & " ��</td>"
			document.write"</tr>"
		document.write"</table>"
		document.write"<br>"
		document.write"<table align=center width=20%>"
			document.write"<tr>"
				document.write"<form action=EXCEL.asp target=Right>"
					document.write"<td align=center><p align=center><input type=submit value=�߂�></td>"
				document.write"</form>"
				document.write"<form action=INitiran.asp target=Right>"
					document.write"<td align=center valign=bottom><input type=submit value=�ꗗ></td>"
				document.write"</form>"
			document.write"<tr>"
		document.write"</table>"
		document.write"</body>"
		document.write"</html>"
	End Function
	
' �G���[���b�Z�[�W
	Function ErrMESSAGE()
		document.write"<html>"
		document.write"<head>"
			document.write"<title>�Ј��Ǘ�</title>"
			document.write"<base target=Right>"
		document.write"</head>"
		document.write"<body>"
			document.write"<h3 align=center>�� �o�̓G���[ ��</h3>"
				document.write"<hr><br>"
			document.write"<h4 align=center><font color=red>Excel�̋N���Ɏ��s���܂����B<br>"
			document.write"���f�[�^�x�[�X�̃f�[�^���o�͂��邱�Ƃ��o���܂���ł����B</font></h4>"
			document.write"<p align=center>"
				document.write "�G���[�F" & Err.description
			document.write"</p>"
			document.write"<p align=center>"
			document.write"<form action=EXCEL.asp target=Right id=form1 name=form1>"
			document.write"<input type=submit value=�߂� id=submit1 name=submit1>"
			document.write"</form></p>"
		document.write"</body>"
	document.write"</html>"
	End Function
//-->
</SCRIPT>
