
<!--#INCLUDE FILE="ExportEXCELhidden.asp"-->

<SCRIPT LANGUAGE="VBS">

'*****************************************************************************************************
'											EXCELの作成
'*****************************************************************************************************

<!-- ブラウザ側のスクリプト

' 変数宣言
	Dim objExcelApp
	Dim FileName
	Dim i
	Dim j
	
	On Error Resume Next
	Err.Clear
	
	FileName = "<%= w_cboName & w_FileName %>.xls"
	'FileName = "Y:\<%= w_cboName %>\<%= w_sFileName %>.xls"
	
' ブラウザのEXCELを立ち上げる
	Set objExcelApp = CreateObject("Excel.Application")

' オブジェクトエラーならばメッセージ
	If Err Then
		ErrMESSAGE()
	Else
		On Error goto 0
		' 既存テンプレートのOpen
		' ※ 新規ワークシートの作成の場合は、変わりに objExcelApp.Workbooks.Add
		objExcelApp.Workbooks.Open "<%= GetURLPath() & "社員マスタ.xls" %>",,True
									'"\\WEBSVR_2\infogram\hirota\No02\社員マスタ.xls",,True
									'<%= GetURLPath() & "demo.xlt" %>
		Set objExcelBook = objExcelApp.ActiveWorkbook
		Set objExcelSheets = objExcelBook.Worksheets
		Set objExcelSheet = objExcelBook.Sheets(1)
		
		objExcelSheet.Activate
		objExcelApp.Application.Visible = True

'------------------------------------------EXCEL出力--------------------------------------------------
        i = 7	' 8行目からの書き出し

        j = 0	' カウント数

            objExcelSheet.Cells(3, 8).Value = "印刷日：" & Date	' 今日の日付
            
        ' EXCELシートに書込み
            <% w_rRs.MoveFirst %>
            <% Do While Not w_rRs.EOF %>
					j = j + 1
                    i = i + 1
                    objExcelSheet.Cells(i, 1).Value = "<%= w_rRs("社員CD") %>"
                    objExcelSheet.Cells(i, 2).Value = "<%= w_rRs("社員名称") %>"
                    objExcelSheet.Cells(i, 4).Value = "<%= w_rRs("生年月日") %>"
                    objExcelSheet.Cells(i, 5).Value = "<%= w_rRs("電話番号1") %>"
                    objExcelSheet.Cells(i, 7).Value = "<%= w_rRs("電話番号2") %>"
                    i = i + 1
                    objExcelSheet.Cells(i, 2).Value = "<%= w_rRs("郵便") %>"
                    objExcelSheet.Cells(i, 4).Value = "<%= w_rRs("住所1") %>"
                    objExcelSheet.Cells(i, 8).Value = "<%= w_rRs("住所2") %>"
                <% w_rRs.MoveNext %>
            <% Loop %>
            
       ' ブックに書き込んだデータを保存
            objExcelBook.SaveAs	"<%= w_cboName & w_FileName %>.xls"
										'objExcelBook.Save	"Y:\廣田\教育用プログラム\Test.xls"
	   ' 開いたブックを閉じる	
            objExcelBook.close
            
       ' EXCELを閉じる
            objExcelApp.Quit
            
       ' オブジェクトの開放
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

'--------------------------------------HTMLメッセージ---------------------------------------------------

' 成功メッセージ
	Function FinishMESSAGE()
		document.write"<html>"
		document.write"<head><title>社員管理</title><base target=Right></head>"
		document.write"<body>"
			document.write"<h3 align=center>★ EXCEL出力 ★</h3><hr>"
			document.write"<h2 align=center><font color=red>EXCEL出力が完了しました！</font></h2>"
		document.write"<table align=center>"
			document.write"<tr>"
				document.write"<td>出力場所</td>"
				document.write"<td>：</td>"
				document.write"<td>" & FileName & "</td>"
			document.write"</tr>"
			document.write"<tr>"
				document.write"<td>出力件数</td>"
				document.write"<td>：</td>"
				document.write"<td>" & j & " 件</td>"
			document.write"</tr>"
		document.write"</table>"
		document.write"<br>"
		document.write"<table align=center width=20%>"
			document.write"<tr>"
				document.write"<form action=EXCEL.asp target=Right>"
					document.write"<td align=center><p align=center><input type=submit value=戻る></td>"
				document.write"</form>"
				document.write"<form action=INitiran.asp target=Right>"
					document.write"<td align=center valign=bottom><input type=submit value=一覧></td>"
				document.write"</form>"
			document.write"<tr>"
		document.write"</table>"
		document.write"</body>"
		document.write"</html>"
	End Function
	
' エラーメッセージ
	Function ErrMESSAGE()
		document.write"<html>"
		document.write"<head>"
			document.write"<title>社員管理</title>"
			document.write"<base target=Right>"
		document.write"</head>"
		document.write"<body>"
			document.write"<h3 align=center>■ 出力エラー ■</h3>"
				document.write"<hr><br>"
			document.write"<h4 align=center><font color=red>Excelの起動に失敗しました。<br>"
			document.write"※データベースのデータを出力することが出来ませんでした。</font></h4>"
			document.write"<p align=center>"
				document.write "エラー：" & Err.description
			document.write"</p>"
			document.write"<p align=center>"
			document.write"<form action=EXCEL.asp target=Right id=form1 name=form1>"
			document.write"<input type=submit value=戻る id=submit1 name=submit1>"
			document.write"</form></p>"
		document.write"</body>"
	document.write"</html>"
	End Function
//-->
</SCRIPT>
