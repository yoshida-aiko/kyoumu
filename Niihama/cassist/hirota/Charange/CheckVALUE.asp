<SCRIPT LANGUAGE="VBS">
<!--クライアントサイドのスクリプト

Sub Submit_OnClick()

	w_StartCD = EXPORT.txtStartCD.value
	w_EndCD = EXPORT.txtEndCD.value
	w_Name = EXPORT.txtName.value
	w_CboCheck = EXPORT.checkDel.Checked
	
    SQL = "SELECT * FROM M_社員 WHERE 社員CD >=0"
    
    if EXPORT.txtFileName.value="" then
		EXPORT.txtFileName.value="Sample"
	end if

' 指定条件がない場合
	if w_StartCD = "" AND w_EndCD = "" AND w_Name = "" AND w_CboCheck = false then
		if CheckFileName()=false then
			Msgbox "ファイル名に不正な名前が使われています。次ぎの文字は使うことが出来ません。" _
					& vbcrlf + vbcrlf & "		\ ; : , * < > | ",16,"社員名称入力エラー"
			window.event.returnValue=false
			EXPORT.txtFileName.select
			Exit Sub
		end if
		MsgStr = Msgbox("指定条件がありません。すべてのデータを出力してもよろしいですか？",vbOkCancel + vbInformation,"エクスポート")
			if MsgStr = vbCancel then
				window.event.returnValue = false
				EXPORT.txtStartCD.focus
				Exit Sub
			end if
			SQL = SQL & " ORDER BY 1 ASC"
			EXPORT.SQL.value = SQL
			Exit Sub
	End if
	
' 社員CDの入力チェック
    If w_StartCD <> "" Then
        If w_EndCD <> "" Then
            If gf_bCheckCD(w_StartCD) = False or gf_bCheckCD(w_EndCD) = false Then
                Msgbox "社員CDに文字が含まれています。整数を入力してください。",16,"社員CD入力エラー"
				window.event.returnValue=false
				EXPORT.txtStartCD.select
				Exit Sub
            End If
            w_StartCD = Cint(w_StartCD)
            w_EndCD = Cint(w_EndCD)
            SQL = SQL & " AND 社員CD >=" & w_StartCD & " AND 社員CD<=" & w_EndCD
        Else
            If gf_bCheckCD(w_StartCD) = False Then
               Msgbox "社員CDに文字が含まれています。整数を入力してください。",16,"社員CD入力エラー"
				window.event.returnValue=false
				EXPORT.txtStartCD.select
				Exit Sub
            End If
            w_StartCD = Cint(w_StartCD)
            SQL = SQL & " AND 社員CD >=" & w_StartCD
        End If
    ElseIf w_EndCD <> "" Then
        If gf_bCheckCD(w_EndCD) = False Then
            Msgbox "社員CDに文字が含まれています。整数を入力してください。",16,"社員CD入力エラー"
			window.event.returnValue=false
			EXPORT.txtEndCD.select
			Exit Sub
        End If
        w_EndCD = Cint(w_EndCD)
        SQL = SQL & " AND 社員CD <=" & Cint(w_EndCD)
    End If
    
' 社員名称の入力チェック
    If w_Name <> "" Then
		if gf_bCheckNAME(w_Name) = false then
			Msgbox "社員名称に"<"，または">"が含まれています。",16,"社員名称入力エラー"
			window.event.returnValue=false
			EXPORT.txtName.select
			Exit Sub
		end if
        SQL = SQL & " AND 社員名称 LIKE '%" & w_Name & "%'"
    End If
    If w_CboCheck = true Then
        SQL = SQL & " AND 使用FLG=1"
    End If
    if CheckFileName()=false then
		Msgbox "ファイル名に不正な名前が使われています。次ぎの文字は使うことが出来ません。" _
				& vbcrlf + vbcrlf & "		\ ; : , * < > | ",16,"社員名称入力エラー"
		window.event.returnValue=false
		EXPORT.txtFileName.select
		Exit Sub
	end if
	
' メッセージ
    MsgStr = Msgbox("この条件で出力してもよろしいですか？",vbOkCancel + vbInformation,"エクスポート")
		if MsgStr = vbCancel then
			window.event.returnValue = false
			EXPORT.txtStartCD.focus
			Exit Sub
		End if
	    SQL = SQL & " ORDER BY 1 ASC"
	    EXPORT.SQL.value = SQL
End Sub

'*****************************************************************
'	入力チェック処理（関数）
'*****************************************************************

Function gf_bCheckCD(p_sCD)
    gf_bCheckCD = false
' 社員CDの入力型が数字になっているか？
    If IsNumeric(p_sCD) = False Then
        Exit Function
    End If
' 文字制限（カンマ、負号、小数点、￥マークは受け付けない）
    If InStr(p_sCD, ".") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, "-") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, "+") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, ",") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, "\") <> 0 Then
        Exit Function
    End If
	if p_sCD < 0 or p_sCD > 9999 then
		Exit Function
	End If
    gf_bCheckCD = True
End Function


'*******************************************************************
'　　タグが入力されたかどうかを判定
'*******************************************************************
Function gf_bCheckNAME(p_sNAME)

	gf_bCheckNAME = false
    If InStr(p_sNAME, "<") <> 0 Then
        Exit Function
    End If
    If InStr(p_sNAME, ">") <> 0 Then
        Exit Function
    End If
    gf_bCheckNAME = true

End Function
//-->
</SCRIPT>