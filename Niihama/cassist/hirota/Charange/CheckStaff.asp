<SCRIPT LANGUAGE="VBS">
<!--クライアントサイドのスクリプト

Sub Submit_OnClick()

Dim w_sCD,w_sName,w_sYEAR,w_sMONTH,w_sDAY
Dim w_sTel1,w_sTel2,w_sTel3,w_sTel4,w_sTel5,w_sTel6
Dim w_sPost1,w_sPost2,w_sAddress1,w_sAddress2,w_sBikou

' テキストに入力されたデータを変数に格納する。
	w_sCD = Tanpyou.txtCD.value
	w_sName = Tanpyou.txtName.value
	w_sYEAR = Tanpyou.txtYEAR.value
	w_sMONTH = Tanpyou.txtMONTH.value
	w_sDAY = Tanpyou.txtDAY.value
	w_sTel1 = Tanpyou.txtTel1.value
	w_sTel2 = Tanpyou.txtTel2.value
	w_sTel3 = Tanpyou.txtTel3.value
	w_sTel4 = Tanpyou.txtTel4.value
	w_sTel5 = Tanpyou.txtTel5.value
	w_sTel6 = Tanpyou.txtTel6.value
	w_sPost1 = Tanpyou.txtPost1.value
	w_sPost2 = Tanpyou.txtPost2.value
	w_sAddress1 = Tanpyou.txtAddress1.value
	w_sAddress2 = Tanpyou.txtAddress2.value
	w_sBikou = Tanpyou.txtBikou.value
	
	if w_sCD = "" or w_sName = "" then
		Msgbox "社員CDと社員名称は必ず入力してください。",16,"入力エラー"
		window.event.returnValue = false
		Tanpyou.txtCD.focus
		exit sub
	end if
	
' 名前チェック（HTMLタグを埋め込まれていないか）
	if f_CheckVALUE(w_sName)=false then
		Msgbox "名前に不適応なデータが入力されています。",16,"入力エラー"
		window.event.returnValue = false
		Tanpyou.txtName.select
		exit sub
	end if

	Tanpyou.CD.value=w_sCD
	Tanpyou.NAME.value= "'" & w_sName & "'"
	
' 生年月日チェック
	if w_sYEAR <> "" AND w_sMONTH <> "" AND w_sDAY <> "" then
		if IsDate(w_sYEAR & "/" & w_sMONTH & "/" & w_sDAY)=false then
			Msgbox "生年月日に存在しない日付が入力されています。",16,"入力エラー"
			Tanpyou.txtDAY.select
			window.event.returnValue = false
			exit sub
		else
			Tanpyou.BIRTHDAY.value ="'" & w_sYEAR & "/" & w_sMONTH & "/" & w_sDAY & "'"
		end if
	elseif w_sYEAR ="" AND w_sMONTH = "" AND w_sDAY = "" then
		Tanpyou.BIRTHDAY.value = "NULL"
	else
		Msgbox "生年月日は年、月、日のすべてに入力して下さい。",16,"入力エラー"
		window.event.returnValue = false
		Tanpyou.txtYEAR.focus
		exit sub
	end if

' 電話番号1チェック
	if w_sTel1 <> "" AND w_sTel2 <> "" AND w_sTel3 <> "" then
		Tanpyou.TELL1.value = "'" & w_sTel1 & "-" & w_sTel2 & "-" & w_sTel3 & "'"
	elseif w_sTel1 = "" AND w_sTel2 = "" AND w_sTel3 = "" then
		Tanpyou.TELL1.value = "NULL"
	else
		Msgbox "電話番号1は必ずハイフン( - )区切りで入力して下さい。",16,"入力エラー"
		window.event.returnValue = false
		Tanpyou.txtTel1.focus
		exit sub
	end if

' 電話番号2チェック
	if w_sTel4 <> "" AND w_sTel5 <> "" AND w_sTel6 <> "" then
		Tanpyou.TELL2.value = "'" & w_sTel4 & "-" & w_sTel5 & "-" & w_sTel6 & "'"
	elseif w_sTel4 = "" AND w_sTel5 = "" AND w_sTel6 = "" then
		Tanpyou.TELL2.value = "NULL"
	else
		Msgbox "電話番号2は必ずハイフン( - )区切りで入力して下さい。",16,"入力エラー"
		window.event.returnValue = false
		Tanpyou.txtTel4.focus
		exit sub
	end if

' 郵便番号チェック
	if w_sPost1 = "" then
		if w_sPost2 = "" then
			Tanpyou.POST.value = "NULL"
		else
			Msgbox "郵便番号は3桁-4桁、又ははじめの3桁を入力して下さい。",16,"入力エラー"
			window.event.returnValue = false
			Tanpyou.txtPost1.focus
			exit sub
		end if
	elseif w_sPost2 = "" then
		if Len(w_sPost1) < 3 then
			Msgbox "郵便番号は3桁-4桁、又ははじめの3桁を入力して下さい。",16,"入力エラー"
			window.event.returnValue = false
			Tanpyou.txtPost1.focus
			exit sub
		end if
		Tanpyou.POST.value = "'" & w_sPost1 & "'"
	else
		if Len(w_sPost1) < 3 or Len(w_sPost2) < 4 then
			Msgbox "郵便番号は3桁-4桁、又ははじめの3桁を入力して下さい。",16,"入力エラー"
			window.event.returnValue = false
			Tanpyou.txtPost1.focus
			exit sub
		end if
		Tanpyou.POST.value = "'" & w_sPost1 & "-" & w_sPost2 & "'"
	end if

' 住所1、住所2チェック
	if w_sAddress1 <> "" then
		if w_sAddress2 <> "" then
			if f_CheckVALUE(w_sAddress1)=false or f_CheckVALUE(w_sAddress2)=false then
				Msgbox "住所に不適応な文字が含まれています。",16,"入力エラー"
				window.event.returnValue = false
				Tanpyou.txtAddress1.focus
				exit sub
			end if
			Tanpyou.ADDRESS1.value= "'" & w_sAddress1 & "'"
			Tanpyou.ADDRESS2.value= "'" & w_sAddress2 & "'"
		else
			if f_CheckVALUE(w_sAddress1)=false then
				Msgbox "住所に不適応な文字が含まれています。",16,"入力エラー"
				window.event.returnValue = false
				Tanpyou.txtAddress1.select
				exit sub
			end if
			Tanpyou.ADDRESS1.value= "'" & w_sAddress1 & "'"
			Tanpyou.ADDRESS2.value= "NULL"	
		end if
	elseif w_sAddress2 <> "" then
		if f_CheckVALUE(w_sAddress2)=false then
			Msgbox "住所に不適応な文字が含まれています。",16,"入力エラー"
			window.event.returnValue = false
			Tanpyou.txtAddress2.select
			exit sub
		end if
		Tanpyou.ADDRESS1.value="NULL"
		Tanpyou.ADDRESS2.value= "'" & w_sAddress2 & "'"
	else
		Tanpyou.ADDRESS1.value="NULL"
		Tanpyou.ADDRESS2.value="NULL"
	end if

' 備考チェック
	if w_sBikou <> "" then
		if f_CheckVALUE(w_sBikou)=false then
			Msgbox "備考に不適応な文字が含まれています。",16,"入力エラー"
			window.event.returnValue = false
			Tanpyou.txtBikou.select
			exit sub
		end if
		Tanpyou.BIKOU.value = "'" & w_sBikou & "'"
	else
		Tanpyou.BIKOU.value = "NULL"
	end if
	MsgOK()
End Sub

'*******************************************************************
'　　タグが入力されたかどうかを判定
'*******************************************************************
Function f_CheckVALUE(p_VALUE)
	f_CheckVALUE = false
    If InStr(p_VALUE, "<") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, ">") <> 0 Then
        Exit Function
    End If
    f_CheckVALUE = true
End Function
//-->
</SCRIPT>