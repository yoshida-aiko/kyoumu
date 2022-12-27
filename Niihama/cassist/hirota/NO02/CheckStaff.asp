<SCRIPT LANGUAGE="VBS">
Sub Submit_OnClick()

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
		exit sub
	end if
	
' 名前チェック（HTMLタグを埋め込まれていないか）
	if f_CheckVALUE(w_sName)=false then
		Msgbox "名前に不適応なデータが入力されています。",16,"入力エラー"
		window.event.returnValue = false
		exit sub
	end if
	
' 生年月日チェック
	if w_sYEAR <> "" AND w_sMONTH <> "" AND w_sDAY <> "" then
		if IsDate(w_sYEAR & "/" & w_sMONTH & "/" & w_sDAY)=false then
			Msgbox "名前に不適応なデータが入力されています。",16,"入力エラー"
			window.event.returnValue = false
			exit sub
		else
			w_sBirthday = w_sYEAR & "年" & w_sMONTH & "月" & w_sDAY & "日"
			w_sBirth = "'" & w_sYEAR & "/" & w_sMONTH & "/" & w_sDAY & "'"
		end if
	elseif w_sYEAR ="" AND w_sMONTH = "" AND w_sDAY = "" then
		w_sBirthday="<font color=red>記入無し</font>"
		w_sBirth = "NULL"
	else
		Msgbox "生年月日は年、月、日のすべてに入力して下さい。",16,"入力エラー"
		window.event.returnValue = false
		exit sub
	end if

' 電話番号1チェック
	if w_sTel1 <> "" AND w_sTel2 <> "" AND w_sTel3 <> "" then
		w_sTelphone1= w_sTel1 & "-" & w_sTel2 & "-" & w_sTel3
		w_sTel1="'" & w_sTel1 & "-" & w_sTel2 & "-" & w_sTel3 & "'"
	elseif w_sTel1 = "" AND w_sTel2 = "" AND w_sTel3 = "" then
		w_sTelphone1="<font color=red>記入無し</font>"
		w_sTel1 ="NULL"
	else
		Msgbox "電話番号1は必ずハイフン( - )区切りで入力して下さい。",16,"入力エラー"
		window.event.returnValue = false
		exit sub
	end if

' 電話番号2チェック
	if w_sTel4 <> "" AND w_sTel5 <> "" AND w_sTel6 <> "" then
		w_sTelphone2=w_sTel4 & "-" & w_sTel5 & "-" & w_sTel6
		w_sTel2="'" & w_sTel4 & "-" & w_sTel5 & "-" & w_sTel6 & "'"
	elseif w_sTel4 = "" AND w_sTel5 = "" AND w_sTel6 = "" then
		w_sTelphone2="<font color=red>記入無し</font>"
		w_sTel2 ="NULL"
	else
		Msgbox "電話番号2は必ずハイフン( - )区切りで入力して下さい。",16,"入力エラー"
		window.event.returnValue = false
		exit sub
	end if

' 郵便番号チェック
	if w_sPost1 = "" then
		if w_sPost2 = "" then
			w_sPostPost="<font color=red>記入無し</font>"
			w_sPost = "NULL"
		else
			Msgbox "郵便番号は3桁-4桁、又ははじめの3桁を入力して下さい。",16,"入力エラー"
			window.event.returnValue = false
			exit sub
		end if
	elseif w_sPost2 = "" then
		if Len(w_sPost1) < 3 then
			Msgbox "郵便番号は3桁-4桁、又ははじめの3桁を入力して下さい。",16,"入力エラー"
			window.event.returnValue = false
			exit sub
		end if
		w_sPostPost=w_sPost1
		w_sPost= "'" & w_sPost1 & "'"
	else
		if Len(w_sPost1) < 3 or Len(w_sPost2) < 4 then
			Msgbox "郵便番号は3桁-4桁、又ははじめの3桁を入力して下さい。",16,"入力エラー"
			window.event.returnValue = false
			exit sub
		end if
		w_sPostPost=w_sPost1 & " - " & w_sPost2
		w_sPost= "'" & w_sPost1 & "-" & w_sPost2 & "'"
	end if

' 住所1、住所2チェック
	if w_sAddress1 <> "" then
		if w_sAddress2 <> "" then
			if f_CheckVALUE(w_sAddress1)=false or f_CheckVALUE(w_sAddress2)=false then
				Msgbox "住所に不適応な文字が含まれています。",16,"入力エラー"
				window.event.returnValue = false
				exit sub
			end if
			w_sAdd =w_sAddress1 & "<br>" & w_sAddress2
			w_sAddress1= "'" & w_sAddress1 & "'"
			w_sAddress2= "'" & w_sAddress2 & "'"
		else
			if f_CheckVALUE(w_sAddress1)=false then
				Msgbox "住所に不適応な文字が含まれています。",16,"入力エラー"
				window.event.returnValue = false
				exit sub
			end if
			w_sAdd=w_sAddress1 & "<br>"
			w_sAddress1= "'" & w_sAddress1 & "'"
			w_sAddress2= "NULL"	
		end if
	elseif w_sAddress2 <> "" then
		if f_CheckVALUE(w_sAddress2)=false then
			Msgbox "住所に不適応な文字が含まれています。",16,"入力エラー"
			window.event.returnValue = false
			exit sub
		end if
		w_sAdd= "<br>" & w_sAddress2
		w_sAddress1="NULL"
		w_sAddress2= "'" & w_sAddress2 & "'"
	else
		w_sAdd="<font color=red>記入無し</font><br>"
		w_sAddress1="NULL"
		w_sAddress2="NULL"
	end if

' 備考チェック
	if w_sBikou <> "" then
		if f_CheckVALUE(w_sBikou)=false then
			Msgbox "備考に不適応な文字が含まれています。",16,"入力エラー"
			window.event.returnValue = false
			exit sub
		end if
		w_sIndex=w_sBikou
		w_sBikou= "'" & w_sBikou & "'"
	else
		w_sIndex="<font color=red>記入無し</font>"
		w_sBikou="NULL"
	end if
End Sub
'*******************************************************************
'　　タグが入力されたかどうかを判定
'*******************************************************************
function f_CheckVALUE(p_VALUE)
	f_CheckVALUE = false
    If InStr(p_VALUE, "<") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, ">") <> 0 Then
        Exit Function
    End If
    f_CheckVALUE = true
end function

</SCRIPT>

