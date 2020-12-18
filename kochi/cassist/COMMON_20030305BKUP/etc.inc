<%

' *************************************************
' *********** その他関数群 ***** by 家入 **********
' *************************************************

' ***** 文字列に指定の桁数に満たない場合0を付ける *****
' 例：  Str = "8"  AddZero(Str, 2)  ---> 08
Function AddZero(s, n)
	AddZero = Right("00000000000000000" & CStr(s), n)
End Function

' ***** システム日付から文字列を生成する *****
Function GetDateStr()
	GetDateStr = Year(Date()) & AddZero(Month(Date()), 2) & AddZero(Day(Date()), 2)
End Function

' ***** システム日付＋n日後の文字列を生成する *****
Function GetDateAddStr(n)
	tDate = DateAdd("d", n, Date())
	GetDateAddStr = Year(tDate) & AddZero(Month(tDate), 2) & AddZero(Day(tDate), 2)
End Function
%>
