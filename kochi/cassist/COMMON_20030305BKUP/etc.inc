<%

' *************************************************
' *********** ���̑��֐��Q ***** by �Ɠ� **********
' *************************************************

' ***** ������Ɏw��̌����ɖ����Ȃ��ꍇ0��t���� *****
' ��F  Str = "8"  AddZero(Str, 2)  ---> 08
Function AddZero(s, n)
	AddZero = Right("00000000000000000" & CStr(s), n)
End Function

' ***** �V�X�e�����t���當����𐶐����� *****
Function GetDateStr()
	GetDateStr = Year(Date()) & AddZero(Month(Date()), 2) & AddZero(Day(Date()), 2)
End Function

' ***** �V�X�e�����t�{n����̕�����𐶐����� *****
Function GetDateAddStr(n)
	tDate = DateAdd("d", n, Date())
	GetDateAddStr = Year(tDate) & AddZero(Month(tDate), 2) & AddZero(Day(tDate), 2)
End Function
%>