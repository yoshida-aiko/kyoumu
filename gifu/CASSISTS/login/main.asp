<%@ Language=VBScript %>

<%
'***********************************************************
'
'�@�V�X�e�����@�F�@���������V�X�e��
'�@���@���@���@�F�@�f�[�^����
'�@�v���O����ID�F�@
'�@�@�@�@�@�@�\�F�@�w�Ѓf�[�^�̌������ʕ\��
'
'-----------------------------------------------------------
'
'�@���@�@�@�@���F�@
'					mode			:���샂�[�h
'										�󔒁F�@�����\��
'										DISP:	�������ʕ\��
'					GAKUNEN			:�w�N
'					GAKKA			:�w��
'					CLASS			:�N���X
'					MEISYO			:����
'					GAKUSEKI_BANGOU	:�w�Дԍ�
'					SEX				:����
'					GAKUSEI_BANGOU	:�w���ԍ�
'					IDOU			:�ٓ�
'					BUKATUDO_T		:���w�N���u
'					BUKATUDO_G		:���݃N���u
'					RYO				:��
'
'�@�ρ@�@�@�@���F�@
'�@���@�@�@�@�n�F�@
'�@���@�@�@�@���F�@
'
'-----------------------------------------------------------
'
'�@��@�@�@�@���F�@2001/03/19�@�Ɠ��@��^
'
'***********************************************************
%>



<% '*** ASP���ʃ��W���[���錾 *** %>

<!-- #include file="../common/common.asp" -->

<%
'*** �O���[�o���ϐ� ***

	Dim CurrentYear, MaxDisp

	CurrentYear = Year(Date())
	'�I�I�I�I���ӁF�e�X�g�łQ�O�O�O�N�������Ă��܂�
	CurrentYear = "2000"
	MaxDisp = 20

	If Request("np") = "" Then
		NowPage = 0
	Else
		NowPage = Request("np")
	End If

'*** ���C������ ***

	'���C�����[�`�����s
	Call Main()

'*** �d�m�c ***

%>



<%

Sub Main()
'***********************************************************
'�@�@�@�\�F�@�{�`�r�o�̃��C�����[�`��
'�@�ԁ@�l�F�@����
'�@���@���F�@����
'�@�ځ@�ׁF�@����
'�@���@�l�F�@����
'***********************************************************

	'On Error Resume Next
	'Err.Clear


	Call ShowPage()

	'On Error Goto 0
	'Err.Clear

End Sub


'*** �֐���` ***


Sub SetBlank()
'***********************************************************
'�@�@�@�\�F�@�S���ڂ��󔒂ɏ�����
'�@�ԁ@�l�F�@����
'�@���@���F�@����
'�@�ځ@�ׁF�@����
'�@���@�l�F�@����
'***********************************************************

End Sub


Sub SearchDisp()
'***********************************************************
'�@�@�@�\�F�@�������ʂ̕\��
'�@�ԁ@�l�F�@����
'�@���@���F�@����
'�@�ځ@�ׁF�@����
'�@���@�l�F�@����
'***********************************************************

	Dim I
	Dim Conn, RS
	Dim OrCon, OrSes

	' ***** CONN *****
	if gf_ConnOpenOLE(OrSes, OrCon) = false then
		Response.Write Err.Description & "<br>"
	end if
	
	' ***** SQL���� *****
	tSql = ""
	tSql = tSql & "		select "
	tSql = tSql & "			T11.T11_GAKUSEI_NO, "
	tSql = tSql & "			T11.T11_SIMEI, "
	tSql = tSql & "			T11.T11_NYUGAKU_KBN, "
	tSql = tSql & "			T13.T13_NENDO, "
	tSql = tSql & "			T13.T13_GAKUSEKI_NO, "
	tSql = tSql & "			T13.T13_GAKKA_CD, "
	tSql = tSql & "			T13.T13_GAKUNEN, "
	tSql = tSql & "			T13.T13_CLASS, "
	tSql = tSql & "			T13.T13_SYUSEKI_NO1, "
	tSql = tSql & "			T13.T13_SYUSEKI_NO2, "
	tSql = tSql & "			T13.T13_CLUB_1, "
	tSql = tSql & "			T13.T13_CLUB_2, "
	tSql = tSql & "			T11.T11_SEIBETU, "
	tSql = tSql & "			T13.T13_RYOSEI_KBN, "
	tSql = tSql & "			KBN_NYUGAKU.M01_SYOBUNRUIMEI NYUGAKU_KBN, "
	tSql = tSql & "			M02.M02_GAKKARYAKSYO, "
	tSql = tSql & "			M17.M17_BUKATUDOMEI, "
	tSql = tSql & "			KBN_RYOSEI.M01_SYOBUNRUIMEI RYOSEI_KBN "
	tSql = tSql & "		from "
	tSql = tSql & "			T11_GAKUSEKI T11, "
	tSql = tSql & "			T13_GAKU_NEN T13, "
	tSql = tSql & "			M01_KUBUN KBN_NYUGAKU, "
	tSql = tSql & "			M01_KUBUN KBN_RYOSEI, "
	tSql = tSql & "			M02_GAKKA M02, "
	tSql = tSql & "			M17_BUKATUDO M17, "

	tSql = tSql & "		(select "
	tSql = tSql & "			T13.T13_GAKUSEI_NO, "
	tSql = tSql & "			max(T13.T13_NENDO) GMAX "
	tSql = tSql & "		from "
	tSql = tSql & "			T11_GAKUSEKI T11, "
	tSql = tSql & "			T13_GAKU_NEN T13 "
	tSql = tSql & "		where "
	tSql = tSql & "			T11.T11_GAKUSEI_NO = T13.T13_GAKUSEI_NO(+) "
	tSql = tSql & "			group by T13.T13_GAKUSEI_NO) T13MAX "

	'*** �ٓ� ***
	'����l���ɕ����N���̃f�[�^������ꍇ�ł��A���ʂɂ͌��݂̊w�N�̂��݂̂̂��o���B
	'�����Ώۂ̈ٓ����̂͂ǂ̊w�N�ł��������ɂ���A�������Ȃ���΂Ȃ�Ȃ����߁A
	'���炩���ߊY������w���ԍ��̂ݒ��o���Ă����A���̒����炳��ɕʂ̏����Œ��o����悤�ɂ���B
	wBuf = Request.Form("IDOU")
	If wBuf <> "%%%" Then
		tSql = tSql & " , (select T13_GAKUSEI_NO "
		tSql = tSql & " from T11_GAKUSEKI T11, T13_GAKU_NEN T13 "
		tSql = tSql & " where T11.T11_GAKUSEI_NO = T13.T13_GAKUSEI_NO(+) "
		tSql = tSql & " and ( "
		tSql = tSql & " T13.T13_IDOU_KBN_1 = '" & wBuf & "' or "
		tSql = tSql & " T13.T13_IDOU_KBN_2 = '" & wBuf & "' or "
		tSql = tSql & " T13.T13_IDOU_KBN_3 = '" & wBuf & "' or "
		tSql = tSql & " T13.T13_IDOU_KBN_4 = '" & wBuf & "' or "
		tSql = tSql & " T13.T13_IDOU_KBN_5 = '" & wBuf & "' "
		tSql = tSql & " ) "
		tSql = tSql & " group by T13.T13_GAKUSEI_NO) T13GAKU "
	End If

	tSql = tSql & "		where "
	tSql = tSql & "			T11.T11_GAKUSEI_NO = T13.T13_GAKUSEI_NO(+) "
	tSql = tSql & "			and KBN_NYUGAKU.M01_NENDO(+) = " & CurrentYear
	tSql = tSql & "			and KBN_NYUGAKU.M01_DAIBUNRUI_CD(+) = 3 "
	tSql = tSql & "			and KBN_NYUGAKU.M01_SYOBUNRUI_CD(+) = T11.T11_NYUGAKU_KBN "
	tSql = tSql & "			and M02.M02_NENDO(+) = " & CurrentYear
	tSql = tSql & "			and M02.M02_GAKKA_CD(+) = T13.T13_GAKKA_CD "
	tSql = tSql & "			and M17.M17_NENDO(+) = " & CurrentYear
	tSql = tSql & "			and M17.M17_BUKATUDO_CD(+) = T13.T13_CLUB_1 "
	tSql = tSql & "			and KBN_RYOSEI.M01_NENDO(+) = " & CurrentYear
	tSql = tSql & "			and KBN_RYOSEI.M01_DAIBUNRUI_CD(+) = 23 "
	tSql = tSql & "			and KBN_RYOSEI.M01_SYOBUNRUI_CD(+) = T13.T13_RYOSEI_KBN "
	tSql = tSql & "			and T13.T13_GAKUSEI_NO = T13MAX.T13_GAKUSEI_NO "
	tSql = tSql & "			and T13.T13_NENDO = T13MAX.GMAX "

	'�ٓ�
	If wBuf <> "%%%" Then
		tSql = tSql & " and T13.T13_GAKUSEI_NO = T13GAKU.T13_GAKUSEI_NO "
	End If

	'�w�N
	wBuf = Request.Form("GAKUNEN")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_GAKUNEN = " & wBuf & " "
	End If

	'�w��
	wBuf = Request.Form("GAKKA")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_GAKKA_CD = '" & wBuf & "' "
	End If

	'�N���X
	wBuf = Request.Form("CLASS")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_CLASS = '" & wBuf & "' "
	End If

	'�w����
	wBuf = Request.Form("MEISYO")
	If wBuf <> "" Then
		tSql = tSql & " and ( "
		tSql = tSql & " T11.T11_SIMEI like '" & wBuf & "%' "
		tSql = tSql & " or T11.T11_SIMEI_KD like '" & wBuf & "%') "
	End If

	'�w���ԍ�
	wBuf = Request.Form("GAKUSEI_BANGOU")
	If wBuf <> "" Then
		tSql = tSql & " and "
		tSql = tSql & " T11.T11_GAKUSEI_NO = '" & wBuf & "' "
	End If

	'����
	wBuf = Request.Form("SEX")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T11.T11_SEIBETU = '" & wBuf & "' "
	End If

	'�w�Дԍ�
	wBuf = Request.Form("GAKUSEKI_BANGOU")
	If wBuf <> "" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_GAKUSEKI_NO = '" & wBuf & "' "
	End If

	'���w�N���u
	wBuf = Request.Form("BUKATUDO_T")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T11.T11_TYU_CLUB = '" & wBuf & "' "
	End If

	'���݃N���u
	wBuf = Request.Form("BUKATUDO_G")
	If wBuf <> "%%%" Then
		tSql = tSql & " and ( "
		tSql = tSql & " T13.T13_CLUB_1 = '" & wBuf & "' "
		tSql = tSql & " or T13.T13_CLUB_2 = '" & wBuf & "') "
	End If

	'��
	wBuf = Request.Form("RYO")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_RYOSEI_KBN = '" & wBuf & "' "
	End If



	' ***** RS *****
	if gf_RSOpenOLE(RS, OrCon, tSql) = false then
		Response.Write Err.Description & "<br>"
	end if


	' ***** �y�[�W�֘A�X�^�[�g *****
%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
<%
	MaxCount = RS.RecordCount

	If NowPage > 0 Then
		Response.Write "<td nowrap width=100><a href='javascript:sbmt(" & NowPage - 1 & ")'>��PREV</td>"
	Else
		Response.Write "<td nowrap width=100>&nbsp</td>"
	End If

	EofFlg = False
	If RS.EOF Then
		Response.Write "<td align='center' width='100%' nowrap>�Y�����鐶�k����������܂���ł����B</td>"
		EofFlg = True
	End If

	If Not EofFlg Then
		RS.MoveNextn NowPage * MaxDisp

		If (NowPage + 1) * MaxDisp < MaxCount Then
			Response.Write "<td align='center' width='100%' nowrap>"
			Response.Write NowPage * MaxDisp + 1 & "�l �` " & (NowPage + 1) * MaxDisp & "�l �^ " & MaxCount & "�l��"
		Else
			Response.Write "<td align='center' width='100%' nowrap>"
			Response.Write NowPage * MaxDisp + 1 & "�l �` " & MaxCount & "�l�^" & MaxCount & "�l��"
		End If

		Response.Write "<br>PAGE: "
		mn = MaxCount
		n = 1
		Do
			If mn <= 0 Then
				Exit Do
			End If
			Response.Write "<a href=javascript:sbmt(" & n - 1 & ")>" & n & "</a> "
			mn = mn - MaxDisp
			n = n + 1
		Loop
		Response.Write "</td>"
	End If


	' ***** �m�d�w�s *****
	If (NowPage + 1) * MaxDisp < MaxCount Then
		Response.Write "<td nowrap width=100 align=right><a href=javascript:sbmt(" & NowPage + 1 & ")>NEXT��</a></td>"
	Else
		Response.Write "<td nowrap width=100>&nbsp</td>"
	End If

	Response.Write "</tr></table>"

	' ***** �y�[�W�֘A�I��� *****


	If Not EofFlg Then
%>

	<table border="1" cellpadding="1" bordercolor="#886688" width="100%">
		<tr>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�w�Дԍ�</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�w�N</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�w��</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�N���X</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�o��<br>�ԍ�1</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�o��<br>�ԍ�2</font></td>
		<td align="center" bgcolor="#886688" height="16" width="140"><font color="white">���@�@��</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">����</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">���w�敪</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�w���ԍ�</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�ٓ�</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">�N���u</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">��</font></td>
		</tr>

<%
	I = 0
	Do Until RS.EOF Or I = MaxDisp
%>
		<tr>
		<% '�w�Дԍ� %>
		<td align="center" height="16"><%= RS(4) %>&nbsp</td>
		<% '�w�@�@�N %>
		<td align="center" height="16"><%= RS(6) %>&nbsp</td>
		<% '�w�@�@�� %>
		<td align="center" height="16"><%= RS(15) %>&nbsp</td>
		<% '�N �� �X %>
		<td align="center" height="16"><%= RS("T13_CLASS") %>&nbsp</td>
		<% '�o�Ȕԍ�1%>
		<td align="center" height="16"><%= RS(8) %>&nbsp</td>
		<% '�o�Ȕԍ�2%>
		<td align="center" height="16"><%= RS(9) %>&nbsp</td>
		<% '���@�@�� %>
		<td align="center" height="16"><a href="../syosai/default.asp?id=<%= RS(0) %>" target="_blank"><%= RS(1) %></a>&nbsp</td>
		<% '���@�@�� %>
		<td align="center" height="16">
			<% If RS(12) = 1 Then %>
				�j
			<% Else %>
				��
			<% End If %>
		</td>
		<% '���w�敪 %>
		<td align="center" height="16"><%= RS(14) %>&nbsp</td>
		<% '�w���ԍ� %>
		<td align="center" height="16"><%= RS(0) %>&nbsp</td>
		<% '�ف@�@�� %>
		<td align="center" height="16"><%= RS(17) %>&nbsp</td>
		<% '�N �� �u %>
		<td align="center" height="16"><%= RS(16) %>&nbsp</td>
		<% '�@ ���@  %>
		<td align="center" height="16">
			<% If RS(13) = 1 Then %>
				��
			<% Else %>
				&nbsp
			<% End If %>
		</td>
		</tr>

<%
		RS.MoveNext
		I = I + 1
	Loop
%>
	</table>
<%


	' ***** �y�[�W�֘A�X�^�[�g *****
%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
<%
	MaxCount = RS.RecordCount

	If NowPage > 0 Then
		Response.Write "<td nowrap width=100><a href='javascript:sbmt(" & NowPage - 1 & ")'>��PREV</td>"
	Else
		Response.Write "<td nowrap width=100>&nbsp</td>"
	End If

	EofFlg = False
	If RS.EOF Then
		EofFlg = True
	End If

	If Not EofFlg Then
		RS.MoveNextn NowPage * MaxDisp

		If (NowPage + 1) * MaxDisp < MaxCount Then
			Response.Write "<td align='center' width='100%' nowrap>"
			Response.Write NowPage * MaxDisp + 1 & "�l �` " & (NowPage + 1) * MaxDisp & "�l �^ " & MaxCount & "�l��"
		Else
			Response.Write "<td align='center' width='100%' nowrap>"
			Response.Write NowPage * MaxDisp + 1 & "�l �` " & MaxCount & "�l�^" & MaxCount & "�l��"
		End If

		Response.Write "<br>PAGE: "
		mn = MaxCount
		n = 1
		Do
			If mn <= 0 Then
				Exit Do
			End If
			Response.Write "<a href=javascript:sbmt(" & n - 1 & ")>" & n & "</a> "
			mn = mn - MaxDisp
			n = n + 1
		Loop
		Response.Write "</td>"
	End If


	' ***** �m�d�w�s *****
	If (NowPage + 1) * MaxDisp < MaxCount Then
		Response.Write "<td nowrap width=100 align=right><a href=javascript:sbmt(" & NowPage + 1 & ")>NEXT��</a></td>"
	Else
		Response.Write "<td nowrap width=100>&nbsp</td>"
	End If

	Response.Write "</tr></table>"

	' ***** �y�[�W�֘A�I��� *****


	End If

	Call gf_RSCloseOLE(RS)
	Call gf_ConnCloseOLE(OrSes, OrCon)
End Sub



Sub ShowPage()
'***********************************************************
'�@�@�@�\�F�@�g�s�l�k���o��
'�@�ԁ@�l�F�@����
'�@���@���F�@����
'�@�ځ@�ׁF�@����
'�@���@�l�F�@����
'***********************************************************

'*** HTML���J�n ***
%>
	<html>
	<head>
	<title>�w�Ѓf�[�^����</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
	<style type="text/css">
		<!--
		body,table,tr,td,th {
			font-size:12px;color:#886688;
		}
		input,select{font-size:12px;}
		h3 {font-size:15px;color:#886688;}
		hr { border-style:solid;  border-color:#0066cc; }
		a:link { color:#cc8866; text-decoration:none; }
		a:visited { color:#cc8866; text-decoration:none; }
		a:active { color:#888866; text-decoration:none; }
		a:hover { color:#888866; text-decoration:underline; }
		b { font-weight: bold }
		//-->
	</style>
	<script language="javascript">
		<!--
		function sbmt(np) {
			document.forms[0].np.value = np;
			document.forms[0].submit();
		}
		//-->
	</script>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">

<%
	' *** ���[�h���� ***
	If Request.Form("mode") = "" Then
		Call FirstHTML()
	Else
		Call DispHTML()
	End If
%>

	<input type="hidden" name="mode" value="<%= Request("mode") %>">
	<input type="hidden" name="GAKUNEN" value="<%= Request("GAKUNEN") %>">
	<input type="hidden" name="GAKKA" value="<%= Request("GAKKA") %>">
	<input type="hidden" name="CLASS" value="<%= Request("CLASS") %>">
	<input type="hidden" name="MEISYO" value="<%= Request("MEISYO") %>">
	<input type="hidden" name="GAKUSEKI_BANGOU" value="<%= Request("GAKUSEKI_BANGOU") %>">
	<input type="hidden" name="SEX" value="<%= Request("SEX") %>">
	<input type="hidden" name="GAKUSEI_BANGOU" value="<%= Request("GAKUSEI_BANGOU") %>">
	<input type="hidden" name="IDOU" value="<%= Request("IDOU") %>">
	<input type="hidden" name="BUKATUDO_T" value="<%= Request("BUKATUDO_T") %>">
	<input type="hidden" name="BUKATUDO_G" value="<%= Request("BUKATUDO_G") %>">
	<input type="hidden" name="RYO" value="<%= Request("RYO") %>">
	<input type="hidden" name="np" value="">

<%
'					mode			:���샂�[�h
'										�󔒁F�@�����\��
'										DISP:	�������ʕ\��
'					GAKUNEN			:�w�N
'					GAKKA			:�w��
'					CLASS			:�N���X
'					MEISYO			:����
'					GAKUSEKI_BANGOU	:�w�Дԍ�
'					SEX				:����
'					GAKUSEI_BANGOU	:�w���ԍ�
'					IDOU			:�ٓ�
'					BUKATUDO_T		:���w�N���u
'					BUKATUDO_G		:���݃N���u
'					RYO				:��
%>
	</form>
	</body>
	</html>

<%
'*** HTML���I�� ***
End Sub


Sub FirstHTML()
'***********************************************************
'�@�@�@�\�F�@���񎞂̕\��
'�@�ԁ@�l�F�@����
'�@���@���F�@����
'�@�ځ@�ׁF�@����
'�@���@�l�F�@����
'***********************************************************
%>
	<br><br>
	<center>
	�y �㕔�̌����t�H�[������������͂��Ă������� �z
	</center>
<%
End Sub
%>

<%
Sub DispHTML()
'***********************************************************
'�@�@�@�\�F�@�������s���̕\��
'�@�ԁ@�l�F�@����
'�@���@���F�@����
'�@�ځ@�ׁF�@����
'�@���@�l�F�@����
'***********************************************************
%>

	<div align="center"><br>�y �� �� �� �� �z<br><br></div>

	<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
	<tr>
		<td width="60">&nbsp</td>
		<td valign="top">

				<%
				' *** �������ʕ\�� ***
				Call SearchDisp()
				%>

		</td>
	</tr>
	</table>

<%
End Sub
%>
