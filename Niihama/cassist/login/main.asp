<%@ Language=VBScript %>

<%
'***********************************************************
'
'　システム名　：　教務事務システム
'　処　理　名　：　データ検索
'　プログラムID：　
'　機　　　　能：　学籍データの検索結果表示
'
'-----------------------------------------------------------
'
'　引　　　　数：　
'					mode			:動作モード
'										空白：　初期表示
'										DISP:	検索結果表示
'					GAKUNEN			:学年
'					GAKKA			:学科
'					CLASS			:クラス
'					MEISYO			:名称
'					GAKUSEKI_BANGOU	:学籍番号
'					SEX				:性別
'					GAKUSEI_BANGOU	:学生番号
'					IDOU			:異動
'					BUKATUDO_T		:中学クラブ
'					BUKATUDO_G		:現在クラブ
'					RYO				:寮
'
'　変　　　　数：　
'　引　　　　渡：　
'　説　　　　明：　
'
'-----------------------------------------------------------
'
'　作　　　　成：　2001/03/19　家入　一真
'
'***********************************************************
%>



<% '*** ASP共通モジュール宣言 *** %>

<!-- #include file="../common/common.asp" -->

<%
'*** グローバル変数 ***

	Dim CurrentYear, MaxDisp

	CurrentYear = Year(Date())
	'！！！！注意：テストで２０００年を代入しています
	CurrentYear = "2000"
	MaxDisp = 20

	If Request("np") = "" Then
		NowPage = 0
	Else
		NowPage = Request("np")
	End If

'*** メイン処理 ***

	'メインルーチン実行
	Call Main()

'*** ＥＮＤ ***

%>



<%

Sub Main()
'***********************************************************
'　機　能：　本ＡＳＰのメインルーチン
'　返　値：　無し
'　引　数：　無し
'　詳　細：　無し
'　備　考：　無し
'***********************************************************

	'On Error Resume Next
	'Err.Clear


	Call ShowPage()

	'On Error Goto 0
	'Err.Clear

End Sub


'*** 関数定義 ***


Sub SetBlank()
'***********************************************************
'　機　能：　全項目を空白に初期化
'　返　値：　無し
'　引　数：　無し
'　詳　細：　無し
'　備　考：　無し
'***********************************************************

End Sub


Sub SearchDisp()
'***********************************************************
'　機　能：　検索結果の表示
'　返　値：　無し
'　引　数：　無し
'　詳　細：　無し
'　備　考：　無し
'***********************************************************

	Dim I
	Dim Conn, RS
	Dim OrCon, OrSes

	' ***** CONN *****
	if gf_ConnOpenOLE(OrSes, OrCon) = false then
		Response.Write Err.Description & "<br>"
	end if
	
	' ***** SQL生成 *****
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

	'*** 異動 ***
	'同一人物に複数年分のデータがある場合でも、結果には現在の学年のもののみを出す。
	'検索対象の異動自体はどの学年でおこったにせよ、そうしなければならないため、
	'あらかじめ該当する学生番号のみ抽出しておき、その中からさらに別の条件で抽出するようにする。
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

	'異動
	If wBuf <> "%%%" Then
		tSql = tSql & " and T13.T13_GAKUSEI_NO = T13GAKU.T13_GAKUSEI_NO "
	End If

	'学年
	wBuf = Request.Form("GAKUNEN")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_GAKUNEN = " & wBuf & " "
	End If

	'学科
	wBuf = Request.Form("GAKKA")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_GAKKA_CD = '" & wBuf & "' "
	End If

	'クラス
	wBuf = Request.Form("CLASS")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_CLASS = '" & wBuf & "' "
	End If

	'学生名
	wBuf = Request.Form("MEISYO")
	If wBuf <> "" Then
		tSql = tSql & " and ( "
		tSql = tSql & " T11.T11_SIMEI like '" & wBuf & "%' "
		tSql = tSql & " or T11.T11_SIMEI_KD like '" & wBuf & "%') "
	End If

	'学生番号
	wBuf = Request.Form("GAKUSEI_BANGOU")
	If wBuf <> "" Then
		tSql = tSql & " and "
		tSql = tSql & " T11.T11_GAKUSEI_NO = '" & wBuf & "' "
	End If

	'性別
	wBuf = Request.Form("SEX")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T11.T11_SEIBETU = '" & wBuf & "' "
	End If

	'学籍番号
	wBuf = Request.Form("GAKUSEKI_BANGOU")
	If wBuf <> "" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_GAKUSEKI_NO = '" & wBuf & "' "
	End If

	'中学クラブ
	wBuf = Request.Form("BUKATUDO_T")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T11.T11_TYU_CLUB = '" & wBuf & "' "
	End If

	'現在クラブ
	wBuf = Request.Form("BUKATUDO_G")
	If wBuf <> "%%%" Then
		tSql = tSql & " and ( "
		tSql = tSql & " T13.T13_CLUB_1 = '" & wBuf & "' "
		tSql = tSql & " or T13.T13_CLUB_2 = '" & wBuf & "') "
	End If

	'寮
	wBuf = Request.Form("RYO")
	If wBuf <> "%%%" Then
		tSql = tSql & " and "
		tSql = tSql & " T13.T13_RYOSEI_KBN = '" & wBuf & "' "
	End If



	' ***** RS *****
	if gf_RSOpenOLE(RS, OrCon, tSql) = false then
		Response.Write Err.Description & "<br>"
	end if


	' ***** ページ関連スタート *****
%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
<%
	MaxCount = RS.RecordCount

	If NowPage > 0 Then
		Response.Write "<td nowrap width=100><a href='javascript:sbmt(" & NowPage - 1 & ")'>←PREV</td>"
	Else
		Response.Write "<td nowrap width=100>&nbsp</td>"
	End If

	EofFlg = False
	If RS.EOF Then
		Response.Write "<td align='center' width='100%' nowrap>該当する生徒が見当たりませんでした。</td>"
		EofFlg = True
	End If

	If Not EofFlg Then
		RS.MoveNextn NowPage * MaxDisp

		If (NowPage + 1) * MaxDisp < MaxCount Then
			Response.Write "<td align='center' width='100%' nowrap>"
			Response.Write NowPage * MaxDisp + 1 & "人 ～ " & (NowPage + 1) * MaxDisp & "人 ／ " & MaxCount & "人中"
		Else
			Response.Write "<td align='center' width='100%' nowrap>"
			Response.Write NowPage * MaxDisp + 1 & "人 ～ " & MaxCount & "人／" & MaxCount & "人中"
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


	' ***** ＮＥＸＴ *****
	If (NowPage + 1) * MaxDisp < MaxCount Then
		Response.Write "<td nowrap width=100 align=right><a href=javascript:sbmt(" & NowPage + 1 & ")>NEXT→</a></td>"
	Else
		Response.Write "<td nowrap width=100>&nbsp</td>"
	End If

	Response.Write "</tr></table>"

	' ***** ページ関連終わり *****


	If Not EofFlg Then
%>

	<table border="1" cellpadding="1" bordercolor="#886688" width="100%">
		<tr>
		<td align="center" bgcolor="#886688" height="16"><font color="white">学籍番号</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">学年</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">学科</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">クラス</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">出席<br>番号1</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">出席<br>番号2</font></td>
		<td align="center" bgcolor="#886688" height="16" width="140"><font color="white">氏　　名</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">性別</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">入学区分</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">学生番号</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">異動</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">クラブ</font></td>
		<td align="center" bgcolor="#886688" height="16"><font color="white">寮</font></td>
		</tr>

<%
	I = 0
	Do Until RS.EOF Or I = MaxDisp
%>
		<tr>
		<% '学籍番号 %>
		<td align="center" height="16"><%= RS(4) %>&nbsp</td>
		<% '学　　年 %>
		<td align="center" height="16"><%= RS(6) %>&nbsp</td>
		<% '学　　科 %>
		<td align="center" height="16"><%= RS(15) %>&nbsp</td>
		<% 'ク ラ ス %>
		<td align="center" height="16"><%= RS("T13_CLASS") %>&nbsp</td>
		<% '出席番号1%>
		<td align="center" height="16"><%= RS(8) %>&nbsp</td>
		<% '出席番号2%>
		<td align="center" height="16"><%= RS(9) %>&nbsp</td>
		<% '氏　　名 %>
		<td align="center" height="16"><a href="../syosai/default.asp?id=<%= RS(0) %>" target="_blank"><%= RS(1) %></a>&nbsp</td>
		<% '性　　別 %>
		<td align="center" height="16">
			<% If RS(12) = 1 Then %>
				男
			<% Else %>
				女
			<% End If %>
		</td>
		<% '入学区分 %>
		<td align="center" height="16"><%= RS(14) %>&nbsp</td>
		<% '学生番号 %>
		<td align="center" height="16"><%= RS(0) %>&nbsp</td>
		<% '異　　動 %>
		<td align="center" height="16"><%= RS(17) %>&nbsp</td>
		<% 'ク ラ ブ %>
		<td align="center" height="16"><%= RS(16) %>&nbsp</td>
		<% '　 寮　  %>
		<td align="center" height="16">
			<% If RS(13) = 1 Then %>
				●
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


	' ***** ページ関連スタート *****
%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
<%
	MaxCount = RS.RecordCount

	If NowPage > 0 Then
		Response.Write "<td nowrap width=100><a href='javascript:sbmt(" & NowPage - 1 & ")'>←PREV</td>"
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
			Response.Write NowPage * MaxDisp + 1 & "人 ～ " & (NowPage + 1) * MaxDisp & "人 ／ " & MaxCount & "人中"
		Else
			Response.Write "<td align='center' width='100%' nowrap>"
			Response.Write NowPage * MaxDisp + 1 & "人 ～ " & MaxCount & "人／" & MaxCount & "人中"
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


	' ***** ＮＥＸＴ *****
	If (NowPage + 1) * MaxDisp < MaxCount Then
		Response.Write "<td nowrap width=100 align=right><a href=javascript:sbmt(" & NowPage + 1 & ")>NEXT→</a></td>"
	Else
		Response.Write "<td nowrap width=100>&nbsp</td>"
	End If

	Response.Write "</tr></table>"

	' ***** ページ関連終わり *****


	End If

	Call gf_RSCloseOLE(RS)
	Call gf_ConnCloseOLE(OrSes, OrCon)
End Sub



Sub ShowPage()
'***********************************************************
'　機　能：　ＨＴＭＬを出力
'　返　値：　無し
'　引　数：　無し
'　詳　細：　無し
'　備　考：　無し
'***********************************************************

'*** HTML部開始 ***
%>
	<html>
	<head>
	<title>学籍データ検索</title>
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
	' *** モード分岐 ***
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
'					mode			:動作モード
'										空白：　初期表示
'										DISP:	検索結果表示
'					GAKUNEN			:学年
'					GAKKA			:学科
'					CLASS			:クラス
'					MEISYO			:名称
'					GAKUSEKI_BANGOU	:学籍番号
'					SEX				:性別
'					GAKUSEI_BANGOU	:学生番号
'					IDOU			:異動
'					BUKATUDO_T		:中学クラブ
'					BUKATUDO_G		:現在クラブ
'					RYO				:寮
%>
	</form>
	</body>
	</html>

<%
'*** HTML部終了 ***
End Sub


Sub FirstHTML()
'***********************************************************
'　機　能：　初回時の表示
'　返　値：　無し
'　引　数：　無し
'　詳　細：　無し
'　備　考：　無し
'***********************************************************
%>
	<br><br>
	<center>
	【 上部の検索フォームより条件を入力してください 】
	</center>
<%
End Sub
%>

<%
Sub DispHTML()
'***********************************************************
'　機　能：　検索実行時の表示
'　返　値：　無し
'　引　数：　無し
'　詳　細：　無し
'　備　考：　無し
'***********************************************************
%>

	<div align="center"><br>【 検 索 結 果 】<br><br></div>

	<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
	<tr>
		<td width="60">&nbsp</td>
		<td valign="top">

				<%
				' *** 検索結果表示 ***
				Call SearchDisp()
				%>

		</td>
	</tr>
	</table>

<%
End Sub
%>
