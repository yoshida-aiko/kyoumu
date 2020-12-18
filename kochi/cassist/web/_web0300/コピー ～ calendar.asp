<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 特別教室予約
' ﾌﾟﾛｸﾞﾗﾑID : web/web0300/calender.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:	SESSION("KYOKAN_CD"):教官CD
'            	SESSION("NENDO")	:年度
'				TUKI				:月
'				cboKyositu			:教室CD
'
' 引      渡:	hidDay     :日にち
'				hidYear    :年
'				hidMonth   :月
'				hidKyositu :教室CD
' 説      明:
'           ■初期表示
'               選択された月のカレンダーを表示
'           ■日付クリック時
'               下のフレームに選択された日付の教室情報を表示
'-------------------------------------------------------------------------
' 作      成: 2001/08/06 伊藤公子
' 変      更:
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iSyoriNen          '教官ｺｰﾄﾞ
    Public m_iKyokanCd          '年度
    Public m_iTuki              '//月
	Public m_iKyosituCd			'//教室CD
	Public m_sKyosituName		'//教室名称
	Public m_SDate
	Public m_EDate
	Public m_sDay    	'//日
	Public m_sKyokanNm

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

Sub Main()
'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    Dim w_iRet              '// 戻り値

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="日毎出欠入力"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '//値の初期化
        Call s_ClearParam()

        '//変数セット
        Call s_SetParam()

'//デバッグ
'Call s_DebugPrint

		'//教室名取得
		w_iRet = f_GetKyousituName()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//日付を取得
        w_iRet = f_GetDate()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '// ページを表示
        Call showPage()
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen  = ""
    m_iKyokanCd  = ""
    m_iTuki      = ""
	m_sKyokanNm  = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen  = Session("NENDO")    
    'm_iKyokanCd  = Session("KYOKAN_CD")
    m_iKyokanCd  = Request("SKyokanCd1")

    m_iTuki      = Request("TUKI")
	m_iKyosituCd = Request("cboKyositu")
	m_sDay       = Request("hidDay")
	m_sKyokanNm  =Request("SKyokanNm1")

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen  = " & m_iSyoriNen  & "<br>"
    response.write "m_iKyokanCd  = " & m_iKyokanCd  & "<br>"
    response.write "m_iTuki      = " & m_iTuki      & "<br>"
    response.write "m_iKyosituCd = " & m_iKyosituCd & "<br>"
    response.write "m_sDay       = " & m_sDay       & "<br>"
    response.write "m_sKyokanNm  = " & m_sKyokanNm  & "<br>"

End Sub

'********************************************************************************
'*  [機能]  日付データを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDate()

	Dim w_iRet
	Dim w_sSQL
	Dim rs
	Dim w_sSDate
	Dim w_sEDate

	On Error Resume Next
	Err.Clear

	f_GetDate = 1

	Do

		'//行事明細テーブルよりカレンダーデータを取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T32.T32_HIDUKE, "
		w_sSQL = w_sSQL & vbCrLf & "  T32.T32_YOUBI_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M T32"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T32.T32_NENDO=" & cInt(m_iSyoriNen)
        w_sSQL = w_sSQL & vbCrLf & "  AND SUBSTR(T32.T32_HIDUKE,6,2)='" & gf_fmtZero(m_iTuki,2) & "'"
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "  T32.T32_HIDUKE,T32.T32_YOUBI_CD"

'response.write w_sSQL & "<BR>"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDate = 99
			Exit Do
		End If

		If rs.EOF = False then
			rs.MoveFirst
			m_SDate = rs("T32_HIDUKE")
			rs.MoveLast
			m_EDate = rs("T32_HIDUKE")
		End If

		'//正常終了
		f_GetDate = 0
		Exit Do

	Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  教室名取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKyousituName()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetKyousituName = 1

    Do
		'//教室名取得
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU.M06_KYOSITUMEI"
		w_sSql = w_sSql & vbCrLf & " FROM M06_KYOSITU"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU.M06_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M06_KYOSITU.M06_KYOSITU_CD=" & m_iKyosituCd

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKyousituName = 99
            Exit Do
        End If

		If rs.EOF = False Then
			m_sKyosituName = rs("M06_KYOSITUMEI")
		End If

        '//正常終了
        f_GetKyousituName = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  該当日付に予定が入っているかどうかにより、TDのCOLORを返す
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_KyosituYoteiInfo(p_iJigenCnt,p_sDay,p_sColor)

    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_iYoyakCnt

    On Error Resume Next
    Err.Clear

    f_KyosituYoteiInfo = 1
	p_sColor = ""

    Do
		'//教室名取得
		'w_sSql = ""
		'w_sSql = w_sSql & vbCrLf & " SELECT "
		'w_sSql = w_sSql & vbCrLf & "  COUNT(*) AS CNT"
		'w_sSql = w_sSql & vbCrLf & " FROM "
		'w_sSql = w_sSql & vbCrLf & "  T58_KYOSITU_YOYAKU T58"
		'w_sSql = w_sSql & vbCrLf & " WHERE "
		'w_sSql = w_sSql & vbCrLf & "  T58.T58_NENDO=" & m_iSyoriNen
		'w_sSql = w_sSql & vbCrLf & "  AND T58.T58_HIDUKE='" & gf_YYYY_MM_DD(p_sDay,"/") & "' "
		'w_sSql = w_sSql & vbCrLf & "  AND T58.T58_KYOSITU=" & m_iKyosituCd

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T58.T58_YOUBI_CD,"
		w_sSql = w_sSql & vbCrLf & "  T58.T58_JIGEN"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T58_KYOSITU_YOYAKU T58"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T58.T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND T58.T58_HIDUKE='" & gf_YYYY_MM_DD(p_sDay,"/") & "' "
		w_sSql = w_sSql & vbCrLf & "  AND T58.T58_KYOSITU=" & m_iKyosituCd
		w_sSql = w_sSql & vbCrLf & "  GROUP BY"
		w_sSql = w_sSql & vbCrLf & "  T58.T58_YOUBI_CD,"
		w_sSql = w_sSql & vbCrLf & "  T58.T58_JIGEN"

'response.write w_sSQL & "<br>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_KyosituYoteiInfo = 99
            Exit Do
        End If

		'w_iYoyakCnt = rs("CNT")
		w_iYoyakCnt = 0
		Do Until rs.EOF

			w_iYoyakCnt = w_iYoyakCnt +1	
		
			RS.MoveNext
		Loop
		'w_iYoyakCnt = rs.RecordCount		

'response.write w_iYoyakCnt

		If cint(w_iYoyakCnt) = 0 Then
			'//予約が入っていない
			p_sColor = ""
		Else

			If cint(w_iYoyakCnt) >= cint(p_iJigenCnt) Then
				'//全ての時限に予約が入っている
				p_sColor = "FILLFULL"
			Else
				'//一部の時限に予約が入っている
				p_sColor = "FILLPART"
			End If
		End If

        '//正常終了
        f_KyosituYoteiInfo = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  カレンダーを作成
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_MakeCalendar()

	Dim myTable()
	Dim w_sTdColor

	On Error Resume Next
	Err.Clear

	f_MakeCalendar = 1

	Do

		myDate    = m_SDate

		myWeekTbl = split("日,月,火,水,木,金,土",",")
		myMonthTbl= split("31,28,31,30,31,30,31,31,30,31,30,31",",")
		myYear = year(m_SDate)

		'//閏年判定
		If (myYear Mod 4=0 And myYear Mod 100<>0 ) or (myYear Mod 400=0) Then
			myMonthTbl(1) = 29	'//2月
		End If

		myMonth = m_iTuki
		myWeek = cint(Weekday(myYear & "/" & myMonth & "/01"))-1	'//ついたちの曜日を取得
		myTblLine = int(((myWeek+myMonthTbl(myMonth-1))/7)+0.9)		'//行数を取得

		ReDim myTable(7*myTblLine)

		'//初期化
		For i=0 To 7*myTblLine-1
			myTable(i)="　"
		Next

		'//日付を格納
		For i = 0 to myMonthTbl(myMonth-1)-1
			myTable(i+myWeek)=i+1
		Next

		'// ***********************
		'// **  カレンダーの表示
		'// ***********************

		'//時限の最大値を取得(教室の予約状況を取得するため)
		w_iRet = f_GetJigen(w_iJigenCnt)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'response.write("<table border='1' class='hyo' width='98%'  >")
		response.write("<table border='1' class='hyo' width='80%'  >")
		response.write("<tr>")

		'=============
		'ヘッダ部
		'=============
		'//曜日を表示
		For i = 0 to 6
		   response.write("<th align='center' class='header'>")
		   response.write(myWeekTbl(i))
		   response.write("</th>")
		Next
		response.write("</tr>")

		'=============
		'明細部
		'=============
		'//日にちを表示
		For i = 0 to myTblLine-1

			'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
			Call gs_cellPtn(w_Class)
		   response.write("<tr>")

		   For j=0 To 7-1
		    myDat = myTable(j+(i*7))

			'//TD色をｾｯﾄ
			w_sTdClassColor=w_Class

			If myDat <> "　" Then
				'=============================================================
				'//該当日付に予定が入っているかどうかにより、TDのCOLORを返す
				w_sDay = myYear & "/" & myMonth & "/" & myDat

				w_sColor = ""
				w_iRet = f_KyosituYoteiInfo(w_iJigenCnt,w_sDay,w_sColor)
				If w_iRet <> 0 Then
					Exit Do
				End If

				If w_sColor <> "" Then
					w_sTdClassColor=w_sColor
				End If
				'=============================================================
			End If

		    response.write("<td align='center' class='" + w_sTdClassColor + "' > ")
			If myDat="　" Then
			    response.write("　")
			Else
				If m_sDay<> "" Then

					If cint(m_sDay) = myDat Then
						'response.write("<span class='select_date'>" & myDat & "</span>")
						response.write("<b>" & myDat & "</b>")
					Else
						response.write("<A HREF='javascript:f_ListClick(" & myDat & ")'>" & myDat & "</A>")
					End If
				Else
					response.write("<A HREF='javascript:f_ListClick(" & myDat & ")'>" & myDat & "</A>")
				End If

			End If

		    response.write("</td>")
			Next
		   response.write("</tr>")
		Next

		response.write("</table>")

		'//正常終了
		f_MakeCalendar = 0
		Exit Do

	Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  時限情報の最大値と最小値を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetJigen(p_iCnt)

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen = 1

    Do

		'w_sSql = ""
		'w_sSql = w_sSql & vbCrLf & " SELECT "
		'w_sSql = w_sSql & vbCrLf & "  MAX(T20_JIKANWARI.T20_JIGEN) AS MAX "
		'w_sSql = w_sSql & vbCrLf & " FROM T20_JIKANWARI"
		'w_sSql = w_sSql & vbCrLf & " WHERE "
		'w_sSql = w_sSql & vbCrLf & "      T20_JIKANWARI.T20_NENDO=" & m_iSyoriNen
		'w_sSql = w_sSql & vbCrLf & "  AND T20_JIKANWARI.T20_GAKKI_KBN=" & Session("GAKKI")

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  MAX(m07_JIGEN.m07_JIKAN) AS MAX "
		w_sSql = w_sSql & vbCrLf & " FROM m07_JIGEN"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "      m07_JIGEN.M07_NENDO=" & m_iSyoriNen
		'w_sSql = w_sSql & vbCrLf & "  AND T20_JIKANWARI.T20_GAKKI_KBN=" & Session("GAKKI")

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetJigen = 99
            Exit Do
        End If

		If ISNULL(rs("MAX")) Then
			p_iCnt = 0
		Else
			p_iCnt = rs("MAX")
		End If

        '//正常終了
        f_GetJigen = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
%>
    <html>
    <head>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>特別教室予約</title>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

    }

    //************************************************************
    //  [機能] カレンダー日付クリック時
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_ListClick(p_date){

		var wArg

		//カレンダーページを再表示
		wArg = ""
		wArg = wArg + "?TUKI=<%=m_iTuki%>"
		wArg = wArg + "&cboKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidDay="+p_date
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=Server.URLEncode(m_iKyokanCd)%>"

		parent.middle.location.href="./calendar.asp"+wArg
		//parent.middle.location.href="./calendar.asp?TUKI=<%=m_iTuki%>&cboKyositu=<%=m_iKyosituCd%>&hidDay="+p_date

		//リストページを再表示
		wArg = ""
		wArg = wArg + "?hidDay="+p_date
		wArg = wArg + "&hidYear=<%=year(m_SDate)%>"
		wArg = wArg + "&hidMonth=<%=month(m_SDate)%>"
		wArg = wArg + "&hidKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=Server.URLEncode(m_iKyokanCd)%>"

		parent.bottom.location.href="./web0300_lst.asp"+wArg

	}

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <br>
    <form name="frm" method="post">
<%
'//デバッグ
'Call s_DebugPrint()
%>

	<center>

    <table class="hyo" border="1" width="80%">

        <tr>
            <th class="header" width="20%" align="center" nowrap><font size="2">利用者</font></th>
            <td class="detail" width="80%" align="left"   nowrap colspan="2"><font size="2"><%=m_sKyokanNm%></font></td>
        </tr>
        <tr>
            <th class="header" width="20%" align="center" nowrap><font size="2">教室</font></th>
            <td class="detail" width="40%" align="left"   nowrap><font size="2"><%=m_sKyosituName%></font></td>
            <td class="detail" width="40%" align="center" nowrap><font size="2"><%=year(m_SDate)%>年　<%=Month(m_SDate)%>月</font></td>
        </tr>
    </table>
	<br>

	<%
	'//カレンダー表示
	Call f_MakeCalendar()
	%>
	<table width="80%" border=0><tr>
	<td align="right" nowrap>
		<span class="msg" ><font size="2">※赤表示：全て埋まっています。<br>黄表示：一部埋まっています。</font></span>
	</td>
	</tr></table>

	<!--値渡用-->
	<input type="hidden" name="TUKI"       value="<%=m_iTuki%>">
	<input type="hidden" name="cboKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">
	<input type="hidden" name="SKyokanCd1" value="<%=m_iKyokanCd%>">

	<input type="hidden" name="hidDay"     value="">
	<input type="hidden" name="hidYear"    value="<%=year(m_SDate) %>">
	<input type="hidden" name="hidMonth"   value="<%=month(m_SDate)%>">
	<input type="hidden" name="hidKyositu" value="<%=m_iKyosituCd%>">

	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>
