<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 特別教室予約
' ﾌﾟﾛｸﾞﾗﾑID : web/web0300/web0300_lst.asp
' 機      能: 教室情報を表示
'-------------------------------------------------------------------------
' 引      数:   NENDO           '//年度
'               KYOKAN_CD       '//教官CD
'				txtMode			:処理モード
'				hidJigen		:時限
'				hidDay			:日にち
'				hidYear			:年
'				hidMonth		:月
'				hidKyositu		:教室CD
'				hidKyosituName	:教室名称
'
' 引      渡:	txtMode			:処理モード
'				hidJigen		:時限
'				hidDay			:日にち
'				hidYear			:年
'				hidMonth		:月
'				hidKyositu		:教室CD
'				hidKyosituName	:教室名称
' 説      明:
'           ■初期表示
'               解除選択されたデータ一覧を表示
'-------------------------------------------------------------------------
' 作      成: 2001/08/08 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	Public m_iSyoriNen          '//年度
	Public m_iKyokanCd          '//教官ｺｰﾄﾞ

	Public m_sYear   			'//年
	Public m_sMonth				'//月
	Public m_sDay    			'//日
	Public m_iKyosituCd			'//教室CD
	Public m_iKaijyoCnt			'//解除チェックボックスカウント
	Public m_sMode				'//処理モード
	Public m_iJigen				'//時限
	Public m_sMokuteki			'//目的
	Public m_sBiko				'//備考
	Public m_sKyosituName		'//教室名称

    'ﾚｺｰﾄﾞセット
    Public m_Rs					'//ﾚｺｰﾄﾞｾｯﾄ

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
    w_sMsgTitle="特別教室予約"
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
'Call s_DebugPrint()

		Select Case m_sMode

			Case "DISP"

				'//表示用データ取得
				w_iRet = f_GetDispData()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//画面を表示
				Call showPage()

			Case "DELETE"
				'//データDelete
				w_iRet = f_DeleteData()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//削除正常終了時
				Call showWhitePage()

			Case Else

		End Select

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(m_Rs)

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
    m_sYear      = ""
    m_sMonth     = ""
    m_sDay       = ""
	m_iKyosituCd = ""
	m_sMode      = ""
	m_iJigen     = ""
	m_sMokuteki  = ""
	m_sBiko      = ""

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
	''m_iKyokanCd  = Request("YoyakKyokanCd")
    m_iKyokanCd  = Request("SKyokanCd1")

    m_sYear      = Request("hidYear")
    m_sMonth     = Request("hidMonth")
    m_sDay       = Request("hidDay")
	m_iKyosituCd = Request("hidKyositu")
	m_sMode      = Request("txtMode")
	m_iJigen     = Request("hidJigen")
	m_sKyosituName = Request("hidKyosituName")

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_sMode      = " & m_sMode      & "<br>"
    response.write "m_iSyoriNen  = " & m_iSyoriNen  & "<br>"
    response.write "m_iKyokanCd  = " & m_iKyokanCd  & "<br>"
    response.write "m_sYear      = " & m_sYear      & "<br>"
    response.write "m_sMonth     = " & m_sMonth     & "<br>"
    response.write "m_sDay       = " & m_sDay       & "<br>"
    response.write "m_iKyosituCd = " & m_iKyosituCd & "<br>"
    response.write "m_iJigen     = " & m_iJigen     & "<br>"
    response.write "m_sMokuteki  = " & m_sMokuteki  & "<br>"
    response.write "m_sBiko      = " & m_sBiko      & "<br>"
    response.write "m_sKyosituName= " & m_sKyosituName & "<br>"

End Sub

'********************************************************************************
'*  [機能]  表示データを取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDispData()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDispData = 1

    Do
		'//日付を作成
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'//教室予約データ取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & " T58_HIDUKE, "
		w_sSql = w_sSql & vbCrLf & " T58_JIGEN, "
		w_sSql = w_sSql & vbCrLf & " T58_MOKUTEKI, "
		w_sSql = w_sSql & vbCrLf & " T58_BIKO"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & " T58_KYOSITU_YOYAKU"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & " T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND T58_HIDUKE='" & w_sDate & "'"
		w_sSql = w_sSql & vbCrLf & " AND T58_JIGEN IN (" & replace(Request("chkKaijyo")," ","") & ")"
		w_sSql = w_sSql & vbCrLf & " AND T58_KYOSITU=" & m_iKyosituCd
		'w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_iKyokanCd & "'"
		w_sSql = w_sSql & vbCrLf & " ORDER BY T58_JIGEN"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetDispData = 99
            Exit Do
        End If

        '//正常終了
        f_GetDispData = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  データUPDATE
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_DeleteData()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_DeleteData = 1

    Do
		'//日付を作成
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'//DELETE
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " DELETE T58_KYOSITU_YOYAKU "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & " T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND T58_HIDUKE='" & w_sDate & "'"
		w_sSql = w_sSql & vbCrLf & " AND T58_JIGEN IN (" & replace(Request("chkKaijyo")," ","") & ")"
		w_sSql = w_sSql & vbCrLf & " AND T58_KYOSITU=" & m_iKyosituCd
		'w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_iKyokanCd & "'"

'response.write w_sSQL & "<br>"

		iRet = gf_ExecuteSQL(w_sSQL)
		If iRet <> 0 Then
			'削除失敗
			msMsg = Err.description
			f_DeleteData = 99
			Exit Do
		End If

		'//正常終了
		f_DeleteData = 0
		Exit Do
	Loop

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

    <!--#include file="../../Common/jsCommon.htm"-->
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
    //  [機能]  キャンセルボタンクリック
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_Cancel() {

		document.frm.action="web0300_lst.asp";
		document.frm.target="bottom";
		document.frm.submit();
    }

    //************************************************************
    //  [機能]  削除ボタンクリック
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_Delete(){

        if (!confirm("予約を解除します。よろしいですか？")) {
           return ;
        }

		document.frm.txtMode.value="DELETE";
		document.frm.action="web0300_del.asp";
		document.frm.target="bottom";
		document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

<%
'//デバッグ
'Call s_DebugPrint()
%>

	<center>
	<!--<form action="yoyaku.asp">-->
	<img src="img/sp.gif" height="3">

		<table border="1" class="hyo" width="700">
			<tr>
			<th CLASS="header" width="100" nowrap>日付</th>
			<th CLASS="header" width="100" nowrap>教室</th>
			<th CLASS="header" width="60" nowrap>時限</th>
			<th CLASS="header" width="200">使用目的</th>
			<th CLASS="header" width="200">備考</th>
			</tr>
			<%Do Until m_Rs.EOF%>
				<tr>
				<td class="detail" width="100" align="center" nowrap><%=gf_fmtWareki(m_Rs("T58_HIDUKE"))%><BR></td>
				<td class="detail" width="100" align="center" nowrap><%=m_sKyosituName%><BR></td>
				<td class="detail" width="60"  align="center" nowrap><%=m_Rs("T58_JIGEN")%>時限</td>
				<td class="detail" width="200" ><%=m_Rs("T58_MOKUTEKI")%><BR></td>
				<td class="detail" width="200" ><%=m_Rs("T58_BIKO")%><BR></td>
				</tr>

				<%m_Rs.MoveNext%>
			<%Loop%>

		</table>

		<br>

		<table width="250">
			<tr>
			<td align="center" colspan="2"><font size="2">以上の予約を解除します。</font></td>
			</tr>
			<tr>
			<td align="center"><input class="button" type="button" value="　解　除　" onclick="javascript:f_Delete()"></td>
			<td align="center"><input class="button" type="button" value="キャンセル" onclick="javascript:f_Cancel()"></td>
			</tr>
		</table>

	<!--値渡し用-->
	<input type="hidden" name="txtMode"    value="">
	<input type="hidden" name="chkKaijyo"  value="<%=Request("chkKaijyo")%>">
	<input type="hidden" name="SKyokanCd1"    value="<%=m_iKyokanCd%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">

	<input type="hidden" name="hidDay"     value="<%=m_sDay%>">
	<input type="hidden" name="hidYear"    value="<%=m_sYear %>">
	<input type="hidden" name="hidMonth"   value="<%=m_sMonth%>">
	<input type="hidden" name="hidKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="hidKyosituName" value="<%=m_sKyosituName%>">

	</form>
	</center>
	</body>
	</html>

<%
End Sub

'********************************************************************************
'*  [機能]  空白ページ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showWhitePage()
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
		alert("予約を解除しました")

		var wArg
		//カレンダーページを再表示
		wArg = ""
		wArg = wArg + "?TUKI=<%=m_sMonth%>"
		wArg = wArg + "&cboKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidDay=<%=m_sDay%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=m_iKyokanCd%>"

		//カレンダーページを再表示
		parent.middle.location.href="./calendar.asp"+wArg
		//parent.middle.location.href="./calendar.asp?TUKI=<%=m_sMonth%>&cboKyositu=<%=m_iKyosituCd%>&hidDay=<%=m_sDay%>"

		//リストページを再表示
		wArg = ""
		wArg = wArg + "?hidDay=<%=m_sDay%>"
		wArg = wArg + "&hidYear=<%=m_sYear%>"
		wArg = wArg + "&hidMonth=<%=m_sMonth%>"
		wArg = wArg + "&hidKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidKyosituName=<%=Server.URLEncode(m_sKyosituName)%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=m_iKyokanCd%>"

		parent.bottom.location.href="./web0300_lst.asp"+wArg

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type="hidden" name="TUKI"       value="<%=m_sMonth%>">
	<input type="hidden" name="cboKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="SKyokanCd1" value="<%=m_iKyokanCd%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">

	<input type="hidden" name="hidDay"     value="<%=m_sDay%>">
	<input type="hidden" name="hidYear"    value="<%=m_sYear %>">
	<input type="hidden" name="hidMonth"   value="<%=m_sMonth%>">
	<input type="hidden" name="hidKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="hidKyosituName" value="<%=m_sKyosituName%>">

	</form>
	</body>
	</html>
<%
End Sub
%>