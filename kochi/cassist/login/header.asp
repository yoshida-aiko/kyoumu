<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: ログイン終了時画面
' ﾌﾟﾛｸﾞﾗﾑID : login/header.asp
' 機      能: ログイン終了時のヘッダー画面
'-------------------------------------------------------------------------
' 引      数    
'               
'           
' 変      数
' 引      渡
'           
'           
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 
' 変      更: 2001/07/26    モチナガ
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

    Dim m_LoginDay      '// ﾛｸﾞｲﾝ日
    Dim m_WaNengappi    '// 和暦生年月日
    Dim m_SchoolName    '// 学校名
	Dim m_bErrFlg

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

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="ヘッダーデータ"
    w_sMsg=""
    w_sRetURL="../default.asp"
    w_sTarget="_parent"

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '//表示データ取得
        if Not f_GetViewRs() then Exit Do

        '//初期表示
        Call showPage()

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    Call gs_CloseDatabase()

End Sub


'********************************************************************************
'*  [機能]  表示データを取得
'*  [引数]  
'*          
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetViewRs()
    Dim w_sSql
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_GetViewRs = False

    '// 今の生年月日
    w_NowDay = gf_YYYY_MM_DD(date,"/")

    '// 元号取得
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "     M09_GENGOMEI, "
    w_sSQL = w_sSQL & "     M09_KAISIBI "
    w_sSQL = w_sSQL & "FROM M09_GENGO "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "     M09_KAISIBI <= '" & w_NowDay & "' "
    w_sSQL = w_sSQL & "Order By M09_KAISIBI desc"

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

    if w_Rs.Eof then
        m_bErrFlg = True
        m_sErrMsg = "元号名取得時にエラーがでました"
        Exit Function
    End if

    '// 元号名
'    w_GENGOMEI = w_Rs("M09_GENGOMEI")

    '// 年（和暦）
'    w_iDiff = DateDiff("yyyy", w_Rs("M09_KAISIBI"), w_NowDay)
'    w_iWaNen = w_iDiff + 1

    '// 学期を取得
    if Not gf_GetKubunName(C_GAKKI,Session("GAKKI"),Session("NENDO"),w_GAKKI) then Exit Function

    '// 西暦年月日
    m_WaNengappi = Session("NENDO") & "年度　" & w_GAKKI

	'// 曜日
	w_Youbi = left(WeekDayName( Weekday(w_NowDay) ),1)

    '// 日付を表示
    m_LoginDay = year(date) & "年　" & Month(date) & "月" & Day(date) & "日" & "(" & w_Youbi & ")"

    '// 学校名取得
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "     M19_NAME "
    w_sSQL = w_sSQL & "FROM M19_GAKKO "
    'w_sSQL = w_sSQL & "Where "
    'w_sSQL = w_sSQL & "     M19_NO = " & C_School_CD

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

    if w_Rs.Eof then
        m_bErrFlg = True
        m_sErrMsg = "学校名取得時にエラーがでました"
        Exit Function
    End if

    '// 学校名
    m_SchoolName = w_Rs("M19_NAME")

    f_GetViewRs = True

    call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear

    %>

    <html>
    <head>
    <title>header</title>

    <STYLE TYPE="text/css">
    <!--
        body,tr,td { font-size:11pt;  color:#ffffff;
        font-family: "ＭＳ Ｐゴシック", Osaka, "ＭＳ ゴシック", Gothic, sans-serif;
        }

        Span.gakoumei { font-size:12pt; color:#ffffff; }

        b { font-weight: bold; }
        hr { border-style:solid;  border-color:#886688; }

        /* A　アンカー 基本*/
        a:link { color:#ffffff; font-size:10pt; text-decoration:none; }
        a:visited { color:#ffffff; font-size:10pt; text-decoration:none; }
        a:active { color:#FF8364; font-size:10pt; text-decoration:none; }
        a:hover { color:#FF8364; font-size:10pt; text-decoration:underline; }
    //-->
    </style>
    </head>

<body rightmargin="0" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table border="0" cellspacing="0" cellpadding="0" width="100%" height="61">
<tr>
<td width="156" height="61" align="left" valign="top">
<img src="images/title.gif">
</td>

<td bgcolor="#56567F" height="61" width="100%" background="images/back_gla.gif">

	<table cellspacing="0" cellpadding="0" border="0" height="61" width="100%">
	<tr>
	<td height="21" width="100%" align="right" valign="middle" nowrap>
	<font color="#ffffff" size="2">
	<%= m_LoginDay %>
	</font><img src="images/sp.gif" height="1" width="20">
	</td>

	<td height="21" align="right" valign="top">
		<table height="21" width="82" cellspacing="0" cellpadding="0" border="0">
		<tr>
		<td height="21">
		<a href="../manual/default.asp" target="_blank">
		<img src="images/help.gif" border="0">
		</a>
		</td>
		<td height="21">
		<a href="../default.asp" target="_top">
		<img src="images/logout.gif" border="0">
		</a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<tr>
	<td height="40" width="100%" align="left" valign="top">
		<table cellspacing="0" cellpadding="0" border="0" height="41" width="100%">
		<tr>
		<td>
		<img src="images/sp.gif" height="1" width="20"><Span class="gakoumei"><%= m_SchoolName %></Span>
		</td>
		<td align="right">
			<table cellspacing="0" cellpadding="0" border="0">
			<tr>
			<td>
			<font color="#ffffff">ユーザー名</font>
			</td>
			<td>
			<font color="#ffffff">:</font>
			</td>
			<td>
			<font color="#ffffff"><%= Session("USER_NM") %></font>
			</td>
			<td>
			<img src="../image/sp.gif" width="15">
			</td>
			</tr>
			</table>
		</td>
		</tr>
		</table>
	</td>
	<td height="40" align="center">
	<font color="#ffffff"><%= m_WaNengappi %></font>
	</td>
	</tr>
	</table>

</td>
</tr>
</table>

</body>
</html>

<% End Sub %>