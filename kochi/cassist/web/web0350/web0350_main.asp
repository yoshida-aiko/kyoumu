<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 空き時間情報検索
' ﾌﾟﾛｸﾞﾗﾑID : web/web0350/web0350_main.asp
' 機      能: 検索結果ページ	 空き時間情報検索を行う
'-------------------------------------------------------------------------
' 引      数:
' 変      数:
' 引      渡:
' 説      明:
'           
'-------------------------------------------------------------------------
' 作      成: 2001/08/17 持永
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_Rs				'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ(空き時間検索)
    Public  m_iJMax				'最大時限数
    Public  mRdiMode			'時限制限
    Public  mJigenSt			'開始時限
    Public  mJigenEd			'終了時限
    Public  m_SplitCell			'クラス指定が入った配列
    Public  m_StrAkijikan		'htmlが入ってる変数

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
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="授業出欠入力"
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
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'// 初期ページを表示
		if gf_IsNull(request("txtDay")) then
			Call showPageDef()
			Exit do
		End if

		'// 検索する
		w_iRet = f_SchAkijikan()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//最大時限数を取得
        Call gf_GetJigenMax(m_iJMax)
		if m_iJMax = "" Then
		    m_bErrFlg = True
		    Exit Do
		end if

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

Function f_SchAkijikan()
'******************************************************************
'機　　能：検索をする
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear

    f_SchAkijikan = 1

    Do

		'// 検索条件取得
		wDay     = request("txtDay")		'<--　日付
		mJigenSt   = request("txtJigenSt")		'<--　開始時限
		mJigenEd   = request("txtJigenEd")		'<--　終了時限
		mRdiMode = request("rdiMode")		'<--　時限制限
		wGakkaCD = request("txtGakka")		'<--　学科コード

		'// 前期後期を取得
		Call gf_GetGakkiInfo(w_sGakki,w_sZenki_Start,w_sKouki_Start,w_sKouki_End)
		if (w_sZenki_Start > gf_YYYY_MM_DD(wDay,"/")) AND (gf_YYYY_MM_DD(wDay,"/") < w_sKouki_Start) then
			w_sGakki = C_GAKKI_KOUKI		'<--　後期
		ElseIf (w_sKouki_Start > gf_YYYY_MM_DD(wDay,"/")) AND (gf_YYYY_MM_DD(wDay,"/") < w_sKouki_End) then
			w_sGakki = C_GAKKI_ZENKI		'<--　前期
		End if

		
		m_sNomiSQL = ""
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " M04.M04_KYOKAN_CD NOT IN "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	(SELECT "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_KYOKAN "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 FROM  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20_JIKANWARI T20  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 WHERE  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_NENDO     = " & Session("NENDO") & " AND  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_GAKKI_KBN = " & w_sGakki & " AND  "
'	If mJigenSt = mJigenEd then '開始と終了が違う場合、期間指定
'        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	    T20.T20_JIGEN     = " & mJigenSt   & " AND  "
'	Else
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	    T20.T20_JIGEN     >= " & mJigenSt   & " AND  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	    T20.T20_JIGEN     <= " & mJigenEd   & " AND  "
'	End If
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_YOUBI_CD  = " & weekday(wDay)
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	group by "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_KYOKAN ) AND "


		'// 学科コード
		if Not wGakkaCD = C_CBO_NULL then
			m_sGakkaSQL = " M04.M04_GAKKA_CD = " & wGakkaCD & " AND "
		End if

        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAN_CD,"
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_GAKKA_CD, "
        m_sSQL = m_sSQL & vbCrLf & " 	M02.M02_GAKKARYAKSYO, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAKEIRETU_KBN, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKANMEI_SEI, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKANMEI_MEI "
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & " 	M02_GAKKA M02, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04_KYOKAN M04 "
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & " 	M02.M02_NENDO     = M04.M04_NENDO     AND "
        m_sSQL = m_sSQL & vbCrLf & " 	M02.M02_GAKKA_CD  = M04.M04_GAKKA_CD  AND "
        m_sSQL = m_sSQL & vbCrLf & 		m_sNomiSQL				'<--時限指定のWHERE文
        m_sSQL = m_sSQL & vbCrLf & 		m_sGakkaSQL									'<--学科コードのWHERE文
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_NENDO     = " & Session("NENDO") & " "
        m_sSQL = m_sSQL & vbCrLf & " GROUP BY "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAN_CD, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_GAKKA_CD, "
        m_sSQL = m_sSQL & vbCrLf & " 	M02.M02_GAKKARYAKSYO, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAKEIRETU_KBN, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKANMEI_SEI, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKANMEI_MEI "
        m_sSQL = m_sSQL & vbCrLf & " ORDER BY "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_GAKKA_CD, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAKEIRETU_KBN, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAN_CD "
'response.write m_sSQL
        w_iRet = gf_GetRecordset(m_Rs,m_sSQL)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

    	f_SchAkijikan = 0
	    Exit Do

    Loop

End Function

Sub showPageDef()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    On Error Resume Next
    Err.Clear

%>
	<html>
	<head>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>

	<body>
	<BR>
	<div align="center">
		<br><br><br>
		<span class="msg">項目を選んで表示ボタンを押してください</span>
	</div>
	</body>
	</html>
<%
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    On Error Resume Next
    Err.Clear

%>
	<html>
	<head>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
	</head>

	<body>
	<BR>
	<div align="center">

	<% if m_Rs.Eof then %>
        <br><br><br>
        <span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
	<% Else %>

        <table class="hyo" border="1" width="400">
            <tr>
                <th nowrap class="header" width="64" align="center">日　付</th>
                <td nowrap class="detail" width="100" align="center"><%=request("txtDay")%></td>
                <th nowrap class="header" width="64" align="center">時　限</th>
				<% if mJigenSt = mJigenEd then %>
	                <td nowrap class="detail" width="150" align="center"><%= mJigenSt %>時限</td>
				<% Else %>
	                <td nowrap class="detail" width="150" align="center"><%= mJigenSt %>-<%= mJigenEd %>時限</td>
				<% End if %>
            </tr>
        </table>
		<BR>

		<table><tr><td>
			<span class="msg"><font size="2">※<%=gf_GetRsCount(m_Rs)%>名の方が空いています</font></span>
		</td></tr></table>


		<table >
			<tr><td valign="top">
			<table class=hyo border="1" bgcolor="#FFFFFF">
				<tr>
					<th nowrap class="header">学　科</th>
					<th nowrap class="header">教科系列</th>
					<th nowrap class="header">教　官　名</th>
				</tr>
				<% 	Call f_MainHyouji()	%>
			</table>
			</td></tr>
		</table>

	<% End if %>
	</div>
	</body>
	</html>
<%
End Sub


Sub s_jigenSuu()
'********************************************************************************
'*  [機能]  時限数を作成（ヘッダー部分）
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	'// 時限指定がなかったら
	if gf_IsNull(mJigen) then
		i = 1
		Do Until i > Cint(m_iJMax)
			%><th nowrap width="80" class="header"><%= i %>限目</th><%
			i = i + 1
		Loop
	Else
		'// 時限指定あり
		Select Case Cint(mRdiMode)
			Case 1
				'// 以前
				i = 1
				Do Until i > Cint(mJigen)
					%><th nowrap width="80" class="header"><%= i %>限目</th><%
					i = i + 1
				Loop

			Case 2
				'// のみ
				i = mJigen
				%><th nowrap width="80" class="header"><%= i %>限目</th><%
				
			Case 3
				'// 以降
				i = Cint(mJigen)
				Do Until i > m_iJMax
					%><th nowrap width="80" class="header"><%= i %>限目</th><%
					i = i + 1
				Loop

		End Select
	End if

End Sub



Function f_MainHyouji()
'********************************************************************************
'*  [機能]  空き時間を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  全部、授業が入っていたら表示しない
'********************************************************************************

	Dim w_iRsCnt,w_iCnt,i

	f_MainHyouji = 1
	
	w_iRsCnt = gf_GetRsCount(m_Rs)
	w_iCnt = INT(w_iRsCnt/2 + 0.9)
	
	Do Until m_Rs.Eof
			i = i + 1

		'// テーブルのクラス指定
        Call gs_cellPtn(m_cell)

		'// 教科系列取得
		Call gf_GetKubunName(C_KYOKA_KEIRETU,m_Rs("M04_KYOKAKEIRETU_KBN"),Session("NENDO"),wKyoukaKeiretu)

		
	%>
		<tr>
			<td nowrap class="<%=m_cell%>"><%=m_Rs("M02_GAKKARYAKSYO")%></td>
			<td nowrap class="<%=m_cell%>"><%=wKyoukaKeiretu%></td>
			<td nowrap class="<%=m_cell%>"><%=m_Rs("M04_KYOKANMEI_SEI")%>　<%=m_Rs("M04_KYOKANMEI_MEI")%></td>
		</tr>
	<% If i =  w_iCnt And w_iRsCnt <> 1 Then 
				'//ｽﾀｲﾙｼｰﾄのｸﾗｽを初期化
				m_cell = ""
			%>
					</table>
				</td>
				<td valign="top">
					<table class="hyo" border="1" >
						<!--ヘッダ-->

				<tr>
				<th nowrap class="header">学　科</th>
				<th nowrap class="header">教科系列</th>
				<th nowrap class="header">教　官　名</th>
			</tr>
	<%End If
	m_Rs.MoveNext
	Loop

	f_MainHyouji = 0
'
End Function

Function f_AkiJikan()
'********************************************************************************
'*  [機能]  空き時間を表示（）
'*  [引数]  なし
'*  [戻値]  True : False
'*  [説明]  
'********************************************************************************

	f_AkiJikan = False
	w_AkinashiFlg = 0

	if gf_IsNull(mJigen) then
		i = 0
		Do Until i => Cint(m_iJMax)
			m_StrAkijikan = m_StrAkijikan & "<td nowrap class='" & m_SplitCell(i) & "'>&nbsp;</td>"
			i = i + 1
		Loop
	Else
		Select Case Cint(mRdiMode)
			Case 1
				'// 以前
				i = 0
				Do Until i >= Cint(mJigen)
					m_StrAkijikan = m_StrAkijikan & "<td nowrap class='" & m_SplitCell(i) & "'>&nbsp;</td>"
					if m_SplitCell(i) = "AKIJIKAN" then
						w_AkinashiFlg = w_AkinashiFlg + 1
					End if
					i = i + 1
				Loop

				if w_AkinashiFlg = Cint(mJigen) then
					Exit Function
				End if

			Case 2
				'// のみ
				i = Cint(mJigen) - 1
				m_StrAkijikan = m_StrAkijikan & "<td nowrap class='" & m_SplitCell(i) & "'>&nbsp;</td>"
				
			Case 3
				'// 以降
				i = Cint(mJigen) - 1
				w_FlgMax = 0
				Do Until i >= (Cint(m_iJMax))
					m_StrAkijikan = m_StrAkijikan & "<td nowrap class='" & m_SplitCell(i) & "'>&nbsp;</td>"
					if m_SplitCell(i) = "AKIJIKAN" then
						w_AkinashiFlg = w_AkinashiFlg + 1
					End if
					i = i + 1
					w_FlgMax = w_FlgMax + 1
				Loop

				if w_AkinashiFlg = w_FlgMax then
					Exit Function
				End if

		End Select
	End if
	m_StrAkijikan = m_StrAkijikan & "</tr>"

	f_AkiJikan = True

End Function
%>
