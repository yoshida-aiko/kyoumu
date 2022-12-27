<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 時間割交換連絡
' ﾌﾟﾛｸﾞﾗﾑID : web/web0310/sousin_main.asp
' 機      能: 下ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
'            モード         ＞      txtMode
'                                   新規 = NEW
'                                   更新 = UPDATE
'            内容           ＞      txtNaiyou
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 前田
' 変      更: 2001/09/03 伊藤公子 教官以外も利用できるように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_iMax          ':最大ページ
    Public  m_iDsp                      '// 一覧表示行数
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官ｺｰﾄﾞ
    Public m_stxtNo         '処理番号
    Public m_stxtMode       'モード
    Public m_sNaiyou        '内容
    Public m_sKaisibi       '開始日
    Public m_sSyuryoubi     '完了日
    Public m_sJoukin        '常勤区分
    Public m_sGakka         '学科区分
    Public m_sKkanKBN       '教官区分
    Public m_sKkeiKBN       '教科系列区分
    Public m_rs
    Public m_Srs            '更新の際の送信先の過去のデータ取得用レコード
    Dim    m_rCnt           '//レコード件数

	Public m_sUserKbn		'//USER区分
	Public m_sSimei			'//氏名

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
    w_sMsgTitle="連絡事項登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_stxtMode = request("txtMode")

    m_sNaiyou   = request("txtNaiyou")
    m_iNendo    = request("txtNendo")
    m_sKaisibi  = request("txtKaisibi")
    m_sSyuryoubi= request("txtSyuryoubi")
    m_sKyokanCd = request("txtKyokanCd")
    m_sJoukin   = request("Joukin")
    m_stxtNo    = request("txtNo")
    m_iDsp = C_PAGE_LINE

    m_sGakka   = Trim(Replace(request("Gakka"),"@@@",""))
    m_sKkanKBN = Trim(Replace(request("KkanKBN"),"@@@",""))
    m_sKkeiKBN = Trim(Replace(request("KkeiKBN"),"@@@",""))
	m_sUserKbn = Trim(Replace(request("UserKbn"),"@@@",""))
	m_sSimei   = request("txtSimei")


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

'//デバッグ
'Call s_DebugPrint


        '//データの表示
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

        Select Case m_stxtMode
            Case "NEW"
                '// ページを表示
                Call showPage()
                Exit Do
            Case "UPD"
                '// ページを表示
                Call UPD_showPage()
                Exit Do
        End Select

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
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_stxtMode  = " & m_stxtMode   & "<br>"
    response.write "m_sKenmei   = " & m_sKenmei    & "<br>"
    response.write "m_sNaiyou   = " & m_sNaiyou    & "<br>"
    response.write "m_sKaisibi  = " & m_sKaisibi   & "<br>"
    response.write "m_sSyuryoubi= " & m_sSyuryoubi & "<br>"
    response.write "m_iNendo    = " & m_iNendo     & "<br>"
    response.write "m_sKyokanCd = " & m_sKyokanCd  & "<br>"
    response.write "m_sGakka    = " & m_sGakka     & "<br>"
    response.write "m_sKkanKBN  = " & m_sKkanKBN   & "<br>"
    response.write "m_sKkeiKBN  = " & m_sKkeiKBN   & "<br>"
    response.write "m_stxtNo    = " & m_stxtNo     & "<br>"
    response.write "m_sUserKbn  = " & m_sUserKbn   & "<br>"
    response.write "m_sSimei    = " & m_sSimei     & "<br>"

End Sub

Function f_GetData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
'
        '//絞り込まれた条件で一覧の表示
        m_sSQL = ""
		m_sSQL = m_sSQL & vbCrLf & " SELECT "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_USER_ID "
		m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_KBN "
		m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_NAME "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_CD "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_GAKKA_CD "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAKEIRETU_KBN "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_KBN"
		m_sSQL = m_sSQL & vbCrLf & " FROM "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN "
		m_sSQL = m_sSQL & vbCrLf & " WHERE "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_KYOKAN_CD = M04_KYOKAN.M04_KYOKAN_CD(+) "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO = M04_KYOKAN.M04_NENDO(+)"
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_iNendo

        If m_sGakka <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_GAKKA_CD= '" & m_sGakka & "' "
        End If

        If m_sKkanKBN <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_KYOKAN_KBN=" & Cint(m_sKkanKBN)
        End If

        If m_sKkeiKBN <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_KYOKAKEIRETU_KBN=" & Cint(m_sKkeiKBN)
        End If

        If m_sUserKbn <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_USER_KBN= " & m_sUserKbn
        End If

        If m_sSimei <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_USER_NAME LIKE '%" & m_sSimei & "%'"
        End If

		m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_GAKKA_CD,M04_KYOKAKEIRETU_KBN,M10_USER_NAME"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

		'//レコード数取得
    	m_rCnt=gf_GetRsCount(m_rs)


        If m_stxtMode = "UPD" Then
            '//送信されている人のデータを取得
            m_sSQL = ""
            m_sSQL = m_sSQL & "SELECT "
            m_sSQL = m_sSQL & " T52_KYOKAN_CD "
            m_sSQL = m_sSQL & "FROM "
            m_sSQL = m_sSQL & " T52_JYUGYO_HENKO "
            m_sSQL = m_sSQL & "WHERE "
            m_sSQL = m_sSQL & " T52_NO = '" & cInt(m_stxtNo) & "'"

            Set m_Srs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(m_Srs, m_sSQL,m_iDsp)
            If w_iRet <> 0 Then
                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                m_bErrFlg = True
                Exit Do 
            End If
        End If

	    f_GetData = 0

	    Exit Do

    Loop

End Function

'********************************************************************************
'*  [機能]  学科記号を取得
'*  [引数]  なし
'*  [戻値]  gf_GetUserNm:
'*  [説明]  
'********************************************************************************
Function f_GetGakkaKigoName(p_sGakkaCd)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetGakkaKigoName = ""
	w_sName = ""

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKA_KIGO"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_GAKKA_CD='" & p_sGakkaCd & "'"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M02_GAKKA_KIGO")
        End If

        Exit Do
    Loop

	'//戻り値ｾｯﾄ
    f_GetGakkaKigoName = w_sName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

Sub S_NEW_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_Half
Dim j
j = 0
%>
<table>
    <tr>
        <td colspan=4 align=center>
            <input type="button" value=" 前 画 面 へ " class=button onclick="javascript:f_before()">
            <input type="button" value="全てチェック" class=button onclick="javascript:f_Allchk()">
            <input type="button" value=" 全てクリア " class=button onclick="javascript:f_AllClear()">
            <input type="button" value=" 選 択 完 了 " class=button onclick="javascript:f_Skanryo()">
        </td>
    </tr>
</table>
<BR>
<div align="center"><span class=CAUTION>※ 送付したい人にチェックをつけて｢選択完了｣ボタンをクリックします。<br>
										※ 全員に送りたい場合は｢全てチェック｣を、全員のチェックをはずしたい場合は｢全てクリア｣をクリックします。
</span></div>

    <%If NOT m_rCnt = "1" Then %>

<table border=0 width=100%>
<tr>

<td align="center" width=50% valign="top">

    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>選択</th>
        <th width=22% class=header>教官</th>
        <th width=4% class=header>学科</th>
        <th width=19% class=header>教科系</th>
        <th width=43% class=header>氏名</th>
    </tr>
    <%
        m_rs.MoveFirst
        w_Half = gf_Round(m_rCnt / 2 ,0)
        Do Until m_rs.EOF
            Call gs_cellPtn(w_cell)
            j = j + 1 
            If w_Half + 1 = j then
            w_cell = ""
            Call gs_cellPtn(w_cell)
    %>
    </table>
</td>
<td align="center" width=50% valign="top">
    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>選択</th>
        <th width=22% class=header>教官</th>
        <th width=4% class=header>学科</th>
        <th width=19% class=header>教科系</th>
        <th width=43% class=header>氏名</th>
    </tr>
    <% End If %>
    <tr>

			<%
			'========================================================
			'//区分名称等取得

			w_sKyokanKbnName = ""
			w_sKeiretuKbnName = ""
			w_sGakkaKigo = ""

			'//教官CDをセット
			w_sKyokanCd = m_rs("M04_KYOKAN_CD")

			'//教官の時(教官CDありの場合)
			If LenB(w_sKyokanCd) <> 0 Then
				'//教官区分名称を取得
				Call gf_GetKubunName(C_KYOKAN,m_rs("M04_KYOKAN_KBN"),m_iNendo,w_sKyokanKbnName)

				'//教科系列区分名称を取得
				Call gf_GetKubunName(C_KYOKA_KEIRETU,m_rs("M04_KYOKAKEIRETU_KBN"),m_iNendo,w_sKeiretuKbnName)

				w_sGakkaKigo = f_GetGakkaKigoName(m_rs("M04_GAKKA_CD"))
			Else

				'//教官以外の場合USER区分名称を表示
				Call gf_GetKubunName(C_USER,m_rs("M10_USER_KBN"),m_iNendo,w_sKyokanKbnName)
				w_sKeiretuKbnName = "―"
				w_sGakkaKigo = "―"

			End If
			'========================================================
			%>

	        <td class=<%=w_cell%> align="center"><input type=checkbox name=chk value="<%=m_rs("M10_USER_ID")%>"></td>
	        <td class=<%=w_cell%>><%=w_sKyokanKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=w_sGakkaKigo%><BR></td>
	        <td class=<%=w_cell%>><%=w_sKeiretuKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=m_rs("M10_USER_NAME")%><BR></td>
    </tr>
    <% m_rs.MoveNext
        Loop
    Else %>

<table border=0 width=50%>
<tr>

<td align="center" width=100% valign="top">

    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>選択</th>
        <th width=22% class=header>教官</th>
        <th width=4% class=header>学科</th>
        <th width=19% class=header>教科系</th>
        <th width=43% class=header>氏名</th>
    </tr>
    <tr>
			<%
			'========================================================
			'//区分名称等取得

			w_sKyokanKbnName = ""
			w_sKeiretuKbnName = ""
			w_sGakkaKigo = ""

			'//教官CDをセット
			w_sKyokanCd = m_rs("M04_KYOKAN_CD")

			'//教官の時(教官CDありの場合)
			If LenB(w_sKyokanCd) <> 0 Then
				'//教官区分名称を取得
				Call gf_GetKubunName(C_KYOKAN,m_rs("M04_KYOKAN_KBN"),m_iNendo,w_sKyokanKbnName)

				'//教科系列区分名称を取得
				Call gf_GetKubunName(C_KYOKA_KEIRETU,m_rs("M04_KYOKAKEIRETU_KBN"),m_iNendo,w_sKeiretuKbnName)

				w_sGakkaKigo = f_GetGakkaKigoName(m_rs("M04_GAKKA_CD"))
			Else
				'//教官以外の場合USER区分名称を表示
				Call gf_GetKubunName(C_USER,m_rs("M10_USER_KBN"),m_iNendo,w_sKyokanKbnName)
				w_sKeiretuKbnName = "―"
				w_sGakkaKigo = "―"
			End If
			'========================================================
            Call gs_cellPtn(w_cell)
			%>
	        <td class=<%=w_cell%> align="center"><input type=checkbox name=chk value="<%=m_rs("M10_USER_ID")%>"></td>
	        <td class=<%=w_cell%>><%=w_sKyokanKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=w_sGakkaKigo%><BR></td>
	        <td class=<%=w_cell%>><%=w_sKeiretuKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=m_rs("M10_USER_NAME")%><BR></td>
    </tr>
<% End If %>

</table></td></tr></table>
<table>
    <tr>
        <td colspan=4 align=center>
            <input type="button" value=" 前 画 面 へ " class=button onclick="javascript:f_before()">
            <input type="button" value="全てチェック" class=button onclick="javascript:f_Allchk()">
            <input type="button" value=" 全てクリア " class=button onclick="javascript:f_AllClear()">
            <input type="button" value=" 選 択 完 了 " class=button onclick="javascript:f_Skanryo()">
        </td>
    </tr>
</table>

<%End sub

Sub S_UPD_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_Half
Dim j
j = 0
%>
<table>
    <tr>
        <td colspan=4 align=center>
            <input type="button" value=" 前 画 面 へ " class=button onclick="javascript:f_before()">
            <input type="button" value="全てチェック" class=button onclick="javascript:f_Allchk()">
            <input type="button" value=" 全てクリア " class=button onclick="javascript:f_AllClear()">
            <input type="button" value=" 選 択 完 了 " class=button onclick="javascript:f_Skanryo()">
        </td>
    </tr>
</table>
<BR>
<div align="center"><span class=CAUTION>※ 送付したい人にチェックをつけて｢選択完了｣ボタンをクリックします。<br>
										※ 全員に送りたい場合は｢全てチェック｣を、全員のチェックをはずしたい場合は｢全てクリア｣をクリックします。
</span></div>

    <%If NOT m_rCnt = "1" Then %>

<table border=0 width=100%>
<tr>

<td align="center" width=50% valign="top">

    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>選択</th>
        <th width=22% class=header>教官</th>
        <th width=4% class=header>学科</th>
        <th width=19% class=header>教科系</th>
        <th width=43% class=header>氏名</th>
    </tr>
    <%
        m_rs.MoveFirst
        w_Half = gf_Round(m_rCnt / 2 ,0)
        Do Until m_rs.EOF
            Call gs_cellPtn(w_cell)
            j = j + 1 
            If w_Half + 1 = j then
            w_cell = ""
            Call gs_cellPtn(w_cell)
    %>
    </table>
</td>
<td align="center" width=50% valign="top">
    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>選択</th>
        <th width=22% class=header>教官</th>
        <th width=4% class=header>学科</th>
        <th width=19% class=header>教科系</th>
        <th width=43% class=header>氏名</th>
    </tr>
    <%
            End If 
            m_Srs.MoveFirst
            w_schk=""
            Do Until m_Srs.EOF
                If m_rs("M10_USER_ID") = m_Srs("T52_KYOKAN_CD") Then
                    w_schk=" checked"
                    Exit Do
                End If
            m_Srs.MoveNext
            Loop
    %>
    <tr>
			<%
			'========================================================
			'//区分名称等取得
			w_sKyokanKbnName = ""
			w_sKeiretuKbnName = ""
			w_sGakkaKigo = ""

			'//教官CDをセット
			w_sKyokanCd = m_rs("M04_KYOKAN_CD")

			'//教官の時(教官CDありの場合)
			If LenB(w_sKyokanCd) <> 0 Then
				'//教官区分名称を取得
				Call gf_GetKubunName(C_KYOKAN,m_rs("M04_KYOKAN_KBN"),m_iNendo,w_sKyokanKbnName)

				'//教科系列区分名称を取得
				Call gf_GetKubunName(C_KYOKA_KEIRETU,m_rs("M04_KYOKAKEIRETU_KBN"),m_iNendo,w_sKeiretuKbnName)

				w_sGakkaKigo = f_GetGakkaKigoName(m_rs("M04_GAKKA_CD"))
			Else
				'//教官以外の場合USER区分名称を表示
				Call gf_GetKubunName(C_USER,m_rs("M10_USER_KBN"),m_iNendo,w_sKyokanKbnName)
				w_sKeiretuKbnName = "―"
				w_sGakkaKigo = "―"
			End If
			'========================================================

			%>
	        <td class=<%=w_cell%> align="center"><input type=checkbox name=chk value="<%=m_rs("M10_USER_ID")%>" <%=w_schk%>></td>
	        <td class=<%=w_cell%>><%=w_sKyokanKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=w_sGakkaKigo%><BR></td>
	        <td class=<%=w_cell%>><%=w_sKeiretuKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=m_rs("M10_USER_NAME")%><BR></td>
    </tr>
    <% m_rs.MoveNext
        Loop
    Else %>

<table border=0 width=50%>
<tr>

<td align="center" width=100% valign="top">

    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>選択</th>
        <th width=22% class=header>教官</th>
        <th width=4% class=header>学科</th>
        <th width=19% class=header>教科系</th>
        <th width=43% class=header>氏名</th>
    </tr>
    <tr>
			<%
			'========================================================
			'//区分名称等取得
			w_sKyokanKbnName = ""
			w_sKeiretuKbnName = ""
			w_sGakkaKigo = ""

			'//教官CDをセット
			w_sKyokanCd = m_rs("M04_KYOKAN_CD")

			'//教官の時(教官CDありの場合)
			If LenB(w_sKyokanCd) <> 0 Then
				'//教官区分名称を取得
				Call gf_GetKubunName(C_KYOKAN,m_rs("M04_KYOKAN_KBN"),m_iNendo,w_sKyokanKbnName)

				'//教科系列区分名称を取得
				Call gf_GetKubunName(C_KYOKA_KEIRETU,m_rs("M04_KYOKAKEIRETU_KBN"),m_iNendo,w_sKeiretuKbnName)

				w_sGakkaKigo = f_GetGakkaKigoName(m_rs("M04_GAKKA_CD"))
			Else
				'//教官以外の場合USER区分名称を表示
				Call gf_GetKubunName(C_USER,m_rs("M10_USER_KBN"),m_iNendo,w_sKyokanKbnName)
				w_sKeiretuKbnName = "―"
				w_sGakkaKigo = "―"
			End If

			'========================================================

            Call gs_cellPtn(w_cell)
			%>
	        <td class=<%=w_cell%> align="center"><input type=checkbox name=chk value="<%=m_rs("M10_USER_ID")%>"  <%=w_schk%>></td>
	        <td class=<%=w_cell%>><%=w_sKyokanKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=w_sGakkaKigo%><BR></td>
	        <td class=<%=w_cell%>><%=w_sKeiretuKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=m_rs("M10_USER_NAME")%><BR></td>
    </tr>
<% End If %>

</table></td></tr></table>
<table>
    <tr>
        <td colspan=4 align=center>
            <input type="button" value=" 前 画 面 へ " class=button onclick="javascript:f_before()">
            <input type="button" value="全てチェック" class=button onclick="javascript:f_Allchk()">
            <input type="button" value=" 全てクリア " class=button onclick="javascript:f_AllClear()">　
            <input type="button" value=" 選 択 完 了 " class=button onclick="javascript:f_Skanryo()">　
           </td>
    </tr>
</table>

<%End sub

Sub showPage_NoData()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
    <html>
    <head>
 <link rel=stylesheet href="../../common/style.css" type=text/css>
   </head>

    <body>

    <center>
		<br><br><br>
		<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
    </center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

%>
<HTML>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>時間割交換連絡</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  前画面へボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_before(){

		// 全て表示されていない場合は動作不可
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  選択完了ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Skanryo(){

		// 全て表示されていない場合は動作不可
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        if (f_chk()==1){
        alert( "登録の対象となる送信者が選択されていません" );
        return;
        }

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.txtMode.value = "NEW2";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  全てチェックボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Allchk(){

		// 全て表示されていない場合は動作不可
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        var i;
        i = 0;

        //1件のとき
        if (document.frm.txtRcnt.value==1){
            document.frm.chk.checked == "True";
        }else{
        //それ以外の時
        do { 
            if(document.frm.chk[i].checked == false){
                document.frm.chk[i].checked = true;
            }
        i++; }  while(i<document.frm.txtRcnt.value);
        }
        return;
    }

    //************************************************************
    //  [機能]  全てクリアボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_AllClear(){

		// 全て表示されていない場合は動作不可
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        var i;
        i = 0;

        //1件のとき
        if (document.frm.txtRcnt.value==1){
            document.frm.chk.checked = false;
        }else{
        //それ以外の時
        do { 
            if(document.frm.chk[i].checked == true){
                document.frm.chk[i].checked = false;
            }
        i++; }  while(i<document.frm.txtRcnt.value);
        }
        return;
    }

    //************************************************************
    //  [機能]  リスト一覧のチェックボックスの確認
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_chk(){
        var i;
        i = 0;

        //0件のとき
        if (document.frm.txtRcnt.value<=0){
            return 1;
            }

        //1件のとき
        if (document.frm.txtRcnt.value==1){
            if (document.frm.chk.checked == false){
                return 1;
            }else{
                return 0;
                }
        }else{
        //それ以外の時
        var checkFlg
            checkFlg=false

        do { 
            
            if(document.frm.chk[i].checked == true){
                checkFlg=true
                break;
             }

        i++; }  while(i<document.frm.txtRcnt.value);
            if (checkFlg == false){
                return 1;
                }
        }
        return 0;
    }

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		parent.frames[0].document.frm.BtnCtrl.value="OK"
    }

    //************************************************************
    //  [機能]  ボタンの制御
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
	function f_BtnCtrl(){

		if(parent.frames[0].document.frm.BtnCtrl.value!="OK"){
			return 1;
		}
		return 0;
	}

    //-->
    </SCRIPT>


<body LANGUAGE="javascript" onload="return window_onload()">
<center>
<FORM NAME="frm" method="post">
<%
    If m_rs.EOF Then
        Call showPage_NoData()
    Else
        Call S_NEW_syousai()
    End If
%>

</table>
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="<%=m_stxtMode%>">
    <INPUT TYPE=HIDDEN  NAME=txtKenmei      value="<%=m_sKenmei%>">
    <INPUT TYPE=HIDDEN  NAME=txtNaiyou      value="<%=m_sNaiyou%>">
    <INPUT TYPE=HIDDEN  NAME=txtKaisibi     value="<%=m_sKaisibi%>">
    <INPUT TYPE=HIDDEN  NAME=txtSyuryoubi   value="<%=m_sSyuryoubi%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_iNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtRcnt        value="<%=m_rCnt%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub

Sub UPD_showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_Half
Dim w_schk
Dim j
j = 0

%>
<HTML>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>時間割交換連絡</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  前画面へボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_before(){
		// 全て表示されていない場合は動作不可
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  選択完了ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Skanryo(){

		// 全て表示されていない場合は動作不可
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        if (f_chk()==1){
        alert( "登録の対象となる送信者が選択されていません" );
        return;
        }

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.txtMode.value = "UPD2";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  全てチェックボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Allchk(){

		// 全て表示されていない場合は動作不可
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        var i;
        i = 0;

        //1件のとき
        if (document.frm.txtRcnt.value==1){
            document.frm.chk.checked == "True";
        }else{
        //それ以外の時
        do { 
            if(document.frm.chk[i].checked == false){
                document.frm.chk[i].checked = true;
            }
        i++; }  while(i<document.frm.txtRcnt.value);
        }
        return;
    }

    //************************************************************
    //  [機能]  全てクリアボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_AllClear(){

		// 全て表示されていない場合は動作不可
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        var i;
        i = 0;

        //1件のとき
        if (document.frm.txtRcnt.value==1){
            document.frm.chk.checked = false;
        }else{
        //それ以外の時
        do { 
            if(document.frm.chk[i].checked == true){
                document.frm.chk[i].checked = false;
            }
        i++; }  while(i<document.frm.txtRcnt.value);
        }
        return;
    }

    //************************************************************
    //  [機能]  リスト一覧のチェックボックスの確認
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_chk(){
        var i;
        i = 0;

        //0件のとき
        if (document.frm.txtRcnt.value<=0){
            return 1;
            }

        //1件のとき
        if (document.frm.txtRcnt.value==1){
            if (document.frm.chk.checked == false){
                return 1;
            }else{
                return 0;
                }
        }else{
        //それ以外の時
        var checkFlg
            checkFlg=false

        do { 
            
            if(document.frm.chk[i].checked == true){
                checkFlg=true
                break;
             }

        i++; }  while(i<document.frm.txtRcnt.value);
            if (checkFlg == false){
                return 1;
                }
        }
        return 0;
    }

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		parent.frames[0].document.frm.BtnCtrl.value="OK"
    }

    //************************************************************
    //  [機能]  ボタンの制御
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
	function f_BtnCtrl(){

		if(parent.frames[0].document.frm.BtnCtrl.value!="OK"){
			return 1;
		}
		return 0;
	}

    //-->
    </SCRIPT>
<body LANGUAGE="javascript" onload="return window_onload()">
<center>

<FORM NAME="frm" method="post">

<%
    If m_rs.EOF Then
        Call showPage_NoData()
    Else
        Call S_UPD_syousai()
    End If
%>

</table>
    <INPUT TYPE=HIDDEN  NAME=txtNo          value="<%=m_stxtNo%>">
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="<%=m_stxtMode%>">
    <INPUT TYPE=HIDDEN  NAME=txtKenmei      value="<%=m_sKenmei%>">
    <INPUT TYPE=HIDDEN  NAME=txtNaiyou      value="<%=m_sNaiyou%>">
    <INPUT TYPE=HIDDEN  NAME=txtKaisibi     value="<%=m_sKaisibi%>">
    <INPUT TYPE=HIDDEN  NAME=txtSyuryoubi   value="<%=m_sSyuryoubi%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_iNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtRcnt        value="<%=m_rCnt%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>