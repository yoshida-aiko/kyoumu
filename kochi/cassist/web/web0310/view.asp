<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 時間割交換連絡
' ﾌﾟﾛｸﾞﾗﾑID : web/web0330/view.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
'            モード         ＞      txtMode
'                                   新規 = NEW
'                                   更新 = UPDATE
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/10 前田
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iMax           ':最大ページ
    Public m_iDsp                       '// 一覧表示行数
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官ｺｰﾄﾞ
    Public m_stxtMode       'モード
    Public m_sNaiyou        '内容
    Public m_sKaisibi       '開始日
    Public m_sSyuryoubi     '完了日
    Public m_sJoukin        '常勤区分
    Public m_sGakka         '学科区分
    Public m_sKkanKBN       '教官区分
    Public m_sKkeiKBN       '教科系列区分
    Public m_stxtNo         '処理番号
    Public m_rs
    Public m_sListCd
    Dim    m_rCnt           '//レコード件数

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
    m_sKaisibi  = request("txtKaisibi")
    m_sSyuryoubi= request("txtSyuryoubi")
    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_stxtNo    = request("txtNo")
    m_sListCd = request("chk")
    m_iDsp = C_PAGE_LINE

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
        
        '//データの取得、表示
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

        Call showPage()
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

Function f_GetData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_sSQL
Dim w_Srs           '詳細用のレコードセット

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
        '//変数の値を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT DISTINCT"
        w_sSQL = w_sSQL & " T52_NAIYO,T52_KAISI,T52_SYURYO "
        w_sSQL = w_sSQL & "FROM "
        w_sSQL = w_sSQL & " T52_JYUGYO_HENKO "
        w_sSQL = w_sSQL & "WHERE "
        w_sSQL = w_sSQL & " T52_NO = '" & cInt(m_stxtNo) & "'"

        Set w_Srs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_Srs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

	    '//取得した値を変数に代入
	    m_sNaiyou   = w_Srs("T52_NAIYO")
	    m_sKaisibi  = w_Srs("T52_KAISI")
	    m_sSyuryoubi= w_Srs("T52_SYURYO")

        '//送信されている人のデータを取得
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
		m_sSQL = m_sSQL & vbCrLf & "  ,T52_JYUGYO_HENKO "
		m_sSQL = m_sSQL & vbCrLf & " WHERE "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_KYOKAN_CD = M04_KYOKAN.M04_KYOKAN_CD(+) "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO = M04_KYOKAN.M04_NENDO(+)"
		m_sSQL = m_sSQL & vbCrLf & "  AND T52_JYUGYO_HENKO.T52_NO = " & cInt(m_stxtNo)
		m_sSQL = m_sSQL & vbCrLf & "  AND T52_JYUGYO_HENKO.T52_KYOKAN_CD = M10_USER.M10_USER_ID(+) "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_iNendo
		m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_GAKKA_CD,M04_KYOKAKEIRETU_KBN,M10_USER_NAME"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
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

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_sClass

%>
<HTML>
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>時間割交換連絡</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  閉じるボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Close(){

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "default.asp";
        document.frm.submit();

    }
    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" action="post">

<br>
<% 
call gs_title("時間割交換連絡","照　会")
%>

<br>
<font>登　録　内　容</font>
<br>
</TD>
</TR>
</TABLE>
<BR>
<div align="center"><span class=CAUTION>※ 送付先の確認を行ないます。<br>
</span></div>

<br>

<table width="500" border=1 CLASS="hyo">
    <TR>
        <TH CLASS="header" width="60">内容</TD>
        <TD CLASS="detail"><%=m_sNaiyou%></TD>
    </TR>
    <TR>
        <TH CLASS="header">期間</TD>
        <TD CLASS="detail"><%=m_sKaisibi%>　～　<%=m_sSyuryoubi%></TD>
    </TR>
    <tr>
	    <td colspan=5 align="right" bgcolor=#9999BD><input class=button type="submit" value="閉じる" class=button onclick="javascript:f_Close()"></td>
    </tr>
    <TR>
        <TH CLASS="header" valign="top">送付先</TD>
        <TD CLASS="detail" >
        <table border=1 class=hyo width=100% height=100%>
    <%
        m_rs.MoveFirst
        Do Until m_rs.EOF
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

            Call gs_cellPtn(w_cell)
			%>

                <td class="CELL2" width=21%><%=w_sKyokanKbnName%><br></td>
                <td class="CELL2" width=21%><%=w_sKeiretuKbnName%><br></td>
                <td class="CELL2" width=6%><%=w_sGakkaKigo%><br></td>
                <td class="CELL2" width=40%><%=m_rs("M10_USER_NAME")%><br></td>

            </tr>
    <%  m_rs.MoveNext
        Loop%>
            </table>
        </TD>
    </TR>
    </TABLE>

    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_iNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>