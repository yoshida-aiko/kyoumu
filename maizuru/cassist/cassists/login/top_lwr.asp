<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: マイページのお知らせ
' ﾌﾟﾛｸﾞﾗﾑID : login/top_lwr.asp
' 機      能: 下ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/23 前田
' 変      更: 2001/08/07 根本 直美     NN対応に伴うソース変更
'           : 2001/08/10 根本 直美     NN対応に伴うソース変更
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_iMax          ':最大ページ
    Public  m_iDsp          '// 一覧表示行数
    Public  m_rs
    Dim     m_iNendo
    Dim     m_sKyokanCd
    Dim     m_sUserId
    Dim     m_sName

    'エラー系
    Public  m_bErrFlg       'ｴﾗｰﾌﾗｸﾞ
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

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

        '// 権限チェックに使用
        session("PRJ_No") = C_LEVEL_NOCHK

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '//リストの一覧データの詳細取得
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

        '// ページを表示
        Call showPage()
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_rs)
    '// 終了処理
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
'response.write "(" & session("KYOKAN_CD") & ")kyoukan <br>"
	m_sUserId = session("LOGIN_ID")
'response.write "(" & session("LOGIN_ID") & ")kyoukan <br>"
    m_iDsp = C_PAGE_LINE

End Sub

Function f_GetData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
	Dim w_user

    On Error Resume Next
    Err.Clear
    f_GetData = 1

	w_user = m_sUserId
	'ユーザが教官であれば、教官CDを代入
	'If m_sKyokanCd <> "" then w_user = m_sKyokanCd

	'T46には、ユーザーIDで登録されているので、教官コードでは該当しない
	'ユーザーIDで抽出、更新するように変更　2001/12/11 伊藤	
	w_user = m_sUserId

    Do
        '//リストの表示
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT * "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T46_RENRAK "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T46_KYOKAN_CD = '" & w_user & "' "
        m_sSQL = m_sSQL & " AND T46_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "'"
        m_sSQL = m_sSQL & " AND T46_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"
        m_sSQL = m_sSQL & " ORDER BY T46_KAKNIN"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
    m_rCnt=gf_GetRsCount(m_rs)

    f_GetData = 0

    Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Function f_Sosin()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_sUserCD
Dim w_rs

    On Error Resume Next
    Err.Clear

    Do
        If m_rs("T46_UPD_USER") <> "" Then
            w_sUserCD = m_rs("T46_UPD_USER")
        Else
            w_sUserCD = m_rs("T46_INS_USER")
        End If

        '//送信者の姓名の取得
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT "
        m_sSQL = m_sSQL & "     M10_USER_NAME "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M10_USER "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     M10_NENDO = " & m_iNendo & " AND "
        m_sSQL = m_sSQL & "     M10_USER_ID = '" & w_sUserCD & "' "

        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

        f_Sosin = w_rs("M10_USER_NAME")

    Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub S_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear

%>
    <table class="disp" width="100%" border="1">
        <colgroup width="10%">
        <colgroup width="50%">
        <colgroup width="15%">
        <colgroup width="25%">
        <tr>
            <td class="disph" align="center" nowrap>　<br></td>
            <td class="disph" align="center" nowrap>用　件</td>
            <td class="disph" align="center" nowrap>日　付</td>
            <td class="disph" align="center" nowrap>送 信 者</td>
        </tr>

        <% m_rs.Movefirst
            Do Until m_rs.EOF 
            call gs_cellPtn(w_cell)%>
        <%Call f_Sosin
            m_sName = f_Sosin()
        %>
        <TR>
            <%If cInt(m_rs("T46_KAKNIN")) = C_KAKU_SUMI Then%>
                    <TD CLASS="<%=w_cell%>" nowrap>　<br></TD>
            <%Else%>
                    <TD CLASS="<%=w_cell%>" align=center nowrap>未</TD>
            <%End If%>
            <TD CLASS="<%=w_cell%>" nowrap>・<a href="#" onclick="NewWin(<%=m_rs("T46_NO")%>,'<%=m_sName%>');"><%=m_rs("T46_KENMEI")%></a></TD>

            <%If m_rs("T46_UPD_USER") <> "" Then%>
                    <TD CLASS="<%=w_cell%>" nowrap><%=m_rs("T46_UPD_DATE")%></TD>
            <%Else%>
                    <TD CLASS="<%=w_cell%>" nowrap><%=m_rs("T46_INS_DATE")%></TD>
            <%End If%>
            <TD CLASS="<%=w_cell%>" nowrap><%=m_sName%></TD>
        </TR>
        <% m_rs.MoveNext : Loop %>
    </table>
    <br>
    <Div align="center"><span class="CAUTION">※ 用件をクリックすると送付内容を確認できます。<br>
	<div align="center"><span class=CAUTION>※ メッセージは、表示期間を過ぎると自動的に削除されます。<br>
	</span></div>


<%
End Sub

Function s_Jikanwari(p_hyoji)
'********************************************************************************
'*  [機能]  時間割変更データの有無の確認
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
	Dim w_user

	s_Jikanwari = false
	p_hyoji = ""
	

	w_user = m_sUserId
	'ユーザが教官であれば、教官CDを代入
	If m_sKyokanCd <> "" then w_user = m_sKyokanCd

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT * "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T52_JYUGYO_HENKO "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & "     T52_KYOKAN_CD = '" & w_user & "' "
    w_sSQL = w_sSQL & " AND T52_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "'"
    w_sSQL = w_sSQL & " AND T52_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"

    Set m_Rds = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(m_Rds, w_sSQL,m_iDsp)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    End If

    If m_Rds.EOF Then
        Exit Function
    End If
	
	p_hyoji = ""
'	p_hyoji = p_hyoji & "<HR>"
	p_hyoji = p_hyoji & "<CENTER>"
	p_hyoji = p_hyoji & "<a href='#' onclick=NewWinJik()> ※時間割の変更連絡を確認する場合はここをクリックして下さい。</a>  "
	p_hyoji = p_hyoji & "</CENTER>"
	p_hyoji = p_hyoji & "<BR>"
	s_Jikanwari = true

End Function

Sub showPage_NoData()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
    <center>
        連絡事項はありません。
    </center>

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
    Dim i			'作業用
    Dim w_bDataF	'表示データフラグ （連絡事項等の連絡がある場合に立てる）
    w_bDataF = false
    i=1

%>
<HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link rel="stylesheet" href="../common/style.css" type="text/css">
    <title>教務事務システム：Canpus Assist トップページ</title>

    <!--#include file="../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [機能]  申請内容表示用ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function NewWin(p_Int,p_sSEIMEI) {
        URL = "view.asp?txtNo="+p_Int+"&txtSEIMEI="+escape(p_sSEIMEI)+"";
        <%if session("browser") = "NN" Then%>
        nWin=window.open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,outerwidth=400,outerheight=300,top=0,left=0");
        <%else%>
        nWin=window.open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=400,height=300,top=0,left=0");
        <%end if%>
        nWin.focus();
        return false;   
    }

    //************************************************************
    //  [機能]  申請内容表示用ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function NewWinJik() {
        URL = "j_view.asp";
        <%if session("browser") = "NN" Then%>
	        nWin=window.open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,outerwidth=400,outerheight=300,top=0,left=0");
        <%else%>
	        nWin=window.open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=450,height=450,top=0,left=0");
        <%end if%>
        nWin.focus();
        return false;   
    }
    //-->
    </SCRIPT>
</head>
<BODY>
<center>
<!--
<hr width="80%" size="1">
-->
<br>
<font size="3">お　知　ら　せ</font>
<br><br>
<FORM NAME="frm" ACTION="post">
<input type="hidden" name="txtNo">
<input type="hidden" name="txtSEIMEI">
<table width="90%"><tr><td>
<%
	'//時間割変更連絡の表示
    If s_Jikanwari(w_hyoji) = true Then
		response.write w_hyoji
        w_bDataF = true
	End If
	'//連絡事項の表示
    If m_rs.EOF = false Then
        Call S_syousai()
        w_bDataF = true
	End If 


	'//上記の二つともない場合
	If w_bDataF = false then 
        Call showPage_NoData()
    End If

%>
</td></tr></table>
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
