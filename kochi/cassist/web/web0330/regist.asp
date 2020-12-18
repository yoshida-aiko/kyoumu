<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 連絡掲示板
' ﾌﾟﾛｸﾞﾗﾑID : web/web0330/regist.asp
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
' 変      更: 2001/09/01 伊藤公子 教官以外も利用できるように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iMax           ':最大ページ
    Public m_iDsp                       '// 一覧表示行数
    Public m_stxtMode           'モード
    Public m_rs
    Public m_sSQL
    Public m_sNendo         '年度
    Public m_sKyokanCd      '教官ｺｰﾄﾞ
    Public m_stxtNo         '処理番号
    Public m_sKenmei        '件名
    Public m_sNaiyou        '内容
    Public m_sKaisibi       '開始日
    Public m_sSyuryoubi     '完了日
    Public m_sListCd

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
    w_sMsgTitle="連絡掲示板"
    w_sMsg=""
    w_sRetURL=C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_stxtMode = request("txtMode")

    m_sKenmei   = request("txtKenmei")
    m_sNaiyou   = request("txtNaiyou")
    m_sKaisibi  = request("txtKaisibi")
    m_sSyuryoubi= request("txtSyuryoubi")
    m_sNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_stxtNo    = request("txtNo")
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

        '//モードによって情報の取得
        Select Case m_stxtMode
            Case "NEW"
        '// ページを表示

			m_sKaisibi = gf_YYYY_MM_DD(Trim(date()),"/")

            Call showPage()
            Exit Do
            Case "UPD"
            w_iRet = f_GetData()
            If w_iRet <> 0 Then
                'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
                m_bErrFlg = True
                Exit Do
            End If
            Case "NEW2"
            w_iRet = f_NUgetData()
            If w_iRet <> 0 Then
                'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
                m_bErrFlg = True
                Exit Do
            End If
            Case "UPD2"
            w_iRet = f_NUgetData()
            If w_iRet <> 0 Then
                'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
                m_bErrFlg = True
                Exit Do
            End If
        End Select

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

Function f_GetData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_rs

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
        '//変数の値を取得
        m_sSQL = ""
        m_sSQL = m_sSQL & "SELECT DISTINCT"
        m_sSQL = m_sSQL & " T46_KENMEI,T46_NAIYO,T46_KAISI,T46_SYURYO "
        m_sSQL = m_sSQL & "FROM "
        m_sSQL = m_sSQL & " T46_RENRAK "
        m_sSQL = m_sSQL & "WHERE "
        m_sSQL = m_sSQL & " T46_NO = '" & cInt(m_stxtNo) & "'"

        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

        '//取得した値を変数に代入
        m_sKenmei   = w_rs("T46_KENMEI")
        m_sNaiyou   = w_rs("T46_NAIYO")
        m_sKaisibi  = w_rs("T46_KAISI")
        m_sSyuryoubi= w_rs("T46_SYURYO")

        If m_stxtMode = "UPD" Then

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
			m_sSQL = m_sSQL & vbCrLf & "  ,T46_RENRAK "
			m_sSQL = m_sSQL & vbCrLf & " WHERE "
			m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_KYOKAN_CD = M04_KYOKAN.M04_KYOKAN_CD(+) "
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO = M04_KYOKAN.M04_NENDO(+)"
			m_sSQL = m_sSQL & vbCrLf & "  AND T46_RENRAK.T46_NO = " & cInt(m_stxtNo)
			m_sSQL = m_sSQL & vbCrLf & "  AND T46_RENRAK.T46_KYOKAN_CD = M10_USER.M10_USER_ID(+) "
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_sNendo
			m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_KYOKAKEIRETU_KBN,M04_GAKKA_CD,M10_USER_NAME"

'response.write m_sSQL & "<BR>"

            Set m_rs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
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

Function f_NUgetData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_NUgetData = 1

    m_sListCd = request("chk")

    Do
        '//送付先のデータ取得

		'//USERID取得＆成型
		w_sUser = ""
		w_sAryUser = split(Replace(Trim(m_sListCd)," ",""),",")
		w_iCnt = UBound(w_sAryUser)

		For i = 0 To w_iCnt
			If w_sUser = "" Then
				w_sUser = "'" & w_sAryUser(i) & "'"
			Else
				w_sUser = w_sUser & ",'" & w_sAryUser(i) & "'"
			End If
		Next

        '//送付先のデータ取得
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
        m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_USER_ID IN (" & w_sUser & ") "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_sNendo
		m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_GAKKA_CD,M04_KYOKAKEIRETU_KBN,M10_USER_NAME"

'response.write m_sSQL & "<BR>"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

        f_NUgetData = 0

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
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_sNendo
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
    On Error Resume Next
    Err.Clear

%>
<HTML>
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>連絡掲示板</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  送信先修正ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Syusei(){

        var iRet;
        // 入力値のﾁｪｯｸ
        iRet = f_CheckData();
        if( iRet != 0 ){
            return;
        }
        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "default.asp";
<%If m_stxtMode = "NEW" or m_stxtMode = "NEW2" Then%>
        document.frm.txtMode.value = "NEW";
<%ElseIf m_stxtMode = "UPD" or m_stxtMode = "UPD2" Then%>
        document.frm.txtMode.value = "UPD";
<%End If%>
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Touroku(){

        var iRet;
        // 入力値のﾁｪｯｸ
        iRet = f_CheckData();
        if( iRet != 0 ){
            return;
        }
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "web0330_edt.asp";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_itiran(){

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "default.asp";
        document.frm.txtMode.value = "";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  入力値のﾁｪｯｸ
    //  [引数]  なし
    //  [戻値]  0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
    //  [説明]  入力値のNULLﾁｪｯｸ、英数字ﾁｪｯｸ、桁数ﾁｪｯｸを行う
    //************************************************************
    function f_CheckData() {
    
        // ■■■NULLﾁｪｯｸ■■■
        // ■件名
        if( f_Trim(document.frm.Kenmei.value) == "" ){
            window.alert("件名が入力されていません");
            document.frm.Kenmei.focus();
            return 1;
        }

        // ■内容
        if( f_Trim(document.frm.Naiyou.value) == "" ){
            window.alert("内容に何も入力されていません");
            document.frm.Naiyou.focus();
            return 1;
        }

        // ■開始日
        if( f_Trim(document.frm.Kaisibi.value) == "" ){
            window.alert("開始日が入力されていません");
            document.frm.Kaisibi.focus();
            return 1;
        }

        // ■完了日
        if( f_Trim(document.frm.Syuryoubi.value) == "" ){
            document.frm.Syuryoubi.value = document.frm.Kaisibi.value;
        }

        // ■■■日付ﾁｪｯｸ■■■
        // ■ 開始日
        if( IsDate(document.frm.Kaisibi.value) != 0 ){
            window.alert("期間の開始日の日付が不正です");
            document.frm.Kaisibi.focus();
            return 1;
        }

	    // ■ 完了日
        if( IsDate(document.frm.Syuryoubi.value) != 0 ){
            window.alert("期間の完了日の日付が不正です");
            document.frm.Syuryoubi.focus();
            return 1;
        }

        // ■■■内容の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.Naiyou.value) > "254" ){
            window.alert("内容の欄は全角127文字以内で入力してください");
            document.frm.Naiyou.focus();
            return 1;
        }

        // ■■■件名の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.Kenmei.value) > "40" ){
            window.alert("件名の欄は全角20文字以内で入力してください");
            document.frm.Kenmei.focus();
            return 1;
        }

        // ■■■期間の取得のﾁｪｯｸ■■■
        if ( f_Trim(document.frm.Syuryoubi.value) != "" ){
	        if( DateParse(document.frm.Kaisibi.value,document.frm.Syuryoubi.value) < 0){
	            window.alert("開始日と終了日を正しく入力してください");
	            document.frm.Kaisibi.focus();
	            return 1;
	        }
        }
        return 0;
    }

    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" method="post">

<br>

<% If m_stxtMode = "NEW"or m_stxtMode = "NEW2" Then 
    call gs_title("連絡掲示板","新　規")
   Else 
    call gs_title("連絡掲示板","修　正")
   End If%>

<br>
<font>登　録　内　容</font>
<br>
<br>
<%If m_stxtMode = "NEW" Then %>
<div align="center"><span class=CAUTION>※ 入力事項を入力し、｢送付先選択｣ボタンをクリックしてください。<br>
</span></div>
<%ElseIf m_stxtMode = "UPD" Then%> 
<div align="center"><span class=CAUTION>※ 修正したい項目を修正し、送付先を変更したい場合は｢送付先選択｣ボタンをクリックしてください。
</span></div>
<%Else%> 
<div align="center"><span class=CAUTION>※ 送付先を変更したい場合は｢送付先選択｣ボタンをクリックしてください。<br>
										※ 登録してよければ登録ボタンをクリックしてください。
</span></div>
<%End If%>
</TD>
</TR>
</TABLE>

<br>
<table width="510" border=1 CLASS="hyo">
    <TR>
        <TH CLASS="header" width="60">件名</TH>
        <TD CLASS="detail"><input type="text" size="57" name="Kenmei" value="<%=m_sKenmei%>" maxlength=40><br>
        <font size=2>（全角20文字以内）</font></TD>
    </TR>
    <TR>
        <TH CLASS="header" width="60">内容</TH>
        <TD CLASS="detail"><textarea rows=6 cols=40 class=text name="Naiyou"><%=m_sNaiyou%></textarea><br>
        <font size=2>（全角127文字以内）</font></TD>
    </TR>
    <TR>
        <TH CLASS="header" width="60">期間</TH>
        <TD CLASS="detail"><input type="text" size="23" name="Kaisibi" value="<%=m_sKaisibi%>" maxlength=10>
        <input type="button" class="button" onclick="fcalender('Kaisibi')" value="選択">
        　～　<input type="text" size="23" name="Syuryoubi" value="<%=m_sSyuryoubi%>" maxlength=10>
        　<input type="button" class="button" onclick="fcalender('Syuryoubi')" value="選択"><br>
        <font size=2>（入力例:<%=Date()%>）</font></TD>
    </TR>
<%
    If m_stxtMode <> "NEW" Then
%>
    <tr>
    <td colspan=2 align=right bgcolor=#9999BD>
	<input type="button" value="登　録" class=button onclick="javascript:f_Touroku()">
	<input type="button" value="キャンセル" class=button onclick="javascript:f_itiran()">
	<input class=button type=button value="送付先選択" onclick="javascript:f_Syusei()"></td>
    </tr>
    <TR>
        <TH CLASS="header" valign="top">送付先</TD>
        <TD CLASS="detail" colspan=2>
        <table border=1 class=hyo width=100% height=100%>
<%
    m_rs.MoveFirst
    Do Until m_rs.EOF
%>
		    <TR>

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
				Call gf_GetKubunName(C_KYOKAN,m_rs("M04_KYOKAN_KBN"),m_sNendo,w_sKyokanKbnName)

				'//教科系列区分名称を取得
				Call gf_GetKubunName(C_KYOKA_KEIRETU,m_rs("M04_KYOKAKEIRETU_KBN"),m_sNendo,w_sKeiretuKbnName)

				w_sGakkaKigo = f_GetGakkaKigoName(m_rs("M04_GAKKA_CD"))
			Else
				'//教官以外の場合USER区分名称を表示
				Call gf_GetKubunName(C_USER,m_rs("M10_USER_KBN"),m_sNendo,w_sKyokanKbnName)
				w_sKeiretuKbnName = "―"
				w_sGakkaKigo = "―"
			End If

			'========================================================

            Call gs_cellPtn(w_cell)
			%>

	        <td class="CELL2"><%=w_sKyokanKbnName%><BR></td>
	        <td class="CELL2"><%=w_sKeiretuKbnName%>
				<input type="hidden" name="KCD" value='<%=m_rs("M10_USER_ID")%>'><BR></td>
	        <td class="CELL2"><%=w_sGakkaKigo%><BR></td>
	        <td class="CELL2"><%=m_rs("M10_USER_NAME")%><BR></td>
		    </TR>
<%
    m_rs.MoveNext
    Loop
%>
        </table>
		</td>
<%
    Else
%>
    <tr>
    <td colspan=6 align=right bgcolor=#9999BD>
	<input type="button" value="キャンセル" class=button onclick="javascript:f_itiran()">
	<input class=button type=button value="送付先選択" onclick="javascript:f_Syusei()"></td>
<%
    End If
%>
    </tr>

</TABLE>
    <INPUT TYPE=HIDDEN  NAME=txtNo value="<%=m_stxtNo%>">
    <INPUT TYPE=HIDDEN  NAME=txtMode value="<%=m_stxtMode%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo   VALUE="<%=m_sNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtListCd      value="<%=m_sListCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    VALUE="<%=m_sKyokanCd%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
