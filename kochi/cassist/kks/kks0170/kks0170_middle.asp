<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 日毎出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0170/kks0170_middle.asp
' 機      能: 下ページ 授業出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数: SESSION("NENDO")           '//処理年
'             SESSION("KYOKAN_CD")       '//教官CD
'             TUKI           '//月
'             cboDate        '//日付
' 変      数:
' 引      渡: NENDO"        '//処理年
'             KYOKAN_CD     '//教官CD
'             GAKUNEN"      '//学年
'             CLASSNO"      '//ｸﾗｽNo
'             cboDate"      '//日付
' 説      明:
'           ■初期表示
'               検索条件にかなう担任ｸﾗｽ生徒情報を表示
'           ■登録ボタンクリック時
'               入力情報を登録する
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bTannin       '//担任ﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public m_iSyoriNen      '//処理年度
    Public m_iKyokanCd      '//教官CD
    Public m_sDate          '//日付
    Public m_iGakunen       '//学年
    Public m_iClassNo       '//クラスNo
    Public m_sClassNm       '//クラス名称
    Public m_iRsCnt         '//クラスﾚｺｰﾄﾞ数
	Public m_sEndDay		'//入力できなくなる日

    'ﾚｺｰﾄﾞセット
    Public m_Rs

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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="日毎出欠入力"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_bTannin = False

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

        '//変数初期化
        Call s_ClearParam()

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '// 担任クラス情報取得
        w_iRet = f_GetClassInfo(m_bTannin)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'入力不可になる日を取得
		call gf_Get_SyuketuEnd(m_iGakunen,m_sEndDay)

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
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_sDate     = ""
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_sClassNm = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = SESSION("NENDO")
    m_iKyokanCd = SESSION("KYOKAN_CD")
    m_sDate     = trim(Request("cboDate"))

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub
    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sDate     = " & m_sDate     & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_sClassNm  = " & m_sClassNm  & "<br>"

End Sub

'********************************************************************************
'*  [機能]  担任クラス情報取得
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetClassInfo(p_bTannin)

    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear
    
    f_GetClassInfo = 1

    Do 

        '// 担任クラス情報
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_TANNIN"
        w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & m_iKyokanCd & "'"

'response.write w_sSQL & "<BR>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetClassInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            p_bTannin = True 
            m_iGakunen = rs("M05_GAKUNEN")
            m_iClassNo = rs("M05_CLASSNO")
            m_sClassNm = rs("M05_CLASSMEI")
        End If

        f_GetClassInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  担任クラス一覧取得
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetClassList()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetClassList = 1

    Do 

        '// 担任クラス情報取得
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_IDOU_NUM, "
        w_sSQL = w_sSQL & vbCrLf & "  B.T11_SIMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  B.T11_GAKUSEI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  C.T30_HIDUKE, "
        w_sSQL = w_sSQL & vbCrLf & "  C.T30_SYUKKETU_KBN,"
        '//"出席"は表示しない
        w_sSQL = w_sSQL & vbCrLf & "  DECODE(D.M01_SYOBUNRUIMEI_R,'出','',D.M01_SYOBUNRUIMEI_R) AS SYUKKETU_MEI"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN A"
        w_sSQL = w_sSQL & vbCrLf & "  ,T11_GAKUSEKI B"
        w_sSQL = w_sSQL & vbCrLf & "  ,(SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     T30_HIDUKE,"
        w_sSQL = w_sSQL & vbCrLf & "     T30_SYUKKETU_KBN,"
        w_sSQL = w_sSQL & vbCrLf & "     T30_GAKUSEKI_NO"
        w_sSQL = w_sSQL & vbCrLf & "    FROM T30_KESSEKI"
        w_sSQL = w_sSQL & vbCrLf & "    WHERE T30_HIDUKE='" & m_sDate & "'"
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_GAKUNEN=" & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_CLASS=" & m_iClassNo & ") C"
        w_sSQL = w_sSQL & vbCrLf & "  ,(SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     M01_SYOBUNRUI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "     M01_SYOBUNRUIMEI_R"
        w_sSQL = w_sSQL & vbCrLf & "    FROM M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & "    WHERE "
        w_sSQL = w_sSQL & vbCrLf & "          M01_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "      AND M01_DAIBUNRUI_CD=" & C_KESSEKI & ") D"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        'w_sSQL = w_sSQL & vbCrLf & "      A.T13_NENDO - A.T13_GAKUNEN + 1 = B.T11_NYUNENDO(+) "
        w_sSQL = w_sSQL & vbCrLf & "      A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO "
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_GAKUSEKI_NO = C.T30_GAKUSEKI_NO(+)"
        w_sSQL = w_sSQL & vbCrLf & "  AND C.T30_SYUKKETU_KBN = D.M01_SYOBUNRUI_CD(+)"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_GAKUNEN=" & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_CLASS=" & m_iClassNo
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T13_GAKUSEKI_NO"

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetClassList = 99
            Exit Do
        End If

        '//ﾚｺｰﾄﾞカウントを取得
        m_iRsCnt = gf_GetRsCount(m_Rs)

        f_GetClassList = 0
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

    On Error Resume Next
    Err.Clear

%>
    <html>
    <head>
    <title>日毎出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

		//スクロール同期制御
		parent.init();

    }

    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Cancel(){
        //初期ページを表示
        //parent.document.location.href="default2.asp"
        document.frm.target = "<%=C_MAIN_FRAME%>";
        document.frm.action = "./default.asp"
        document.frm.submit();
        return;


    }

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Touroku(){

		parent.frames["main"].f_Touroku();
    }


    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">
<%
'//デバッグ
'Call s_DebugPrint()
%>
    <center>
    <form name="frm" method="post">
    <%call gs_title("日毎出欠入力","一　覧")%>
    <%Do
        '//担任クラスがない場合
        If m_bTannin = False Then
        %>
        <br><br>
        <span class="msg">受持クラスがありません。</span>
        <%
            Exit Do
        End If
		%>

        <table>
		<tr><td>
	        <table class="hyo" border="1" width="400">
	            <tr>
	                <th nowrap class="header" width="64"  align="center">クラス</th>
	                <td nowrap class="detail" width="50"  align="center"><%=m_iGakunen%>年</td>
	                <td nowrap class="detail" width="130" align="center"><%=m_sClassNm%></td>
	                <th nowrap class="header" width="100" align="center">入力対象日</th>
	                <td nowrap class="detail" width="150" align="center"><%=m_sDate & "(" & gf_GetYoubi(Weekday(m_sDate)) & ")"%></td>
	            </tr>
	        </table>
		</td></tr><tr>
		<td align="center">
<%		if m_sEndDay < m_sDate then %>
            <table>
				<tr>
                <td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="　登　録　"></td>
                <td ><input class="button" type="button" onclick="javascript:f_Cancel();" value="キャンセル"></td>
				</tr>
            </table>
<% Else %>
            <table>
				<tr>
                <td ><input class="button" type="button" onclick="javascript:f_Cancel();" value=" 戻　る "></td>
				</tr>
            </table>
<% End If %>
		</td></tr>
        </table>

        <!--明細ヘッダ部-->
        <table >
<%		if m_sEndDay < m_sDate then %>
            <tr>
                <td align="center" colspan=3 valign="bottom">
                    <span class="CAUTION">※ 出欠状況欄をクリックして、出欠状況を入力してください。（欠→遅→早→空欄(出席)の順で表示されます）</span>
                </td>
            </tr>
<% Else%>
            <tr>
                <td align="center" colspan=3 valign="bottom">
                    <span class="CAUTION">※ 出欠状況を変更することはできません。</span>
                </td>
            </tr>

<% End If %>

            <tr><td valign="top">

                <!--ヘッダ-->
                <table class=hyo border="1" bgcolor="#FFFFFF">
                    <tr>
                        <th nowrap class="header" width="80"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
                        <th nowrap class="header" width="150" align="center">氏　名</th>
                        <th nowrap class="header" width="80" align="center">出欠状況</th>
                    </tr>

            <%If i = w_iCnt Then
                '//リストを改行する

                '//ｽﾀｲﾙｼｰﾄのｸﾗｽを初期化
				w_Class = ""
                %>
                </table>
                </td><td width="10"></td><td valign="top">
                <!--ヘッダ-->
                <table class="hyo" border="1" >
                    <tr>
                        <th nowrap class="header" width="80"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
                        <th nowrap class="header" width="150" align="center">氏　名</th>
                        <th nowrap class="header" width="80" align="center">出欠状況</th>
            <%End If%>
                </table>
                </td></tr>
            </table>

        <%Exit Do%>
    <%Loop%>

    <!--値渡し用-->
    <input type="hidden" name="NENDO"     value="<%=m_iSyoriNen%>">
    <input type="hidden" name="KYOKAN_CD" value="<%=m_iKyokanCd%>">
    <input type="hidden" name="GAKUNEN"   value="<%=m_iGakunen%>">
    <input type="hidden" name="CLASSNO"   value="<%=m_iClassNo%>">
    <input type="hidden" name="cboDate"   value="<%=m_sDate%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
'********************************************************************************
'*  [機能]  空白HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showWhitePage()
%>
    <html>
    <head>
    <title>日毎出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
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
    //-->
    </SCRIPT>
    </head>

	<body LANGUAGE=javascript onload="return window_onload()">
    </body>
    </html>
<%
End Sub
%>

