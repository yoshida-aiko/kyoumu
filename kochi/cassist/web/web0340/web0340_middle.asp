<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 個人履修選択科目決定
' ﾌﾟﾛｸﾞﾗﾑID : web/web0340/web0340_main.asp
' 機      能: 下ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/25 前田
' 変      更: 2001/08/28 伊藤公子 ヘッダ部切り離し対応
' 変      更: 2015/08/19 清本 1年間番号の幅を50→70に変更
' 変      更: 2015/08/27 藤林 科目のデータ取得方法変更(T15_RISYU→T16_RISYU_KOJIN)
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_iMax          ':最大ページ
    Public  m_iDsp          '// 一覧表示行数
    Public  m_sPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_Krs           '科目用レコードセット
    Public  m_KSrs          '科目数のレコードセット
    Dim     m_iNendo        '//年度
    Dim     m_sKyokanCd     '//教官コード
    Dim     m_sGakunen      '//学年
    Dim     m_sClass        '//クラス
    Dim     m_sKBN          '//区分
    Dim     m_sGRP          '//グループ区分
    Dim     m_KrCnt         '//科目のレコードカウント
    Dim     m_KSrCnt        '//科目数のレコードカウント
    Dim     m_cell          '配色の設定
	Dim		m_sRisyuJotai	'履修状態フラグ add 2001/10/25
    Dim     i               
    Dim     j               
    Dim     k               

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
    w_sRetURL=C_RetURL & C_ERR_RETURL
    w_sTarget=""

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

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '//科目の情報取得
        w_iRet = f_KamokuData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

		If m_Krs.EOF Then
			Call showPage_NoData()
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

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_Krs)
    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    'Call gf_closeObject(m_Grs)
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

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sGakunen  = request("txtGakunen")
    m_sClass    = request("txtClass")
    m_sKBN      = Cint(request("txtKBN"))
    m_sGRP      = Cint(request("txtGRP"))
    m_iDsp      = C_PAGE_LINE
    m_sRisyuJotai = Request("txtRisyu")

End Sub

Function f_KamokuData()
'******************************************************************
'機　　能：科目のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear
    f_KamokuData = 1

    Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

        '//科目のデータ取得
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT DISTINCT "
        m_sSQL = m_sSQL & vbCrLf & "     T16_KAMOKUMEI,T16_KAMOKU_CD,T16_HAITOTANI"
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T16_RISYU_KOJIN "
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T16_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_HISSEN_KBN = " & C_HISSEN_SEN & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_HAITOTANI <> " & C_T15_HAITO & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_GRP = " & m_sGRP & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_KAMOKU_KBN = " & m_sKBN & " "
        m_sSQL = m_sSQL & vbCrLf & " AND EXISTS ( SELECT 'X' "
        m_sSQL = m_sSQL & vbCrLf & "              FROM  "
        m_sSQL = m_sSQL & vbCrLf & "                    T11_GAKUSEKI,T13_GAKU_NEN "
        m_sSQL = m_sSQL & vbCrLf & "              WHERE  "
        m_sSQL = m_sSQL & vbCrLf & "                    T13_NENDO = T16_NENDO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_GAKUSEI_NO = T16_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_GAKUSEI_NO = T11_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T11_NYUNENDO = " & w_iNyuNendo & " "
        m_sSQL = m_sSQL & vbCrLf & "             ) "

'response.write m_sSQL & "<BR>"

        Set m_Krs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Krs, m_sSQL,m_iDsp)

'response.write "w_iRet = " & w_iRet & "<BR>"
'response.write m_Krs.EOF & "<BR>"

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
    m_KrCnt=gf_GetRsCount(m_Krs)

    f_KamokuData = 0

    Exit Do

    Loop

End Function

Function f_KamokusuData()
'******************************************************************
'機　　能：科目数のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_KamokusuData = 1

    m_KSrCnt=""

    Do
        '//科目数のデータ取得
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT T16_KAMOKU_CD "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T16_RISYU_KOJIN ,T13_GAKU_NEN "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T16_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & " AND T16_NENDO = T13_NENDO "
        m_sSQL = m_sSQL & " AND T16_GAKUSEI_NO = T13_GAKUSEI_NO "
        m_sSQL = m_sSQL & " AND T16_HAITOGAKUNEN = T13_GAKUNEN "
        m_sSQL = m_sSQL & " AND T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & " AND T16_SELECT_FLG = " & C_SENTAKU_YES & " "
        m_sSQL = m_sSQL & " AND T16_KAMOKU_CD = '" & m_Krs("T16_KAMOKU_CD") & "' "
        m_sSQL = m_sSQL & " AND T16_HAITOGAKUNEN = " & m_sGakunen & " "
        m_sSQL = m_sSQL & " AND T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO & " "
'response.write m_sSQL
        Set m_KSrs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_KSrs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
	    m_KSrCnt=gf_GetRsCount(m_KSrs)

        If m_KSrs.EOF Then
            m_KSrCnt = "0"%>
	        <td class=disph><%=m_Krs("T16_KAMOKUMEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=m_KSrCnt%>" class="CELL2" name=Kamoku<%=i%> readonly></td>
        <%Else%>
	        <td class=disph><%=m_Krs("T16_KAMOKUMEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=m_KSrCnt%>" class="CELL2" name=Kamoku<%=i%> readonly></td>
        <%End If

	    f_KamokusuData = 0

	    Exit Do

    Loop


    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_KSrs)

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If

End Function

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
	<SCRIPT language="javascript">
	<!--
    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		parent.location.href = "white.asp?txtMsg=個人履修選択科目のデータがありません。"
        return;
    }
	//-->
	</SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>
    </center>
	<form name="frm" method="post">

	<input type="hidden" name="txtMsg" value="個人履修選択科目のデータがありません。">

	</form>
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
Dim w_iKhalf
Dim w_iGhalf
Dim n

    On Error Resume Next
    Err.Clear

i = 0
k = 0
n = 0
%>
<HTML>
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>個人履修選択科目決定</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Cansel(){

        //空白ページを表示
        parent.document.location.href="default2.asp"
    
    }
    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){
        parent.bottom.f_Touroku();
    }
    //-->
    </SCRIPT>

	<center>

	<FORM NAME="frm" method="post">
	<table class=disp border=1>
	    <tr>
	        <th class=header rowspan=2 width=16>選択数</th>
	    <%
	        m_Krs.MoveFirst
	        w_iKhalf = gf_Round(m_KrCnt / 2,0)
	        Do Until m_Krs.EOF
	            i = i + 1 
	            If w_iKhalf + 1 = i then
				    %>
				    </tr>
				    <tr>
				    <%
				End If 
		        Call f_KamokusuData() 
		        m_Krs.MoveNext
	        Loop%>
	    </tr>
	</table>
  <% If cint(m_sRisyuJotai) = C_K_RIS_MAE then %>
	<span class=CAUTION>※ 決定する科目をクリックし、○印をつけてください。(数字は各学生の希望順位、○は決定)</span>
	<table>
	    <tr>
	        <td align=center><input type=button class=button value="　登　録　" onclick="javascript:f_Touroku()"></td>
	        <td align=center><input type=button class=button value="キャンセル" onclick="javascript:f_Cansel()"></td>
	    </tr>
	</table>
  <% Else %>
	<BR>
	<table border="0">
	    <tr>
	        <td align=center><FONT size="1">	<BR><BR></FONT></td>
	    </tr>
	</table>
  <% End If %>
	<table class=hyo border=1>
	    <tr>
	        <th class=header width=70 height=34><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
	        <th class=header width=120>氏　名</th>
	    <%

	        m_Krs.MoveFirst
	        Do Until m_Krs.EOF
		        n = n + 1
			    %>
		        <th class=header2 width=96 valign=middle><%=m_Krs("T16_KAMOKUMEI")%>
		        <input type=hidden name=kamokuCd<%=n%> value="<%=m_Krs("T16_KAMOKU_CD")%>">
		        <input type=hidden name=Tanisuu<%=n%> value="<%=m_Krs("T16_HAITOTANI")%>"></th>
			    <%
		        m_Krs.MoveNext
	        Loop%>
	    </tr>
	</table>

	</FORM>
	</center>
	</BODY>
	</HTML>
<%
End Sub
%>