<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: レベル別科目決定
' ﾌﾟﾛｸﾞﾗﾑID : web/web0390/web0390_middle.asp
' 機      能: 下ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/10/26 谷脇 良也
' 変      更: 
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
    Dim     m_sKamokuCD     '//科目コード
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
        w_iRet = f_KyokanData()
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
resposne.write w_sMsg
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
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
    m_iDsp      = C_PAGE_LINE
	m_sKamokuCD = request("cboKamokuCode")
	m_sRisyuJotai = request("txtRisyu")
End Sub

Function f_KyokanData()
'******************************************************************
'機　　能：教官のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_KyokanData = 1

    Do


        '//科目のデータ取得
        m_sSQL = ""
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT "
        m_sSQL = m_sSQL & vbCrLf & "     T27_KYOKAN_CD,"
        m_sSQL = m_sSQL & vbCrLf & "     M04_KYOKANMEI_SEI,"
        m_sSQL = m_sSQL & vbCrLf & "     M04_KYOKANMEI_MEI"
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T27_TANTO_KYOKAN,M04_KYOKAN"
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T27_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T27_GAKUNEN = " & m_sGakunen & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T27_KAMOKU_CD = '" & m_sKamokuCD & "' "
        m_sSQL = m_sSQL & vbCrLf & " AND M04_KYOKAN_CD = T27_KYOKAN_CD"
        m_sSQL = m_sSQL & vbCrLf & " AND M04_NENDO = T27_NENDO"
        m_sSQL = m_sSQL & vbCrLf & " GROUP BY M04_KYOKANMEI_MEI,M04_KYOKANMEI_SEI,T27_KYOKAN_CD"
        m_sSQL = m_sSQL & vbCrLf & " ORDER BY T27_KYOKAN_CD"

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

    f_KyokanData = 0

    Exit Do

    Loop

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
			For i = 0 to m_KrCnt - 1
	            If w_iKhalf = i then
				    %>
				    </tr>
				    <tr>
				    <%
				End If 
			w_sTmp = "txtLKCnt"&i
				%>
	        <td class=disph><%=m_Krs("M04_KYOKANMEI_SEI")%> <%=m_Krs("M04_KYOKANMEI_MEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=Request(w_sTmp)%>" class="CELL2" name="KYOKAN<%=i%>" readonly></td>
	        <%m_Krs.MoveNext
	          Next%>
	    </tr>
	</table>
  <% If cint(m_sRisyuJotai) = C_K_RIS_MAE then %>
	<span class=CAUTION>※ 担当する教官の下の枠クリックし、○印をつけてください。(○は決定)</span>
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
	        <td align=center><input type=button class=button value=" 戻　る " onclick="javascript:f_Cansel()"></td>
	    </tr>
	</table>
  <% End If %>
	<table class=hyo border=1>
	    <tr>
	        <th class=header width="50"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
	        <th class=header width="120">氏　名</th>
	    <%

	        m_Krs.MoveFirst
	        Do Until m_Krs.EOF
		        n = n + 1
			    %>
		        <th class=header2 width=96 valign=middle><%=m_Krs("M04_KYOKANMEI_SEI")%> <%'m_Krs("M04_KYOKANMEI_MEI")%>
		        	<input type=hidden name=kyokanCd<%=n%> value="<%=m_Krs("T27_KYOKAN_CD")%>">
				</th>
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