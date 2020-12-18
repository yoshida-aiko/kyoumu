<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 部活動部員一覧
' ﾌﾟﾛｸﾞﾗﾑID : web/web0360/web0360_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:   txtClubCd		:部活CD
'
' 引      渡:   txtClubCd		:部活CD
'               cboGakunenCd	:学年
'               cboClassCd		:クラスNO
'               txtTyuClubCd	:中学校部活CD
' 説      明:
'           ■初期表示
'               学年、クラス、中学校部活のコンボボックスを表示
'-------------------------------------------------------------------------
' 作      成: 2001/08/22 伊藤公子
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iSyoriNen          '//教官ｺｰﾄﾞ
    Public m_iKyokanCd          '//年度
    Public m_sClubCd			'//部活CD
    Public m_iGakunen			'//学年
	Public m_iClassNo           '//クラスNO
	Public m_sTyuClubCd			'//中学校クラブCD

    '//コンボ用Where条件等
    Public m_sClubWhere
    Public m_sGakunenWhere      '//学年の条件
    Public m_sClassWhere        '//クラスの条件

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
    w_sMsgTitle="部活動部員一覧"
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

        '//学年コンボに関するWHEREを作成する
        Call s_MakeGakunenWhere() 

        '//クラスコンボに関するWHEREを作成する
        Call s_MakeClassWhere() 

        '//中学校クラブコンボに関するWHEREを作成する
        Call s_MakeClubWhere() 

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

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
	m_sClubCd = ""
	m_iClassNo   = ""
	m_sTyuClubCd = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")
	m_sClubCd   = Request("txtClubCd")
    m_iGakunen  = Request("cboGakunenCd")   '//学年

	m_iClassNo   = Request("cboClassCd")	'//クラス
	m_sTyuClubCd = replace(Request("txtTyuClubCd"),"@@@","")	'//中学校クラブCD

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
    response.write "m_sClubCd   = " & m_sClubCd   & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
	response.write "m_iClassNo   = " & m_iClassNo   & "<br>"
	response.write "m_sTyuClubCd = " & m_sTyuClubCd & "<br>"

End Sub

'********************************************************************************
'*  [機能]  学年コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeGakunenWhere()

    m_sGakunenWhere = ""
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyorinen
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

End Sub

'********************************************************************************
'*  [機能]  クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeClassWhere()

    m_sClassWhere = ""
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen

    If m_iGakunen = "" Then
        '//初期表示時は1年1組を表示する
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = 1"
    Else
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_iGakunen)
    End If

End Sub

'********************************************************************************
'*  [機能]  クラブコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeClubWhere()

    m_sClubWhere = ""
    m_sClubWhere = m_sClubWhere & " M17_NENDO =" & m_iSyoriNen  '//処理年度
    m_sClubWhere = m_sClubWhere & " AND M17_BUJYOKYO_KBN = 0"	'//部活動状況区分

End Sub

'********************************************************************************
'*  [機能]  部活名を取得する
'*  [引数]  p_sClubCd:部活CD
'*  [戻値]  f_GetClubName：部活名
'*  [説明]  
'********************************************************************************
Function f_GetClubName(p_sClubCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubName = ""
	w_sClubName = ""

	Do

		'//部活CDが空の時
		If trim(gf_SetNull2String(p_sClubCd)) = "" Then
			Exit Do
		End If

		'//部活動情報取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD=" & p_sClubCd

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//部活名
			w_sClubName = rs("M17_BUKATUDOMEI")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	f_GetClubName = w_sClubName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>

    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>部活動部員一覧</title>

    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
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
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        document.frm.action="./web0360_insMain.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  学年が変更されたとき、本画面を再表示
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./web0360_insTop.asp";
        document.frm.target="topFrame";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post" onClick="return false;">

    <%call gs_title("部活動部員一覧","登録")%>
    <center>

	<br>
    <table bordeer="0">
        <tr>
        <td class="search">
            <table border="0">
            <tr><td align="left">　―　<%=f_GetClubName(m_sClubCd)%>　入部登録　―</td></tr>
            <tr><td>
                <table border="0" cellpadding="1" cellspacing="1">
	                <tr>
		                <td nowrap align="left" >学年</td>
		                <td nowrap align="left">
		                    <% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' ",False,m_iGakunen) %>
		                </td>
		                <td nowrap align="left" width="40" >クラス</td>
		                <td nowrap align="left" >
		                    <% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"style='width:80px;' " & m_sClassOption,true,m_iClassNo) %>
		                </td>
		                <td nowrap align="left"><br></td>
<!--
					</tr>
					<tr>
		                <td nowrap align="left" colspan="2">中学校部活</td>
		                <td nowrap align="left" colspan="2">
							<% call gf_ComboSet("txtTyuClubCd",C_CBO_M17_BUKATUDO,m_sClubWhere," style='width:140px;'",True,"") %>
		                </td>
-->
				        <td valign="bottom" align="right">
				        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();" name="btnShow">
				        </td>
	                </tr>
                </table>

            </td>
            </tr>
            </table>
        </td>
        </tr>
    </table>
    </center>

    <!--値渡し用-->
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
	<input type="hidden" name="txtClubCd" value="<%=m_sClubCd%>">

    </form>
    </body>
    </html>
<%
End Sub
%>