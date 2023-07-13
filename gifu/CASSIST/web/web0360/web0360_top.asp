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
'
' 説      明:
'           ■初期表示
'               クラブのコンボボックスを表示
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
    Public m_sClubCd

    '//コンボ用Where条件等
    Public m_sClubWhere

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
'call s_DebugPrint()

        '//クラブコンボに関するWHEREを作成する
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
    response.write "m_sClubCd   = " & m_sClubCd & "<br>"

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

	    var n = document.frm.txtClubCd.selectedIndex;
		if(document.frm.txtClubCd.options[n].value=="@@@"){
		    alert("クラブを選択してください");
			return;
		}

        document.frm.action="./web0360_main.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">

    <center>
    <%call gs_title("部活動部員一覧","一　覧")%>

    <table bordeer="0">
        <tr>
        <td class="search">
            <table border="0">
            <tr>
            <td>

                <table border="0" cellpadding="1" cellspacing="1">
	                <tr>
		                <td nowrap align="left">クラブ名</td>
		                <td nowrap align="left" >
							<% call gf_ComboSet("txtClubCd",C_CBO_M17_BUKATUDO,m_sClubWhere," style='width:140px;'",True,cstr(gf_SetNull2String(m_sClubCd))) %>
						</td>
				        <td valign="bottom" align="right">
				        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
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

    </form>
    </body>
    </html>
<%
End Sub
%>