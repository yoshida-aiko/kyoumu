<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 追試受講者一覧
' ﾌﾟﾛｸﾞﾗﾑID : saisi/saisi0400/top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
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
    Const DebugFlg = 6
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_sNendo             '年度
    Public m_PgMode             '処理別フラグ
    Public m_sMsgTitle          'ﾀｲﾄﾙ

    Public m_Rs

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
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    On Error Resume Next

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="再試受講者一覧"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    'm_PgMode=request("p_mode")
	'Select Case m_PgMode
	'	Case "P_HAN0100"
	'	    m_sMsgTitle="成績一覧表"
	'	Case "P_KKS0200"
	'	    m_sMsgTitle="欠課一覧表"
	'	Case "P_KKS0210"
	'	    m_sMsgTitle="遅刻一覧表"
	'	Case "P_KKS0220"
	'	    m_sMsgTitle="行事欠課一覧表"
	'	Case Else
	'End Select
	'w_sMsgTitle = m_sMsgTitle

    Err.Clear

    m_bErrFlg = False

    m_sNendo    = session("NENDO")
    m_iDsp = C_PAGE_LINE

    Do

		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '// ページを表示
        Call showPage()
        Exit Do
    Loop

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
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>再試受講者一覧</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Touroku(){

        //リスト情報をsubmit
        //document.frm.target="_parent";
        //document.frm.action="regist.asp";
        //document.frm.txtMode.value ="NEW";
        //document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript>
    <form name="frm" method="post">
    <center>
<%call gs_title("再試受講者一覧","一　覧")%>
<br>
    <!--INPUT TYPE=HIDDEN NAME=txtMode     VALUE=""-->
    <INPUT TYPE=HIDDEN NAME=txtNendo    VALUE="<%=m_sNendo%>">
    <!--INPUT TYPE=HIDDEN NAME=txtKyokanCd VALUE="<%=m_sKyokanCd%>"-->

    </center>

    </form>
    </body>
    </html>
<%
    '---------- HTML END   ----------
End Sub
%>
