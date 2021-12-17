<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 連絡掲示板
' ﾌﾟﾛｸﾞﾗﾑID : web/web0330/sousin_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
'            モード         ＞      txtMode
'                                   新規 = NEW
'                                   更新 = UPDATE
'            件名           ＞      txtkenmei
'            内容           ＞      txtNaiyou
'            開始日         ＞      txtKaisibi
'            完了日         ＞      txtSyuryobi
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
    Const DebugFlg = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_sNendo             '年度
    Public m_sKyokanCd          '教官ｺｰﾄﾞ
    Public m_stxtMode           'モード
    Public m_stxtNo             '処理番号
    Public m_sKenmei            '件名
    Public m_sNaiyou            '内容
    Public m_sKaisibi           '開始日
    Public m_sSyuryoubi         '完了日
    Public m_rs

    Public m_sUserKbn
    Public m_sSimei
    Public m_sGakkaOption
    Public m_sKeiretuOption
    Public m_sKyokaKbnOption

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
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

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

	m_sUserKbn = Replace(request("UserKbn"),"@@@","")
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

		'//対象者区分が教官以外の場合は、教官区分等を選択できないようにする
		Call s_CtrlDisabled()

	    If m_stxtMode = "NEW" Then
	        Call showPage()
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
    Call gf_closeObject(m_Rs)
    '// 終了処理
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  コンボボックスのDISABLEDをセット
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_CtrlDisabled()

	m_sGakkaOption=""
	m_sKeiretuOption=""
	m_sKyokaKbnOption=""

	'//教官以外の場合
	If cint(gf_SetNull2Zero(m_sUserKbn)) <> C_USER_KYOKAN Then
		m_sGakkaOption="DISABLED"
		m_sKeiretuOption="DISABLED"
		m_sKyokaKbnOption="DISABLED"
	End If


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
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>連絡掲示板</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
    //************************************************************
    //  [機能]  表示ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Hyouji(){
		document.frm.BtnCtrl.value=""

        //リスト情報をsubmit
        document.frm.target = "main" ;
        document.frm.action = "sousin_main.asp";
        document.frm.submit();

    }
    //************************************************************
    //  [機能]  ﾘﾛｰﾄﾞ時
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./sousin_top.asp";
        document.frm.target="_self";
        document.frm.submit();

    }

	//-->
    </SCRIPT>

<center>

<FORM NAME="frm" action="post" onsubmit="return false">

<br>
<% 
    Select Case m_stxtMode
        Case "NEW"
            call gs_title("連絡掲示板","新　規")
        Case "UPD"
            call gs_title("連絡掲示板","修　正")
    End Select      
%>
<font>送　付　先　選　択</font>
<br>
</TD>
</TR>
</TABLE>

<br>

<table class=hyo width=46%>
    <tr>
        <td align="center" width=20%><font color="white">件　名</font></td>
        <td class=detail width=80%><%=m_sKenmei%></td>
    </tr>
</table>
<BR>
<div align="center"><span class=CAUTION>※ 表示ボタンをクリックし、送付先一覧を表示します。<br>
										※ コンボから条件を選択し送付先一覧の絞込みを行なう事ができます。<br>
										※ 対象者区分が教官の場合のみ、学科、教科系列区分、教官区分が選択可能となります。
</span></div>
<table border="0" width="78%">
    <tr>
<!--
        <td width="80%">
            <table border="0" class=search cellpadding="1" cellspacing="1" width="100%">
                <tr>
-->
                    <td align="left" class=search>
                        <table border="0" cellpadding="1" cellspacing="1" width="100%">
                            <tr>
                                <td align="left" width="10%" Nowrap>対象者区分</td>
                                <td width="20%" Nowrap colspan="1">：
                                <%
                                call gf_ComboSet("UserKbn",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_USER & " AND M01_NENDO=" & m_sNendo & " AND M01_SYOBUNRUI_CD>0","onchange='javascript:f_ReLoadMyPage()'",True,Request("UserKbn"))
                                %>
                                </td>

                                <td align="left" width=10% Nowrap>学　科</td>
                                <td widt=20% Nowrap>：
                                <%
                                call gf_ComboSet("Gakka",C_CBO_M02_GAKKA,"M02_NENDO='" & m_sNendo & "'",m_sGakkaOption,True,"")
                                %>
                                </td>
                                <td align="left" width=10% Nowrap>教科系列区分</td>
                                <td width=20% Nowrap>：
                                <%
                                call gf_ComboSet("KkeiKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_KYOKA_KEIRETU &" AND M01_NENDO=" & m_sNendo & " AND M01_SYOBUNRUIMEI IS NOT NULL ",m_sKeiretuOption,True,"")
                                %>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" width=10% Nowrap>教官区分</td>
                                <td width=20% Nowrap>：
                                <%
                                call gf_ComboSet("KkanKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD='" & C_KYOKAN &"' AND M01_NENDO='" & m_sNendo & "'",m_sKyokaKbnOption,True,"")
                                %>
                                </td>

                                <td align="left" width="10%" Nowrap>氏名</td>
                                <td width="20%"  colspan="1" Nowrap>： <input type="text" name="txtSimei" size="25" value="<%=Request("txtSimei")%>"></td>

                                <td align="left" width="10%" Nowrap><br></td>
						        <td width="30%" valign="bottom" align="left" colspan="2">　
						        <input class=button type="button" value="　表　示　" onClick = "javascript:f_Hyouji()">
						        </td>
							</tr>

                        </table>
                    </td>
<!--
                </tr>
            </table>
        </td>
-->
    </tr>
</table>
    <INPUT TYPE=HIDDEN  NAME=txtNo          value="<%=m_stxtNo%>">
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="<%=m_stxtMode%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_sNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtKenmei      value="<%=m_sKenmei%>">
    <INPUT TYPE=HIDDEN  NAME=txtNaiyou      value="<%=m_sNaiyou%>">
    <INPUT TYPE=HIDDEN  NAME=txtKaisibi     value="<%=m_sKaisibi%>">
    <INPUT TYPE=HIDDEN  NAME=txtSyuryoubi   value="<%=m_sSyuryoubi%>">

    <INPUT TYPE=HIDDEN  NAME=BtnCtrl value="<%=Request("BtnCtrl")%>">

</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>