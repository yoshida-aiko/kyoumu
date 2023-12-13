<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 教官参照選択画面
' ﾌﾟﾛｸﾞﾗﾑID : Common/com_select/SEL_KYOKAN/SEL_KYOKAN_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/19 前田
' 変      更: 2001/08/08 根本 直美     NN対応に伴うソース変更
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iNendo             '年度
    Public m_sKyokanCd          '教官ｺｰﾄﾞ
    Public m_iI                 '
    Public m_sKNm               '教官名
    Public m_sGakkaCd               '所属学科

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
    w_sMsgTitle="利用者選択画面"
    w_sMsg=""
    w_sRetURL="../../../../default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iI        = request("txtI")
    m_sKNm      = request("txtKNm")
    m_sGakkaCd      = request("txtGakka")

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

		'//対象者区分が教官以外の場合は、教官区分等を選択できないようにする
		Call s_CtrlDisabled()

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

<link rel="stylesheet" href="../../style.css" type="text/css">
    <title>利用者選択画面</title>

    <!--#include file="../../jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [機能]  表示ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Hyouji(){

        //リスト情報をsubmit
        document.frm.target = "main" ;
        document.frm.action = "SEL_KYOKAN_main.asp";
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

        document.frm.action="./SEL_KYOKAN_top.asp";
        document.frm.target="_self";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" action="post" onsubmit="return false">

<br>
<% 
    call gs_title("利用者選択画面","一　覧")
%>
</TD>
</TR>
</TABLE>

<table border="0" width="100%">
    <tr>
        <td width="100%">
            <table border="0" class="search" cellpadding="1" cellspacing="1" width="100%">
                <tr>
                    <td align="left" class="search" nowrap>
                        <table border="0" cellpadding="1" cellspacing="1" width="100%">
                            <tr>
                                <td align="left" width="16%" Nowrap>利用者区分</td>
                                <td width="34%" Nowrap colspan="1">
                                <%
                                call gf_ComboSet("UserKbn",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_USER & " AND M01_NENDO=" & m_iNendo & " AND M01_SYOBUNRUI_CD>0","onchange='javascript:f_ReLoadMyPage()'",True,Request("UserKbn"))
                                %>
                                </td>

                                <td align="left" width="16%" Nowrap>学　科</td>
                                <td width="34%" Nowrap colspan="3">
                                <%
                                call gf_ComboSet("Gakka",C_CBO_M02_GAKKA,"M02_NENDO='" & m_iNendo & "'",m_sGakkaOption,True,m_sGakkaCd)
                                %>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" width="16%" Nowrap>教官区分</td>
                                <td width="34%" Nowrap>
                                <%
                                call gf_ComboSet("KkanKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD='" & C_KYOKAN &"' AND M01_NENDO='" & m_iNendo & "'",m_sKyokaKbnOption,True,"")
                                %>
                                </td>
                                <td align="left" width="16%" Nowrap>教科系列区分</td>
                                <td width="34%" Nowrap>
                                <%
                                call gf_ComboSet("KkeiKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_KYOKA_KEIRETU &" AND M01_NENDO=" & m_iNendo & " AND M01_SYOBUNRUIMEI IS NOT NULL ",m_sKeiretuOption,True,"")
                                %>
                                </td>

                            </tr>
                            <tr>

                                <td align="left" width="16%" Nowrap>氏名</td>
                                <td width="34%"  colspan="1" Nowrap><input type="text" name="txtSimei" size="25" value="<%=Request("txtSimei")%>"></td>
                                <td align="left" width="16%" Nowrap><br></td>
						        <td align="right" width="34%" Nowrap ><input class="button" type="button" value="　表　示　" onClick = "javascript:f_Hyouji()"></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<span class="CAUTION">※ 利用者区分が教官の場合のみ、学科、教官区分、教科系列区分が選択可能となります。 </span>
    <input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
    <input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
    <input type="hidden" name="txtI"        value="<%=m_iI%>">
    <input type="hidden" name="txtKNm"      value="<%=m_sKNm%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>