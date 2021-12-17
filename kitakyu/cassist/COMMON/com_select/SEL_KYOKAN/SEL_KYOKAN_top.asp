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
    Const DebugFlg = 6
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iNendo             '年度
    Public m_sKyokanCd          '教官ｺｰﾄﾞ
    Public m_iI                 '
    Public m_sKNm               '教官名
    Public m_sGakkaCd               '所属学科

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
    w_sMsgTitle="教官参照選択画面"
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

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
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
    Call gf_closeObject(m_Rs)
    '// 終了処理
    Call gs_CloseDatabase()
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
    <title>教官参照選択画面</title>

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
    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" action="post" onClick = "return;false;">

<br>
<% 
    call gs_title("教官参照選択画面","一　覧")
%>
</TD>
</TR>
</TABLE>

<table border="0" width="100%">
    <tr>
<!--
        <td width="100%">
            <table border="0" class="search" cellpadding="1" cellspacing="1" width="100%">
                <tr>
-->
                    <td align="left" class="search" nowrap>
                        <table border="0" cellpadding="1" cellspacing="1" width="100%">
                            <tr>
                                <td align="center" width="16%" Nowrap>学　科</td>
                                <td width="84%" Nowrap colspan="3">
                                <%
                                call gf_ComboSet("Gakka",C_CBO_M02_GAKKA,"M02_NENDO='" & m_iNendo & "'","",True,m_sGakkaCd)
                                %>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" width="16%" Nowrap>教官区分</td>
                                <td width="34%" Nowrap>
                                <%
                                call gf_ComboSet("KkanKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD='" & C_KYOKAN &"' AND M01_NENDO='" & m_iNendo & "'","",True,"")
                                %>
                                </td>
                                <td align="center" width="26%" Nowrap>教科系列区分</td>
                                <td width="24%" Nowrap>
                                <%
                                'call gf_ComboSet("KkeiKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD='" & C_KYOKA_KEIRETU &"' AND M01_NENDO='" & m_iNendo & "'","",True,"")
                                call gf_ComboSet("KkeiKBN",C_CBO_M01_KUBUN,"M01_DAIBUNRUI_CD=" & C_KYOKA_KEIRETU &" AND M01_NENDO=" & m_iNendo & " AND M01_SYOBUNRUIMEI IS NOT NULL ","",True,"")


                                %>
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
    <tr>
        <td align="right"><input class="button" type="button" value="　表　示　" onClick = "javascript:f_Hyouji()"></td>
    </tr>
</table>
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