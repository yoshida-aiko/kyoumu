<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 進路先情報検索
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0133/top.asp
' 機      能: 上ページ 就職先マスタの検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           txtSinroCD      :進路区分
'           txtSingakuCd    :進学区分
'           txtSinroName        :就職先名称（一部）

'           :txtSyusyokuName        :就職先名称（一部） '/2001/07/31追加
'           :txtMode                :モード             '/2001/07/31追加
'           :txtFLG                 :                   '/2001/07/31追加
'           :txtSNm                 :                   '/2001/07/31追加
'           :txtNendo               :年度               '/2001/07/31追加
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう就職先を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/15 岩下　幸一郎
' 変      更: 2001/07/31 根本 直美  引数・引渡追加
'           :                       進路先名称テキストボックスMAXLENGTH追加
'           :                       変数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    '市町村選択用のWhere条件
    Public m_sSinroWhere    '進路の条件
    Public m_sSingakuWhere  '進学コンボの条件
    Public m_sSingakuOption '進学コンボのオプション
    Public m_sSyusyokuName  ':就職先名称（一部）
    Public m_iSinroCD       ':進路区分      '/2001/07/31変更
    Public m_iSingakuCd     ':進学区分      '/2001/07/31変更
    Public m_iNendo         ':年度
    Public m_sMode          ':モード
    Public m_iFLG
    Public m_sSNm

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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="進路先情報検索"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
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
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
        '進路に関するWHREを作成する
        Call f_MakeSinroWhere() 
        '進学&就職に関するWHREを作成する
        Call f_MakeSingakuWhere()   

        '// ページを表示
        Call showPage()
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
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

    m_iSinroCD = Request("txtSinroCD")              ':進路区分
    m_iSingakuCd = Request("txtSingakuCD")          ':進学区分
    m_sSyusyokuName = Request("txtSyusyokuName")    ':就職先名称（一部）
    m_sMode = request("txtMode")                    ':モード    
    m_iNendo = Session("NENDO")                     ':年度
    m_iFLG = request("txtFLG")
    m_sSNm = request("txtSNm")
End Sub


Sub f_MakeSinroWhere()
'********************************************************************************
'*  [機能]  進路コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sSinroWhere=""

    m_sSinroWhere = " M01_DAIBUNRUI_CD = " & C_SINRO & "  AND "
    m_sSinroWhere = m_sSinroWhere & " M01_NENDO = " & m_iNendo & ""

End Sub

Sub f_MakeSingakuWhere()
'********************************************************************************
'*  [機能]  進学コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sSingakuWhere=""
    m_sSingakuOption=""

	'// 進学
    If m_iSinroCD = "1" Then
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_SINGAKU & "  AND "
        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
	'// 就職
	ElseIf m_iSinroCD = "2" Then
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_GYOSYU_KBN & "  AND "
        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
	'// その他
    Else
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = 0 "
        m_sSingakuOption = " DISABLED "
    End IF

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

<title>進路先情報検索</title>
    <!--#include file="../../Common/jsCommon.htm"-->
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

    //************************************************************
    //  [機能]  進路が修正されたとき、再表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="top.asp";
        document.frm.target="_self";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        document.frm.action="./main.asp";
        document.frm.target="main";
        document.frm.txtMode.value = "Search";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  クリアボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Clear(){

        document.frm.txtSinroCD.value = "@@@";
        document.frm.txtSingakuCd.value = "@@@";
        document.frm.txtSyusyokuName.value = "";
    
    }

    //-->
    </SCRIPT>

    <link rel=stylesheet href="../../common/style.css" type=text/css>

    </HEAD>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<center>

<form name="frm" Method="POST"  onSubmit="return false" onClick="return false;">

<%
call gs_title("進路先情報検索","一　覧")
%>
<br>
    <table border="0">
    <tr>
    <td>

        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left" class=search>

                <table border="0" cellpadding="1" cellspacing="1">
                <tr>
                <td Nowrap align="left">
            進路区分<img src="../../image/sp.gif" width="15">
<%          '共通関数から進路に関するコンボボックスを出力する（年度条件）
            call gf_ComboSet("txtSinroCD",C_CBO_M01_KUBUN,m_sSinroWhere,"onchange = 'javascript:f_ReLoadMyPage()' ",True,m_iSinroCD)%>

                </td>
                <td Nowrap align="left">種別区分

<%          '共通関数から進学に関するコンボボックスを出力する（年度、進路区分が条件）（進路区分が入力されていないときは、DISABLEDとなる）
            call gf_ComboSet("txtSingakuCd",C_CBO_M01_KUBUN,m_sSingakuWhere,m_sSingakuOption & " style='width:100px;'",True,m_iSingakuCd)%>
                </td>
                </tr>

                <tr>
                <td align="left" colspan="2" nowrap>
                進路先名称
                <input type="text" name="txtSyusyokuName" size="20" Value="<%=m_sSyusyokuName%>" maxlength="60">   <!--'//2001/07/31修正-->
	            <font size="2">※進路先名称の一部で検索します</font>
                </td>
                </tr>
				<tr>
					<td valign="bottom" align="right" colspan="2">
			        <input type="button" class="button" value=" ク　リ　ア " onclick="javasript:f_Clear();">
					<input class="button" type="button" value="　表　示　" onClick = "javascript:f_Search()">
					</td>
				</tr>
                </table>
	        </td>
        </tr>
        </table>
    </td>
  </tr>
</table>
<input type="hidden" name="txtFLG" value="<%=m_iFLG%>">
<input type="hidden" name="txtSNm" value="<%=m_sSNm%>">
<input type="hidden" name="txtMode" value="<%=m_sMode%>">
<input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
</form>

</center>

</body>

</html>






<%
    '---------- HTML END   ----------
End Sub
%>
