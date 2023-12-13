<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ登録
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0133/top.asp
' 機      能: 上ページ 就職先マスタの登録を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           txtSinroCD2     :進路コード
'           txtSingakuCD2       :進学コード
'           txtSinroName        :就職先名称（一部）
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう就職先を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/22 岩下　幸一郎
' 変      更: 2001/07/13 谷脇　良也
' 変      更: 2001/08/22 伊藤　公子　業種区分追加対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    '市町村選択用のWhere条件
    Public m_sSinroWhere            '進路の条件
    Public m_sSingakuWhere      '進学コンボの条件
    Public m_sSingakuOption     '進学コンボのオプション
    Public m_sSyusyokuName
    Public m_sSinroCD
    Public m_sSingakuCD
    Public m_iNendo
    
    Public m_sMode

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
    w_sMsgTitle="就職先マスタ"
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
        '進路に関するWHREを作成する
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

    m_sSinroCD = Request("txtSinroCD")      ':進路コード
	If m_sSinroCD = "@@@" Then
		m_sSinroCD = ""
	End If
    m_sSingakuCD = Request("txtSingakuCD")  ':進学コード
    m_sSyusyokuName = Request("txtSyusyokuName")
    m_sMode = "search"
    m_iNendo = Session("nendo")

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

'---2001/08/22 ito 業種区分追加対応
	'// 進学
    If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_SINGAKU & "  AND "
        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
	'// 就職
	ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
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

Dim w_sSelectSinroCd
Dim w_sSelectSingakuCd

%>


<html>

<head>

<title>就職先マスタ登録</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

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

        document.frm.action = "./main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Touroku(){

        document.frm.action = "syusei.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "Sinki";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>


<link rel=stylesheet href="../../common/style.css" type=text/css>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" method="POST"  onSubmit="return false" onClick="return false">
<%call gs_title("進路先情報登録","一　覧")%>

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
					<% '共通関数から進路に関するコンボボックスを出力する（年度条件）
					    call gf_ComboSet("txtSinroCD",C_CBO_M01_KUBUN,m_sSinroWhere," onchange = 'javascript:f_ReLoadMyPage();' ",True,m_sSinroCD) %>

                </td>
                <td Nowrap align="left">
					<%
					'If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
					'	w_sTitle = "業種区分"
					'Else
					'	w_sTitle = "進学区分"
					'End If
					%>
					　種別区分
					<% '共通関数から進学に関するコンボボックスを出力する（年度、進路区分が条件）（進路区分が入力されていないときは、DISABLEDとなる）
					   call gf_ComboSet("txtSingakuCD",C_CBO_M01_KUBUN,m_sSingakuWhere,"style='width:100px;' " & m_sSingakuOption,True,m_sSingakuCd) %>

                </td>
                </tr>

                <tr>
                <td align="left" colspan="2">
	                進路先名称
	                <input type="text" name="txtSyusyokuName" size="20" Value="<%=m_sSyusyokuName%>" maxlength="60">   <!--'//2001/07/31修正-->
	                <font size="2">※進路先名称の一部で検索します</font>
                </td>
                </tr>
                <tr>
                <td Nowrap align="right" colspan="2">
			    <input class=button type="button" value="　表　示　" onClick="javascript:f_Search()">
                </td>
                </tr>
                </table>

        </td>
        </tr>
        </table>

    </td>
    <td valign="top">
    <a href="javascript:f_Touroku()" onClick="javascript:f_Touroku()">新規登録はこちら</a><br><img src="../../image/sp.gif" height="10"><br>
    </td>
  </tr>
</table>
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
