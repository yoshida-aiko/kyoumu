<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 高等学校情報検索
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0123/top.asp
' 機      能: 上ページ 高等学校マスタの検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           txtKubun        :区分コード
'           txtKenCd        :県コード
'           txtSityoCd      :市町村コード
'           txtSyuName      :高等学校名称（一部）
'           txtSyuKbn       :高等学校区分
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう高等学校を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/15 高丘 知央
' 変      更: 2001/06/20 岩下 幸一郎
'           : 2001/07/31 根本 直美  引数・引渡追加
'           :                       高等学校名称テキストボックスMAXLENGH追加
'           :                       関数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    '市町村選択用のWhere条件
    Public m_sKenWhere          '県の条件
    Public m_sSityoWhere        '市町村コンボの条件
    Public m_sSityoOption       '市町村コンボのオプション
    Public m_sSyuWhere          '高等学校の条件
    Public m_sSyuOption         '高等学校コンボのオプション
    Public m_sKenSentakuWhere   '選択した県
    Public m_sSityoSentakuWhere '選択した市町村
    Public m_sSyuSentakuWhere   '選択した高等学校区分

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
    w_sMsgTitle="高等学校情報検索"
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

    If Request("txtMode") = "Search" Then

        '県に関するWHREをQuery.Stringから作成する
        Call s_QueryKenWhere()  
        '市町村に関するWHREをQuery.Stringから作成する
        Call s_QuerySityoWhere()
        '高等学校に関するWHREをQuery.Stringから作成する
        Call s_QuerySyuWhere()  
        Else

        '県に関するWHREを作成する
        Call s_MakeKenWhere()   
        '市町村に関するWHREを作成する
        Call s_MakeSityoWhere() 
        '高等学校に関するWHREを作成する
        Call s_MakeSyuWhere()   
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
    
    '// 終了処理
    Call gs_CloseDatabase()
End Sub


Sub s_MakeKenWhere()'/2001/07/31変更
'********************************************************************************
'*  [機能]  県コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sKenWhere=""
    m_sKenSentakuWhere=""
        m_sKenWhere = "     M16_NENDO = '" & Session("NENDO") & "' "
        m_sKenSentakuWhere = Request("txtKenCd")
End Sub

Sub s_MakeSityoWhere()'/2001/07/31変更
'********************************************************************************
'*  [機能]  市町村コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sSityoWhere=""
    m_sSityoOption=""

    If Request("txtKenCd") <> "" Then
        m_sSityoWhere = "     M12_KEN_CD = '" & Request("txtKenCd") & "' "
        m_sSityoWhere = m_sSityoWhere & " GROUP BY M12_SITYOSON_CD,M12_SITYOSONMEI "
    Else
        m_sSityoOption = " DISABLED "
        m_sSityoWhere  = " M12_Ken_CD = '0' "
    End IF

End Sub

Sub s_MakeSyuWhere()'/2001/07/31変更
'********************************************************************************
'*  [機能]  高等学校コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sSyuWhere=""
    m_sSyuSentakuWhere=""
        m_sSyuWhere = "     M01_NENDO = '" & Session("NENDO") & "' "
        m_sSyuWhere = m_sSyuWhere & " AND M01_DAIBUNRUI_CD = " & C_SYUSSINKO
        m_sSyuSentakuWhere = Request("txtSyuKbn")
End Sub

Sub s_QueryKenWhere()'/2001/07/31変更
'********************************************************************************
'*  [機能]  県コンボに関するWHREをQuery.Stringから作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    m_sKenWhere=""
    m_sKenSentakuWhere=""

        m_sKenWhere = "     M16_NENDO = '" & Session("NENDO") & "' "
        m_sKenSentakuWhere = Request("txtKenCd")
End Sub


Sub s_QuerySityoWhere()'/2001/07/31変更
'********************************************************************************
'*  [機能]  市町村コンボに関するWHREをQuery.Stringから作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sSityoSentakuWhere=""
    m_sSityoWhere=""

    If Request("txtKenCd")<>"" Then
        m_sSityoWhere = "     M12_KEN_CD = '" & Request("txtKenCd") & "' "
        m_sSityoWhere = m_sSityoWhere & " GROUP BY M12_SITYOSON_CD,M12_SITYOSONMEI "
        m_sSityoSentakuWhere = Request("txtSityoCd")
    Else
        m_sSityoOption=" DISABLED "
        m_sSityoWhere = " M12_Ken_CD = '0' "
    End IF

End Sub

Sub s_QuerySyuWhere()'/2001/07/31変更
'********************************************************************************
'*  [機能]  高等学校コンボに関するWHREをQuery.Stringから作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    m_sSyuWhere=""
    m_sSyuSentakuWhere=""

        m_sSyuWhere = "     M01_NENDO = '" & Session("NENDO") & "' "
        m_sSyuWhere = m_sSyuWhere & " AND M01_DAIBUNRUI_CD = " & C_SYUSSINKO
        m_sSyuSentakuWhere = Request("txtSyuKbn")
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

<title>高等学校マスタ参照</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  県が修正されたとき、再表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./top.asp";
        document.frm.target="top";
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

        document.frm.txtSyuKbn.value = "@@@";
        document.frm.txtKenCd.value = "@@@";
        document.frm.txtSityoCd.value = "@@@";
        document.frm.txtSyuName.value = "";
    
    }

    //-->
    </SCRIPT>

    <link rel=stylesheet href=../../common/style.css type=text/css>
</head>

<body>
<div align="center">

<form name="frm" Method="POST" onSubmit="return false" onClick="return false;">
<input type="hidden" name="txtMode" value="">

<%call gs_title("高等学校情報検索","一　覧")%>

<img src="../../image/sp.gif" height="10"><br>
    
<table>
	<tr>
		<td class=search>

	        <table border="0">
		        <tr>
					<td Nowrap>区分</td>
			        <td Nowrap>
						<%  '共通関数から学校区分に関するコンボボックスを出力する（年度条件）
						        call gf_ComboSet("txtSyuKbn",C_CBO_M01_KUBUN,m_sSyuWhere,"",True,m_sSyuSentakuWhere)
						%>
			        </td>
			        <td align="left" valign="top" Nowrap>都道府県<!-- <select name="gakunen"> -->
				        <%  '共通関数から県に関するコンボボックスを出力する（年度条件）
				        call gf_ComboSet("txtKenCd",C_CBO_M16_KEN,m_sKenWhere,"onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sKenSentakuWhere)%>
			        </td Nowrap>
	        		<td align="center" valign="top" Nowrap>　市町村  <!-- <select name="gakka"> -->
			        <%  '共通関数から市町村に関するコンボボックスを出力する（年度、県が条件）（県が入力されていないときは、DISABLEDとなる）
			        call gf_ComboSet("txtSityoCd",C_CBO_M12_SITYOSON,m_sSityoWhere,"style='width:200px;' " & m_sSityoOption,True,m_sSityoSentakuWhere)%>
			        </td>
		        </tr>
		        <tr>
					<td Nowrap>高校名称</td>
					<td Nowrap><input type="text" size="20" name="txtSyuName" value="<%=Request("txtSyuName")%>" maxlength="60"></font></td>
			        <td colspan="1" Nowrap><font size="2">※高校名称の一部で検索します</font></td>
					<td valign="bottom" align="right" Nowrap>
			        <input type="button" class="button" value=" ク　リ　ア " onclick="javasript:f_Clear();">
					<input class="button" type="button" value="　表　示　" onClick = "javascript:f_Search()">
					</td>
		        </tr>
	        </table>

		</td>
	</tr>
</table>

</form>
</div>
</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
