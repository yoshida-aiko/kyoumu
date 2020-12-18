<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 留年該当者一覧
'* ﾌﾟﾛｸﾞﾗﾑID : han/han0121/header.asp
'* 機      能: 上ページ 留年該当者一覧の検索を行う
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*           :session("PRJ_No")      '権限ﾁｪｯｸのキー
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*           cboGakunenCd      :学年コード
'* 説      明:
'*           ■初期表示
'*               コンボボックスは学年を表示
'*           ■表示ボタンクリック時
'*               下のフレームに指定した条件の留年該当者一覧を表示させる
'*-------------------------------------------------------------------------
'* 作      成: 2001/08/08 前田　智史
'* 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '選択用のWhere条件
    Public m_sGakunenWhere      '学年の条件
    
    '取得したデータを持つ変数
    Public  m_iNendo         ':処理年度
    Public  m_iKyokanCd         ':教官コード
    
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
    w_sWinTitle= "キャンパスアシスト"
    w_sMsgTitle= "留年該当者一覧"
    w_sMsg= ""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget= ""


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

        '学年コンボに関するWHEREを作成する
        Call s_MakeGakunenWhere() 

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

Sub s_SetParam()
'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_iNendo = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub

Sub s_MakeGakunenWhere()
'********************************************************************************
'*  [機能]  学年コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    
    m_sGakunenWhere = ""
    
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iNendo
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"
    
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
<link rel=stylesheet href="../../common/style.css" type=text/css>
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
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        if( f_Trim(document.frm.cboGakunenCd.value) == "<%=C_CBO_NULL%>" ){
            window.alert("学年の選択を行ってください");
            document.frm.cboGakunenCd.focus();
            return ;
		}

        document.frm.action="ichiran.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" method="post">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
	<tr>
		<td valign="top" align="center">
    	<%call gs_title("留年該当者一覧","一　覧")%>
<br>
    		<table border="0">
    			<tr>
    				<td class=search>
				        <table border="0" cellpadding="1" cellspacing="1">
					        <tr>
						        <td align="left">
						            <table border="0" cellpadding="1" cellspacing="1">
							            <tr>
								            <td align="left">学年</td>
								            <td align="left">
								            <% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere," style='width:40px;' ",True,"") %>
								            年</td>
										    <td valign="bottom">
									        <input type="button" value="　表　示　" onClick = "javascript:f_Search()" class=button>
										    </td>
							            </tr>
						            </table>
						        </td>
					        </tr>
				        </table>
				    </td>
			    </tr>
		    </table>
		</td>
	</tr>
</table>

	<input type=hidden name=txtMode value="Hyouji">
	<input type=hidden name=txtNendo value="<%=m_iNendo%>">
	<input type=hidden name=txtKyokanCd value="<%=m_iKyokanCd%>">

</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
