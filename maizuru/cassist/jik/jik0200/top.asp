<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 教官別授業時間一覧
' ﾌﾟﾛｸﾞﾗﾑID : jik/jik0200/top.asp
' 機      能: 上ページ 教官別授業時間の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           :session("PRJ_No")      '権限ﾁｪｯｸのキー
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           cboKyokaKeiCd   :科目系列コード
'           cboKyokanCd     :教官コード
'           txtMode         :動作モード
'                            (BLANK)    :初期表示
'                            Reload     :リロード
' 説      明:
'           ■初期表示
'               コンボボックスは科目系列と教官を表示
'           ■表示ボタンクリック時
'               下のフレームに授業一覧を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/03 根本 直美
' 変      更: 2001/07/30 根本 直美  戻り先URL変更
' 　      　:                       定数未対応につき修正
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '取得したデータを持つ変数
    Public  m_iSyoriNen         '処理年度
    Public  m_iKyokanCd         '教官コード
    Public  m_sKyokanName       '教官コード
    
    'データ取得用

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
    w_sMsgTitle="教官別授業時間一覧"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False


        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
         

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

'response.write "AB"

        '//教官名称を取得
        w_iRet = f_GetKyokanNm(m_iKyokanCd,m_iSyoriNen,m_sKyokanName)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

'response.write "CD"

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

    m_iSyoriNen = Session("NENDO")
    'm_iSyoriNen = 2001      '//テスト用
    m_iKyokanCd = Session("KYOKAN_CD")

    '//教官CD
    If Request("SKyokanCd1") = "" Then
        m_iKyokanCd = Session("KYOKAN_CD")
    Else
        m_iKyokanCd = Request("SKyokanCd1")
    End If

End Sub

Function f_GetKyokanNm(p_sCD,p_iNENDO,p_sName)
'********************************************************************************
'*  [機能]  教官の氏名を取得
'*  [引数]  なし
'*  [戻値]  p_sName
'*  [説明]  
'********************************************************************************
Dim rs
Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetKyokanNm = 1
    w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M04_KYOKAN_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M04_NENDO = " & p_iNENDO & " "

'response.write w_sSQL

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKyokanNm = 99
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M04_KYOKANMEI_SEI") & "　" & rs("M04_KYOKANMEI_MEI")
        End If

        f_GetKyokanNm = 0
        Exit Do
    Loop

    p_sName = w_sName

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
        if(document.frm.SKyokanNm1.value==""){
			alert("教官を選択してください")
			return;
		}
        document.frm.action="main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  教官参照選択画面ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function KyokanWin(p_iInt,p_sKNm) {
		var obj=eval("document.frm."+p_sKNm)

        URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=610,top=0,left=0");
        nWin.focus();
        return true;    
    }

    //************************************************************
    //  [機能]  クリアボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function fj_Clear(){

		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}
    //-->
    </SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" action="./main.asp" target="main" Method="POST">
    <input type="hidden" name="txtMode">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
    <%call gs_title("教官別授業時間一覧","一　覧")%>
<%
If m_sMode = "" Then
%>
<br>
    <table border="0">
    <tr>
    <td class=search>
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left">
            <table border="0" cellpadding="1" cellspacing="1">
            <tr>
	            <td align="left" nowrap>
	            教官
	            </td>
	            <td align="left" nowrap colspan="2">
	                <input type="text" class="text" name="SKyokanNm1" VALUE='<%=m_sKyokanName%>' readonly>
	                <input type="hidden" name="SKyokanCd1" VALUE='<%=m_iKyokanCd%>'>
	                <input type="button" class="button" value="選択" onclick="KyokanWin(1,'SKyokanNm1')">
					<input type="button" class="button" value="クリア" onClick="fj_Clear()">
	            </td>
		    <td align="right" valign="bottom">　
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
<%
End IF
%>
</td>
</tr>
</table>
</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
