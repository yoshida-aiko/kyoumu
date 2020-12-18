<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: マイページのお知らせ
' ﾌﾟﾛｸﾞﾗﾑID : login/view.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/23 前田
' 変      更: 2001/08/07 根本 直美     NN対応に伴うソース変更
'           : 2001/08/10 根本 直美     NN対応に伴うソース変更
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iMax           ':最大ページ
    Public m_iDsp                       '// 一覧表示行数
    Public m_stxtNo         '処理番号
    Public m_stxtSEIMEI     '送信者の姓名
    Public m_sKenmei    
    Public m_sNaiyou    
    Public m_sKyokanCd  
    Public m_rs

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
    w_sMsgTitle="連絡事項登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_stxtNo        = request("txtNo")
    m_stxtSEIMEI    = request("txtSEIMEI")

	'T46には、ユーザーIDで登録されているので、教官コードでは該当しない
	'ユーザーIDで抽出、更新するように変更　2001/12/11 伊藤	
    'm_sKyokanCd     = session("KYOKAN_CD")
    m_sKyokanCd     = session("LOGIN_ID")

    m_iDsp          = C_PAGE_LINE

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

        '// 権限チェックに使用
        session("PRJ_No") = C_LEVEL_NOCHK

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '//データの取得、表示
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If
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

Function f_GetData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_sSQL
Dim w_Srs           '詳細用のレコードセット

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
        '//変数の値を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT DISTINCT"
        w_sSQL = w_sSQL & " T46_KENMEI,T46_NAIYO "
        w_sSQL = w_sSQL & "FROM "
        w_sSQL = w_sSQL & " T46_RENRAK "
        w_sSQL = w_sSQL & "WHERE "
        w_sSQL = w_sSQL & " T46_NO = '" & cInt(m_stxtNo) & "'"

        Set w_Srs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_Srs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

    '//取得した値を変数に代入
    m_sKenmei   = w_Srs("T46_KENMEI")
    m_sNaiyou   = w_Srs("T46_NAIYO")

    '//確認フラグを１にする。
        w_sSQL = ""
        w_sSQl = w_sSQL & " UPDATE T46_RENRAK SET "
        w_sSQL = w_sSQL & "     T46_KAKNIN = 1 "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "     T46_NO = " & cInt(m_stxtNo) & ""
        w_sSQL = w_sSQL & " AND T46_KYOKAN_CD = '" & m_sKyokanCd & "'"

        iRet = gf_ExecuteSQL(w_sSQL)
        If iRet <> 0 Then
            msMsg = Err.description
            f_GetData = 99
            Exit Do
        End If


    f_GetData = 0

    Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_sClass

%>
<HTML>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link rel="stylesheet" href="../common/style.css" type="text/css">
    <title>お知らせ</title>

    <!--#include file="../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [機能]  オンロード時
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_open(){

        //リスト情報をsubmit
        //document.frm.target = "<%=C_MAIN_FRAME%>_low" ;
        //document.frm.action = "top_lwr.asp";
        //document.frm.submit();

	opener.location.reload();
        window.focus();

    }

    //************************************************************
    //  [機能]  閉じるボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Close(){

        window.close()

    }
    //-->
    </SCRIPT>
</head>
<BODY onload="f_open()">
<center>
<FORM NAME="frm" action="post">
<br>
<% 
call gs_title("お　知　ら　せ","詳　細")
%>
<br>
<table width="300" border="1" CLASS="hyo">
    <TR>
        <TH CLASS="header" width="60">件名</TH>
        <TD CLASS="detail"><%=m_sKenmei%></TH>
    </TR>
    <TR>
        <TH CLASS="header">内容</TD>
        <TD CLASS="detail"><%=m_sNaiyou%></TD>
    </TR>
    <TR>
        <TH CLASS="header">送信者</TD>
        <TD CLASS="detail"><%=m_stxtSEIMEI%></TD>
    </TR>
</table>
<br>
<table border="0" width="350">
    <tr>
    <td valign="top" align="center">
    <input type="button" value="閉じる" class="button" onclick="javascript:f_Close()">
    </td>
    </tr>
</table>
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>