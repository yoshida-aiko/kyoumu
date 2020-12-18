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
' 変      更: 
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
    Public m_sKyokanCd  
    Public m_iNendo 
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

    m_sKyokanCd     = session("KYOKAN_CD")
    m_iNendo        = session("NENDO")
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
	Dim w_user

    On Error Resume Next
    Err.Clear
    f_GetData = 1

	w_user = session("LOGIN_ID")
	'ユーザが教官であれば、教官CDを代入
	If m_sKyokanCd <> "" then w_user = m_sKyokanCd

    Do
        '//変数の値を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     A.T52_NAIYO,A.T52_INS_DATE,B.M10_USER_NAME "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T52_JYUGYO_HENKO A,M10_USER B "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "     A.T52_KYOKAN_CD = '" & w_user & "' AND "
        w_sSQL = w_sSQL & "     A.T52_INS_USER = B.M10_USER_ID AND "
        w_sSQL = w_sSQL & "     B.M10_NENDO = " & m_iNendo & " "

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
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
<link rel=stylesheet href="../common/style.css" type=text/css>
    <title>時間割変更連絡</title>

    <!--#include file="../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
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

<BODY>
<center>
<FORM NAME="frm" action="post">

<% call gs_title("時間割変更連絡","詳　細") %>
<br>
<table width="400" border=1 CLASS="hyo">
    <TR>
        <TH CLASS="header" width=65%>送　信　者</TD>
        <TH CLASS="header" width=35%>登　録　日</TD>
    </TR>
    <TR>
        <TH CLASS="header" width=100% colspan=2>　内　容　</TD>
    </TR>
	<%
	    m_rs.MoveFirst
	    Do Until m_rs.EOF
			%>
		    <TR>
		        <TD CLASS="CELL1" ><%=m_rs("M10_USER_NAME")%></TD>
		        <TD CLASS="CELL1" ><%=m_rs("T52_INS_DATE")%></TD>
		    </TR>
		    <TR>
		        <TD CLASS="CELL2" colspan=2><%=m_rs("T52_NAIYO")%></TD>
		    </TR>
			<%
	    m_rs.MoveNext
    Loop
	%>
    </TABLE>

	<br>
    <table border="0" width="350">
	    <tr>
		    <td valign="top" align="center">
		    <input type="button" value="閉じる" class=button onclick="javascript:f_Close()">
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