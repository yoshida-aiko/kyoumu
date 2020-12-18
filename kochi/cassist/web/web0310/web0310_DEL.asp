<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 時間割交換連絡
' ﾌﾟﾛｸﾞﾗﾑID : web/web0310/web0310_DEL.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 前田
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_iMax          ':最大ページ
    Public  m_iDsp          '// 一覧表示行数
    Public  m_rs
    Public  m_stxtMode      'モード
    Dim     m_iNendo
    Dim     m_sKyokanCd
    Dim     m_sNo
    Dim     m_sDelNo
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
    w_sMsgTitle="時間割交換連絡"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sNo   = request("Delchk")
    m_sDelNo    = request("txtDelNo")
    m_stxtMode = request("txtMode")
    m_iDsp = C_PAGE_LINE

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

        Select Case m_stxtMode

            Case "","DELKNIN"
                '//リストの一覧データの詳細取得
                w_iRet = f_GetData()
                If w_iRet <> 0 Then
                    'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
                    m_bErrFlg = True
                    Exit Do
                End If
                '// ページを表示
                Call showPage()
                Exit Do

            Case "Delete"

                w_iRet = f_DeleteData()
                If w_iRet <> 0 Then
                    'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
                    m_bErrFlg = True
                    Exit Do
                End If
                '// ページを表示
                Call DEL_showPage()
                Exit Do
        End Select

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

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do

        '//リストの表示
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT DISTINCT "
        m_sSQL = m_sSQL & "     T52_NO,T52_NAIYO "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T52_JYUGYO_HENKO "
        m_sSQL = m_sSQL & " WHERE "
        If m_stxtMode = "" Then
            m_sSQL = m_sSQL & "     T52_NO IN (" & Trim(m_sNo) & ") "
        ElseIf m_stxtMode = "DELKNIN" Then
            m_sSQL = m_sSQL & "     T52_NO = '" & Trim(m_sDelNo) & "' "
        End If

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
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

Function f_DeleteData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_DeleteData = 1

    Do
        '//リストの表示
        m_sSQL = ""
        m_sSQL = m_sSQL & " DELETE FROM T52_JYUGYO_HENKO "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T52_NO IN (" & Trim(m_sDelNo) & ") "

        iRet = gf_ExecuteSQL(m_sSQL)
        If iRet <> 0 Then
            msMsg = Err.description
            f_DeleteData = 99
            Exit Do
        End If

        f_DeleteData = 0

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

    On Error Resume Next
    Err.Clear
%>

<html>
    <head>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  削除ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_delete(){

        if (!confirm("<%=C_SAKUJYO_KAKUNIN%>")) {
           return ;
        }
        document.frm.action="web0310_DEL.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "Delete";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
</head>
<body>

<center>

<%call gs_title("時間割交換連絡","削　除")%>
<br>
削　除　内　容
<br><br>
    <table border="1" class=hyo width="80%">
<form name="frm" action="post">

    <tr>
    <th class=header>処理番号</th>
    <th class=header>内　容</th>
    </tr>

<%
    m_rs.MoveFirst
    Do Until m_rs.EOF
%>
    <tr>
    <td align="center" class=detail width="20%"><%=m_rs("T52_NO")%></td>
    <td class=detail width="80%"><%=m_rs("T52_NAIYO")%></td>
    </tr>
<%
    m_rs.MoveNext
    Loop
 %>

    </table>
<br>
以上の内容を削除します。
<br><br>
<table border="0">
<tr>
<td align=left>
<input type="button" class=button value="　削　除　" Onclick="javascript:f_delete()">
</td>
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_iNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">

<%
    If m_stxtMode = "" Then
%>
        <INPUT TYPE=HIDDEN  NAME=txtDelNo       value="<%=m_sNo%>">
<%
    ElseIf m_stxtMode = "DELKNIN" Then
%>
        <INPUT TYPE=HIDDEN  NAME=txtDelNo       value="<%=m_sDelNo%>">
<%
        End If
%>

</form>
<form action="default.asp" target="<%=C_MAIN_FRAME%>" method="post">
<td align=right>
<input type="submit" class=button value="キャンセル">
</td>
</form>
</tr>
</table>

</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub

Sub DEL_showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>

    <html>
    <head>
    <title>時間割交換連絡</title>
    <link rel=stylesheet href=../../font.css type=text/css>

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

        location.href = "default.asp"
        return;
    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>