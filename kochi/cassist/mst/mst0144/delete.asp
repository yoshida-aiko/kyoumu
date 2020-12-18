<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0133/main.asp
' 機      能: 就職先マスタの削除を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           txtSinroKBN     :進路先コード
'           txtSingakuCd        :進学コード
'           txtSinroName        :就職先名称（一部）
'           txtPageCD       :表示済表示頁数（自分自身から受け取る引数）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           txtRenrakusakiCD    :選択された連絡先コード
'           txtPageCD       :表示済表示頁数（自分自身に引き渡す引数）
' 説      明:
'           ■初期表示
'               検索条件にかなう就職・進学先を表示
'           ■次へ、戻るボタンクリック時
'               指定した条件にかなう就職・進学を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/29 岩下　幸一郎
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_sSinroCD      ':進路先コード
    Public  m_sSingakuCd        ':進学コード
    Public  m_sSinroCD2     ':進路先コード
    Public  m_sSingakuCd2       ':進学コード
    Public  m_sSyusyokuName     ':就職先名称（一部）
    Public  m_sPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_skubun
    Public  m_Rs            'recordset
    Public  m_iDisp         ':表示件数の最大値をとる
    Public  m_sRenrakusakiCD
    Public  m_iNendo        ':年度


    'ページ関係
    Public  m_iMax          ':最大ページ
    Public  m_iDsp                      '// 一覧表示行数

'   call gs_viewForm(request.form)
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

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '// 削除の実行
        Call S_delete()

        '// ページを表示
        Call showPage()

End Sub


'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_sMode      = Request("txtMode")
    m_sRenrakusakiCD = Request("txtRenrakusakiCD")      ':連絡先コード

    m_sSinroCD2 = Request("txtSinroCD2")            ':進路コード
    'コンボ未選択時
    If m_sSinroCD2="@@@" Then
        m_sSinroCD2=""
    End If

    m_sSingakuCD2 = Request("txtSingakuCD2")        ':進学コード
    'コンボ未選択時
    If m_sSingakuCD2="@@@" Then
        m_sSingakuCD2=""
    End If

    m_sSyusyokuName = Request("txtSyusyokuName")        ':就職先名称（一部）


    '// BLANKの場合は行数ｸﾘｱ
    If Request("txtMode") = "Search" Then
        m_sPageCD = 1
    Else
        m_sPageCD = INT(Request("txtPageCD"))       ':表示済表示頁数（自分自身から受け取る引数）
    End If

    If m_sSinroCD = "1" Then                ':ヘッダーの区分名称変更
        m_skubun = "進学区分"
    else
        m_skubun = "進路区分"
    End If

    m_iDisp = Request("txtDisp")                ':ページ件数最大値

    m_iNendo = Request("txtNendo")              ':年度

End Sub


Sub S_delete()
'********************************************************************************
'*  [機能]  削除実行
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_slink
Dim w_iCnt
Dim i
Dim w_iSinrosakiCD
Dim w_iSinrosakiCDs
Dim w_iSinrosakiCDe
Dim w_iSinrosakiCDStr

w_slink = "　"

w_iCnt = 0

w_iSinrosakiCD = ""
w_sSQL = ""

w_iSinrosakiCD = Request("deleteNO")
w_iSinrosakiCDs = Split(w_iSinrosakiCD)
for each w_iSinrosakiCDe In w_iSinrosakiCDs
    w_iSinrosakiCDe = "'" & Replace(w_iSinrosakiCDe,",","") & "'"
    w_iSinrosakiCDStr = w_iSinrosakiCDStr & w_iSinrosakiCDe
next
    w_iSinrosakiCDStr = Replace(w_iSinrosakiCDStr,"''","','")
    'response.write w_iSinrosakiCDStr


    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWHERE            '// WHERE文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="進路先情報登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    Do
            '// ﾃﾞｰﾀﾍﾞｰｽ接続
            w_iRet = gf_OpenDatabase()
           If w_iRet <> 0 Then
                'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
                m_bErrFlg = True
                Exit Do
                m_sErrMsg = "データベースとの接続に失敗しました。"
            End If

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

        w_sSQL = w_sSQL & vbCrLf & " delete "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & " M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        '抽出条件の作成
        w_sSQL = w_sSQL & vbCrLf & " M32.M32_SINRO_CD in (" & w_iSinrosakiCDStr & ")"

'response.write w_sSQL
        if gf_ExecuteSQL(w_sSQL) <> 0 then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            '//ﾛｰﾙﾊﾞｯｸ
            Call gs_RollbackTrans()
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If


    'すべての処理が正しく終了
    Exit do
    Loop
    

    '//ｺﾐｯﾄ
    Call gs_CommitTrans()

    '// 終了処理
    Call gs_CloseDatabase()
    
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If


    'LABEL_showPage_OPTION_END
End sub

'********************************************************************************
'*  [機能]  HTMLを表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
%>





<html>
<link rel=stylesheet href=../common/style.css type=text/css>
    <head>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    function gonext() {
    alert('<%=C_SAKUJYO_OK_MSG%>');
            document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>
<body onLoad="gonext();">

<center>
<form action="default.asp" name="frm" target=<%=C_MAIN_FRAME%> method="post">
<img src="../../image/sp.gif" width="20" height="1">
<input type="hidden" name="txtMode" value="search">
<input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
</form>

</center>

</body>

</html>





<%
    '---------- HTML END   ----------
End Sub

Sub showPage_NoData()
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
    </head>

    <body>

    <center>
		<br><br><br>
		<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
    <input type="button" value="戻　る" onclick="javascript:history.back()">
    </center>

    </body>

    </html>
<%
End Sub
%>