<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 使用教科書登録
' ﾌﾟﾛｸﾞﾗﾑID : web/WEB0320/delete.asp
' 機      能: 登録されている教科書の削除を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           w_sDelKyokasyoCD    :選択された教科書コード
'           txtPageCD       :表示済表示頁数（自分自身に引き渡す引数）
' 説      明:
'           ■初期表示
'               検索条件にかなう就職・進学先を表示
'           ■次へ、戻るボタンクリック時
'               指定した条件にかなう就職・進学を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/16 岩下　幸一郎
' 変      更: 2001/08/01 前田　智史
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_sMode
    Public  m_sNendo        ':年度
    Public  m_sNo           


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
    m_sNendo     = Request("txtNendo")              ':年度
    m_sNo        = Request("txtNo")

End Sub


'********************************************************************************
'*  [機能]  削除実行
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub S_delete()
Dim i
Dim w_iKyokasyoCD
Dim w_iRet              '// 戻り値
Dim w_sSQL              '// SQL文
Dim w_sWHERE            '// WHERE文
Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    w_iKyokasyoCD = ""
    w_sSQL = ""

    w_iKyokasyoCD = Request("deleteNO")

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="就職マスタ"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

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

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

        w_sSQL = w_sSQL & vbCrLf & " delete "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & " T47_KYOKASYO T47 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        '抽出条件の作成
        w_sSQL = w_sSQL & vbCrLf & " T47.T47_NO in (" & m_sNo & ")"
        w_sSQL = w_sSQL & vbCrLf & " and T47.T47_NENDO = " & m_sNendo
        
'response.write ("<BR>w_sSQL = " & w_sSQL)

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

    
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    Call gs_CloseDatabase()

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
<input type="hidden" name="txtMode" value="DELETE">
<input type="hidden" name="SKyokanCd1" value="<%=Request("SKyokanCd1")%>">
</form>

</center>

</body>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>