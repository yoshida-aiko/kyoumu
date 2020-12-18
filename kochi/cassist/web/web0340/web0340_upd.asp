<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 個人履修選択科目決定
' ﾌﾟﾛｸﾞﾗﾑID : web/web0340/web0340_edt.asp
' 機      能: 下ページ 個人履修選択科目決定の更新
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
' 変      数:
' 引      渡: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 前田 智史
' 変      更: 2001/08/28 伊藤公子 ヘッダ部切り離し対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
    Const DebugPrint = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Dim     m_iNendo        '//年度
    Dim     m_sKyokanCd     '//教官CD
    Dim     n_Max           '//最大数
    Dim     k_Max           '//最大数

    Public  m_iMax          '最大ページ
    Public  m_iDsp          '一覧表示行数

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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="行事出欠入力"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
'--------2001/08/28 ito 
'    m_sgakuNo   = Trim(request("gakuNo"))
'    m_skamokuCd = Trim(request("kamokuCd"))
    n_Max       = request("n_Max")
    k_Max       = request("k_Max")
    m_iDsp = C_PAGE_LINE

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

        '// 個人履修選択科目決定
        w_iRet = f_Update()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '// ページを表示
        Call showPage()

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gs_CloseDatabase()

End Sub

Function f_Update()
'********************************************************************************
'*  [機能]  ヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Dim w_sGakuNo
Dim w_sKamokuCd
Dim n
Dim k



    On Error Resume Next
    Err.Clear
    
    f_Update = 1

    Do 

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

        For k=1 To k_Max                '//学生Noを回す
            For n=1 To n_Max            '//科目CDを回す

                w_sGakuNo = request("GakuNo"& k)
                w_sKamokuCd = request("KamokuCd"& n)

                If request("MAE"& k &"_"& n) <> "○" AND request("ATO"& k &"_"& n) = "○" Then

                    w_sSQL = ""
                    w_sSQL = w_sSQL & " UPDATE T16_RISYU_KOJIN SET "
                    w_sSQL = w_sSQL & "   T16_SELECT_FLG = " & C_SENTAKU_YES & ", "
                    w_sSQL = w_sSQL & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "',"
                    w_sSQL = w_sSQL & "   T16_UPD_USER = '" & Session("LOGIN_ID") & "' "
                    w_sSQL = w_sSQL & " WHERE "
                    w_sSQL = w_sSQL & "   T16_NENDO = " & m_iNendo & " "
                    w_sSQL = w_sSQL & " AND T16_GAKUSEI_NO = '" & w_sGakuNo & "' "
                    w_sSQL = w_sSQL & " AND T16_KAMOKU_CD = '" & w_sKamokuCd & "' "

                    iRet = gf_ExecuteSQL(w_sSQL)
                    If iRet <> 0 Then
                        '//ﾛｰﾙﾊﾞｯｸ
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_Update = 99
                        Exit For
                    End If

                ElseIf request("MAE"& k &"_"& n) = "○" AND request("ATO"& k &"_"& n) <> "○" Then

                    w_sSQL = ""
                    w_sSQL = w_sSQL & " UPDATE T16_RISYU_KOJIN SET "
                    w_sSQL = w_sSQL & "   T16_SELECT_FLG = " & C_SENTAKU_NO & ", "
                    w_sSQL = w_sSQL & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "',"
                    w_sSQL = w_sSQL & "   T16_UPD_USER = '" & Session("LOGIN_ID") & "' "
                    w_sSQL = w_sSQL & " WHERE "
                    w_sSQL = w_sSQL & "   T16_NENDO = " & m_iNendo & " "
                    w_sSQL = w_sSQL & " AND T16_GAKUSEI_NO = '" & w_sGakuNo & "' "
                    w_sSQL = w_sSQL & " AND T16_KAMOKU_CD = '" & w_sKamokuCd & "' "

                    iRet = gf_ExecuteSQL(w_sSQL)
                    If iRet <> 0 Then
                        '//ﾛｰﾙﾊﾞｯｸ
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_Update = 99
                        Exit For
                    End If
                End If
            Next
        Next

        '//ｺﾐｯﾄ
        Call gs_CommitTrans()

        '//正常終了
        f_Update = 0
        Exit Do
    Loop

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
    <title>個人履修選択科目決定</title>
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

		alert("<%=C_TOUROKU_OK_MSG%>");

//        document.frm.action="default2.asp";
//        document.frm.target="main";
//        document.frm.submit();

	    document.frm.target = "main";
	    document.frm.action = "./web0340_main.asp"
	    document.frm.submit();
	    return;
    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type="hidden" name="txtNendo"    value="<%=Request("txtNendo")%>">
	<input type="hidden" name="txtKyokanCd" value="<%=Request("txtKyokanCd")%>">
	<input type="hidden" name="txtGakunen"  value="<%=Request("txtGakunen")%>">
	<input type="hidden" name="txtClass"    value="<%=Request("txtClass")%>">
	<input type="hidden" name="txtKBN"      value="<%=Request("txtKBN")%>">
	<input type="hidden" name="txtGRP"      value="<%=Request("txtGRP")%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

