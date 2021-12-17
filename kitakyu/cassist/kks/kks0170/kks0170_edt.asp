<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 日毎出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0170/kks0170_edt.asp
' 機      能: 下ページ日毎出欠入力の登録、更新
'-------------------------------------------------------------------------
' 引      数: NENDO        '//処理年
'             KYOKAN_CD    '//教官CD
'             GAKUNEN      '//学年
'             CLASSNO      '//ｸﾗｽNo
'             cboDate      '//日付
' 変      数:
' 引      渡: cboDate      '//日付
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public m_iSyoriNen
    Public m_iKyokanCd
    Public m_iGakunen
    Public m_iClassNo
    Public m_sDate

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
    w_sMsgTitle="日毎出欠入力"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

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

        '//変数初期化
        Call s_ClearParam()

        '// MainﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

'//デバッグ
'Call s_DebugPrint()

        '// 日毎出欠登録
        w_iRet = f_AbsUpdate()
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

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_sDate     = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = trim(Request("NENDO"))
    m_iKyokanCd = trim(Request("KYOKAN_CD"))
    m_iGakunen  = trim(Request("GAKUNEN"))
    m_iClassNo  = trim(Request("CLASSNO"))
    m_sDate     = trim(Request("cboDate"))

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub
    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_sDate     = " & m_sDate     & "<br>"

End Sub

''********************************************************************************
''*  [機能]  ログイン者情報を取得
''*  [引数]  なし
''*  [戻値]  0:情報取得成功 99:失敗
''*  [説明]  
''********************************************************************************
'Function f_Get_UserInfo(p_UserName)
'Dim rs
'Dim w_sSQL
'
'    f_Get_UserInfo = 1
'    p_UserName = ""
'
'    Do
'        w_sSQL = ""
'        w_sSQL = w_sSQL & vbCrLf & " SELECT "
'        w_sSQL = w_sSQL & vbCrLf & "   M04_KYOKANMEI_SEI"
'        w_sSQL = w_sSQL & vbCrLf & " FROM M04_KYOKAN "
'        w_sSQL = w_sSQL & vbCrLf & " WHERE"
'        w_sSQL = w_sSQL & vbCrLf & "       M04_NENDO = " & cInt(m_iSyoriNen)
'        w_sSQL = w_sSQL & vbCrLf & "   AND M04_KYOKAN_CD = '" & m_iKyokanCd & "'"
'
'        iRet = gf_GetRecordset(rs, w_sSQL)
'        If iRet <> 0 Then
'            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
'            msMsg = Err.description
'            f_Get_UserInfo = 99
'            Exit Do
'        End If
'
'        If rs.EOF = False Then
'            p_UserName = rs("M04_KYOKANMEI_SEI")
'        End If
'
'        f_Get_UserInfo = 0
'        Exit Do
'    Loop
'
'    Call gf_closeObject(rs)
'
'End Function

'********************************************************************************
'*  [機能]  教科別出欠登録
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_AbsUpdate()

    Dim w_sSQL
    Dim w_Rs
    Dim w_sUserId
    Dim w_iKekka

    On Error Resume Next
    Err.Clear
    
    f_AbsUpdate = 1

    Do 

		'//ﾕｰｻﾞｰIDを取得
		w_sUserId = Session("LOGIN_ID")

        '//学籍Noを取得
        m_sGakuseiNo = split(replace(Request("GAKUSEKI_NO")," ",""),",")
        m_iGakusekiCnt = UBound(m_sGakuseiNo)

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

        '//クラスの人数分処理を実行
        For i=0 To m_iGakusekiCnt

            '//出欠CDを取得
            w_iKekka = trim(Request("hidKBN" & m_sGakuseiNo(i)))

            If w_iKekka = "---" Then
                '//出欠入力不可のときは更新処理をしない
            Else

                w_sSQL = ""
                w_sSQL = w_sSQL & vbCrLf & " SELECT "
                w_sSQL = w_sSQL & vbCrLf & "  T30_NENDO, "
                w_sSQL = w_sSQL & vbCrLf & "  T30_HIDUKE, "
                w_sSQL = w_sSQL & vbCrLf & "  T30_GAKUNEN, "
                w_sSQL = w_sSQL & vbCrLf & "  T30_CLASS, "
                w_sSQL = w_sSQL & vbCrLf & "  T30_GAKUSEKI_NO"
                w_sSQL = w_sSQL & vbCrLf & " FROM T30_KESSEKI"
                w_sSQL = w_sSQL & vbCrLf & " WHERE "
                w_sSQL = w_sSQL & vbCrLf & "      T30_NENDO=" & m_iSyoriNen
                w_sSQL = w_sSQL & vbCrLf & "  AND T30_HIDUKE='" & m_sDate & "' "
                w_sSQL = w_sSQL & vbCrLf & "  AND T30_GAKUNEN= " & m_iGakunen
                w_sSQL = w_sSQL & vbCrLf & "  AND T30_CLASS= " & m_iClassNo
                w_sSQL = w_sSQL & vbCrLf & "  AND T30_GAKUSEKI_NO='" & m_sGakuseiNo(i) & "'"

                iRet = gf_GetRecordset(rs, w_sSQL)
                If iRet <> 0 Then
                    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                    msMsg = Err.description
                    f_AbsUpdate = 99
                    Exit Do
                End If

                If rs.EOF Then

                    If w_iKekka <> "" and cstr(w_iKekka)<>"0" Then

                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T30_KESSEKI"
                        w_sSQL = w_sSQL & vbCrLf & "  ("
                        w_sSQL = w_sSQL & vbCrLf & "  T30_NENDO, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_HIDUKE, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_YOUBI_CD, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_GAKUNEN, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_CLASS, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_GAKUSEKI_NO, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_SYUKKETU_KBN, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_INS_DATE, "
                        w_sSQL = w_sSQL & vbCrLf & "  T30_INS_USER"
                        w_sSQL = w_sSQL & vbCrLf & "  )VALUES("
                        w_sSQL = w_sSQL & vbCrLf & "   "  & cInt(m_iSyoriNen) & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & Trim(m_sDate)     & "',"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & Weekday(m_sDate)  & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   "  & cInt(m_iGakunen)  & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "   "  & cInt(m_iClassNo)  & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & m_sGakuseiNo(i)   & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   "  & Trim(w_iKekka)    & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & Date()            & "',"
                        w_sSQL = w_sSQL & vbCrLf & "  '"  & w_sUserId       & "'"
                        w_sSQL = w_sSQL & vbCrLf & "  )"

                        iRet = gf_ExecuteSQL(w_sSQL)
'response.write w_sSQL & "<br>"
'response.write "INSERT iRet = " & iRet & "<br>"
                        If iRet <> 0 Then
                            '//ﾛｰﾙﾊﾞｯｸ
                            Call gs_RollbackTrans()
                            msMsg = Err.description
                            f_AbsUpdate = 99
                            Exit Do
                        End If

                    End If

                Else

                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " UPDATE T30_KESSEKI SET "
                    w_sSQL = w_sSQL & vbCrLf & "  T30_SYUKKETU_KBN = "  & Trim(w_iKekka)    & " ," 
                    w_sSQL = w_sSQL & vbCrLf & "  T30_UPD_DATE =    '"  & Date()            & "',"
                    w_sSQL = w_sSQL & vbCrLf & "  T30_UPD_USER =    '"  & w_sUserId         & "'"
                    w_sSQL = w_sSQL & vbCrLf & " WHERE "
                    w_sSQL = w_sSQL & vbCrLf & "      T30_NENDO="    & m_iSyoriNen
                    w_sSQL = w_sSQL & vbCrLf & "  AND T30_HIDUKE='"  & m_sDate & "' "
                    w_sSQL = w_sSQL & vbCrLf & "  AND T30_GAKUNEN= " & m_iGakunen
                    w_sSQL = w_sSQL & vbCrLf & "  AND T30_CLASS= "   & m_iClassNo
                    w_sSQL = w_sSQL & vbCrLf & "  AND T30_GAKUSEKI_NO='" & m_sGakuseiNo(i) & "'"

                    iRet = gf_ExecuteSQL(w_sSQL)

'response.write w_sSQL & "<br>"
'response.write "UPDATE iRet = " & iRet & "<br>"

                    If iRet <> 0 Then
                        '//ﾛｰﾙﾊﾞｯｸ
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_AbsUpdate = 99
                        Exit Do
                    End If

                    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
                    Call gf_closeObject(rs)
                End If
            End If

        Next

        '//ｺﾐｯﾄ
        Call gs_CommitTrans()

        '//正常終了
        f_AbsUpdate = 0
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
    <title>日毎出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
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
		alert("<%= C_TOUROKU_OK_MSG %>");

        //リスト情報をsubmit
        document.frm.target = "main";
        document.frm.action = "./kks0170_bottom.asp"
        document.frm.submit();
        return;

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">

    <form name="frm" method="post">
    <input type="hidden" name="cboDate"   value="<%=Request("cboDate")%>">
    </form>

    </center>
    </body>
    </html>
<%
End Sub
%>