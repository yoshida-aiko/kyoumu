<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 時間割交換連絡
' ﾌﾟﾛｸﾞﾗﾑID : web/web0310/web0310_edt.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
'            モード         ＞      txtMode
'                                   新規 = NEW
'                                   更新 = UPDATE
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 前田
' 変      更: 2001/09/03 伊藤公子 教官以外も利用できるように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iMax           ':最大ページ
    Public m_iDsp           '// 一覧表示行数
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官ｺｰﾄﾞ
    Public m_stxtMode       'モード
    Public m_sNaiyou        '内容
    Public m_sKaisibi       '開始日
    Public m_sSyuryoubi     '完了日
    Public m_sJoukin        '常勤区分
    Public m_sGakka         '学科区分
    Public m_sKkanKBN       '教官区分
    Public m_sKkeiKBN       '教科系列区分
    Public m_stxtNo         '処理番号
    Public m_rs
    Public m_sListCd
    Dim    m_rCnt           '//レコード件数

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
    m_stxtMode = request("txtMode")

    m_sNaiyou   = request("Naiyou")
    m_iNendo    = request("txtNendo")
    m_sKaisibi  = request("Kaisibi")
    m_sSyuryoubi= request("Syuryoubi")
    m_sKyokanCd = request("txtKyokanCd")
    m_stxtNo    = request("txtNo")
    m_sListCd   = request("KCD")
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
            Case "NEW2"
            '//データの取得
            w_iRet = f_insertData()
            If w_iRet <> 0 Then
                'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
                m_bErrFlg = True
                Exit Do
            End If
            Call showPage()
            Exit Do
            
            Case "UPD","UPD2"
            '//データの取得、表示
            w_iRet = f_updateData()
            If w_iRet <> 0 Then
                'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
                m_bErrFlg = True
                Exit Do
            End If
            Call showPage()
            Exit Do

        End Select
        '// ページを表示
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

Function f_insertData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_sSQL
Dim w_rs
Dim w_sKyokanList
Dim w_sListCd
Dim w_sKyokanCd
Dim w_iMaxNo
Dim i

    On Error Resume Next
    Err.Clear
    f_insertData = 1

    Do

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

        '//Noの最大値を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & "  MAX(T52_NO) AS MAXNO "
        w_sSQL = w_sSQL & "FROM "
        w_sSQL = w_sSQL & "  T52_JYUGYO_HENKO "

        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

        If IsNull(w_rs("MAXNO")) Then
            w_iMaxNo = 1
        Else
            w_iMaxNo = cInt(w_rs("MAXNO")) + 1
        End If

        '//送信先選択画面でチェックされたデータを配列で取得
        w_sKyokanList = split(replace(m_sListCd," ",""),",")

        iMax = UBound(w_sKyokanList)

'---------2001/09/03 ito
'        m_sSQL = ""
'        m_sSQL = m_sSQL & "SELECT "
'        m_sSQL = m_sSQL & "  M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
'        m_sSQL = m_sSQL & "FROM "
'        m_sSQL = m_sSQL & "  M04_KYOKAN "
'        m_sSQL = m_sSQL & "WHERE "
'        m_sSQL = m_sSQL & "  M04_KYOKAN_CD IN (" & Trim(m_sListCd) & ") "
'
'        Set m_rs = Server.CreateObject("ADODB.Recordset")
'        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
'        If w_iRet <> 0 Then
'            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
'            m_bErrFlg = True
'            Exit Do 
'        End If

    For i=0 to iMax
        w_sKyokanCd = w_sKyokanList(i)

        '//学年･クラスのデータ
        m_sSQL = ""
        m_sSQL = m_sSQL & "INSERT INTO T52_JYUGYO_HENKO " 
        m_sSQL = m_sSQL & " ( " 
        m_sSQL = m_sSQL & "  T52_NO,T52_KYOKAN_CD,T52_NAIYO,T52_KAISI, " 
        m_sSQL = m_sSQL & "  T52_SYURYO,T52_INS_DATE,T52_INS_USER " 
        m_sSQL = m_sSQL & ") " 
        m_sSQL = m_sSQL & " VALUES " 
        m_sSQL = m_sSQL & "( " 
        m_sSQL = m_sSQL & " '" & cInt(w_iMaxNo) & "', " 
        m_sSQL = m_sSQL & "'" & Trim(w_sKyokanCd) & "', " 
        m_sSQL = m_sSQL & "'" & Trim(m_sNaiyou) & "', " 
        m_sSQL = m_sSQL & "'" & gf_YYYY_MM_DD(Trim(m_sKaisibi),"/") & "', " 
        m_sSQL = m_sSQL & "'" & gf_YYYY_MM_DD(Trim(m_sSyuryoubi),"/") & "', " 
        m_sSQL = m_sSQL & "'" & gf_YYYY_MM_DD(date(),"/") & "', " 
        m_sSQL = m_sSQL & "'" & Session("LOGIN_ID") & "' " 
        m_sSQL = m_sSQL & "   )"

        iRet = gf_ExecuteSQL(m_sSQL)
        If iRet <> 0 Then
            '//ﾛｰﾙﾊﾞｯｸ
            Call gs_RollbackTrans()
            msMsg = Err.description
            f_insertData = 99
            Exit Do
        End If
    Next

    '//ｺﾐｯﾄ
    Call gs_CommitTrans()

    f_insertData = 0

    Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Function f_updateData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_sSQL
Dim w_Srs           '削除用のレコードセット
Dim w_Brs           '以前のレコードセット
Dim w_Nrs           '現在のレコードセット
Dim w_sKyokanList
Dim w_sKyokanCd
Dim w_sUpdFlg
Dim i

    On Error Resume Next
    Err.Clear
    f_updateData = 1

    Do

        Call gs_BeginTrans()

        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & "  T52_NO,T52_KYOKAN_CD "
        w_sSQL = w_sSQL & "FROM "
        w_sSQL = w_sSQL & "  T52_JYUGYO_HENKO "
        w_sSQL = w_sSQL & "WHERE "
        w_sSQL = w_sSQL & "  T52_NO = '" & cInt(m_stxtNo) & "' "

        Set w_Brs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_Brs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

        '//送信先選択画面でチェックされたデータを配列で取得
        w_sKyokanList = split(replace(m_sListCd," ",""),",")

        iMax = UBound(w_sKyokanList)

        '//テーブルに書き込む
        For i=0 to iMax
            w_sKyokanCd = w_sKyokanList(i)

            w_Brs.MoveFirst
            Do Until w_Brs.EOF

                UpdFlg = False
                If w_Brs("T52_KYOKAN_CD") = Trim(w_sKyokanCd) Then

                    '//T52_JYUGYO_HENKOにすでに生徒情報がある場合はUPDATE
                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " UPDATE T52_JYUGYO_HENKO SET "
                    w_sSQL = w_sSQL & vbCrLf & "   T52_NAIYO = '"  & Trim(m_sNaiyou) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "   T52_KAISI = '"  & gf_YYYY_MM_DD(Trim(m_sKaisibi),"/") & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "   T52_SYURYO = '"  & gf_YYYY_MM_DD(Trim(m_sSyuryoubi),"/") & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "   T52_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")            & "',"
                    w_sSQL = w_sSQL & vbCrLf & "   T52_UPD_USER = '"    & Session("LOGIN_ID") & "'"
                    w_sSQL = w_sSQL & vbCrLf & " WHERE "
                    w_sSQL = w_sSQL & vbCrLf & "        T52_NO = '" & m_stxtNo & "'  "
                    w_sSQL = w_sSQL & vbCrLf & "    AND T52_KYOKAN_CD = '" & Trim(w_sKyokanList(i)) & "' "

                    iRet = gf_ExecuteSQL(w_sSQL)
                    If iRet <> 0 Then
                        '//ﾛｰﾙﾊﾞｯｸ
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_updateData = 99
                        Exit Do
                    End If
                UpdFlg = True
                Exit Do
                End If 
                w_Brs.MoveNext
            Loop

                If UpdFlg = False Then

                    '//T52_JYUGYO_HENKOに生徒情報がない場合INSERT
                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T52_JYUGYO_HENKO  "
                    w_sSQL = w_sSQL & vbCrLf & "   ("
                    w_sSQL = w_sSQL & vbCrLf & "   T52_NO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T52_KYOKAN_CD, "
                    w_sSQL = w_sSQL & vbCrLf & "   T52_NAIYO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T52_KAISI, "
                    w_sSQL = w_sSQL & vbCrLf & "   T52_SYURYO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T52_INS_DATE, "
                    w_sSQL = w_sSQL & vbCrLf & "   T52_INS_USER "
                    w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
                    w_sSQL = w_sSQL & vbCrLf & "    '" & cInt(m_stxtNo) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & Trim(w_sKyokanList(i)) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & Trim(m_sNaiyou) & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & gf_YYYY_MM_DD(Trim(m_sKaisibi),"/") & "',"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & gf_YYYY_MM_DD(Trim(m_sSyuryoubi),"/") & "' ,"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & gf_YYYY_MM_DD(date(),"/") & "',"
                    w_sSQL = w_sSQL & vbCrLf & "    '" & Session("LOGIN_ID") & "' "
                    w_sSQL = w_sSQL & vbCrLf & "   )"

                    iRet = gf_ExecuteSQL(w_sSQL)
                    If iRet <> 0 Then
                        '//ﾛｰﾙﾊﾞｯｸ
                        Call gs_RollbackTrans()
                        msMsg = Err.description
                        f_updateData = 99
                        Exit For
                    End If
                End If
        Next

    '//ｺﾐｯﾄ
    Call gs_CommitTrans()

    '//削除する
    Call gs_BeginTrans()

            w_sSQL = ""
            w_sSQL = w_sSQL & "SELECT "
            w_sSQL = w_sSQL & "  T52_NO,T52_KYOKAN_CD "
            w_sSQL = w_sSQL & "FROM "
            w_sSQL = w_sSQL & "  T52_JYUGYO_HENKO "
            w_sSQL = w_sSQL & "WHERE "
            w_sSQL = w_sSQL & "  T52_NO = '" & cInt(m_stxtNo) & "' "

            Set w_Srs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(w_Srs, w_sSQL,m_iDsp)
            If w_iRet <> 0 Then
                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                m_bErrFlg = True
                Exit Do 
            End If
    
        w_Srs.MoveFirst
        Do Until w_Srs.EOF
    
            For i=0 to iMax
                UpdFlg = False
                w_sKyokanCd = w_sKyokanList(i)
    
                If w_Srs("T52_KYOKAN_CD") = w_sKyokanList(i) Then
                    UpdFlg = True
                    Exit For
                End If
            Next
            If UpdFlg = False Then
    
                w_sSQL = ""
                w_sSQL = w_sSQL & vbCrLf & " DELETE FROM T52_JYUGYO_HENKO  "
                w_sSQL = w_sSQL & vbCrLf & " WHERE "
                w_sSQL = w_sSQL & vbCrLf & "     T52_NO = '" & cInt(m_stxtNo) & "' "
                w_sSQL = w_sSQL & vbCrLf & " AND T52_KYOKAN_CD = '" & w_Srs("T52_KYOKAN_CD") & "' "

                iRet = gf_ExecuteSQL(w_sSQL)
                If iRet <> 0 Then
                    '//ﾛｰﾙﾊﾞｯｸ
                    Call gs_RollbackTrans()
                    msMsg = Err.description
                    f_updateData = 99
                    Exit Do
                End If
            End If
            w_Srs.MoveNext
        Loop

    '//ｺﾐｯﾄ
    Call gs_CommitTrans()

    f_updateData = 0

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

        location.href = "./default.asp"
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