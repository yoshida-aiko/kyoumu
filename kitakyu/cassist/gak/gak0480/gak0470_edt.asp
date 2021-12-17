<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 各種委員登録
' ﾌﾟﾛｸﾞﾗﾑID : gah/gak0470/gaku0470_edt.asp
' 機      能: 下ページ 各種委員登録の登録、更新
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
' 作      成: 2001/07/02 前田
' 変      更: 
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
    Public  m_sGakunen
    Public  m_sClassNo
    Dim     m_iNendo
    Dim     m_sKyokanCd

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
    w_sMsgTitle="各種委員登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
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

        '// 行事出欠登録
        w_iRet = f_AbsUpdate()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

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

Function f_AbsUpdate()
'********************************************************************************
'*  [機能]  ヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************

    Dim w_sUserName
    Dim w_sSQL
    Dim w_Rs
    Dim i
    i = 1

    On Error Resume Next
    Err.Clear
    
    f_AbsUpdate = 1

    w_iMax = request("HIDMAX")
'	w_iGAKKI = 

    Do 

'        '// ログイン者名称を取得
'        w_iRet = f_Get_UserInfo(w_sUserName)
'        If w_iRet <> 0 Then
'            Exit Do
'        End If

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

            '//行事数分処理を実行
            For i=1 to w_iMax

                If request("gakuNo" & i ) <> "" Then

                    If request("Before" & i ) = "" Then

                        '//T06_GAKU_IINに生徒情報がない場合INSERT
                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T06_GAKU_IIN  "
                        w_sSQL = w_sSQL & vbCrLf & "   ("
                        w_sSQL = w_sSQL & vbCrLf & "   T06_NENDO, "
						w_sSQL = w_sSQL & vbCrLf & "   T06_GAKKI_KBN, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_GAKUSEI_NO, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_DAIBUN_CD, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_SYOBUN_CD, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_INS_DATE, "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_INS_USER "
                        w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
                        w_sSQL = w_sSQL & vbCrLf & "    '"  & cInt(m_iNendo) & "' ,"
						w_sSQL = w_sSQL & vbCrLf & "   " & request("GAKKI") & " ,"
                        w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(request("gakuNo" & i )) & "',"
                        w_sSQL = w_sSQL & vbCrLf & "    '"  & cInt(request("iinDai" & i )) & "' ,"
                        w_sSQL = w_sSQL & vbCrLf & "    '"  & cInt(request("iinSyo" & i )) & "' ,"
                        w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_YYYY_MM_DD(date(),"/") & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   '"  & Session("LOGIN_ID") & "' "
                        w_sSQL = w_sSQL & vbCrLf & "   )"

                        iRet = gf_ExecuteSQL(w_sSQL)
                        If iRet <> 0 Then
                            '//ﾛｰﾙﾊﾞｯｸ
                            Call gs_RollbackTrans()
                            msMsg = Err.description
                            f_AbsUpdate = 99
                            Exit Do
                        End If

                    ElseIf request("gakuNo" & i ) <> request("Before" & i ) Then

                        '//T06_GAKU_IINにすでに生徒情報がある場合はUPDATE
                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " UPDATE T06_GAKU_IIN SET "
                        w_sSQL = w_sSQL & vbCrLf & "   T06_GAKUSEI_NO = '"  & Trim(request("gakuNo" & i))    & "' ,"
                        w_sSQL = w_sSQL & vbCrLf & "   T06_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")            & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   T06_UPD_USER = '"    & Session("LOGIN_ID")       & "'"
                        w_sSQL = w_sSQL & vbCrLf & " WHERE "
                        w_sSQL = w_sSQL & vbCrLf & "        T06_NENDO = '" & m_iNendo & "'  "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_GAKUSEI_NO = '" & Trim(request("Before" & i )) & "' "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_DAIBUN_CD = '" & cInt(request("iinDai" & i )) & "' "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_SYOBUN_CD = '" & cInt(request("iinSyo" & i )) & "' "
						w_sSQL = w_sSQL & vbCrLf & "	AND T06_GAKKI_KBN = " & request("GAKKI") & " "

                        iRet = gf_ExecuteSQL(w_sSQL)
                        If iRet <> 0 Then
                            '//ﾛｰﾙﾊﾞｯｸ
                            Call gs_RollbackTrans()
                            msMsg = Err.description
                            f_AbsUpdate = 99
                            Exit Do
                        End If
                    End If
				Else
                    If request("Before" & i ) <> "" Then

                        '//T06_GAKU_IINにすでに生徒情報があり、現在表示上で空白の場合DELETE
                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " DELETE FROM T06_GAKU_IIN "
                        w_sSQL = w_sSQL & vbCrLf & " WHERE "
                        w_sSQL = w_sSQL & vbCrLf & "        T06_NENDO = '" & m_iNendo & "'  "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_GAKUSEI_NO = '" & Trim(request("Before" & i )) & "' "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_DAIBUN_CD = '" & cInt(request("iinDai" & i )) & "' "
                        w_sSQL = w_sSQL & vbCrLf & "    AND T06_SYOBUN_CD = '" & cInt(request("iinSyo" & i )) & "' "
						w_sSQL = w_sSQL & vbCrLf & "	AND T06_GAKKI_KBN = " & request("GAKKI") & " "

                        iRet = gf_ExecuteSQL(w_sSQL)
                        If iRet <> 0 Then
                            '//ﾛｰﾙﾊﾞｯｸ
                            Call gs_RollbackTrans()
                            msMsg = Err.description
                            f_AbsUpdate = 99
                            Exit Do
                        End If
                    End If
                End If

                    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
                    Call gf_closeObject(rs)
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
    <title>各種委員登録</title>
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

		alert("<%= C_TOUROKU_OK_MSG %>");

        location.href = "./default.asp"
        return;
    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    <INPUT TYPE=HIDDEN NAME=CLASS   VALUE="<%=Request("CLASS")%>">
    <INPUT TYPE=HIDDEN NAME=GAKUNEN VALUE="<%=Request("GAKUNEN")%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>