<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0110/kks0110_edt.asp
' 機      能: 下ページ授業出欠入力の登録、更新
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             TUKI           '//月
' 変      数:
' 引      渡: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             TUKI           '//月
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 伊藤公子
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
    Public m_sGakunen
    Public m_sClassNo
    Public m_sTuki
    Public m_sKamokuCd

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
    w_sMsgTitle="授業出欠入力"
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

        '// 教科別出欠登録
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

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sTuki     = ""
    m_sKamokuCd = ""

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
    m_sGakunen  = trim(Request("GAKUNEN"))
    m_sClassNo  = trim(Request("CLASSNO"))
    m_sTuki     = trim(Request("TUKI"))
    m_sKamokuCd = trim(Request("KAMOKU_CD"))

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
Exit Sub
    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"

End Sub

'********************************************************************************
'*  [機能]  教科別出欠登録
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_AbsUpdate()

    Dim w_sSQL
    Dim w_Rs
    Dim w_iKekka

    On Error Resume Next
    Err.Clear
    
    f_AbsUpdate = 1

    Do 

		'//ﾕｰｻﾞｰIDを取得
		w_sUserId = Session("LOGIN_ID")

        '//時間割情報を取得
        m_Date_Jigen = split(replace(Request("JIKANWARI")," ",""),",")
        m_iJikanCnt = UBound(m_Date_Jigen)

        '//学籍Noを取得
        m_sGakuseiNo = split(replace(Request("GAKUSEI")," ",""),",")
        m_iGakusekiCnt = UBound(m_sGakuseiNo)

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

        '//クラスの人数分処理を実行
        For i=0 To m_iGakusekiCnt

            '//時間数分処理を実行
            For j=0 to m_iJikanCnt

                '//出欠CDを取得
                w_iKekka = trim(Request("hidKBN" & m_sGakuseiNo(i) & "_" & replace(m_Date_Jigen(j),"/","")))

                If w_iKekka = "---" Or Trim(m_sGakuseiNo(i)) = "" Then
                    '//出欠入力不可のときは更新処理をしない
                Else

                    w_DJ = split(m_Date_Jigen(j),"_")
                    w_Date  = w_DJ(0)
                    w_Jigen = replace(w_DJ(1),"$",".")

                    '//学年、クラスNOを取得
                    iRet = f_Get_NenClass(m_sGakuseiNo(i),w_Gakunen,w_Class,w_Gakuseki)
                    If iRet <> 0 Then
                        Exit Do
                    End If

                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " SELECT "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_NENDO, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_HIDUKE, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_GAKUSEKI_NO, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIGEN, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_KAMOKU, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_KYOKAN, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_SYUKKETU_KBN, "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIMU_FLG"
                    w_sSQL = w_sSQL & vbCrLf & " FROM T21_SYUKKETU"
                    w_sSQL = w_sSQL & vbCrLf & " WHERE "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_NENDO=" & cInt(m_iSyoriNen) & " AND "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_HIDUKE='" & w_Date & "' AND "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_GAKUSEKI_NO='" & w_Gakuseki & "' AND "
                    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIGEN=" & w_Jigen

                    iRet = gf_GetRecordset(rs, w_sSQL)
                    If iRet <> 0 Then
                        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                        msMsg = Err.description
                        f_AbsUpdate = 99
                        Exit Do
                    End If

                    If rs.EOF Then

                        If w_iKekka <> "" and cstr(w_iKekka)<>"0" Then

                            '//T22_GYOJI_SYUKKETUに生徒情報がない場合で、欠席数が入力されている場合はINSERT
                            w_sSQL = ""
                            w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T21_SYUKKETU  "
                            w_sSQL = w_sSQL & vbCrLf & "   ("
                            w_sSQL = w_sSQL & vbCrLf & "  T21_NENDO, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_HIDUKE, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_YOUBI_CD, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_GAKUNEN, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_CLASS, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_GAKUSEKI_NO, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_JIGEN, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_KAMOKU, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_KYOKAN, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU_KBN, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_JIMU_FLG, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_INS_DATE, "
                            w_sSQL = w_sSQL & vbCrLf & "  T21_INS_USER"
                            w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
                            w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_iSyoriNen) & " ,"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(w_Date)      & "',"
                            w_sSQL = w_sSQL & vbCrLf & "    "  & cint(Weekday(w_Date))   & ","
                            w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(w_Gakunen)   & " ,"
                            w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(w_Class)     & " ,"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(w_Gakuseki)  & "',"
                            w_sSQL = w_sSQL & vbCrLf & "    "  & w_Jigen     & " ,"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(m_sKamokuCd) & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(m_iKyokanCd) & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(w_iKekka)    & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & cstr(C_JIMU_FLG_NOTJIMU) & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_YYYY_MM_DD(date(),"/")            & "',"
                            w_sSQL = w_sSQL & vbCrLf & "   '"  & w_sUserId         & "' "
                            w_sSQL = w_sSQL & vbCrLf & "   )"

'response.write w_sSQL & "<br>"
'response.write "INSERT iRet = " & iRet & "<br>"

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

                        '//T21_SYUKKETUにすでに生徒情報がある場合はUPDATE
                        w_sSQL = ""
                        w_sSQL = w_sSQL & vbCrLf & " UPDATE T21_SYUKKETU SET "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_SYUKKETU_KBN ='" & Trim(w_iKekka)    & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   T21_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")            & "',"
                        w_sSQL = w_sSQL & vbCrLf & "   T21_UPD_USER = '"    & w_sUserId         & "' "
                        w_sSQL = w_sSQL & vbCrLf & " WHERE "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_NENDO="          & cInt(m_iSyoriNen) & "  AND "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_HIDUKE='"        & Trim(w_Date)      & "' AND "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_GAKUSEKI_NO='"   & Trim(w_Gakuseki)  & "' AND "
                        w_sSQL = w_sSQL & vbCrLf & "   T21_JIGEN="          & w_Jigen

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

                    End If

                    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
                    Call gf_closeObject(rs)

                End If

            Next
        Next

        '//ｺﾐｯﾄ
        Call gs_CommitTrans()

        '//正常終了
        f_AbsUpdate = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  学年、学績NOを取得
'*  [引数]  p_Gakuseki：学生NO
'*  [戻値]  p_Gakunen：学年
'*          p_Class：クラス
'*          p_Gakuseki:学績NO
'*  [説明]  
'********************************************************************************
Function f_Get_NenClass(p_Gakusei,p_Gakunen,p_Class,p_Gakuseki)

    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear
    
    f_Get_NenClass = 1

    p_Gakunen = ""
    p_Class = ""
    p_Gakuseki = ""

    Do 

        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLASS,"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO"
        w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & " T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO='" & trim(p_Gakusei) & "'"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_NenClass = 99
            Exit Do
        End If

        If rs.EOF = false Then
            p_Gakunen  = rs("T13_GAKUNEN")
            p_Class    = rs("T13_CLASS")
            p_Gakuseki = rs("T13_GAKUSEKI_NO")
        End If

        f_Get_NenClass = 0
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
    <title>授業出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>

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

		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("再表示しています　しばらくお待ちください")%>"

	    parent.main.document.frm.target = "main";
        //parent.main.document.frm.action="./WaitAction.asp";
	    parent.main.document.frm.action = "./kks0110_bottom.asp"
	    parent.main.document.frm.submit();
	    return;


    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    <input type="hidden" name="Tuki_Zenki_Start" value="<%=Request("Tuki_Zenki_Start")%>">
    <input type="hidden" name="Tuki_Kouki_Start" value="<%=Request("Tuki_Kouki_Start")%>">
    <input type="hidden" name="Tuki_Kouki_End"   value="<%=Request("Tuki_Kouki_End")%>">
    <INPUT TYPE=HIDDEN NAME="NENDO"     value="<%=Request("NENDO")%>">
    <INPUT TYPE=HIDDEN NAME="KYOKAN_CD" value="<%=Request("KYOKAN_CD")%>">
    <INPUT TYPE=HIDDEN NAME="TUKI"      value="<%=Request("TUKI")%>">
    <INPUT TYPE=HIDDEN NAME="GAKKI"     value="<%=Request("GAKKI")%>">
    <INPUT TYPE=HIDDEN NAME="GAKUNEN"   value="<%=Request("GAKUNEN")%>">
    <INPUT TYPE=HIDDEN NAME="CLASSNO"   value="<%=Request("CLASSNO")%>">
    <INPUT TYPE=HIDDEN NAME="KAMOKU_CD" value="<%=Request("KAMOKU_CD")%>">
    <INPUT TYPE=HIDDEN NAME="SYUBETU"   value="<%=Request("SYUBETU")%>">

    <INPUT TYPE=HIDDEN NAME="KAMOKU_NAME" value="<%=Request("KAMOKU_NAME")%>">
    <INPUT TYPE=HIDDEN NAME="CLASS_NAME"  value="<%=Request("CLASS_NAME")%>">

    <input TYPE="HIDDEN" NAME="txtURL" VALUE="kks0110_bottom.asp">
    <input TYPE="HIDDEN" NAME="txtMsg" VALUE="<%=Server.HTMLEncode("再表示しています　しばらくお待ちください")%>">

    </form>
    </body>
    </html>
<%
End Sub
%>