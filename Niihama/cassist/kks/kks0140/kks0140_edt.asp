<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 行事出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0140/kks0140_edt.asp
' 機      能: 下ページ 行事出欠入力の登録、更新
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
    Const DebugPrint = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public m_iSyoriNen
    Public m_iKyokanCd
    Public m_sGakunen
    Public m_sClassNo
    Public m_sTuki

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
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sTuki     = ""

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

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"

End Sub

'********************************************************************************
'*  [機能]  ヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_AbsUpdate()

    Dim w_sSQL
    Dim w_Rs
    Dim w_sUserName
    Dim w_iKekka

    On Error Resume Next
    Err.Clear
    
    f_AbsUpdate = 1

    Do 

		'//ﾕｰｻﾞｰIDを取得
		w_sUserId = Session("LOGIN_ID")

        '//行事CDを取得
        m_sGyojiCD = Request("GYOJI_CD")

        '//学籍Noを取得
        m_sGakusekiNo = split(replace(Request("GAKUSEKI_NO")," ",""),",")
        m_iGakusekiCnt = UBound(m_sGakusekiNo)

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

        '//クラスの人数分処理を実行
        For i=0 To m_iGakusekiCnt

            '//欠席数を取得
            w_iKekka = replace(trim(Request("SU_" & m_sGakusekiNo(i))),"+","")

            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " SELECT "
            w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_KEKKA"
            w_sSQL = w_sSQL & vbCrLf & " FROM T22_GYOJI_SYUKKETU"
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "   T22_NENDO=" & cInt(m_iSyoriNen) & " AND "
            w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN=" & cInt(m_sGakunen) & " AND "
            w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS=" & cInt(m_sClassNo) & " AND "
            w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUSEKI_NO='" & Trim(m_sGakusekiNo(i)) & "' AND "
            w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD='" & Trim(m_sGyojiCD) & "'"

            iRet = gf_GetRecordset(rs, w_sSQL)
            If iRet <> 0 Then
                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                msMsg = Err.description
                f_AbsUpdate = 99
                Exit Do
            End If

            If rs.EOF Then

                If w_iKekka <> "" Then

                    '//T22_GYOJI_SYUKKETUに生徒情報がない場合で、欠席数が入力されている場合はINSERT
                    w_sSQL = ""
                    w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T22_GYOJI_SYUKKETU  "
                    w_sSQL = w_sSQL & vbCrLf & "   ("
                    w_sSQL = w_sSQL & vbCrLf & "   T22_NENDO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUSEKI_NO, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_KEKKA, "
                    w_sSQL = w_sSQL & vbCrLf & "   T22_INS_DATE,"
                    w_sSQL = w_sSQL & vbCrLf & "   T22_INS_USER"
                    w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
                    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_iSyoriNen)      & " ,"
                    w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(m_sGyojiCD)       & "',"
                    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_sGakunen)       & " ,"
                    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_sClassNo)       & " ,"
                    w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_SetNull2Zero(Trim(m_sGakusekiNo(i))) & "',"
                    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(gf_SetNull2Zero(trim(w_iKekka))) & " ,"
                    w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_YYYY_MM_DD(date(),"/")                 & "',"
                    w_sSQL = w_sSQL & vbCrLf & "   '"  & w_sUserId              & "' "
                    w_sSQL = w_sSQL & vbCrLf & "   )"

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

                '//T22_GYOJI_SYUKKETUにすでに生徒情報がある場合はUPDATE
                w_sSQL = ""
                w_sSQL = w_sSQL & vbCrLf & " UPDATE T22_GYOJI_SYUKKETU SET "
                w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_KEKKA =" & cInt(gf_SetNull2Zero(trim(w_iKekka))) & " ,"
                w_sSQL = w_sSQL & vbCrLf & "   T_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")              & "',"
                w_sSQL = w_sSQL & vbCrLf & "   T_UPD_USER = '"    & w_sUserId           & "'"
                w_sSQL = w_sSQL & vbCrLf & " WHERE "
                w_sSQL = w_sSQL & vbCrLf & "   T22_NENDO="        & cInt(m_iSyoriNen) & " AND "
                w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN="      & cInt(m_sGakunen)  & " AND "
                w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS="        & cInt(m_sClassNo)  & " AND "
                w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUSEKI_NO='" & Trim(m_sGakusekiNo(i)) & "' AND "
                w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD='"    & Trim(m_sGyojiCD)  & "'"

                iRet = gf_ExecuteSQL(w_sSQL)
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
    <title>行事出欠入力</title>
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

        //リスト情報をsubmit
        document.frm.target = "main";
        document.frm.action = "./kks0140_bottom.asp"
        document.frm.submit();
        return;


    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

    <input type="hidden" name="NENDO"     value="<%=Request("NENDO")%>">
    <input type="hidden" name="KYOKAN_CD" value="<%=Request("KYOKAN_CD")%>">
    <input type="hidden" name="GAKUNEN"   value="<%=Request("GAKUNEN")%>">
    <input type="hidden" name="CLASSNO"   value="<%=Request("CLASSNO")%>">
    <input type="hidden" name="TUKI"      value="<%=Request("TUKI")%>">
    <INPUT TYPE=HIDDEN NAME="GYOJI_CD"  value = "<%=Request("GYOJI_CD")%>">
    <INPUT TYPE=HIDDEN NAME="GYOJI_MEI" value = "<%=Request("GYOJI_MEI")%>">
    <INPUT TYPE=HIDDEN NAME="KAISI_BI"  value = "<%=Request("KAISI_BI")%>">
    <INPUT TYPE=HIDDEN NAME="SYURYO_BI" value = "<%=Request("SYURYO_BI")%>">
    <INPUT TYPE=HIDDEN NAME="SOJIKANSU" value = "<%=Request("SOJIKANSU")%>">

	</form>
    </body>
    </html>
<%
End Sub
%>