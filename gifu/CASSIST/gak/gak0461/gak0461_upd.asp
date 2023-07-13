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
' 変      数:
' 引      渡: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2001/07/18 前田 智史
' 変      更：2001/08/30 伊藤 公子     検索条件を2重に表示しないように変更
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
    Dim     m_sKyokanCd     '//教官CD
'    Dim     m_sKSyoken      '//行動所見
'    Dim     m_sSyumiTokugi  '//趣味等
    Dim     m_sBikou        '//備考
    Dim     m_sTokukatu     '//特別活動
    Dim     m_sNSyoken      '//年毎所見
    Dim     m_sNendo        '//
    Dim     m_sGakuNo       '//学生No
    Dim     m_sGakunen
    Dim     m_sClass
    Dim     m_sClassNm
    Dim     m_sGakusei

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

    m_sKyokanCd = session("KYOKAN_CD")
'    m_sKSyoken      = request("KSyoken")
'    m_sSyumiTokugi  = request("SyumiTokugi")
    m_sBikou        = request("Bikou")
    m_sTokukatu     = request("Tokukatu")
    m_sNSyoken      = request("NSyoken")
    m_sNendo        = request("txtNendo")
    m_sGakuNo       = request("txtGakuNo")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass  = Cint(request("txtClass"))
	m_sClassNm  = request("txtClassNm")
	m_sGakusei  = request("GakuseiNo")
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
        w_iRet = f_Update()
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

Function f_Update()
'********************************************************************************
'*  [機能]  ヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    
    f_Update = 1

    Do 

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

            '//T11_GAKUSEKIにUPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE T11_GAKUSEKI SET "
'            w_sSQL = w_sSQL & vbCrLf & "   T11_KODOSYOKEN = '"  & Trim(m_sKSyoken) & "' ,"
'            w_sSQL = w_sSQL & vbCrLf & "   T11_SYUMITOKUGI = '"  & Trim(m_sSyumiTokugi) & "' ,"
            w_sSQL = w_sSQL & vbCrLf & "   T11_TYOSA_BIK = '"  & Trim(m_sBikou) & "' ,"
            w_sSQL = w_sSQL & vbCrLf & "   T11_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/")  & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T11_UPD_USER = '"    & Session("LOGIN_ID")       & "'"
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T11_GAKUSEI_NO = '" & m_sGakuNo & "'  "

            iRet = gf_ExecuteSQL(w_sSQL)
            If iRet <> 0 Then
                '//ﾛｰﾙﾊﾞｯｸ
                Call gs_RollbackTrans()
                msMsg = Err.description
                f_Update = 99
                Exit Do
            End If

            '//T13_GAKU_NENにUPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN SET "
            w_sSQL = w_sSQL & vbCrLf & "   T13_TOKUKATU_DET = '"  & Trim(m_sTokukatu) & "' ,"
            w_sSQL = w_sSQL & vbCrLf & "   T13_NENSYOKEN = '"  & Trim(m_sNSyoken) & "' ,"
            w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/") & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_USER = '"    & Session("LOGIN_ID")       & "'"
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T13_NENDO = '" & m_sNendo & "'  "
            w_sSQL = w_sSQL & vbCrLf & "    AND T13_GAKUSEI_NO = '" & m_sGakuNo & "' "

            iRet = gf_ExecuteSQL(w_sSQL)
            If iRet <> 0 Then
                '//ﾛｰﾙﾊﾞｯｸ
                Call gs_RollbackTrans()
                msMsg = Err.description
                f_Update = 99
                Exit Do
            End If

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
    <title>調査書所見等登録</title>
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

		parent.topFrame.location.href = "white.htm";

		<%
		'//登録ボタン押下時、初期画面に戻る
		If trim(Request("GakuseiNo")) = "" Then%>
	        document.frm.action="default.asp";
			document.frm.target="<%=C_MAIN_FRAME%>";
		<%
		'//前へOR次へボタン押下時、引き続き入力画面を表示する
		Else %>
    	    document.frm.action="gak0461_main.asp";
	        document.frm.target="main";
		<%End If %>

        document.frm.submit();

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
		<input type="hidden" name="txtNendo" value="<%=m_sNendo%>">
		<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
		<input type="hidden" name="GakuseiNo" value="<%=m_sGakusei%>">
		<input type="hidden" name="txtClass" value="<%=m_sClass%>">
		<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">

		<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

