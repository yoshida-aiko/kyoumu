<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 欠席日数登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/sei0600/sei0600_upd.asp
' 機      能: 下ページ 欠席日数の登録、更新
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
' 作      成: 2001/09/26 谷脇 良也
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
    Dim     m_iSikenKBN

    Public  m_iMax          '最大ページ
    Public  m_iDsp          '一覧表示行数
	Public  m_rCnt

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
    w_sMsgTitle="欠席日数登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_iSikenKBN = Cint(request("txtSikenKBN"))
	m_rCnt = cint(Request("txtCnt"))

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
    dim w_sSikenKBN_KE,w_sSikenKBN_KI
	dim w_Gno,w_Kss,w_kbk

    On Error Resume Next
    Err.Clear
    
    f_Update = 1

	select case cint(m_iSikenKBN)
		case C_SIKEN_ZEN_TYU '前期中間
			w_sSikenKBN_KE = "T13_KESSEKI_TYUKAN_Z"
			w_sSikenKBN_KI = "T13_KIBIKI_TYUKAN_Z"
		case C_SIKEN_ZEN_KIM '前期期末
			w_sSikenKBN_KE = "T13_KESSEKI_KIMATU_Z"
			w_sSikenKBN_KI = "T13_KIBIKI_KIMATU_Z"
		case C_SIKEN_KOU_TYU '後期中間
			w_sSikenKBN_KE = "T13_KESSEKI_TYUKAN_K"
			w_sSikenKBN_KI = "T13_KIBIKI_TYUKAN_K"
		case C_SIKEN_KOU_KIM '後期期末（学年末）
			w_sSikenKBN_KE = "T13_SUMKESSEKI"
			w_sSikenKBN_KI = "T13_SUMKIBTEI"
	End select

    Do 

	For i = 1 to m_rCnt 
	
        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()
		w_Gno = "txtGAKUSEINO_"&i
		w_Kss = "txtKESSEKI_"&i
		w_kbk = "txtKIBIKI_"&i

            '//T11_GAKUSEKIにUPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN SET "
            w_sSQL = w_sSQL & vbCrLf & "   " & w_sSikenKBN_KE & "= '"  & gf_SetNull2Zero(Request(w_Kss)) & "',"
            w_sSQL = w_sSQL & vbCrLf & "   " & w_sSikenKBN_KI & "= '"  & gf_SetNull2Zero(Request(w_kbk)) & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_DATE = '"    & gf_YYYY_MM_DD(date(),"/") & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_USER = '"    & Session("LOGIN_ID")       & "'"
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T13_GAKUSEI_NO = '" & Request(w_Gno) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "   AND  T13_NENDO = " & session("NENDO") & "  "
            
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
	Next
	
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
    <title>欠席日数登録</title>
    <link rel=stylesheet href="../../common/style.css" type=text/css>

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

        document.frm.action="default.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

		<input type="hidden" name="txtSikenKBN" value="<%=m_iSikenKBN%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

