<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0100_upd_tuku.asp
' 機      能: 下ページ 成績登録の登録、更新
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
' 変      数:
' 引      渡:
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2001/07/27 前田 智史
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
    Dim     m_sKyokanCd     '//教官CD
    Dim     m_iNendo 
    Dim     m_sSikenKBN
    Dim     m_sKamokuCd
    Dim     i_max 
    Dim     m_sGakuNo	'//学年
    Dim     m_sGakkaCd	'//学科

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
    w_sMsgTitle="成績登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sKyokanCd     = request("txtKyokanCd")
    m_iNendo        = request("txtNendo")
	m_sSikenKBN     = Cint(request("txtSikenKBN"))
	m_sKamokuCd     = request("KamokuCd")
	i_max           = request("i_Max")
	m_sGakuNo	= Cint(request("txtGakuNo"))	'//学年
	m_sGakkaCd	= request("txtGakkaCd")			'//学科

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

        '// 成績登録
'response.write "AAA" & "<BR>"

        w_iRet = f_Update(m_sSikenKBN)
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

'Function f_Update()
Function f_Update(p_sSikenKBN)
'********************************************************************************
'*  [機能]  ヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Dim i

    On Error Resume Next
    Err.Clear
    
    f_Update = 1

    Do 

		For i=1 to i_max

            '//T16_RISYU_KOJINにUPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE T34_RISYU_TOKU SET "

		'Select Case m_sSikenKBN
		Select Case p_sSikenKBN

			Case C_SIKEN_ZEN_TYU
				w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_TYUKAN_Z		= " & Cint(gf_SetNull2Zero(request("Kekka"&i))) & ", "
				w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_NASI_TYUKAN_Z		= " & Cint(gf_SetNull2Zero(request("KekkaGai"&i))) & ", "
				w_sSQL = w_sSQL & vbCrLf & " 	T34_CHIKAI_TYUKAN_Z		= " & Cint(gf_SetNull2Zero(request("Chikai"&i))) & ", "
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_KIMATU_Z		= " & Cint(gf_SetNull2Zero(request("Kekka"&i))) & ", "
				w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_NASI_KIMATU_Z		= " & Cint(gf_SetNull2Zero(request("KekkaGai"&i))) & ", "
				w_sSQL = w_sSQL & vbCrLf & " 	T34_CHIKAI_KIMATU_Z		= " & Cint(gf_SetNull2Zero(request("Chikai"&i))) & ", "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_TYUKAN_K		= " & Cint(gf_SetNull2Zero(request("Kekka"&i))) & ", "
				w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_NASI_TYUKAN_K		= " & Cint(gf_SetNull2Zero(request("KekkaGai"&i))) & ", "
				w_sSQL = w_sSQL & vbCrLf & " 	T34_CHIKAI_TYUKAN_K		= " & Cint(gf_SetNull2Zero(request("Chikai"&i))) & ", "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_KIMATU_K		= " & Cint(gf_SetNull2Zero(request("Kekka"&i))) & ", "
				w_sSQL = w_sSQL & vbCrLf & " 	T34_KEKA_NASI_KIMATU_K		= " & Cint(gf_SetNull2Zero(request("KekkaGai"&i))) & ", "
				w_sSQL = w_sSQL & vbCrLf & " 	T34_CHIKAI_KIMATU_K		= " & Cint(gf_SetNull2Zero(request("Chikai"&i))) & ", "
		End Select

            w_sSQL = w_sSQL & vbCrLf & "   T34_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
            w_sSQL = w_sSQL & vbCrLf & "   T34_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T34_NENDO = " & Cint(m_iNendo) & " "
            w_sSQL = w_sSQL & vbCrLf & "    AND T34_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "    AND T34_TOKUKATU_CD = '" & Trim(m_sKamokuCd) & "'  "

'response.write w_sSQL & "<BR>"

            iRet = gf_ExecuteSQL(w_sSQL)

            If iRet <> 0 Then

                '//ﾛｰﾙﾊﾞｯｸ
                msMsg = Err.description
                f_Update = 99
                Exit Do
            End If

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
    <title>成績登録</title>
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

		alert("<%=C_TOUROKU_OK_MSG%>");

	    document.frm.target = "main";
	    document.frm.action = "./sei0100_bottom.asp"
	    document.frm.submit();
	    return;

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type=hidden name=txtNendo    value="<%=trim(Request("txtNendo"))%>">
	<input type=hidden name=txtKyokanCd value="<%=trim(Request("txtKyokanCd"))%>">
	<input type=hidden name=txtSikenKBN value="<%=trim(Request("txtSikenKBN"))%>">
	<input type=hidden name=txtGakuNo   value="<%=trim(Request("txtGakuNo"))%>">
	<input type=hidden name=txtClassNo  value="<%=trim(Request("txtClassNo"))%>">
	<input type=hidden name=txtKamokuCd value="<%=trim(Request("txtKamokuCd"))%>">
	<input type=hidden name=txtGakkaCd  value="<%=trim(Request("txtGakkaCd"))%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

