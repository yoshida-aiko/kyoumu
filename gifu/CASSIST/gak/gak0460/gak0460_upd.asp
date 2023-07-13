<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 調査書所見等登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0460/gak0460_upd.asp
' 機      能: 下ページ 調査書所見等登録の登録、更新
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
' 作      成: 2001/07/19 前田 智史
' 変      更：2001/08/30 伊藤 公子     検索条件を2重に表示しないように変更
' 変      更：2002/10/08 廣田 耕一郎   担任所見、資格等の項目を追加
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
    Dim     m_sGakuNo       '//学生番号
    Dim     m_sSGSyoken 
    Dim     m_sBikou 
    Dim     m_sTanninSyoken  '担任所見		'2002.10.08 Add hirota
    Dim     m_sTanninBikou   '資格等  		'2002.10.08 Add hirota
	Dim     m_sNendo         '処理年度　　　'2002.10.08 Add hirota
    Dim     m_sSinroCd 
    Dim     m_sSRondai 
    Dim     m_sSKyokanCd1 
    Dim     m_sSKyokanCd2 
    Dim     m_sSKyokanCd3
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
    w_sMsgTitle="調査書所見等登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sKyokanCd     = session("KYOKAN_CD")
    m_sGakuNo       = request("txtGakuNo")
    m_sSGSyoken     = request("SGSyoken")
    m_sBikou        = request("Bikou")
    m_sSGSyoken     = request("SGSyoken")
    m_sTanninSyoken = request("TanninSyoken")		'2002.10.08 Add hirota
    m_sTanninBikou  = request("TanninBikou")		'2002.10.08 Add hirota
	m_sNendo        = request("txtNendo")			'2002.10.08 Add hirota
    m_sSRondai      = request("SRondai")
    m_sSKyokanCd1   = request("SKyokanCd1")
    m_sSKyokanCd2   = request("SKyokanCd2")
    m_sSKyokanCd3   = request("SKyokanCd3")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass  = Cint(request("txtClass"))
	m_sClassNm  = request("txtClassNm")
	m_sGakusei  = request("GakuseiNo")
    m_iDsp = C_PAGE_LINE

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        If gf_OpenDatabase() <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '//指導要録所見更新
		Call gs_BeginTrans()		'トランザクション開始         	'2002.10.08 hirota

        If f_Update() <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		if Not gf_GetGakkoNO(w_sGakkoNO) then
            m_bErrFlg = True
            Exit Do
		end if

		if w_sGakkoNO = cstr(C_NCT_KUMAMOTO) then

	        If f_Update_T13() <> 0 Then
	            m_bErrFlg = True
	            Exit Do
	        End If

		end if

		Call gs_CommitTrans() 		'トランザクションをコミット   	'2002.10.08 hirota

        '// ページを表示
        Call showPage()

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
		Call gs_RollbackTrans() 	'トランザクションをロールバック '2002.10.08 hirota
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  ヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Update()
	
	On Error Resume Next
	Err.Clear
	
	f_Update = 1
	
	Do
		'//T11_GAKUSEKIにUPDATE
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " UPDATE T11_GAKUSEKI SET "
		w_sSQL = w_sSQL & vbCrLf & "   T11_SOGOSYOKEN = '"  & Trim(m_sSGSyoken) & "' ,"
		w_sSQL = w_sSQL & vbCrLf & "   T11_KOJIN_BIK = '"  & Trim(m_sBikou) & "' ,"
		
		If m_sGakunen = 5 Then
			w_sSQL = w_sSQL & vbCrLf & "   T11_SINRO = '"  & Trim(m_sSinroCd) & "' ,"
			w_sSQL = w_sSQL & vbCrLf & "   T11_SOTUKEN_DAI = '"  & Trim(m_sSRondai) & "' ,"
			w_sSQL = w_sSQL & vbCrLf & "   T11_SOTU_KYOKAN_CD1 = '"  & Trim(m_sSKyokanCd1) & "' ,"
			w_sSQL = w_sSQL & vbCrLf & "   T11_SOTU_KYOKAN_CD2 = '"  & Trim(m_sSKyokanCd2) & "' ,"
			w_sSQL = w_sSQL & vbCrLf & "   T11_SOTU_KYOKAN_CD3 = '"  & Trim(m_sSKyokanCd3) & "', "
		End If
		
		w_sSQL = w_sSQL & vbCrLf & "   T11_UPD_DATE = '"  & gf_YYYY_MM_DD(date(),"/") & "', "
		w_sSQL = w_sSQL & vbCrLf & "   T11_UPD_USER = '"  & Session("LOGIN_ID") & "' "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "        T11_GAKUSEI_NO = '" & m_sGakuNo & "'  "
		
		If gf_ExecuteSQL(w_sSQL) <> 0 Then
			msMsg = Err.description
			f_Update = 99
			Exit Do
		End If
		
		'//正常終了
		f_Update = 0
		Exit Do
	Loop
	
End Function

'********************************************************************************
'*  [機能]  T13更新処理(担当所見,資格等)
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'*  [作成]  廣田 : 2002.10.08
'********************************************************************************
Function f_Update_T13()
	
	On Error Resume Next
	Err.Clear
	
	f_Update_T13 = 1
	
	Do
		'//T13_GAKU_NENにUPDATE
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN SET "
		w_sSQL = w_sSQL & vbCrLf & "   T13_TANNINSYOKEN = '"  & Trim(m_sTanninSyoken) & "' ,"
		w_sSQL = w_sSQL & vbCrLf & "   T13_TANNIN_BIK = '"  & Trim(m_sTanninBikou) & "' ,"

		w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_DATE = '"  & gf_YYYY_MM_DD(date(),"/") & "', "
		w_sSQL = w_sSQL & vbCrLf & "   T13_UPD_USER = '"  & Session("LOGIN_ID") & "' "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKUSEI_NO = '" & m_sGakuNo & "'  "
		w_sSQL = w_sSQL & vbCrLf & "   AND T13_NENDO = " & m_sNendo

		If gf_ExecuteSQL(w_sSQL) <> 0 Then
			msMsg = Err.description
			f_Update_T13 = 99
			Exit Do
		End If
		
		'//正常終了
		f_Update_T13 = 0
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

		<%
		'//登録ボタン押下時、初期画面に戻る
		If trim(Request("GakuseiNo")) = "" Then%>
	        document.frm.action="default.asp";
			document.frm.target="<%=C_MAIN_FRAME%>";
		<%
		'//前へOR次へボタン押下時、引き続き入力画面を表示する
		Else %>
    	    document.frm.action="gak0460_main.asp";
        	document.frm.target="main";
		<%End If %>
        document.frm.submit();

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">
		<input type="hidden" name="txtNendo" value="<%=request("txtNendo")%>">
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

