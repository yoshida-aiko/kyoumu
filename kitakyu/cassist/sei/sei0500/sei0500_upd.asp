<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0500/sei0500_upd.asp
' 機      能: 下ページ 成績登録の登録、更新
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
' 変      数:
' 引      渡:
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2001/09/07 モチナガ
' 変      更: 2016/05/18 Nishimura 異動(休学者)の場合更新できない障害対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="実力試験成績登録"
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

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()


		'// 成績を更新する
		If f_SeisekiUpdate() then
			'//ｺﾐｯﾄ
			Call gs_CommitTrans()
		Else
			'// ﾛｰﾙﾊﾞｯｸ
			Call gs_RollbackTrans()
			Exit Do
		End if

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

Function f_SeisekiUpdate()
'********************************************************************************
'*  [機能]  成績を更新する
'*  [引数]  なし
'*  [戻値]  True : False
'*  [説明]  
'********************************************************************************
Dim i

    On Error Resume Next
    Err.Clear
    
    f_SeisekiUpdate = False

	'// ﾊﾟﾗﾒｰﾀ取得
	w_iNendo	= request("txtNendo")
	w_sKyokanCd	= request("txtKyokanCd")
	w_sSiKenCd	= Cint(request("txtShikenCd"))
	w_sGakuNo	= Cint(request("txtGakuNo"))
	w_sClassNo	= Cint(request("txtClassNo"))
	w_sKamokuCd	= request("txtKamokuCd")

	m_rCnt      = Cint(request("hidRecCnt"))

	i = 1
	Do Until i > m_rCnt

		w_iIdoCnt = request("hidIdoCnt" & i )	'Ins 2016/05/18 Nishimura 休学者の場合エラーになるため、IF文 分岐追加
		IF w_iIdoCnt = 1 Then					'Ins 2016/05/18 Nishimura 

		w_SQL = ""
		w_SQL = w_SQL & vbCrLf & " Update T33_SIKEN_SEISEKI Set "
'2015/10/20 UPDATE URAKAWA NULLの時0を登録しない。
'		w_SQL = w_SQL & vbCrLf & "	T33_TOKUTEN		 =  " & Cint(gf_SetNull2Zero(request("Seiseki" & i))) & ", "
		w_SQL = w_SQL & vbCrLf & "	T33_TOKUTEN		 =  '" & gf_SetNull2String(request("Seiseki" & i)) & "', "
		w_SQL = w_SQL & vbCrLf & "	T33_UPD_DATE	 = '" & gf_YYYY_MM_DD(date(),"/") & "',"
		w_SQL = w_SQL & vbCrLf & "	T33_UPD_USER	 = '" & w_sKyokanCd & "'"
		w_SQL = w_SQL & vbCrLf & " WHERE "
		w_SQL = w_SQL & vbCrLf & "	T33_NENDO		 =  " & w_iNendo & " AND"
		w_SQL = w_SQL & vbCrLf & "	T33_SIKEN_KBN	 =  " & C_SIKEN_JITURYOKU & " AND"
		w_SQL = w_SQL & vbCrLf & "	T33_SIKEN_CD	 =  " & w_sSiKenCd & " AND"
		w_SQL = w_SQL & vbCrLf & "	T33_SIKEN_KAMOKU = '" & w_sKamokuCd & "' AND"
		w_SQL = w_SQL & vbCrLf & "	T33_GAKUSEKI_NO  = '" & request("hidGakusekiNo" & i ) & "' AND"
		w_SQL = w_SQL & vbCrLf & "	T33_GAKUNEN  	 =  " & w_sGakuNo & " AND"
		w_SQL = w_SQL & vbCrLf & "	T33_CLASS		 =  " & w_sClassNo

		iRet = gf_ExecuteSQL(w_SQL)
		If iRet <> 0 Then
			m_bErrFlg = True
			msMsg = Err.description
			Exit Function
		End If

		END IF	'Ins 2016/05/18 Nishimura 

		i = i + 1
	Loop

	'//正常終了
	f_SeisekiUpdate = True

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
	    document.frm.action = "sei0500_bottom.asp"
	    document.frm.submit();

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">

	<input type=hidden name=txtNendo    value="<%=trim(Request("txtNendo"))%>">
	<input type=hidden name=txtKyokanCd value="<%=trim(Request("txtKyokanCd"))%>">
	<input type=hidden name=txtShikenCd value="<%=trim(Request("txtShikenCd"))%>">
	<input type=hidden name=txtGakuNo   value="<%=trim(Request("txtGakuNo"))%>">
	<input type=hidden name=txtClassNo  value="<%=trim(Request("txtClassNo"))%>">
	<input type=hidden name=txtKamokuCd value="<%=trim(Request("txtKamokuCd"))%>">

	</form>
    </body>
    </html>
<%
End Sub
%>

