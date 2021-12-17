<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 放送大学成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0160_upd.asp
' 機      能: 下ページ 放送大学成績登録の登録、更新
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
' 変      数:
' 引      渡:
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Dim m_bErrFlg		'//ｴﾗｰﾌﾗｸﾞ
	
    '取得したデータを持つ変数

    Dim m_iNendo				'//年度
    Dim m_sKyokanCd				'//教官コード
    Dim m_sGakunen				'//学年
    Dim m_sClass				'//クラス

    Dim m_sBunruiCD		 		'//分類コード
    Dim m_sBunruiNM		 		'//分類名称
    Dim m_sTani		 			'//単位
    Dim i_max

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()
	Dim w_iRet              '// 戻り値
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="放送大学成績登録"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = False
	
	Do
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		if gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		end If
		
		'データ取得
		Call s_SetParam()
		
		'//不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//トランザクション開始
		Call gs_BeginTrans()
		
		'//履修認定テーブル削除処理
		if f_Delete() <> 0 Then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//履修認定テーブル更新処理
		If f_Update() <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'// ページを表示
		Call showPage()
		
		Exit Do
	Loop
	
    '//ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        '//ロールバック
        Call gs_RollbackTrans()
        
        w_sMsg = gf_GetErrMsg()
        'response.write "w_sMsg =" & w_sMsg & "<BR>"
        'response.end
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    else
    	'//コミット
    	Call gs_CommitTrans()
    End If
    
    '// 終了処理
    Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()

    	m_iNendo    	= request("txtNendo")		'年度
    	m_sKyokanCd 	= request("txtKyokanCd")	'教官コード


	m_sGakunen  	= request("txtGakunen")         '学年
	m_sClass    	= request("txtClass")           'クラス
	m_sBunruiCD 	= request("txtBunruiCd")	'分類コード
	m_sBunruiNm 	= request("txtBunruiNm")	'分類名称
	m_sTani     	= request("txtTani")		'単位

	i_max           = request("i_Max")

End Sub

'********************************************************************************
'*  [機能]  履修認定テーブル(T100_RISYU_NINTEI)削除処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Delete()
	Dim i
	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Delete = 99
	
	Do

		For i=1 to i_max

			w_sSQL = ""
			w_sSQL = w_sSQL & " DELETE"
			w_sSQL = w_sSQL & " FROM"
			w_sSQL = w_sSQL & " T100_RISYU_NINTEI"
			w_sSQL = w_sSQL & " WHERE"
			w_sSQL = w_sSQL & "      T100_GAKUSEI_NO = '" & request("txtGseiNo"&i)  & "'"	
			w_sSQL = w_sSQL & " AND  T100_BUNRUI_CD = '" & m_sBunruiCD & "'"
		
			'削除
			if gf_ExecuteSQL(w_sSQL) <> 0 then Exit Do
 'response.write w_sSQL
		Next
		
		'//正常終了
		f_Delete = 0
		
		Exit Do
	Loop
	
End Function


'********************************************************************************
'*  [機能]  履修認定テーブル(T100_RISYU_NINTEI)更新処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Update()
	Dim i
	Dim w_sSQL
	Dim w_Seiseki
	Dim w_Hyoka
	Dim w_sHyokaFuka
	Dim w_SumiTani

	
	On Error Resume Next
	Err.Clear
	
	f_Update = 99
	
	Do

		For i=1 to i_max

			w_Seiseki = gf_SetNull2String(request("Seiseki"&i))
			w_Hyoka   = gf_SetNull2String(request("hidHyoka"&i))

			'成績・評価が入力されたときのみ登録
			if w_Seiseki <> "" or w_Hyoka <> "" then 

				'修得年度が入力されていれば、修得とみなす 
				If request("SyuNendo"&i) <> "" Then
					w_sHyokaFuka = "0"	'合格
					w_SumiTani = m_sTani	'済単位
				Else
					w_sHyokaFuka = "1"	'不合格
					w_SumiTani = "0"	
				End If

				w_sSQL = ""
				w_sSQL = w_sSQL & " INSERT INTO T100_RISYU_NINTEI "
				w_sSQL = w_sSQL & " ("
				w_sSQL = w_sSQL & " T100_GAKUSEI_NO"
				w_sSQL = w_sSQL & ", T100_BUNRUI_CD"
				'w_sSql = w_sSql & ", T100_KYU_CD"
				w_sSQL = w_sSQL & ", T100_SYUTOKU_NENDO"
				w_sSQL = w_sSQL & ", T100_SYUTOKU_GAKUNEN"
				w_sSQL = w_sSQL & ", T100_GAKUSEKI_NO"
				w_sSQL = w_sSQL & ", T100_GAKKA_CD"
				w_sSQL = w_sSQL & ", T100_CLASS"
				w_sSQL = w_sSQL & ", T100_COURSE_CD"
				w_sSQL = w_sSQL & ", T100_BUNRUI_MEISYO"
				'w_sSql = w_sSql & ", T100_KYU_MEI"
				w_sSQL = w_sSQL & ", T100_HAITOTANI"
				w_sSQL = w_sSQL & ", T100_TANI_SUMI"
				'w_sSql = w_sSql & ", T100_HYOTEI"
				w_sSQL = w_sSQL & ", T100_HYOKA"
				w_sSQL = w_sSQL & ", T100_HYOKA_FUKA_KBN"
				'w_sSql = w_sSql & ", T100_NINTEIBI"
				w_sSQL = w_sSQL & ", T100_SEISEKI"
				w_sSQL = w_sSQL & ", T100_INS_DATE"
				w_sSQL = w_sSQL & ", T100_INS_USER"

				w_sSQL = w_sSQL & " )VALUES("

				w_sSQL = w_sSQL & " '" & request("txtGseiNo"&i)  & "'"		'T100_GAKUSEI_NO"
				w_sSQL = w_sSQL & ",'" & m_sBunruiCD & "'"			'T100_BUNRUI_CD"
				'w_sSql = w_sSql & ", T100_KYU_CD"
				w_sSQL = w_sSQL & ", " & f_CnvNumNull(request("SyuNendo"&i)) 	'T100_SYUTOKU_NENDO"
				w_sSQL = w_sSQL & ", " & m_sGakunen				'T100_SYUTOKU_GAKUNEN"
				w_sSQL = w_sSQL & ",'" & request("txtGsekiNo"&i) & "'"		'T100_GAKUSEKI_NO"
				w_sSQL = w_sSQL & ",'" & request("txtGakkaCD"&i) & "'"		'T100_GAKKA_CD"
				w_sSQL = w_sSQL & ",'" & request("txtClass"&i) 	& "'"		'T100_CLASS"
				w_sSQL = w_sSQL & ",'" & request("txtCorceCD"&i) & "'"		'T100_COURSE_CD"
				w_sSQL = w_sSQL & ",'" & m_sBunruiNm & "'"			'T100_BUNRUI_MEISYO"
				'w_sSql = w_sSql & ", T100_KYU_MEI"
				w_sSQL = w_sSQL & ", " & m_sTani				'T100_HAITOTANI"
				w_sSQL = w_sSQL & ", " & w_SumiTani 				'T100_TANI_SUMI"
				'w_sSql = w_sSql & ", T100_HYOTEI"
				w_sSQL = w_sSQL & ",'" & w_Hyoka & "'"				'T100_HYOKA"
				w_sSQL = w_sSQL & ", " & w_sHyokaFuka 				'T100_HYOKA_FUKA_KBN"
				'w_sSql = w_sSql & ", T100_NINTEIBI"
				w_sSQL = w_sSQL & ", " & f_CnvNumNull(w_Seiseki)		'T100_SEISEKI
				w_sSQL = w_sSQL & ",'" & gf_YYYY_MM_DD(date(),"/") & "'" 	'T100_INS_DATE"
				w_sSQL = w_sSQL & ",'" & Trim(Session("LOGIN_ID")) & "'" 	'T100_INS_USER"
				w_sSQL = w_sSQL & " )"
'response.write w_sSQL

				'実行
				if gf_ExecuteSQL(w_sSQL) <> 0 then Exit Do
			End If
		Next
		'//正常終了
		f_Update = 0
 		
		Exit Do
	Loop
	
End Function

'********************************************************************************
'*  [機能]  数値型項目の更新時の設定
'*  [引数]  値
'*  [戻値]  なし
'*  [説明]  数値が入っている場合は[値]、無い場合は"NULL"を返す
'********************************************************************************
Function f_CnvNumNull(p_vAtai)

	If Trim(p_vAtai) = "" Then
		f_CnvNumNull = "NULL"
	Else
		f_CnvNumNull = cInt(p_vAtai)
    End If

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
%>
    <html>
    <head>
    <title>放送大学成績登録</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
	
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
	
    //************************************************************
    //  [機能]  ページロード時処理
    //************************************************************
    function window_onload() {
	alert("<%=C_TOUROKU_OK_MSG%>");
	document.frm.action="sei0160_top.asp";
	document.frm.target="topFrame";
	document.frm.submit();
	document.frm.target = "main";
	document.frm.action = "sei0160_bottom.asp"
	document.frm.submit();
	}
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE="javascript" onload="window_onload();">
    <form name="frm" method="post">
	
	<input type="hidden" name="txtNendo"     value="<%=trim(Request("txtNendo"))%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=trim(Request("txtKyokanCd"))%>">
	<input type="hidden" name="txtGakunen"   value="<%=trim(Request("txtGakunen"))%>">
	<input type="hidden" name="txtClass"     value="<%=trim(Request("txtClass"))%>">
	<input type="hidden" name="txtBunruiCd"  value="<%=trim(Request("txtBunruiCd"))%>">
	<input type="hidden" name="txtBunruiNm"  value="<%=trim(Request("txtBunruiNm"))%>">
	<input type="hidden" name="txtTani"      value="<%=trim(Request("txtTani"))%>">
	<input type="hidden" name="i_Max"        value="<%=request("i_Max")%>">
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>