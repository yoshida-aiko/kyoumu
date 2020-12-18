<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0112/kks0112_edit.asp
' 機      能: 前ページで(kks0112_bottom.asp)登録した出欠状況を登録する
'-------------------------------------------------------------------------
' 引      数: 
'             
'             
'             
'             
' 変      数: 
' 引      渡: 
'             
'             
'             
'             
' 説      明: 
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2002/05/16 shin
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Dim m_bErrFlg           'エラーフラグ
	
    '取得したデータを持つ変数
    Dim m_iSyoriNen
    Dim m_iKyokanCd
    
	Dim m_sGakunen
	Dim m_sClassNo
	Dim m_sKamokuCd
	
	Dim m_sDate
	Dim m_iJigen
	Dim m_sUserId
	Dim m_iKamokuKbn
	
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
        If gf_OpenDatabase() <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
        '//変数初期化
        Call s_ClearParam()
		
        '//MainﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
		
        '//教科別出欠登録
        If not f_AbsEdit() Then
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
	m_sUserId = ""
    m_iSyoriNen = ""
    m_iKyokanCd = ""
    
	m_sGakunen = ""
	m_sClassNo = ""
	m_sKamokuCd = ""
	
	m_sDate = ""
	m_iJigen = ""
	m_iKamokuKbn = 0
	
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	m_sUserId = Session("LOGIN_ID")
	
    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")
    
	m_sGakunen = trim(Request("hidGakunen"))
	m_sClassNo = trim(Request("hidClassNo"))
	m_sKamokuCd = trim(Request("hidKamokuCd"))
	
	m_sDate = gf_YYYY_MM_DD(trim(Request("hidDate")),"/")
	m_iJigen = trim(Request("hidJigen"))
	
	m_iKamokuKbn = cint(request("hidSyubetu"))
	
End Sub

'********************************************************************************
'*  [機能]  教科別出欠登録
'*  [引数]  なし
'*  [戻値]  false:情報取得成功 true:失敗
'*  [説明]  デリート後、インサートする
'********************************************************************************
Function f_AbsEdit()
	Dim w_sSQL
    Dim w_Rs
    Dim w_iKekka
	Dim w_sUserId
	Dim w_sGakusekiNo,w_iCount
	Dim w_State
	Dim w_JikanNum
	
    On Error Resume Next
    Err.Clear
    
    f_AbsEdit = false
	
	'//学籍NO
	w_sGakusekiNo = split(replace(Request("hidGakusekiNo")," ",""),",")
	
	'//学生数
	w_iCount = UBound(w_sGakusekiNo)
	
	'//ﾄﾗﾝｻﾞｸｼｮﾝ開始
    Call gs_BeginTrans()
	
	'//delete
	if not f_AbsDelete() then
		'//ﾛｰﾙﾊﾞｯｸ
		Call gs_RollbackTrans()
		exit function
	end if
	
	for i=0 to w_iCount
		w_State = 0
		w_State = gf_SetNull2Zero(trim(request("hidState" & w_sGakusekiNo(i))))
		
		w_JikanNum = 0
		w_JikanNum = gf_SetNull2Zero(trim(request("hidJikanState" & w_sGakusekiNo(i))))
		
		'//状況が選択されていたら、insert
		if w_State <> 0 then
			if not f_AbsInsert(w_sGakusekiNo(i),w_State,w_JikanNum) then
				'//ﾛｰﾙﾊﾞｯｸ
				Call gs_RollbackTrans()
				exit function
			end if
		end if
	next
	
	'//ｺﾐｯﾄ
	Call gs_CommitTrans()
    
    f_AbsEdit = true
    
End Function

'********************************************************************************
'*  [機能]  
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Function f_AbsInsert(p_GakusekiNo,p_State,p_JikanNum)
	
	Dim w_sSQL
    
    On Error Resume Next
    Err.Clear
    
    f_AbsInsert = false
	
	'if cInt(p_State) <> 1 then p_JikanNum = 0
	
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
	w_sSQL = w_sSQL & vbCrLf & "  T21_JIKANSU, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_JIMU_FLG, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_KAMOKU_KBN, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_INS_DATE, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_INS_USER"
    w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
    w_sSQL = w_sSQL & vbCrLf & "    "  & m_iSyoriNen				& " ,"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & m_sDate					& "',"
    w_sSQL = w_sSQL & vbCrLf & "    "  & cint(Weekday(m_sDate))		& ","
    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_sGakunen)			& " ,"
    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_sClassNo)			& " ,"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & p_GakusekiNo				& "',"
    w_sSQL = w_sSQL & vbCrLf & "    "  & m_iJigen					& " ,"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & m_sKamokuCd				& "',"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(m_iKyokanCd)			& "',"
    w_sSQL = w_sSQL & vbCrLf & "   "   & p_State					& ","
    w_sSQL = w_sSQL & vbCrLf & "   "   & p_JikanNum					& ","
    w_sSQL = w_sSQL & vbCrLf & "   "   & C_JIMU_FLG_NOTJIMU			& ","
    w_sSQL = w_sSQL & vbCrLf & "   "   & m_iKamokuKbn				& ","
    w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_YYYY_MM_DD(date(),"/")	& "',"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & m_sUserId					& "' "
    w_sSQL = w_sSQL & vbCrLf & "   )"
	
	if gf_ExecuteSQL(w_sSQL) <> 0 Then exit function
    
	'//正常終了
    f_AbsInsert = true
    
End Function


'********************************************************************************
'*  [機能]  
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Function f_AbsDelete()
	
	Dim w_sSQL
    
    On Error Resume Next
    Err.Clear
    
    f_AbsDelete = false
	
    w_sSQL = ""
    w_sSQL = w_sSQL & " delete from T21_SYUKKETU  "
    w_sSQL = w_sSQL & "  where "
    w_sSQL = w_sSQL & "      T21_NENDO			=  " & m_iSyoriNen
    w_sSQL = w_sSQL & "  and T21_HIDUKE			= '" & m_sDate				& "' "
    w_sSQL = w_sSQL & "  and T21_GAKUNEN		=  " & cInt(m_sGakunen)
    w_sSQL = w_sSQL & "  and T21_CLASS			=  " & cInt(m_sClassNo)
    w_sSQL = w_sSQL & "  and T21_JIGEN			=  " & m_iJigen
	w_sSQL = w_sSQL & "  and T21_KAMOKU			= '" & m_sKamokuCd			& "' "
    w_sSQL = w_sSQL & "  and T21_KYOKAN			= '" & Trim(m_iKyokanCd)	& "' "
    'w_sSQL = w_sSQL & "  and T21_GAKUSEKI_NO	= '" & p_GakusekiNo			& "' "
    
    if gf_ExecuteSQL(w_sSQL) <> 0 Then exit function
    
	'//正常終了
    f_AbsDelete = true
    
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
        parent.main.document.frm.action = "kks0112_bottom.asp"
	    parent.main.document.frm.submit();
	    return;
	}
	
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
	
	<input TYPE="HIDDEN" NAME="txtURL" VALUE="kks0112_bottom.asp">
    <input TYPE="HIDDEN" NAME="txtMsg" VALUE="<%=Server.HTMLEncode("再表示しています　しばらくお待ちください")%>">
	
	<input type="hidden" name="hidGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="hidClassNo" value="<%=m_sClassNo%>">
	<input type="hidden" name="hidKamokuCd" value="<%=m_sKamokuCd%>">
	
	<input type="hidden" name="txtDate" value="<%=m_sDate%>">
	<input type="hidden" name="sltJigen" value="<%=m_iJigen%>">
	
	<input type="hidden" name="hidKamokuName" value="<%=request("hidKamokuName")%>">
	<input type="hidden" name="hidClassName" value="<%=request("hidClassName")%>">
	
	<input type="hidden" name="hidSyubetu" value="<%=m_iKamokuKbn%>">
	
    </form>
    </body>
    </html>
<%
End Sub
%>