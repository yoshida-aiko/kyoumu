<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠表示(詳細)
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0112/kks0112_subwin_top.asp
' 機      能: 上ページ 授業出欠覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             
' 変      数:
' 引      渡: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
' 説      明:
'            
'-------------------------------------------------------------------------
' 作      成: 2002/05/07 shin
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	
	Public m_bErrFlg		'//エラーフラグ
	
	Public m_iSyoriNen		'//処理年度
	Public m_sGakunenCd		'//学年
	Public m_sClassCd		'//クラスCD
	
	Public m_sKamokuCd		'//科目コード
    
    Public m_iMonth			'//月
	
	Public m_sGakki
	Public m_sZenki_Start
	Public m_sKouki_Start
	Public m_sKouki_End
	
	Public m_Rs				'//レコードセット
	
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
            w_sMsg = "データベースとの接続に失敗しました。"
            'm_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//変数初期化
		Call s_ClearParam()
		
		'// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
		
		'//前期・後期情報を取得
		if gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End) <> 0 then
			m_bErrFlg = True
			exit do
		end if
		
		'//ヘッダ情報取得
		if not f_HeadData() then
			m_bErrFlg = True
			exit do
		end if
		
		if m_Rs.EOF then
			'Call showWhitePage("対象となる、授業情報がありません")
			exit do
		end if
		
		'//ページ表示
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
	
	m_iSyoriNen = 0
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	
    m_sKamokuCd = ""
    
    m_iMonth = ""
	
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen = Session("NENDO")
	
	m_sGakunenCd = request("hidGakunen")
	m_sClassCd = request("hidClassNo")
	
    m_sKamokuCd = request("hidKamokuCd")
    
    m_iMonth = request("sltMonth")
	
End Sub

'********************************************************************************
'*  [機能]  ヘッダ情報取得
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
function f_HeadData()
	Dim w_sSQL
	Dim w_sSDate,w_sEDate
	
	On Error Resume Next
	Err.Clear
	
	f_HeadData = false
	
	w_sSDate = ""
	w_sEDate = ""
	
	'//指定月の開始日、終了日をゲット
	Call f_GetTukiRange(w_sSDate,w_sEDate)
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T21_HIDUKE, "
	w_sSQL = w_sSQL & "  T21_JIGEN "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T21_SYUKKETU "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "     T21_HIDUKE >='" & w_sSDate & "' "
	w_sSQL = w_sSQL & " and T21_HIDUKE <'" & w_sEDate & "' "
	w_sSQL = w_sSQL & " and T21_NENDO =" & m_iSyoriNen 
	w_sSQL = w_sSQL & " and T21_GAKUNEN =" & m_sGakunenCd 
	w_sSQL = w_sSQL & " and T21_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & " and T21_KAMOKU ='" & m_sKamokuCd & "'"
	w_sSQL = w_sSQL & " and T21_SYUKKETU_KBN in(1,2,3)"
	
	w_sSQL = w_sSQL & " group by T21_HIDUKE,T21_JIGEN "
	w_sSQL = w_sSQL & " order by T21_HIDUKE,T21_JIGEN "
	
	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	m_Rs.movefirst
	
	f_HeadData = true
	
end function


'********************************************************************************
'*	[機能]	月の検索条件を作成(7月…　"MONTH>=2001/07/01 AND MONTH<2001/08/01" として使用)
'*	[引数]	なし
'*	[戻値]	p_sSDate
'*			p_sEDate
'*	[説明]	
'********************************************************************************
Function f_GetTukiRange(p_sSDate,p_sEDate)
	
	p_sSDate = ""
	p_sEDate = ""
	
	if 4 <= cInt(m_iMonth) and cInt(m_iMonth) <=12 then
		w_iNen = cint(m_iSyoriNen)

		'//開始日
		If cint(month(m_sZenki_Start)) = Cint(m_iMonth) Then
			p_sSDate = m_sZenki_Start
		Else
			p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iMonth),2) & "/01"
		End If
	
		'//終了日
		If cint(month(m_sKouki_Start)) = Cint(m_iMonth) Then
			p_sEDate = m_sKouki_Start
		Else 
			If Cint(m_iMonth) = 12 Then
				p_sEDate = cstr(w_iNen+1) & "/01/01"
			Else
				p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iMonth+1),2) & "/01"
			End If
		End If
		
	Else
		'//後期の年
		If cint(m_iMonth) <=4 Then
			w_iNen = cint(m_iSyoriNen) + 1
		Else
			w_iNen = cint(m_iSyoriNen)
		End If

		'//開始日
		If cint(month(m_sKouki_Start)) = Cint(m_iMonth) Then
			p_sSDate = m_sKouki_Start
		Else
			p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iMonth),2) & "/01"
		End If

		'//終了日
		If cint(month(m_sKouki_End)) = Cint(m_iMonth) Then
			'p_sEDate = m_sKouki_End
			p_sEDate = DateAdd("d",1,m_sKouki_End)
		Else 
			If Cint(m_sTuki) = 12 Then
				p_sEDate = cstr(w_iNen+1) & "/01/01"
			Else
				p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iMonth+1),2) & "/01"
			End If
		End If

	End If

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	On Error Resume Next
    Err.Clear
%>
    <html>
    <head>
    <title>授業出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
	
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
	
    //************************************************************
    //  [機能]  ページロード時処理
    //************************************************************
    function window_onload() {
		//スクロール同期制御
		//parent.init();
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    	<BR>
		
		<table>
			<tr>
				<td valign="bottom" nowrap>
					<table class="hyo" border="1">
						<tr>
							<th width="80" class="header" rowspan="2" nowrap>学籍番号</th>
							<th width="80" class="header" rowspan="2" nowrap>氏名</th>
							
							<% if not (m_Rs is nothing) then %>
								
								<th width="40" class="header" align="center" nowrap>日付</th>
								
								<% do until m_Rs.EOF %>
									<th width="40" class="header" align="center" nowrap><%=day(m_Rs("T21_HIDUKE"))%></th>
									
									<% m_Rs.movenext %>
								<% loop %>
							<% end if %>
							
						</tr>
						
						<tr>
							<% if not (m_Rs is nothing) then %>
								<% m_Rs.movefirst %>
								
								<th width="40" class="header" align="center" nowrap>時限</th>
								
								<% do until m_Rs.EOF %>
									<th width="40" class="header" align="center" nowrap><%=m_Rs("T21_JIGEN")%></th>
									
									<% m_Rs.movenext %>
								<% loop %>
							<% end if %>	
						</tr>
					</table>
				</td>
			</tr>
		</table>
    </form>
    </center>
    </body>
    </html>
<%
End Sub

'********************************************************************************
'*	[機能]	空白HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
	<html>
	<head>
	<title>授業出欠入力</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//	[機能]	ページロード時処理
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function window_onload() {
	}
	//-->
	</SCRIPT>

	</head>
	<body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" mothod="post">

	<center>
	<br><br><br>
		<span class="msg"><%=Server.HTMLEncode(p_Msg)%></span>
	</center>

	<input type="hidden" name="txtMsg" value="<%=Server.HTMLEncode(p_Msg)%>">
	</form>
	</body>
	</html>
<%
End Sub
%>
