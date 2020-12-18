<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0111/kks0111_detail_bottom.asp
' 機      能: 下ページ 授業出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             TUKI           '//月
' 変      数:
' 引      渡: NENDO          '//処理年
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
' 説      明:
'           ■初期表示
'               検索条件にかなう行事出欠入力を表示
'           ■登録ボタンクリック時
'               入力情報を登録する
'-------------------------------------------------------------------------
' 作      成: 2002/05/07 shin
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public m_bErrFlg			'//エラーフラグ
	Public m_bNoDataFlg			'//データなしフラグ
	
	'取得したデータを持つ変数
	Public m_iSyoriNen      '//処理年度
	
	Public m_sGakunenCd
	Public m_sClassCd
	Public m_sFromDate
	Public m_sToDate
	Public m_sGakusekiNo
	
	Public m_sKamokuCd		'//科目コード
	
	Public m_sSyubetu		'//種別
	Public m_iMonth
	
	Public m_AryJigen()		'//
	Public m_AryState()		
	
	Public m_Count
	
	Public m_Rs				'//レコードセット
	Public m_JigenCount
	Public m_AryXCount
	Public m_StudentCount
	
	Public m_sGakki
	Public m_sZenki_Start
	Public m_sKouki_Start
	Public m_sKouki_End
	
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
	m_bNoDataFlg = false
	
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
		
		'// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
		
		'//前期・後期情報を取得
		if gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End) <> 0 then
			m_bErrFlg = True
			exit do
		end if
		
		'//時限情報の取得
		if not f_HeadData() then
			m_bErrFlg = True
			exit do
		end if
		
		'//時限情報がないとき
		if m_bNoDataFlg = true then
			Call showWhitePage("選択された条件での、授業情報はありません")
			exit do
		end if
		
		'//生徒欠課情報の取得
		if not f_Get_KekkaData() then
			m_bErrFlg = True
			exit do
		end if
		
		'//生徒情報がないとき
		if m_bNoDataFlg = true then
			Call showWhitePage("生徒情報がありません")
			exit do
		end if
		
		Call showPage()
		
        Exit Do
    Loop
	
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gf_closeObject(m_Rs)
    
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
    
    m_sSyubetu = ""
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
    
    m_sSyubetu = request("hidSyubetu")
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
	Dim w_iRet
	Dim w_sSDate,w_sEDate
	Dim w_num
	
	On Error Resume Next
	Err.Clear
	
	f_HeadData = false
	
	w_sSDate = ""
	w_sEDate = ""
	
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
	
	w_sSQL = w_sSQL & " group by T21_HIDUKE,T21_JIGEN"
	w_sSQL = w_sSQL & " order by T21_HIDUKE,T21_JIGEN"
	
	w_iRet = gf_GetRecordset(m_Rs,w_sSQL)
	
	if w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	m_JigenCount = gf_GetRsCount(m_Rs)
	
	if m_JigenCount = 0 then
		m_bNoDataFlg = true
		f_HeadData = true
		exit function
	end if
	
	ReDim Preserve m_AryJigen(m_JigenCount,3)
	
	for w_num = 0 to m_JigenCount-1
		
		m_AryJigen(w_num,0) = day(m_Rs("T21_HIDUKE"))
		m_AryJigen(w_num,1) = m_Rs("T21_JIGEN")
		m_AryJigen(w_num,2) = m_Rs("T21_HIDUKE")
		
		m_Rs.movenext
	next
	
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
	Dim w_iNen
	
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
'*	[機能]	生徒情報取得,欠課･遅刻等数の取得
'*	[引数]	
'*	[戻値]	true:情報取得成功 false:失敗
'*	[説明]	
'********************************************************************************
function f_Get_KekkaData()
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_num,w_Jnum
	
	On Error Resume Next
	Err.Clear
	
	f_Get_KekkaData = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T13.T13_GAKUSEKI_NO,"
	w_sSQL = w_sSQL & "  T11.T11_SIMEI "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T13_GAKU_NEN T13,"
	w_sSQL = w_sSQL & "  T11_GAKUSEKI T11 "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  T13.T13_GAKUSEI_NO = T11.T11_GAKUSEI_NO "
	w_sSQL = w_sSQL & "  and T13.T13_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and T13.T13_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T13.T13_CLASS =" & m_sClassCd
	
	w_sSQL = w_sSQL & "  group by T11.T11_SIMEI,T13.T13_GAKUSEKI_NO "
	w_sSQL = w_sSQL & "  order by T13.T13_GAKUSEKI_NO "
	
	w_iRet = gf_GetRecordset(m_Rs_Student,w_sSQL)
	
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	m_StudentCount = gf_GetRsCount(m_Rs_Student)		'//生徒数
	
	'//データなし
	if m_StudentCount = 0 then
		m_bNoDataFlg = true
		f_Get_KekkaData = true
		exit function
	end if
	
	m_AryXCount = 2 + m_JigenCount						'//(学籍NO+生徒氏名) + 時限数
	
	ReDim Preserve m_AryState(m_AryXCount,m_StudentCount)	'//欠課･遅刻数セット配列
	
	for w_num = 0 to m_StudentCount
		
		m_AryState(0,w_num) = m_Rs_Student(0)		'学籍NO
		m_AryState(1,w_num) = m_Rs_Student(1)		'生徒氏名
		
		for w_Jnum = 0 to m_JigenCount - 1
			'w_Jnum時限の欠課区分
			if not f_SetKekka(m_Rs_Student(0),m_AryJigen(w_Jnum,1),m_AryJigen(w_Jnum,2),m_AryState(2+w_Jnum,w_num)) then exit function
		next
		
		m_Rs_Student.movenext
		
	next
	
	f_Get_KekkaData = true
	
end function

'********************************************************************************
'*	[機能]	欠課･遅刻数の取得
'*	[引数]	p_GakusekiNo→学籍NO
'*			p_Jigen→時限
'*			p_Type→C_KEKKA:欠課数,C_TIKOKU:遅刻数
'*			p_Kikan→C_ZENKI:前期,C_KOUKI:後期,C_KOUKI:前期･後期以外
'*
'*	[戻値]	0:情報取得成功 99:失敗、p_KekkaNum→欠課･遅刻数
'*	[説明]	
'********************************************************************************
function f_SetKekka(p_GakusekiNo,p_Jigen,p_Hiduke,p_KekkaType)
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_Rs
	Dim w_KekkaName
	
	On Error Resume Next
	Err.Clear
	
	f_SetKekka = false
	
	p_KekkaNum = 0
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T21_SYUKKETU_KBN, "
	w_sSQL = w_sSQL & "  T21_JIKANSU, "
	w_sSQL = w_sSQL & "  M01_SYOBUNRUIMEI_R "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T21_SYUKKETU, "
	w_sSQL = w_sSQL & "  M01_KUBUN "
	
	w_sSQL = w_sSQL & " where T21_NENDO = " & m_iSyoriNen
	
	w_sSQL = w_sSQL & "  and M01_DAIBUNRUI_CD =" & C_KESSEKI
	w_sSQL = w_sSQL & "  and M01_NENDO =" & m_iSyoriNen
	w_sSQL = w_sSQL & "  and T21_SYUKKETU_KBN = M01_SYOBUNRUI_CD(+) "
	
	w_sSQL = w_sSQL & "  and T21_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T21_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & "  and T21_GAKUSEKI_NO ='" & p_GakusekiNo & "'"
	w_sSQL = w_sSQL & "  and T21_JIGEN =" & p_Jigen
	w_sSQL = w_sSQL & "  and T21_HIDUKE ='" & p_Hiduke & "'"
	
	w_iRet = gf_GetRecordset(w_Rs,w_sSQL)
	
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	'Dim w_IdouType,w_KubunName
	'w_IdouType = cint(gf_SetNull2Zero(gf_Get_IdouChk(p_GakusekiNo,p_Hiduke,m_iSyoriNen,w_KubunName)))
	
	if cInt(gf_SetNull2Zero(w_Rs(0))) <> cInt(C_KETU_KEKKA) then
		p_KekkaType = w_Rs(2)
	else
		p_KekkaType = w_Rs(1) & w_Rs(2)
	end if
	
	f_SetKekka = true
	
end function 

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	Dim w_Class
	Dim w_Jnum,w_num
	
	w_Class = ""
	
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
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    	<table>
			<tr>
				<td valign="top" nowrap>
					<table class="hyo"  border="1">
						<% for w_num = 0 to m_StudentCount-1 %>
							<% Call gs_cellPtn(w_Class) %>
							<tr>
								<td width="80" class="<%=w_Class%>" align="center" nowrap><%=m_AryState(0,w_num)%></td>
								<td width="130"  class="<%=w_Class%>" align="center" nowrap><%=m_AryState(1,w_num)%></td>
								
								<%for w_Jnum = 0 to m_JigenCount-1 %>
									<td width="40" class="<%=w_Class%>" align="center" nowrap><%=gf_HTMLTableSTR(m_AryState(2+w_Jnum,w_num))%></td>
								<%next%>
							</tr>
							
						<% next%>
						
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
	<body LANGUAGE="javascript" onload="return window_onload()">
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
