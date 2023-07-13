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
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	
	'取得したデータを持つ変数
	Public m_iSyoriNen      '//処理年度
	
	Public m_sGakunenCd
	Public m_sClassCd
	Public m_sFromDate
	Public m_sToDate
	Public m_sGakusekiNo
	
	Public m_AryKekkaMei()  '//欠課名称格納配列
	
	Public m_Count
	
	Public m_Rs				'//レコードセット
	Public m_bDataNon		'//詳細データフラグ
	Public m_RecCount		'//レコード件数
	
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
		
		if not f_SetKekka() then
			m_bErrFlg = True
			Exit Do
		end if
		
		'詳細データがなし
		if m_bDataNon = true then
			Call showWhitePage("詳細データは、ありません")
			exit do
		end if
		
		if not f_Get_SyukketuKbn() then
			m_bErrFlg = True
			Exit Do
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
	m_sGakunenCd = ""
	m_sClassCd = ""
	m_sGakusekiNo = ""
	
	m_Count = 0
	
    m_iSyoriNen = ""
    
	m_bDataNon = false
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	m_sGakunenCd = request("Nen")
	m_sClassCd = request("Class")
	
	m_sFromDate = request("FromDate")
	m_sToDate = request("ToDate")
	m_sGakusekiNo = request("GakusekiNo")
	
    m_iSyoriNen = Session("NENDO")
    
End Sub

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
function f_SetKekka()
	
	Dim w_sSQL
	Dim w_iRet
	
	On Error Resume Next
	Err.Clear
	
	f_SetKekka = false
	
	p_KekkaNum = 0
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T21_HIDUKE, "
	w_sSQL = w_sSQL & "  T21_JIGEN, "
	w_sSQL = w_sSQL & "  T21_KYOKAN, "
	
	w_sSQL = w_sSQL & "  M03_KAMOKUMEI, "
	w_sSQL = w_sSQL & "  M04_KYOKANMEI_SEI, "
	w_sSQL = w_sSQL & "  M04_KYOKANMEI_MEI, "
	w_sSQL = w_sSQL & "  T21_SYUKKETU_KBN, "
	w_sSQL = w_sSQL & "  T21_JIKANSU "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T21_SYUKKETU, "
	w_sSQL = w_sSQL & "  M03_KAMOKU, "
	w_sSQL = w_sSQL & "  M04_KYOKAN "
	
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "      T21_SYUKKETU.T21_KAMOKU = M03_KAMOKU.M03_KAMOKU_CD "
	w_sSQL = w_sSQL & "  and T21_SYUKKETU.T21_KYOKAN = M04_KYOKAN.M04_KYOKAN_CD(+) "
	
	w_sSQL = w_sSQL & "  and T21_SYUKKETU.T21_NENDO = M03_KAMOKU.M03_NENDO "
	w_sSQL = w_sSQL & "  and T21_SYUKKETU.T21_NENDO = M04_KYOKAN.M04_NENDO "
	
	w_sSQL = w_sSQL & "  and T21_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T21_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & "  and T21_GAKUSEKI_NO ='" & m_sGakusekiNo & "'"
	
	w_sSQL = w_sSQL & "  and T21_HIDUKE >='" & m_sFromDate & "'"
	w_sSQL = w_sSQL & "  and T21_HIDUKE <='" & m_sToDate & "'"
	
	w_sSQL = w_sSQL & "  and (T21_SYUKKETU_KBN =" & C_KETU_TIKOKU
	w_sSQL = w_sSQL & "       or T21_SYUKKETU_KBN = " & C_KETU_SOTAI
	w_sSQL = w_sSQL & "       or T21_SYUKKETU_KBN = " & C_KETU_KEKKA
	w_sSQL = w_sSQL & "      )"
	
	w_sSQL = w_sSQL & " order by T21_HIDUKE,T21_JIGEN "
	
	w_iRet = gf_GetRecordset(m_Rs,w_sSQL)
	
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	'データがない
	if m_Rs.EOF then m_bDataNon = true
	
	m_RecCount = gf_GetRsCount(m_Rs)
	
	f_SetKekka = true
	
end function 

'********************************************************************************
'*	[機能]	出欠区分名称の取得(配列にセット)
'*	[引数]	なし
'*	[戻値]	
'*	[説明]	
'********************************************************************************
function f_Get_SyukketuKbn()
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_Rs_Kekka
	
	On Error Resume Next
	Err.Clear
	
	f_Get_SyukketuKbn = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  M01_SYOBUNRUIMEI,M01_SYOBUNRUI_CD "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  M01_KUBUN "
	
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "      M01_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and M01_DAIBUNRUI_CD = " & C_KESSEKI
	w_sSQL = w_sSQL & " order by  M01_SYOBUNRUI_CD "
	
	w_iRet = gf_GetRecordset(w_Rs_Kekka,w_sSQL)
	
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	m_Count = gf_GetRsCount(w_Rs_Kekka) - 1
	
	ReDim Preserve m_AryKekkaMei(2,m_Count)
	
	for w_num = 0 to m_Count
		m_AryKekkaMei(0,w_num) = w_Rs_Kekka("M01_SYOBUNRUI_CD")
		m_AryKekkaMei(1,w_num) = w_Rs_Kekka("M01_SYOBUNRUIMEI")
		w_Rs_Kekka.movenext
	Next
	
	f_Get_SyukketuKbn = true
	
end function 

'********************************************************************************
'*	[機能]	出欠名称取得
'*	[引数]	p_SyukketuKbn:出欠区分
'*	[戻値]	出欠名称
'*	[説明]	
'********************************************************************************
function f_Set_SyukketuMei(p_SyukketuKbn)
	Dim w_num
	
	for w_num = 0 to m_Count
		
		if cInt(m_AryKekkaMei(0,w_num)) = cInt(p_SyukketuKbn) then
			f_Set_SyukketuMei = m_AryKekkaMei(1,w_num)
			exit function
		end if
		
	Next
	
	f_Set_SyukketuMei = ""
	
end function


'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	Dim w_Name			'名前
	Dim w_SyukketuName	'出欠名称
	Dim w_Class			'td classセット
	
	w_Name = ""
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
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    	<table>
	        <tr>
	        	<td>
	        		<table width="540" class="hyo"  border="1" >
	        			<% if not (m_Rs is nothing) then %>
					    	<% Do until m_Rs.EOF %>
								<% Call gs_cellPtn(w_Class) %>
								
								<tr>
									<td width="130" class="<%=w_Class%>" align="center" nowrap><%=m_Rs("T21_HIDUKE")%></td>
									<td width="60"  class="<%=w_Class%>" align="center" nowrap><%=m_Rs("T21_JIGEN")%></td>
									<td width="150" class="<%=w_Class%>" align="center" nowrap><%=m_Rs("M03_KAMOKUMEI")%></td>
									
									<%
										w_Name = ""
										
										if cstr(m_Rs("T21_KYOKAN")) = "1" then
											w_Name = "教務"
										else
											w_Name = m_Rs("M04_KYOKANMEI_SEI") & "　" & m_Rs("M04_KYOKANMEI_MEI")
										end if
									%>
									
									<td width="120" class="<%=w_Class%>" align="center" nowrap><%=w_Name%></td>
									
									<%
										w_SyukketuName = ""
										
										if cInt(m_Rs("T21_SYUKKETU_KBN")) = cInt(C_KETU_KEKKA) then
											w_SyukketuName = gf_SetNull2String(m_Rs("T21_JIKANSU")) & f_Set_SyukketuMei(m_Rs("T21_SYUKKETU_KBN"))
										else
											w_SyukketuName = f_Set_SyukketuMei(m_Rs("T21_SYUKKETU_KBN"))
										end if
									%>
									
						            <td width="80" class="<%=w_Class%>" align="center" nowrap><%=w_SyukketuName%></td>
			            		</tr>
			            		
			            		<% m_Rs.movenext %>
			            	<% Loop %>
						<% end if %>
	            		
	            	</table>
	            </td>
	        </tr>
	        
	        <% if m_RecCount >= 20 then %><!-- 20件以上だと表示する -->
	        <tr>
				<td align="center"><input type="button" value="閉じる" onClick="javascript:parent.close();"></td>
			</tr>
			<% end if %>
			
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
