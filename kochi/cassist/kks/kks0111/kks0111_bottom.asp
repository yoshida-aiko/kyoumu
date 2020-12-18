<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠参照
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0111/kks0111_bottom.asp
' 機	  能: 下ページ 授業出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引	  数: NENDO 		 '//処理年
'			  GAKUNEN		 '//学年
'			  CLASSNO		 '//ｸﾗｽNo
' 変	  数:
' 引	  渡: NENDO 		 '//処理年
'			  GAKUNEN		 '//学年
'			  CLASSNO		 '//ｸﾗｽNo
' 説	  明:
'			■初期表示
'				検索条件にかなう行事出欠入力を表示
'			■登録ボタンクリック時
'				入力情報を登録する
'-------------------------------------------------------------------------
' 作	  成: 2002/05/07 shin
' 変	  更: 2015.03.19 kiyomoto Win7対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
	Const C_KEKKA = 1		'//欠課数
	Const C_TIKOKU = 2		'//遅刻数
	
	Const C_ZENKI = 1		'//前期
	Const C_KOUKI = 2		'//後期
	Const C_OTHER = 3		'//前期･後期以外
	
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public	m_bErrFlg		'//エラーフラグ
	
	Public m_iSyoriNen		'//処理年度
	
	Public m_sGakki 		'//学期
	Public m_sZenki_Start	'//前期開始日
	Public m_sKouki_Start	'//後期開始日
	Public m_sKouki_End 	'//後期終了日
	
	'ﾚｺｰﾄﾞセット
	Public m_Rs_Student
	Public m_Rs_Jigen
	
	Public m_StudentCount	'//生徒数
	Public m_AryKekka()		'//欠課･遅刻数セット配列
	Public m_AryKei()		'//前期･後期別の欠課･遅刻数セット配列
	
	Public m_JigenCount		'//時限数
	Public m_AryXCount		'//配列の列数
	
	Public m_sGakunenCd
	Public m_sClassCd
	Public m_sFromDate
	Public m_sToDate
	
	Public m_bDataNon		'データ存在フラグ
	
'///////////////////////////メイン処理/////////////////////////////

	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*	[機能]	本ASPのﾒｲﾝﾙｰﾁﾝ
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub Main()
	
	Dim w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget
	
	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="授業出欠入力"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = False
	m_bDataNon = false
	
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
		
		'//パラメータSET
		Call s_SetParam()
		
		'//前期・後期情報を取得
		if gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End) <> 0 then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//処理年度の時限数取得
		If not f_Get_JigenData() Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//時限情報がないとき(M07_JIGEN)
		if m_bDataNon = true then
			Call showWhitePage("時限情報がありません")
			Exit Do
		end if
		
		'//生徒、欠課、遅刻情報の取得
		If not f_Get_KekkaData() Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//生徒情報がないとき(T13_GAKU_NEN,T11_GAKUSEKI)
		if m_bDataNon = true then
			Call showWhitePage("生徒情報がありません")
			Exit Do
		end if
		
		'// データ表示ページを表示
		Call showPage()
		
		Exit Do
	Loop
	
	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
	End If
	
	'// 終了処理
	Call gf_closeObject(m_Rs_Student)
	Call gf_closeObject(m_Rs_Jigen)
	
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	変数初期化
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_ClearParam()
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	m_sFromDate = ""
	m_sToDate = ""
	
	m_iSyoriNen = 0
	
	m_sGakki	= ""
	m_sZenki_Start = ""
	m_sKouki_Start = ""
	m_sKouki_End = ""
	
End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_SetParam()
	
	m_sGakunenCd = request("cboGakunenCd")
	m_sClassCd = request("cboClassCd")
	m_sFromDate = gf_YYYY_MM_DD(request("txtFromDate"),"/")
	m_sToDate = gf_YYYY_MM_DD(request("txtToDate"),"/")
	
	m_iSyoriNen = Session("NENDO")
	
End Sub

'********************************************************************************
'*	[機能]	処理年度の時限数の取得
'*	[引数]	
'*	[戻値]	true:成功 false:失敗
'*	[説明]	
'********************************************************************************
function f_Get_JigenData()
	Dim w_sSQL
	
	On Error Resume Next
	Err.Clear
	
	f_Get_JigenData = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  MAX(M07_JIKAN) "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  M07_JIGEN "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  M07_NENDO = " & m_iSyoriNen
	
	If gf_GetRecordset(m_Rs_Jigen,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	'//データなし
	if m_Rs_Jigen.EOF then
		m_bDataNon = true
		f_Get_JigenData = true
		exit function
	end if
	
	m_JigenCount = cInt(m_Rs_Jigen(0))
	
	f_Get_JigenData = true
	
end function

'********************************************************************************
'*	[機能]	生徒情報取得,欠課･遅刻等数の取得
'*	[引数]	
'*	[戻値]	true:情報取得成功 false:失敗
'*	[説明]	
'********************************************************************************
function f_Get_KekkaData()
	
	Dim w_sSQL
	Dim w_num,w_Jnum
	Dim w_KekkaNum
	
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
	
	If gf_GetRecordset(m_Rs_Student,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	'//データなし
	if m_Rs_Student.EOF then
		m_bDataNon = true
		f_Get_KekkaData = true
		exit function
	end if
	
	m_StudentCount = gf_GetRsCount(m_Rs_Student) - 1		'生徒数
	
	m_AryXCount = 2 + m_JigenCount * 2						'(学籍NO+生徒氏名) + 時限数 * 2
	
	ReDim Preserve m_AryKekka(m_AryXCount,m_StudentCount)	'//欠課･遅刻数セット配列
	ReDim Preserve m_AryKei(4,m_StudentCount)				'//前期･後期別の欠課･遅刻数セット配列
	
	for w_num = 0 to m_StudentCount
		
		m_AryKekka(0,w_num) = m_Rs_Student(0)		'学籍NO
		m_AryKekka(1,w_num) = m_Rs_Student(1)		'生徒氏名
		
		for w_Jnum = 1 to m_JigenCount
			'w_Jnum時限の欠課数の取得
			if not f_SetKekka(m_Rs_Student(0),w_Jnum,C_KEKKA,C_OTHER,m_AryKekka(w_Jnum*2+1,w_num)) then exit function
			
			'w_Jnum時限の遅刻･早退数の取得
			if not f_SetKekka(m_Rs_Student(0),w_Jnum,C_TIKOKU,C_OTHER,m_AryKekka(w_Jnum*2+2,w_num)) then exit function
		next
		
		'前期の欠課数の取得
		if not f_SetKekka(m_Rs_Student(0),0,C_KEKKA,C_ZENKI,m_AryKei(0,w_num)) then exit function
		
		'前期の遅刻･早退数の取得
		if not f_SetKekka(m_Rs_Student(0),0,C_TIKOKU,C_ZENKI,m_AryKei(1,w_num)) then exit function
		
		'後期の欠課数の取得
		if not f_SetKekka(m_Rs_Student(0),0,C_KEKKA,C_KOUKI,m_AryKei(2,w_num)) then exit function
		
		'後期の遅刻･早退数の取得
		if not f_SetKekka(m_Rs_Student(0),0,C_TIKOKU,C_KOUKI,m_AryKei(3,w_num)) then exit function
		
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
function f_SetKekka(p_GakusekiNo,p_Jigen,p_Type,p_Kikan,p_KekkaNum)
	
	Dim w_sSQL
	Dim w_Rs
	
	On Error Resume Next
	Err.Clear
	
	f_SetKekka = false
	
	p_KekkaNum = 0
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  DECODE(sum(T21_JIKANSU),0,'',sum(T21_JIKANSU)) "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T21_SYUKKETU "
	w_sSQL = w_sSQL & " where T21_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and T21_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T21_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & "  and T21_GAKUSEKI_NO ='" & p_GakusekiNo & "'"
	
	select case p_Kikan
		case C_OTHER	'時限(前期･後期以外)
			w_sSQL = w_sSQL & "  and T21_JIGEN =" & p_Jigen
			w_sSQL = w_sSQL & "  and T21_HIDUKE >='" & m_sFromDate & "'"
			w_sSQL = w_sSQL & "  and T21_HIDUKE <='" & m_sToDate & "'"
			
		case C_ZENKI	'前期
			w_sSQL = w_sSQL & "  and T21_HIDUKE >='" & m_sZenki_Start & "'"
			w_sSQL = w_sSQL & "  and T21_HIDUKE <'"  & m_sKouki_Start & "'"
			
		case C_KOUKI	'後期
			w_sSQL = w_sSQL & "  and T21_HIDUKE >='" & m_sKouki_Start & "'"
			w_sSQL = w_sSQL & "  and T21_HIDUKE <='"  & m_sKouki_End & "'"
			
	end select
	
	if p_Type = C_KEKKA then
		'欠課数
		w_sSQL = w_sSQL & "  and (T21_SYUKKETU_KBN =" & C_KETU_KEKKA & " or T21_SYUKKETU_KBN = " & C_KETU_KEKKA_1 & ")"
	elseif p_Type = C_TIKOKU then
		'遅刻･早退数
		w_sSQL = w_sSQL & "  and (T21_SYUKKETU_KBN =" & C_KETU_TIKOKU & " or T21_SYUKKETU_KBN = " & C_KETU_SOTAI & ")"
	end if
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	p_KekkaNum = w_Rs(0)
	
	f_SetKekka = true
	
end function 

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

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	dim w_str	'表示メッセージ
	
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
		
		//スクロール同期制御
		parent.init();
		
		if(location.href.indexOf('#')==-1){
			document.frm.target = "topFrame";
			document.frm.action = "kks0111_middle.asp"
			document.frm.submit();
		}
		return;
	}
	
    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Touroku(){
		parent.frames["main"].f_Touroku();
		return;
    }
	
    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Cancel(){
        //空白ページを表示
        parent.document.location.href="default.asp"
    }
	
	//************************************************************
    //  [機能]  詳細ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //	
    //************************************************************
    function f_Detail(pGAKUSEI_NO,pName){
		var PositionX,PositionY,w_position;
		
		url = "kks0111_detail.asp";
		url = url + "?GakusekiNo=" + pGAKUSEI_NO;
		url = url + "&FromDate=<%=m_sFromDate%>";
		url = url + "&ToDate=<%=m_sToDate%>";
		url = url + "&Nen=<%=m_sGakunenCd%>";
		url = url + "&Class=<%=m_sClassCd%>";
		
		w   = 800;
		h   = 600;
		
		PositionX = window.screen.availWidth  / 2 - w / 2;
		PositionY = window.screen.availHeight / 2 - h / 2;
		
		w_position = ",left=" + PositionX + ",top=" + PositionY;
		
		opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
		if (w > 0)
			opt = opt + ",width=" + w;
		if (h > 0)
			opt = opt + ",height=" + h;
		
		opt = opt + w_position;
		
		newWin = window.open(url,"detail_subwin", opt);
	}
	
	//************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Back(){
        //空白ページを表示
        parent.document.location.href="default.asp"
    }
    
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    
		<!-- 2015.03.19 Upd Start kiyomoto-->
    	<!--<table width=800>-->
    	<table>
		<!-- 2015.03.19 Upd End kiyomoto-->
        <tr>
        	<td valign="top" align="center" nowrap>
        		<table class="hyo"  border="1">
        			
        			<%dim i%>
        			<%for i=0 to m_StudentCount%>
						<% Call gs_cellPtn(w_Class) %>
        					<tr>
					            <td class="<%=w_Class%>" align="center" width="100" height="27" nowrap><%=m_AryKekka(0,i)%></td>
					            <td class="<%=w_Class%>" align="left" width="100" height="27" nowrap><%=trim(m_AryKekka(1,i))%></td>
					            <td class="<%=w_Class%>" align="center" width="50" height="27" nowrap><input type="button" name="btnDetail" value="詳細" onClick="javascript:f_Detail('<%=m_AryKekka(0,i)%>','<%=m_AryKekka(1,i)%>');"></td>
					            
					            <% Dim j%>
					            <% for j = 3 to m_AryXCount %>
					            	<td class="<%=w_Class%>" align="center" width="20" height="27" nowrap><%=gf_SetNull2String(m_AryKekka(j,i))%></td>
		            			<% next %>	
							</tr>
		            <%next%>
            	</table>
            </td>
            
            <td width="10" height="27" valign="top" nowrap><br></td>
            
            <td align="center" width="120" valign="top" nowrap>
				
				<table width="120" class="hyo" border="1">
		            <% w_Class = "" %>
		            
		            <% Dim w_kei_num %>
		            <% for w_kei_num=0 to m_StudentCount %>
		            	<% Call gs_cellPtn(w_Class) %>
						
			            <tr>
			            	<td class="<%=w_Class%>" align="center" width="30" height="27" nowrap><%=gf_SetNull2String(m_AryKei(0,w_kei_num))%></td>
				            <td class="<%=w_Class%>" align="center" width="30" height="27" nowrap><%=gf_SetNull2String(m_AryKei(1,w_kei_num))%></td>
				            <td class="<%=w_Class%>" align="center" width="30" height="27" nowrap><%=gf_SetNull2String(m_AryKei(2,w_kei_num))%></td>
				            <td class="<%=w_Class%>" align="center" width="30" height="27" nowrap><%=gf_SetNull2String(m_AryKei(3,w_kei_num))%></td>
			            </tr>
		            	
		            <% next %>
	            </table>
			</td>
            
        </tr>
        
        </table>
		
		<table>
			<tr>
				<td align="center" nowrap>
					<input class="button" type="button" onclick="javascript:f_Back();" value=" 戻　る ">
				</td>
			</tr>
	    </table>
		
	<INPUT type="hidden" name="txtFromDate"	value = "<%=m_sFromDate%>">
	<INPUT type="hidden" name="txtToDate"	value = "<%=m_sToDate%>">
	<INPUT type="hidden" name="cboGakunenCd"	value = "<%=m_sGakunenCd%>">
	<INPUT type="hidden" name="cboClassCd"	value = "<%=m_sClassCd%>">
	
	<INPUT type="hidden" name="JigenCount"  value="<%=m_JigenCount%>">
	
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>