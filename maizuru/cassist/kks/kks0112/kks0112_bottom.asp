<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0112/kks0112_bottom.asp
' 機	  能: 下ページ 授業出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引	  数: 
'			  
'			  
'			  
'			  
' 変	  数: 
' 引	  渡: 
'			  
'			  
'			  
'			  
' 説	  明:
'			■初期表示
'				検索条件にかなう生徒一覧を表示
'			■登録ボタンクリック時
'				入力情報を登録する
'			■戻るボタンクリック時
'				前ページに戻る
'-------------------------------------------------------------------------
' 作	  成: 2002/05/16 shin
' 変	  更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public	m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ

	'取得したデータを持つ変数
	Public m_iSyoriNen		'//処理年度
	Public m_sGakunen		'//学年
	Public m_sClassNo		'//ｸﾗｽNO
	
	Public m_sKamokuCd		'//課目CD
	Public m_iKamokuKbn		'//授業種別(C_JIK_JUGYO)
	
	'ﾚｺｰﾄﾞセット
	Public m_Rs_Student		'//recordset生徒
	
	Public m_sDate
	Public m_iJigen
	
	Public m_Count
	
	Const C_SELECT = 1	'//選択科目で選択
	
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
		
		'// ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()
		
		'//生徒情報取得
		if not f_Get_Student() Then
			m_bErrFlg = True
			Exit Do
		End If
		
		if m_Rs_Student.EOF Then
			Call showWhitePage("対象となる、生徒情報がありません")
			Exit Do
		End If
		
		'// データ表示ページを表示
		Call showPage_bottom()
		
		Exit Do
	Loop
	
	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
	End If
	
	'// 終了処理
	Call gf_closeObject(m_Rs_Student)
	
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	変数初期化
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_ClearParam()

	m_iSyoriNen = 0
	
	m_sGakunen	= 0
	m_sClassNo	= 0
	
	m_sKamokuCd = ""
	m_iKamokuKbn	= 0
	
	m_sDate = ""
	m_iJigen = ""
	
End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen = Session("NENDO")
	
	m_iKamokuKbn	= cint((Request("hidSyubetu")))
	
	m_sGakunen	= trim(Request("hidGakunen"))
	m_sClassNo	= trim(Request("hidClassNo"))
	m_sKamokuCd = trim(Request("hidKamokuCd"))
	
	m_sDate = gf_YYYY_MM_DD(trim(Request("txtDate")),"/")
	m_iJigen = trim(Request("sltJigen"))
	
End Sub

'********************************************************************************
'*	[機能]	生徒情報取得,欠課･遅刻等数の取得
'*	[引数]	
'*	[戻値]	true:情報取得成功 false:失敗
'*	[説明]	
'********************************************************************************
function f_Get_Student()
	Dim w_sSQL
	
	On Error Resume Next
	Err.Clear
	
	f_Get_Student = false
	
	if m_iKamokuKbn = C_JIK_JUGYO then
		'//通常授業のとき
		w_sSQL = ""
		w_sSQL = w_sSQL & " select "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & " 	T11.T11_SIMEI,"
		w_sSQL = w_sSQL & " 	T21.T21_JIKANSU, "
		w_sSQL = w_sSQL & " 	T21.T21_SYUKKETU_KBN, "
		w_sSQL = w_sSQL & " 	T16.T16_HISSEN_KBN as HISSEN_KBN , "
		w_sSQL = w_sSQL & " 	T16.T16_SELECT_FLG as SELECT_FLG "
		
		w_sSQL = w_sSQL & " from "
		w_sSQL = w_sSQL & " 	T13_GAKU_NEN T13, "
		w_sSQL = w_sSQL & " 	T11_GAKUSEKI T11, "
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN T16 ,"
		w_sSQL = w_sSQL & " 	( "
		w_sSQL = w_sSQL & " 	 select * from T21_SYUKKETU "
		w_sSQL = w_sSQL & " 	 where T21_HIDUKE ='" & m_sDate & "' "
		w_sSQL = w_sSQL & " 	 and T21_JIGEN =" & m_iJigen & " ) T21 "
		
		w_sSQL = w_sSQL & " where "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEI_NO = T11.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " and T13.T13_GAKUSEI_NO = T16.T16_GAKUSEI_NO "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = T21.T21_NENDO(+) "
		w_sSQL = w_sSQL & " and T13.T13_GAKUNEN = T21.T21_GAKUNEN(+) "
		w_sSQL = w_sSQL & " and T13.T13_CLASS = T21.T21_CLASS(+) "
		w_sSQL = w_sSQL & " and T13.T13_GAKUSEKI_NO = T21.T21_GAKUSEKI_NO(+) "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = T16.T16_NENDO "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & " and T13.T13_GAKUNEN = " & m_sGakunen
		w_sSQL = w_sSQL & " and T13.T13_CLASS = " & m_sClassNo
		w_sSQL = w_sSQL & " and T16.T16_KAMOKU_CD ='" & m_sKamokuCd & "'"
		
		w_sSQL = w_sSQL & " group by "
		w_sSQL = w_sSQL & " 	T11.T11_SIMEI,"
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO,"
		w_sSQL = w_sSQL & " 	T21.T21_JIKANSU,"
		w_sSQL = w_sSQL & " 	T21.T21_SYUKKETU_KBN,"
		w_sSQL = w_sSQL & " 	T16.T16_HISSEN_KBN,"
		w_sSQL = w_sSQL & " 	T16.T16_SELECT_FLG "
		w_sSQL = w_sSQL & " order by "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO "
	else
		'//特別活動
		w_sSQL = ""
		w_sSQL = w_sSQL & " select "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & " 	T11.T11_SIMEI,"
		w_sSQL = w_sSQL & " 	T21.T21_JIKANSU, "
		w_sSQL = w_sSQL & " 	T21.T21_SYUKKETU_KBN, "
		w_sSQL = w_sSQL & " 	1 as HISSEN_KBN , "
		w_sSQL = w_sSQL & " 	0 as SELECT_FLG "
		
		w_sSQL = w_sSQL & " from "
		w_sSQL = w_sSQL & " 	T13_GAKU_NEN T13, "
		w_sSQL = w_sSQL & " 	T11_GAKUSEKI T11, "
		w_sSQL = w_sSQL & " 	T34_RISYU_TOKU T34 ,"
		w_sSQL = w_sSQL & " 	( "
		w_sSQL = w_sSQL & " 	 select * from T21_SYUKKETU "
		w_sSQL = w_sSQL & " 	 where T21_HIDUKE ='" & m_sDate & "' "
		w_sSQL = w_sSQL & " 	 and T21_JIGEN =" & m_iJigen & " ) T21 "
		
		w_sSQL = w_sSQL & " where "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEI_NO = T11.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " and T13.T13_GAKUSEI_NO = T34.T34_GAKUSEI_NO "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = T21.T21_NENDO(+) "
		w_sSQL = w_sSQL & " and T13.T13_GAKUNEN = T21.T21_GAKUNEN(+) "
		w_sSQL = w_sSQL & " and T13.T13_CLASS = T21.T21_CLASS(+) "
		w_sSQL = w_sSQL & " and T13.T13_GAKUSEKI_NO = T21.T21_GAKUSEKI_NO(+) "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = T34.T34_NENDO "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & " and T13.T13_GAKUNEN = " & m_sGakunen
		w_sSQL = w_sSQL & " and T13.T13_CLASS = " & m_sClassNo
		w_sSQL = w_sSQL & " and T34.T34_TOKUKATU_CD ='" & m_sKamokuCd & "'"
		
		w_sSQL = w_sSQL & " group by "
		w_sSQL = w_sSQL & " 	T11.T11_SIMEI,"
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO,"
		w_sSQL = w_sSQL & " 	T21.T21_JIKANSU,"
		w_sSQL = w_sSQL & " 	T21.T21_SYUKKETU_KBN "
		w_sSQL = w_sSQL & " order by "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO "
		
	end if
	
	If gf_GetRecordset(m_Rs_Student,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	f_Get_Student = true
	
end function

'********************************************************************************
'*	[機能]	出欠名称取得
'*	[引数]	p_SyukketuKbn:出欠区分
'*	[戻値]	出欠名称
'*	[説明]	
'********************************************************************************
function f_Set_SyukketuMei(p_SyukketuKbn,p_JikanNum)
	Dim w_num
	Dim w_KubunName
	
	if gf_SetNull2String(p_SyukketuKbn) = "" then 
		f_Set_SyukketuMei = ""
		exit function
	end if
	
	'//出席区分名取得
	if not gf_GetKubunName(19,p_SyukketuKbn,m_iSyoriNen,w_KubunName) then
		f_Set_SyukketuMei = ""
		exit function
	end if
	
	if cint(p_SyukketuKbn) = 1 then
		f_Set_SyukketuMei = p_JikanNum & w_KubunName
	else
		f_Set_SyukketuMei = w_KubunName
	end if
	
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
	</head>
	
	<body LANGUAGE="javascript">
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
'*	[機能]	HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub showPage_bottom()
	Dim w_Class
	Dim w_num
	Dim w_SyukketuName
	Dim w_SelectFlg
	Dim w_HissenFlg
	Dim w_IdouType,w_KubunName
	
	w_num = 1
	
	On Error Resume Next
	Err.Clear
	
%>
	<html>
	<head>
	<title>行事用出欠入力</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<!--#include file="../../Common/jsCommon.htm"-->
	
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	
	//************************************************************
	//	[機能]	ページロード時処理
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function window_onload() {
		//スクロール同期制御
		parent.init();
		
		if(location.href.indexOf('#')==-1){
			//ヘッダ部を表示submit
			document.frm.target = "topFrame";
			document.frm.action = "kks0112_middle.asp?<%=Request.Form.Item%>"
			document.frm.submit();
		}
		
		return;
		
	}
	
	//************************************************************
	//	[機能]	登録ボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Insert(){
		if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
		   return false;
		}
		
		//ヘッダ部空白表示
		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("登録しています・・・・　　しばらくお待ちください")%>"
		
		//リスト情報をsubmit
		document.frm.target = "main";
		document.frm.action = "kks0112_edit.asp";
		document.frm.submit();
	}
	
    //************************************************************
    //  [機能]  戻るボタンが押されたとき
    //  [引数]  
    //  [戻値]  
    //  [説明]
    //************************************************************
    function f_Back(){
        //空白ページを表示
        parent.document.location.href="default.asp";
    }
	
	//************************************************************
	//	[機能]	出欠入力
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//************************************************************
	function chg(chgInp){
		var w_num,i,w_StateValue,w_KekkaNum;
		var w_txtKekka;
		w_txtKekka = eval("parent.topFrame.document.frm.txtKekka");
		
		w_num = 0;
		w_KekkaNum = w_txtKekka.value;
		
		//どのラジオボタンが選択されているかチェック
		for(i=0;i<4;i++){
			if(parent.topFrame.document.frm.rdoType[i].checked == true){
				w_num = i + 1;
				if(w_num == 1 && w_KekkaNum == ""){
					f_InpChkErr("欠課数を入力して下さい",w_txtKekka);
					return false;
				}else if(w_num == 1 && (w_KekkaNum < 1 || w_KekkaNum > 9 || isNaN(w_KekkaNum))){
					f_InpChkErr("欠課数が不正です",w_txtKekka);
					return false;
				}
				break;
			}
		}
		
		switch(w_num){
			case 0 : alert("入力したい「入力区分」を選択後、該当する学生の出欠状況覧をクリックして下さい。");
					 return false;
					 break;
					 
			case 1 : w_StateValue = w_KekkaNum + "欠課";
					 break;
					 
			case 2 : w_StateValue = "遅刻";
					 w_KekkaNum = 1;
					 break;
					 
			case 3 : w_StateValue = "早退";
					 w_KekkaNum = 1;
					 break;
					 
			case 4 : w_StateValue = ""; 
					 w_KekkaNum = 0;
					 w_num = 0;
					 break;
					 
			default: break;
		}
		
		chgInp.value = w_StateValue;
		
		var ob = new Array();
		ob[0] = eval("document.frm.hid"+chgInp.name);
		ob[0].value = w_num;
		
		ob[1] = eval("document.frm.hidJikan"+chgInp.name);
		ob[1].value = w_KekkaNum;
		
	}
	
	//************************************************************
    //  [機能]  入力チェックエラー時のalert,focus,select処理
    //************************************************************
    function f_InpChkErr(p_AlertMsg,p_Object){
		alert(p_AlertMsg);
		p_Object.focus();
		p_Object.select();
	}
	
	//-->
	</SCRIPT>
	
	</head>
	<body LANGUAGE="javascript" onload="window_onload()">
	<form name="frm" method="post">
	
	<center>
		<table width="545">
			<tr>
				<td align="center" valign="top" nowrap>
					<table class="hyo"	border="1" width="300">
						
						<%
						Do until m_Rs_Student.EOF
							Call gs_cellPtn(w_Class)
							
							w_HissenFlg = cint(m_Rs_Student("HISSEN_KBN"))	'必修・選択フラグ
							w_SelectFlg = cint(m_Rs_Student("SELECT_FLG"))	'選択授業を選択しているかフラグ
							
							if (w_HissenFlg = C_HISSEN_HIS) or (w_HissenFlg = C_HISSEN_SEN and w_SelectFlg = C_SELECT) then
						%>
								<tr>
									<td class="<%=w_Class%>" width="80" align="center" nowrap><%=m_Rs_Student("T13_GAKUSEKI_NO")%></td>
									<input type="hidden" name="hidGakusekiNo" value="<%=m_Rs_Student("T13_GAKUSEKI_NO")%>">
									
									<td class="<%=w_Class%>" width="150" nowrap><%=m_Rs_Student("T11_SIMEI")%></td>
									<td class="<%=w_Class%>" width="70" height="28" align="center" nowrap>
										<% 
											'出欠状況
											w_SyukketuName = ""
											w_SyukketuName = f_Set_SyukketuMei(m_Rs_Student("T21_SYUKKETU_KBN"),m_Rs_Student("T21_JIKANSU"))
											
											'異動情報取得
											w_IdouType = cint(gf_SetNull2Zero(gf_Get_IdouChk(m_Rs_Student("T13_GAKUSEKI_NO"),m_sDate,m_iSyoriNen,w_KubunName)))
										%>
										
										<% if w_IdouType = 0 or w_IdouType = C_IDO_FUKUGAKU or w_IdouType = C_IDO_TEI_KAIJO then %>
											<input type="button" class="<%=w_Class%>" name="State<%=m_Rs_Student("T13_GAKUSEKI_NO")%>" style="border-style:none;text-align:center;" tabindex="-1" value="<%=w_SyukketuName%>" onclick="return chg(this);">
										<% else %>
											<font color="red"><%=w_KubunName%></font>
										<% end if %>
										
										<input type="hidden" name='hidState<%=m_Rs_Student("T13_GAKUSEKI_NO")%>' value='<%=gf_SetNull2Zero(m_Rs_Student("T21_SYUKKETU_KBN"))%>'>
										<input type="hidden" name='hidJikanState<%=m_Rs_Student("T13_GAKUSEKI_NO")%>' value='<%=gf_SetNull2Zero(m_Rs_Student("T21_JIKANSU"))%>'>
									</td>
								</tr>
						<%	
							end if
							
							w_num = w_num + 1
							m_Rs_Student.movenext
						loop
						
						%>
						
					</table>
				</td>
			</tr>
			
			<tr>
				<td align="center" valign="top" nowrap>
					<table>
						<tr>
							<td nowrap><input type="button" name="btnInsert" value="　登　録　" onClick="f_Insert();"></td>
							<td nowrap><input type="button" name="btnBack" value="　戻　る　" onClick="f_Back();"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</center>
	
	<input type="hidden" name="hidGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="hidClassNo" value="<%=m_sClassNo%>">
	<input type="hidden" name="hidKamokuCd" value="<%=m_sKamokuCd%>">
	
	<input type="hidden" name="hidDate" value="<%=m_sDate%>">
	<input type="hidden" name="hidJigen" value="<%=m_iJigen%>">
	
	<input type="hidden" name="hidKamokuName" value="<%=request("hidKamokuName")%>">
	<input type="hidden" name="hidClassName" value="<%=request("hidClassName")%>">
	<input type="hidden" name="hidSyubetu" value="<%=m_iKamokuKbn%>">
	
	</form>
	</body>
	</html>
<%
End Sub
%>
