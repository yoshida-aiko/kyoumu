<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0150_23_middle.asp
' 機      能: 下ページ 成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:
'           :
' 変      数:
' 引      渡:
'           :
' 説      明:
'           ■初期表示
'				
'			■表示ボタンクリック時
'				
'-------------------------------------------------------------------------
' 作      成: 2003/05/01 廣田
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Dim  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	
	'//氏名選択用のWhere条件
	Dim m_iNendo		'年度
	Dim m_sSikenKBN		'試験区分
	Dim m_iGakunen		'学年
	Dim m_sClassNo		'学科
	Dim m_sKamokuCd		'科目コード
	Dim m_sGakkaCd
	
	Dim m_FromSei
	Dim m_ToSei
	Dim m_FromKekka
	Dim m_ToKekka
	
	Dim m_bSeiInpFlg		'入力期間フラグ
	Dim m_bKekkaNyuryokuFlg	'欠課入力可能ﾌﾗｸﾞ(True:入力可 / False:入力不可)
	
	Dim m_UpdateDate
	
	'2002/06/21
	Dim m_iKamokuKbn
	Dim m_sKamokuBunrui
	Dim m_iSeisekiInpType
	
	Dim m_iDataCount
	Dim m_AryHyokaData()
	
	Dim m_iCount
	Dim m_sMiHyoka
	Dim m_Checked
	Dim m_Disabled
	Dim m_SchoolFlg
	Dim m_HyokaDispFlg
	Dim m_KekkaGaiDispFlg
	
	Dim m_TableWidth

	Dim m_sGakkoNO	'学校番号
	
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
	Dim w_sSQL
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="成績登録"
	w_sMsg=""
	w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = False
	
	Do
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If
		
		'//ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()
		
		'//不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'成績入力方法が文字入力のとき、科目評価データ取得
		if m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_STRING) then
			if not gf_GetKamokuHyokaData(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iDataCount,m_AryHyokaData) then 
				m_bErrFlg = True
				Exit Do
			end if
		end if
		
		'学校番号を取得
		if Not gf_GetGakkoNO(m_sGakkoNO) then
	        m_bErrFlg = True
		end if

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

'********************************************************************************
'*	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()
	
	m_iNendo	= request("txtNendo")
	m_sSikenKBN	= Cint(request("sltShikenKbn"))
	m_iGakunen	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sGakkaCd	= request("txtGakkaCd")
	
	m_bSeiInpFlg	= cbool(request("hidKikan"))
	m_bKekkaNyuryokuFlg	= request("hidKekkaNyuryokuFlg")
	
	m_iKamokuKbn	 	= request("hidKamokuKbn")
	m_sKamokuBunrui 	= request("hidKamokuBunrui")
	m_iSeisekiInpType 	= cint(request("hidSeisekiInpType"))
	
	m_UpdateDate = request("txtUpdDate")
	
	m_iCount = cint(request("i_Max"))
	m_sMiHyoka = request("hidMihyoka")
	m_SchoolFlg = cbool(request("hidSchoolFlg"))
	m_HyokaDispFlg = cbool(request("hidHyokaDispFlg"))
	m_KekkaGaiDispFlg = cbool(request("hidKekkaGaiDispFlg"))
	
	m_TableWidth = cint(request("hidTableWidth"))
	
	m_FromSei = gf_SetNull2String(request("hidFromSei"))
	m_ToSei = gf_SetNull2String(request("hidToSei"))
	m_FromKekka = gf_SetNull2String(request("hidFromKekka"))
	m_ToKekka = gf_SetNull2String(request("hidToKekka"))
	
	m_Checked  = ""
	m_Disabled = ""
	
End Sub

'********************************************************************************
'*	[機能]	試験名取得
'********************************************************************************
Function f_ShikenMei()
	Dim w_Rs
	
	On Error Resume Next
	Err.Clear
	
	f_ShikenMei = ""
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD = " & cint(m_sSikenKBN)
	w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
	w_sSQL = w_sSQL & " AND M01_NENDO = " & cint(m_iNendo)
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function
	
	If not w_Rs.EOF Then
		f_ShikenMei = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))
	End If
	
End Function

'********************************************************************************
'*  [機能]  履修テーブルより科目名称を取得
'********************************************************************************
Function f_GetKamokuName(p_Gakunen,p_GakkaCd,p_KamokuCd)
	Dim w_sSQL
	Dim w_Rs
	Dim w_GakkaCd
	
	On Error Resume Next
	Err.Clear
	
	f_GetKamokuName = ""
	
	w_sSQL = ""
	
	If m_iKamokuKbn = C_TUKU_FLG_TUJO Then '通常授業と特別活動で取り先を変える。
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	T15_KAMOKUMEI AS KAMOKUMEI"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T15_RISYU"
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 	T15_NYUNENDO=" & cint(m_iNendo) - cint(p_Gakunen) + 1
		w_sSQL = w_sSQL & " AND T15_GAKKA_CD='" & p_GakkaCd & "'"
		w_sSQL = w_sSQL & " AND T15_KAMOKU_CD='" & p_KamokuCd & "'"
	Else
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	M41_MEISYO AS KAMOKUMEI"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	M41_TOKUKATU"
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 	M41_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & " AND M41_TOKUKATU_CD='" & p_KamokuCd & "'"
	End If
	
	if gf_GetRecordset(w_Rs, w_sSQL) <> 0 then exit function
	
	If not w_Rs.EOF Then f_GetKamokuName = w_Rs("KAMOKUMEI")
	
	Call gf_closeObject(w_Rs)
	
End Function

'********************************************************************************
'*  [機能]  成績登録が文字の場合評価コンボを作成
'********************************************************************************
Sub s_SetHyokaCombo()
	Dim w_Str,w_lIdx
	
	w_Str = ""
	w_Str = w_Str & "<select name='sltHyoka'>"
	
	for w_lIdx = 0 to m_iDataCount-1
		
		w_Str = w_Str & "<option value='" & m_AryHyokaData(w_lIdx,0)
		'w_Str = w_Str & "#@#" & m_AryHyokaData(w_lIdx,1)
		w_Str = w_Str & "#@#" & m_AryHyokaData(w_lIdx,2)
		w_Str = w_Str & "'>" & m_AryHyokaData(w_lIdx,0)
		
	next
	
	w_Str = w_Str & "<option value=''>クリア"
	w_Str = w_Str & "</select>"
	
	response.write w_Str
	
End Sub

'********************************************************************************
'*  [機能]  未評価の設定
'********************************************************************************
Sub setHyokaType()
	
	'科目が未評価
	if cint(gf_SetNull2Zero(m_sMiHyoka)) = cint(C_MIHYOKA) then
		m_Checked = "checked"
	end if
	
	'入力期間外
	if not m_bSeiInpFlg then
		m_Disabled = "disabled"
	end if
	
End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
Sub showPage()
	Dim w_sInputClass
	Dim w_sColSpan

	'//NN対応
	If session("browser") = "IE" Then
		w_sInputClass = "class='num'"
	Else
		w_sInputClass = ""
	End If

	w_sColSpan = 11

	'試験区分が学年末の場合
	if m_sSikenKBN = C_SIKEN_KOU_KIM then
		w_sColSpan = 6
	end if

%>

<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--
	//************************************************************
    //  [機能]  ページロード時処理
    //************************************************************
    function window_onload(){
		//スクロール同期制御
		parent.init();
	}
	
	//************************************************************
    //  [機能]  登録ボタンが押されたとき
    //************************************************************
    function f_Touroku(){
        parent.main.f_Touroku();
    }
	
	//************************************************************
	//	[機能]	キャンセルボタンが押されたとき
	//************************************************************
	function f_Cancel(){
		//初期ページを表示
        parent.document.location.href="default.asp";
	}
	
	//************************************************************
	//	[機能]	ペーストボタンが押されたとき
	//************************************************************
	function f_Paste(pType){
		parent.main.document.frm.PasteType.value=pType;
		
		//submitで画面を開くとウィンドウのステータスが設定できないため､
		//一旦空ページを開いてから、新ウィンドウに対してsubmitする。
		nWin=open("","Paste","location=no,menubar=no,resizable=yes,scrollbars=no,scrolling=no,status=no,toolbar=no,width=300,height=600,top=0,left=0");
		parent.main.document.frm.target="Paste";
		parent.main.document.frm.action="sei0150_23_paste.asp";
		parent.main.document.frm.submit();
	}
	
	//************************************************************
	//	[機能]	未評価がチェックされたとき
	//************************************************************
	function setHyoka(){
		var w_num,w_type;
		var ob = new Array();
		
		if(document.frm.chkMiHyoka.checked){
			parent.main.document.frm.hidMihyoka.value=<%=C_MIHYOKA%>;
			w_type = true;
		}else{
			parent.main.document.frm.hidMihyoka.value="";
			w_type = false;
		}
		
		for(w_num=1;w_num<<%=m_iCount%>;w_num++){
			ob[0] = eval("parent.main.document.frm.chkHyokaFuno" + w_num);
			
			<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
				ob[1] = eval("parent.main.document.frm.Seiseki" + w_num);
			<% end if %>
			
			<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
				ob[2] = eval("parent.main.document.frm.hidSeiseki" + w_num);
			<% end if %>
			
			if(typeof(ob[0]) != "undefined" && ob[0].type == "checkbox"){
				if(w_type){
					ob[0].checked = false;
					<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
						ob[1].value = "";
					<% end if %>
					
					<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
						ob[2].value = "";
					<% end if %>
				}
				
				ob[0].disabled = w_type;
				
				<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
					ob[1].disabled = w_type;
				<% end if %>
			}
		}
		<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
			parent.main.f_GetTotalAvg();
		<% end if %>
	}

	//-->
	</SCRIPT>
	</head>
	<body onload="window_onload();">
	<table border="0" cellpadding="0" cellspacing="0" height="245" width="100%">
		<tr>
			<td>
				<%
				If m_bSeiInpFlg Then
					call gs_title(" 成績登録 "," 登　録 ")
				Else
					call gs_title(" 成績登録 "," 表　示 ")
				End If
				%>
			</td>
		</tr>
		<tr>
			<td align="center" nowrap>
			<form name="frm" method="post">
				<table border=1 class=hyo width=670>
					<tr>
						<th class="header3" colspan="6" nowrap align="center">
						成績入力期間　<%=f_ShikenMei()%>　　　更新日：<%=m_UpdateDate%>
						</th>
					</tr>
					<tr>
						<th class=header3 width="96"  align="center">成績入力期間</th><td class=detail width="239"  align="center" colspan="2"><%=m_FromSei%> 〜 <%=m_ToSei%></td>
						<th class=header3 width="96"  align="center">
							<%if m_sGakkoNO = C_NCT_KURUME then%>
								<font size=1>受講時間入力期間</font>
							<%else%>
								欠課入力期間
							<%end if%>
						</th><td class=detail width="239"  align="center" colspan="2"><%=m_FromKekka%> 〜 <%=m_ToKekka%></td>
					</tr>
					<tr>
						<th class=header3 width="96"  align="center">実施科目</th>
						<%
							w_str = m_iGakunen & "年　" & gf_GetClassName(m_iNendo,m_iGakunen,m_sClassNo) & "　" & f_GetKamokuName(m_iGakunen,m_sGakkaCd,m_sKamokuCd)
						%>
						<td class=detail colspan="5" align="center"><%=w_str%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<span class=msg2>
				<%if m_bSeiInpFlg then %>※ヘッダの文字色が「<FONT COLOR="#99CCFF">成績</FONT>」のようになっている部分をクリックすると、Excel貼り付け用の画面が開きます。<br><%end if%>
				</span><br>
				<% if m_bSeiInpFlg or m_bKekkaNyuryokuFlg Then %>
					<% if m_sSikenKBN = C_SIKEN_KOU_KIM and Not m_bSeiInpFlg then %>
					<% else %>
						<input type="button" class="button" value="　登　録　" onclick="f_Touroku();">　
					<% end if %>
				<% end if %>
				<input type="button" class="button" value="キャンセル" onclick="f_Cancel();">
				
			</td>
		</tr>
		<tr>
			<td align="center" valign="bottom" nowrap>
				<table class="hyo" border="1" align="center" width="<%=m_TableWidth%>">
					<tr>
						<th class="header3" colspan="<%= w_sColSpan %>" nowrap align="center">
							総授業数
							<%If m_bSeiInpFlg Then%>
								<input type="text" <%=w_sInputClass%> maxlength="3" style="width:30px" name="txtSouJyugyou" value="<%= Request("hidSouJyugyou") %>">
							<% Else %>
								&nbsp;<%= Request("hidSouJyugyou") %>
							<% End if%>
							&nbsp;&nbsp;純授業数
							<%If m_bSeiInpFlg Then%>
								<input type="text" <%=w_sInputClass%> maxlength="3" style="width:30px" name="txtJunJyugyou" value="<%= Request("hidJunJyugyou") %>">
							<% Else %>
								&nbsp;<%= Request("hidJunJyugyou") %>
							<% End if%>
							<%
							if m_bSeiInpFlg then
								'成績入力方法が文字入力のとき、評価コンボ表示
								if m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_STRING) then
									Call s_SetHyokaCombo()
								end if
							end if
							%>
						</th>
						<% if m_sSikenKBN = C_SIKEN_KOU_KIM then %>
							<th class="header3" colspan="5" nowrap align="center">
								総授業数計：&nbsp;<%= Request("hidSouJyugyou_KK") %>
								&nbsp;&nbsp;純授業数計：&nbsp;<%= Request("hidJunJyugyou_KK") %>
							</th>
						<% end if %>
					</tr>                                                                                                                                                 
					<tr>
						<th class="header3" rowspan="2" width="65" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
						<th class="header3" rowspan="2" width="150" nowrap>氏　名</th>
						<th class="header3" colspan="4" nowrap>成績履歴</th>
						
						<% if m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_NUM) and m_bSeiInpFlg then %>
							<th class="header3" rowspan="2" width="50" nowrap onClick="f_Paste('Seiseki')"><FONT COLOR="#99CCFF">成績</FONT></th>
						<% else %>
							<th class="header3" rowspan="2" width="50" nowrap>成績</th>
						<% end if %>

						<% If cstr(m_iKamokuKbn) = cstr(C_JIK_JUGYO) then %>
							<th class="header3" colspan="4" nowrap>欠課</th>
						<% end if %>
					</tr>
					
					<tr>
						<th class="header2" width="50" nowrap><span style="font-size:10px;">前中</span></th>
						<th class="header2" width="50" nowrap><span style="font-size:10px;">前末</span></th>
						<th class="header2" width="50" nowrap><span style="font-size:10px;">後中</span></th>
						<th class="header2" width="50" nowrap><span style="font-size:10px;">学末</span></th>

						<% If cstr(m_iKamokuKbn) = cstr(C_JIK_JUGYO) and m_bKekkaNyuryokuFlg and m_sSikenKBN <> C_SIKEN_KOU_KIM then %>
							<th class="header2" width="50" nowrap><span style="font-size:10px;" onClick="f_Paste('txtKekka')"><FONT COLOR="#99CCFF">欠課</FONT></span></th>
							<th class="header2" width="50" nowrap><span style="font-size:10px;" onClick="f_Paste('txtKibi')"><FONT COLOR="#99CCFF">忌引</FONT></span></th>
							<th class="header2" width="50" nowrap><span style="font-size:10px;" onClick="f_Paste('txtTeisi')"><FONT COLOR="#99CCFF">停止</FONT></span></th>
							<th class="header2" width="50" nowrap><span style="font-size:10px;" onClick="f_Paste('txtHaken')"><FONT COLOR="#99CCFF">派遣</FONT></span></th>
						<% Else %>
							<th class="header2" width="50" nowrap><span style="font-size:10px;">欠課</span></th>
							<th class="header2" width="50" nowrap><span style="font-size:10px;">忌引</span></th>
							<th class="header2" width="50" nowrap><span style="font-size:10px;">停止</span></th>
							<th class="header2" width="50" nowrap><span style="font-size:10px;">派遣</span></th>
						<% End If %>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	
	<input type="hidden" name="hidSeisekiInpType" value="<%=m_iSeisekiInpType%>">
	<input type="hidden" name="hidKekkaGaiDispFlg" value="<%=m_KekkaGaiDispFlg%>">
	<input type="hidden" name="hidKekkaNyuryokuFlg" value="<%=m_bKekkaNyuryokuFlg%>">
	
	</body>
	</html>
<%
End sub
%>