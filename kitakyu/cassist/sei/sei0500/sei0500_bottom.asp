<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 実力試験成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0500/sei0500_bottom.asp
' 機      能: 下ページ 実力試験の成績を入力する
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 説      明:

'-------------------------------------------------------------------------
' 作      成: 2001/09/06 モチナガ
' 変      更: 2016/05/18 Nishimura 異動(休学者)の場合更新できない障害対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

	'氏名選択用のWhere条件
    Public m_iNendo			'年度
    Public m_sKyokanCd		'教官コード
    Public m_sGakuNo		'学年
    Public m_sClassNo		'学科
    Public m_sKamokuCd		'科目コード
    Public m_sSiKenCd		'試験コード

    Public m_SeitoRs		'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ(生徒)
    Public m_rCnt			'ﾚｺｰﾄﾞｶﾝｳﾄ(生徒)
    Public m_lKikan			'0：成績入力期間内、1：成績入力期間外

	Public	m_iMax			'最大ページ
	Public  m_Half			'最大ページの半分
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
    Dim w_sSQL              '// SQL文
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="実力試験成績登録"
	w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_lKikan = 0
	
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

	    '// ﾊﾟﾗﾒｰﾀSET
	    Call s_SetParam()

		'//期間データの取得
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			m_lKikan = 1
			'// ページを表示
			'Call No_showPage()
			'Exit Do
		End If
		
		w_iRet = 0
		
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If

		'//クラス別生徒データ取得
        w_iRet = f_GetClassData()
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
		If m_SeitoRs.EOF Then
			Call ShowPage_No()
			Exit Do
		End If

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
    Call gf_closeObject(m_SeitoRs)
    Call gs_CloseDatabase()

End Sub

Sub s_SetParam()
'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSiKenCd	= Cint(request("txtShikenCd"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")

End Sub

Function f_GetClassData()
'********************************************************************************
'*	[機能]	生徒を取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
	Dim w_iNyuNendo

	On Error Resume Next
	Err.Clear
	f_GetClassData = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_TOKUTEN,  "
		w_sSQL = w_sSQL & vbCrLf & " 	T11.T11_SIMEI,  "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_GAKUSEKI_NO "
		w_sSQL = w_sSQL & vbCrLf & " FROM  "
		w_sSQL = w_sSQL & vbCrLf & " 	T11_GAKUSEKI T11, "
		w_sSQL = w_sSQL & vbCrLf & " 	T13_GAKU_NEN T13, "
		w_sSQL = w_sSQL & vbCrLf & " 	T33_SIKEN_SEISEKI T33  "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "	T33.T33_GAKUSEKI_NO  = T13.T13_GAKUSEKI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "	T13.T13_GAKUSEI_NO   = T11.T11_GAKUSEI_NO  AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_NENDO        = T13.T13_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_NENDO        =  " & m_iNendo          & " AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_SIKEN_CD     =  " & m_sSiKenCd        & " AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_SIKEN_KAMOKU = '" & m_sKamokuCd       & "' AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_GAKUNEN      =  " & m_sGakuNo         &" AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_CLASS        =  " & m_sClassNo
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY T33_GAKUSEKI_NO "

        iRet = gf_GetRecordset(m_SeitoRs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			m_bErrFlg = True
            msMsg = Err.description
            f_GetClassData = 99
            Exit Do
        End If

		'//ﾚｺｰﾄﾞカウント取得
		m_rCnt = gf_GetRsCount(m_SeitoRs)

		f_GetClassData = 0
		Exit Do
	Loop

End Function

Function f_Nyuryokudate()
'********************************************************************************
'*	[機能]	入力期間取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SEISEKI_KAISI, "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SEISEKI_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & "  FROM  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND "	'試験区分(実力試験のみ)
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO        =  " & m_iNendo	& "  AND "		'処理年度
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_CD     =  " & m_sSiKenCd  & "  AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KAMOKU = '" & m_sKamokuCd & "' AND "
		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_SEISEKI_KAISI  <= '" & gf_YYYY_MM_DD(date(),"/") & "' AND"
		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If m_DRs.EOF Then
			Exit Do
		End If

	    Call gf_closeObject(m_DRs)

		f_Nyuryokudate = 0
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
	On Error Resume Next
	Err.Clear

	'// 生徒数の半分
	m_Half = gf_Round(m_rCnt / 2, 0)

%>
	<html>
	<head>
	<link rel=stylesheet href="../../common/style.css" type=text/css>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT language="javascript">
	<!--
	    //************************************************************
	    //  [機能]  ページロード時処理
	    //  [引数]
	    //  [戻値]
	    //  [説明]
	    //************************************************************
	    function window_onload(){

			//スクロール同期制御
			parent.init();

			//成績合計値の取得
			f_GetTotalAvg();

	        //submit
	        document.frm.target = "topFrame";
	        document.frm.action = "sei0500_middle.asp?<%=Request.Form.Item%>"
	        document.frm.submit();

	    }
		//************************************************************
		//	[機能]	キャンセルボタンが押されたとき
		//	[引数]	なし
		//	[戻値]	なし
		//	[説明]
		//************************************************************
		function f_Cansel(){

	        //初期ページを表示
	        parent.document.location.href="default.asp"
		
		}
	    //************************************************************
	    //  [機能]  登録ボタンが押されたとき
	    //  [引数]  なし
	    //  [戻値]  なし
	    //  [説明]
	    //************************************************************
	    function f_Touroku(){

			// 数字ﾁｪｯｸ
			for(i=1; i < <%= m_rCnt %>; i++){
//Ins_s 2016/05/18 Nishimura
//異動の場合はチェックしない
				objIdo = eval("document.frm.hidIdoCnt"+i)

				if (objIdo.value == 1){
//Ins_e 2016/05/18 Nishimura
					obj = eval("document.frm.Seiseki"+i)
					if (obj.value.match(/[^0-9]/) ){
			            alert("入力値が不正です");
						obj.focus();
						return;
					}
				}
			}

			//登録処理
	        document.frm.action = "sei0500_upd.asp?<%=Request.Form.Item%>";
	        document.frm.target = "main";
	        document.frm.submit();

	    }

		//************************************************
		//  [機能]  Enter キーで下の入力フォームに動くようになる
		//  [引数]  p_inpNm	対象入力フォーム名
		//          p_frm	対象フォーム
		//          i		現在の番号
		//  [説明]  
		//************************************************
		function f_MoveCur(p_inpNm,p_frm,i){
			if (event.keyCode == 13){		//押されたキーがEnter(13)の時に動く。
				i++;
				if (i > <%=m_rCnt%>) i = 1; //iが最大値を超えると、はじめに戻る。
				inpForm = eval("p_frm."+p_inpNm+i);
				inpForm.focus();			//フォーカスを移す。
				inpForm.select();			//移ったテキストボックス内を選択状態にする。
			}else{
				return false;
			}
			return true;
		}

	//-->
	</SCRIPT>
	</head>

    <body onLoad="window_onload();">
	<form name="frm" method="post" onClick="return false;">
	<center>

	<table >
		<tr>
			<td valign="top">

				<table class="hyo" border="1" align="center" width="280">
				<%

					Dim w_IdouCnt
					Dim w_sIdouMei

					i = 1
					Do until m_SeitoRs.Eof or i > m_Half

						'**異動処理　Add 2001.12.22 oakda*********************
						w_IdouCnt = gf_Set_Idou(Cstr(m_SeitoRs("T33_GAKUSEKI_NO")),m_iNendo,w_sIdouMei)

						if w_sIdouMei <> "" then
							w_sIdouMei = "[" & w_sIdouMei & "]"
						End if
						'*****************************************************

						Call gs_cellPtn(w_cell)
						%>
							<tr>
								<td class="<%=w_cell%>" width="50" align="center"><%=m_SeitoRs("T33_GAKUSEKI_NO")%></td>
								<td class="<%=w_cell%>" width="200"><%=m_SeitoRs("T11_SIMEI")%><%=w_sIdouMei%></td>
								<input type="hidden" align="center" name="hidIdoCnt<%=i%>" value="<%= w_IdouCnt %>">

<% IF w_IdouCnt = 1 Then %>
							<%If m_lKikan = 1 Then%>
								<td class="<%=w_cell%>" width="30"align="right">
								<input type="hidden" class='num' align="center"  name="Seiseki<%=i%>" value="<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)">
								<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>
								</td>
							<% Else %>
								<td class="<%=w_cell%>" width="30"><input type="text" class='num' align="center"  name="Seiseki<%=i%>" value="<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)"></td>
							<% End IF %>
								<input type="hidden" align="center" name="hidGakusekiNo<%=i%>" value="<%= m_SeitoRs("T33_GAKUSEKI_NO") %>">
<% Else %>
								<td class="<%=w_cell%>" align="center" width="30">-</td>
<% End IF %>
							</tr>
						<%
						i = i + 1
						m_SeitoRs.MoveNext
					Loop
				%>
				</table>

			</td>
			<td valign="top">

				<table class="hyo" border="1" align="center" width="280">
				<%
					Do until m_SeitoRs.Eof
						
						'**異動処理　Add 2001.12.22 oakda*********************
						w_IdouCnt = gf_Set_Idou(Cstr(m_SeitoRs("T33_GAKUSEKI_NO")),m_iNendo,w_sIdouMei)

						if w_sIdouMei <> "" then
							w_sIdouMei = "[" & w_sIdouMei & "]"
						End if
						'*****************************************************

						Call gs_cellPtn(w_cell)

						%>
							<tr>
								<td class="<%=w_cell%>" width="50" align="center"><%=m_SeitoRs("T33_GAKUSEKI_NO")%></td>
								<td class="<%=w_cell%>" width="200"><%=m_SeitoRs("T11_SIMEI")%><%=w_sIdouMei%></td>
								<input type="hidden" align="center" name="hidIdoCnt<%=i%>" value="<%= w_IdouCnt %>">

<% IF w_IdouCnt = 1 Then %>
							<%If m_lKikan = 1 Then%>
								<td class="<%=w_cell%>" width="30" align="right">
								<input type="hidden" class='num' align="center"  name="Seiseki<%=i%>" value="<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)">
								<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>
								</td>
							<% Else %>
								<td class="<%=w_cell%>" width="30"><input type="text" class='num' align="center"  name="Seiseki<%=i%>" value="<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)"></td>
							<% End IF %>

								<input type="hidden" align="center" name="hidGakusekiNo<%=i%>" value="<%= m_SeitoRs("T33_GAKUSEKI_NO") %>">
<% Else %>
								<td class="<%=w_cell%>" align="center" width="30">-</td>
<% End IF %>
							</tr>
						<%
						i = i + 1
						m_SeitoRs.MoveNext
					Loop%>




				</table>

			</td>
		</tr>
		<tr>
			<td colspan=2>
				<table class="hyo" border="1" align="center" width="100%">
					<tr>
						<td class="header" nowrap align="right">
							<FONT COLOR="#FFFFFF"><B>成績合計</B></FONT>
							<input type="text" name="txtTotal" size="5" <%=w_sInputClass%> readonly>
						</td>
					</tr>
					<tr>
						<td class="header" nowrap align="right">
							<FONT COLOR="#FFFFFF"><B>平均点</B></FONT>
							<input type="text" name="txtAvg" size="5" <%=w_sInputClass%> readonly>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<table width="50%">
		<tr>
			<td align="center" nowrap>
				<input type="button" class="button" value="　登　録　" onclick="javascript:f_Touroku()">　
				<input type="button" class="button" value="キャンセル" onclick="javascript:f_Cansel()">
			</td>
		</tr>
	</table>

	<input type="hidden" name="hidRecCnt" value="<%= m_rCnt %>">
	<input type="hidden" name="i_Max"       value="<%=i%>">
	<input type="hidden" name="PasteType" value="">
	<input type="hidden" name="i_Maxherf" value="<%=m_Half%>">
	</FORM>
	</center>
	</body>
	<SCRIPT>
		//************************************************************
		//	[機能]	成績が変更されたとき
		//	[引数]	なし
		//	[戻値]	なし
		//	[説明]	成績の合計と平均を求める
		//	[備考]	学生の総数が分かるのは最後であるため、この位置に書く。
		//************************************************************
		function f_GetTotalAvg(){
			var i;
			var total;
			var avg;
			var cnt;

			total = 0;
			cnt = 0;
			avg = 0;

	<%If m_iKikan <> "NO" Then	'入力期間中%>

			//学生数でのループ
			for(i=0;i<<%=i%>;i++) {

				//存在するかどうか
				textbox = eval("document.frm.Seiseki" + (i+1));
				if (textbox) {
					//未入力チェック
					if (textbox.value != "") {
						//数字でないのは無視する
						if (!isNaN(textbox.value)) {
							total = total + parseInt(textbox.value);
						}
					}
					cnt = cnt + 1;
				}
			}

	<% Else	'入力期間中ではない%>
		total = <%=w_lSeiTotal%>;
		cnt   = <%=w_lGakTotal%>;
	<% End If%>

			document.frm.txtTotal.value=total;

			//四捨五入
			if (cnt!=0){
				avg = total/cnt;
				avg = avg * 10;
				avg = Math.round(avg);
				avg = avg / 10;
			}
			
			document.frm.txtAvg.value=avg;
		}
	</SCRIPT>

	</html>
<%
End sub

Sub No_showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

    <body>
	<form name="frm" method="post">
	<center>
	<br><br><br>
		<span class="msg">成績入力期間外です。</span>
	</center>

	<input type="hidden" name="txtMsg" value="成績入力期間外です。">

	</form>
	</body>
	</html>

<%
End Sub
Sub showPage_No()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

    <body>
	<form name="frm" method="post">
	</head>

	<body>
	<br><br><br>
	<center>
		<span class="msg">データが存在しません。</span>
	</center>

	<input type="hidden" name="txtMsg" value="データが存在しません。">

	</form>
	</body>
	</html>

<%
End Sub
%>