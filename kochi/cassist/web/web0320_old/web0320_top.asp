<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 使用教科書登録
' ﾌﾟﾛｸﾞﾗﾑID : web/web0320/web0320_top.asp
' 機	  能: 上ページ 使用教科書登録の検索を行う
'-------------------------------------------------------------------------
' 引	  数:教官コード 	＞		SESSIONより（保留）
'			:年度			＞		SESSIONより（保留）
' 変	  数:なし
' 引	  渡:教官コード 	＞		SESSIONより（保留）
'			:年度			＞		SESSIONより（保留）
' 説	  明:
'			■初期表示
'				コンボボックスは空白で表示
'			■表示ボタンクリック時
'				下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作	  成: 2001/08/01 前田 智史
' 変	  更: 2001/08/07 根本 直美	   NN対応に伴うソース変更
' 変	  更: 2001/08/18 伊藤　公子 次年度の学期情報がない時は次年度の入力が出来ないようにする
' 変	  更: 2001/08/22 伊藤　公子 教官を選択できるように変更
' 変	  更: 2001/12/01 田部 雅幸 所属学科のみを変更するように修正
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public	m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ

	'市町村選択用のWhere条件
	Public m_iNendo 		'年度
	Public m_sKyokanCd		'教官コード
	Public m_sKyokanSimei	'教官氏名
	Public m_bJinendoGakki	'次年度学期情報

	Public m_iGakunen
	Public m_sGakkaCd
	Public m_sGakunenWhere
	Public m_sGakkaWhere

	Public	m_Rs
	Public	m_iMax			'最大ページ
	Public	m_iDsp			'一覧表示行数

	Public m_sSyozokuGakka		'//2001/12/01 Add ログインした教官の所属する学科

'///////////////////////////メイン処理/////////////////////////////

	'ﾒｲﾝﾙｰﾁﾝ実行
	Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

Sub Main()
'********************************************************************************
'*	[機能]	本ASPのﾒｲﾝﾙｰﾁﾝ
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	Dim w_iRet				'// 戻り値
	Dim w_sSQL				'// SQL文
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="使用教科書登録"
	w_sMsg=""
	w_sRetURL="../../login/default.asp" 	
	w_sTarget="_top"


	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

'	 m_iNendo	 = session("NENDO")

'	If Request("SKyokanCd1") <> "" Then
'		m_sKyokanCd = Request("SKyokanCd1")
'	Else
'	 	m_sKyokanCd = session("KYOKAN_CD")
'	End If

	m_iDsp = C_PAGE_LINE

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

		Call s_SetParam()

		'//次年度情報があるかチェック
		w_iRet = f_GetJinendoGakki(m_bJinendoGakki)
		If w_iRet  = False Then
			m_bErrFlg = True
			exit do
		End If

'//デバッグ
'Call s_DebugPrint

'		'//教官氏名を取得
'		 w_iRet = f_KyokanSimei()
'		 If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//学年のコンボを取得
		Call s_MakeGakunenWhere()

		'//学科のコンボを取得
		Call s_MakeGakkaWhere()

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
	Call gf_closeObject(m_Rs)
	Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_SetParam()

	If Request("txtNendo") = "" Then
		m_iNendo   = session("NENDO")

		'//次年度情報がある場合は、次年度の教科書登録を初期設定の対象とする
		If m_bJinendoGakki = True Then
			m_iNendo   = m_iNendo + 1
		End If
	Else
		m_iNendo = Request("txtNendo")
	End If

	m_iGakunen = Request("txtGakunenCd")
	m_sGakkaCd = Request("txtGakkaCD")

End Sub

'********************************************************************************
'*	[機能]	デバッグ用
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iNendo   = " & m_iNendo	& "<br>"
	response.write "m_iGakunen = " & m_iGakunen & "<br>"
	response.write "m_sGakkaCd = " & m_sGakkaCd & "<br>"

End Sub

'********************************************************************************
'*	[機能]	次年度の学期情報があるかどうかチェックする
'*	[引数]	なし
'*	[戻値]	p_bJinendoGakki=true:学期情報あり
'*			p_bJinendoGakki=false:学期情報なし
'*	[説明]	
'********************************************************************************
Function f_GetJinendoGakki(p_bJinendoGakki)
	Dim w_iRet				'// 戻り値
	Dim w_sSQL				'// SQL文
	dim w_Rs

	on error resume next
	err.clear

	f_GetJinendoGakki = False
	p_bJinendoGakki = False

	'//次年度の学期情報があるかどうか
	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI"
	w_sSQL = w_sSQL & vbCrLf & " FROM M01_KUBUN"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_NENDO=" & cint(SESSION("NENDO"))+1
	w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & C_KAISETUKI

	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		m_bErrFlg = True
		Exit Function
	End If

	'//データがあった時
	If w_Rs.EOF = False Then
		p_bJinendoGakki = True
	End If

	Call gf_closeObject(w_Rs)

	f_GetJinendoGakki = True

End Function

'********************************************************************************
'*	[機能]	学年コンボに関するWHEREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_MakeGakunenWhere()

	m_sGakunenWhere = ""
	m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iNendo
	m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

End Sub

'********************************************************************************
'*	[機能]	学科コンボに関するWHREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_MakeGakkaWhere()
	m_sGakkaWhere=""
	m_sGakkaWhere = " M02_NENDO = " & m_iNendo
	m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD <> '00' "

End Sub

'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'		(リストダウンボックス選択表示用)
'[引数] pData1 : データ１
'		pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'					
'****************************************************
Function f_Selected(pData1,pData2)

	f_Selected = ""

	If IsNull(pData1) = False And IsNull(pData2) = False Then
		If trim(cStr(pData1)) = trim(cstr(pData2)) Then
			f_Selected = "selected" 
		Else
		End If
	End If

End Function

Sub showPage()
'********************************************************************************
'*	[機能]	HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
%>
<html>
<head>
<title>使用教科書登録</title>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//	[機能]	表示ボタンクリック時の処理
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Search(){

		// ■■■NULLﾁｪｯｸ■■■
		// ■年度
		if( f_Trim(document.frm.txtNendo.value) == "" ){
			window.alert("年度の選択を行ってください");
			document.frm.txtNendo.focus();
			return ;
		}

		document.frm.action="web0320_main.asp";
		document.frm.target="main";
		document.frm.submit();
	
	}
	//************************************************************
	//	[機能]	登録ボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Touroku(){

		document.frm.action="./touroku.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.txtMode.value = "Touroku";
		document.frm.submit();
	
	}

	//************************************************************
	//	[機能]	教官参照選択画面ウィンドウオープン
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function KyokanWin(p_iInt,p_sKNm) {
		var obj=eval("document.frm."+p_sKNm)

		URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
		nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=600,top=0,left=0");
		nWin.focus();
		return true;	
	}

	//************************************************************
	//	[機能]	クリアボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function fj_Clear(){

		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}

	//************************************************************
	//	[機能]	年度が変更されたとき、本画面を再表示
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_ReLoadMyPage(){

		document.frm.action="./web0320_top.asp";
		document.frm.target="_self";
		document.frm.submit();

	}

	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>

	<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<center>
	<form name="frm" method="POST">
	<% call gs_title("使用教科書登録","一　覧") %>
	<br>
	<table border="0">
	<tr>
	<td valign="bottom">
		<table border="0" cellpadding="1" cellspacing="1">
		<tr>
		<td align="left" class="search">
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
			<td Nowrap>
				<select name="txtNendo" onchange = 'javascript:f_ReLoadMyPage()'>
					<%If m_bJinendoGakki = True Then%>
						<%w_iNen=Session("NENDO")%>
						<option VALUE="<%= w_iNen + 1 %>" <%=f_Selected(Request("txtNendo"),w_iNen + 1)%> ><%= w_iNen + 1 %>
						<option VALUE="<%= w_iNen %>"	  <%=f_Selected(Request("txtNendo"),w_iNen)%> ><%= w_iNen %>
					<%Else%>
						<option VALUE="<%= m_iNendo %>" 			><%= m_iNendo %>
					<%End If%>

				</select>
			</td>
			<td>年度&nbsp;&nbsp;</td>

			<td>学年</td>
			<td nowrap align="left">
				<% call gf_ComboSet("txtGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere," style='width:40px;' ",True,m_iGakunen) %>
			</td>

			<td nowrap>学科</td>
			<td nowrap align="left">

			<%	'共通関数から学科に関するコンボボックスを出力する
				call gf_ComboSet_Gakka("txtGakkaCD",C_CBO_M02_GAKKA,m_sGakkaWhere,"style='width:175px;' ",True,m_sGakkaCd)%>
			</td>

			</tr><tr>
			<td Nowrap align="right" colspan=6>
			<input class="button" type="button" value="　表　示　" onClick = "javascript:f_Search();">
			</td>
			</tr>
			</table>
		</td>
		</tr>
		</table>
	</td>
	<td valign="top">
	<a href="javascript:f_Touroku();">新規登録はこちら</a><br><img src="../../image/sp.gif" height="10"><br>
	</td>
	</tr>
	</table>
	<input type="hidden" name="txtMode" value="">
	</form>
	</center>
	</body>
	</html>

<%
	'---------- HTML END   ----------
End Sub
%>
