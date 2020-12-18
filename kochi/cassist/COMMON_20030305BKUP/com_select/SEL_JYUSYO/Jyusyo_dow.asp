<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 市町村検索画面
' ﾌﾟﾛｸﾞﾗﾑID : Common/com_select/SEL_JYUSYO/Jyusyo_dow.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:	
' 	           	hidSchMode	= 検索ﾌﾗｸﾞ
'   	        hidKenCd	= 県コード
'       	    hidShiCd	= 市町村コード
'           	txtCyouiki	= 町域コード
' 
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/30 持永
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

	Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	Public  m_bSchMode			'検索ﾌﾗｸﾞ
	Public  m_KenCd				'県コード
	Public  m_ShiCd				'市町村コード
	Public  m_Cyouiki 			'町域コード
	Public  m_CyouRs			'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ(町)

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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="連絡事項登録"
    w_sMsg=""
    w_sRetURL="../../login/top.asp"
    w_sTarget="_parent"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

		m_bSchMode = request("hidSchMode")	'検索ﾌﾗｸﾞ
		m_KenCd    = request("hidKenCd")	'県コード
		m_ShiCd    = request("hidShiCd")	'市町村コード
		m_Cyouiki  = request("txtCyouiki")	'町域コード

		'// 町データを取得
		if Not f_GetCyou() then Exit Do

        '// ページを表示
        Call showPage()

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_Rs)

    '// 終了処理
    Call gs_CloseDatabase()

End Sub

'******************************************************************
'機　　能：町データ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'******************************************************************
Function f_GetCyou()

	On Error Resume Next
	Err.Clear

	f_GetCyou = False

	
	if m_bSchMode then

		w_iWhereFlg = 0		'where句が前にあるFlg

		'// 県のWHERE句
		w_sKenSQL = ""
		if m_KenCd <> "" then
			w_sKenSQL = " M12_KEN_CD = '" & gf_fmtZero(m_KenCd,2) & "' "
			w_iWhereFlg = 1
		End if

		'// 市町村のWHERE句
		w_sShiSQL = ""
		if m_ShiCd <> "" then
			'// すでにWHERE文があったら"AND"でつなぐ
			if w_iWhereFlg = 1 then 
				w_sShiSQL = w_sShiSQL & " AND "
			End if
			w_sShiSQL = w_sShiSQL & " M12_SITYOSON_CD = '" & m_ShiCd & "' "
			w_iWhereFlg = 1
		End if

		'// 町域指定のWHERE句
		w_sCyouikiSQL = ""
		if m_Cyouiki <> "" then
			'// すでにWHERE文があったら"AND"でつなぐ
			if w_iWhereFlg = 1 then 
				w_sShiSQL = w_sShiSQL & " AND "
			End if
			w_sCyouikiSQL = w_sCyouikiSQL & " M12_TYOIKIMEI like '%" & m_Cyouiki & "%' "
		End if

		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	M12_YUBIN_BANGO, "
		w_sSQL = w_sSQL & " 	M12_SITYOSONMEI, "
		w_sSQL = w_sSQL & " 	M12_RENBAN, "
		w_sSQL = w_sSQL & " 	M12_TYOIKIMEI "
		w_sSQL = w_sSQL & " FROM  "
		w_sSQL = w_sSQL & " 	M12_SITYOSON "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & w_sKenSQL
		w_sSQL = w_sSQL & w_sShiSQL
		w_sSQL = w_sSQL & w_sCyouikiSQL

		iRet = gf_GetRecordset(m_CyouRs, w_sSQL)
		If iRet <> 0 Then
			m_bErrFlg = True
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Function
		End If

	End if

	f_GetCyou = True

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
<link rel=stylesheet href="../../style.css" type=text/css>
<!--#include file="../../jsCommon.htm"-->
<script language="JavaScript">
<!--
    //************************************************************
    //  [機能]  オープナーにデータを返す
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
	function jf_ReturnDate(pZipCode,pJusyoNum,pRenban) {

		parent.opener.document.frm.txtYUBINBANGO.value = pZipCode;		// 郵便番号
		pJusyo1 = eval("document.frm.hidSITYOSONMEI_"+pJusyoNum+".value");
		pJusyo2 = eval("document.frm.hidTYOIKIMEI_"+pJusyoNum+".value");
		parent.opener.document.frm.txtJUSYO1.value  = pJusyo1;			// 住所１
		parent.opener.document.frm.txtJUSYO2.value  = pJusyo2;			// 住所２
		parent.opener.document.frm.txtJUSYO3.value  = "";				// 住所３

		parent.opener.document.frm.txtRenban.value  = pRenban;			// 連番(キー)
		parent.opener.document.frm.txtKenCd.value   = document.frm.hidKenCd.value;	// 県コード(キー)
		parent.opener.document.frm.txtSityoCd.value = document.frm.hidShiCd.value;	// 市町村コード(キー)

		parent.window.close();

	}
//-->
</script>
</head>

<body>
<form name="frm" method="post">
<div align="center">

<% if m_bSchMode then %>
<table>
	<tr>
		<td align="center">

			<table border="1" class="hyo">
				<tr>
					<th align="center" class="header" width="80">郵便番号</th>
					<th align="center" class="header" width="266">町域名</th>
				</tr>
				<% if m_CyouRs.Eof then %>
					<tr>
						<td colspan="2" class="CELL1">該当するデータがありません。</td>
					</tr>
				<% End if %>
				<% 
					i = 1		'連番
					Do Until m_CyouRs.Eof
					
					%>
					<tr>
						<td class="CELL1" align="center"><a href="javascript:jf_ReturnDate('<%= m_CyouRs("M12_YUBIN_BANGO") %>','<%=i%>','<%= m_CyouRs("M12_RENBAN") %>');"><%= m_CyouRs("M12_YUBIN_BANGO") %></a>
							<input type="hidden" name="hidSITYOSONMEI_<%=i%>" value="<%= m_CyouRs("M12_SITYOSONMEI") %>">
							<input type="hidden" name="hidTYOIKIMEI_<%=i%>"   value="<%= m_CyouRs("M12_TYOIKIMEI") %>"></td>
						<td class="CELL1" nowrap><%= m_CyouRs("M12_SITYOSONMEI") & m_CyouRs("M12_TYOIKIMEI") %></td>
					</tr>
					<%
						i = i + 1
						m_CyouRs.MoveNext
					Loop
				%>
			</table>

		</td>
	</tr>
	<tr>
		<td align="right">
			<input type="button" class="button" value="キャンセル" onClick="parent.window.close();">
		</td>
	</tr>
</table>
<% End if %>

<input type="hidden" name="hidKenCd" value="<%= m_KenCd %>">
<input type="hidden" name="hidShiCd" value="<%= m_ShiCd %>">
</div>
</form>
</body>
</html>
<%
End Sub
%>