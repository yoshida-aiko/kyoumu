<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 市町村検索画面
' ﾌﾟﾛｸﾞﾗﾑID : Common/com_select/SEL_JYUSYO/Jyusyo_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:	hidKenCd = 県コード（自分自身へサブミット）
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

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_KenRs				'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ（県マスタ）
    Public  m_KenCd				'県コード
    Public  m_ShiCd				'市コード
    Public  m_ShiRs				'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ（市町村)
    Public  m_Cyouiki 			'町域コード
    Public  m_JUSYO1			'県市区(検索条件)
    Public  m_JUSYO2            '町    (検索条件)
    Public  m_NoHitFlg			'検索ﾋｯﾄなしﾌﾗｸﾞ

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
	m_NoHitFlg = 0

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

		'//-- ユーザー入力条件 --
		m_JUSYO1 = request("JUSYO1")		'県市区
		m_JUSYO2 = request("JUSYO2")		'町

'		m_JUSYO1 = Session("m_JUSYO1")
'		m_JUSYO2 = Session("m_JUSYO2")

		'//----------------------//
		m_KenCd   = request("hidKenCd")		'県コード
		m_Cyouiki = request("txtCyouiki")	'町域コード

		'// ユーザー入力条件から県コードを取得
		if Not f_GetKenCode() then Exit Do

		'// 県データを取得
		if Not f_GetKenMaster() then Exit Do

		'// ユーザー入力条件から市町村コードを取得
		if Not f_GetShicyosonCode() then Exit Do

		'// 市町村を取得
		if Not f_GetShicyoson() then Exit Do
        
        '// ページを表示
        Call showPage()

		'// ｾｯｼｮﾝ削除
'		Session("m_JUSYO1") = ""
'		Session("m_JUSYO2") = ""

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_KenRs)
    Call gf_closeObject(m_ShiRs)

    '// 終了処理
    Call gs_CloseDatabase()

End Sub

'******************************************************************
'機　　能：ユーザー入力条件から県コードを取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'******************************************************************
Function f_GetKenCode()

	On Error Resume Next
	Err.Clear

	f_GetKenCode = False

	'// 住所１がNullじゃなかったら
	if m_JUSYO1 <> "" then

		'// 県CDを取得する
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	M16_KEN_CD "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	M16_KEN "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 		M16_NENDO  =  " & Session("NENDO")
		w_sSQL = w_sSQL & " 	AND M16_KENMEI = '" & m_JUSYO1 & "'"

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			m_bErrFlg = True
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Function
		End If

		'// Nullじゃなかったら変数に入れる
		if Not w_Rs.Eof then
			m_KenCd = w_Rs("M16_KEN_CD")	
		Else
			m_NoHitFlg = 1						'// ﾋｯﾄなしﾌﾗｸﾞ
		End if

	End if

    Call gf_closeObject(w_Rs)
	f_GetKenCode = True

End Function

'******************************************************************
'機　　能：県データ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'******************************************************************
Function f_GetKenMaster()

	On Error Resume Next
	Err.Clear

	f_GetKenMaster = False

	'// 県マスタを取得する
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M16_KEN_CD,  "
	w_sSQL = w_sSQL & " 	M16_KENMEI "
	w_sSQL = w_sSQL & " FROM  "
	w_sSQL = w_sSQL & " 	M16_KEN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M16_NENDO = " & Session("NENDO")

	iRet = gf_GetRecordset(m_KenRs, w_sSQL)
	If iRet <> 0 Then
		m_bErrFlg = True
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit Function
	End If

	if m_KenRs.Eof then
		Call gs_SetErrMsg("都道府県を取得時に、エラーが発生しました")
		m_bErrFlg = True
		Exit Function
	End if

	f_GetKenMaster = True

End Function

'******************************************************************
'機　　能：ユーザー入力条件から市町村コードを取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'******************************************************************
Function f_GetShicyosonCode()

	On Error Resume Next
	Err.Clear

	f_GetShicyosonCode = False

	if m_NoHitFlg = 1 then

		'// 県コードがわかっていたらWHERE条件に加える
		if m_KenCd <> "" then
			w_sKenSQL = " AND M12_KEN_CD = '" & m_KenCd & "'"
		End if
		
		'// 市町村CDを取得する
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	M12_KEN_CD,  "
		w_sSQL = w_sSQL & " 	M12_SITYOSON_CD,  "
		w_sSQL = w_sSQL & " 	M12_SITYOSONMEI "
		w_sSQL = w_sSQL & " FROM  "
		w_sSQL = w_sSQL & " 	M12_SITYOSON "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 	M12_SITYOSONMEI Like '%" & m_JUSYO1 & "%' " 
		w_sSQL = w_sSQL & 		w_sKenSQL
		w_sSQL = w_sSQL & " Group by "
		w_sSQL = w_sSQL & " 	M12_KEN_CD,  "
		w_sSQL = w_sSQL & " 	M12_SITYOSON_CD,  "
		w_sSQL = w_sSQL & " 	M12_SITYOSONMEI "

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			m_bErrFlg = True
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Function
		End If

		'// Nullじゃなかったら変数に入れる
		if Not w_Rs.Eof then
			m_KenCd = w_Rs("M12_KEN_CD")
			m_ShiCd = w_Rs("M12_SITYOSON_CD")
			m_NoHitFlg = 0
		End if

	End if

    Call gf_closeObject(w_Rs)
	f_GetShicyosonCode = True

End Function

'******************************************************************
'機　　能：市町村データ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'******************************************************************
Function f_GetShicyoson()

	On Error Resume Next
	Err.Clear

	f_GetShicyoson = False

	'// 県コードがNUllじゃなかったら
	if m_KenCd <> "" then

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_KEN_CD,  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSON_CD,  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSONMEI "
		w_sSQL = w_sSQL & vbCrLf & " FROM  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSON "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_KEN_CD = '" & gf_fmtZero(m_KenCd,2) & "' "
		w_sSQL = w_sSQL & vbCrLf & " Group by "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_KEN_CD,  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSON_CD,  "
		w_sSQL = w_sSQL & vbCrLf & " 	M12_SITYOSONMEI "

		iRet = gf_GetRecordset(m_ShiRs, w_sSQL)

		If iRet <> 0 Then
			m_bErrFlg = True
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Function
		End If

	End if

	f_GetShicyoson = True

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

	'// 住所２の値を町域にいれる
	if m_JUSYO2 <> "" then
		m_Cyouiki = m_JUSYO2
	End if

	'// オンロード時にサブミットする
	if m_NoHitFlg = 1 AND m_JUSYO2 = "" then
		w_jFName = ""
	Else
		if m_JUSYO1 & m_JUSYO2 <> "" then
			w_jFName = " onLoad='jf_OnLoadSubmit();'"
		End if
	End if

	'// ネスケとIEによってtextboxのサイズを変える
	if session("browser") = "IE" then
		w_FormSize = "61"
	Else
		w_FormSize = "44"
	End if

%>
<html>

<head>
<link rel=stylesheet href="../../style.css" type=text/css>
<!--#include file="../../jsCommon.htm"-->
<script language="JavaScript">
<!--


    //************************************************************
    //  [機能]  ロード時にサブミットする
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
	function jf_OnLoadSubmit(){
		document.frm.action         = "Jyusyo_dow.asp";
		document.frm.target         = "dow";
		document.frm.submit();
	}


    //************************************************************
    //  [機能]  県が選ばれたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
	function jf_KenSelect(){
		w_KenCd = document.frm.selKen.value;
		document.frm.hidKenCd.value = w_KenCd;
		document.frm.action         = "Jyusyo_top.asp";
		document.frm.target         = "top";
		document.frm.submit();
	}

    //************************************************************
    //  [機能]  検索ボタンが選ばれたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
	function jf_ShiSelect(){
		if( document.frm.selKen.value == "" && document.frm.selShi.value == "" ){
			window.alert("県名指定をしてください");
			document.frm.selKen.focus();
	        return false;
		}

		w_ShiCd = document.frm.selShi.value;
		document.frm.hidShiCd.value = w_ShiCd;
		document.frm.action         = "Jyusyo_dow.asp";
		document.frm.target         = "dow";
		document.frm.submit();
	}

    //************************************************************
    //  [機能]  町域検索が押された時
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
	function jf_Cyouiki(){

		if( f_Trim(document.frm.txtCyouiki.value) == ""){
			window.alert("町域指定をしてください");
			document.frm.txtCyouiki.focus();
	        return false;
		}

		document.frm.action         = "Jyusyo_dow.asp";
		document.frm.target         = "dow";
		document.frm.submit();
	}

//-->
</script>
</head>

<body <%= w_jFName %>>
<form name="frm" method="post">
<div align="center">

<% 
    call gs_title("市町村検索","検　索")
%>
<table>
	<tr>
		<td align="center">

			<table border="1" class="hyo">
				<tr>
					<th align="center" class="header">県名</th>
					<th align="center" class="header">市町村名</th>
				<tr>
				<tr>
					<td class="CELL1">
						<select name="selKen" onChange="jf_KenSelect();">
							<option value="">　　　　
							<%
								Do Until m_KenRs.Eof
									'// 変数に"selected"を入れる
									w_Selected = ""
									if Cint(m_KenCd) = Cint(m_KenRs("M16_KEN_CD")) then
										w_Selected = "Selected"
									End if
								%>
								<option value="<%= m_KenRs("M16_KEN_CD") %>" <%=w_Selected%>><%= m_KenRs("M16_KENMEI") %>
								<%
									m_KenRs.MoveNext
								Loop
							%>
						</select>
					</td>
					<td class="CELL1">
						<select name="selShi" style="width:230px;">
							<option value="">　　　　　　　　　　　　　　　　　　　　　　　　　　
							<%
								if m_KenCd <> "" then
									Do Until m_ShiRs.Eof
									'// 変数に"selected"を入れる
									w_Selected = ""
									if m_ShiCd = m_ShiRs("M12_SITYOSON_CD") then
										w_Selected = "selected"
									End if
									%>
									<option value="<%= m_ShiRs("M12_SITYOSON_CD") %>" <%=w_Selected%>><%= m_ShiRs("M12_SITYOSONMEI") %>
									<%
										m_ShiRs.MoveNext
									Loop
								End if
							%>
						</select>
						<input type="button" class="button" value="検索" onClick="jf_ShiSelect();">
					</td>
				</tr>
			</table>

		</td>
	</tr>
	<tr>
		<td align="center">

			<table border="1" class="hyo">
				<tr><th align="center" class="header">検索町域名</th></tr>
				<tr><td class="CELL1"><input type="text" name="txtCyouiki" value="<%= m_Cyouiki %>" size="<%=w_FormSize%>">
						<input type="button" class="button" value="検索" onClick="return jf_Cyouiki();"></td></tr>
			</table>

		</td>
	</tr>
</table>

</div>
<input type="hidden" name="hidKenCd" value="<%= m_KenCd %>">
<input type="hidden" name="hidShiCd">
<input type="hidden" name="hidSchMode" value="True">
</form>
</body>
</html>
<%
End Sub
%>