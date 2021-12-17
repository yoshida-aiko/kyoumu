<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績一覧
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0100_top.asp
' 機      能: 上ページ 成績一覧の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 説      明:
'           ■初期表示
'				コンボボックスは空白で表示
'			■表示ボタンクリック時
'				下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/08/08 前田 智史
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bErrMsg           'ｴﾗｰﾒｯｾｰｼﾞ

    Public m_iNendo				'年度
    Public m_sKyokanCd			'教官コード
	'試験区分用のWhere条件
    Public m_sSikenKBN			'試験区分コンボボックスに入る値
    Public m_sSikenKBNWhere		'氏名コンボボックスの条件
	'クラス用のWhere条件
    Public m_sGakuNo			'クラスの学年コンボボックスに入る値
    Public m_sGakuNoWhere		'クラスの学年コンボボックスの条件
    Public m_sClassNo			'クラスの学科コンボボックスに入る値
    Public m_sClassNoWhere		'クラスの学科コンボボックスの条件

    Public m_sGakkaNo			'学科

	'科目用のWhere条件
    Public m_sKBN				'区分名コンボボックスに入る値
    Public m_sKBNWhere			'区分名コンボボックスの条件

    Public m_sOption			'クラスの学科コンボボックスの使用可、不可の判別
    Public m_sKengen			'表示権限

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
	w_sMsgTitle="成績一覧"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_iNendo	= session("NENDO")
	m_sKyokanCd	= session("KYOKAN_CD")

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

		'//権限を取得
		w_iRet = f_GetKengen_Sei0200(m_sKengen)
		If w_iRet <> 0 Then
            m_bErrFlg = True
            m_sErrMsg = "権限を取得できませんでした"
			Exit Do
		End If

		if m_sKengen = C_SEI0200_ACCESS_TANNIN then
	        '//学年の対象のデータ取得
	        w_iRet = f_getData()
	        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		ElseIf m_sKengen = C_SEI0200_ACCESS_GAKKA then
			w_iRet = f_GetGakkaInfo(m_sKengen)
	        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
		End if

		'試験区分用のWhere条件
        Call f_SikenKBNWhere()

		'クラスの学年用のWhere条件
        Call f_GakuNoWhere()

	If m_sKengen <> C_SEI0200_ACCESS_GAKKA then
		'クラスの組用のWhere条件
		Call  f_ClassNoWhere()
	End If

		'区分用のWhere条件
		Call f_KBNWhere()

	   '// ページを表示
	   Call showPage()
	   Exit Do

	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
    
    '// 終了処理
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  個人履修選択科目決定処理の権限を取得する
'*  [引数]  なし
'*  [戻値]  p_sKengen
'*  [説明]  
'********************************************************************************
Function f_GetKengen_Sei0200(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0340 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('SEI0200','SEI0201','SEI0202') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKengen_Sei0200 = 99
            Exit Do
        End If
		if wLevRs.Eof then
            msMsg = "権限を取得できませんでした"
            f_GetKengen_Sei0200 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "SEI0200" : p_sKengen = C_SEI0200_ACCESS_FULL	  		'//アクセス権限FULLアクセス可
			Case "SEI0201" : p_sKengen = C_SEI0200_ACCESS_TANNIN        '//アクセス担任アクセス
			Case "SEI0202" : p_sKengen = C_SEI0200_ACCESS_GAKKA        '//アクセス担任アクセス
		End Select

		'== 閉じる ==
	    Call gf_closeObject(wLevRs)

        f_GetKengen_Sei0200 = 0
        Exit Do
    Loop

End Function

Function f_getData()
'********************************************************************************
'*  [機能]  学年の対象のデータ取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_getData = 1

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     M05_CLASS "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     M05_NENDO  = " & session("NENDO")
        w_sSQL = w_sSQL & " AND M05_TANNIN = '" & session("KYOKAN_CD") & "' "

        iRet = gf_GetRecordset(wClsRs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_getData = 99
            Exit Do
        End If

		m_sGakuNo  = wClsRs("M05_GAKUNEN")
		m_sClassNo = wClsRs("M05_CLASSNO")

        f_getData = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  権限チェック（ユーザ学科情報取得）
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  ○担任アクセス権限が設定されているUSERでも、実際に担任クラスを
'*			受け持っていない場合には参照不可とする
'********************************************************************************
Function f_GetGakkaInfo(p_sKengen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetGakkaInfo = 1
	p_sKengen = ""

	Do 

		'// 担任クラス情報
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M04_GAKKA_CD "
		w_sSQL = w_sSQL & vbCrLf & " FROM M04_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M04_NENDO=" & session("NENDO")
		w_sSQL = w_sSQL & vbCrLf & "  AND M04_KYOKAN_CD='" & session("KYOKAN_CD") & "'"
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetGakkaInfo = 99
			Exit Do
		End If

		If rs.EOF Then
			'//クラス情報が取得できないとき
            m_sErrMsg = "参照権限がありません。"
			Exit Do
		Else
			p_sKengen = C_SEI0200_ACCESS_GAKKA 
			m_sGakkaNo  = rs("M04_GAKKA_CD")
'			m_sGakkaMei = rs("M02_GAKKAMEI")

			'//権限が担任の場合は、担任クラス以外は選択できない
'			m_sGakuNoOption = " DISABLED "
'			m_sClassNoOption = " DISABLED "
		End If

		f_GetGakkaInfo = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function

Function f_SikenKBNWhere()
'********************************************************************************
'*	[機能]	試験区分コンボに関するWHEREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	m_sSikenKBNWhere=""

		m_sSikenKBNWhere = " M01_NENDO = " & m_iNendo & " "
		m_sSikenKBNWhere = m_sSikenKBNWhere & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN) & " "
		m_sSikenKBNWhere = m_sSikenKBNWhere & " AND M01_SYOBUNRUI_CD < " & cint(C_SIKEN_JITURYOKU) & " "

	m_sSikenKBN = request("txtSikenKBN")

End Function

Function f_GakuNoWhere()
'********************************************************************************
'*	[機能]	クラスの学年コンボに関するWHEREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	m_sGakuNoWhere=""

	m_sGakuNoWhere = " M05_NENDO = " & m_iNendo & " "
	m_sGakuNoWhere = m_sGakuNoWhere & " GROUP BY M05_GAKUNEN "

	if gf_IsNull(m_sGakuNo) then
		m_sGakuNo = request("txtGakuNo")
		If request("txtGakuNo") = C_CBO_NULL Then m_sGakuNo = ""
	End if

End Function

Sub f_ClassNoWhere()
'********************************************************************************
'*	[機能]	クラスの組コンボに関するWHEREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	m_sClassNoWhere=""
	m_sOption=""

	If m_sGakuNo <> "" Then
		m_sClassNoWhere = " M05_NENDO = " & m_iNendo & " AND "
		m_sClassNoWhere = m_sClassNoWhere & " M05_GAKUNEN = " & m_sGakuNo & " "

		if gf_IsNull(m_sClassNo) then
			m_sClassNo = request("txtClassNo")
		End if

	Else
		m_sOption = " DISABLED "
		m_sClassNoWhere  = " M05_GAKUNEN = 99 "
	End IF

End Sub

Sub f_KBNWhere()
'********************************************************************************
'*	[機能]	科目名コンボに関するWHEREを作成する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	m_sKBNWhere=""

		m_sKBNWhere = " M01_NENDO = " & m_iNendo & " AND "
		m_sKBNWhere = m_sKBNWhere & " M01_DAIBUNRUI_CD = " & Cint(C_HISSEN) & " "

End Sub

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
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
//************************************************************
//	[機能]	年度が修正されたとき、再表示する
//	[引数]	なし
//	[戻値]	なし
//	[説明]
//
//************************************************************
function f_ReLoadMyPage(){

	document.frm.action="sei0200_top.asp";
	document.frm.target="top";
	document.frm.submit();

}

//************************************************************
//	[機能]	クリアボタンが押された場合
//	[引数]	なし
//	[戻値]	なし
//	[説明]
//
//************************************************************
function f_Clear(){

	<% if Cint(m_sKengen) = 0 then %>
        document.frm.txtGakuNo.value = "";
        document.frm.txtClassNo.value = "";
	<% End if %>
        document.frm.txtKBN.value = "";

}

//************************************************************
//	[機能]	表示ボタンクリック時の処理
//	[引数]	なし
//	[戻値]	なし
//	[説明]
//
//************************************************************
function f_Search(){

	// ■■■NULLﾁｪｯｸ■■■
	// ■学年
	if( f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
		window.alert("学年の選択を行ってください");
		document.frm.txtGakuNo.focus();
		return ;
	}
<% If m_sKengen <> C_SEI0200_ACCESS_GAKKA then %>
	// ■クラス
	if( f_Trim(document.frm.txtClassNo.value) == "<%=C_CBO_NULL%>" ){
		window.alert("クラスの選択を行ってください");
		document.frm.txtClassNo.focus();
		return ;
	}
<% End If %>

	// ■区分名
	if( f_Trim(document.frm.txtKBN.value) == "<%=C_CBO_NULL%>" ){
		window.alert("区分の選択を行ってください");
		document.frm.txtKBN.focus();
		return ;
	}
	// ■学年
	if( f_Trim(document.frm.txtGakuNo.value) == "" ){
		window.alert("学年の選択を行ってください");
		document.frm.txtGakuNo.focus();
		return ;
	}

<% If m_sKengen <> C_SEI0200_ACCESS_GAKKA then %>
	// ■クラス
	if( f_Trim(document.frm.txtClassNo.value) == "" ){
		window.alert("クラスの選択を行ってください");
		document.frm.txtClassNo.focus();
		return ;
	}
<% End If %>

	// ■区分名
	if( f_Trim(document.frm.txtKBN.value) == "" ){
		window.alert("区分の選択を行ってください");
		document.frm.txtKBN.focus();
		return ;
	}

	document.frm.action="sei0200_main.asp";
	document.frm.target="main";
	document.frm.submit();

}
//-->
</SCRIPT>
<link rel=stylesheet href="../../common/style.css" type=text/css>
</head>

<body>
<center>
<form name="frm" METHOD="post" onClick="return false;">

<% call gs_title(" 成績一覧 "," 一　覧 ") %>
<br>
<table border="0">
	<tr><td valign="bottom">

		<table border="0">
			<tr><td class="search">

				<table border="0">
					<tr>
						<td align="left">試験区分</td>
						<td><%call gf_ComboSet("txtSikenKBN",C_CBO_M01_KUBUN,m_sSikenKBNWhere, "style='width:150px;' ",False,m_sSikenKBN) %></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td align="left">クラス</td>
						<td><% if Cint(m_sKengen) = 1 then wDISABLED = "DISABLED" %>
							<%call gf_ComboSet("txtGakuNo",C_CBO_M05_CLASS_G,m_sGakuNoWhere,"style='width:40px;' onchange='javascript:f_ReLoadMyPage()' " & wDISABLED,True,m_sGakuNo)%>年
<% If m_sKengen <> C_SEI0200_ACCESS_GAKKA then %>
							<%call gf_ComboSet("txtClassNo",C_CBO_M05_CLASS,m_sClassNoWhere,"style='width:80px;' "& m_sOption & wDISABLED ,True,m_sClassNo)%>
<%Else%>
							　<%=gf_GetGakkaNm(m_iNendo,m_sGakkaNo)%>
<%End If%>
						</td>
						<td align="right">　区　分</td>
						<td><%call gf_ComboSet("txtKBN",C_CBO_M01_KUBUN,m_sKBNWhere,"style='width:96px;' ",True,m_sKBN)%></td>
					</tr>
					<tr>
				        <td colspan="4" align="right">
				        <input type="button" class="button" value=" ク　リ　ア " onclick="javasript:f_Clear();">
				        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
				        </td>
					</tr>
				</table>
			</td></tr>
		</table>
	</tr>
</table>
<% if m_sKengen = C_SEI0200_ACCESS_TANNIN then %>
<input type="hidden" name="txtGakuNo"  value="<%=m_sGakuNo %>">
<input type="hidden" name="txtClassNo" value="<%=m_sClassNo%>">
<% ElseIf m_sKengen = C_SEI0200_ACCESS_GAKKA then %>
<input type="hidden" name="txtGakkaNo" value="<%=m_sGakkaNo%>">
<% End if %>
<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">
<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
<input type="hidden" name="txtKengen" value="<%=m_sKengen%>">

</form>
</center>
</body>
</html>
<%
End Sub
%>