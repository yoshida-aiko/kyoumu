<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 使用教科書登録
' ﾌﾟﾛｸﾞﾗﾑID : web/WEB0321/WEB0321_main.asp
' 機	  能: 使用教科書の登録を行う
'-------------------------------------------------------------------------
' 引	  数:教官コード 	＞		SESSIONより（保留）
' 変	  数:なし
' 引	  渡:教官コード 	＞		SESSIONより（保留）
' 説	  明:
'			■フレームページ
'-------------------------------------------------------------------------
' 作	  成: 2001/08/01 前田 智史
' 変	  更: 2001/08/22 伊藤 公子 教官を選択できるように変更
' 変	  更: 2001/12/01 田部 雅幸 所属学科のみを変更するように修正
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
	Public	m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ
	Public	m_iNendo			'年度
	Public	m_sKyokan_CD		'教官CD
	Public	m_sPageCD			':表示済表示頁数（自分自身から受け取る引数）
	Public	m_iMax
	Public	m_Rs
	Public	w_sSQL
	Public	m_iDsp
	Public	m_iDisp 		':表示件数の最大値をとる
	Public	m_sGakka		 '学科名称

	Public m_iGakunen
	Public m_sGakkaCd

	Public	m_sKyokanNm		'//ログイン教官名

	Public m_sSyozokuGakka		'//2001/12/01 Add ログインした教官の所属する学科
	Public m_sKamokuCD()		'//2001/12/01 Add 担当する科目の一覧

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

	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="使用教科書登録"
	w_sMsg=""
	w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""


	On Error Resume Next
	Err.Clear

	m_bErrFlg = False
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

		'// 権限チェックに使用
		session("PRJ_No") = "WEB0321"

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'// 値を変数に入れる
		Call s_SetParam()

		'// 表示用ﾃﾞｰﾀを取得する
		if f_GetData() = False then
			exit do
		end if

		'// 担当科目ﾃﾞｰﾀを取得する
		if f_GetTantoKamoku() = False then
			exit do
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
	Call gf_closeObject(m_Rs)
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	値を変数に入れる
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_SetParam()

	m_iNendo	 = request("txtNendo")					   '年度
	m_iGakunen = trim(replace(Request("txtGakunenCd"),"@@@",""))
	m_sGakkaCd = trim(replace(Request("txtGakkaCD"),"@@@",""))
	m_iDisp = C_PAGE_LINE		'１ページ最大表示数

	'// BLANKの場合は行数ｸﾘｱ
	If Request("txtMode") = "" Then
		m_sPageCD = 1
	Else
		m_sPageCD = INT(Request("txtPageCD"))	':表示済表示頁数（自分自身から受け取る引数）
	End If
	If m_sPageCD = 0 Then m_sPageCD = 1

End Sub

'********************************************************************************
'*	[機能]	表示ﾃﾞｰﾀを取得する
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
function f_GetData()
	Dim w_sSQL				'// SQL文
	Dim w_iRet				'// 戻り値

	Dim w_oRecord			'//2001/12/01 Add 所属学科取得のため

	f_GetData = False

	'2001/12/01 Add ---->
	'//所属学科の取得
	w_sSQL = ""
	w_sSQL = w_sSQL & "SELECT "
	w_sSQL = w_sSQL & "M04_GAKKA_CD "
	w_sSQL = w_sSQL & "From "
	w_sSQL = w_sSQL & "M04_KYOKAN "
	w_sSQL = w_sSQL & "Where "
	w_sSQL = w_sSQL & "M04_NENDO = " & m_iNendo & " "
	w_sSQL = w_sSQL & "And "
	w_sSQL = w_sSQL & "M04_KYOKAN_CD = '" & Session("KYOKAN_CD") & "'"

	
	w_iRet = gf_GetRecordset(w_oRecord, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit function
	End If

	If w_oRecord.EOF <> True Then
		m_sSyozokuGakka = w_oRecord("M04_GAKKA_CD")
	Else
		m_sSyozokuGakka =""
	End If

	'//閉じる
	w_oRecord.Close
	Set w_oRecord = Nothing

	'2001/12/01 Add <----


'	 w_sSQL = w_sSQL & vbCrLf & " SELECT "
'	 w_sSQL = w_sSQL & vbCrLf & " T47.T47_NENDO "			 ''年度
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKI_KBN "		 ''学期区分
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_NO"				 ''No
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKUNEN "		 ''学年
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKA_CD "		 ''学科
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KAMOKU " 		 ''科目ｺｰﾄﾞ
'''  w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKAN "
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKASYO "		 ''教科書名
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SYUPPANSYA " 	 ''出版社
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_TYOSYA " 		 ''著者
'''  w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKUSEISU "
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKANYOUSU "	 ''教官用数
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SIDOSYOSU "		 ''指導書数
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_BIKOU "			 ''備考
'''  w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_NENDO "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKA_CD "
'	 w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKAMEI "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_NENDO "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKU_CD "
'	 w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_NENDO "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKAN_CD "
'	 w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
'	 w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
'	 w_sSQL = w_sSQL & vbCrLf & " FROM "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47_KYOKASYO T47 "
'	 w_sSQL = w_sSQL & vbCrLf & "    ,M02_GAKKA M02 "
'	 w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
'	 w_sSQL = w_sSQL & vbCrLf & "    ,M04_KYOKAN M04 "
'	 w_sSQL = w_sSQL & vbCrLf & " WHERE "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M02.M02_NENDO(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_GAKKA_CD  = M02.M02_GAKKA_CD(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M03.M03_NENDO(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M04.M04_NENDO(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = M04.M04_KYOKAN_CD(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO = " & m_iNendo & " "
'	 'w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = '" & m_sKyokan_CD & "' "
'	 w_sSQL = w_sSQL & vbCrLf & " ORDER BY T47.T47_GAKKA_CD "



	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "  T47_KYOKASYO.T47_GAKKI_KBN "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_NO "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_GAKUNEN "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_GAKKA_CD "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KAMOKU "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KYOKAN "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KYOKASYO "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_SYUPPANSYA "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_TYOSYA"
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "  T47_KYOKASYO"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "  T47_KYOKASYO.T47_NENDO=" & m_iNendo

	If m_iGakunen <> "" Then
		w_sSQL = w_sSQL & vbCrLf & "  AND T47_KYOKASYO.T47_GAKUNEN=" & m_iGakunen
	End If

'2001/12/01 Mod ---->
'	If m_sGakkaCd <> "" Then
'		w_sSQL = w_sSQL & vbCrLf & "  AND T47_KYOKASYO.T47_GAKKA_CD='" & m_sGakkaCd & "'"
'	End If

	w_sSQL = w_sSQL & vbCrLf & "  AND T47_KYOKASYO.T47_GAKKA_CD='" & m_sSyozokuGakka & "'"

'2001/12/01 Mod <----

	w_sSQL = w_sSQL & vbCrLf & " ORDER BY "
	w_sSQL = w_sSQL & vbCrLf & "  T47_KYOKASYO.T47_GAKUNEN"
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_GAKKA_CD"
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KAMOKU"
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KYOKAN"
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KYOKASYO "

'response.write("<BR>w_sSQL = " & w_sSQL)

	Set m_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		m_bErrFlg = True
		Exit function
	Else
		'ページ数の取得
		m_iMax = gf_PageCount(m_Rs,m_iDsp)
	End If

	f_GetData = True

End Function

'********************************************************************************
'*	[機能]	学科の略称を取得
'*	[引数]	p_sGakkaCd : 学科CD
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_GetGakkaNm_R(p_sGakkaCd)
	Dim w_sSQL				'// SQL文
	Dim w_iRet				'// 戻り値
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakkaNm_R = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKARYAKSYO"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_GAKKA_CD='" & p_sGakkaCd & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M02_GAKKARYAKSYO")
		End If 

		Exit do 
	Loop

	'//戻り値をセット
	f_GetGakkaNm_R = w_sName

	'//RS Close
	Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*	[機能]	科目名称を取得
'*	[引数]	p_sGakkaCd : 学科CD
'*			p_sKamokuCd
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_GetKamokuNm(p_sGakkaCd,p_sKamokuCd)
	Dim w_sSQL				'// SQL文
	Dim w_iRet				'// 戻り値
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetKamokuNm = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & m_iNendo

		if cstr(gf_HTMLTableSTR(p_sGakkaCd)) <> cstr(C_CLASS_ALL) then
			w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & p_sGakkaCd & "'"
		End If
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_sKamokuCd & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("T15_KAMOKUMEI")
		End If 

		Exit do 
	Loop

	'//戻り値をセット
	f_GetKamokuNm = w_sName

	'//RS Close
	Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*	[機能]	詳細を表示
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub S_syousai()

	Dim w_iCnt
	Dim w_i
	Dim w_cell

	Dim w_lCnt			'カウンタ
	Dim w_bTantoYes		'担当している

	w_iCnt	= 0
	w_i 	= 0
	w_cell = ""

	Dim w_sCurGakkaCD		'2001/12/01 Add 処理中の学科ＣＤ

	Do While not m_Rs.EOF

		w_i = w_i + 1

		call gs_cellPtn(w_cell)
%>

		<Tr>
		<Td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_GAKUNEN")) %>年</Td>
<%
		If CStr(gf_HTMLTableSTR(m_Rs("T47_GAKKA_CD"))) = CStr(C_CLASS_ALL) Then
			m_sGakka = "全学科"
			w_sCurGakkaCD = ""								'2001/12/01 Add
		Else
			m_sGakka = f_GetGakkaNm_R(m_Rs("T47_GAKKA_CD"))
			w_sCurGakkaCD = CStr(m_Rs("T47_GAKKA_CD"))		'2001/12/01 Add
		End If

		w_bTantoYes= False

		For w_lCnt = 0 To Ubound(m_sKamokuCD)
			If m_sKamokuCD(w_lCnt) = CStr(m_Rs("T47_KAMOKU")) Then
				w_bTantoYes = True
				Exit For
			End if
		Next

		'2001/12/01 Mod ---->
		If w_bTantoYes = True Then
		'== 全学科か所属する学科の場合 ==
%>
<!-- <%= m_sGakka %> -->
		<Td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_sGakka) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(f_GetKamokuNm(m_Rs("T47_GAKKA_CD"),m_Rs("T47_KAMOKU"))) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(gf_GetKyokanNm(m_iNendo,m_Rs("T47_KYOKAN"))) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><A HREF='javascript:f_LinkClick(<%=m_Rs("T47_NO")%>);'><%=gf_HTMLTableSTR(m_Rs("T47_KYOKASYO")) %></A></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_SYUPPANSYA")) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_TYOSYA")) %></Td>
		<Td align="center" width="30"  class=<%=w_cell%>><input class=button type="button" value=">>" onclick="javascript:f_Update(<%=gf_HTMLTableSTR(m_Rs("T47_NO")) %>)"></Td>
		<Td align="center" width="30"  class=<%=w_cell%>><input type="checkbox" name="deleteNO" value="<%=gf_HTMLTableSTR(m_Rs("T47_NO")) %>"></Td>

<%
		Else
		'== 全学科でも所属する学科でもない場合 ==
%>
<!-- <%= m_sGakka %> -->

		<Td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_sGakka) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(f_GetKamokuNm(m_Rs("T47_GAKKA_CD"),m_Rs("T47_KAMOKU"))) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(gf_GetKyokanNm(m_iNendo,m_Rs("T47_KYOKAN"))) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><A HREF='javascript:f_LinkClick(<%=m_Rs("T47_NO")%>);'><%=gf_HTMLTableSTR(m_Rs("T47_KYOKASYO")) %></A></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_SYUPPANSYA")) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_TYOSYA")) %></Td>
		<Td align="center" width="30"  class=<%=w_cell%>>　</Td>
		<Td align="center" width="30"  class=<%=w_cell%>>　</Td>

<%
		End If

		m_Rs.MoveNext

		If w_iCnt >= C_PAGE_LINE-1 Then
			Exit Sub
		Else
			w_iCnt = w_iCnt + 1
		End If
	Loop

	m_iDisp= w_i

End sub

Sub showPage()
'********************************************************************************
'*	[機能]	HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
	Dim w_bFlg				'// ﾃﾞｰﾀ有無
	Dim w_bNxt				'// NEXT表示有無
	Dim w_bBfr				'// BEFORE表示有無
	Dim w_iNxt				'// NEXT表示頁数
	Dim w_iBfr				'// BEFORE表示頁数
	Dim w_iCnt				'// ﾃﾞｰﾀ表示ｶｳﾝﾀ
	Dim w_pageBar			'ページBAR表示用
	
	On Error Resume Next
	Err.Clear

	'ページBAR表示
	Call gs_pageBar(w_pageBar)

	Dim w_iRecordCnt		'//レコードセットカウント

	On Error Resume Next
	Err.Clear

	w_iCnt	= 1
	w_bFlg	= True

%>

	<html>

	<head>

	<title>使用教科書登録</title>
<!-- <%= m_sSyozokuGakka %>-->
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

	<!--

	//************************************************************
	//	[機能]	一覧表の次・前ページを表示する
	//	[引数]	p_iPage :表示頁数
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_PageClick(p_iPage){

		document.frm.action="";
		document.frm.target="";
		document.frm.txtMode.value = "PAGE";
		document.frm.txtPageCD.value = p_iPage;
		document.frm.submit();
	
	}

	//************************************************************
	//	[機能]	更新画面へ
	//	[引数]	p_iPage :表示頁数
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Update(p_No){

		document.frm.action="./touroku.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.txtMode.value = "Kousin";
		document.frm.txtUpdNo.value = p_No;
		document.frm.submit();

	}

	//************************************************************
	//	[機能]	削除ページへ
	//	[引数]	p_iPage :表示頁数
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_Delete(){

		if (f_chk()==1){
		alert( "削除の対象となる教科書が選択されていません" );
		return;
		}

		document.frm.action="del_kakunin.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.txtMode.value = "DELETE";
		document.frm.submit();
	}
	//************************************************************
	//	[機能]	リスト一覧のチェックボックスの確認
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//
	//************************************************************
	function f_chk(){

		var i;
		i = 0;

		//0件のとき
		if (document.frm.txtDisp.value<=0){
			return 1;
			}

		//1件のとき
		if (document.frm.txtDisp.value==1){
			if (document.frm.deleteNO.checked == false){
				return 1;
			}else{
				return 0;
				}
		}else{
		//それ以外の時
		var checkFlg
			checkFlg=false

		do { 
			
			if(document.frm.deleteNO[i].checked == true){
				checkFlg=true
				break;
			 }

		i++; }	while(i<document.frm.txtDisp.value);
			if (checkFlg == false){
				return 1;
				}
		}
		return 0;
	}

	//************************************************************
	//	[機能]	リンククリック
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function f_LinkClick(p_No){
		document.frm.txtUpdNo.value = p_No;
		document.frm.action="view.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.submit();
	}

	//-->
	</SCRIPT>
	<link rel=stylesheet href="../../common/style.css" type=text/css>

	</head>

	<body>

	<center>
	<br>

	<form name="frm" action="touroku.asp" target="" Method="POST">

	<%
	'データなしの場合
	If m_Rs.EOF Then
	%>
		<br><br><br>
		<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
	<%Else%>


		<span class="msg"><font size="2">※教科書名をクリックすると詳細内容を参照できます。</font></span>
	<%Call gs_pageBar(m_Rs,m_sPageCD,m_iDsp,w_pageBar)%>

		<table width=90%>
			<Tr><Td><%=w_pageBar %></Td></Tr>

			<Tr><Td>
				<table border="1" width="90%" class=hyo>
				<Tr>
					<Th width="70"	class=header nowrap>学年</Th>
					<Th width="70"	class=header nowrap>学科</Th>
					<Th width="110" class=header nowrap>科目</Th>
					<Th width="110" class=header nowrap>教官名</Th>
					<Th width="150" class=header nowrap>教科書名</Th>
					<Th width="90"	class=header nowrap>出版社</Th>
					<Th width="90"	class=header nowrap>著者</Th>
					<Th width="30"	class=header >修正</Th>
					<Th width="30"	class=header>削除</Th>
				</Tr>

					<% S_syousai() %>
				<Tr>
					<Td colspan=9 align=right bgcolor=#9999BD>
					<input class=button type=button value="×削除" Onclick="f_Delete()"></Td>
				</Tr>

				</table>
			</Td></Tr>
<!--
<% = Ubound(m_sKamokuCD) %>
<%
		For w_lCnt = 0 To Ubound(m_sKamokuCD)
			response.write(m_sKamokuCD(w_lCnt))
		Next
%>
-->
			<Tr><Td><%=w_pageBar %></Td></Tr>
		<table>
	<%End If%>

	<!--値渡用-->
	<input type="hidden" name="txtMode" value="Touroku">
	<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
	<input type="hidden" name="txtDisp" value="<%= m_iDisp %>">
	<input type="hidden" name="txtUpdNo" value="">
	<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">

	<input type="hidden" name="KeyNendo" value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokan_CD%>">
	<input type="hidden" name="SKyokanCd1" value="<%=m_sKyokan_CD%>">

	<input type="hidden" name="txtGakunenCd" value="<%= Request("txtGakunenCd") %>">
	<input type="hidden" name="txtGakkaCD"	 value="<%= Request("txtGakkaCD") %>">

	</form>
	</center>
	</body>
	</html>

<%
End Sub


Function f_GetTantoKamoku()
'********************************************************************************
'*	[機能]　担当教官かどうかチェック
'*	[引数]　 なし
'*	[戻値]　True:担当教官をしている、False:担当教官をしていない
'*	[説明]
'********************************************************************************
	Dim w_iRet			'戻り値
	Dim w_sSQL			'SQL
	Dim w_oRecord		'レコード
	Dim w_lCnt			'レコードカウント

	f_GetTantoKamoku = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "     T27_KAMOKU_CD"
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & "     T27_TANTO_KYOKAN T27"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "     T27_NENDO = " & SESSION("NENDO") & " "
	If request("txtGakunen") <> "" Then
		w_sSQL = w_sSQL & " AND "
		w_sSQL = w_sSQL & " T27_GAKUNEN = " & request("txtGakunen") & " "
	End If
	w_sSQL = w_sSQL & " AND "
	w_sSQL = w_sSQL & " T27_KYOKAN_CD = '" & SESSION("KYOKAN_CD") & "' "
	w_sSQL = w_sSQL & " GROUP BY T27_KAMOKU_CD"
	w_sSQL = w_sSQL & " ORDER BY T27_KAMOKU_CD"

	Set w_oRecord = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset_OpenStatic(w_oRecord, w_sSQL)

	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
'		response.write(w_sSQL)
		Exit Function
	End If

	'//担当していない場合
	If w_oRecord.EOF = True Then
		ReDim m_sKamokuCD(0)

		f_GetTantoKamoku = True
		Exit Function
	End If

	w_lCnt = gf_GetRsCount(w_oRecord)
'	w_oRecord.MoveFirst

	ReDim m_sKamokuCD(w_lCnt)

	'// 担当科目の保持
	For w_lCnt = 0 To Ubound(m_sKamokuCD) - 1
		m_sKamokuCD(w_lCnt) = CStr(w_oRecord("T27_KAMOKU_CD"))

		w_oRecord.MoveNext
	Next

	w_oRecord.Close
	Set w_oRecord = Nothing

	f_GetTantoKamoku = True
'response.write("True<BR>")

End Function






%>