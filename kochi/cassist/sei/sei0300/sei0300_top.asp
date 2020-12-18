<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績一覧
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0300/sei0300_top.asp
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
' 作      成: 2001/09/04 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////

	Public CONST C_KENGEN_SEI0300_FULL = "FULL"	'//アクセス権限FULL
	Public CONST C_KENGEN_SEI0300_TAN = "TAN"	'//アクセス権限担任
	Public CONST C_KENGEN_SEI0300_GAK = "GAK"	'//アクセス権限学科

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
    Public m_sGakusei			'クラスの学年コンボボックスに入る値
    Public m_sGakuseiWhere		'クラスの学年コンボボックスの条件
    Public m_sGakkaNo			'学科コード

    Public m_sKengen			'権限
'    Public m_bTannin
'    Public m_bGakka

    Public m_sOption			'クラスの学科コンボボックスの使用可、不可の判別
    Public m_sGakuNoOption
    Public m_sClassNoOption
    Public m_sGakuseiOption


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
	w_sMsgTitle="個人別成績一覧"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

'	m_iNendo	= session("NENDO")
'	m_sKyokanCd	= session("KYOKAN_CD")

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

		'//パラメータセット
		Call s_SetParam

		'//権限チェック
		w_iRet = f_CheckKengen(w_sKengen)
		If w_iRet <> 0 Then
            m_bErrFlg = True
			m_sErrMsg = "参照権限がありません。"
			Exit Do
		End If

		'//権限が担任の場合は担任クラス情報を取得する
		If w_sKengen = C_KENGEN_SEI0300_TAN Then

			'//担任クラス情報取得
			'//情報が取得できない場合は担任クラスが無い為、参照不可とする
			w_iRet = f_GetClassInfo(m_sKengen)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				m_sErrMsg = "参照権限がありません。"
				Exit Do
			End If

		ElseIf w_sKengen = C_KENGEN_SEI0300_GAK Then

			'//学科情報取得
			'//情報が取得できない場合は学科が無い為、参照不可とする
			w_iRet = f_GetGakkaInfo(m_sKengen)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				m_sErrMsg = "参照権限がありません。"
				Exit Do
			End If

		End If

		'試験区分用のWhere条件
        Call f_SikenKBNWhere()

		'クラスの学年用のWhere条件
        Call f_GakuNoWhere()

	If w_sKengen <> C_KENGEN_SEI0300_GAK Then
		'クラスの組用のWhere条件
		Call  f_ClassNoWhere()
	End If
	
		'区分用のWhere条件
		Call f_GakuseiWhere()
		
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
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

	m_iNendo	= session("NENDO")
	m_sKyokanCd	= session("KYOKAN_CD")

	m_sSikenKBN = trim(Request("txtSikenKBN"))
	m_sGakuNo   = Replace(trim(Request("txtGakuNo")),"@@@","")
	m_sClassNo  = Replace(trim(Request("txtClassNo")),"@@@","")
	m_sGakusei  = Replace(trim(Request("txtGakusei")),"@@@","")

End Sub

'********************************************************************************
'*	[機能]	権限チェック
'*	[引数]	なし
'*	[戻値]	w_sKengen
'*	[説明]	ログインUSERの処理レベルにより、参照可不可の判断をする
'*			①FULLアクセス権限保持者は、全ての生徒の成績情報を参照できる
'*			②担任アクセス権限保持者は、受け持ちクラス生徒の成績情報を参照できる
'*			③上記以外のUSERは参照権限なし
'********************************************************************************
Function f_CheckKengen(p_sKengen)
    Dim w_iRet
    Dim w_sSQL
	 Dim rs

	 On Error Resume Next
	 Err.Clear

	 f_CheckKengen = 1

	 Do

		'T51より権限情報取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL.T51_ID "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL.T51_ID IN ('SEI0300','SEI0301','SEI0302')"
		w_sSql = w_sSql & vbCrLf & "  AND T51_SYORI_LEVEL.T51_LEVEL" & Session("LEVEL") & " = 1"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			m_sErrMsg = Err.description
			f_CheckKengen = 99
			Exit Do
		End If

		If rs.EOF Then
			m_sErrMsg = "参照権限がありません。"
			Exit Do
		Else
			Select Case rs("T51_ID")
				Case "SEI0300"	'//フルアクセス権限あり
					p_sKengen = C_KENGEN_SEI0300_FULL
				Case "SEI0301"	'//担任権限有り
					p_sKengen = C_KENGEN_SEI0300_TAN
				Case "SEI0302"	'//学科権限有り
					p_sKengen = C_KENGEN_SEI0300_GAK
			End Select

		End If

		f_CheckKengen = 0
		Exit Do
	 Loop


	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  権限チェック（担任クラス情報取得）
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  ○担任アクセス権限が設定されているUSERでも、実際に担任クラスを
'*			受け持っていない場合には参照不可とする
'********************************************************************************
Function f_GetClassInfo(p_sKengen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassInfo = 1
	p_bTannin = False

	Do 

		'// 担任クラス情報
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS.M05_CLASSNO "
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & m_sKyokanCd & "'"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetClassInfo = 99
			Exit Do
		End If

		If rs.EOF Then
			'//クラス情報が取得できないとき
            m_sErrMsg = "参照権限がありません。"
			Exit Do
		Else
			p_sKengen = C_KENGEN_SEI0300_TAN 
			m_sGakuNo  = rs("M05_GAKUNEN")
			m_sClassNo = rs("M05_CLASSNO")

			'//権限が担任の場合は、担任クラス以外は選択できない
			m_sGakuNoOption = " DISABLED "
			m_sClassNoOption = " DISABLED "
		End If

		f_GetClassInfo = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)

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
		w_sSQL = w_sSQL & vbCrLf & "      M04_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M04_KYOKAN_CD='" & m_sKyokanCd & "'"
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
			p_sKengen = C_KENGEN_SEI0300_GAK 
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
	Else
		m_sClassNoOption = " DISABLED "
		m_sClassNoWhere  = " M05_GAKUNEN = 99 "
	End IF

End Sub

Sub f_GakuseiWhere()
'********************************************************************************
'*  [機能]  氏名コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sGakuseiWhere=""

 	If m_sClassNo <> "" Then
	    m_sGakuseiWhere = " T11_GAKUSEI_NO = T13_GAKUSEI_NO "
'	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T11_NYUNENDO = T13_NENDO - T13_GAKUNEN + 1"
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_GAKUNEN = " & m_sGakuNo
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_CLASS = " & m_sClassNo
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_NENDO = " & m_iNendo
		m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO

	ElseIf m_sKengen = C_KENGEN_SEI0300_GAK AND m_sGakuNo <> "" Then
	    m_sGakuseiWhere = " T11_GAKUSEI_NO = T13_GAKUSEI_NO "
'	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T11_NYUNENDO = T13_NENDO - T13_GAKUNEN + 1"
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_GAKUNEN = " & m_sGakuNo
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_GAKKA_CD = " & m_sGakkaNo
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_NENDO = " & m_iNendo
		m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO

	Else
		m_sGakuseiOption = " DISABLED "
		m_sGakuseiWhere  = " T13_GAKUNEN = 99 "
	End IF

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
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
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

		document.frm.action="sei0300_top.asp";
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

		document.frm.txtGakuNo.value = "";
		document.frm.txtClassNo.value = "";
		document.frm.txtGakusei.value = "";

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
	<%If m_sKengen <> C_KENGEN_SEI0300_GAK Then%>
		// ■クラス
		if( f_Trim(document.frm.txtClassNo.value) == "<%=C_CBO_NULL%>" ){
			window.alert("クラスの選択を行ってください");
			document.frm.txtClassNo.focus();
			return ;
		}
	<%End If%>
		// ■学生
		if( f_Trim(document.frm.txtGakusei.value) == "<%=C_CBO_NULL%>" ){
		    window.alert("学生の選択を行ってください");
		    document.frm.txtGakusei.focus();
		    return ;
		}

		// ■学年
		if( f_Trim(document.frm.txtGakuNo.value) == "" ){
			window.alert("学年の選択を行ってください");
			document.frm.txtGakuNo.focus();
			return ;
		}
	<%If m_sKengen <> C_KENGEN_SEI0300_GAK Then%>
		// ■クラス
		if( f_Trim(document.frm.txtClassNo.value) == "" ){
			window.alert("クラスの選択を行ってください");
			document.frm.txtClassNo.focus();
			return ;
		}
	<%End If%>
		// ■学生
		if( f_Trim(document.frm.txtGakusei.value) == "<%=C_CBO_NULL%>" ){
		    window.alert("学生の選択を行ってください");
		    document.frm.txtGakusei.focus();
		    return ;
		}

		document.frm.action="sei0300_main.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.submit();

	}
	//-->
	</SCRIPT>
	<link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>

	<body>
	<center>
	<form name="frm" METHOD="post">

	<% call gs_title(" 個人別成績一覧 "," 一　覧 ") %>
	<br>
	<table border="0">
		<tr><td valign="bottom">

			<table border="0">
				<tr><td class="search">

					<table border="0">
						<tr>
							<td align="left" nowrap>試験区分</td>
							<td nowrap><%call gf_ComboSet("txtSikenKBN",C_CBO_M01_KUBUN,m_sSikenKBNWhere, "style='width:150px;' ",False,m_sSikenKBN) %></td>
						<td></td>
						<td></td>
						</tr>
						<tr>
							<td align="left" nowrap>クラス</td>
							<td nowrap><%
								call gf_ComboSet("txtGakuNo",C_CBO_M05_CLASS_G,m_sGakuNoWhere,"style='width:40px;' onchange='javascript:f_ReLoadMyPage()' " & m_sGakuNoOption,True,m_sGakuNo)%>年　<%
							If m_sKengen <> C_KENGEN_SEI0300_GAK then 
								call gf_ComboSet("txtClassNo",C_CBO_M05_CLASS,m_sClassNoWhere,"style='width:80px;'  onchange='javascript:f_ReLoadMyPage()' "& m_sClassNoOption,True,m_sClassNo)
							else %>
								<%=gf_GetGakkaNm(m_iNendo,m_sGakkaNo)%>
						 <% end if %>
							</td>
					            <td Nowrap align="center">　氏　名　
								<%call gf_PluComboSet("txtGakusei",C_CBO_T11_GAKUSEKI_N,m_sGakuseiWhere, "style='width:250px;'"& m_sGakuseiOption,True,m_sGakusei)%>
							</td>
						</tr>
						<tr>
					        <td colspan="4" align="right" valign="bottom"  nowrap>
					        <input type="button" class="button" value=" ク　リ　ア " onclick="javasript:f_Clear();">
					        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
					        </td>
						</tr>
					</table>
				</td></tr>
			</table>
		</tr>
	</table>

	<%If m_sKengen=C_KENGEN_SEI0300_TAN Then%>
		<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">
		<input type="hidden" name="txtClassNo" value="<%=m_sClassNo%>">
	<%ElseIf m_sKengen=C_KENGEN_SEI0300_GAK Then%>
		<input type="hidden" name="txtGakkaNo" value="<%=m_sGakkaNo%>">
	<%End If%>

	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>