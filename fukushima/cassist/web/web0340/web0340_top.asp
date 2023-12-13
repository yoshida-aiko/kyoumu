<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 個人履修選択科目決定
' ﾌﾟﾛｸﾞﾗﾑID : web/web0340/web0340_top.asp
' 機      能: 上ページ 個人履修選択科目決定の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/23 前田 智史
' 変      更: 2001/08/07 根本 直美     NN対応に伴うソース変更
' 変      更: 2015/03/20 清本 千秋     Win7対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '市町村選択用のWhere条件
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官コード
    Public m_sKBN           '区分コンボボックスに入る値
    Public m_sGRP           '氏名コンボボックスに入る値
    Public m_sKBNWhere      '年度コンボボックスの条件
    Public m_sGRPWhere      '氏名コンボボックスの条件
    Public m_sOption        '氏名コンボボックスの使用可、不可の判別
    Public m_sGakunen       '学年
    Public m_sClass         'クラス
    Public m_sGakka         '学科コード
    Public m_rs             '
    Public m_sGakunenWhere      '//学年の条件
    Public m_sGakunenOption     '//学年コンボのオプション
    Public m_sClassWhere        '//クラスの条件
    Public m_sClassOption       '//クラスコンボのオプション
    Public m_sKengen

    Public  m_iMax          '最大ページ
    Public  m_iDsp          '一覧表示行数

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
    w_sMsgTitle="個人履修選択科目決定"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

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
		w_iRet = gf_GetKengen_web0340(m_sKengen)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'//データを変数にセット
		Call s_SetParam()

		'//区分コンボに関するWHEREを作成する
        w_iRet = f_KBNWhere()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//グループコンボに関するWHEREを作成する
        Call f_GRPWhere()

'//デバッグ
'call s_DebugPrint


        '//学年コンボに関するWHEREを作成する
        Call s_MakeGakunenWhere() 

        '//クラスコンボに関するWHEREを作成する
        Call s_MakeClassWhere() 

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
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()


    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp = C_PAGE_LINE

	'//権限が担任の場合は、担任クラスのみ登録を可能とする
	If m_sKengen = C_WEB0340_ACCESS_TANNIN Then

		'//担任教官の場合は、担任クラスの年組を取得する
		Call f_Gakunen()
	Else
		'//担任以外の場合
	    m_sGakunen  = Request("cboGakunenCd")   '//学年
	    m_sClass    = Request("cboClassCd")     '//クラス

	End If

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iNendo     = " & m_iNendo     & "<br>"
    response.write "m_sKyokanCd  = " & m_sKyokanCd  & "<br>"
    response.write "m_sGakunen   = " & m_sGakunen   & "<br>"
    response.write "m_sClass     = " & m_sClass     & "<br>"
    response.write "m_sGakka     = " & m_sGakka     & "<br>"

End Sub

'********************************************************************************
'*  [機能]  学年コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeGakunenWhere()

    m_sGakunenWhere = ""
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iNendo
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

	If m_sKengen = C_WEB0340_ACCESS_TANNIN Then
		m_sGakunenOption = "DISABLED"
	End If

End Sub

'********************************************************************************
'*  [機能]  クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeClassWhere()

    m_sClassWhere = ""
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iNendo

    If m_sGakunen = "" Then
        '//初期表示時は1年1組を表示する
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = 1"
    Else
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_sGakunen)
    End If

	'//権限が担任の場合は、担任クラス以外の登録は出来ない
	If m_sKengen = C_WEB0340_ACCESS_TANNIN Then
		m_sClassOption = "DISABLED"
	End If

End Sub

Function f_KBNWhere()
'********************************************************************************
'*  [機能]  区分コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_KBNWhere = 1

    Do

        m_sKBNWhere = ""
        m_sKBNWhere = m_sKBNWhere & " M01_DAIBUNRUI_CD = " & C_KAMOKU & " AND "
        m_sKBNWhere = m_sKBNWhere & " M01_NENDO        = " & m_iNendo & " AND "
        m_sKBNWhere = m_sKBNWhere & " M01_SYOBUNRUI_CD <> 2 "

        m_sKBN = request("txtKBN")

        If request("txtKBN") = C_CBO_NULL Then m_sKBN = ""

        f_KBNWhere = 0
        Exit Do
    Loop


End Function

Sub f_GRPWhere()
'********************************************************************************
'*  [機能]  グループコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_iNyuNendo

	Dim w_sGroup,w_bFlg

    m_sGRPWhere=""
    m_sOption=""
	w_sGroup = ""
	w_bDispFlg = True

    If m_sKBN <> "" Then

		'=============
		'//学科の取得
		'=============
		Call f_GetGakka(m_sGakunen,m_sClass)

		'==============================================================
		'//権限が専門教官の場合は、担当科目関連のみ登録を可能とする
		'==============================================================
		If m_sKengen = C_WEB0340_ACCESS_SENMON Then

			'//専門教官の場合は、関連教科の科目のみ入力可とする
			'//科目のグループを取得(複数)
			Call f_GetGroup(m_sGakunen,m_sClass,w_sGroup)

			If Trim(w_sGroup) = "" Then
				w_bDispFlg = False
			End If

		End If

		'==============================================================
		'//グループコンボを取得
		'==============================================================
		If w_bDispFlg = True Then

	        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

	        m_sGRPWhere = " T15_HISSEN_KBN = " & C_HISSEN_SEN & " AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_GAKKA_CD = " & m_sGakka & " AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_NYUNENDO = " & w_iNyuNendo & " AND "
	        m_sGRPWhere = m_sGRPWhere & " T15_KAMOKU_KBN = " & cInt(m_sKBN) & " AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_NYUNENDO = T15_NYUNENDO(+) AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_GRP = T15_GRP(+) AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_GAKKA_CD = T15_GAKKA_CD(+) AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_GRP <> " & C_T18_GRP & " "

			If w_sGroup <> "" Then
		        m_sGRPWhere = m_sGRPWhere & " AND T18_GRP IN (" & w_sGroup & ")"
			End If

	        m_sGRPWhere = m_sGRPWhere & " GROUP BY T18_GRP,T18_SYUBETU_MEI "

		Else
	        m_sOption = " DISABLED "
	        m_sGRPWhere  = " T18_GAKKA_CD = 00 "
		End If

    Else
        m_sOption = " DISABLED "
        m_sGRPWhere  = " T18_GAKKA_CD = 00 "
    End IF

End Sub

Sub f_Gakunen()
'********************************************************************************
'*  [機能]  学年の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    '//学年･クラスのデータ
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT"
    w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_GAKKA_CD "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     M05_CLASS "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & "     M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "

    Set m_rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(m_rs, w_sSQL,m_iDsp)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Sub
    End If

	If m_rs.EOF = false Then
	    m_sGakka   = m_rs("M05_GAKKA_CD")
	    m_sGakunen = cInt(m_rs("M05_GAKUNEN"))
	    m_sClass   = cInt(m_rs("M05_CLASSNO"))
	End If

   Call gf_closeObject(m_rs)

End Sub

Sub f_GetGakka(p_sGakuNen,p_sClass)
'********************************************************************************
'*  [機能]  学科の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	Dim rs
	Dim w_sSQL

    '//学年･クラスより学科を取得
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT"
    w_sSQL = w_sSQL & "     M05_GAKKA_CD "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     M05_CLASS "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & "     M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "     AND M05_GAKUNEN = " & p_sGakuNen
    w_sSQL = w_sSQL & "     AND M05_CLASSNO = " & p_sClass

    w_iRet = gf_GetRecordset(rs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Sub
    End If

	If rs.EOF = false Then
	    m_sGakka   = rs("M05_GAKKA_CD")
	End If

   Call gf_closeObject(rs)

End Sub

Function f_GetGroup(p_sGakuNen,p_sClass,p_sGroup)
'********************************************************************************
'*  [機能]  学科の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	Dim rs
	Dim w_sSQL
	Dim w_sKamoku

	w_sKamoku = ""
	p_sGroup = ""

	Do 

	    '//教官の関連科目を取得する
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN.T27_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN.T27_NENDO= " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_GAKUNEN=" & p_sGakuNen
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_CLASS=" & p_sClass
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_KYOKAN_CD=" & m_sKyokanCd
'response.write w_sSQL
'メイン教官フラグが立っている教官のみ入力可能　add 2001/10/29 tani
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_MAIN_FLG=" & C_MAIN_KYOKAN_YES

'response.write w_sSQL

	    w_iRet = gf_GetRecordset(rs, w_sSQL)
	    If w_iRet <> 0 Then
	        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
	        m_bErrFlg = True
	        Exit Do
	    End If

		If rs.EOF = True Then
			Exit Do
		Else

			'//科目CDを取得
			Do Until rs.EOF
				If w_sKamoku = "" Then
				    w_sKamoku = rs("T27_KAMOKU_CD")
				Else
					w_sKamoku = w_sKamoku  & "," & rs("T27_KAMOKU_CD")
				End If
				rs.MoveNext
			Loop

		End If

		'//科目CDが取得できないとき
		If Trim(w_sKamoku) = "" Then
			Exit Do
		End If

		'//科目のグループの種類を取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_GRP"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(p_sGakuNen) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD= '" & m_sGakka & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD IN (" & w_sKamoku & ")"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_KBN=" & m_sKBN 
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_HISSEN_KBN=" & C_HISSEN_SEN '//選択科目のみ

	    w_iRet = gf_GetRecordset(rs_K, w_sSQL)
	    If w_iRet <> 0 Then
	        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
	        m_bErrFlg = True
	        Exit Do
	    End If

		If rs_K.EOF Then
			Exit Do
		Else

			Do Until rs_K.EOF

				If p_sGroup = "" Then
				    p_sGroup = rs_K("T15_GRP")
				Else
				    p_sGroup = p_sGroup & "," & rs_K("T15_GRP")
				End If

				rs_K.MoveNext
			Loop

		End If

		Exit Do
	Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
   Call gf_closeObject(rs)
   Call gf_closeObject(rs_K)

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

<title>個人履修選択科目決定</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [機能]  年度が修正されたとき、再表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="web0340_top.asp";
        document.frm.target="top";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        // ■■■NULLﾁｪｯｸ■■■
        // ■年度
        if( f_Trim(document.frm.txtKBN.value) == "" ){
            window.alert("区分の選択を行ってください");
            document.frm.txtKBN.focus();
            return ;
        }
        // ■年度
        if( f_Trim(document.frm.txtKBN.value) == "<%=C_CBO_NULL%>" ){
            window.alert("区分の選択を行ってください");
            document.frm.txtKBN.focus();
            return ;
        }
        // ■選択種別
<% if m_sOption <> "" then '選択できる選択種別がないときは、フォーカスしない %>
        if( f_Trim(document.frm.txtGRP.value) == "" ){
            window.alert("選択できる選択種別がありません。");
            return ;
        }
        // 
        if( f_Trim(document.frm.txtGRP.value) == "<%=C_CBO_NULL%>" ){
            window.alert("選択できる選択種別がありません。");
            return ;
        }
<% Else 	'選択してないとき%>
        // 
        if( f_Trim(document.frm.txtGRP.value) == "" ){
            window.alert("選択種別の選択を行ってください");
            document.frm.txtGRP.focus();
            return ;
        }
        // 
        if( f_Trim(document.frm.txtGRP.value) == "<%=C_CBO_NULL%>" ){
            window.alert("選択種別の選択を行ってください");
            document.frm.txtGRP.focus();
            return ;
        }
<% End If %>
		//学年、クラスをセット
		document.frm.txtGakunen.value = document.frm.cboGakunenCd.value
		document.frm.txtClass.value =document.frm.cboClassCd.value

        document.frm.action="web0340_main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }
    //************************************************************
    //  [機能]  クリアボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Clear(){

        document.frm.txtKBN.value = "@@@";
        document.frm.txtGRP.value = "@@@";
    
    }

    //-->
    </SCRIPT>

    <link rel="stylesheet" href="../../common/style.css" type="text/css">

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<center>

<form name="frm" METHOD="post">

<table cellspacing="0" cellpadding="0" border="0" width="100%">
    <tr>
        <td valign="top" align="center">
<%call gs_title("個人履修選択科目決定","一　覧")%>
<br>
            <table border="0">
                <tr>
                    <td class="search">
                        <table border="0" cellpadding="1" cellspacing="1">
                            <tr>
                                <td align="left">
                                    <table border="0" cellpadding="1" cellspacing="1">

                                        <tr>
                                            <td Nowrap align="left">学　年
											<% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' " &  m_sGakunenOption ,False,m_sGakunen) %>　クラス
											<!-- 2015.03.20 Upd width:80->180 -->
											<% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:180px;' " & m_sClassOption,False,m_sClass) %>
                                            </td>
                                        </tr>

                                        <tr>
                                            <td Nowrap align="left">区　分
											<%call gf_ComboSet("txtKBN",C_CBO_M01_KUBUN,m_sKBNWhere,"style='width:120px;' onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sKBN)%>
                                            </td>
                                            <td Nowrap align="left">　選択種別
											<%call gf_PluComboSet("txtGRP",C_CBO_T18_SEL_SYUBETU,m_sGRPWhere, "style='width:160px;' "& m_sOption,True,m_sGRP)%>
                                            </td>
                                        </tr>
										<tr>
											<td colspan="2" align="right">
									        <input type="button" class="button" value=" ク　リ　ア " onclick="javasript:f_Clear();">
											<input class="button" type="button" value="　表　示　" onClick = "javascript:f_Search()">
											</td>
										</tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>

<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
<input type="hidden" name="txtClass"    value="<%=m_sClass%>">
<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">

</form>

</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
