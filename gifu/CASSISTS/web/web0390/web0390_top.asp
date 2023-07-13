<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: レベル別科目決定
' ﾌﾟﾛｸﾞﾗﾑID : web/web0390/web0390_top.asp
' 機      能: 上ページ レベル別科目決定の検索を行う
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
' 作      成: 2001/10/26 谷脇　良也
' 変      更: 
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
    w_sMsgTitle="レベル別科目決定"
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
		w_iRet = gf_GetKengen_web0390(m_sKengen)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'権限の定数
		'C_WEB0390_ACCESS_FULL  
		'C_WEB0390_ACCESS_SENMON
		'C_WEB0390_ACCESS_TANNIN

		'//データを変数にセット
		Call s_SetParam()

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
	If m_sKengen = C_WEB0390_ACCESS_TANNIN Then

		'//担任教官の場合は、担任クラスの年組を取得する
		Call f_Gakunen()
	Else
		'//担任以外の場合
	    m_sGakunen  = Request("cboGakunenCd")   '//学年
		if m_sGakunen = "@@@" OR m_sGakunen = "" then m_sGakunen = "1"
	    m_sClass    = Request("cboClassCd")     '//クラス
		if m_sClass = "@@@" OR m_sClass = "" then m_sClass = "1"

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

	If m_sKengen = C_WEB0390_ACCESS_TANNIN Then
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
	If m_sKengen = C_WEB0390_ACCESS_TANNIN Then
		m_sClassOption = "DISABLED"
	End If

End Sub

'********************************************************************************
'*  [機能]  レベル別科目コンボ生成
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_KamokuCbo(p_sKamokuCBO,p_sOption)

    Dim w_sSQL,w_Rs,w_iRet
    Dim w_NyuNen 		'入学年度
	Dim WEB0391_Flg

    On Error Resume Next
    Err.Clear

    f_Get_KamokuCbo = 1
    p_sKamokuCBO = ""
    p_sOption = ""
	WEB0391_Flg = false

	m_iNendo = cint(gf_SetNull2Zero(m_iNendo))
	m_sGakunen = cint(gf_SetNull2Zero(m_sGakunen))
	m_sClass = cint(gf_SetNull2Zero(m_sClass))

	If (m_iNendo = 0 OR m_sGakunen = 0 OR m_sClass = 0) then 
		p_sKamokuCBO = "<option value=''>科目がありません</option>" & vbCrLf
		p_sOption = " DISABLED"
	    f_Get_KamokuCbo = 0
		exit Function
	End If
        '================
        '//科目情報取得
        '================
	w_NyuNen = m_iNendo - cInt(m_sGakunen) + 1

        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_KAMOKU_CD, "
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_KAMOKUMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_LEVEL_FLG"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS M05,T15_RISYU T15"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  M05.M05_GAKKA_CD = T15.T15_GAKKA_CD AND "
        w_sSQL = w_sSQL & vbCrLf & "  M05.M05_NENDO="      & cInt(m_iNendo) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  M05.M05_GAKUNEN="    & cInt(m_sGakunen)  & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  M05.M05_CLASSNO="    & cInt(m_sClass)  & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_KAISETU" 	 & cInt(m_sGakunen)  & " < 3 AND"
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_LEVEL_FLG = 1 AND"
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_NYUNENDO= "  & w_NyuNen

'response.write w_sSQL & "<br>"
'response.end
	'レコードセットの取得失敗
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_KamokuCbo = 99
            Exit Function
        End If

	'科目が取れなかった場合
	If w_Rs.EOF Then 
		p_sKamokuCBO = "<option value=''>科目がありません</option>" & vbCrLf
		p_sOption = " DISABLED"
	    f_Get_KamokuCbo = 0
		exit Function
	End If

	'データがあれば、コンボボックスを生成
	Do Until w_Rs.EOF 
		If m_sKengen = C_WEB0390_ACCESS_SENMON then 
			If f_KyokanData(w_Rs("T15_KAMOKU_CD")) = true then 
				p_sKamokuCBO = p_sKamokuCBO & "<option value='" & w_Rs("T15_KAMOKU_CD") & "'>"
				p_sKamokuCBO = p_sKamokuCBO & w_Rs("T15_KAMOKUMEI")
				p_sKamokuCBO = p_sKamokuCBO & "</option>" & vbCrLf
				WEB0391_Flg = true
			End If
		Else
				p_sKamokuCBO = p_sKamokuCBO & "<option value='" & w_Rs("T15_KAMOKU_CD") & "'>"
				p_sKamokuCBO = p_sKamokuCBO & w_Rs("T15_KAMOKUMEI")
				p_sKamokuCBO = p_sKamokuCBO & "</option>" & vbCrLf
		End If
		w_Rs.MoveNext
	Loop

	If m_sKengen = C_WEB0390_ACCESS_SENMON then 
			If WEB0391_Flg = false then 
				p_sKamokuCBO = "<option value=''>科目がありません</option>" & vbCrLf
				p_sOption = " DISABLED"
			    f_Get_KamokuCbo = 0
				exit Function
			End If
	End If

    '正常終了
    f_Get_KamokuCbo = 0

End Function

Function f_KyokanData(p_sKamokuCD)
'******************************************************************
'機　　能：教官のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
	Dim w_rs,w_iRet,w_sSQL

    On Error Resume Next
    Err.Clear
    f_KyokanData = false

    Do


        '//科目のデータ取得
        w_sSQL = ""
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     T27_KYOKAN_CD"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "     T27_TANTO_KYOKAN"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "     T27_NENDO = " & m_iNendo & " "
        w_sSQL = w_sSQL & vbCrLf & " AND T27_GAKUNEN = " & m_sGakunen & " "
        w_sSQL = w_sSQL & vbCrLf & " AND T27_KAMOKU_CD = '" & p_sKamokuCD & "' "
        w_sSQL = w_sSQL & vbCrLf & " AND T27_KYOKAN_CD = '" & m_sKyokanCd & "' "

'response.write w_sSQL & vbCrLf & "<BR>"

        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, w_sSQL,m_iDsp)

'response.write "w_iRet = " & w_iRet & "<BR>"
'response.write w_rs.EOF & "<BR>"& vbCrLf
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            Exit Do 
        End If

        If w_rs.EOF Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            Exit Do 
        End If

    f_KyokanData = true

    Exit Do

    Loop

   Call gf_closeObject(w_rs)

End Function

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

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
	Dim w_sKamokuCBO	'レベル別科目コンボ
	Dim w_sOption		'レベル別科目コンボオプション
	Dim w_iRet
	w_iRet = f_Get_KamokuCbo(w_sKamokuCBO,w_sOption)
    If w_iRet <> 0 Then m_bErrFlg = True : Exit Sub

%>
<html>

<head>

<title>レベル別科目決定</title>
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

        document.frm.action="web0390_top.asp";
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
        // ■
        if( f_Trim(document.frm.cboKamokuCode.value) == "" ){
            window.alert("科目が選択されていません。");
//            document.frm.cboKamokuCode.focus();
            return ;
        }

		//学年、クラスをセット
		document.frm.txtGakunen.value = document.frm.cboGakunenCd.value
		document.frm.txtClass.value =document.frm.cboClassCd.value

        document.frm.action="web0390_main.asp";
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

		document.frm.cboGakunenCd.value = ""
		document.frm.cboClassCd.value = ""

		f_ReLoadMyPage();
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
<%call gs_title("レベル別科目決定","一　覧")%>
<br>
            <table border="0">
                <tr>
                    <td class="search">
                        <table border="0" cellpadding="1" cellspacing="1">
                            <tr>
                                <td align="left">
                                    <table border="0" cellpadding="1" cellspacing="1">

                                        <tr>
                                            <td Nowrap align="left">クラス
											<% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' " &  m_sGakunenOption ,false,m_sGakunen) %>　
											<% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere," style='width:80px;' " & m_sClassOption,false,m_sClass) %>
                                            </td>
                                            <td Nowrap align="left">科　目
											<Select name="cboKamokuCode" style='width:200px;'<%=w_sOption%>>
												<%=w_sKamokuCBO%>
 											</Select>
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
