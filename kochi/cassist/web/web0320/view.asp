<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 使用教科書登録
' ﾌﾟﾛｸﾞﾗﾑID : web/web0320/default.asp
' 機      能: 使用教科書の詳細確認
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/09/04 伊藤　公子
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系

    Public m_bErrFlg        'ｴﾗｰﾌﾗｸﾞ
    Public m_iNendo         '年度
    Public m_sKyokan_CD     '教官CD
    Public m_sTitle         ''新規登録・修正の表示用

    ''ﾃﾞｰﾀ表示用
    Public m_sNo
    Public m_sNendo
    Public m_sGakkiCD
    Public m_sGakunenCD
    Public m_sGakkaCD
    Public m_sKamokuCD
    Public m_sCourseCD
    Public m_sKyokan_NAME       '教官
    Public m_sKyokasyo_NAME     '教科書
    Public m_sSyuppansya        '出版社
    Public m_sTyosya            '著者名
    Public m_sSidousyo          '指導書
    Public m_sKyokanyo          '教官用
    Public m_sBiko              '備考

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

    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="就職先マスタ登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

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

        '// 値を変数に入れる
        Call s_SetParam()

		'//表示データ取得
		if f_GetData() = False then
			exit do
		end if

        '// 教官の名称を取得する
        if f_GetData_Kyokan() = False then
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
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  値を変数に入れる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo     = Session("NENDO")
    m_sMode      = Request("txtMode")       ':モード
	m_sTitle = "参照"
	m_sNo = Request("txtUpdNo")     ''更新用No格納

	'//一覧表示中ページを保存
    m_sPageCD    = Request("txtPageCD")

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iNendo		= " & m_iNendo			& "<br>"
	response.write "m_sMode			= " & m_sMode			& "<br>"
	response.write "m_sTitle		= " & m_sTitle			& "<br>"
	response.write "m_sDBMode		= " & m_sDBMode			& "<br>"
	response.write "m_sPageCD		= " & m_sPageCD			& "<br>"
	response.write "m_sNendo		= " & m_sNendo			& "<br>"
	response.write "m_sNo			= " & m_sNo				& "<br>"
	response.write "m_sKyokan_CD	= " & m_sKyokan_CD		& "<br>"
	response.write "m_sGakkiCD		= " & m_sGakkiCD		& "<br>"
	response.write "m_sGakunenCD	= " & m_sGakunenCD		& "<br>"
	response.write "m_sGakkaCD		= " & m_sGakkaCD		& "<br>"
	response.write "m_sKamokuCD		= " & m_sKamokuCD		& "<br>"
	response.write "m_sCourseCD		= " & m_sCourseCD		& "<br>"
	response.write "m_sKyokan_NAME	= " & m_sKyokan_NAME	& "<br>"
	response.write "m_sKyokasyo_NAME= " & m_sKyokasyo_NAME	& "<br>"
	response.write "m_sSyuppansya	= " & m_sSyuppansya		& "<br>"
	response.write "m_sTyosya		= " & m_sTyosya			& "<br>"
	response.write "m_sSidousyo		= " & m_sSidousyo		& "<br>"
	response.write "m_sKyokanyo		= " & m_sKyokanyo		& "<br>"
	response.write "m_sBiko			= " & m_sBiko			& "<br>"

End Sub

'********************************************************************************
'*  [機能]  教官の名称を取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
function f_GetData_Kyokan()
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    dim w_Rs

    f_GetData_Kyokan = False

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " M04.M04_NENDO "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKAN_CD "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN M04 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    M04_NENDO = " &  m_iNendo & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN_CD = '" & m_sKyokan_CD & "'"

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    Else
        'ページ数の取得
        m_iMax = gf_PageCount(w_Rs,m_iDsp)
    End If

	m_sKyokan_NAME = ""
	If w_Rs.EOF = False Then
	    m_sKyokan_NAME = w_Rs("M04_KYOKANMEI_SEI") & "  " & w_Rs("M04_KYOKANMEI_MEI")
	End If

    w_Rs.close

    f_GetData_Kyokan = True

end function

'********************************************************************************
'*  [機能]  更新時の表示ﾃﾞｰﾀを取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
function f_GetData()
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_Rs

    f_GetData = False

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " T47.T47_NENDO "            ''年度
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKI_KBN "       ''学期区分
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKUNEN "         ''学年
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKA_CD "        ''学科
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_COURSE_CD "       ''ｺｰｽｺｰﾄﾞ
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KAMOKU "          ''科目ｺｰﾄﾞ
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKASYO "        ''教科書名
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SYUPPANSYA "      ''出版社
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_TYOSYA "          ''著者
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKANYOUSU "     ''教官用数
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SIDOSYOSU "       ''指導書数
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_BIKOU "           ''備考
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKAN "           ''教官
    w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKAMEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    T47_KYOKASYO T47 "
    w_sSQL = w_sSQL & vbCrLf & "    ,M02_GAKKA M02 "
    w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
    w_sSQL = w_sSQL & vbCrLf & "    ,M04_KYOKAN M04 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M02.M02_NENDO(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_GAKKA_CD  = M02.M02_GAKKA_CD(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M03.M03_NENDO(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M04.M04_NENDO(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = M04.M04_KYOKAN_CD(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO = " & Request("KeyNendo") & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NO = " & m_sNo & ""

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    Else
        'ページ数の取得
        m_iMax = gf_PageCount(w_Rs,m_iDsp)
    End If

    m_sNendo   = gf_HTMLTableSTR(w_Rs("T47_NENDO"))
    m_sGakkiCD   = gf_HTMLTableSTR(w_Rs("T47_GAKKI_KBN"))
    m_sGakunenCD = gf_HTMLTableSTR(w_Rs("T47_GAKUNEN"))
    m_sGakkaCD   = gf_HTMLTableSTR(w_Rs("T47_GAKKA_CD"))
    m_sKamokuCD  = gf_HTMLTableSTR(w_Rs("T47_KAMOKU"))
    m_sCourseCD  = gf_HTMLTableSTR(w_Rs("T47_COURSE_CD"))
    m_sKyokasyo_NAME  = gf_HTMLTableSTR(w_Rs("T47_KYOKASYO"))       '教科書
    m_sSyuppansya  = gf_HTMLTableSTR(w_Rs("T47_SYUPPANSYA"))        '出版社
    m_sTyosya  = gf_HTMLTableSTR(w_Rs("T47_TYOSYA"))                '著者名
    m_sSidousyo  = gf_HTMLTableSTR(w_Rs("T47_SIDOSYOSU"))           '指導書
    m_sKyokanyo  = gf_HTMLTableSTR(w_Rs("T47_KYOKANYOUSU"))         '教官用
    m_sBiko  = gf_HTMLTableSTR(w_Rs("T47_BIKOU"))                   '備考

    m_sKyokan_CD = gf_HTMLTableSTR(w_Rs("T47_KYOKAN"))
    w_Rs.close
    f_GetData = True

end function

'********************************************************************************
'*  [機能]  学科の略称を取得
'*  [引数]  p_sGakkaCd : 学科CD
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetGakkaNm(p_sGakkaCd)
    Dim w_sSQL              '// SQL文
    Dim w_iRet              '// 戻り値
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakkaNm = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKAMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_sNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_GAKKA_CD='" & p_sGakkaCd & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M02_GAKKAMEI")
		End If 

		Exit do 
	Loop

	'//戻り値をセット
	f_GetGakkaNm = w_sName

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'****************************************************
'[機能] コース名称を取得
'[引数] pData1 : データ１
'[戻値] f_Selected : "SELECTED" OR ""
'****************************************************
Function f_GetCourseNm()
    Dim w_sSQL              '// SQL文
    Dim w_iRet              '// 戻り値
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetCourseNm = ""
	w_sName = ""

	Do

		If Trim(m_sCourseCD) = "" Then
			Exit Do
		End If

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M20_COURSEMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M20_COURSE "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M20_NENDO         =  " & m_sNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M20_GAKKA_CD  = '" & m_sGakkaCD & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND M20_GAKUNEN   =  " & m_sGakunenCD
		w_sSQL = w_sSQL & vbCrLf & "  AND M20_COURSE_CD = '" & m_sCourseCD & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M20_COURSEMEI")
		End If 

		Exit do 
	Loop

	'//戻り値をセット
	f_GetCourseNm = w_sName

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*  [機能]  科目名称を取得
'*  [引数]  p_sGakkaCd : 学科CD
'*          p_sKamokuCd
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKamokuNm(p_sGakkaCd,p_sKamokuCd)
    Dim w_sSQL              '// SQL文
    Dim w_iRet              '// 戻り値
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
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & m_sNendo

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

<title>使用教科書登録</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  メインページへ戻る
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Back(){
		document.frm.action="./default.asp";
        document.frm.target="";
        document.frm.submit();
    
    }

    //-->
    </script>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">

	</head>
	<body>
	<form name="frm" action="" target="" method="post">

<%'call s_DebugPrint%>

	<center>
	<% call gs_title("使用教科書登録",m_sTitle) %>
	<br>
<table border="0" cellpadding="1" cellspacing="1" width="400">
    <tr>
        <td align="left">
            <table width="100%" border=1 CLASS="hyo">
                <tr>
                <th height="16" width="75" class=header nowrap>年　度</th>
                <td height="16" width="325" class=detail nowrap>
					<%=Request("KeyNendo")%><br>
                </td>
                </tr>

                <tr>
                <th height="16" width="75" class="header" nowrap>学　期</th>
				<%Call gf_GetKubunName(C_KAISETUKI,m_sGakkiCD,m_sNendo,w_KubunName)%>
                <td height="16" width="325" class="detail" nowrap><%=w_KubunName%><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>学　年</th>
                <td height="16" width="325" class=detail nowrap><%=m_sGakunenCD%>年</td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>学　科</th>
                <td height="16" width="325" class=detail nowrap><%=f_GetGakkaNm(m_sGakkaCD)%><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>コース</font></th>
                <td height="16" width="325" class=detail><%=f_GetCourseNm()%><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>科　目</font></th>
                <td height="16" width="325" class=detail><%=f_GetKamokuNm(m_sGakkaCD,m_sKamokuCD)%><br></td>
                </tr>

                <tr>

                <th height="16" width="80" class=header nowrap>教官</font></th>
                <td height="16" width="325" class=detail nowrap><%=m_sKyokan_NAME%><br></td>
                </tr>

                <tr>
                <th height="16" width="80" class=header nowrap>教科書名</font></th>
                <td height="16" width="325" class=detail nowrap><%= m_sKyokasyo_NAME %><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>出版社</font></th>
                <td height="16" width="325"  class=detail nowrap><%= m_sSyuppansya %><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>著者名</font></th>
                <td height="16" width="325" class=detail nowrap><%= m_sTyosya %><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>教官用</font>
                </th>
                <td height="16" width="325" class=detail nowrap><%= m_sKyokanyo %>冊</td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>指導書</font>
                </th>
                <td height="16" width="325" class=detail nowrap><%= m_sSidousyo %>冊</td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>備　考</font></th>
                <td height="35" width="325" class=detail nowrap valign="top"><%= trim(m_sBiko) %><br></td>
                </TR>
            </TABLE>
        </td>
    </TR>
</TABLE>
		<table border="0" width=300>
		<tr>
		<td valign="top" align=center>
		<input type="Button" class=button value="キャンセル" OnClick="f_Back()">
		</td>
		</tr>
		</table>
		</center>
		
	    <input type="hidden" name="txtNendo"     value="<%= Request("txtNendo") %>">
	    <input type="hidden" name="txtGakunenCd" value="<%= Request("txtGakunenCd") %>">
	    <input type="hidden" name="txtGakkaCD"   value="<%= Request("txtGakkaCD") %>">
	    <input type="hidden" name="txtPageCD"    value="<%= Request("txtPageCD") %>">
		
		<input type="hidden" name="txtMode" value="<%=Request("txtMode")%>">
		
		<input type="hidden" name="hidYear" value="<%=request("hidYear")%>">
		<input type="hidden" name="hidGakunen" value="<%=request("hidGakunen")%>">
		<input type="hidden" name="hidGakka" value="<%=request("hidGakka")%>">
	</form>
	</body>
	</html>

<%
End Sub
%>