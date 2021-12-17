<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: ログイン画面
' ﾌﾟﾛｸﾞﾗﾑID : login/default.asp
' 機      能: ログイン画面
'-------------------------------------------------------------------------
' 引      数    Request("txtLogin")   = ﾛｸﾞｲﾝID
'               Request("txtPass")    = ﾊﾟｽﾜｰﾄﾞ
'               Request("hidLoginFlg) = ﾛｸﾞｲﾝ画面から来たしるし
'           
' 変      数
' 引      渡    Session("NENDO")            '年度
'               Session("LOGIN_ID")         'ログインＣＤ
'               Session("USER_NM")          'ユーザネーム
'               Session("LEVEL")            '権限
'               Session("USER_KBN")         'ユーザ区分
'               Session("KYOKAN_CD")        '教官CD

' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 
' 変      更: 2001/07/26    モチナガ
'*************************************************************************/
%>
<!--#include file="../common/com_All.asp"-->
<%

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_Rs                'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_iMiKengenFlg		'権限ﾁｪｯｸをしないﾌﾗｸﾞ


'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////


'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sRetURL           '// ｴﾗｰﾒｯｾｰｼﾞ用戻り先URL
    Dim w_sTarget           '// ｴﾗｰﾒｯｾｰｼﾞ用戻り先ﾌﾚｰﾑ
    Dim w_sWinTitle         '// ｴﾗｰﾒｯｾｰｼﾞ用ﾀｲﾄﾙ
    Dim w_sMsgTitle         '// ｴﾗｰﾒｯｾｰｼﾞ用ﾀｲﾄﾙ
    
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="学籍データ検索結果"
    w_sMsg=""
    w_sRetURL="../default.asp"
    w_sTarget="_parent"

     Do

		'// ログイン画面から来た場合は、ログインチェックをする
		if Cint(Request("hidLoginFlg")) = Cint(C_LOGIN_FLG) then

	        '// ﾊﾟﾗﾒｰﾀ取得
	        w_LoginID  = Request("txtLogin")
	        w_PassWord = Request("txtPass")

			'// ﾃﾞｰﾀﾍﾞｰｽ接続
			w_iRet = gf_OpenDatabase()
			If w_iRet <> 0 Then
				'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
				m_bErrFlg = True
				m_sErrMsg = "データベースとの接続に失敗しました。"
				Exit Do
			End If

			'// 年度取得
			If Not f_GetNendo() then Exit Do End if

			'// ﾛｸﾞｲﾝﾁｪｯｸ
			If Not f_login(w_LoginID,w_PassWord) Then
				Call ErrPage()          '// エラーページを表示
			End if

		End if

		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		Call showPage()         '// ページを表示

        '// 正常終了
        Exit Do
    LOOP

   '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, m_sErrMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  年度を出力
'*  [引数]  
'*          
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetNendo()
    Dim w_sSql
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_GetNendo = False

    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & "     A.M00_KANRI "
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & "     M00_KANRI A "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & "     A.M00_NENDO  = " & C_M00_NENDO & " AND "
    w_sSql = w_sSql & "     A.M00_NO     = 0 "

    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_NendoRs, w_sSQL)

    If w_iRet <> 0 Then
    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

    '// 各種情報をセッションに格納
    session("NENDO") = w_NendoRs("M00_KANRI")           '年度

    f_GetNendo = True

    call gf_closeObject(w_NendoRs)

End Function

'********************************************************************************
'*  [機能]  ログイン処理
'*  [引数]  p_id   = ﾕｰｻﾞｰが入力したﾕｰｻﾞｰID
'*          p_pass = ﾕｰｻﾞｰが入力したﾊﾟｽﾜｰﾄﾞ
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_login(p_id,p_pass)
    Dim w_sSql
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_login = false

    '// Nullなら抜ける
    if trim(p_id) = "" then
        Exit Function
    Elseif trim(p_pass) = "" then
        Exit Function
    End if

    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & "     M10_USER_ID, "      '0
    w_sSql = w_sSql & "     M10_KYOKAN_CD, "    '1
    w_sSql = w_sSql & "     M10_USER_NAME, "    '2
    w_sSql = w_sSql & "     M10_USER_KBN, "     '3
    w_sSql = w_sSql & "     M10_LEVEL "         '4
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & "     M10_USER  "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & "     M10_NENDO    =  " & session("NENDO") & " AND "
    w_sSql = w_sSql & "     M10_USER_ID  = '" & p_id & "' AND "
    w_sSql = w_sSql & "     M10_PASSWORD = '" & p_pass  & "' "

    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

	'// ﾚｺｰﾄﾞｾｯﾄがなかったら抜ける
	If m_Rs.Eof then
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
	End if

    '// １時限の単位数情報取得
    w_iRet = f_GetJigen_Tani(w_iTani)
    If w_iRet <> 0 Then
        m_bErrFlg = True
        Exit Function
    End If

    '// 各種情報をセッションに格納
    Session("LOGIN_ID")  = m_RS("M10_USER_ID")          'ログインＣＤ
    Session("USER_NM")   = m_Rs("M10_USER_NAME")        'ユーザネーム
    Session("LEVEL")     = m_Rs("M10_LEVEL")            '権限
    Session("USER_KBN")  = m_Rs("M10_USER_KBN")         'ユーザ区分
    Session("KYOKAN_CD") = m_Rs("M10_KYOKAN_CD")        '教官CD
	
	application("KYOKAN_CD") = m_Rs("M10_KYOKAN_CD")        '教官CD
	
	Session("JIKAN_TANI") = cint(w_iTani)				'１時限の単位(時間)数
    f_login = true


    call gf_closeObject(m_Rs)

End Function

'********************************************************************************
'*  [機能]  １時限の単位数を取得
'*  [引数]  なし
'*  [戻値]  p_iTani : １時限の時間数を取得(基本は、１対１)
'*  [説明]  
'********************************************************************************
Function f_GetJigen_Tani(p_iTani)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetJigen_Tani = 1
	p_iTani = ""

    Do 

		'//時限マスタより単位数を取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(M07_TANISU) AS TANISU "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M07_JIGEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M07_NENDO=" & session("NENDO")

'response.write w_sSQL  & "<BR>"

         iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSikenKbn = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_iTani = cint(w_Rs("TANISU"))
		End If

		'//データが取得できないときは、１にする。
		If p_iTani = "" Then
			p_iTani = 1
		End If

        f_GetJigen_Tani = 0

        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  ログインエラー表示ページ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub ErrPage()
%>
 <HTML>
   <HEAD>
    <meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
    <TITLE></TITLE>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_loginErr(){
        alert("ログインエラー\nログインIDとパスワードをお確かめの上、\n再度ログインしてください。");
        location.href="<%=C_RetURL%>default.asp"
        return true;
    }

            </SCRIPT>
        </HEAD>
        <BODY onLoad="f_loginErr();">
            <br>
        </BODY>
    </HTML>
<%
    End Sub

'********************************************************************************
'*  [機能]  HTML表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
%>
<html>

<head>
<title>Campus Assist</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>

<frameset rows=61,*,24 frameborder="0" FRAMESPACING="0" BORDER="0">
    <frame src="header.asp" scrolling="no" NAME="topHead">
    <frameset cols=166,* frameborder="0" FRAMESPACING=0 frameborder="0">
        <frame src="menu.asp" scrolling="auto" noresize name="menu">
        <frame src="top.asp" scrolling="auto" noresize name="<%=C_MAIN_FRAME%>">
    </frameset>
        <frame src="foot.asp" scrolling="auto" noresize name="foot">
</frameset>

</html>
<%
End Sub
%>