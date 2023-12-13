<%@ Language=VBScript %>
<%Response.Expires = 0%>
<%Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: パスワード変更
' ﾌﾟﾛｸﾞﾗﾑID : web/web0400/default.asp
' 機      能: ログインパスワードを変更します。
'-------------------------------------------------------------------------
' 引      数:SESSION(""):教官コード     ＞      SESSIONより
' 変      数:なし
' 引      渡:SESSION(""):教官コード     ＞      SESSIONより
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/10/04 谷脇
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ システム系
    Public m_bErrFlgPara           'ｴﾗｰﾌﾗｸﾞ パラメータチェック系
    Public m_iNendo
    Public m_sUser
    Public m_sPass
    Public m_sPassN1
    Public m_sPassN2

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()
response.end
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
    Dim w_iLevel

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="パスワード変更"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
	m_bErrFlgPara = False
    Do

		'// 変数初期化
		call f_paraSet()

			'// ﾃﾞｰﾀﾍﾞｰｽ接続
			w_iRet = gf_OpenDatabase()
			If w_iRet <> 0 Then
				'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
				m_bErrFlg = True
				m_sErrMsg = "データベースとの接続に失敗しました。"
				Exit Do
			End If

			'// 権限チェックに使用
	'		session("PRJ_No") = "WEB0400"

			'// 不正アクセスチェック
			'Call gf_userChk(session("PRJ_No"))
			
			'// 不正アクセスチェック
			w_iRet = f_login(m_sUser,m_sPass,w_iLevel)
			If w_iRet = false Then
				'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
				If m_bErrFlgPara = true then 
					w_sMsg = "ログインIDとパスワードが一致しませんでした。<BR>もう一度、ログインID、パスワードを確認の上、変更ボタンを押してください。"
				else
					m_bErrFlg = True
					m_sErrMsg = "ログインデータの取得ができませんでした。"
				End If
				Exit Do
			End If

			'// 権限チェック
			w_iRet = f_TT51(w_iLevel)
			If w_iRet = false Then
				'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
				If m_bErrFlgPara = true then 
					w_sMsg = "パスワードを変更する権限がありません。"
				else
					m_bErrFlg = True
					m_sErrMsg = "ログインデータの取得ができませんでした。"
				End If
				Exit Do
			End If

			'// 更新処理
			w_iRet = f_Update()
			If w_iRet = false Then
				'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
					m_bErrFlg = True
					m_sErrMsg = "データの更新ができませんでした。"
				Exit Do
			End If

        '// 変更ページを表示
        Call showPage()
'        Call showErrPage("大成功")
        Exit Do
    Loop

    '// パラメータのｴﾗｰの場合はパラメータｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlgPara = True Then
        Call showErrPage(w_sMsg)
    End If
    
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Sub

Sub f_paraSet()
'*******************************************************************************
' 機　　能：変数の初期化と代入
' 引　　数：なし
' 機能詳細：
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
m_iNendo = session("NENDO")
m_sMode = Request("txtMode")
m_sUser = Request("txtUser")
m_sPass = Request("txtPass")
m_sPassN1 = Request("txtPassN1")
m_sPassN2 = Request("txtPassN2")
'm_iNendo = 2001
	
End Sub


Function f_login(p_id,p_pass,p_level)
'********************************************************************************
'*  [機能]  ログイン処理
'*  [引数]  p_id   = ﾕｰｻﾞｰが入力したﾕｰｻﾞｰID
'*          p_pass = ﾕｰｻﾞｰが入力したﾊﾟｽﾜｰﾄﾞ
'*  [戻値]  p_level= ﾕｰｻﾞｰの権限
'*  [説明]  
'********************************************************************************
    Dim w_sSql
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_login = false

  Do
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
    w_sSql = w_sSql & "     M10_NENDO    =  " & m_iNendo & " AND "
    w_sSql = w_sSql & "     M10_USER_ID  = '" & p_id & "' AND "
    w_sSql = w_sSql & "     M10_PASSWORD = '" & p_pass  & "' "

    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit do 'GOTO LABEL_MAIN_END
    End If

	'// ﾚｺｰﾄﾞｾｯﾄがなかったら抜ける
	If m_Rs.Eof then
        m_bErrFlgPara = True
        Exit do 'GOTO LABEL_MAIN_END
	End if

	'// 権限取得
	p_level = m_Rs("M10_LEVEL")

    f_login = true
    exit do
  Loop

    call gf_closeObject(m_Rs)

End Function

Function f_TT51(p_level)
'********************************************************************************
'*  [機能]  権限チェック
'*  [引数]  p_level = ﾕｰｻﾞｰの権限
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    Dim w_sSql
    Dim w_sLevel
    
    On Error Resume Next
    Err.Clear
    m_bErrFlg = False

    f_TT51 = false

  Do
    '// Nullなら抜ける
    if trim(p_level) = "" then
        Exit Function
    End if

    w_sLevel = "T51_LEVEL" & trim(p_level)

    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & w_sLevel
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & "     TT51_SYORI_LEVEL  "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & "     T51_ID         = 'WEB0400' AND "
    w_sSql = w_sSql & "     T51_SYORI_KBN  = 12 "

    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit do
    End If

    '// ﾚｺｰﾄﾞｾｯﾄがなかったら抜ける
    If m_Rs.Eof then
        m_bErrFlgPara = True
        Exit do
    End if

    '// 1：OK
    If m_Rs(w_sLevel) <> "1" then
        m_bErrFlgPara = True
        Exit do
    End If

    f_TT51 = true
    exit do
  Loop

    call gf_closeObject(m_Rs)

End Function

Function f_Update()
'********************************************************************************
'*  [機能]  パスワード変更
'*  [引数]  なし
'*  [戻値]  true:成功 false:失敗
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    
    f_Update = false

    Do 

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

            '//T11_GAKUSEKIにUPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE M10_USER SET "
            w_sSQL = w_sSQL & vbCrLf & "   M10_PASSWORD = '"  & Trim(m_sPassN1) & "' ,"
            w_sSQL = w_sSQL & vbCrLf & "   M10_UPD_DATE = '"  & gf_YYYY_MM_DD(date(),"/") & "', "
            w_sSQL = w_sSQL & vbCrLf & "   M10_UPD_USER = '"  & Session("LOGIN_ID") & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        M10_USER_ID = '" & Trim(m_sUser) & "' AND "
            w_sSQL = w_sSQL & vbCrLf & "        M10_NENDO = " & m_iNendo & " "

            iRet = gf_ExecuteSQL(w_sSQL)
            If iRet <> 0 Then
                '//ﾛｰﾙﾊﾞｯｸ
                Call gs_RollbackTrans()
                msMsg = Err.description
                f_Update = 99
                Exit Do
            End If

        '//ｺﾐｯﾄ
        Call gs_CommitTrans()

        '//正常終了
        f_Update = true
        Exit Do
    Loop

End Function

Sub showErrPage(p_msg)
'********************************************************************************
'*  [機能]  エラーHTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
<html>
<head>
    <title>パスワード変更</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
</head>
<body>
<center>
    <%call gs_title("ログインパスワード変更","更　新")%>

<BR>
	<table class="hyo" border="0" width="70%">
		<FORM action="default.asp" name="frm" method="post" target="_self">
			<tr><td colspan="2" height="30" class="detail"></td></tr>
			<tr><td colspan="2" height="15" class="detail" align="center"><%=p_msg%></td></tr>
			<tr><td colspan="2" height="30" class="detail"></td></tr>
			<tr><td colspan="2" height="30" class="detail" align="center">
				<input type="submit" name="submit" value=" 戻 る " maxlength="16"> <!-- 2023.10.25 Upd Kiyomoto パスワードを10桁⇒16桁に変更 -->
	            </td>
	        </tr>
<input type="hidden" name="txtUser" value="<%=m_sUser%>">
<input type="hidden" name="txtPass" value="<%=m_sPass%>">
<input type="hidden" name="txtPassN1" value="<%=m_sPassN1%>">
<input type="hidden" name="txtPassN2" value="<%=m_sPassN2%>">
		</FORM>
	</table>
</center>
</body>
</head>
</html>
<%
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  エラーHTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
<html>
<head>
    <title>パスワード変更</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
		function f_load(){
			alert("パスワードを変更しました。");
			top.location.href="../../default.asp";
		}
	//-->
	</SCRIPT>
</head>
<body onload="f_load();">
<center>
    <%call gs_title("ログインパスワード変更","更　新")%>
</center>
</body>
</head>
</html>
<%
End Sub
%>
