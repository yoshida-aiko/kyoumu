
<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 
' ﾌﾟﾛｸﾞﾗﾑID : default.asp
' 機      能: ログインID・パスワードの照合を行う
'-------------------------------------------------------------------------
' 引      数:ログインID、パスワード
' 変      数:なし
' 引      渡:
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/06/15 高丘 知央
' 変      更: 2001/06/15 岩下 幸一郎
'*************************************************************************/
%>
<!--#include file="Common/com_All.asp"-->
<%


'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bErrMsg           'ｴﾗｰﾒｯｾｰｼﾞ
	Public  m_SchoolName		'学校名
	
'///////////////////////////メイン処理/////////////////////////////


    'バージョンの表示
    'Response.Write "[ ORACLE Ver:" & OraSession.OIPVersionNumber & " ]"
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
	w_sMsgTitle="トップ"
	w_sMsg=""
	w_sRetURL= C_RetURL     
	w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do

        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        If gf_OpenDatabase() <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_bErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

	'// セッションクリアー
	Call s_SessionClear

        '//パラメータセット
        If SetPara() = false Then
        	m_bErrFlg = True
        	Exit Do
        End If
        
        
        '//学校名を取得
        If not f_GetSchoolName() Then
        	m_bErrFlg = True
        	Exit Do
        End If
        
	    '// ページを表示
'Call ErrPage("テストだ")
	    Call showPage()
	    Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
Call ErrPage(w_sMsg)
'		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
    
    '// 終了処理
    Call gs_CloseDatabase()
End Sub

Sub s_SessionClear()
'********************************************************************************
'*  [機能]  セッションをクリアーする
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
'/** Dim Item
Dim w_OraDatabase 
Dim w_Qurs

	'// 必要なセッションは変数に入れる
	w_User_ID =	Session("USER_ID")
	w_PASS    =	Session("PASS")
	w_CONNECT =	Session("CONNECT")
	w_TYUGAKU_TIZU_PATH = Session("TYUGAKU_TIZU_PATH")

    SET w_OraDatabase = Session("OraDatabase")
    SET w_Qurs = Session("qurs")

	'セッションクリアー
	for Each name in Session.Contents
'/** Response.Write name & " " & session(name) & "***"
                session(name) = ""
	next

	'// セッションに戻す
	Session("USER_ID") = w_User_ID
	Session("PASS")    = w_PASS
	Session("CONNECT") = w_CONNECT
	Session("TYUGAKU_TIZU_PATH") = w_TYUGAKU_TIZU_PATH

    SET Session("OraDatabase") = w_OraDatabase
    SET Session("qurs") = w_Qurs

End Sub

Function SetPara() 
'********************************************************************************
'*  [機能]  変数をセット
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
	Dim w_sSQL,w_Rs,w_iRet
	Dim w_nendo
	
	SetPara = false

  Do
	w_sSQL = ""
	w_sSQL = w_sSQL & "SELECT "
	w_sSQL = w_sSQL & "M00_KANRI AS NENDO "
	w_sSQL = w_sSQL & "FROM "
	w_sSQL = w_sSQL & "M00_KANRI "
	w_sSQL = w_sSQL & "WHERE "
	w_sSQL = w_sSQL & "M00_NENDO = 9999 AND "
	w_sSQL = w_sSQL & "M00_NO = 0 AND "
	w_sSQL = w_sSQL & "M00_SYUBETU = 0 "

	Set w_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

	If w_iRet <> 0 Then
	'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		m_bErrFlg = True
		Exit Do 'GOTO LABEL_MAIN_END
	End If

	'//　処理年度を入れる
	Session("NENDO") = w_Rs("NENDO")
	Exit Do
  Loop	

	w_Rs = ""

  Do
	w_sSQL = ""
	w_sSQL = w_sSQL & "SELECT "
	w_sSQL = w_sSQL & "M00_KANRI AS GAKKI "
	w_sSQL = w_sSQL & "FROM "
	w_sSQL = w_sSQL & "M00_KANRI "
	w_sSQL = w_sSQL & "WHERE "
	w_sSQL = w_sSQL & "M00_NENDO = " & Session("NENDO") & " AND "
	w_sSQL = w_sSQL & "M00_NO = 11 AND "
	w_sSQL = w_sSQL & "M00_SYUBETU = 0 "

	Set w_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

	If w_iRet <> 0 Then
	'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		m_bErrFlg = True
		Exit Do 'GOTO LABEL_MAIN_END
	End If

	'//　学期をセッションに入れる。
	If w_Rs("GAKKI") > gf_YYYY_MM_DD(date(),"/") Then 
		Session("GAKKI") = C_GAKKI_ZENKI
	Else 
		Session("GAKKI") = C_GAKKI_KOUKI
	End If
	SetPara = true
	Exit Do
  Loop	

	'// ﾌﾞﾗｳｻﾞｰ情報取得
	wBrauza = request.servervariables("HTTP_USER_AGENT")
	if InStr(wBrauza,"IE") <> 0 then
		session("browser") = "IE"
	Else
		session("browser") = "NN"
	End if

	call gf_closeObject(m_Rs)

End Function

'********************************************************************************
'*  [機能]  学校名を取得する。
'*  [引数]  
'*  [説明]  
'********************************************************************************
Function f_GetSchoolName()
	
	Dim w_sSQL
	Dim w_Rs
	Dim w_FieldName
	Dim w_Table,w_TableName,w_KamokuName
	
	On Error Resume Next
	Err.Clear
	
	f_GetSchoolName = false
	m_SchoolName = ""
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M19_NAME "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M19_GAKKO "
	
	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function
	
	if not w_Rs.EOF then
		m_SchoolName = w_Rs("M19_NAME") & "　専攻科"
		Call gf_closeObject(w_Rs)
	end if
	
	f_GetSchoolName = true
	
End Function

Sub ErrPage(p_sMsg)
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
<html>
<head>
<title>キャンパスアシスト</title>
</head>
<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" bgcolor="#F6F7FC" >
<center>
<table width="100%" height="100%" cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td width="100%" height="40%" background="image/back.gif" style="background-repeat: repeat-y "><img src="image/title.gif" width="504" height="214"><br><br></td>
	</tr>
	<tr>
		<td align="center" background="image/back.gif" style="background-repeat: repeat-y;">
		<%=p_sMsg%><BR><BR>
		<font color="#FF0000">
		データベース接続に失敗しました<BR>
		管理者に連絡してください。
		</font>
		</td>
	</tr>
</table>
</center>
</body>
</html>
<%
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear

	'// フォームサイズ
	if session("browser") = "NN" then
		wformSize = "15"
	Else
		wformSize = "20"
	End if

%>
<html>

<head>
<title>Campus Assist</title>
<!-- <link rel=stylesheet href="common/style.css" type=text/css> -->
<LINK REL="SHORTCUT ICON" href="image/CAtitle.ico">
<script language="javascript">
<!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		document.frm.txtLogin.focus();
    }

    //************************************************************
    //  [機能]  リセットボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //  [作成日] 
    //************************************************************
	function f_clear() {
		document.frm.reset();
		return false;
	}
//-->
</script>
<style type="text/css">
<!--
   input { font-size:12px;}
   A {	 text-decoration:none; 
   		font-size:9pt;
   		text-align:center;
   	 }

   a:link {color:#222268;}
   a:visited {color:#222268;}
   a:active {color:#222268;}
   a:hover {color:#682222; text-decoration:underline; }
//-->
</style>
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" bgcolor="#ffffff" onLoad="window_onload();">

<table width="100%" height="100%" cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td nowrap width="25%" valign="top" rowspan="3"><%= "[ ORACLE Ver:" & OraSession.OIPVersionNumber & " ]" %></td>
		<td width="504" height="40%" background="image/back.gif"><img src="image/title.gif" width="504" height="214"><br><br></td>
		<td nowrap width="25%" rowspan="3">&nbsp;</td>
	</tr>
	
	<tr>
		<td height="50%" width="504" align="center" background="image/back.gif">
			
			<table cellspacing="0" cellpadding="0" width="244" height="140" border="0">
				<tr><td align="center" colspan="3"><font size="-1" color="#222268"><%=m_SchoolName%></font></td></tr>
				<tr><td colspan="3">&nbsp;</td></tr>
				
				<tr>
					<td height="5" width="5"><img src="image/table1.gif"></td>
					<td height="5" width="230" background="image/table2.gif"><img src="image/sp.gif"></td>
					<td height="5" width="9"><img src="image/table3.gif"></td>
				</tr>
				
				<tr>
					<td height="139" width="5" background="image/table4.gif"><img src="image/sp.gif"></td>
					<td height="139" width="230" bgcolor="#ffffff" align="center" background="image/sp.gif">
						
						<img src="image/sp.gif" height="1"><br>
						<form action="login/default.asp" name="frm" method="post">
						<table border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td><img src="image/login.gif" border="0"></td>
								<td><input type="text" size="<%=wformSize%>" name="txtLogin" value="<%= DC_USERADMIN %>"></td>
							</tr>
							<tr>
								<td colspan="2"><img src="image/sp.gif" height="5"></td>
							</tr>
							<tr>
								<td><img src="image/pass.gif" border="0"></td>
								<td><input type="password" size="<%=wformSize%>" name="txtPass" value="<%= DC_USERADMIN %>"></td>
							</tr>
							<tr>
								<td colspan="2"><img src="image/sp.gif" height="10"></td>
							</tr>
							<tr>
								<td colspan="2" align="center" valign="bottom"><input type="image" border="0" src="image/login_b.gif"><img src="image/sp.gif" width="35"><input type="image" border="0" src="image/clear.gif" onclick="return f_clear()"></td>
							</tr>
						</table>
		<% if gf_empPasChg() then %>
				<a href="web/web0400/default.asp">- パスワード変更はこちら -</a>
		<% End if %>
					</td>
					<td height="139" width="9" background="image/table5.gif"><img src="image/sp.gif"></td>
				</tr>
				<tr>
					<td height="5" width="5"><img src="image/table6.gif"></td>
					<td height="5" width="230" background="image/table7.gif"><img src="image/sp.gif"></td>
					<td height="5" width="9"><img src="image/table8.gif"></td>
				</tr>
				<tr><td colspan="3">&nbsp;</td></tr>
			</table>
			
		</td>
	</tr>
	<tr>
		<td height="10%" width="504" valign="bottom" align="center" background="image/back.gif"><img src="image/info_logo.gif"></td>
	</tr>
</table>


<input type="hidden" name="hidLoginFlg" value="<%= C_LOGIN_FLG %>">

</form>
</body>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>
