<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索詳細
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/kojin_ue.asp
' 機      能: 検索された学生の詳細を表示する
'-------------------------------------------------------------------------
' 引      数 	Session("GAKUSEI_NO")  = 学生番号
'            	Session("HyoujiNendo") = 表示年度
' 変      数
' 引      渡
'           
'           
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public m_bErrFlg		'ｴﾗｰﾌﾗｸﾞ
	Public m_Rs				'ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ

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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="学生情報検索結果"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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

		'//表示項目を取得
		w_iRet = f_GetDetail()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//初期表示
        if m_TxtMode = "" then
            Call showPage()
            Exit Do
        end if

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  表示項目を取得
'*  [引数]  なし
'*  [戻値]  0:正常終了	1:任意のエラー  99:システムエラー
'*  [説明]  
'********************************************************************************
Function f_GetDetail()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetail = 1

	Do
		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO, "
		w_sSql = w_sSql & " 	A.T13_GAKUSEKI_NO,  "
		w_sSql = w_sSql & " 	A.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	A.T13_CLASS,  "
		w_sSql = w_sSql & " 	B.T11_SIMEI "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A, "
		w_sSql = w_sSql & " 	T11_GAKUSEKI B "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO(+) AND "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' AND "
		w_sSql = w_sSql & " 	A.T13_NENDO		 =  " & Session("HyoujiNendo")

		iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetDetail = 99
			Exit Do
		End If

		'//正常終了
		f_GetDetail = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()


	m_GAKUSEI_NO  = ""
	m_GAKUSEKI_NO = ""
	m_GAKUNEN     = ""
	m_CLASS       = ""
	m_SIMEI       = ""

	if Not m_Rs.Eof then
		m_GAKUSEI_NO  = m_Rs("T13_GAKUSEI_NO")
		m_GAKUSEKI_NO = m_Rs("T13_GAKUSEKI_NO")
		m_GAKUNEN     = m_Rs("T13_GAKUNEN")
		m_CLASS       = m_Rs("T13_CLASS")
		m_SIMEI       = m_Rs("T11_SIMEI")
	End if

%>
	<html>
	<head>
	<title>学籍データ参照</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<%call gs_title("学生情報検索","詳細")%>

	<br>
	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td>

				<table border="1" class="disp">
					<tr>
						<% if gf_empItem(C_T13_GAKUSEI_NO) then %>
							<td class="disph" nowrap width="100" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_5NEN)%></td>
							<td class="disp" nowrap width="80"><%= m_GAKUSEI_NO %>&nbsp;</td>
						<% End if %>
						<% if gf_empItem(C_T13_GAKUSEKI_NO) then %>
							<td class="disph" nowrap width="100" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_1NEN)%></td>
							<td class="disp" nowrap width="80"><%= m_GAKUSEKI_NO %>&nbsp;</td>
						<% End if %>
					</tr>
				</table>

			</td>
		</tr>
		<tr>
			<td>

				<table border="1" class="disp">
					<tr>
						<% if gf_empItem(C_T13_GAKUNEN) then %>
							<td class="disph" nowrap width="100" height="16">学　　年</td>
							<td class="disp" nowrap width="80"><%= m_GAKUNEN %>&nbsp;</td>
						<% End if %>
						<% if gf_empItem(C_T13_CLASS) then %>
							<td class="disph" nowrap width="100" height="16">ク ラ ス</td>
							<td class="disp" nowrap width="80"><%= m_CLASS %>組&nbsp;</td>
						<% End if %>
						<% if gf_empItem(C_T11_SIMEI) then %>
							<td class="disph" nowrap width="100" height="16">氏　　名</td>
							<td class="disp" nowrap><%= m_SIMEI %>&nbsp;</td>
						<% End if %>
					</tr>
				</table>

			</td>
		</tr>
	</table>

	</div>
	</form>
	</body>
	</html>
<% End Sub %>