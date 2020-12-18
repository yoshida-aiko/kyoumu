<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/default.asp
' 機      能: フレームページ 学生情報検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg          'ｴﾗｰﾌﾗｸﾞ
    Public  m_PgMode           '処理モード
    Public  m_sMsgTitle        'ﾀｲﾄﾙ

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
    w_sMsgTitle="成績一覧表"
    w_sRetURL="../../login/default.asp"
    w_sTarget="_parent"

    m_PgMode=request("p_mode")
	Select Case m_PgMode
		Case "P_HAN0100"
		    m_sMsgTitle="成績一覧表"
		Case "P_KKS0200"
		    m_sMsgTitle="欠課一覧表"
		Case "P_KKS0210"
		    m_sMsgTitle="遅刻一覧表"
		Case "P_KKS0220"
		    m_sMsgTitle="行事欠課一覧表"
		Case "P_HAN0111"
		    m_sMsgTitle="評点一覧表"
		Case Else
	End Select
	w_sMsgTitle = m_sMsgTitle

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

		'// 権限チェックに使用
		'session("PRJ_No") = "GAK0300"

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

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
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
    '---------- HTML START  ----------
    On Error Resume Next
    Err.Clear

%>

<html>
	<head>
	<title><%=m_sMsgTitle%></title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<script language=javascript>
	</script>
	<frameset rows="65,*" border="1" framespacing="0" frameborder="no"> 
		<frame src="top.asp?p_mode=<%=m_PgMode%>" name="fTop" marginwidth="0" noresize scrolling="no" frameborder="no">
		<frame src="main.asp?p_mode=<%=m_PgMode%>&txtMode=Search" name="fMain" marginwidth="0" marginheight="0" scrolling="auto" frameborder="0">
	</frameset>
	</head>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>

