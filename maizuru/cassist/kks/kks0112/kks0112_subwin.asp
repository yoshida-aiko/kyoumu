<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠一覧
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0110/kks0111_detail.asp
' 機      能: フレームページ 授業出欠表示
'-------------------------------------------------------------------------
' 引      数:
' 変      数:
' 引      渡:
' 説      明:
'           
'-------------------------------------------------------------------------
' 作      成: 2002/05/07 shin
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="授業出欠入力"
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
			m_bErrFlg = True
			w_sMsg = "データベースとの接続に失敗しました。"
			'm_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If
		
		'// 権限チェックに使用
		session("PRJ_No") = "KKS0112"
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
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
<title>授業出欠入力</title>

<script language="javascript1.2">
<!--

window.onload=init;

function init(){
	if(!document.all){
		document.all={}
		if(!frames["topFrame"].document.body)frames["topFrame"].document.body={scrollLeft:0,scrollTop:0}
			frames["topFrame"].setInterval('if(parent.frames["main"].pageXOffset!=(document.body.scrollLeft=self.pageXOffset))document.body.onscroll()',10)
		if(!frames["main"].document.body)frames["main"].document.body={scrollLeft:0,scrollTop:0}
			frames["main"].setInterval('if(parent.frames["topFrame"].pageXOffset!=(document.body.scrollLeft=self.pageXOffset))document.body.onscroll()',10)
	}
	
	if(document.all){
		frames["topFrame"].document.body.onscroll=function(){
			frames["main"].scrollTo(frames["topFrame"].document.body.scrollLeft,frames["main"].document.body.scrollTop)
		}
		
		frames["main"].document.body.onscroll=function(){
			frames["topFrame"].scrollTo(frames["main"].document.body.scrollLeft,frames["topFrame"].document.body.scrollTop)
		}
	}
}

//-->
</script>

</head>
<frameset cols="200px,1,*" border="1" frameborder="no" onBlur="window.focus();">
	<frame src="kks0112_subwin_left.asp?<%=Request.QueryString%>" scrolling="no" name="leftFrame" border="1">
	<frame src="../../common/bar.html" scrolling="no" name="barH" noresize>
	
    <frameset rows="80px,1,*" border="1" frameborder="no" border="1" onload="init();">
		<frame src="kks0112_subwin_top.asp?<%=Request.QueryString%>" scrolling="yes" name="topFrame">
		<frame src="../../common/bar.html" scrolling="no" name="barW" noresize>
		<frame src="kks0112_subwin_bottom.asp?<%=Request.QueryString%>" scrolling="yes" name="main">
	</frameset>
    
</frameset>

</head>
</html>
<%
End Sub
%>