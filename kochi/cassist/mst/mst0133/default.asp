<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0133/default.asp
' 機      能: フレームページ 就職先マスタの参照を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           :txtSinroCD             :進路コード     '/2001/07/31追加
'           :txtSingakuCD           :進学コード     '/2001/07/31追加
'           :txtSyusyokuName        :就職先名称（一部） '/2001/07/31追加
'           :txtPageCD              :表示頁数           '/2001/07/31追加
'           :txtMode                :モード             '/2001/07/31追加
'           :txtSentakuSinroCD
'           :txtFLG
'           :txtSNm
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/06/18 岩下　幸一郎
' 変      更: 2001/07/31 根本 直美  引数・引渡追加
'           :                       変数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_sMode             'モード
    Public  m_iSinroCD          '進路区分           '/2001/07/31変更
    Public  m_iSingakuCd        '進学区分           '/2001/07/31変更
    Public  m_sSyusyokuName     '就職先名称（一部）
    Public  m_sPageCD           '表示頁数
    Public  m_sSentakuSinroCD   
    Public  m_iFLG
    Public  m_sSNm

    Public  m_sArg              '引数   '/2001/07/31変更
    Public  m_sArg_top          '引数   '/2001/07/31変更

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
    w_sMsgTitle="進路先情報検索"
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

        '// 権限チェックに使用
        session("PRJ_No") = "MST0133"

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

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
'*  [機能]  任意のページへパラメータを渡す
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSinroCD = Request("txtSinroCD")      ':進路区分
    'コンボ未選択時
    If m_iSinroCD="@@@" Then
        m_iSinroCD=""
    End If

    m_iSingakuCd = Request("txtSingakuCD")      ':進学コード
    'コンボ未選択時
    If m_iSingakuCd="@@@" Then
        m_iSingakuCd=""
    End If

    m_sSyusyokuName = Request("txtSyusyokuName")':就職先名称（一部）

    m_sPageCD = Request("txtPageCD")        ':表示されるページ

    m_sMode = Request("txtMode")            ':モード

    m_iFLG = request("txtFLG")
    m_sSNm = request("txtSNm")

    m_sArg = "?"
    m_sArg = m_sArg & "txtMode=" & m_sMode 
    m_sArg = m_sArg & "&txtSinroCD=" & m_iSinroCD 
    m_sArg = m_sArg & "&txtSingakuCD=" & m_iSingakuCd 
    m_sArg = m_sArg & "&txtSyusyokuName=" & Server.URLEncode(m_sSyusyokuName) 
    m_sArg = m_sArg & "&txtPageCD=" & m_sPageCD 
    m_sArg = m_sArg & "&txtSentakuSinroCD=" & m_sSentakuSinroCD 

    m_sArg_top = "?"
    m_sArg_top = m_sArg_top & "txtSinroCD=" & m_iSinroCD 
    m_sArg_top = m_sArg_top & "&txtSingakuCD=" & m_iSingakuCd 
    m_sArg_top = m_sArg_top & "&txtSyusyokuName=" & Server.URLEncode(m_sSyusyokuName) 
    m_sArg_top = m_sArg_top & "&txtFLG=" & m_iFLG 
    m_sArg_top = m_sArg_top & "&txtSNm=" & Server.URLEncode(m_sSNm) 

End Sub


Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	If m_sMode = "" Then
	    w_FrameUrl = "default2.asp"
	Else
	    w_FrameUrl = "main.asp" & m_sArg
	End If

%>
<html>
<head>
<title>進路先情報検索</title>
</head>

<frameset rows=190,1,* frameborder="0">
	<frame src="top.asp<%=m_sArg_top%>" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
	<frame src="<%=w_FrameUrl%>" scrolling="auto" noresize name="main">
</frameset>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>
