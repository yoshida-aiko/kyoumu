<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 中学校情報検索
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0113/default.asp
' 機      能: フレームページ 中学校マスタの参照を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           :txtKenCd       :都道府県コード     '/2001/07/30追加
'           :txtSityoCd     :市町村コード       '/2001/07/30追加
'           :txtTyuName     :中学校名           '/2001/07/30追加
'           :txtPageTyu     :表示済表示頁数     '/2001/07/30追加
'           :txtMode        :モード             '/2001/07/30追加
'           :txtTyuKbn      :中学校区分         '/2001/07/30追加
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/30追加
'           :txtKenCd       :都道府県コード     '/2001/07/30追加
'           :txtSityoCd     :市町村コード       '/2001/07/30追加
'           :txtTyuName     :中学校名           '/2001/07/30追加
'           :txtPageTyu     :表示済表示頁数     '/2001/07/30追加
'           :txtMode        :モード             '/2001/07/30追加
'           :txtTyuKbn      :中学校区分         '/2001/07/30追加
'           :txtQueryKenCd          :都道府県コード     '/2001/07/30追加
'           :txtQuerySityoCd        :市町村コード       '/2001/07/30追加
'           :txtQueryTyuName        :中学校名           '/2001/07/30追加
'           :txtQueryPageTyu        :表示済表示頁数     '/2001/07/30追加
'           :txtQueryTyuKbn         :中学校区分         '/2001/07/30追加
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/06/15 高丘 知央
' 変      更: 2001/07/27 根本　直美(DB変更に伴う修正)
'           : 2001/07/30 根本 直美  変数名命名規則に基く変更
'           :                       引数・引渡追加
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    Public  m_sKenCd            ':都道府県コード
    Public  m_sSityoCd          ':市町村コード
    'Public  m_sSinroName       '
    Public  m_sPageTyu          ':表示済表示頁数
    'Public  m_sSentakuSinroCD  '
    Public  m_sMode             ':モード

    Public  m_sArg              '引数     '/2001/07/30変更
    Public  m_sArg_top          '引数     '/2001/07/30変更

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
    w_sMsgTitle="中学校情報検索"
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
        session("PRJ_No") = "MST0113"

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

    m_sKenCd = Request("txtKenCd")      ':県コード
    m_sSityoCd = Request("txtSityoCd")  ':市町村コード
    m_sTyuName = Request("txtTyuName")  ':中学校名
    m_sPageTyu = Request("txtPageTyu")  ':表示済表示頁数
    m_sMode = Request("txtMode")        ':モード
    
    m_iTyuKbn = Request("txtTyuKbn")        ':中学校区分

    m_sArg = "?"
    m_sArg = m_sArg & "txtMode=" & m_sMode 
    m_sArg = m_sArg & "&txtKenCd=" & m_sKenCd 
    m_sArg = m_sArg & "&txtSityoCd=" & m_sSityoCD 
    m_sArg = m_sArg & "&txtTyuName=" & m_sTyuName 
    m_sArg = m_sArg & "&txtPageTyu=" & m_sPageTyu
    m_sArg = m_sArg & "&txtTyuKbn=" & m_iTyuKbn

    m_sArg_top = "?"
    m_sArg_top = m_sArg_top & "txtMode=" & m_sMode 
    m_sArg_top = m_sArg_top & "&txtQueryKenCd=" & m_sKenCd 
    m_sArg_top = m_sArg_top & "&txtQuerySityoCd=" & m_sSityoCD 
    m_sArg_top = m_sArg_top & "&txtQueryPageTyu=" & m_sPageTyu 
    m_sArg_top = m_sArg_top & "&txtQueryTyuName=" & m_sTyuName 
    m_sArg_top = m_sArg_top & "&txtQueryTyuKbn=" & m_iTyuKbn

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

<title>中学校情報検索</title>

<frameset rows=165,1,* frameborder="0">

<frame src="./top.asp<%= m_sArg_top %>" scrolling="auto" noresize name="top">
<frame src="../../common/bar.html" scrolling="auto" noresize name="bar">

<frame src="
<% If m_sMode = "" Then %>
    default2.asp
<% Else %>
    main.asp<%= m_sArg %>
<% End If %>
" scrolling="auto" noresize name="main">
</frameset>

</head>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>
