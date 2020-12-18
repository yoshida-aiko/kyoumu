<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 高等学校情報検索
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0123/default.asp
' 機      能: フレームページ 高等学校マスタの参照を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           :txtKenCd       :都道府県コード     '/2001/07/31追加
'           :txtSityoCd     :市町村コード       '/2001/07/31追加
'           :txtSyuName     :高等学校名           '/2001/07/31追加
'           :txtPageSyu     :表示済表示頁数     '/2001/07/31追加
'           :txtMode        :モード             '/2001/07/31追加
'           :txtTyuKbn      :中学校区分         '/2001/07/31追加
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           :txtKenCd       :都道府県コード     '/2001/07/31追加
'           :txtSityoCd     :市町村コード       '/2001/07/31追加
'           :txtSyuName     :高等学校名           '/2001/07/31追加
'           :txtPageSyu     :表示済表示頁数     '/2001/07/31追加
'           :txtMode        :モード             '/2001/07/31追加
'           :txtSyuKbn      :高等学校区分         '/2001/07/31追加
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/06/20 岩下　幸一郎
' 変      更: 2001/07/27 根本　直美(DB変更に伴う修正)
'           : 2001/07/31 根本 直美  変数名命名規則に基く変更
'           :                       引数・引渡追加
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    Public  m_sKenCd            '都道府県コード
    Public  m_sSityoCd          '市町村コード
    'Public  m_sSinroName       
    Public  m_sPageSyu          '表示済表示頁数
    'Public  m_sSentakuSinroCD  
    Public  m_sMode             'モード
    Public  m_sSyuName          '高等学校M名

    Public  m_iSyuKbn       ':高等学校区分

    Public  m_sArg          ':引数'/2001/07/31変更
    Public  m_sArg_top      ':引数'/2001/07/31変更

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
    w_sMsgTitle="高等学校情報検索"
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
        session("PRJ_No") = "MST0123"

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

    m_sKenCd = Request("txtKenCd")      ':進路コード
    m_sSityoCd = Request("txtSityoCd")  ':進学コード
    m_sSyuName = Request("txtSyuName")  ':高等学校名
    m_sPageSyu = Request("txtPageSyu")  ':表示済表示頁数
    m_sMode = Request("txtMode")        ':モード
    
    m_iSyuKbn = Request("txtSyuKbn")        ':高等学校区分

    m_sArg = "?"
    m_sArg = m_sArg & "txtMode=" & m_sMode 
    m_sArg = m_sArg & "&txtKenCd=" & m_sKenCd 
    m_sArg = m_sArg & "&txtSityoCd=" & m_sSityoCD 
    m_sArg = m_sArg & "&txtSyuName=" & m_sSyuName 
    m_sArg = m_sArg & "&txtPageSyu=" & m_sPageSyu
    m_sArg = m_sArg & "&txtSyuKbn=" & m_iSyuKbn

    m_sArg_top = "?"
    m_sArg_top = m_sArg_top & "txtMode=" & m_sMode 
    m_sArg_top = m_sArg_top & "&txtKenCd=" & m_sKenCd 
    m_sArg_top = m_sArg_top & "&txtSityoCd=" & m_sSityoCD 
    m_sArg_top = m_sArg_top & "&txtPageSyu=" & m_sPageSyu 
    m_sArg_top = m_sArg_top & "&txtSyuName=" & m_sSyuName 
    m_sArg_top = m_sArg_top & "&txtSyuKbn=" & m_iSyuKbn

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

<title>高等学校情報検索</title>

<frameset rows=180,1,* frameborder="0">

<frame src="./top.asp<%= m_sArg_top %>" scrolling="auto" noresize name="top">
<frame src="../../common/bar.html" scrolling="auto" noresize name="bar">

<frame src="
<%If m_sMode = "" Then%>
    default2.asp
<%Else%>
    main.asp<%= m_sArg %>
<%End If%>
" scrolling="auto" noresize name="main">
</frameset>


</head>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>