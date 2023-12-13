<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: クラス別授業時間一覧
' ﾌﾟﾛｸﾞﾗﾑID : jik/jik0210/top.asp
' 機      能: 上ページ クラス別授業時間の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           :session("PRJ_No")      '権限ﾁｪｯｸのキー
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           cboGakunenCd      :学年コード
'           cboClassCd      :クラスコード
'           txtMode         :動作モード
'                           (BLANK) :初期表示
'                           Reload  :リロード
' 説      明:
'           ■初期表示
'               コンボボックスは学年とクラスを表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件の授業一覧を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/06 根本 直美
' 変      更: 2001/07/30 根本 直美  戻り先URL変更
'           : 2001/08/09 根本 直美     NN対応に伴うソース変更
'           : 2015/03/20 清本 千秋  Win7対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '選択用のWhere条件
    Public m_sGakunenWhere      '学年の条件
    Public m_sClassWhere        'クラスの条件
    Public m_sGakunenOption     ':学年コンボのオプション
    Public m_sClassOption       ':クラスコンボのオプション
    
    '取得したデータを持つ変数
    Public  m_iSyoriNen         ':処理年度
    Public  m_iKyokanCd         ':教官コード
    Public  m_iGakunen          ':学年コード
    Public  m_sMode             ':動作モード
    
    'データ取得用
    Public  m_iTanninG          ':担任（学年）
    Public  m_iTanninC          ':担任（クラス）

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
    w_sWinTitle= "キャンパスアシスト"
    w_sMsgTitle= "クラス別授業時間一覧"
    w_sMsg= ""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget= ""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False


        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

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

        '担任学年・クラス取得
        Call f_GetTannin() 
        '// 学年コード取得
        Call s_SetGakunenCd()
        
        '学年コンボに関するWHEREを作成する
        Call s_MakeGakunenWhere() 
        'クラスコンボに関するWHEREを作成する
        Call s_MakeClassWhere() 

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

Sub s_SetParam()
'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

    m_sMode = Request("txtMode")

End Sub

Sub s_SetGakunenCd()
'********************************************************************************
'*  [機能]  学年コードを設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    if m_sMode = "" Then
        m_iGakunen = m_iTanninG
    else
        m_iGakunen = Request("cboGakunenCd")
    end if
    
End Sub

Sub s_MakeGakunenWhere()
'********************************************************************************
'*  [機能]  学年コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    
    m_sGakunenWhere = ""
    m_sGakunenOption = ""
    
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyorinen
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"
    
    if m_sMode = "" Then
        m_sGakunenOption = m_iTanninG
    else
        m_sGakunenOption = m_iGakunen
    end if
    
End Sub

Sub s_MakeClassWhere()
'********************************************************************************
'*  [機能]  クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sClassWhere = ""
    m_sClassOption = ""
    
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen
    if m_sMode = "" Then
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & m_iTanninG
    else
        if m_iGakunen <> "" Then
            m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & m_iGakunen
        end if
    end if

    if m_sMode = "" Then
        m_sClassOption = m_iTanninC
    end if
    
End Sub

'********************************************************************************
'*  [機能]  担任学年・クラスの取得
'*  [引数]  なし
'*  [戻値]  0:情報取得成功、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetTannin()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    f_GetTannin = 0
    m_iTanninG = ""
    m_iTanninC = ""

    Do

        '// 学年・クラスマスタを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "M05_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "M05_CLASSNO "
        w_sSQL = w_sSQL & vbCrLf & "FROM "
        w_sSQL = w_sSQL & vbCrLf & "M05_CLASS "
        w_sSQL = w_sSQL & vbCrLf & "WHERE "
        'w_sSQL = w_sSQL & vbCrLf & "M05_NENDO = 2200"
        w_sSQL = w_sSQL & vbCrLf & "M05_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND M05_TANNIN = '" & m_iKyokanCd & "'"
        
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
        
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            f_GetSentaku = 99
            Exit Do 'GOTO LABEL_f_GetTannin_END
        Else
        End If
        
        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            'm_bErrFlg = True
            'm_sErrMsg = "対象ﾚｺｰﾄﾞなし"
            'f_GetTannin = 1
            m_iTanninG = 1
            m_iTanninC = 1

            Exit Do 'GOTO LABEL_f_GetTannin_END
        End If

            '// 取得した値を格納
            m_iTanninG = w_Rs("M05_GAKUNEN")
            m_iTanninC = w_Rs("M05_CLASSNO")
        '// 正常終了
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs)

'// LABEL_f_GetTannin_END
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
<link rel="stylesheet" href="../../common/style.css" type="text/css">
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

    }

    //************************************************************
    //  [機能]  戻るボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_BackClick(){

    }

    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        document.frm.action="main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }
    //************************************************************
    //  [機能]  試験が選択されたとき、再表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="top.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();
    
    }



    //-->
    </SCRIPT>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" action="./main.asp" target="main" Method="POST">

    <input type="hidden" name="txtMode">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
    <%call gs_title("クラス別授業時間一覧","一　覧")%>
<br>
    <table border="0">
    <tr>
    <td class="search">
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left">
            <table border="0" cellpadding="1" cellspacing="1">
            <tr valign="middle">
            <td align="left">
            クラス
            </td>
            <td align="left">
            <% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage();' style='width:40px;' ",False,m_sGakunenOption) %>
            </td>
            <td align="left">
            年
            </td>
            <td><img src="../../image/sp.gif" height="10"></td>
            <td align="left">
			<!-- 2015.03.20 Upd width:80->180 -->
            <% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"style='width:180px;' ",False,m_sClassOption) %>
            </td>
		        <td colspan="6" align="right">　
		        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
		        </td>
			</tr>
            </table>
        </td>
        </tr>
        </table>
    </td>
    </tr>
    </table>
</td>
</tr>
</table>

</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
