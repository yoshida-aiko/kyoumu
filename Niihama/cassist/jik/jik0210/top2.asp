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
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           cboGakunenCd      :学年コード
'           cboClassCd      :クラスコード
'           txtMode         :動作モード
'                               BLANK   :初期表示
' 説      明:
'           ■初期表示
'               コンボボックスは学年とクラスを表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう授業一覧を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/06 根本 直美
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '選択用のWhere条件
    Public m_sGakunenWhere        '学年の条件
    Public m_sClassWhere        'クラスの条件
    Public m_sClassOption          ':クラスコンボのオプション
    
    '取得したデータを持つ変数
    Public  m_iSyoriNen      ':処理年度
    Public  m_iKyokanCd      ':教官コード
    Public  m_iGakunen      ':学年コード
    
    'データ取得用

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
    w_sMsgTitle="クラス別授業時間一覧"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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
    'm_iSyoriNen = 2001      '//テスト用
    m_iKyokanCd = Session("KYOKAN_CD")

    m_iGakunen = ""
    m_iGakunen = Request("cboGakunenCd")


End Sub

'Sub s_MakeGakunenWhere()
'********************************************************************************
'*  [機能]  学年コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
'
'    
'    m_sGakuenWhere = ""
'    
'    m_sGakuenWhere = m_sGakuenWhere & " M05_NENDO = " & m_iSyorinen
'    'm_sGakuenWhere = m_sGakuenWhere & " M05_NENDO = " & 2000  '//テスト用
'End Sub

Sub s_MakeClassWhere()
'********************************************************************************
'*  [機能]  教官コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sClassWhere = ""
    m_sClassOption = ""
    
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen
    if m_iGakunen <> "" Then
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & m_iGakunen
    end if
    
    if m_iGakunen = "" Then
        m_sClassOption = " DISABLED "
    end if
    
End Sub

Sub s_SetGakCbo()
'********************************************************************************
'*  [機能]  学年コンボのSelectを表示させる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_iCount

    For w_iCount = 1 To 5
        response.write "<option value=" & w_iCount
            If CStr(m_iGakunen) = CStr(w_iCount) Then
                response.write " Selected "
            End If
        response.write " >" & w_iCount
    Next

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
<link rel=stylesheet href="../../common/style.css" type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
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

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();
    
    }



    //-->
    </SCRIPT>

</head>
<body>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
    <input type="hidden" name="txtMode">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
    <%call gs_title("クラス別授業時間一覧","一　覧")%>
<%
If m_sMode = "" Then
%>
    <table border="0">
    <tr>
    <td>
        <table border="0" class=search cellpadding="1" cellspacing="1">
        <tr>
        <td align="left" class=search>
            <table border="0" cellpadding="1" cellspacing="1">
            <tr>
            <td align="left" class=search>
            学年
            </td>
            <td align="left" class=search>
            <% 'call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS,m_sGakuenWhere,"onchange = 'javascript:f_ReLoadMyPage()' ",False,m_iGakunen) %>
            <select name="cboGakunenCd" onchange = 'javascript:f_ReLoadMyPage()'>
                <option>
                <%Call s_SetGakCbo()%>
            </select>
            年生
            </td>
            <td align="left" class=search>
            クラス
            <% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,m_sClassOption,False,"") %>
            </td>
            </tr>
            </table>
        </td>
        <td><input type="button" value="表示" onClick="javascript:f_Search()" class=button></td>
        </tr>
        </table>
    </td>
    </tr>
    </table>
<%
End IF
%>
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
