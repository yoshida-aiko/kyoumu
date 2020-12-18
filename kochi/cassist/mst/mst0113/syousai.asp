<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 中学校マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0113/syossai.asp
' 機      能: 下ページ 中学校マスタの詳細表示を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtKenCd        :県コード
'           txtSityoCd      :市町村コード
'           txtTyuName      :中学校名称（一部）
'           txtPageTyu      :表示済表示頁数（自分自身から受け取る引数）
'           txtTyuCd        :選択された中学校コード
'           txtTyuKbn       :選択された中学校区分
'           txtMode         :モード
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           txtKenCd        :県コード（戻るとき）
'           txtSityoCd      :市町村コード（戻るとき）
'           txtTyuName      :中学校名称（戻るとき）
'           txtPageTyu      :表示済表示頁数（戻るとき）
'           txtTyuKbn       :中学校区分（戻るとき）
'           txtMode         :モード
' 説      明:
'           ■初期表示
'               指定された中学校の詳細データを表示
'           ■地図画像ボタンクリック時
'               指定した条件にかなう中学校地図を表示する（別ウィンドウ）
'-------------------------------------------------------------------------
' 作      成: 2001/06/16 高丘 知央
' 変      更: 2001/07/26 根本　直美　'DB変更に伴う修正
'             2001/07/31 根本  直美  変数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_sKenCd        ':県コード
    Public  m_sSityoCd      ':市町村コード
    Public  m_sTyuName      ':中学校名称（一部）
    Public  m_iPageTyu      ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_sTyuCd        ':選択された中学校コード
    Public  m_Rs            'recordset
    Public  m_iNendo        ':年度      '//2001/07/31変更
    Public  m_sMode         ':モード
    Public  m_iTyuKbn       ':中学校区分
    
    Public  m_iTyuKbnD      ':中学校区分(DB)
    Public  m_iJyoKbnD      ':学校状況区分(DB)
    Public  m_sTyuKbnMei    ':中学校区分名
    Public  m_sJyoKbnMei    ':学校状況区分名
    

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
    Dim w_sWHERE            '// WHERE文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="中学校情報検索"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sMode = request("txtMode")

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

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '中学校マスタを取得
        w_sWHERE = ""

        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M13.M13_TYUGAKKO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKOMEI "     
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TEL "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_YUBIN_BANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_GAKKOJYOKYO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TIZUFILENAME "
        w_sSQL = w_sSQL & vbCrLf & " ,M12.M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M16.M16_KENMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM M13_TYUGAKKO M13 "
        w_sSQL = w_sSQL & vbCrLf & " , M16_KEN M16 "
        w_sSQL = w_sSQL & vbCrLf & " , M12_SITYOSON M12 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & "      M13.M13_KEN_CD = M16.M16_KEN_CD (+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_KEN_CD = M12.M12_KEN_CD (+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_SITYOSON_CD = M12.M12_SITYOSON_CD (+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13_TYUGAKKO_CD = '" & m_sTyuCd & "' "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_NENDO = " & m_iNendo & ""

'Response.Write w_sSQL & "<br>"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If
        
        '// DBから区分を取得
        Call s_SetDB()
        Call s_GetTyugakkoKbn()
        Call s_GetJyokyoKbn()
        
        
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
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo = Session("NENDO")     ':年度

    m_sKenCd = Request("txtKenCd")          ':県コード
    'コンボ未選択時
    If m_sKenCd="@@@" Then
        m_sKenCd=""
    End If

    m_sSityoCd = Request("txtSityoCd")      ':市町村コード
    'コンボ未選択時
    If m_sSityoCd="@@@" Then
        m_sSityoCd=""
    End If

    m_sTyuName = Request("txtTyuName")      ':中学校名称（一部）
    m_sTyuCd = Request("txtTyuCd")      ':中学校名称（一部）


    '// BLANKの場合は行数ｸﾘｱ
    'If Request("txtMode") = "Search" Then
    '    m_iPageTyu = 1
    'Else
        m_iPageTyu = INT(Request("txtPageTyu"))     ':表示済表示頁数（自分自身から受け取る引数）
    'End If

    m_iTyuKbn = Request("txtTyuKbn")        ':中学校区分

End Sub


'********************************************************************************
'*  [機能]  DB値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetDB()

m_iTyuKbnD = m_Rs("M13_TYUGAKKO_KBN")
m_iJyoKbnD = m_Rs("M13_GAKKOJYOKYO_KBN")

End Sub

'********************************************************************************
'*  [機能]  中学校区分名の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_GetTyugakkoKbn()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    m_sTyuKbnMei = ""
    
    On Error Resume Next
    Err.Clear

    Do
        
        '// 区分マスタﾚｺｰﾄﾞｾｯﾄを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & " M01_SYOBUNRUIMEI"
        w_sSQL = w_sSQL & " FROM M01_KUBUN "
        w_sSQL = w_sSQL & " WHERE M01_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD = " & C_TYUGAKKO_KBN
        w_sSQL = w_sSQL & " AND M01_SYOBUNRUI_CD = " & gf_SetNull2Zero(trim(m_iTyuKbnD))

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_sTyuKbnMei = "　"
            'm_sErrMsg = "ﾚｺｰﾄﾞｾｯﾄの取得失敗"
            Exit Do 
        End If
        
        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            m_sTyuKbnMei = "　"
            'm_sErrMsg = "対象ﾚｺｰﾄﾞなし"
            Exit Do 
        End If
        
        '// 取得した値を格納
        m_sTyuKbnMei = w_Rs("M01_SYOBUNRUIMEI")
        '// 正常終了
        Exit Do

    Loop

    gf_closeObject(w_Rs)

End Sub

'********************************************************************************
'*  [機能]  学校状況区分名の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_GetJyokyoKbn()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    m_sJyoKbnMei = ""
    
    On Error Resume Next
    Err.Clear

    Do
        
        '// 区分マスタﾚｺｰﾄﾞｾｯﾄを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & " M01_SYOBUNRUIMEI"
        w_sSQL = w_sSQL & " FROM M01_KUBUN "
        w_sSQL = w_sSQL & " WHERE M01_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD = " & C_GAKKO_JYOKYO
        w_sSQL = w_sSQL & " AND M01_SYOBUNRUI_CD = " & gf_SetNull2Zero(trim(m_iJyoKbnD))

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_sJyoKbnMei = "　"
            'm_sErrMsg = "ﾚｺｰﾄﾞｾｯﾄの取得失敗"
            Exit Do 
        End If
        
        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            m_sJyoKbnMei = "　"
            'm_sErrMsg = "対象ﾚｺｰﾄﾞなし"
            Exit Do 
        End If
        
        '// 取得した値を格納
        m_sJyoKbnMei = w_Rs("M01_SYOBUNRUIMEI")
        '// 正常終了
        Exit Do

    Loop

    gf_closeObject(w_Rs)

End Sub

''********************************************************************************
''*  [機能]  全項目に引き渡されてきた値を設定
''*  [引数]  なし
''*  [戻値]  なし
''*  [説明]  
''********************************************************************************
'Sub s_MapHTML()
'
'    If ISNULL(m_Rs("M13_TIZUFILENAME")) OR m_Rs("M13_TIZUFILENAME")="" Then
'        Response.Write("登録されていません")
'    Else
'        Response.Write("<a Href=""javascript:f_OpenWindow('" & Session("TYUGAKU_TIZU_PATH") & m_Rs("M13_TIZUFILENAME") & "')"">周辺地図</a>")
'    End If
'    
'End Sub
'


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

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageTyu.value = p_iPage;
        document.frm.submit();
    
    }

    function f_OpenWindow(p_Url){
    //************************************************************
    //  [機能]  子ウィンドウをオープンする
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
        var window_location;
        window_location=window.open(p_Url,"window","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,scrolling=no,Width=500,Height=500");
        window_location.focus();
    }

    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_GoSyosai(p_sTyuCd){

        document.frm.action="syousai.asp";
        document.frm.target="";
        document.frm.txtMode.value = "Syosai";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    </head>

<body>

<center>

<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">

<%call gs_title("中学校情報検索","詳　細")%>

<img src="../../image/sp.gif" height="10"><br>

    <table border="1" class=disp width="400">
        <tr>
            <td class=disph align="left" width="100">中学校名</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_TYUGAKKOMEI")) %></td>
        </tr>
        <!-- tr>
            <td class=disph align="left" width="100">中学略称</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_TYUGAKKORYAKSYO")) %></td>
        </tr -->
        <tr>
            <td class=disph align="left" width="100">区分</td>
            <td class=disp align="left" width="300"><%=m_sTyuKbnMei%></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">学校状況</td>
            <td class=disp align="left" width="300"><%=m_sJyoKbnMei%></td>
        </tr>

        <!-- tr>
            <td class=disph align="left" width="100">県</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M16_KENMEI")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">市区町村</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M12_SITYOSONMEI")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所（１）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_JUSYO1")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所（２）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_JUSYO2")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所（３）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_JUSYO3")) %></td>
        </tr -->

        <tr>
            <td class=disph align="left" width="100">郵便番号</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_YUBIN_BANGO")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所</td>
            <td class=disp align="left" width="300">
                <%=gf_HTMLTableSTR(m_Rs("M13_JUSYO1"))%><BR>
                <%=gf_HTMLTableSTR(m_Rs("M13_JUSYO2"))%>
                <%=gf_HTMLTableSTR(m_Rs("M13_JUSYO3"))%></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">電話番号</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_TEL")) %></td>
        </tr>
        <!--<tr>
            <td class=disph align="left" width="100">地図</td>
            <td class=disp align="left" width="300">
<%
    ' 地図の有無を表示・リンクするファンクション
    'Call s_MapHTML()
%>
            </td>
        </tr>-->
    </table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form action="./default.asp" target="<%=C_MAIN_FRAME%>">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtKenCd" value="<%=m_sKenCd%>">
        <input type="hidden" name="txtSityoCd" value="<%=m_sSityoCd%>">
        <input type="hidden" name="txtTyuName" value="<%=m_sTyuName%>">
        <input type="hidden" name="txtPageTyu" value="<%=m_iPageTyu%>">
        <input type="hidden" name="txtTyuCd" value="">
        <input type="hidden" name="txtTyuKbn" value="<%=m_iTyuKbn%>">
    <input class=button type="submit" value="戻　る">
    </form>
    </td>
    </tr>
    </table>

</td>
</tr>
</table>

    </center>


    </body>

    </html>





<%
    '---------- HTML END   ----------
End Sub
%>










