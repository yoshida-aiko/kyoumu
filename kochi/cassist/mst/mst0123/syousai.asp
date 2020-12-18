<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 高等学校マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0123/syossai.asp
' 機      能: 下ページ 高等学校マスタの詳細表示を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtKenCd        :県コード
'           txtSityoCd      :市町村コード
'           txtSyuName      :高等学校名称（一部）
'           txtPageSyu      :表示済表示頁数（自分自身から受け取る引数）
'           txtSyuCd        :選択された高等学校コード
'           txtSyuKbn       :高等学校区分
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           txtKenCd        :県コード（戻るとき）
'           txtSityoCd      :市町村コード（戻るとき）
'           txtSyuName      :高等学校名称（戻るとき）
'           txtPageSyu      :表示済表示頁数（戻るとき）
'           txtSyuKbn       :高等学校区分（戻るとき）
' 説      明:
'           ■初期表示
'               指定された高等学校の詳細データを表示
'           ■地図画像ボタンクリック時
'               指定した条件にかなう高等学校地図を表示する（別ウィンドウ）
'-------------------------------------------------------------------------
' 作      成: 2001/06/20 岩下　幸一郎
' 変      更: 2001/07/26 根本　直美　'DB変更に伴う修正
'           : 2001/07/31 根本 直美  変数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_sKubunCd      ':県コード
    Public  m_sKenCd        ':県コード
    Public  m_sSityoCd      ':市町村コード
    Public  m_sSyuName      ':高等学校名称（一部）
    Public  m_iPageSyu      ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_sSyuCd        ':選択された中学校コード
    Public  m_Rs            'recordset
    Public  m_iNendo        ':年度      '/2001/07/31変更
    Public  m_sMode         ':モード
    
    Public  m_iSyuKbn       ':高等学校区分


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

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '高等学校マスタを取得
        w_sWHERE = ""
        
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M31.M31_GAKKOMEI "        
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_GAKKORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_TEL "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_YUBIN_BANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_TIZUFILENAME "
        w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M12.M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M16.M16_KENMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM M31_SYUSSINKO M31 "
        w_sSQL = w_sSQL & vbCrLf & " , M16_KEN M16 "
        w_sSQL = w_sSQL & vbCrLf & " , M12_SITYOSON M12 "
        w_sSQL = w_sSQL & vbCrLf & " , M01_KUBUN M01 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & "      M01.M01_DAIBUNRUI_CD = " & C_SYUSSINKO
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_KEN_CD = M16.M16_KEN_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_KEN_CD = M12.M12_KEN_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_SITYOSON_CD = M12.M12_SITYOSON_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M31_GAKKO_CD = '" & m_sSyuCd & "' "
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_NENDO = " & m_iNendo & ""
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_NENDO = M01.M01_NENDO(+) " 
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_GAKKO_KBN = M01.M01_SYOBUNRUI_CD(+) "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If
        
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

    m_sKubunCd = Request("txtKubunCd")      ':区分コード
    'コンボ未選択時
    If m_sKubunCd="@@@" Then
        m_sKubunCd=""
    End If

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

    m_sSyuName = Request("txtSyuName")      ':高等学校名称（一部）
    m_sSyuCd = Request("txtSyuCd")          ':高等学校コード

    m_iNendo = Session("NENDO")             ':年度      '/2001/07/31変更
    m_sMode = request("txtMode")            ':モード

    '// BLANKの場合は行数ｸﾘｱ
    If Request("txtMode") = "Search" Then
        m_iPageSyu = 1
    Else
        m_iPageSyu = INT(Request("txtPageSyu"))     ':表示済表示頁数（自分自身から受け取る引数）
    End If
    
    m_iSyuKbn = Request("txtSyuKbn")        ':高等学校区分

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
        document.frm.txtPageSyu.value = p_iPage;
        document.frm.submit();
    
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

<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
<td valign="top" align="center">

<%call gs_title("高等学校情報検索","詳　細")%>

    <table border="1" class=disp width="400">
        <tr>
            <td class=disph align="left" width="100">高等学校名</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_GAKKOMEI")) %></td>
        </tr>
        <!-- tr>
            <td class=disph align="left" width="100">高等学校略称</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_GAKKORYAKSYO")) %></td>
        </tr -->
        <tr>
            <td class=disph align="left" width="100">区分</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M01_SYOBUNRUIMEI")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">郵便番号</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_YUBIN_BANGO")) %></td>
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
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_JUSYO1")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所（２）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_JUSYO2")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所（３）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_JUSYO3")) %></td>
        </tr -->

        <tr>
            <td class=disph align="left" width="100">住所</td>
            <td class=disp align="left" width="300">
                <%=gf_HTMLTableSTR(m_Rs("M31_JUSYO1")) %><BR>
                <%=gf_HTMLTableSTR(m_Rs("M31_JUSYO2")) %>
                <%=gf_HTMLTableSTR(m_Rs("M31_JUSYO3")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">電話番号</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_TEL")) %></td>
        </tr>
    </table>

</td>
</tr>
</table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form action="./default.asp" target="<%=C_MAIN_FRAME%>">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtKubunCd" value="<%= m_sKubunCd %>">
        <input type="hidden" name="txtKenCd" value="<%= m_sKenCd %>">
        <input type="hidden" name="txtSityoCd" value="<%= m_sSityoCd %>">
        <input type="hidden" name="txtSyuName" value="<%= m_sSyuName %>">
        <input type="hidden" name="txtPageSyu" value="<%= m_iPageSyu %>">
        <input type="hidden" name="txtSyuCd" value="">
        <input type="hidden" name="txtSyuKbn" value="<%=m_iSyuKbn%>">
    <input type="submit" class=button value="戻　る">
    </form>
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










