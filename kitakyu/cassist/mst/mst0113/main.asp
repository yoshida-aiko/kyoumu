<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 中学校情報検索
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0113/main.asp
' 機      能: 下ページ 中学校マスタの一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtKenCd        :県コード
'           txtSityoCd      :市町村コード
'           txtTyuName      :中学校名称（一部）
'           txtPageTyu      :表示済表示頁数（自分自身から受け取る引数）
'           txtTyuKbn       :中学校区分
'           txtMode         :モード
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           txtTyuCd        :選択された中学校コード
'           txtPageTyu      :表示済表示頁数（自分自身に引き渡す引数）
'           txtMode         :モード
' 説      明:
'           ■初期表示
'               検索条件にかなう中学校を表示
'           ■次へ、戻るボタンクリック時
'               指定した条件にかなう中学校を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/16 高丘 知央
' 変      更: 2001/07/31 根本 直美  変数名命名規則に基く変更
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
    Public  m_iTyuKbn       ':中学校区分
    Public  m_iTyuKbnD      ':中学校区分(DB)
    Public  m_sTyuKbnMei    ':中学校区分名
    Public  m_iPageTyu      ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_iNendo        ':年度      '//2001/07/31変更
    Public  m_sMode         ':モード
    Public  m_Rs            'recordset

    'ページ関係
    Public  m_iMax          ':最大ページ
    Public  m_iDsp                      '// 一覧表示行数

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
    m_iDsp = C_PAGE_LINE

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
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TEL "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M12.M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M16.M16_KENMEI "
        'w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM M13_TYUGAKKO M13 "
        w_sSQL = w_sSQL & vbCrLf & " , M16_KEN M16 "
        w_sSQL = w_sSQL & vbCrLf & " , M12_SITYOSON M12 "
        'w_sSQL = w_sSQL & vbCrLf & " , M01_KUBUN M01 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & "      M13.M13_KEN_CD = M16.M16_KEN_CD "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_KEN_CD = M12.M12_KEN_CD "
        w_sSQL = w_sSQL & vbCrLf & "  AND M12.M12_YUBIN_BANGO LIKE '%00' "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_SITYOSON_CD = M12.M12_SITYOSON_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_NENDO = M16.M16_NENDO(+) "
        'w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_NENDO = M01.M01_NENDO(+) "
        'w_sSQL = w_sSQL & vbCrLf & "  AND M01.M01_DAIBUNRUI_CD = " & C_TYUGAKKO_KBN
        'w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_TYUGAKKO_KBN = M01.M01_SYOBUNRUI_CD(+) "

'response.write w_sSQL

        '抽出条件の作成
        If m_sKenCd<>"" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M13_KEN_CD = '" & m_sKenCd & "' "
        End If
        If m_sSityoCd<>"" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M13_SITYOSON_CD = '" & m_sSityoCd & "' "
        End If
        If m_sTyuName<>"" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M13_TYUGAKKOMEI Like '%" & m_sTyuName & "%' "
        End If
        If m_iTyuKbn <> "" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M13_TYUGAKKO_KBN = " & m_iTyuKbn
        End If
        
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M13_TYUGAKKO_CD"

'Response.Write w_sSQL & "<br>"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        Else
            'ページ数の取得
            m_iMax = gf_PageCount(m_Rs,m_iDsp)
        End If
        
            If m_Rs.EOF Then
            '// ページを表示
            Call showPage_NoData()
        Else
            '// ページを表示
            Call showPage()
        End If
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
        w_sSQL = w_sSQL & " AND M01_SYOBUNRUI_CD = " & m_iTyuKbnD
        
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

    m_sMode = Request("txtMode")


    '// BLANKの場合は行数ｸﾘｱ
    If Request("txtMode") = "Search" Then
        m_iPageTyu = 1
    Else
        m_iPageTyu = INT(Request("txtPageTyu"))     ':表示済表示頁数（自分自身から受け取る引数）
    End If
    
    m_iTyuKbn = Request("txtTyuKbn")
    'コンボ未選択時
    If m_iTyuKbn="@@@" Then
        m_iTyuKbn= ""
    else
        m_iTyuKbn = CInt(m_iTyuKbn)
    End If
    
End Sub

'********************************************************************************
'*  [機能]  DB値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetDB()

m_iTyuKbnD = m_Rs("M13_TYUGAKKO_KBN")

if m_iTyuKbnD = "" Then
    m_iTyuKbnD = 0
end if

End Sub

Sub showPage_NoData()
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
   </head>

    <body>

    <center>
		<br><br><br>
		<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
    </center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    Dim w_pageBar           'ページBAR表示用

    Dim w_iRecordCnt        '//レコードセットカウント
    Dim w_iCnt
    Dim w_bFlg
    
    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

    'ページBAR表示
    Call gs_pageBar(m_Rs,m_iPageTyu,m_iDsp,w_pageBar)

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

        document.frm.action="main.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageTyu.value = p_iPage;
        document.frm.submit();
    
    }
    
    //************************************************************
    //  [機能]  選択した中学校の詳細を表示する。
    //  [引数]  p_sTyuCd    :選択した中学校コード
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_GoSyosai(p_sTyuCd){

        document.frm.action="syousai.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtTyuCd.value = p_sTyuCd;
        document.frm.txtMode.value = "<%=m_sMode%>";
        document.frm.txtPageTyu.value = "<%=m_iPageTyu%>";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>

    <link rel=stylesheet href=../../common/style.css type=text/css>

    </head>

    <body>

    <center>
<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr><td align="center">
<br>
<span class=CAUTION>※ 中学校名をクリックすると詳細を確認できます。</span>
<%=w_pageBar %>
    <table border=1 class=hyo width="100%">
    <COLGROUP WIDTH="15%">
    <COLGROUP WIDTH="15%">
    <COLGROUP WIDTH="40%">
    <COLGROUP WIDTH="15%">
    <COLGROUP WIDTH="15%">
    <tr>
        <th class=header>区分</th>
        <th class=header>県</th>    
        <th class=header>市町村</th>
        <th class=header>中学校名</th>
        <th class=header>電話番号</th>
    </tr>

        <%
        'm_Rs.MoveFirst
        Do While (w_bFlg)
        call gs_cellPtn(w_cell)
        call s_SetDB()
        call s_GetTyugakkoKbn()
        %>
        <tr>
        <td class="<%=w_cell%>" align = "left"><%=m_sTyuKbnMei%></td>
        <td class="<%=w_cell%>" align = "left"><%=m_Rs("M16_KENMEI") %></td>
        <td class="<%=w_cell%>" align = "left"><%=m_Rs("M12_SITYOSONMEI") %></td>
        <td class="<%=w_cell%>" align = "left"><a href="javascript:f_GoSyosai('<%=Trim(m_Rs("M13_TYUGAKKO_CD")) %>')"><%=Trim(m_Rs("M13_TYUGAKKOMEI")) %></a></td>
        <td class="<%=w_cell%>"><font size="2"><%=m_Rs("M13_TEL") %></font></td>
        </tr>
        <%
            m_Rs.MoveNext

            If m_Rs.EOF Then
                w_bFlg = False
            ElseIf w_iCnt >= C_PAGE_LINE Then
                w_bFlg = False
            Else
                w_iCnt = w_iCnt + 1
            End If
        Loop

    'LABEL_showPage_OPTION_END
    %>
        </table>

<%=w_pageBar %>
</td></tr></table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form name ="frm"  Method="POST">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtKenCd" value="<%=m_sKenCd%>">
        <input type="hidden" name="txtSityoCd" value="<%=m_sSityoCd%>">
        <input type="hidden" name="txtTyuName" value="<%=m_sTyuName%>">
        <input type="hidden" name="txtPageTyu" value="<%=m_iPageTyu%>">
        <input type="hidden" name="txtNendo" value="<%= Session("NENDO") %>">
        <input type="hidden" name="txtTyuCd" value="">
        <input type="hidden" name="txtTyuKbn" value="<%=m_iTyuKbn%>">
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
