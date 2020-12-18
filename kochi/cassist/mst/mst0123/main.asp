<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 高等学校マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0123/main.asp
' 機      能: 下ページ 高等学校マスタの一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtKubunCd      :区分コード
'           txtKenCd        :県コード
'           txtSityoCd      :市町村コード
'           txtSyuName      :高等学校名称（一部）
'           txtPageSyu      :表示済表示頁数（自分自身から受け取る引数）
'           txtSyuKbn       :高等学校区分
'           txtMode
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           txtSyuCd        :選択された高等学校コード
'           txtPageSyu      :表示済表示頁数（自分自身に引き渡す引数）
'           txtMode
' 説      明:
'           ■初期表示
'               検索条件にかなう高等学校を表示
'           ■次へ、戻るボタンクリック時
'               指定した条件にかなう高等学校を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/20 岩下 幸一郎
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
    Public  m_iKubunCd      ':区分コード    '/2001/07/31変更
    Public  m_sKenCd        ':県コード
    Public  m_sSityoCd      ':市町村コード
    Public  m_sSyuName      ':高等学校名称（一部）
    Public  m_sNendo        ':年度
    Public  m_sMode         ':モード
    Public  m_Rs            'recordset
    Public  m_sPageSyu      ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_iSyuKbn       ':中学校区分
    
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
    w_sMsgTitle="高等学校情報検索"
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

        '高等学校マスタを取得
        w_sWHERE = ""
        
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M31.M31_GAKKO_CD  "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_GAKKOMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_TEL "
        w_sSQL = w_sSQL & vbCrLf & " ,M12.M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M16.M16_KENMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & " 	M31_SYUSSINKO M31, "
        w_sSQL = w_sSQL & vbCrLf & " 	M16_KEN M16, "
        w_sSQL = w_sSQL & vbCrLf & " 	("
        w_sSQL = w_sSQL & vbCrLf & " 	select * "
        w_sSQL = w_sSQL & vbCrLf & " 	from "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & " 	where "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_DAIBUNRUI_CD  = " & C_SYUSSINKO & " and "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_NENDO = " & m_sNendo
        w_sSQL = w_sSQL & vbCrLf & " 	) M01, "
        w_sSQL = w_sSQL & vbCrLf & " 	("
        w_sSQL = w_sSQL & vbCrLf & " 	select * "
        w_sSQL = w_sSQL & vbCrLf & " 	from "
        w_sSQL = w_sSQL & vbCrLf & " 		M12_SITYOSON "
        w_sSQL = w_sSQL & vbCrLf & " 	where "
        w_sSQL = w_sSQL & vbCrLf & " 		M12_YUBIN_BANGO LIKE '%00' "
        w_sSQL = w_sSQL & vbCrLf & " 	) M12 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_NENDO = " & m_sNendo & " and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_KEN_CD = M16.M16_KEN_CD (+) and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_NENDO = M16.M16_NENDO(+) and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_KEN_CD = M12.M12_KEN_CD (+) and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_SITYOSON_CD = M12.M12_SITYOSON_CD (+) and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_GAKKO_KBN = M01.M01_SYOBUNRUI_CD (+) and "
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_NENDO = M01.M01_NENDO (+) "
        
        '抽出条件の作成
        If m_iKubunCd <> "" Then
            w_sSQL = w_sSQL & " AND M01.M01_SYOBUNRUI_CD = " & m_iKubunCd   '//2001/07/31変更
        End If
        
        If m_sKenCd <> "" Then
            w_sSQL = w_sSQL & " AND M16.M16_KEN_CD = '" & m_sKenCd & "' "
        End If
        
        If m_sSityoCd <> "" Then
            w_sSQL = w_sSQL & " AND M12.M12_SITYOSON_CD = '" & m_sSityoCd & "' "
        End If
        
        If m_sSyuName <> "" Then
            w_sSQL = w_sSQL & " AND M31.M31_GAKKOMEI Like '%" & m_sSyuName & "%' "
        End If
        
        If m_iSyuKbn <> "" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M31_GAKKO_KBN = " & m_iSyuKbn
        End If
		
        w_sSQL = w_sSQL & " ORDER BY M31.M31_GAKKO_CD"
		
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

            If m_sMode = "" Then
            '// ページを表示
            Call NoPage()
            Exit Do
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
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
    m_sNendo = Session("NENDO")     ':年度
	
    m_iKubunCd = Request("txtKubunCd")  ':区分コード
    'コンボ未選択時
    If m_iKubunCd="@@@" Then
        m_iKubunCd=""
    End If
	
    m_sKenCd = Request("txtKenCd")      ':県コード
    'コンボ未選択時
    If m_sKenCd="@@@" Then
        m_sKenCd=""
    End If

    m_sMode = Request("txtMode")        ':モード

    m_sSityoCd = Request("txtSityoCd")  ':市町村コード
    'コンボ未選択時
    If m_sSityoCd="@@@" Then
        m_sSityoCd=""
    End If

    m_sSyuName = Request("txtSyuName")  ':高校名称（一部）
	
    '// BLANKの場合は行数ｸﾘｱ
	If m_sMode = "Search" Then
        m_sPageSyu = 1
    Else
        m_sPageSyu = INT(Request("txtPageSyu"))     ':表示済表示頁数（自分自身から受け取る引数）
    End If
    
    m_iSyuKbn = Request("txtSyuKbn")
    'コンボ未選択時
    If m_iSyuKbn="@@@" Then
        m_iSyuKbn= ""
    elseif m_iSyuKbn = "" Then
        m_iSyuKbn = ""
    else
        m_iSyuKbn = CInt(m_iSyuKbn)
    End If
	
End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage_NoData()
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

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub NoPage()
%>
	<html>
	<head>
    </head>
	<body>
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
    Call gs_pageBar(m_Rs,m_sPageSyu,m_iDsp,w_pageBar)
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
    //  [機能]  高等学校名をクリックした場合
    //  [引数]  p_sSyuCd :高等学校コード
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_GoSyosai(p_sSyuCd){

        document.frm.action="syousai.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtSyuCd.value = p_sSyuCd;
        document.frm.txtMode.value = "<%=m_sMode%>";
        document.frm.txtPageSyu.value = "<%=m_sPageSyu%>";
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
<span class=CAUTION>※ 高等学校名をクリックすると詳細を確認できます。</span>
<%=w_pageBar %>

        <table border="1" class=hyo width="100%">
        <COLGROUP WIDTH="10%">
        <COLGROUP WIDTH="10%">
        <COLGROUP WIDTH="40%">
        <COLGROUP WIDTH="20%">
        <COLGROUP WIDTH="20%">
        <tr>
        <th class=header>区分</th>
        <th class=header>県</th>
        <th class=header>市町村</th>
        <th class=header>高等学校名</th>
        <th class=header>電話番号</th>
        </tr>

        <%
        Do While (w_bFlg)
        call gs_cellPtn(w_cell)
        %>
        <tr>
        <td class="<%=w_cell%>" align="left"><%=m_Rs("M01_SYOBUNRUIMEI") %></td>
        <td class="<%=w_cell%>" align="left"><%=m_Rs("M16_KENMEI") %></td>
        <td class="<%=w_cell%>" align="left"><%=m_Rs("M12_SITYOSONMEI") %></td>
        <td class="<%=w_cell%>" align="left"><a href="javascript:f_GoSyosai('<%=m_Rs("M31_GAKKO_CD") %>')"><%=Trim(m_Rs("M31_GAKKOMEI"))%></a></td>
        <td class="<%=w_cell%>" align="left"><%=m_Rs("M31_TEL") %></td>
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
	%>
        </table>

<%=w_pageBar %>
</td></tr></table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form name ="frm" action="" target="">
        <input type="hidden" name="txtMode" value="">
        <input type="hidden" name="txtKubunCd" value="<%=m_iKubunCd%>">
        <input type="hidden" name="txtKenCd" value="<%=m_sKenCd%>">
        <input type="hidden" name="txtSityoCd" value="<%=m_sSityoCd%>">
        <input type="hidden" name="txtSyuName" value="<%=m_sSyuName%>">
        <input type="hidden" name="txtPageSyu" value="<%=m_sPageSyu%>">
        <input type="hidden" name="txtNendo" value="<%= Session("NENDO") %>">
        <input type="hidden" name="txtSyuCd" value="">
        <input type="hidden" name="txtSyuKbn" value="<%=m_iSyuKbn%>">
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
