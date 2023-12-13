<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 進路先情報検索
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0133/main.asp
' 機      能: 下ページ 就職先マスタの一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSinroCD      :進路コード
'           txtSingakuCD        :進学コード
'           txtSinroName        :就職先名称（一部）
'           txtPageSyusyoku     :表示済表示頁数（自分自身から受け取る引数）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           txtSinroCD      :進路コード             '/2001/07/31追加
'           txtSingakuCD        :進学コード         '/2001/07/31追加
'           txtSinroName        :就職先名称（一部）
'           txtSentakuSinroCD   :選択された進路コード
'           txtSentakuSinroKbn   :選択された進路区分
'           txtPageSyusyoku     :表示済表示頁数（自分自身に引き渡す引数）
' 説      明:
'           ■初期表示
'               検索条件にかなう就職・進学先を表示
'           ■次へ、戻るボタンクリック時
'               指定した条件にかなう就職・進学を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/18 岩下　幸一郎
' 変      更: 2001/07/31 根本 直美  引数・引渡追加
'           :                       変数名命名規則に基く変更
'           : 2001/08/10 根本 直美  NN対応に伴うソース変更
'           : 2001/08/22 伊藤 公子  検索SQL文変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_iSinroCD          ':進路コード        '/2001/07/31変更
    Public  m_iSingakuCd        ':進学コード        '/2001/07/31変更
    Public  m_sSyusyokuName     ':就職先名称（一部）
    Public  m_iPageCD           ':表示済表示頁数（自分自身から受け取る引数）'/2001/07/31変更
    Public  m_skubun            ':区分名称
    Public  m_Rs                'recordset
    Public  m_iNendo            ':年度
    Public  m_sMode             ':モード
    Public  m_iFLG              ':
    Public  m_sSNm              ':
    'Public  m_sSinroKBN        ':進路区分
    Public  m_iSinroKbn         ':進路区分コード
    

    'ページ関係
    Public  m_iMax              ':最大ページ
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
    w_sMsgTitle="進路先情報検索"
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

        '就職マスタを取得
        w_sWHERE = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " 	M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINRO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINGAKU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_GYOSYU_KBN"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & " 	M32_SINRO M32, "
        w_sSQL = w_sSQL & vbCrLf & " 	("
        w_sSQL = w_sSQL & vbCrLf & " 	select * "
        w_sSQL = w_sSQL & vbCrLf & " 	from "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & " 	where "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_DAIBUNRUI_CD  = " & C_SINRO & " and "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & vbCrLf & " 	) M01"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " 	M32.M32_SINRO_KBN = M01.M01_SYOBUNRUI_CD (+) and "
		w_sSQL = w_sSQL & vbCrLf & " 	M32.M32_NENDO = M01.M01_NENDO (+) and "
		w_sSQL = w_sSQL & vbCrLf & " 	M32.M32_NENDO = " & m_iNendo & ""
		
        '抽出条件の作成
        'If m_sSinroKBN <> "" Then
        If m_iSinroCD <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN =" & m_iSinroCD & " "
        End If
        
        If m_iSingakuCd <> "" Then
			if m_iSinroCD = 1 then
	            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN =" & m_iSingakuCd & " "
			ElseIf m_iSinroCD = 2 then
	            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_GYOSYU_KBN =" & m_iSingakuCd & " "
			End if
        End If
        
        If m_sSyusyokuName<>"" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINROMEI Like '%" & m_sSyusyokuName & "%' "
        End If

        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M32.M32_SINRO_CD "
		
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

            If m_iFLG = "1" Then
                '// ページを表示
                Call showPage_SHOW()
            Else
                '// ページを表示
                Call showPage()
            End If
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

    m_iNendo = Session("NENDO")         ':年度

    m_iSinroCD = Request("txtSinroCD")      ':進路区分
    'コンボ未選択時
    If m_iSinroCD = "@@@" Then
        m_iSinroCD = ""
    End If

    m_iSingakuCd = Request("txtSingakuCd")      ':進学区分
    'コンボ未選択時
    If m_iSingakuCd="@@@" Then
        m_iSingakuCd=""
    End If

    m_sMode = Request("txtMode")            ':モード

    m_sSyusyokuName = Request("txtSyusyokuName")    ':就職先名称（一部）

    If m_sMode = "Search" Then
        m_iPageCD = 1
    Else
        m_iPageCD = INT(Request("txtPageSyusyoku")) ':表示済表示頁数（自分自身から受け取る引数）
    End If

'    If cstr(gf_SetNull2String(m_iSinroCD)) = "1" Then            ':ヘッダーの区分名称変更
'        m_skubun = "進学区分"
'	ElseIf cstr(gf_SetNull2String(m_iSinroCD)) = "2" Then
'        m_skubun = "業種区分"
'    else
'        'm_skubun = "進路区分"
'        m_skubun = "種別区分"
'    End If

    m_iDisp = C_PAGE_LINE       '１ページ最大表示数

    m_iFLG = request("txtFLG")
    m_sSNm = request("txtSNm")

	if gf_IsNull(request("txtSNm")) then
		m_sSNm = request("SearchNm")
	End if

    m_iSinroKbn = ""
    
End Sub


Sub S_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_slink
Dim w_iCnt

w_iCnt = 0

If m_iFLG <> "1" Then

    Do While not m_Rs.EOF

    w_slink = "　"
    m_iSinroKbn = m_Rs("M32_SINRO_KBN")

    if m_Rs("M32_SINRO_URL") <> "" Then 
        'w_sLink= "<a href='" & gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "'>" 
        w_sLink= "<a href='" & m_Rs("M32_SINRO_URL") & "' target='_site'>" 
        w_sLink= w_sLink &  m_Rs("M32_SINRO_URL") & "</a>"
    End if

        '//テーブルセル背景色
        call gs_cellPtn(w_cellT)
        %>
        <tr>

		<%
		'//初期化
		w_sKbn = ""
		w_sKbnName = ""

		'//進路区分OR業種区分を取得
		Select case cint(gf_SetNull2Zero(m_Rs("M32_SINRO_KBN")))
			Case C_SINRO_SINGAKU	'//進路区分が進学の場合

				'//進学区分名称を取得
				w_sKbn = trim(m_Rs("M32_SINGAKU_KBN"))
				If w_sKbn <> "" Then
					Call gf_GetKubunName(C_SINGAKU,m_Rs("M32_SINGAKU_KBN"),m_iNendo,w_sKbnName)
				End If

			Case C_SINRO_SYUSYOKU	'//進路区分が就職の場合

				'//業種区分名称を取得
				w_sKbn = trim(m_Rs("M32_GYOSYU_KBN"))
				If w_sKbn <> "" Then
					Call gf_GetKubunName(C_GYOSYU_KBN,m_Rs("M32_GYOSYU_KBN"),m_iNendo,w_sKbnName)
				End If

			Case C_SINRO_SONOTA	'//進路区分がその他の場合

		End Select

		%>

        <td align="left" class=<%=w_cellT%>><%=m_Rs("M01_SYOBUNRUIMEI") %></td>
        <td align="left" class=<%=w_cellT%>><%=w_sKbnName%></td>
        <td align="left" class=<%=w_cellT%>><a href="javascript:f_GoSyosai('<%=m_Rs("M32_SINRO_CD") %>','<%=m_iSinroKbn%>')"><%=trim(m_Rs("M32_SINROMEI")) %></a></td>
        <td align="left" class=<%=w_cellT%>><%=m_Rs("M32_DENWABANGO") %></td>
        <td align="left" class=<%=w_cellT%>><%=w_slink%></td>
        </tr>
        <%
        m_Rs.MoveNext

        If w_iCnt >= C_PAGE_LINE Then
            Exit Do
        Else
            w_iCnt = w_iCnt + 1
        End If
    Loop

Else 

    Do While not m_Rs.EOF
        Call gs_cellPtn(w_cell)

        %>
        <tr>
        <td align="left" class=<%=w_cell%>><%=m_Rs("M01_SYOBUNRUIMEI") %></td>
        <td align="left" class=<%=w_cell%>>
        <input type=button class=<%=w_cell%> name="SinroNm_<%=w_iCnt%>" value='<%=m_Rs("M32_SINROMEI") %>' onclick="iinSelect(<%=w_iCnt%>)">
        <input type=hidden name="SinroCd_<%=w_iCnt%>" value='<%=m_Rs("M32_SINRO_CD") %>'>
        </td>
        <td align="left" class=<%=w_cell%>><%=m_Rs("M32_DENWABANGO") %></td>
        </tr>
        <%
        m_Rs.MoveNext

        If w_iCnt >= C_PAGE_LINE Then
            Exit Do
        Else
            w_iCnt = w_iCnt + 1
        End If
    Loop

End If

    'LABEL_showPage_OPTION_END
End sub

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

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

    'ページBAR表示
    Call gs_pageBar(m_Rs,m_iPageCD,m_iDsp,w_pageBar)

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
        document.frm.txtPageSyusyoku.value = p_iPage;
        document.frm.submit();
    
    }
    
    //************************************************************
    //  [機能]  詳細ページを表示
    //  [引数]  p_sSinroCD:進路コード
    //          p_sSinroKbn:進路区分
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_GoSyosai(p_sSinroCD,p_sSinroKbn){

        document.frm.action="syousai.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtSentakuSinroCD.value = p_sSinroCD;
        document.frm.txtSentakuSinroKbn.value = p_sSinroKbn;
        document.frm.txtMode.value = "Search";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
    </head>

    <body>

    <center>
<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr><td align="center">
<br>
<span class=CAUTION>※ 進路名をクリックすると詳細を確認できます。</span>
<%=w_pageBar %>

        <table border=1 class=hyo width="100%">
        <COLGROUP WIDTH="15%">
        <COLGROUP WIDTH="15%">
        <COLGROUP WIDTH="30%">
        <COLGROUP WIDTH="25%">
        <COLGROUP WIDTH="30%">
        <tr>
        <th class=header>進路区分</th>
        <th class=header>種別区分</th>
        <th class=header>進路名</th>
        <th class=header>TEL</th>
        <th class=header>URL</th>
        </tr>

    <% S_syousai() %>

        </table>

<%=w_pageBar %>

</td></tr></table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form name ="frm" action="" target="">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtSinroCD" value="<%= m_iSinroCD %>">
        <input type="hidden" name="txtSingakuCD" value="<%= m_iSingakuCd %>">
        <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
        <input type="hidden" name="txtPageSyusyoku" value="<%= m_iPageCD %>">
        <input type="hidden" name="txtSentakuSinroCD" value="">
        <input type="hidden" name="txtSentakuSinroKbn" value="">
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

Sub showPage_SHOW()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    Dim w_pageBar           'ページBAR表示用

    Dim w_iRecordCnt        '//レコードセットカウント

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

    'ページBAR表示
    Call gs_pageBar(m_Rs,m_iPageCD,m_iDsp,w_pageBar)

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
        document.frm.target="main";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageSyusyoku.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  申請内容表示用ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function iinSelect(p_sct) {

        //挿入元のフォームを取得
            w_sctNm = eval("document.frm.SinroNm_"+p_sct);
            w_sctNo = eval("document.frm.SinroCd_"+p_sct);

        //挿入処理
            parent.opener.document.frm.SinroNm.value = w_sctNm.value;
            parent.opener.document.frm.SinroCd.value = w_sctNo.value;

            document.frm.SearchNm.value = w_sctNm.value;
            document.frm.SearchNo.value = w_sctNo.value;

        return true;    
        //window.close()

    }

    //************************************************************
    //  [機能]  クリアボタンをクリックした場合
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_Clear(p_No) {

        document.frm.SearchNm.value = "";
        document.frm.SearchNo.value = "";

        //挿入させたいフォームを取得
            w_NmStr = parent.opener.document.frm.SinroNm;
            w_NoStr = parent.opener.document.frm.SinroCd;

        //挿入処理

            w_NmStr.value = document.frm.SearchNm.value;
            w_NoStr.value = document.frm.SearchNo.value;
        return true;    
    }
    //-->
    </SCRIPT>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    </head>

    <body>
    <center>
    <form name="frm" method="post">

    <table width="90%" border="0">
        <tr>
            <td align="center">
                <table width="80%" class="hyo">
                    <tr>
                        <td align="center" width="30%"><font color="white">進　路　名</font></td>
                        <td align="center" class="detail"><input type="text" class="noBorder" name="SearchNm" value="<%=m_sSNm%>" readonly><input type="hidden" name="SearchNo" value=""></td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <span class="CAUTION">※ 選択をするには進路名をクリックしてください。</span>
    <table border="0" align="center">
    <tr>
    <td valign="top">
        <input type="button" value=" クリア " class="button" onclick="javascript:f_Clear()">
        <input type="button" value="閉じる" class="button" onclick="javascript:parent.window.close()">
    </td>
    </tr>
    </table>

                <%=w_pageBar %>
                <table border="1" class="hyo" width="100%">
                    <tr>
                        <th class="header" width="10%" nowrap>進路区分</th>
                        <th class="header" width="50%">進路名</th>
                        <th class="header" width="40%">TEL</th>
                    </tr>
                    <% S_syousai() %>
                </table>
                <%=w_pageBar %>
            </td>
        </tr>
    </table>

    <table border="0" align="center">
    <tr>
    <td valign="top">
        <input type="button" value=" クリア " class="button" onclick="javascript:f_Clear()">
        <input type="button" value="閉じる" class="button" onclick="javascript:parent.window.close()">
    </td>
    </tr>
    </table>

	    <input type="hidden" name="txtMode" value="<%=m_sMode%>">
	    <input type="hidden" name="txtSinroCD" value="<%= m_iSinroCD %>">
	    <input type="hidden" name="txtSingakuCD" value="<%= m_iSingakuCd %>">
	    <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
	    <input type="hidden" name="txtPageSyusyoku" value="<%= m_iPageCD %>">
	    <input type="hidden" name="txtSentakuSinroCD" value="">
	    <input type="hidden" name="txtSentakuSinroKbn" value="">
	    <input type="hidden" name="txtFLG" value="<%=m_iFLG%>">
    </form>

    </center>
    </body>
    </html>



<%
    '---------- HTML END   ----------
End Sub
%>