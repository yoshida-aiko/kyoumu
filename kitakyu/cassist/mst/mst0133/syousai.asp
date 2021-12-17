<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 進路先情報検索
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0133/syousai.asp
' 機      能: 下ページ 就職先マスタの詳細表示を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSinroCD      :進路コード
'           txtSingakuCd        :進学コード
'           txtSinroName        :進路名称（一部）
'           txtPageSinro        :表示済表示頁数（自分自身から受け取る引数）
'           txtSentakuSinroCD       :選択された進路コード
'           txtSentakuSinroKbn       :選択された進路区分
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 　      　:session("PRJ_No")      '権限ﾁｪｯｸのキー '/2001/07/31追加
'           txtSinroCD          :進路区分（戻るとき）
'           txtSingakuCd        :進学コード（戻るとき）
'           txtSinroName        :進路名称（戻るとき）
'           txtPageSinro        :表示済表示頁数（戻るとき）
' 説      明:
'           ■初期表示
'               指定された進学先・就職先の詳細データを表示
'           ■地図画像ボタンクリック時
'               指定した条件にかなう進学先・就職先を表示する（別ウィンドウ）
'-------------------------------------------------------------------------
' 作      成: 2001/06/21 岩下 幸一郎
' 変      更: 2001/07/25 根本　直美
'           : 2001/07/31 根本　直美     変数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_iSinroCD          ':進路区分  '/2001/07/31変更
    Public  m_iSingakuCd        ':進学区分
    Public  m_sSyusyokuName     ':進路名称（一部）
    Public  m_iPageCD           ':表示済表示頁数（自分自身から受け取る引数）'/2001/07/31変更
    Public  m_Rs                'recordset
    Public  m_iNendo            ':年度
    Public  m_sSentakuSinroCD   ':コンボボックスで選択された進路CD
    Public  m_sMode             ':モード
    Public  m_iSentakuSinroKbn  ':選択された進路区分
    
    Public m_sKbn               ':区分
    Public m_sSinromei          ':進路名
    Public m_sSinromeiKan       ':進路名（カナ）
    Public m_sSinromeiRya       ':進路名（略称）
    Public m_sJyusyo1           ':住所（１）
    Public m_sJyusyo2           ':住所（２）
    Public m_sJyusyo3           ':住所（３）
    Public m_sTel               ':進路先電話番号
    Public m_sYubin             ':進路先郵便番号
    Public m_sUrl               ':URL
    Public m_iGyosyuKbn         ':業種区分
    Public m_iSihonkin          ':資本金（単位：万円）
    Public m_iSihonkinY         ':資本金（単位：円）
    Public m_iJyugyoin_Suu      ':従業員数
    Public m_iSyoninkyu         ':初任給
    Public m_sBiko              ':備考
    Public m_iSinroKbn          ':進路区分
    Public m_sKbnName			':種別名称

    'Public Const C_SYORYAKU_KETA = 4    '//表示時に省略する桁数（資本金）
    

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

        '// 値の初期化
        Call s_SetBlank()
        
        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '就職先マスタを取得
        w_sWHERE = ""

        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M01_1.M01_NENDO "
        w_sSQL = w_sSQL & vbCrLf & " ,M01_1.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_NENDO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI_KANA "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_YUBIN_BANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINGAKU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_GYOSYU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SIHONKIN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JYUGYOIN_SUU "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SYONINKYU "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_BIKO "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & "    ,M01_KUBUN M01_1 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "        M01_1.M01_NENDO = " & m_iNendo & ""
        w_sSQL = w_sSQL & vbCrLf & "    AND M01_1.M01_DAIBUNRUI_CD = " & C_SINRO & " "
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_NENDO = " & m_iNendo & ""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN = M01_1.M01_SYOBUNRUI_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "    AND M32_SINRO_CD = '" & m_sSentakuSinroCD & "' "

'Response.Write w_sSQL & "<br>"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If
        
        '//DBから値を取得
        Call s_SetDB()
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
    call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  全値を初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetBlank()
    
    m_sSentakuSinroCD = ""
    m_iNendo = ""
    m_iSinroCD = ""
    m_iSingakuCd = ""
    m_sMode = ""
    m_sSyusyokuName = ""
    m_iPageCD = ""
    
    m_sKbn = ""
    m_sSinromei = ""
    m_sSinromeiKan = ""
    m_sSinromeiRya = ""
    m_sJyusyo1 = ""
    m_sJyusyo2 = ""
    m_sJyusyo3 = ""
    m_sTel = ""
    m_sYubin = ""
    m_sUrl = ""
    m_iGyosyuKbn = ""
    m_iSihonkin = ""
    m_iJyugyoin_Suu = ""
    m_iSyoninkyu = ""
    m_sBiko = ""
    m_iSinroKbn = ""
    m_iSinroKbnY = ""
    m_iSentakuSinroKbn = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_sSentakuSinroCD = Request("txtSentakuSinroCD")    ':進路コード

    m_iNendo = Session("NENDO")     ':年度

    m_iSinroCD = Request("txtSinroCD")      ':進路区分
    'コンボ未選択時
    If m_iSinroCD="@@@" Then
        m_iSinroCD=""
    End If

    m_iSingakuCd = Request("txtSingakuCd")      ':進学区分
    'コンボ未選択時
    If m_iSingakuCd="@@@" Then
        m_iSingakuCd=""
    End If

    m_sMode = Request("txtMode")        ':モード

    m_sSyusyokuName = Request("txtSyusyokuName")    ':就職先名称（一部）

    '// BLANKの場合は行数ｸﾘｱ
    If Request("txtMode") = "Search" Then
        m_iPageCD = 1
    Else
        m_iPageCD = INT(Request("txtPageSyusyoku"))     ':表示済表示頁数（自分自身から受け取る引数）
    End If
    
    m_iSentakuSinroKbn = CInt(Request("txtSentakuSinroKbn"))    ':進路区分
    
End Sub

'********************************************************************************
'*  [機能]  DBから取得した値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetDB()

Dim w_iSihonkin

	if IsNull(m_Rs("M01_SYOBUNRUIMEI")) = False Then
	    m_sKbn = m_Rs("M01_SYOBUNRUIMEI")
	end if

	if IsNull(m_Rs("M32_SINROMEI")) = False Then
	    m_sSinromei = m_Rs("M32_SINROMEI")
	end if

	if IsNull(m_Rs("M32_SINROMEI_KANA")) = False Then
	    m_sSinromeiKan = m_Rs("M32_SINROMEI_KANA")
	end if
	if IsNull(m_Rs("M32_SINRORYAKSYO")) = False Then
	    m_sSinromeiRya = m_Rs("M32_SINRORYAKSYO")
	end if
	if IsNull(m_Rs("M32_JUSYO1")) = False Then
	    m_sJyusyo1 = m_Rs("M32_JUSYO1")
	end if
	if IsNull(m_Rs("M32_JUSYO2")) = False Then
	    m_sJyusyo2 = m_Rs("M32_JUSYO2")
	end if
	if IsNull(m_Rs("M32_JUSYO3")) = False Then
	    m_sJyusyo3 = m_Rs("M32_JUSYO3")
	end if
	if IsNull(m_Rs("M32_DENWABANGO")) = False Then
	    m_sTel = m_Rs("M32_DENWABANGO")
	end if
	if IsNull(m_Rs("M32_YUBIN_BANGO")) = False Then
	    m_sYubin = m_Rs("M32_YUBIN_BANGO")
	end if
	if IsNull(m_Rs("M32_SINRO_URL")) = False Then
	    m_sUrl = m_Rs("M32_SINRO_URL")
	end if

	if IsNull(m_Rs("M32_GYOSYU_KBN")) = False Then
	    m_iGyosyuKbn = m_Rs("M32_GYOSYU_KBN")
	end if

	if IsNull(m_Rs("M32_SIHONKIN")) = False Then
	    m_iSihonkinY = m_Rs("M32_SIHONKIN")
	    w_iSihonkin = CInt(Len(m_iSihonkinY)) - C_SYORYAKU_KETA
	    m_iSihonkin = Mid(m_iSihonkinY,1,w_iSihonkin)
	end if

	if IsNull(m_Rs("M32_JYUGYOIN_SUU")) = False Then
	    m_iJyugyoin_Suu = m_Rs("M32_JYUGYOIN_SUU")
	end if
	if IsNull(m_Rs("M32_SYONINKYU")) = False Then
	    m_iSyoninkyu = m_Rs("M32_SYONINKYU")
	end if
	if IsNull(m_Rs("M32_BIKO")) = False Then
	    m_sBiko = m_Rs("M32_BIKO")
	end if
	if IsNull(m_Rs("M32_SINRO_KBN")) = False Then
	    m_iSinroKbn = m_Rs("M32_SINRO_KBN")
	end if


	'//進路区分OR業種区分名称を取得
	Select case cint(gf_SetNull2Zero(m_Rs("M32_SINRO_KBN")))
		Case C_SINRO_SINGAKU	'//進路区分が進学の場合

			'//進学区分名称を取得
			w_sKbn = trim(m_Rs("M32_SINGAKU_KBN"))
			If w_sKbn <> "" Then
				Call gf_GetKubunName(C_SINGAKU,m_Rs("M32_SINGAKU_KBN"),m_iNendo,m_sKbnName)
			End If

		Case C_SINRO_SYUSYOKU	'//進路区分が就職の場合

			'//業種区分名称を取得
			w_sKbn = trim(m_Rs("M32_GYOSYU_KBN"))
			If w_sKbn <> "" Then
				Call gf_GetKubunName(C_GYOSYU_KBN,m_Rs("M32_GYOSYU_KBN"),m_iNendo,m_sKbnName)
			End If

		Case C_SINRO_SONOTA	'//進路区分がその他の場合
			m_sKbnName = ""
	End Select


End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
'Sub s_MapHTML()

'   If ISNULL(m_Rs("M13_TIZUFILENAME")) OR m_Rs("M13_TIZUFILENAME")="" Then
'       Response.Write("登録されていません")
'   Else
'       Response.Write("<a Href=""javascript:f_OpenWindow('" & Session("TYUGAKU_TIZU_PATH") & m_Rs("M13_TIZUFILENAME") & "')"">周辺地図</a>")
'   End If
    
'End Sub


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

w_slink = "　"

if m_Rs("M32_SINRO_URL") <> "" Then 
    w_sLink= "<a href='" & gf_HTMLTableSTR(m_sUrl) & "'>" 
    w_sLink= w_sLink &  gf_HTMLTableSTR(m_sUrl) & "</a>"
End if

        %>
        <%=w_slink%>
        <%
            m_Rs.MoveNext


    'LABEL_showPage_OPTION_END
End sub


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
        document.frm.txtPageSinro.value = p_iPage;
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

    //-->
    </SCRIPT>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    </head>

    <body>

    <center>

<%
m_sSubtitle = "詳　細"

call gs_title("進路先情報検索",m_sSubtitle)
%>

    <table border=1 class=disp width="400">
        <tr>
            <td class=disph align="left" width="100">名称</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sSinromei) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">名称（カナ）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sSinromeiKan) %></td>
        </tr>

        <tr>
            <td class=disph align="left" width="100">略称</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sSinromeiRya) %></td>
        </tr>

        <tr>
            <td class=disph align="left" width="100">進路区分</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sKbn) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">種別区分</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sKbnName) %></td>
        </tr>

        <tr>
            <td class=disph align="left" width="100">郵便番号</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sYubin) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所（１）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sJyusyo1) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所（２）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sJyusyo2) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">住所（３）</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sJyusyo3) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">電話番号</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sTel) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">URL</td>
            <td class=disp align="left" width="300"><% S_syousai() %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">資本金</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_iSihonkin) %>万円</td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">従業員数</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_iJyugyoin_Suu) %>人</td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">初任給</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_iSyoninkyu) %>円</td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">備考</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sBiko) %></td>
        </tr>

    </table>


    <br>


    <table border="0">
    <tr>
    <td valign="top">
    <form name ="frm" action="./default.asp" target="<%=C_MAIN_FRAME%>">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtSinroCD" value="<%= m_iSinroCD %>">
        <input type="hidden" name="txtSingakuCD" value="<%= m_iSingakuCd %>">
        <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
        <input type="hidden" name="txtPageCD" value="<%= m_iPageCD %>">
    <input class=button type="submit" value="戻　る">
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