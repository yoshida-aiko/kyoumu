<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 進路先情報登録
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0144/kakunin.asp
' 機      能: 下ページ 進路先マスタの登録確認を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           txtSinroCD      :進路コード
'           txtSingakuCd        :進学コード
'           txtSyusyokuName     :進路名称（一部）
'           txtPageSinro        :表示済表示頁数（自分自身から受け取る引数）
'           Sinro_syuseiCD      :選択された進路コード
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           txtSinroCD      :進路コード（戻るとき）
'           txtSingakuCd        :進学コード（戻るとき）
'           txtSyusyokuName     :進路名称（戻るとき）
'           txtPageSinro        :表示済表示頁数（戻るとき）
' 説      明:
'           ■初期表示
'               指定された進学先・就職先の詳細データを表示
'           ■地図画像ボタンクリック時
'               指定した条件にかなう進学先・就職先を表示する（別ウィンドウ）
'-------------------------------------------------------------------------
' 作      成: 2001/06/22 岩下 幸一郎
' 変      更: 2001/07/12 谷脇 良也
' 　      　: 2001/07/24 根本 直美(DB変更に伴う修正)
' 　　　　　: 2001/08/22 伊藤　公子　業種区分追加対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数

    Public  m_sSINROMEI         ':進路名
    'Public m_sSINROMEI_EIGO    ':進路名英語
    Public  m_sSINROMEI_KANA    ':進路名カナ
    Public  m_sSINRORYAKSYO     ':進路略称
    'Public m_sJUSYO            ':住所
    Public  m_sJUSYO1           ':住所1
    Public  m_sJUSYO2           ':住所2
    Public  m_sJUSYO3           ':住所3
    Public  m_iKenCd            ':県コード
    Public  m_iSityoCd          ':市町村コード
    Public  m_sSityoson         ':市町村名
    Public  m_sDENWABANGO       ':電話番号
    Public  m_iYUBINBANGO       ':郵便番号
    Public  m_sSINRO_URL        ':URL
    Public  m_sSinrokubun       ':データベースから取得した進路区分
    Public  m_sSingakukubun     ':データベースから取得した進学区分
    Public  m_Rs                ':recordset
    Public  m_iNendo            ':年度
    Public  m_sYubin            ':郵便番号
    Public  m_iGyosyu_Kbn       ':業種区分
    Public  m_iSihonkin         ':資本金
    Public  m_iJyugyoin_Suu     ':従業員
    Public  m_iSyoninkyu        ':初任給
    Public  m_sBiko             ':備考

    Public  m_sRenrakusakiCD    ':連絡先コード
    Public  m_sSinroCD      ':進路コード
    Public  m_sSingakuCD        ':進学コード
    Public  m_sSinroCD2     ':Mainから取得した進路コード
    Public  m_sSingakuCD2       ':Mainから取得した進学コード
    Public  m_sSyusyokuName     ':進路名称（一部）
    Public  m_sPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_sMode
    Public  m_bReFlg
    Public  m_sMsgFlg       ':エラーフラグ
    Public  m_sMsg

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
    w_sMsgTitle="進路先情報登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_sMsgFlg = False

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

        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI_KANA "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_KEN_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SITYOSON_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo & " AND "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO_CD = '" & m_sRenrakusakiCD & "' "

'Response.Write w_sSQL & "<br>"

        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If

		'//新規時、データの重複チェック
		If m_sMode = "Sinki" Then
	        If m_Rs.EOF = False Then
	            m_sMsgFlg = True
	            m_sMsg = "入力された進路先コードはすでに使用されています"
	        End If
		End If

        '// ページを表示
        Call showPage()
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    Dim strErrMsg

    strErrMsg = ""

    m_sMode          = Request("txtMode")               'モードの設定
    m_bReFlg         = Request("txtReFlg")              'リロード有無の設定

    m_sRenrakusakiCD = Request("txtRenrakusakiCD")      ':連絡先コード
    m_sSINROMEI      = Request("txtSINROMEI")           ':進路名
    m_sSINROMEI_KANA = Request("txtSINROMEI_KANA")      ':進路名カナ
    'm_sSINROMEI_EIGO = Request("txtSINROMEI_EIGO")     ':進路名英語
    'If m_sSINROMEI_EIGO="" Then m_sSINROMEI_EIGO="　"
    m_sSINRORYAKSYO  = Request("txtSINRORYAKSYO")       ':進路略称
    If m_sSINRORYAKSYO="" Then m_sSINRORYAKSYO="　"
    'm_sJUSYO         = Request("txtJUSYO")             ':住所
    m_iKenCd         = Request("txtKenCd")              ':県コード
    m_iSityoCd       = Request("txtSityoCd")            ':市町村コード（住所1）
    m_sJUSYO1        = Request("txtJUSYO1")             ':住所1
    m_sJUSYO2        = Request("txtJUSYO2")             ':住所2
    m_sJUSYO3        = Request("txtJUSYO3")             ':住所3
    m_iKenCd         = Request("txtKenCd")              ':県コード
    m_iSityoCd       = Request("txtSityoCd")            ':市町村コード
    m_sDENWABANGO    = Request("txtDENWABANGO")         ':電話番号
    m_sSinroCD       = Request("txtSinroCD")            ':進路区分
    m_sSingakuCD     = Request("txtSingakuCd")          ':進学区分
    m_sSINRO_URL     = Request("txtSINRO_URL")          'URL
    if Instr(m_sSINRO_URL,"http://") = 0 then m_sSINRO_URL = "http://" & m_sSINRO_URL
    if m_sSINRO_URL  = "http://" then m_sSINRO_URL = ""
    m_sSinroCD2      = Request("txtSinroCD2")           ':戻り用の進路区分
    m_sSingakuCD2    = Request("txtSingakuCd2")         ':戻り用の進学区分
    m_sSyusyokuName  = Request("txtSyusyokuName")       '戻り用の:就職先名称（一部）
    m_sPageCD        = INT(Request("txtPageCD"))        '戻り用の表示頁

    m_iNendo        = Request("txtNendo")               ':年度
    m_sYubin        = Request("txtYUBINBANGO")          ':郵便番号
    'm_iGyosyu_Kbn  = Request("txtGYOSYU_KBN")          ':業種区分
    m_iSihonkin = Request("txtSIHONKIN")                ':資本金
    m_iJyugyoin_Suu = Request("txtJYUGYOIN_SUU")        ':従業員数
    m_iSyoninkyu    = Request("txtSYONINKYU")           ':初任給
    m_sBiko         = Request("txtBIKO")                ':備考

    
    If strErrmsg <> "" Then
        ' エラーを表示するファンクション
        Call err_page(strErrMsg)
        response.end
    End If
'   call s_viewForm(request.form)   'デバッグ用　引数の内容を見る
End Sub

'********************************************************************************
'*  [機能]  市町村名を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SityosonMei()

        '// 進路区分名を取得
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M12_SITYOSON "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "        M12_KEN_CD = '" & m_iKenCd & "'"
        w_sSQL = w_sSQL & vbCrLf & "    AND M12_SITYOSON_CD = " & m_iSityoCd & " "
        w_sSQL = w_sSQL & vbCrLf & "    GROUP BY M12_SITYOSONMEI "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Sub
        End If

    m_sSityoson = w_Rs("M12_SITYOSONMEI")


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

	Do While not m_Rs.EOF

	w_slink = "　"

	if m_Rs("M32_SINRO_URL") <> "" Then 
	    w_sLink= "<a href='" & gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "'>" 
	    w_sLink= w_sLink &  gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "</a>"
	End if

	        %>
	        <%=w_slink%>
	        <%
	            m_Rs.MoveNext
    Loop

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
    <!--#include file="../../Common/jsCommon.htm"-->
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

    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_GoSyosai(p_sSinroKBN){

        document.frm.action="./syousai.asp";
        document.frm.target="";
        document.frm.txtMode.value = "Syosai";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  戻るボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_BackClick(){

        document.frm.action="./syusei.asp";
        document.frm.target="_self";
        document.frm.txtReFlg.value = "<%=m_bReFlg%>";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_SinkiClick(){

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

	    document.frm.action="update.asp";
	    document.frm.target="_self";
	    document.frm.submit();
    }

    //************************************************************
    //  [機能]  ウインドウオープン時
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function Window_onload(){
    <%
    If m_sMsgFlg = True Then
    %>
    alert("<%=m_sMsg%>")

      document.frm.action="./syusei.asp";
      document.frm.target="_self";
      document.frm.txtMode.value = "Sinki";
      document.frm.submit();

    <%
    End If
    %>
    }
    //-->
    </SCRIPT>

    <link rel=stylesheet href="../../common/style.css" type=text/css>

    </head>

<body onload="Window_onload()">

<center>

<form name="frm" action="update.asp" target="_self" method=post>

<%
If m_sMode = "Sinki" Then
  m_sSubtitle = "新規登録"
else
  m_sSubtitle = "修　正"
End If

call gs_title("進路先情報登録",m_sSubtitle)
%>
<br>
進　路　先　情　報
<br><br>

    <table border="0" class=form width=75%>
    <tr>
    <td class=form align="left" width="100">進路先コード</td>
    <td class=form align="left">
    <%= m_sRenrakusakiCD %>
    <input type="hidden" name="txtRenrakusakiCD" value="<%= m_sRenrakusakiCD %>">
    </td>
    </tr>

    <tr>
    <td class=form align="left">名　称</td>
    <td class=form align="left">
    <%= m_sSINROMEI %>
    <input type="hidden" name="txtSINROMEI" value="<%= m_sSINROMEI %>">
    </td>
    </tr>
    <!--
    <tr>
    <td class=form align="left">名　称（英語）</td>
    <td class=form align="left">
    <%= m_sSINROMEI_EIGO %>
    <input type="hidden" name="txtSINROMEI_EIGO" value="<%= m_sSINROMEI_EIGO %>">
    </td>
    </tr>
    //-->
    <tr>
    <td class=form align="left">名　称（カナ）</td>
    <td class=form align="left">
    <%= m_sSINROMEI_KANA %>
    <input type="hidden" name="txtSINROMEI_KANA" value="<%= m_sSINROMEI_KANA %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">略　称</td>
    <td class=form align="left">
    <%= m_sSINRORYAKSYO %>
    <input type="hidden" name="txtSINRORYAKSYO" value="<%= m_sSINRORYAKSYO %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">郵便番号</td>
    <td class=form align="left">
    <%= m_sYubin %>
    <input type="hidden" name="txtYUBINBANGO" value="<%= m_sYubin %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">住　所（１）</td>
    <td class=form align="left">
    <%= m_sJUSYO1 %>
    <input type="hidden" name="txtJUSYO1" value="<%= m_sJUSYO1 %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">住　所（２）</td>
    <td class=form align="left">
    <%= m_sJUSYO2 %>
    <input type="hidden" name="txtJUSYO2" value="<%= m_sJUSYO2 %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">住　所（３）</td>
    <td class=form align="left">
    <%= m_sJUSYO3 %>
    <input type="hidden" name="txtJUSYO3" value="<%= m_sJUSYO3 %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">電話番号</td>
    <td class=form align="left">
    <%= m_sDENWABANGO %>
    <input type="hidden" name="txtDENWABANGO" value="<%= m_sDENWABANGO %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">進路区分</td>
    <td class=form align="left">

    <%
	'// 進路区分名の確定
	 Call gf_GetKubunName(C_SINRO,m_sSinroCD,m_iNendo,m_sSinrokubun)
	 response.write m_sSinrokubun 
	%>

    <input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD %>">
    </td>
    </tr>

	<tr>
	<%
	'=================================
	'//進路区分により表示を変える
	'=================================
	w_sKbnName = ""
	Select case cint(gf_SetNull2Zero(m_sSinroCD))
		Case C_SINRO_SINGAKU	'//進路区分が進学の場合

			'//進学区分名称を取得
			Call gf_GetKubunName(C_SINGAKU,m_sSingakuCD,m_iNendo,w_sKbnName)

		Case C_SINRO_SYUSYOKU	'//進路区分が就職の場合

			'//業種区分名称を取得
			Call gf_GetKubunName(C_GYOSYU_KBN,m_sSingakuCD,m_iNendo,w_sKbnName)

		Case C_SINRO_SONOTA	'//進路区分がその他の場合
	End Select
	%>

    <td class=form align="left">種別区分</td>
    <td class=form align="left"><%=w_sKbnName%>
    <input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">資本金</td>
    <td class=form align="left">
    <%= m_iSihonkin %>
    <input type="hidden" name="txtSIHONKIN" value="<%= m_iSihonkin %>">万円
    </td>
    </tr>
    <tr>
    <td class=form align="left">従業員数</td>
    <td class=form align="left">
    <%= m_iJyugyoin_Suu %>
    <input type="hidden" name="txtJYUGYOIN_SUU" value="<%= m_iJyugyoin_Suu %>">人
    </td>
    </tr>
    <tr>
    <td class=form align="left">初任給</td>
    <td class=form align="left">
    <%= m_iSyoninkyu %>
    <input type="hidden" name="txtSYONINKYU" value="<%= m_iSyoninkyu %>">円
    </td>
    </tr>
    <tr>
    <td class=form align="left">Ｕ　Ｒ　Ｌ</td>
    <td class=form align="left">
    <%= m_sSINRO_URL %>
    <input type="hidden" name="txtSINRO_URL" value="<%= m_sSINRO_URL %>">
    </td>
    </tr>
    <tr>
    <td class=form align="left">備　考</td>
    <td class=form align="left">
    <%= m_sBiko %>
    <input type="hidden" name="txtBIKO" value="<%= m_sBiko %>">
    </td>
    </tr>
    </table>
<br>
以上の内容で登録します。
<br><br>
<table border="0">
<tr>
<td valign="top">

<input type="button" class=button value="　登　録　" Onclick="f_SinkiClick()">
<img src="../../image/sp.gif" width="20" height="1">
<input type="button" class=button value="キャンセル" Onclick="f_BackClick()">

</td>
</tr>
</table>

<input type="hidden" name="txtMode" value="<%= m_sMode %>">
<input type="hidden" name="txtReFlg" value="<%= m_bReFlg %>">
<input type="hidden" name="txtSinroCD2" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD2" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<!--<input type="hidden" name="txtNendo" value="<%= Session("SYORI_NENDO") %>">-->
<input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
<input type="hidden" name="txtKenCd" value="<%= m_iKenCd %>">
<input type="hidden" name="txtSityoCd" value="<%= m_iSityoCd %>">
</form>


</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
<%
'**********  エラーを表示するファンクション  *********
Function err_page(myErrMsg)
%>
    <html>
    <head>
    <title>項目エラー</title>
    <link rel=stylesheet href=bar.css type=text/css>
    </head>

    <body bgcolor="#ffffff">
    <center>
    <form>
    <font size="2">
    Error:項目エラー<br><br>
    以下の項目のエラーがでています。<br><br>

    <%=myErrMsg%>

    <br><br>
    以上の項目を入力して再度送信してください。<p>
    <input class=button type="button" class=button value="キャンセル" onclick="JavaScript:history.back();">

    </font>

    </form>
    </center>
    </body>
    </html>
<%
End Function
%>