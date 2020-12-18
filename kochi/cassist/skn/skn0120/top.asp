<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験期間教官予定登録
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0120/top.asp
' 機      能: 上ページ 予定登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSikenKbn      :試験区分
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSikenKbn      :試験区分
'           txtSikenCd      :試験コード（実力・追試験//A:1,B:2）
'           txtMode         :動作モード
'                               BLANK   :初期表示
'                               Reroad  :（条件選択後）再表示
'                               Search  :検索
' 説      明:
'           ■初期表示
'               コンボボックスは試験名称を表示
'           ■表示ボタンクリック時
'               下のフレームに指定した試験条件にかなう教官予定を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/18 高丘 知央
' 変      更: 2001/06/26 根本 直美
'           : 2001/08/02 根本 直美  '試験コンボ表示変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '試験選択用のWhere条件
    Public m_sSikenWhere        '試験の条件
    Public m_sSikenOption       '試験コンボのオプション
    Public  m_sSikenCdWhere     '試験コンボのオプション（試験コード）
    
    '取得したデータを持つ変数
    Public  m_iSikenKbn      ':試験区分
    Public  m_iSyoriNen      ':処理年度
    Public  m_iKyokanCd      ':教官コード

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
    w_sMsgTitle="試験監督免除登録"
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

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '//現在の日付に一番近い試験区分を取得
        '//(初期表示は現在の日付に一番近い試験での時間割一覧を表示する)
        If m_sTxtMode = "" Then
            w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_JISSI_KIKAN,0)
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If
        End If
        
        '試験コンボに関するWHEREを作成する
        Call s_MakeSikenWhere() 
        
        '試験コンボに関するWHEREを作成する
        Call s_MakeSikenCdWhere() 
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

    m_iSikenKbn = ""
    m_iSikenKbn = Request("txtSikenKbn")

    'Session("NENDO")= 2000  '//テスト用
    m_iSyoriNen = Session("NENDO")
'    m_iSyoriNen = 2001     '//テスト用
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub


Sub s_MakeSikenWhere()
'********************************************************************************
'*  [機能]  試験コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    
    m_sSikenWhere = ""
    
	m_sSikenWhere = m_sSikenWhere & " M01_NENDO = " & m_iSyorinen
'    m_sSikenWhere = m_sSikenWhere & " M01_NENDO = " & 2000  '//テスト用

'--2001/07/15 CONSTに変更
    'm_sSikenWhere = m_sSikenWhere & " AND M01_DAIBUNRUI_CD = 13 "
    m_sSikenWhere = m_sSikenWhere & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
    m_sSikenWhere = m_sSikenWhere & " AND M01_SYOBUNRUI_CD <= 4 "						'<!--8/16修正
 
End Sub

Sub s_MakeSikenCdWhere()
'********************************************************************************
'*  [機能]  試験コンボに関するWHEREを作成する（試験コード）
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************


    m_sSikenCdWhere = ""
    m_sSikenOption = ""
    
'--2001/07/15 CONSTに変更
        'C_SIKEN_JITURYOKU = 5  '実力試験
        'C_SIKEN_TUISI = 6      '追試験

    'If m_iSikenKbn = 5 or m_iSikenKbn = 6  Then
    If cint(m_iSikenKbn) = Cint(C_SIKEN_JITURYOKU) or cint(m_iSikenKbn) = cInt(C_SIKEN_TUISI)  Then
        'm_sSikenCdWhere = m_sSikenCdWhere & " M27_NENDO = " & m_iSyoriNen
        m_sSikenCdWhere = m_sSikenCdWhere & " M27_NENDO = " & m_iSyoriNen
        m_sSikenCdWhere = m_sSikenCdWhere & " AND M27_SIKEN_KBN = " & m_iSikenKbn
    else
        m_sSikenOption = " DISABLED "
    End If
    
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
<link rel=stylesheet href=../../common/style.css type=text/css>
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

        document.frm.action="../../menu/siken.asp";
        document.frm.target="_parent";
        document.frm.submit();
    
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
        document.frm.txtMode.value = "Search";
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

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Touroku(){

        document.frm.action = "syusei.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "Sinki";
        document.frm.submit();
    
    }
    //************************************************************
    //  [機能]  登録画面を表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_NewClick(){
    
        document.frm.action="syousai.asp";
        document.frm.target = "<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "BLANK";
        document.frm.submit();
        
    }


    //-->
    </SCRIPT>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" action="" target="main" Method="POST">
<input type="hidden" name="txtMode">

<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
<%call gs_title("試験監督免除申請登録","一　覧")%>
<br>
<%
If m_sMode = "" Then
%>
    <table border="0" cellpadding="0" cellspacing="0" align="center">
    <tr>
    <td class=search >
                <table border="0" cellpadding="1" cellspacing="1">
                <tr>
                <td align="left" >
                <% call gf_ComboSet("txtSikenKbn",C_CBO_M01_KUBUN,m_sSikenWhere," style='width:150px;' ",False,m_iSikenKbn) %>
                </td>
				<td valign="bottom" align="right" colspan="1"><input class="button" type="button" onclick="javascript:f_Search();" value="　表　示　"></td>
<!--
                <td align="left" >&nbsp;&nbsp;
                <% call gf_ComboSet("txtSikenCd",C_CBO_M27_SIKEN,m_sSikenCdWhere,m_sSikenOption & " style='width:120px;' ",true,"") %>
                </td>
//-->
                </tr>
                </table>
    </td>
    <td valign="top" align="right" rowspan="2">
        <table border="0" cellpadding="0" cellspacing="0">
        <tr>
        <td align="right" class="form" valign="top">
        <a href="#" onClick = "javascript:f_NewClick()">新規登録はこちら</a><br><img src="../../image/sp.gif" height="10"><br>
        </td>
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
