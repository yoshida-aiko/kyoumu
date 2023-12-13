<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験期間教官予定一覧
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0180/skn0180_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:   NENDO           '//:年
'               KYOKAN_CD       '//教官CD
'               cboGakunenCd    '//学年
'               cboClassCd      '//クラス
'               cboSikenKbn     '//試験区分
'               cboSikenCd      '//試験CD
' 説      明:
'           ■初期表示
'               現在の日付に一番近い試験での時間割一覧を表示する
'-------------------------------------------------------------------------
' 作      成: 2001/07/23 伊藤公子
' 変      更: 2001/08/02 根本 直美  '試験コンボ表示変更
'           : 2001/08/08 根本 直美     NN対応に伴うソース変更
'           : 2001/08/09 根本 直美     NN対応に伴うソース変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iSyoriNen          '//教官ｺｰﾄﾞ
    Public m_iKyokanCd          '//年度

    '//コンボ用Where条件等
    Public m_sSikenWhere        '試験の条件
    Public m_sSikenOption       '試験コンボのオプション
    Public m_sSikenCdWhere      '試験コンボのオプション（試験コード）
    Public m_sKyokanName        '//教官名称
    Public m_iSikenKbn          '//試験区分
    Public m_sSikenCd           '//試験CD
    Public m_sTxtMode           '//動作モード

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="試験期間教官予定一覧"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '//値の初期化
        Call s_ClearParam()

        '//変数セット
        Call s_SetParam()

        '//教官名称を取得
        w_iRet = f_GetKyokanNm(m_iKyokanCd,m_iSyoriNen,m_sKyokanName)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//現在の日付に一番近い試験区分を取得
        '//(初期表示は現在の日付に一番近い試験での時間割一覧を表示する)
        If m_sTxtMode = "" Then
            w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_JISSI_KIKAN,0)
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If
        End If

        '//試験コンボに関するWHEREを作成する
        Call s_MakeSikenWhere() 
        
        '//試験コンボ(追試等)に関するWHEREを作成する
        Call s_MakeSikenCdWhere() 

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
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_iSikenKbn = ""
    m_sTxtMode = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")
    m_iSikenKbn = Request("cboSikenKbn")    '//試験区分
    m_sTxtMode  = Request("txtMode")

    '//教官CD
    If Request("SKyokanCd1") = "" Then
        m_iKyokanCd = Session("KYOKAN_CD")
    Else
        m_iKyokanCd = Request("SKyokanCd1")
    End If

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_iSikenKbn = " & m_iSikenKbn & "<br>"
    response.write "m_sTxtMode  = " & m_sTxtMode  & "<br>"

End Sub

'********************************************************************************
'*  [機能]  試験区分コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeSikenWhere()

    m_sSikenWhere = ""
    m_sSikenWhere = m_sSikenWhere & " M01_NENDO = " & m_iSyorinen
    m_sSikenWhere = m_sSikenWhere & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
    m_sSikenWhere = m_sSikenWhere & " AND M01_SYOBUNRUI_CD <= 4 "						'<!--8/16修正

End Sub

'********************************************************************************
'*  [機能]  試験ｺｰﾄﾞコンボに関するWHEREを作成する（試験コード）
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeSikenCdWhere()

    m_sSikenCdWhere = ""
    m_sSikenOption = ""

    If cint(m_iSikenKbn) = Cint(C_SIKEN_JITURYOKU) or cint(m_iSikenKbn) = cInt(C_SIKEN_TUISI)  Then
        m_sSikenCdWhere = m_sSikenCdWhere & " M27_NENDO = " & m_iSyoriNen
        m_sSikenCdWhere = m_sSikenCdWhere & " AND M27_SIKEN_KBN = " & m_iSikenKbn

    else
        m_sSikenOption = " DISABLED "

    End If
End Sub

'********************************************************************************
'*  [機能]  現在の日付に一番近い試験区分を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  初期表示は現在の日付に一番近い試験での時間割一覧を表示する
'********************************************************************************
Function f_Get_SikenKbn(p_iSiken_Kbn,p_sSiken_CD)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_Get_SikenKbn = 1
    p_iSiken_Kbn = ""
    p_sSiken_CD  = ""

    Do
        '現在の日付に一番近い試験区分を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "    T24_SIKEN_KBN,"
        w_sSQL = w_sSQL & "    T24_SIKEN_CD"
        w_sSQL = w_sSQL & " FROM T24_SIKEN_NITTEI"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "       T24_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & "   AND T24_JISSI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"
        w_sSQL = w_sSQL & " ORDER BY T24_JISSI_SYURYO ASC"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_SikenKbn = 99
            Exit Do
        End If

        If rs.EOF = False Then
            p_iSiken_Kbn = rs("T24_SIKEN_KBN")
            p_sSiken_CD  = rs("T24_SIKEN_CD")
        End If

        '//正常終了
        f_Get_SikenKbn = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  教官の氏名を取得
'*  [引数]  なし
'*  [戻値]  p_sName
'*  [説明]  
'********************************************************************************
Function f_GetKyokanNm(p_sCD,p_iNENDO,p_sName)
Dim rs
Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetKyokanNm = 1
    w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M04_KYOKAN_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M04_NENDO = " & p_iNENDO & " "

'response.write w_sSQL

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKyokanNm = 99
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M04_KYOKANMEI_SEI") & "　" & rs("M04_KYOKANMEI_MEI")
        End If

        f_GetKyokanNm = 0
        Exit Do
    Loop

    p_sName = w_sName

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
    <title>試験期間教官予定一覧</title>

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
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        if(document.frm.SKyokanNm1.value==""){
			alert("教官を選択してください")
			return;
		}

        document.frm.action="./skn0180_main.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  試験コンボが変更されたとき、本画面を再表示
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./skn0180_top.asp";
        document.frm.target="top";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  クリアボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function fj_Clear(){

		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}

    //************************************************************
    //  [機能]  教官参照選択画面ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function KyokanWin(p_iInt,p_sKNm) {
		var obj=eval("document.frm."+p_sKNm)

        URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=610,top=0,left=0");
        nWin.focus();
        return true;    
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">

<%
'//デバッグ
'Call s_DebugPrint()
%>

    <center>
    <%call gs_title("試験期間教官予定一覧","一　覧")%>
<br>
    <table>
        <tr>
        <td class="search">
            <table border="0">
            <tr>
            <td>
                <table border="0" cellpadding="1" cellspacing="1">
                <tr valign="middle">
                <td align="left" nowrap>
                試験
                </td>
                <td align="left" nowrap>
                    <% call gf_ComboSet("cboSikenKbn",C_CBO_M01_KUBUN,m_sSikenWhere,  "onchange = 'javascript:f_ReLoadMyPage()' style='width:150px;' ",False,m_iSikenKbn) %>
                </td>
<!--
                <td align="left" nowrap>
                    <% call gf_ComboSet("cboSikenCd" ,C_CBO_M27_SIKEN,m_sSikenCdWhere,m_sSikenOption & " style='width:120px;' ",True,"") %>
                </td>
//-->
                </tr>
                <tr valign="middle">
                <td align="left" nowrap>
                教官
                </td>
                <td align="left" nowrap colspan="2">
                    <input type="text" class="text" name="SKyokanNm1" VALUE='<%=m_sKyokanName%>' readonly>
                    <input type="hidden" name="SKyokanCd1" VALUE='<%=m_iKyokanCd%>'>
                    <input type="button" class="button" value="選択" onclick="KyokanWin(1,'SKyokanNm1')">
					<input type="button" class="button" value="クリア" onClick="fj_Clear()">
                    <!--<input type="button" class="button" value="選択" onclick="KyokanWin(1,'<%=m_sKyokanName%>')">-->
                </td>
                </tr>
                </table>
            </td>
            </tr>
			<tr>
		        <td valign="bottom" clspan="1" align="right">
		        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
		        </td>
			</tr>
            </table>
        </td>
        </tr>
    </table>

    </center>

    <!--値渡し用-->
    <INPUT TYPE="HIDDEN" NAME="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
    </form>
    </body>
    </html>
<%
    '---------- HTML END   ----------
End Sub
%>
