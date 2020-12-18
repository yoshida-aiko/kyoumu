<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 調査書所見等登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0460/gak0460_top.asp
' 機      能: 上ページ 調査書所見等登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/18 前田 智史
' 変      更: 2001/08/07 根本 直美     NN対応に伴うソース変更
' 変      更：2001/08/30 伊藤 公子     検索条件を2重に表示しないように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '市町村選択用のWhere条件
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官コード
    Public m_sNendo         '年度コンボボックスに入る値
    Public m_sGakuNo        '氏名コンボボックスに入る値
    Public m_sNendoWhere    '年度コンボボックスの条件
    Public m_sGakuNoWhere   '氏名コンボボックスの条件
    Public m_sOption        '氏名コンボボックスの使用可、不可の判別

    Public m_sGakunen        
    Public m_sClass          
    Public m_sClassNm        

    Public  m_Rs
    Public  m_iMax          '最大ページ
    Public  m_iDsp          '一覧表示行数

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
    w_sMsgTitle="調査書所見等登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
	m_sNendo = request("txtNendo")
    m_sGakuNo   = request("txtGakuNo")

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

        w_iRet = f_NendoWhere()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

        Call f_GakuNoWhere()
        
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

Function f_NendoWhere()
'********************************************************************************
'*  [機能]  設定クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_NendoWhere = 1

    Do

        m_sNendoWhere=""
            m_sNendoWhere = " M05_NENDO > " & m_iNendo - 5 & "  AND "
            m_sNendoWhere = m_sNendoWhere & " M05_NENDO <= " & m_iNendo & "  AND "
            m_sNendoWhere = m_sNendoWhere & " M05_TANNIN = '" & m_sKyokanCd & "' "

            m_sNendo = request("txtNendo")

        If request("txtNendo") = C_CBO_NULL Then m_sNendo = ""

        If m_sNendo <> "" Then
            w_sSQL = ""
            w_sSQL = w_sSQL & " SELECT "
            w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
            w_sSQL = w_sSQL & " FROM "
            w_sSQL = w_sSQL & "     M05_CLASS "
            w_sSQL = w_sSQL & " WHERE"
            w_sSQL = w_sSQL & "     M05_NENDO = '" & m_sNendo & "' "
            w_sSQL = w_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "

            Set m_Rs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
            If w_iRet <> 0 Then
                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                f_NendoWhere = 99
                m_bErrFlg = True
                Exit Do 
            End If

			m_sGakunen	= m_Rs("M05_GAKUNEN")
			m_sClass	= m_Rs("M05_CLASSNO")
			m_sClassNm	= m_Rs("M05_CLASSMEI")

        End If

        f_NendoWhere = 0
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub f_GakuNoWhere()
'********************************************************************************
'*  [機能]  氏名コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sGakuNoWhere=""
    m_sOption=""

    If m_sNendo <> "" Then
        If m_Rs.EOF Then
            m_sOption = " DISABLED "
            m_sGakuNoWhere  = " T11_GAKUSEI_NO = '' "
        Else
            m_sGakuNoWhere = " T11.T11_GAKUSEI_NO = T13.T13_GAKUSEI_NO AND "
'            m_sGakuNoWhere = m_sGakuNoWhere & " T11.T11_NYUNENDO = T13.T13_NENDO - T13.T13_GAKUNEN + 1 AND "
            m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_GAKUNEN = " & m_sGakunen & " AND "
            m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_CLASS = " & m_sClass & " AND "
            m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_NENDO = " & m_sNendo & " "
        End If
    Else
        m_sOption = " DISABLED "
        m_sGakuNoWhere  = " T11_GAKUSEI_NO = '' "
    End IF

End Sub

Sub f_Syosai()
'********************************************************************************
'*  [機能]  氏名コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    If m_sNendo = "" Then
%>
        <td width="48" Nowrap>　</td>
        <td width="96" Nowrap>　</td>
<%
    Else
%>
        <td width="48" Nowrap align="right"><%=m_sGakunen%>年</td>
        <td width="96" Nowrap><%=m_sClassNm%></td>
<%
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
    <title>調査書所見等登録</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [機能]  年度が修正されたとき、再表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="gak0461_top.asp";
        document.frm.target="topFrame";
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

        // ■■■NULLﾁｪｯｸ■■■
        // ■年度
        if( f_Trim(document.frm.txtNendo.value) == "" ){
            window.alert("年度の選択を行ってください");
            document.frm.txtNendo.focus();
            return ;
        }
        // ■年度
        if( f_Trim(document.frm.txtNendo.value) == "<%=C_CBO_NULL%>" ){
            window.alert("年度の選択を行ってください");
            document.frm.txtNendo.focus();
            return ;
        }
        // ■学生
        if( f_Trim(document.frm.txtGakuNo.value) == "" ){
			if (document.frm.txtGakuNo.length == 1) {
	            window.alert("指定年度の学生のデータがありません");
	            document.frm.txtNendo.focus();
			} else {
	            window.alert("学生の選択を行ってください");
    	        document.frm.txtGakuNo.focus();
			}
            return ;
        }

        // ■学生
        if( f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
			if (document.frm.txtGakuNo.length == 1) {
	            window.alert("指定年度の学生のデータがありません");
	            document.frm.txtNendo.focus();
			} else {
	            window.alert("学生の選択を行ってください");
    	        document.frm.txtGakuNo.focus();
			}
        	    return ;
        }

        document.frm.action="gak0461_main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }
    //************************************************************
    //  [機能]  クリアボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Clear(){

        document.frm.txtNendo.value = "";
        document.frm.txtGakuNo.value = "";
    
    }

    //-->
    </SCRIPT>

    <link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>
<body>
<center>
<form name="frm" METHOD="post">
<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
<td valign="top" align="center">
<%call gs_title("調査書所見等登録","登　録")%>
<br>
    <table border="0">
    <tr>
    <td class="search" Nowrap>
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td Nowrap>
        <%call gf_ComboSet("txtNendo",C_CBO_M05_CLASS_N,m_sNendoWhere,"style='width:70px;' onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sNendo)%></td><td>年度</td>
        <%Call f_Syosai()%>
        <td Nowrap>　氏　名　</td>
		<td Nowrap>
        <%call gf_PluComboSet("txtGakuNo",C_CBO_T11_GAKUSEKI_N,m_sGakuNoWhere,"style='width:250px;' "& m_sOption,True,m_sGakuNo)%>
        </td>
        </tr>
		<tr>
	        <td colspan="6" align="right">
	        <input type="button" class="button" value=" ク　リ　ア " onclick="javasript:f_Clear();">
	        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
	        </td>
		</tr>
        </table>
    </td>
    </tr>
    </table>
</td>
</tr>
</table>
	<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
</form>
</center>
</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>
