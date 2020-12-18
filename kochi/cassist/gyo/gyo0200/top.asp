<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 行事日程一覧
' ﾌﾟﾛｸﾞﾗﾑID : gyo/gyo0200/top.asp
' 機      能: 上ページ 行事日程の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           cboGyojiDate      :行事日付
'           chkGyojiCd      :行事コード
' 説      明:
'           ■初期表示
'               コンボボックスは月を表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう行事一覧を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/26 根本 直美
' 変      更: 2001/07/27 伊藤公子　M40_CALENDERテーブル削除に対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bErrMsg           'ｴﾗｰﾒｯｾｰｼﾞ
    
    '月選択用のWhere条件
    Public m_sGyojiMWhere        '月の条件
    
    '取得したデータを持つ変数
    Public  m_iSyoriNen      ':処理年度
    Public  m_iKyokanCd      ':教官コード
    
    'データ取得用
    Public  m_iDate             ':今日の日付(yyyy/mm/dd)
    Public  m_iDay              ':今日の日

    Public  m_iTuki		'//当月

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
    w_sMsgTitle="行事日程一覧"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False


        '// ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()

        '// 日付SET
		Call s_SetDate()
         

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_bErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
        
        '月コンボに関するWHEREを作成する
        'Call s_MakeGyojiMWhere() 
        
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

Sub s_SetParam()
'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub

Sub s_SetDate()
'********************************************************************************
'*  [機能]  今日の日付を設定（コンボボックス用）
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

'    m_iDate = gf_YYYY_MM_DD(date(),"/")
'    m_iDay = Day(m_iDate)

	m_iTuki = month(date())

End Sub

'****************************************************
'[機能]	データ1とデータ2が同じ時は "SELECTED" を返す
'		(リストダウンボックス選択表示用)
'[引数]	pData1 : データ１
'		pData2 : データ２
'[戻値]	f_Selected : "SELECTED" OR ""
'					
'****************************************************
Function f_Selected(pData1,pData2)

	If IsNull(pData1) = False And IsNull(pData2) = False Then
		If trim(cStr(pData1)) = trim(cstr(pData2)) Then
			f_Selected = "selected"	
		Else 
			f_Selected = ""	
		End If
	End If

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
<link rel=stylesheet href="../../common/style.css" type=text/css>
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

        document.frm.action="../../menu/sansyo.asp";
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
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" action="./main.asp" target="main" Method="POST">
    <input type="hidden" name="txtMode">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
    <%call gs_title("行事日程一覧","一　覧")%>
<%
If m_sMode = "" Then
%>
    <table border="0">
    <tr>
    <td class=search>
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left" >
            <table border="0" cellpadding="1" cellspacing="1">
            <tr>
            <td align="left" >
            <% 'call gf_calComboSet("cboGyojiDate",C_CBO_M40_CALENDER,m_sGyojiMWhere,"",False,m_iDate,"mm") %>
			<select name="cboGyojiDate">
				<option value="4"  <%=f_Selected("4" ,cstr(m_iTuki))%> >4
				<option value="5"  <%=f_Selected("5" ,cstr(m_iTuki))%> >5
				<option value="6"  <%=f_Selected("6" ,cstr(m_iTuki))%> >6
				<option value="7"  <%=f_Selected("7" ,cstr(m_iTuki))%> >7
				<option value="8"  <%=f_Selected("8" ,cstr(m_iTuki))%> >8
				<option value="9"  <%=f_Selected("9" ,cstr(m_iTuki))%> >9
				<option value="10" <%=f_Selected("10",cstr(m_iTuki))%> >10
				<option value="11" <%=f_Selected("11",cstr(m_iTuki))%> >11
				<option value="12" <%=f_Selected("12",cstr(m_iTuki))%> >12
				<option value="1"  <%=f_Selected("1" ,cstr(m_iTuki))%> >1
				<option value="2"  <%=f_Selected("2" ,cstr(m_iTuki))%> >2
				<option value="3"  <%=f_Selected("3" ,cstr(m_iTuki))%> >3
			</select>月
            </td>
            <td align="left" >&nbsp;&nbsp;&nbsp;<input type="checkbox" name="chkGyojiCd">行事のみ表示</td>
            </tr>
            </table>
        </td>
        </tr>
        </table>
    </td>
    <td valign="bottom">
    <input type="button" value="　表　示　" onClick="javascript:f_Search()" class=button>
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
