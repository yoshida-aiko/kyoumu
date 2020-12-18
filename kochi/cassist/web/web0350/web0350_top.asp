<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 空き時間情報検索
' ﾌﾟﾛｸﾞﾗﾑID : web/web0350/web0350_top.asp
' 機      能: 検索ページ	 空き時間情報検索を行う
'-------------------------------------------------------------------------
' 引      数:
' 変      数:
' 引      渡:
' 説      明:
'           
'-------------------------------------------------------------------------
' 作      成: 2001/08/17 持永
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  s_sGakkaWhere		'学科のWHERE文
    Public  m_iJMax				'時限最大数

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
    w_sMsgTitle="空き時間情報検索"
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

        '学科コンボに関するWHEREを作成する
        Call s_MakeGakkaWhere() 

        '//最大時限数を取得
        Call gf_GetJigenMax(m_iJMax)
		if m_iJMax = "" Then
		    m_bErrFlg = True
		    Exit Do
		end if

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
'*  [機能]  学科コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeGakkaWhere()
    
    s_sGakkaWhere = ""
    s_sGakkaWhere = m_sGakkaWhere & " M02_NENDO = " & Session("NENDO")  '//表示年度

End Sub

Function f_JigenCbo(p_name)
	Dim i,w_val,w_iRet
	i=0::w_val = "":w_iRet = 0
	f_JigenCbo = ""

	w_iRet = gf_GetJigenMax(w_iJMax)

	If w_iRet <> 0 then 
		f_JigenCbo = f_JigenCbo & vbCrLf & "<SELECT name='" & p_name & "' disabled>"
			f_JigenCbo = f_JigenCbo & vbCrLf & "<option></option>"
			f_JigenCbo = f_JigenCbo & vbCrLf & "</SELECT>"
	Else
			f_JigenCbo = f_JigenCbo & vbCrLf & "<SELECT name='" & p_name & "'>"
		For i = 1 to cint(w_iJMax)
			w_val = right("  "&i,2)
			f_JigenCbo = f_JigenCbo & vbCrLf & "<option value='" & i & "'>" & w_val & "</option>"
		Next
			f_JigenCbo = f_JigenCbo & vbCrLf & "</SELECT>"
	End If
	

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    On Error Resume Next
    Err.Clear

	'//今日の日付を取得する
	w_Date = gf_YYYY_MM_DD(date(),"/")

%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

		// 日付型チェック
		if (IsDate(document.frm.txtDay.value) == 1 ){
			window.alert("日付の入力に誤りがあります");
			document.frm.txtDay.focus();
			return ;
		}

		// 開始時限時限指定
		strSt = document.frm.txtJigenSt.selectedIndex;
		strEd = document.frm.txtJigenEd.selectedIndex;
			if (strSt > strEd){
			    alert("開始時限より大きい値を選択して下さい");
				document.frm.txtJigenEd.focus();
				return ;
			}
		
        document.frm.action="web0350_main.asp";
        document.frm.target="main";
        document.frm.submit();
        
    }

    //-->
    </SCRIPT>

    </head>
    <body>
    <%call gs_title("空き時間情報検索","一　覧")%>

	<div align="center">

	<table border="0">
        <tr>
	        <td class="search">

				<table border="0">
    <form name="frm" method="post">
					<tr>
						<td nowrap>日　付</td>
						<td nowrap><input type="text" name="txtDay" value="<%=w_Date%>"></td>
						<td nowrap><input type="button" value="選択" onclick="fcalender('txtDay')"></td>
						<td nowrap>空き時限</td>
						<td nowrap><%=f_JigenCbo("txtJigenSt")%>限から</td>
						<td nowrap><%=f_JigenCbo("txtJigenEd")%>限の間</td>
<!--
						<td nowrap><input type="text" name="txtJigenSt" size="3">限から</td>
						<td nowrap><input type="text" name="txtJigenEd" size="3">限の間</td>
-->
					</tr>
					<tr>
						<td nowrap>学　科</td>
						<td nowrap><% call gf_ComboSet("txtGakka",C_CBO_M02_GAKKA,s_sGakkaWhere," style='width:115px;'",True,m_sGakkaCD) %></td>
						<td nowrap colspan="4" align="right"><input class="button" type="reset" value=" ク　リ　ア ">
							<input class="button" type="button" value="　表　示　" onClick="javascript:f_Search();"></td>
					</tr>
    </form>
				</table>

	        </td>
		</tr>
	</table>
	<BR>
<table><tr><td>
<span class="msg"><font size="2">
	※ 調べたい空き時間を入力して、「表示」を押してください<BR>
	※ その間すべてに、空き時間のある教官の一覧が表示されます
</font></span>
</td></tr></table>

    </div>
    </body>
    </html>
<%
End Sub
%>