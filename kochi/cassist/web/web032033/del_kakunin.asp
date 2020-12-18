<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 使用教科書登録　削除確認画面
' ﾌﾟﾛｸﾞﾗﾑID : web/WEB0320/del_kakunin.asp
' 機      能: 下ページ 使用教科書登録の削除確認画面を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           txtSinroKBN     :進路先コード
'           txtSingakuCd        :進学コード
'           txtSinroName        :就職先名称（一部）
'           txtPageCD       :表示済表示頁数（自分自身から受け取る引数）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           txtRenrakusakiCD    :選択された連絡先コード
'           txtPageCD       :表示済表示頁数（自分自身に引き渡す引数）
' 説      明:
'           ■初期表示
'               検索条件にかなう就職・進学先を表示
'           ■次へ、戻るボタンクリック時
'               指定した条件にかなう就職・進学を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/14 岩下　幸一郎
' 変      更: 2001/08/01 前田　智史
' 変      更: 2001/08/22 伊藤 公子 教官を選択できるように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_sNendo         '// 年度
    Public  m_sKyokan_CD     ':教官コード
    Public  m_sMode          ':モード
'    Public  w_sDelKyokasyoCD
    Public  m_Rs             'recordset
'    Public  w_iDisp          ':表示件数の最大値をとる
    Public  m_sPageCD        ':ページ数
    Public  m_sNo            ':defaultの削除にチェックいれたものの配列
    Public  m_sGakka         '学科名称

    'ページ関係
    Public  m_cell
    Public  m_iMax      ':最大ページ
    Public  m_iDsp      '// 一覧表示行数

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
    w_sMsgTitle="使用教科書登録　削除確認画面"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"


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

		'//削除するデータ取得
		w_iRet = f_syousaiitiran()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
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

    'm_sNendo        = request("txtNendo")
    m_sNendo        = request("KeyNendo")
    'm_sKyokan_CD    = request("txtKyokanCd")
    m_sKyokan_CD    = request("SKyokanCd1")

    m_sMode         = Request("txtMode")
    m_sPageCD       = Request("txtPageCD")
	m_sNo           = request("deleteNO")
'    w_iDisp  = Request("txtDisp")           ':ページ最大値

End Sub


Function f_syousaiitiran()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_syousaiitiran = 1

    Do

	    w_sSQL = w_sSQL & vbCrLf & " SELECT "
	    w_sSQL = w_sSQL & vbCrLf & "  T47.T47_GAKUNEN "         ''学年
	    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKA_CD "        ''学科
	    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKASYO "        ''教科書名
	    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SYUPPANSYA "      ''出版社
	    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_TYOSYA "          ''著者
	    w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKAMEI "
	    w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI "
	    w_sSQL = w_sSQL & vbCrLf & " FROM "
	    w_sSQL = w_sSQL & vbCrLf & "    T47_KYOKASYO T47 "
	    w_sSQL = w_sSQL & vbCrLf & "    ,M02_GAKKA M02 "
	    w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
	    w_sSQL = w_sSQL & vbCrLf & "    ,M04_KYOKAN M04 "
	    w_sSQL = w_sSQL & vbCrLf & " WHERE "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NO IN (" & Trim(m_sNo) & ") AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M02.M02_NENDO(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_GAKKA_CD  = M02.M02_GAKKA_CD(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M03.M03_NENDO(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M04.M04_NENDO(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = M04.M04_KYOKAN_CD(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO = " & m_sNendo
	    'w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = '" & m_sKyokan_CD & "' "
	    w_sSQL = w_sSQL & vbCrLf & " ORDER BY T47.T47_GAKKA_CD "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If

        f_syousaiitiran = 0

        Exit Do
    Loop
    
    'LABEL_showPage_OPTION_END
End Function

Sub S_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	Do Until m_Rs.EOF
		Call gs_cellPtn(m_cell)
        %>
        <tr>
	        <td align="center" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_GAKUNEN")) %>年</td>
        <%
        if cstr(gf_HTMLTableSTR(m_Rs("T47_GAKKA_CD"))) = cstr(C_CLASS_ALL) then
            m_sGakka="全学科"
        else
            m_sGakka=gf_HTMLTableSTR(m_Rs("M02_GAKKAMEI"))
        end if
        %>
	        <td align="left" class=<%=m_cell%>><%=m_sGakka %></td>
	        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M03_KAMOKUMEI")) %></td>
	        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_KYOKASYO")) %></td>
	        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_SYUPPANSYA")) %></td>
        </tr>

        <%
    m_Rs.MoveNext
	Loop
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

    On Error Resume Next
    Err.Clear
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
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  syosai_frmへのパラメータの受け渡し
    //  [引数]  p_sSyuseiCD
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Henko(p_sSyuseiCD){

        document.frm.action="syusei.asp";
        document.frm.target="";
        document.frm.txtRenrakusakiCD.value = p_sSyuseiCD;
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    }

    //************************************************************
    //  [機能]  削除ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_delete(){

        if (!confirm("<%=C_SAKUJYO_KAKUNIN%>")) {
           return ;
        }

        document.frm.action="./delete.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "Delete";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Back(){
        document.frm.action="./default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "Back";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
</head>
<body>

<center>

<%

If m_sMode = "DELETE" Then
  m_sSubtitle = "削　除"
End If

call gs_title("使用教科書登録",m_sSubtitle)
%>
<br>
使　用　教　科　書
<br><br>
<form name="frm" action="" target="" method="post">
<table border="1" class=hyo width="75%">
    <tr>
	    <th class=header>学年</th>
	    <th class=header>学科</th>
	    <th class=header>科目</th>
	    <th class=header>教科書名</th>
	    <th class=header>出版社</th>
    </tr>

    <% S_syousai() %>

</table>
<br>
以上の内容を削除します。
<br><br>
<table border="0" width="75%">
	<tr>
		<td align=center colspan=5>
		<input type="button" class=button value="　削　除　" onclick="f_delete()">
		<input type="button" class=button value="キャンセル" onclick="f_Back()">
		</td>
	</tr>
</table>
	<input type="hidden" name="txtMode" value="">
	<input type="hidden" name="txtDelKyokasyoCD" value="<%= w_sDelKyokasyoCD %>">
	<input type="hidden" name="txtNendo" value="<%= m_sNendo %>">
	<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
	<input type="hidden" name="txtDisp" value="<%= w_iDisp %>">
    <input type="hidden" name="txtNo" value="<%=m_sNo%>">

    <input type="hidden" name="SKyokanCd1" value="<%=m_sKyokan_CD%>">

</form>

</center>

</body>

</html>





<%
    '---------- HTML END   ----------
End Sub

Sub NoDataPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
    <html>
    <head>
    </head>

    <body>

    <center>
        削除の対象となるデータが選択されていません。<br><br><br>
    <input type="button" class=button value="戻　る" onclick="javascript:history.back()">
    </center>

    </body>

    </html>
<%
End Sub
%>