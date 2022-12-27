<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0133/main.asp
' 機      能: 下ページ 就職先マスタの一覧リスト表示を行う
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
' 作      成: 2001/06/27 岩下　幸一郎
' 変      更: 2001/07/13 谷脇　良也
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_sSinroCD      ':進路先コード
    Public  m_sSingakuCd        ':進学コード
    Public  m_sSinroCD2     ':進路先コード
    Public  m_sSingakuCd2       ':進学コード
    Public  m_sSyusyokuName     ':就職先名称（一部）
    Public  m_sPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_skubun
    Public  m_Rs            'recordset
    Public  w_iDisp         ':表示件数の最大値をとる
    Public  m_sRenrakusakiCD
    Public  w_i
    w_i     = 1
    Public  w_iThisPgCnt
    Public  w_iSinrosakiCD
    Public  m_sSinroName
    Public  m_iNendo        ':年度
    Public  m_sMode

    'ページ関係
    Public  m_cell
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
    w_sMsgTitle="就職マスタ"
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
        w_sSQL = w_sSQL & vbCrLf & " M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_NENDO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & "    ,M01_KUBUN M01 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M01_NENDO = " & m_iNendo & " AND "
        w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo & " AND "
If m_sSinroCD <> 1 Then
        w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD (+) = "&C_SINRO&""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN = M01.M01_SYOBUNRUI_CD (+)"
Else
        w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD (+) = "&C_SINGAKU&""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN = M01.M01_SYOBUNRUI_CD (+)"
End If

        '抽出条件の作成
        If m_sSinroCD<>"" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN =" & m_sSinroCD & " "
        End If
        If m_sSingakuCd<>"" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN =" & m_sSingakuCd & " "
        End If
        If m_sSinroName<>"" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINROMEI Like '%" & m_sSinroName & "%' "
        End If

        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M32.M32_SINRO_CD "

'Response.Write w_sSQL & "<br>"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        Else
            'ページ数の取得
            m_iMax = gf_PageCount(m_Rs,m_iDsp)
'Response.Write "m_iMax:" & m_iMax & "<br>"
        End If

'If m_sRenrakusakiCD = "" Then
'   Call NoDataPage()
'Exit Sub
'End If

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

    m_sMode      = Request("txtMode")
    m_sRenrakusakiCD = Request("txtRenrakusakiCD")  ':連絡先コード

    m_sSinroCD2 = Request("txtSinroCD2")        ':進路コード
    'コンボ未選択時
    If m_sSinroCD2="@@@" Then
        m_sSinroCD2=""
    End If

'response.write m_sSinroCD2

    m_sSingakuCD2 = Request("txtSingakuCD2")    ':進学コード
    'コンボ未選択時
    If m_sSingakuCD2="@@@" Then
        m_sSingakuCD2=""
    End If

    m_sSyusyokuName = Request("txtSyusyokuName")    ':就職先名称（一部）

    m_sSinroName = Request("txtSinroName")      ':就職先名称（一部）

    '// BLANKの場合は行数ｸﾘｱ
    If Request("txtMode") = "Delete" Then
        m_sPageCD = 1
    Else
        m_sPageCD = INT(Request("txtPageCD"))   ':表示済表示頁数（自分自身から受け取る引数）
    End If

    If m_sSinroCD = "1" Then            ':ヘッダーの区分名称変更
        m_skubun = "進学区分"
    else
        m_skubun = "進路区分"
    End If

    m_iNendo = Session("NENDO")         ':年度

    w_iDisp  = Request("txtDisp")           ':ページ最大値

End Sub


Sub S_syousaiitiran()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_slink
Dim w_iCnt
Dim i

w_iThisPgCnt = 0
w_slink = "　"

w_iCnt = 0


For i = 1 to w_iDisp


w_iSinrosakiCD = ""
w_sSQL = ""

If Request("deleteNO" & i) <> "" Then

w_iSinrosakiCD = Request("deleteNO" & i)

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWHERE            '// WHERE文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="就職マスタ"
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

        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_NENDO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & "    ,M01_KUBUN M01 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo & " AND "
If m_sSinroCD <> 1 Then
        w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD (+) = "&C_SINRO&""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN = M01.M01_SYOBUNRUI_CD (+)"
Else
        w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD (+) = "&C_SINGAKU&""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN = M01.M01_SYOBUNRUI_CD (+)"
End If

'response.write w_sSQL

        '抽出条件の作成
        If m_sSinroCD <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN =" & m_sSinroCD & " "
        End If
        If m_sSingakuCd <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN =" & m_sSingakuCd & " "
        End If
        If m_sSinroName <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINROMEI Like '%" & m_sSinroName & "%' "
        End If
        If w_iSinrosakiCD <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_CD = '" & w_iSinrosakiCD & "' "

        End If

        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M32.M32_SINRO_CD "

'response.write w_sSQL


        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        Else
            'ページ数の取得
            m_iMax = gf_PageCount(m_Rs,m_iDsp)

'Response.Write "m_iMax:" & m_iMax & "<br>"
        End If

		w_iThisPgCnt = w_iThisPgCnt + 1

        If m_Rs.EOF Then
            '// ページを表示
            Call showPage_NoData()
        Else
            '// ページを表示
            Call S_syousai()
        End If
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If
    
    '// 終了処理
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End If



Next

    'LABEL_showPage_OPTION_END
End sub


Sub S_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_iCnt
Dim w_cell
w_iCnt  = 0

call gs_cellPtn(m_cell)
        %>

        <tr>
        <td align="center" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M01_SYOBUNRUIMEI")) %></td>
        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M32_SINROMEI")) %></td>
        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M32_DENWABANGO")) %></td>
        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) %></td>
        <input type="hidden" name="deleteNO" value="<%=gf_HTMLTableSTR(m_Rs("M32_SINRO_CD")) %>">
        </tr>

        <%
w_i = w_i + 1
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
    //  [機能]  戻るボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_back(){

        document.frm.action="./default.asp";
        document.frm.target="fTopMain";
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
</head>
<body>

<center>


<%
If m_sMode = "Delete" Then
  m_sSubtitle = "削　除"
End If

call gs_title("進路先情報登録",m_sSubtitle)
%>
<br>
進　路　先　情　報
<br><br>
    <table border="1" class=hyo width="75%">
<form name="frm" action="delete.asp" target="_self" method="post">

    <tr>
    <th class=header>区分</th>
    <th class=header>進路名</th>
    <th class=header>TEL</th>
    <th class=header>URL</th>
    </tr>

    <% S_syousaiitiran() %>

    </table>
<br>
以上の内容を削除します。
<br><br>
<table border="0" width=50%>
<tr>
<td align=left>
<input type="button" class=button value="　削　除　" Onclick="f_delete()">
<input type="hidden" name="txtMode" value="">
<input type="hidden" name="txtRenrakusakiCD" value="<%= m_sRenrakusakiCD %>">
<input type="hidden" name="txtSinroCD2" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD2" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
<input type="hidden" name="txtDisp" value="<%= w_iThisPgCnt %>">
</td>
</form>
<form action="default.asp" target="<%=C_MAIN_FRAME%>" method="post">
<td align=right>
<input type="submit" class=button value="キャンセル">
<input type="hidden" name="txtMode" value="search">
<input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
</td>
</form>
</tr>
</table>

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