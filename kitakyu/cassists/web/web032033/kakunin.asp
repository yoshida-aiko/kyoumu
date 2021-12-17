<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 使用教科書登録確認
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0144/kakunin.asp
' 機      能: 下ページ 就職先マスタの詳細変更を行う
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
' 作      成: 2001/07/12 岩下 幸一郎
' 変      更: 2001/08/22 伊藤 公子 教官を選択できるように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_sDBMode           'DBのﾓｰﾄﾞの設定

    '取得したデータを持つ変数
    Public  m_Rs            'recordset
    Public  m_sNendo
    Public  m_sGakkiCD
    Public  m_sNo
    Public  m_sGakunenCD
    Public  m_sGakkaCD
    Public  m_sCourseCD
    Public  m_sKamokuCD
    Public  m_sKyokanMei
    Public  m_sKyokasyoName
    Public  m_sSyuppansya
    Public  m_sTyosya
    Public  m_sKyokanyo
    Public  m_sSidousyo
    Public  m_sBiko
    Public  m_sKyokan_CD

    ''名称
    Public  m_sSYOBUNRUI_CD
    Public  m_sSYOBUNRUIMEI
    Public  m_sGAKKAMEI
    Public  m_sCOURSEMEI
    Public  m_sKAMOKUMEI

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
    w_sMsgTitle="使用教科書登録確認"
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
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '// 画面に表示する名称を取得
        if f_Get_Name = False then
            exit do
        end if

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


'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    Dim strErrMsg

    strErrMsg = ""
    m_sDBMode    = Request("txtMode")
    m_sNendo     = Request("txtNendo")      '年度の取得
    m_sGakkiCD   = Request("txtGakkiCD")    '学期の取得
    m_sNo = Request("txtUpdNo")             ''更新用No格納
    m_sGakunenCD     = Request("txtGakunenCD")  '学年の取得
    m_sGakkaCD   = Request("txtGakkaCD")    '学科の取得
    m_sCourseCD  = Request("txtCourseCD")   'コースの取得
    m_sKamokuCD  = Request("txtKamokuCD")   '科目の取得
    m_sKyokanMei     = Request("txtKyokanMei")  '教官名の取得
    m_sKyokasyoName  = Request("txtKyokasyoName")   '教科書名の取得
    m_sSyuppansya    = Request("txtSyuppansya") '出版社の取得
    m_sTyosya    = Request("txtTyosya")     '著者の取得
    m_sKyokanyo  = Request("txtKyokanyo")   '教官用の取得
    m_sSidousyo  = Request("txtSidousyo")   '指導書の取得
    m_sBiko      = Request("txtBiko")       '教官用の取得

    m_sKyokan_CD = Request("SKyokanCd1")

    m_sSYOBUNRUI_CD = ""
    m_sSYOBUNRUIMEI = ""
    m_sGAKKAMEI = ""
    m_sCOURSEMEI = ""
    m_sKAMOKUMEI = ""

	If m_sKyokanyo = "" Then
	  m_sKyokanyo = 0
	End If

	If m_sSidousyo = "" Then
	  m_sSidousyo = 0
	End If

    If strErrmsg <> "" Then
        ' エラーを表示するファンクション
        Call err_page(strErrMsg)
        response.end
    End If

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
Exit Sub

	response.write("<BR>m_sDBMode = " & m_sDBMode)
	response.write("<BR>m_sNendo = " & m_sNendo)
	response.write("<BR>m_sGakkiCD = " & m_sGakkiCD)
	response.write("<BR>m_sNo = " & m_sNo)
	response.write("<BR>m_sGakunenCD = " & m_sGakunenCD)
	response.write("<BR>m_sGakkaCD = " & m_sGakkaCD)
	response.write("<BR>m_sCourseCD = " & m_sCourseCD)
	response.write("<BR>m_sKamokuCD = " & m_sKamokuCD)
	response.write("<BR>m_sKyokanMei = " & m_sKyokanMei)
	response.write("<BR>m_sKyokasyoName = " & m_sKyokasyoName)
	response.write("<BR>m_sSyuppansya = " & m_sSyuppansya)
	response.write("<BR>m_sTyosya = " & m_sTyosya)
	response.write("<BR>m_sKyokanyo = " & m_sKyokanyo)
	response.write("<BR>m_sSidousyo = " & m_sSidousyo)
	response.write("<BR>m_sBiko = " & m_sBiko)

End Sub

'********************************************************************************
'*  [機能]  名称を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
function f_Get_Name
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_Rs

    f_Get_Name = False

	'==============
    ''学期名称取得
	'==============
    w_sSQL = ""
    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " M01.M01_SYOBUNRUIMEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_SYOBUNRUI_CD "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    M01_KUBUN M01 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    'w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD  =  " & 51 & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD  =  " & C_KAISETUKI & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M01.M01_SYOBUNRUI_CD  =  " & m_sGakkiCD & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M01.M01_NENDO         =  " & m_sNendo

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    End If

	If w_Rs.EOF = false Then
	    m_sSYOBUNRUI_CD = gf_HTMLTableSTR(w_Rs("M01_SYOBUNRUI_CD"))
	    m_sSYOBUNRUIMEI = gf_HTMLTableSTR(w_Rs("M01_SYOBUNRUIMEI"))
	End If

    w_Rs.close
    set w_Rs = nothing

	'=================
    ''学科情報を取得
	'=================
    If cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) Then
        m_sGAKKAMEI = "全学科"
    else
	    w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M02.M02_GAKKAMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKA M02 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M02.M02_NENDO         =  " & m_sNendo & " AND "
        If cstr(m_sGakkaCD) <> cstr(C_CLASS_ALL) Then
                w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKA_CD          = '" & m_sGakkaCD & "'"
        End If

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Function
        End If

		If w_Rs.EOF = false Then
	        m_sGAKKAMEI = gf_HTMLTableSTR(w_Rs("M02_GAKKAMEI"))
		End If

        w_Rs.close
        set w_Rs = nothing
    end if

	'=================
    ''ｺｰｽ情報を取得
	'=================
    If cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) Then
        m_sCOURSEMEI = ""
    else

        If m_sCOURSECD <> "@@@" AND m_sCOURSECD <> "" Then
		    w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " SELECT "
            w_sSQL = w_sSQL & vbCrLf & " M20.M20_COURSEMEI "
            w_sSQL = w_sSQL & vbCrLf & " FROM "
            w_sSQL = w_sSQL & vbCrLf & "    M20_COURSE M20 "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "    M20.M20_NENDO         =  " & m_sNendo & " AND "
            w_sSQL = w_sSQL & vbCrLf & "    M20_GAKKA_CD          = '" & m_sGakkaCD & "' AND "
            w_sSQL = w_sSQL & vbCrLf & "    M20_GAKUNEN           =  " & m_sGakunenCD & " AND "
            w_sSQL = w_sSQL & vbCrLf & "    M20_COURSE_CD         = '" & m_sCOURSECD & "'"

            w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
            If w_iRet <> 0 Then
                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                m_bErrFlg = True
                Exit Function
            End If

			If w_Rs.EOF = false Then
	            m_sCOURSEMEI = gf_HTMLTableSTR(w_Rs("M20_COURSEMEI"))
			End If

            w_Rs.close
            set w_Rs = nothing
        else
            m_sCOURSEMEI = ""
        end if
    end if

	'=================
    ''科目情報を取得
	'=================
    w_sSQL = ""
    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " T15.T15_KAMOKUMEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    T15_RISYU T15 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    T15.T15_NYUNENDO      =  " & (m_sNendo - m_sGakunenCD + 1) & " AND "
    If cstr(m_sGakkaCD) <> cstr(C_CLASS_ALL) Then
        w_sSQL = w_sSQL & vbCrLf & "    T15_GAKKA_CD          = '" & m_sGakkaCD & "' AND "
    else
        w_sSQL = w_sSQL & vbCrLf & "    T15_KAMOKU_KBN          = 0 AND "
    End If
    w_sSQL = w_sSQL & vbCrLf & "    T15_KAMOKU_CD         = '" & m_sKAMOKUCD & "'"

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    End If

	If w_Rs.EOF = false Then
	    m_sKAMOKUMEI = gf_HTMLTableSTR(w_Rs("T15_KAMOKUMEI"))
	End If

    w_Rs.close
    set w_Rs = nothing

    f_Get_Name = True

end function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()

%>

<html>

    <head>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_BackClick(){

	    document.frm.txtMode.value = "Disp";
	    document.frm.action="./touroku.asp";
	    document.frm.target="_self";
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

	    document.frm.txtMode.value = "<%=m_sDBMode%>";
	    document.frm.action='./db.asp';
	    document.frm.target="_self";
	    document.frm.submit();
    }

    //-->
    </SCRIPT>

    <link rel=stylesheet href="../../common/style.css" type=text/css>

    </head>

	<body>

	<center>

	<form name="frm" action="" target="_self">

	<br>

	<% call gs_title("使用教科書登録",Request("txtTitle")) %>

	<br>

	登　録　内　容
	<br><br>

	    <table border="0" class=form width=75%>
	    <tr>
	    <td class=form align="left" width="100">年　度</td>
	    <td class=form align="left">
	    <%= m_sNendo %>年
	    <input type="hidden" name="txtNendo" value="<%= m_sNendo %>">
	    <BR></td>
	    </tr>

	    <tr>
	    <td class=form align="left">学　期</td>
	    <td class=form align="left">
	    <%= m_sSYOBUNRUIMEI %>
	    <input type="hidden" name="txtGakkiCD" value="<%= m_sSYOBUNRUI_CD %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">学　年</td>
	    <td class=form align="left">
	    <%= m_sGakunenCD %>年
	    <input type="hidden" name="txtGakunenCD" value="<%= m_sGakunenCD %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">学　科</td>
	    <td class=form align="left">
	<%
	If cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) then
	    response.write "全学科"
	  Else
	    response.write m_sGAKKAMEI
	End If
	%>
	    <input type="hidden" name="txtGakkaCD" value="<%= m_sGakkaCD %>">
	    <BR></td>
	    </tr>

	    <tr>
	    <td class=form align="left">コース</td>
	    <td class=form align="left">
	<%
	If m_sCOURSECD <> "@@@" AND m_sCOURSECD <> "" Then
	response.write m_sCOURSEMEI
	End If
	%>
	    <input type="hidden" name="txtCourseCD" value="<%If m_sCourseCD = "@@@" Then
	response.write ""
	Else
	response.write m_sCourseCD 
	End If%>">
	    <BR></td>
	    </tr>

	    <tr>
	    <td class=form align="left">科　目</td>
	    <td class=form align="left">
	    <%= m_sKAMOKUMEI %>
	    <input type="hidden" name="txtKamokuCD" value="<%= m_sKamokuCD %>">
	    <BR></td>
	    </tr>

	    <tr>
	    <td class=form align="left">教　官</td>
	    <td class=form align="left">
	    <%= m_sKyokanMei %>
	    <input type="hidden" name="txtKyokanMei" value="<%= m_sKyokanMei %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">教科書名</td>
	    <td class=form align="left">
	    <%= m_sKyokasyoName %>
	    <input type="hidden" name="txtKyokasyoName" value="<%= m_sKyokasyoName %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">出版社</td>
	    <td class=form align="left">
	    <%= m_sSyuppansya %>
	    <input type="hidden" name="txtSyuppansya" value="<%= m_sSyuppansya %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">著者名</td>
	    <td class=form align="left">
	    <%= m_sTyosya %>
	    <input type="hidden" name="txtTyosya" value="<%= m_sTyosya %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">教官用</td>
	    <td class=form align="left"><%= m_sKyokanyo %>冊
	    <input type="hidden" name="txtKyokanyo" value="<%= m_sKyokanyo %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">指導書</td>
	    <td class=form align="left"><%= m_sSidousyo %>冊
	    <input type="hidden" name="txtSidousyo" value="<%= m_sSidousyo %>">
	    <BR></td>
	    </tr>
	    <tr>
	    <td class=form align="left">備考</td>
	    <td class=form align="left">
	    <%= m_sBiko %>
	    <input type="hidden" name="txtBiko" value="<%= m_sBiko %>">

	    <BR></td>
	    </tr>
	    </table>
	<br>
	以上の内容で登録します。
	<br><br>
	<table border="0">
	<tr>
	<td valign="top">
	<input type="button" class=button value="　登　録　" Onclick="f_SinkiClick()">
	<input type="hidden" name="txtTitle" value="<%= Request("txtTitle") %>">
	<input type="hidden" name="txtUpdNo" value="<%= Request("txtUpdNo") %>">
	<img src="../../image/sp.gif" width="20" height="1">
	<input type="button" class=button value="キャンセル" Onclick="f_BackClick()">

	<!--値渡し用-->
	<input type="hidden" name="txtMode" value="">
    <input type="hidden" name="SKyokanCd1" value="<%=m_sKyokan_CD%>">

    <input type="hidden" name="KeyNendo" value="<%=Request("KeyNendo")%>">

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