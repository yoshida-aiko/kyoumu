<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 共通処理
' ﾌﾟﾛｸﾞﾗﾑID : COM_LOGGET
' 機      能: エラーログ取得共通処理
'-------------------------------------------------------------------------
' 作      成: 2001.01.01 高丘
' 変      更: 2001/06/16 高丘     Gf_IIF追加
' 　      　: 2001/07/23 伊藤     gf_GetRecordset_OpenStatic追加
' 　      　: 2001/07/24 谷脇     関数一覧追加
'           : 2002/05/06 大前     
'             　　　　　　　　    グローバル変数の追加
'                                 gf_GetRecordset_oo4oの追加
'*************************************************************************/

'//////////////////////////////////////////////////////////////////////////////////////////
'
'	関数一覧
'
'//////////////////////////////////////////////////////////////////////////////////////////
'ＤＢオープン			 			gf_OpenDatabase()
'ＯＤＢＣＤＢオープン	 			gf_openODBCDatabase(dbOpen)	
'オラクルＤＢオープン				gf_openORAADODatabase(dbOpen)
'ＤＢオープン						AutoOpen(m_oConObj,p_sDbName,p_sConnect)
'ＤＢクローズ						gs_CloseDatabase()
'オブジェクトクローズ				gf_closeObject(oClose)
'右側にスペースをつけてそろえる。	gf_Keta(p_Str,p_iSpace)
'トランザクション開始				gs_BeginTrans()
'トランザクションをロールバック		gs_RollbackTrans()
'トランザクションをコミット			gs_CommitTrans()
'レコードセットの取得 ADO版			gf_GetRecordset(p_Rs, p_sSQL)
'レコードセットの取得(pagesize有)	gf_GetRecordsetExt(p_Rs, p_sSQL, p_iPageSize)
'レコードセットの抽出				gf_GetRecordset_OpenStatic(p_Rs,p_sSQL)
'ページ数の計算						gf_PageCount(p_Rs, p_iDsp)
'レコードカウント取得				gf_GetRsCount(p_Rs)
'絶対値ページの設定					gs_AbsolutePage(p_Rs,p_iPage, p_iDsp)
'空だったら全角スペースを返す		gf_HTMLTableSTR(p_Data)
'フィールド値の取得					gf_GetFieldValue(p_Rs, p_sField)
'ＳＱＬの実行						gf_ExecuteSQL(p_sSQL)
'エラー情報ページ表示				gs_showMsgPage(p_sWinTitle, p_sMsgTitle, p_sMsgString, p_sRetURL, p_sTarget)
'エラーメッセージを設定				gs_SetErrMsg(p_sErrMsg)
'ｴﾗｰﾒｯｾｰｼﾞﾀｲﾄﾙを取得する			gf_GetErrMsgTitle(p_sURL,p_sMsgTitle)
'エラーメッセージを取得する			gf_GetErrMsg()
'エラーメッセージをクリアする		gs_ClearErrMsg()
'テキストファイルのOPEN				gf_OpenFile(p_File,p_sPath)
'NULLが空白をハイフンに変換する		gf_SetNull2Haifun(p_sStr)
'nullチェック（IsNullもどき）		gf_IsNull(p_null)
'NULLを空文字に変換					gf_SetNull2String(p_sStr)
'NULLを0に変換						gf_SetNull2Zero(p_iNum)
'YYYY/MM/DDにフォーマット			gf_YYYY_MM_DD(p_sDate,p_sDelimit)
'和暦フォーマット					gf_fmtWareki(pDate)
'年月日整形							gf_FormatDate(p_Date,p_Delimiter)
'日付型フォーマット関数(推奨)		FormatTime(datTime,strFormat)
'文字列のバイト数を返す				f_LenB(p_str)
'IIF関数の実現						gf_IIF(p_Judge, p_tStr, p_fStr)
'数値桁合わせ(０フォーマット)		gf_fmtZero(w_str,w_kazu)
'区分マスタから各種データを取得		gf_GetKubunName(pDAIBUNRUI,pSYOBUNRUI,pNendo,pKubunName)
'区分マスタから各種データを取得(略称)gf_GetKubunName_R(pDAIBUNRUI,pSYOBUNRUI,pNendo,pKubunName)
'全角を半角に変換					gf_Zen2Han(pStr)


'** 定数定義 **
'** キャッシュオフ **
Response.Expires = 0
Response.AddHeader "Pragma", "No-Cache"

'** 構造定義 **
'** 変数宣言 ** 
'** 外部ﾌﾟﾛｼｰｼﾞｬ定義 **
%>
<!--#include file="adovbs.inc"-->
<!--#include file="CACommon.asp"-->
<!--#include file="common_combo.asp"-->
<!--#include file="com_const.asp"-->
<!--#include file="com_const_web.asp"-->
<%

Const WebRootPath = "C:\Inetpub\wwwroot/cassist"

'**EXCELﾌｧｲﾙ関連のｺﾝｽﾄ**
Const C_KK_PATH         = "C:\Inetpub\wwwroot"  '工程管理ﾌﾟﾛｼﾞｪｸﾄのパス


'DB名称
Const C_DB_PATH             = ""
Const C_BAK_PATH            = ""        'C_CSV_PATHと揃えるべき
Const C_HOME_DIR            = ""
Const C_ROOT_URL            = ""
Const C_CSV_PATH            = ""        'C_BAK_PATHと揃えるべき
Const C_DB_FILE_NAME        = ""        'MDBファイル名

CONST C_DB_NAME             = ""

'// ｴﾗｰｺｰﾄﾞ
Const C_ERR_DATA_EXIST      = -2147217900       '// ｷｰ違反の場合に発生するｴﾗｰｺｰﾄﾞ
Const C_ERR_DATA_EXIST2     = -2147467259       '// ｷｰ違反の場合に発生するｴﾗｰｺｰﾄﾞ
Const C_CommandTimeout      = 600               '// 接続を確立するまで待つ秒数
Const C_ConnectionTimeout   = 60                '// 接続確立までの待ち時間

Public m_sGrpKey                    '//ｸﾞﾙｰﾌﾟｷｰ
Public m_objDB
Public m_sErrMsg

'// ﾌｧｲﾙｵﾌﾞｼﾞｪｸﾄ
Public m_oFile
Const C_CSV_TANTO = "TANTO.CSV"     '// 担当者一覧


Set m_objDB = Server.CreateObject("ADODB.Connection")

'////////////////////////////////////////////////////////////////////////
'// データベースのオープン
'//
'// 引　数：
'// 戻り値：正常終了    : 0
'//         異常終了    : -1
'////////////////////////////////////////////////////////////////////////
Function gf_OpenDatabase()

    Dim w_bRetCode              '// Boolean戻り値
    Dim w_bErrFlg               '// ｴﾗｰﾌﾗｸﾞ
    Dim w_bErrMsg
    
    On Error Resume Next
    Err.Clear
'Response.Write "OpenDatabase1"
    gf_OpenDatabase = -1
    w_bErrFlg = True
    
'Response.Write "OpenDatabase2"
    '// ﾃﾞｰﾀﾍﾞｰｽ接続(オラクル) 
    If gf_openORAADODatabase(m_objDB)=False Then
'Response.Write "OpenDatabase3"
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        m_sErrMsg = Err.description & vbCrLf 
        w_bErrFlg = True
    else
'Response.Write "OpenDatabase4"
        '正常終了
        gf_OpenDatabase = 0
    End If
    
    Err.Clear
    
End Function

'////////////////////////////////////////////////////////////////////////
'// ＯＤＢＣデータベースのオープン
'//
'// 引　数：OUT dbOpen      : オープンするＤＢ
'// 戻り値：正常終了    : True
'//         異常終了    : False
'////////////////////////////////////////////////////////////////////////
Function gf_openODBCDatabase(dbOpen)

    On Error Resume Next
    gf_openODBCDatabase = False
    If Err <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        Response.Write "OpenODBCDataBase関数前にエラー"
    End If
    '// ﾃﾞｰﾀﾍﾞｰｽ接続
    Set dbOpen = Server.CreateObject("ADODB.Connection")
    If Err <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        Exit Function
    End If

    dbOpen.CommandTimeout = C_CommandTimeout        '// 接続を確立するまで待つ秒数
    dbOpen.ConnectionTimeout = C_ConnectionTimeout  '// 接続確立までの待ち時間
    dbOpen.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & server.MapPath (C_HOME_DIR) & C_DB_PATH

    
    If Err <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        Exit Function
    End If
    gf_openODBCDatabase = True
    
End Function

'////////////////////////////////////////////////////////////////////////
'// オラクルデータベース（ADO)のオープン
'//
'// 引　数：OUT dbOpen      : オープンするＤＢ
'// 戻り値：正常終了    : True
'//         異常終了    : False
'////////////////////////////////////////////////////////////////////////
Function gf_openORAADODatabase(dbOpen)
    On Error Resume Next

    gf_openORAADODatabase = False
    
    If Err <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        Response.Write "OpenODBCDataBase関数前にエラー"
    End If
    
    '// ﾃﾞｰﾀﾍﾞｰｽ接続
    Set dbOpen = Server.CreateObject("ADODB.Connection")
    If Err <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        Exit Function
    End If
    
    If NOT AutoOpen(dbopen,Session("CONNECT"),Session("USER_ID") & "/" & Session("PASS")) Then
        Exit Function
    End If
    
    If Err <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        Exit Function
    End If
    gf_openORAADODatabase = True
    
End Function


Public Function AutoOpen(m_oConObj,ByVal p_sDbName, ByVal p_sConnect)
'*************************************************************************************
' 機    能: データベース接続 For Oracle
' 返    値: True - 成功 / False - 失敗
' 引    数: p_sDbName - データベース名
'           p_sConnect - 接続文字列 (ユーザ名 "/" パスワード)
' 機能詳細:
' 備    考:
'*************************************************************************************
    On Error Resume Next 
    AutoOpen = False
    Dim w_vSplit            'オラクルクラスと同じ引数にする。
    
    w_vSplit = Split(p_sConnect, "/")

    '/* ADO でｺﾈｸﾄ
    ' 20220601 kiyomoto Edit ---------------------------------------------ST
    'm_oConObj.ConnectionString = "Provider=MSDAORA;"
    m_oConObj.ConnectionString = "Provider=OraOLEDB.Oracle;"
    ' 20220601 kiyomoto Edit ---------------------------------------------ED

    m_oConObj.ConnectionString = m_oConObj.ConnectionString & "Data Source=" & p_sDbName & ";"

    If Err <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        Exit Function
    End If

    m_oConObj.Open , w_vSplit(0), w_vSplit(1)
    If Err <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        Exit Function
    End If
 
   AutoOpen = True

End Function


'////////////////////////////////////////////////////////////////////////
'// データベースのクローズ
'//
'// 引　数：
'// 戻り値：
'////////////////////////////////////////////////////////////////////////
Sub gs_CloseDatabase()

    'ﾃﾞｰﾀﾍﾞｰｽをｸﾛｰｽﾞする
    gf_closeObject(m_objDB)
    
End Sub

'////////////////////////////////////////////////////////////////////////
'// オブジェクトのクローズ（データベース、レコードセット）
'//
'// 引　数：OUT oClose      : クローズするオブジェクト
'// 戻り値：正常終了    : True
'//         異常終了    : False
'////////////////////////////////////////////////////////////////////////
Function gf_closeObject(oClose)

    On Error Resume Next

   'ADO関連
    gf_closeObject = False
    oClose.Close
    set oClose = Nothing

    gf_closeObject = True

    On Error Goto 0
    Err.Clear

End Function

'********************************************************************************
'*  [機能]  文字列の桁数を右側にｽﾍﾟｰｽをつけて揃える
'*  [引数]  p_EigyoCD：営業所コード
'*  [戻値]  営業所コード
'*  [説明]  
'********************************************************************************
Function gf_Keta(p_Str,p_iSpace)

    Dim i
    Dim w_sCd
    
    On Error Resume Next
    Err.Clear

    For i = 0 To p_iSpace - f_LenB(p_Str)
        w_sCd = w_sCd & "&nbsp;"
    Next

    gf_Keta = p_Str & w_sCd
    
End Function

'////////////////////////////////////////////////////////////////////////
'// エラー情報ページ表示
'//
'// 引　数：IN  
'// 戻り値：正常終了    : True
'//         異常終了    : False
'////////////////////////////////////////////////////////////////////////
Sub gs_showMsgPage(p_sWinTitle, p_sMsgTitle, p_sMsgString, p_sRetURL, p_sTarget)
    Dim i

	'// ｴﾗｰﾒｯｾｰｼﾞｺﾝﾊﾞｰﾄ
	wErrMsg = Replace(p_sMsgString, Chr(13), "\n")
	wErrMsg = Replace(wErrMsg, Chr(10), "\n")

	'===========================================================
	'=[説明]  環境変数"URL"を取得して、
	'=		  その中のフォルダー名からｴﾗｰﾒｯｾｰｼﾞﾀｲﾄﾙを取得する
	'===========================================================
	w_sURL = request.servervariables("URL")
	Call gf_GetErrMsgTitle(w_sURL,w_sMsgTitle)

	'===========================================================
	'=[説明]  環境変数"URL"の中に
	'=		  "login/"が入ってたらCassist/default.aspにもどす。
	'=		  入ってなかったら、login/top.aspにもどす。
	'===========================================================
	'// ｴﾗｰ時戻り先 & ターゲット取得
	if InStr(w_sURL,"login/") <> 0 then 
		w_sRetURL= C_RetURL & "default.asp"
	Else
	    w_sRetURL= C_RetURL & C_ERR_RETURL
	End if
	w_sTarget="_top"

	'// 不正アクセス時
	if gf_IsNull(Session("LOGIN_ID")) then
	    w_sMsgTitle="ログインエラー"
	    w_sRetURL = C_RetURL & "default.asp"
	End if

	%>
	<HTML>
	<HEAD>
	<TITLE><%=p_sWinTitle%></TITLE>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//  [機能]  ページロード時処理
	//  [引数]
	//  [戻値]
	//  [説明]
	//************************************************************
	function window_onload() {

		window.alert("<%=w_sMsgTitle%>\n\n<%= wErrMsg %>");

		document.frm.target = "<%=w_sTarget%>";
		document.frm.action = "<%=w_sRetURL%>";
		document.frm.submit();

	}
	    
	//-->
	</SCRIPT>
	</HEAD>
	<BODY bgcolor="#ffffff" LANGUAGE=javascript onload="return window_onload()">
	<FORM method="POST" name="frm">
	<!--
	<FORM method="POST" name="frm" action="<%=p_sRetURL%>" target="<%=p_sTarget %>">
	<TABLE border="0" cellpadding="0" cellspacing="0" width="100%">
	    <TR>
	        <TD nowrap align="center" valign="middle" height="20">
	            <FONT size="4" color="#cc0000" face="ＭＳ ゴシック">
	            <B><%=p_sMsgTitle%></B></FONT>
	        </TD>
	    </TR>
	    <TR>
	        <TD nowrap align="center" valign="middle" height="10"></TD>
	    </TR>
	    <TR>
	        <TD nowrap align="center" valign="middle"><FONT size="2" face="ＭＳ ゴシック">
	        <%
	            For i = 1 To Len(p_sMsgString)
	                If Mid(p_sMsgString, i, 1) <> Chr(10) Then
	                    Response.Write Mid(p_sMsgString, i, 1)
	                Else
	                    Response.Write "<BR>"
	                End If
	            Next
	        %>
	        </FONT></TD>
	    </TR>
	    <TR>
	        <TD nowrap align="center" valign="middle" height="10"></TD>
	    </TR>
	    <TR>
	        <TD nowrap align="center" valign="middle" height="20">
				<INPUT type="submit" value="戻　る" name="Back" tabindex="15">
	        </TD>
	    </TR>
	</TABLE>
	//-->
	</FORM>
	</BODY>
	</HTML>

<%
End Sub

'********************************************************************************
'*  [機能]  ｴﾗｰﾒｯｾｰｼﾞﾀｲﾄﾙを取得する
'*  [引数]  p_sURL      : 環境変数"URL"
'*  [戻値]  p_sMsgTitle : ｴﾗｰﾒｯｾｰｼﾞﾀｲﾄﾙ
'*  [説明]  
'********************************************************************************
Function gf_GetErrMsgTitle(p_sURL,p_sMsgTitle)
    
    On Error Resume Next
    Err.Clear

	if InStr(p_sURL,"kks0110/") <> 0 then p_sMsgTitle = "授業出欠入力"
	if InStr(p_sURL,"kks0140/") <> 0 then p_sMsgTitle = "行事出欠入力"
	if InStr(p_sURL,"kks0170/") <> 0 then p_sMsgTitle = "日毎出欠入力"
	if InStr(p_sURL,"skn0130/") <> 0 then p_sMsgTitle = "試験実施科目登録"
	if InStr(p_sURL,"skn0120/") <> 0 then p_sMsgTitle = "試験監督免除申請登録"
	if InStr(p_sURL,"sei0100/") <> 0 then p_sMsgTitle = "成績登録"
	if InStr(p_sURL,"sei0200/") <> 0 then p_sMsgTitle = "成績一覧"
	if InStr(p_sURL,"sei0300/") <> 0 then p_sMsgTitle = "個人成績一覧"
	if InStr(p_sURL,"skn0170/") <> 0 then p_sMsgTitle = "試験時間割(クラス別）"
	if InStr(p_sURL,"skn0180/") <> 0 then p_sMsgTitle = "試験期間教官予定一覧"
	if InStr(p_sURL,"han0121/") <> 0 then p_sMsgTitle = "留年該当者一覧"
	if InStr(p_sURL,"gyo0200/") <> 0 then p_sMsgTitle = "行事日程一覧"
	if InStr(p_sURL,"jik0210/") <> 0 then p_sMsgTitle = "クラス別授業時間一覧"
	if InStr(p_sURL,"jik0200/") <> 0 then p_sMsgTitle = "教官別授業時間一覧"
	if InStr(p_sURL,"web0310/") <> 0 then p_sMsgTitle = "時間割交換連絡"
	if InStr(p_sURL,"mst0144/") <> 0 then p_sMsgTitle = "進路先情報登録"
	if InStr(p_sURL,"web0320/") <> 0 then p_sMsgTitle = "使用教科書登録"
	if InStr(p_sURL,"gak0460/") <> 0 then p_sMsgTitle = "指導要録所見等登録"
	if InStr(p_sURL,"gak0461/") <> 0 then p_sMsgTitle = "調査書所見等登録"
	if InStr(p_sURL,"gak0470/") <> 0 then p_sMsgTitle = "各種委員登録"
	if InStr(p_sURL,"web0340/") <> 0 then p_sMsgTitle = "個人履修選択科目決定"
	if InStr(p_sURL,"gak0300/") <> 0 then p_sMsgTitle = "学生情報検索"
	if InStr(p_sURL,"mst0113/") <> 0 then p_sMsgTitle = "中学校情報検索"
	if InStr(p_sURL,"mst0123/") <> 0 then p_sMsgTitle = "高等学校情報検索"
	if InStr(p_sURL,"mst0133/") <> 0 then p_sMsgTitle = "進路先情報検索"
	if InStr(p_sURL,"web0300/") <> 0 then p_sMsgTitle = "特別教室予約"
	if InStr(p_sURL,"web0330/") <> 0 then p_sMsgTitle = "連絡掲示板"
	if InStr(p_sURL,"web0350/") <> 0 then p_sMsgTitle = "空き時間情報検索"
	if InStr(p_sURL,"web0360/") <> 0 then p_sMsgTitle = "部活動部員一覧"
	if InStr(p_sURL,"sei0400/") <> 0 then p_sMsgTitle = "成績毎所見登録"
	if InStr(p_sURL,"sei0500/") <> 0 then p_sMsgTitle = "実力試験成績登録"

End Function

'********************************************************************************
'*  [機能]  ﾄﾗﾝｻﾞｸｼｮﾝ開始
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub gs_BeginTrans()

   On Error Resume Next
   m_objDB.BeginTrans
   On Error Goto 0
   Err.Clear
   
End Sub

'********************************************************************************
'*  [機能]  ﾄﾗﾝｻﾞｸｼｮﾝをﾛｰﾙﾊﾞｯｸする
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub gs_RollbackTrans()

   On Error Resume Next
   
   m_objDB.RollbackTrans
   
   On Error Goto 0
   Err.Clear
   
End Sub

'********************************************************************************
'*  [機能]  ﾄﾗﾝｻﾞｸｼｮﾝをｺﾐｯﾄする
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub gs_CommitTrans()

   On Error Resume Next
   
   m_objDB.CommitTrans
   
   On Error Goto 0
   Err.Clear
   
End Sub

'********************************************************************************
'*  [機能]  ﾚｺｰﾄﾞｾｯﾄを取得する
'*  [引数]  p_Rs            :取得したﾚｺｰﾄﾞｾｯﾄ
'*          p_sSQL          :ﾚｺｰﾄﾞｾｯﾄ取得のためのSQL
'*  [戻値]  0:正常終了、その他:ｴﾗｰ
'*  [説明]  
'********************************************************************************
Function gf_GetRecordset(p_Rs, p_sSQL)


    On Error Resume Next
    Err.Clear
    '正常終了を設定
    gf_GetRecordset = 0
    Do
        'ﾚｺｰﾄﾞｾｯﾄの取得
        set p_Rs = m_objDB.Execute(p_sSQL)
'	Set p_Rs = Server.CreateObject("ADODB.Recordset")
'	p_Rs.Open p_sSQL,m_objDB,adOpenStatic
        If Err <> 0 Then
            Call gs_SetErrMsg("gf_GetRecordset:" & Replace(Err.description,vbCrLf," "))
'response.write Err.description
            gf_GetRecordset = Err.number
            Exit Do
        End If
        
        '正常終了
        Exit Do
    Loop
    
    Err.Clear
    
End Function

'********************************************************************************
'*  [機能]  ﾚｺｰﾄﾞｾｯﾄを取得する
'*  [引数]  p_Rs            :取得したﾚｺｰﾄﾞｾｯﾄ
'*          p_sSQL          :ﾚｺｰﾄﾞｾｯﾄ取得のためのSQL
'*          p_iPageSize     :ﾍﾟｰｼﾞｻｲｽﾞ
'*  [戻値]  0:正常終了、その他:ｴﾗｰ
'*  [説明]  
'********************************************************************************
Function gf_GetRecordsetExt(p_Rs, p_sSQL, p_iPageSize)

    On Error Resume Next
    Err.Clear
    
    '正常終了を設定
    gf_GetRecordsetExt = 0

    Do
        'ﾚｺｰﾄﾞｾｯﾄの取得
        p_Rs.ActiveConnection = m_objDB
        p_Rs.CursorType = adOpenKeyset
        'p_Rs.CursorType = adOpenForwardOnly
        p_Rs.Source = p_sSQL
        p_Rs.LockType = adLockOptimistic
        p_Rs.Pagesize = p_iPageSize
        p_Rs.Open
        If Err <> 0 Then
            Call gs_SetErrMsg("gf_GetRecordsetExt:" & Replace(Err.description,vbCrLf," "))
            gf_GetRecordsetExt = Err.number
'response.write Err.description
            Exit Do
        End If
        
        '正常終了
        Exit Do
    Loop
    
    Err.Clear
    
End Function

'*****************************************************
'*	[概要]	データの抽出（レコードセットを取得）抽出用
'*	[引数]	p_sSQL		= SQL
'*	[戻値]	p_Rs		= レコードセット
'*  [説明]	p_Rs.MovePrevious使用可
'*@***************************************************
Function gf_GetRecordset_OpenStatic(p_Rs,p_sSQL)

	gf_GetRecordset_OpenStatic=0
	
	On Error Resume Next
	Err.Clear 

	Set p_Rs = Server.CreateObject("ADODB.Recordset")
	p_Rs.Open p_sSQL,m_objDB,adOpenStatic

	If Err.number <>0 Then
		Call gs_SetErrMsg("gf_GetRecordset_OpenStatic:" & Replace(Err.description,vbCrLf," "))
		gf_GetRecordset_OpenStatic = Err.number
		Exit Function
	End If

End Function

'********************************************************************************
'*  [機能]  ページ数の計算
'*  [引数]  p_Rs            :取得したﾚｺｰﾄﾞｾｯﾄ
'*          p_sField        :１ページごとの表示件数
'*  [戻値]  ページ数
'*  [説明]  （”レコードセット.PageCount”が使用できない場合のみ使用してください）
'********************************************************************************
Function gf_PageCount(p_Rs, p_iDsp)

    dim w_iRecCount
    On Error Resume Next
    Err.Clear
    gf_PageCount=0

    w_iRecCount=0
    p_Rs.MoveFirst
    Do Until p_Rs.EOF
        p_Rs.MoveNext
        w_iRecCount=w_iRecCount+1
    Loop
    p_Rs.MoveFirst

    gf_PageCount = INT(w_iRecCount/m_iDsp) + gf_IIF(w_iRecCount mod m_iDsp = 0,0,1)
    'gf_PageCount = INT(w_iRecCount/m_iDsp) + gf_IIF(m_Rs.RecordCount mod m_iDsp = 0,0,1)

    Err.Clear
   
End Function

'********************************************************************************
'*  [機能]  ﾚｺｰﾄﾞカウント取得
'*  [引数]  p_Rs
'*  [戻値]  gf_GetRsCount:ﾚｺｰﾄﾞ数
'*  [説明]  p_Rs.RecordCountが使えない場合
'********************************************************************************
Function gf_GetRsCount(p_Rs)
Dim w_iRecCount

    On Error Resume Next
    Err.Clear

    w_iRecCount= 0

    If p_Rs.EOF = False Then
        p_Rs.MoveFirst
        Do Until p_Rs.EOF
            p_Rs.MoveNext
            w_iRecCount=w_iRecCount+1
        Loop
        p_Rs.MoveFirst
    End If

    gf_GetRsCount = w_iRecCount
    Err.Clear

End Function

'********************************************************************************
'*  [機能]  絶対値ページの設定
'*  [引数]  p_Rs            :取得したﾚｺｰﾄﾞｾｯﾄ
'*          p_iPage         :指定したいページ
'*          p_sField        :１ページごとの表示件数
'*  [戻値]  なし
'*  [説明]  （”レコードセット.AbsolutePage”が使用できない場合のみ使用してください）
'********************************************************************************
Sub gs_AbsolutePage(p_Rs,p_iPage, p_iDsp)
    dim w_iRecCount
    On Error Resume Next
    Err.Clear

    '絶対値ページの設定
    p_Rs.MoveFirst
    for w_iRecCount=1 to p_iDsp*(p_iPage-1)
        p_Rs.MoveNext
    Next    

    Err.Clear
    
End Sub

'********************************************************************************
'*  [機能]  空だったら全角スペースを返す
'*  [引数]  p_Data          :表示したいデータ
'*  [戻値]  変換文字列
'*  [説明]  
'********************************************************************************
Function gf_HTMLTableSTR(p_Data)
    On Error Resume Next
    Err.Clear
    
    gf_HTMLTableSTR=gf_IIF(ISNULL(p_Data) OR p_DATA="","　",p_DATA)

    Err.Clear
    
End Function



'********************************************************************************
'*  [機能]  フィールドの値取得
'*  [引数]  p_Rs            :取得したﾚｺｰﾄﾞｾｯﾄ
'*          p_sField        :field name
'*  [戻値]  取得文字列
'*  [説明]  
'********************************************************************************
Function gf_GetFieldValue(p_Rs, p_sField)

    On Error Resume Next
    Err.Clear
    
    '正常終了を設定
    gf_GetFieldValue = p_Rs(p_sField)

    Err.Clear
    
End Function


'********************************************************************************
'*  [機能]  SQLを実行する
'*  [引数]  p_sSQL          :ﾚｺｰﾄﾞｾｯﾄ取得のためのSQL
'*  [戻値]  0:正常終了、その他:ｴﾗｰ
'*  [説明]  
'********************************************************************************
Function gf_ExecuteSQL(p_sSQL)
    On Error Resume Next
    Err.Clear
    
    '正常終了を設定
    gf_ExecuteSQL = 0

    Do
        'ﾚｺｰﾄﾞｾｯﾄの取得
        set p_Rs = m_objDB.Execute(p_sSQL)
        If Err <> 0 Then
            Call gs_SetErrMsg("gf_ExecuteSQL:" & Replace(Err.description,vbCrLf," "))
            gf_ExecuteSQL = Err.number
            Exit Do
        End If
        
        '正常終了
        Exit Do
    Loop
    
    Err.Clear
    
End Function

'********************************************************************************
'*  [機能]  ｴﾗｰﾒｯｾｰｼﾞを設定する
'*  [引数]  p_sErrMsg       :ｴﾗｰﾒｯｾｰｼﾞ
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub gs_SetErrMsg(p_sErrMsg)

    m_sErrMsg = p_sErrMsg

End Sub

'********************************************************************************
'*  [機能]  ｴﾗｰﾒｯｾｰｼﾞを取得する
'*  [引数]  なし
'*  [戻値]  ｴﾗｰﾒｯｾｰｼﾞ
'*  [説明]  
'********************************************************************************
Function gf_GetErrMsg()

    gf_GetErrMsg = m_sErrMsg

End Function

'********************************************************************************
'*  [機能]  ｴﾗｰﾒｯｾｰｼﾞをｸﾘｱする
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub gs_ClearErrMsg()

    m_sErrMsg = ""

End Sub

'********************************************************************************
'*  [機能]  ﾃｷｽﾄﾌｧｲﾙのOpen
'*  [引数]  p_File  ：ﾌｧｲﾙｵﾌﾞｼﾞｪｸﾄ
'*  　　　  p_sPath ：ﾊﾟｽ
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function gf_OpenFile(p_File,p_sPath)

    On Error Resume Next
    gf_OpenFile = False

    Set m_oFile = Server.CreateObject("Scripting.FileSystemObject")
    If Err <> 0 Then
        'ﾌｧｲﾙｵﾌﾞｼﾞｪｸﾄ生成失敗
        Call gs_SetErrMsg("gf_OpenFile:" & Replace(Err.description,vbCrLf," "))
        Exit Function
    End If

    Set p_File = m_oFile.CreateTextFile(p_sPath,True,False) 
    If Err <> 0 Then
        'ﾌｧｲﾙｵｰﾌﾟﾝ失敗
        Call gs_SetErrMsg("gf_OpenFile:" & Replace(Err.description,vbCrLf," "))
        Exit Function
    End If

    gf_OpenFile = True
    
End Function


'********************************************************************************
'*  [機能]  NULLか空白をハイフンに変換する
'*  [引数]  確認する文字列
'*  [戻値]  NULLか空白:ハイフン、その他:そのまま
'*  [説明]  
'********************************************************************************
Function gf_SetNull2Haifun(p_sStr)

    If isNull(p_sStr) Then
        gf_SetNull2Haifun = "-"
    Else
        If Trim(p_sStr) = "" Then
            gf_SetNull2Haifun = "-"
        Else
            gf_SetNull2Haifun = p_sStr
        End If          
    End If

End Function

'********************************************************************************
'*  [機能]  NULLを空文字に変換
'*  [引数]  p_sStr      :文字列
'*  [戻値]  変換後の文字列
'*  [説明]  指定された文字列がNULLの場合空文字を設定し、NULLでない場合そのままを設定
'********************************************************************************
Function gf_SetNull2String(p_sStr)

    If IsNull(p_sStr) Then
        gf_SetNull2String = ""
    Else
        gf_SetNull2String = p_sStr
    End If

End Function

'********************************************************************************
'*  [機能]  NULLを0に変換
'*  [引数]  p_iNum      :数値
'*  [戻値]  変換後の数値
'*  [説明]  指定された文字列がNULLの場合0を設定し、NULLでない場合そのままを設定
'********************************************************************************
Function gf_SetNull2Zero(p_iNum)

    If IsNull(p_iNum) or p_iNum="" Then
        gf_SetNull2Zero = 0
    Else
        gf_SetNull2Zero = p_iNum
    End If

End Function

'************************************
'*  nullチェック（IsNullもどき）
'*  [引数]
'*        判定対象
'*  [戻り値]
'*        true, false
'************************************
function gf_IsNull(p_null)
	w_jud = false
	if IsNull(p_null) = true then
		w_jud = true
	elseif IsEmpty(p_null) = true then
		w_jud = true
	elseif p_null = "" then
		w_jud = true
	end if
	gf_IsNull = w_jud
end function


'********************************************************************************
'*  [機能]  日付文字列をYYYY/MM/DDにﾌｫｰﾏｯﾄ
'*  [引数]  p_sDate     :日付文字列
'*          p_sDelimit  :区切り文字
'*  [戻値]  YYYY/MM/DD文字列（ｽﾗｯｼｭ部分は区切り文字による）
'*  [説明]  日付でない場合、年が４桁で与えられていない場合はそのままを返す
'********************************************************************************
Function gf_YYYY_MM_DD(p_sDate,p_sDelimit)

    Dim w_sStr
    
    gf_YYYY_MM_DD = ""
    
    If IsNull(p_sDate)  Then
        gf_YYYY_MM_DD = ""
        Exit Function
    End If
    
    '空白、日付じゃない、年が４桁ではない場合は空白を返す
    If p_sDate = "" Then 
        gf_YYYY_MM_DD = p_sDate
        Exit Function
    End If
    If IsDate(CDate(p_sDate)) <> True Then
        gf_YYYY_MM_DD = p_sDate
        Exit Function
    End If
    
    w_sMM = DatePart("m",CDate(p_sDate))
    w_sDD = DatePart("d",CDate(p_sDate))
    
    If Len(w_sMM) = 1 Then w_sMM = "0" & w_sMM
    If Len(w_sDD) = 1 Then w_sDD = "0" & w_sDD
	
    w_sStr = DatePart("yyyy",CDate(p_sDate)) & p_sDelimit
    w_sStr = w_sStr & w_sMM & p_sDelimit
    w_sStr = w_sStr & w_sDD
    
    gf_YYYY_MM_DD = w_sStr
    
End Function

'****************************************************
'[機能]	和暦フォーマット	:MM月DD日（曜日）
'[引数]	pDate : 対象日付(YYYY/MM/DD)
'[戻値]	
'****************************************************
Function gf_fmtWareki(pDate)

	gf_fmtWareki = ""

	'// Nullなら抜ける
	if gf_IsNull(trim(pDate)) then	Exit Function

	'// MM月DD日作成
	w_MM = Mid(FormatYYYYMMDD(pDate),6,2) & "月"
	w_DD = Right(FormatYYYYMMDD(pDate),2) & "日"

	'// 曜日を取得
	w_Youbi = WeekdayName(Weekday(FormatYYYYMMDD(pDate))) & "<BR>"
	w_Youbi = "（" & Left(w_Youbi,1) & "）"

	gf_fmtWareki = w_MM & w_DD & w_Youbi

End Function

'********************************************************************************
'*  [機能]  年月日整形
'*  [引数]  日付
'*  [戻値]  区切り文字付き日付(エラー時、引数をそのまま)
'*  [説明]  年月日をYYYY/MM/DDの形に変換。数字のみの年月日を区切り文字で分ける。
'********************************************************************************
Function FormatYYYYMMDD( target )
	dim yyyy, mm, dd

	yyyy = year(target)
	mm = month(target)
	dd = day(target)

	FormatYYYYMMDD = yyyy & "/" & right( "00" & mm, 2 ) & "/" & right( "00" & dd, 2)

End Function

'********************************************************************************
'*  [機能]  年月日整形
'*  [引数]  数字のみ日付(YYYYMMDD形式のもののみ)
'*  [戻値]  区切り文字付き日付(エラー時、引数をそのまま)
'*  [説明]  数字のみの年月日を区切り文字で分ける。
'*          gf_YYYY_MM_DDはYYYYMMDDを受け付けないようなので。
'********************************************************************************
Function gf_FormatDate(p_Date,p_Delimiter)
    Dim w_sDate 
    Dim w_sYear
    Dim w_sMonth
    Dim w_sDay

    '空白ならエラー
    If IsNull(p_Date)  Then
        gf_FormatDate = p_Date
        Exit Function
    End If
    If p_Date = "" Then 
        gf_FormatDate = p_Date
        Exit Function
    End If

    '数字でないならエラー
    If Not IsNumeric(p_Date) Then 
        gf_FormatDate = p_Date
        Exit Function
    End If

    '8桁でないならエラー
    If Len(p_Date) <> 8 Then
        gf_FormatDate = p_Date
        Exit Function
    End If

    w_sYear  = Mid(p_Date,1,4)
    w_sMonth = Mid(p_Date,5,2)
    w_sDay   = Mid(p_Date,7,2)

    w_sDate = w_sYear & p_Delimiter 
    w_sDate = w_sDate & w_sMonth & p_Delimiter
    w_sDate = w_sDate & w_sDay

    '最終的に日付でないならエラー
    If Not IsDate(w_sDate) Then 
        gf_FormatDate = p_Date
        Exit Function
    End If

    gf_FormatDate = w_sDate

End Function

'*********************************************************************
'　　日付型フォーマット関数　　　　　　　　　　　　　ver 1.0  00.10.19
'
'　　引数(1)：[Date]   フォーマットしたい日付型
'　　　　(2)：[String] フォーマット型（ページ後方に記載）
'　　戻値   ：[String] フォーマットされた文字列
'*********************************************************************

Function FormatTime(datTime,strFormat)

	Dim tmpFormat
	Dim cntType
	Dim FormatType

	FormatType = Split("YYYY/YY/MM/M/DD/D/HH24/H24/HH/H/II/I/SS/S/XX/ZZ","/")

	tmpFormat = Cstr(strFormat)

	For cntType = 0 To Ubound(FormatType)

		If InStr(tmpFormat,FormatType(cntType)) > 0 Then

			Select Case FormatType(cntType)
			Case "HH24"
				tmpFormat = Replace(tmpFormat,"HH24",Right(CStr(Hour(datTime) + 100),2))
			Case "H24"
				tmpFormat = Replace(tmpFormat,"H24",CStr(Hour(datTime)))
			Case "HH"
				tmpFormat = Replace(tmpFormat,"HH",Right(CStr((Hour(datTime) Mod 12) + 100),2))
			Case "H"
				tmpFormat = Replace(tmpFormat,"H",CStr(Hour(datTime) Mod 12))		
			Case "II"
				tmpFormat = Replace(tmpFormat,"II",Right(CStr(Minute(datTime) + 100),2))
			Case "I"
				tmpFormat = Replace(tmpFormat,"I",CStr(Minute(datTime)))
			Case "SS"
				tmpFormat = Replace(tmpFormat,"SS",Right(CStr(Second(datTime) + 100),2))
			Case "S"
				tmpFormat = Replace(tmpFormat,"S", CStr(Second(datTime)))
			Case "YYYY"
				If Len(CStr(Year(datTime))) = 2 Then
					If Year(datTime) > 30 Then
						tmpFormat = Replace(tmpFormat,"YYYY","19" & CStr(Year(datTime)))
					Else
						tmpFormat = Replace(tmpFormat,"YYYY","20" & CStr(Year(datTime)))
					End If
				Else
					tmpFormat = Replace(tmpFormat,"YYYY",CStr(Year(datTime)))
				End If
			Case "YY"
				tmpFormat = Replace(tmpFormat,"YY",Right(CStr(Year(datTime)),2))
			Case "MM"
				tmpFormat = Replace(tmpFormat,"MM",Right(CStr(Month(datTime) + 100),2))
			Case "M"
				tmpFormat = Replace(tmpFormat,"M",CStr(Month(datTime)))
			Case "DD"
				tmpFormat = Replace(tmpFormat,"DD",Right(CStr(Day(datTime) + 100),2))
			Case "D"
				tmpFormat = Replace(tmpFormat,"D",CStr(Day(datTime)))
			Case "XX"
				If Hour(datTime) < 12 Then
					tmpFormat = Replace(tmpFormat,"XX","午前")
				Else
					tmpFormat = Replace(tmpFormat,"XX","午後")
				End If
			Case "ZZ"
				If Hour(datTime) < 12 Then
					tmpFormat = Replace(tmpFormat,"ZZ","AM")
				Else
					tmpFormat = Replace(tmpFormat,"ZZ","PM")
				End If
			End Select
		
		End If

	Next

	FormatTime = CStr(tmpFormat)

End Function

'*********************************************************************
'　フォーマット指定できる型について（日付型からの変換）
'　　YYYY	西暦４桁
'　　YY		西暦２桁
'　　MM		月２桁
'　　M		月１桁
'　　DD		日２桁
'　　D		日１桁
'　　HH24	時２桁（２４時間）
'　　H24	時１桁（２４時間）
'　　HH		時２桁（１２時間）
'　　H		時１桁（１２時間）
'　　II		分２桁
'　　I		分１桁
'　　SS		秒２桁
'　　S		秒１桁
'　　XX		午前/午後
'　　ZZ		AM/PM
'*********************************************************************


'********************************************************************
'*  文字列のバイト数を返す
'*  [引数]
'*          p_str   :   調べる文字列
'*  [戻り値] 
'*          文字列のバイト数
'********************************************************************
Function f_LenB(p_str)

    Dim w_sbyte, w_dbyte, w_len, w_idx

    w_len = Len(p_str & "")

    For w_idx = 1 To w_len
        If Len(Hex(Asc(Mid(p_str, w_idx, 1)))) > 2 Then
            w_dbyte = w_dbyte + 1
        End If
    Next

    w_sbyte = w_len - w_dbyte

    f_LenB = w_sbyte + (w_dbyte * 2)

End Function

function gf_IIF(p_Judge, p_tStr, p_fStr)
'************************************
'*  VBのIIF関数をASPで実現
'*  [引数]
'*        VBのIIFに同じ
'*  [戻り値]
'*        VBのIIFに同じ
'************************************
    if p_Judge = true then
        gf_iif = p_tStr
    else
        gf_iif = p_fStr
    end if
end function

'***************************************************
'*  Format関数 ｾﾞﾛﾌｫｰﾏｯﾄ
'*  引数:
'*      対象の数値:w_str
'*      桁数:w_kazu
'*      例)fmtZero(125,7) ----> 0000125
'*          
'***************************************************
Function gf_fmtZero(w_str,w_kazu)
    gf_fmtZero = Right((String(w_kazu,"0") & w_str),w_kazu)
End Function

'********************************************************************
'*  不正アクセスを防ぐ
'*  [引数]
'*      p_LoginURL : ログイン画面url
'*
'********************************************************************
Function gf_UseValidRoute(p_LoginURL)

    'ﾕｰｻﾞｰ名がｾｯｼｮﾝ所持されてない場合ﾛｸﾞｲﾝ画面へ
    If Len(Session("LOGIN_ID")) = 0 Then
        Response.Redirect(p_LoginURL)
    End If

End Function


'********************************************************************************
'*  [機能]  区分マスタから各種データを取得
'*  [引数]  pDAIBUNRUI	= 大分類CD
'*			pSYOBUNRUI  = 小分類CD
'*			pNendo		= 年度
'*
'*  [戻値]  True:正常終了	False:エラー（該当なし）
'*			pKubunName  = 取得した値
'*  [説明]  
'********************************************************************************
Function gf_GetKubunName(pDAIBUNRUI,pSYOBUNRUI,pNendo,pKubunName)
	Dim w_iRet
	Dim w_sSQL
	Dim wKubunRs

	On Error Resume Next
	Err.Clear

	gf_GetKubunName = False

	'// 小分類に値が入ってなかったら抜ける
	if gf_IsNull(pSYOBUNRUI) then
		pKubunName = ""
		gf_GetKubunName = True
		Exit Function
	End if

	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	A.M01_SYOBUNRUIMEI "
	w_sSql = w_sSql & " FROM  "
	w_sSql = w_sSql & " 	M01_KUBUN A "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	 A.M01_NENDO = " & pNendo
	w_sSql = w_sSql & "  AND A.M01_DAIBUNRUI_CD = " & pDAIBUNRUI
	w_sSql = w_sSql & "  AND A.M01_SYOBUNRUI_CD = " & pSYOBUNRUI

	iRet = gf_GetRecordset(wKubunRs, w_sSql)
	If iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit Function
	End If

	if wKubunRs.Eof then
		gf_GetKubunName = True
		Exit Function
	End if

	pKubunName = wKubunRs("M01_SYOBUNRUIMEI")

    If Not IsNull(wKubunRs) Then gf_closeObject(wKubunRs)

	'//正常終了
	gf_GetKubunName = True

End Function

'********************************************************************************
'*  [機能]  区分マスタから各種データを取得(略称を取得)
'*  [引数]  pDAIBUNRUI	= 大分類CD
'*			pSYOBUNRUI  = 小分類CD
'*			pNendo		= 年度
'*
'*  [戻値]  True:正常終了	False:エラー（該当なし）
'*			pKubunName  = 取得した値
'*  [説明]  
'********************************************************************************
Function gf_GetKubunName_R(pDAIBUNRUI,pSYOBUNRUI,pNendo,pKubunName)
	Dim w_iRet
	Dim w_sSQL
	Dim wKubunRs

	On Error Resume Next
	Err.Clear

	gf_GetKubunName_R = False

	'// 小分類に値が入ってなかったら抜ける
	if gf_IsNull(pSYOBUNRUI) then
		pKubunName = ""
		gf_GetKubunName_R = True
		Exit Function
	End if

	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	A.M01_SYOBUNRUIMEI_R "
	w_sSql = w_sSql & " FROM  "
	w_sSql = w_sSql & " 	M01_KUBUN A "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	 A.M01_NENDO = " & pNendo
	w_sSql = w_sSql & "  AND A.M01_DAIBUNRUI_CD = " & pDAIBUNRUI
	w_sSql = w_sSql & "  AND A.M01_SYOBUNRUI_CD = " & pSYOBUNRUI

	iRet = gf_GetRecordset(wKubunRs, w_sSql)
	If iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit Function
	End If

	if wKubunRs.Eof then
		gf_GetKubunName_R = True
		Exit Function
	End if

	pKubunName = wKubunRs("M01_SYOBUNRUIMEI_R")

    If Not IsNull(wKubunRs) Then gf_closeObject(wKubunRs)

	'//正常終了
	gf_GetKubunName_R = True

End Function

'********************************************************************************
'*  [機能]  全角を半角に
'*  [引数]  pStr = 変換したい文字列
'*  [戻値]  なし 
'*  [説明]  
'********************************************************************************
function gf_Zen2Han(pStr)

	zenStr = "０,１,２,３,４,５,６,７,８,９,"
	zenStr = zenStr & "ア,イ,ウ,エ,オ,カ,キ,ク,ケ,コ,サ,シ,ス,セ,ソ,タ,チ,ツ,テ,ト,ナ,ニ,ヌ,ネ,ノ,ハ,ヒ,フ,ヘ,ホ,マ,ミ,ム,メ,モ,ヤ,ユ,ヨ,ラ,リ,ル,レ,ロ,ワ,ヲ,ン,"
	zenStr = zenStr & "ガ,ギ,グ,ゲ,ゴ,ザ,ジ,ズ,ゼ,ゾ,ダ,ヂ,ヅ,デ,ド,バ,ビ,ブ,ベ,ボ,パ,ピ,プ,ペ,ポ,ヴ,"
	zenStr = zenStr & "ァ,ィ,ゥ,ェ,ォ,ッ,ー,－,　"
	hanStr = "0,1,2,3,4,5,6,7,8,9,"
	hanStr = hanStr & "ｱ,ｲ,ｳ,ｴ,ｵ,ｶ,ｷ,ｸ,ｹ,ｺ,ｻ,ｼ,ｽ,ｾ,ｿ,ﾀ,ﾁ,ﾂ,ﾃ,ﾄ,ﾅ,ﾆ,ﾇ,ﾈ,ﾉ,ﾊ,ﾋ,ﾌ,ﾍ,ﾎ,ﾏ,ﾐ,ﾑ,ﾒ,ﾓ,ﾔ,ﾕ,ﾖ,ﾗ,ﾘ,ﾙ,ﾚ,ﾛ,ﾜ,ｦ,ﾝ,"
	hanStr = hanStr & "ｶﾞ,ｷﾞ,ｸﾞ,ｹﾞ,ｺﾞ,ｻﾞ,ｼﾞ,ｽﾞ,ｾﾞ,ｿﾞ,ﾀﾞ,ﾁﾞ,ﾂﾞ,ﾃﾞ,ﾄﾞ,ﾊﾞ,ﾋﾞ,ﾌﾞ,ﾍﾞ,ﾎﾞ,ﾊﾟ,ﾋﾟ,ﾌﾟ,ﾍﾟ,ﾎﾟ,ｳﾞ,"
	hanStr = hanStr & "ｧ,ｨ,ｩ,ｪ,ｫ,ｯ,ｰ,-, "
	wZen = split(zenStr,",")	
	wHan = split(hanStr,",")

	wStr = pStr
	wLen = len(wStr)
	for pf_iCnt = 0 to 89
		'response.write pf_iCnt & "---" & wZen(pf_iCnt) & "----" & wHan(pf_iCnt) & "<br>"
		bChg = false
		while not bChg
			wCnt = instr(1,wStr,wZen(pf_iCnt))
			if wCnt <> 0 then
				'response.write "wCnt=" & wCnt & "  wLen=" & wLen & "   wLen-wCnt=" & wLen-wCnt & "   wStr=" & wStr & "<br>" 
				if len(wHan(pf_iCnt)) = 2 then
					wLen = wLen + 1
					wStr = left(wStr,wCnt-1) & wHan(pf_iCnt) & right(wStr,wLen-wCnt-1)
				else 
					wStr = left(wStr,wCnt-1) & wHan(pf_iCnt) & right(wStr,wLen-wCnt)
				end if
			else 
				bChg = true
			end if
		wend
	next
	gf_Zen2Han = wStr
end function

%>