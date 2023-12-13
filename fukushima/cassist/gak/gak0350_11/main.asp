<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索結果
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0350_11/main.asp
' 機      能: 下ページ 学籍データの検索結果を表示する
'-------------------------------------------------------------------------
' 変      数:なし
' 引      数:処理年度       ＞      SESSIONより（保留）
'			txtGakunen             :学年
'           txtGakkaCD             :学科
'           txtClass               :クラス
'           txtName                :名称
'           txtGakusekiNo          :学籍番号
'           txtGakuseiNo           :学生番号
'           txtMode                :動作モード
' 説      明:
'           ■初期表示
'               タイトルのみ表示
'           ■結果表示
'               上ページで設定された検索条件にかなう学生写真を表示する
'-------------------------------------------------------------------------
' 作      成: 2006/04/28 熊野
' 変      更: 2011/04/05 iwata 学生写真データを　Sessionからでなく、データベースから取得する。
' 変      更: 2017/05/17 清本 学生写真データSession作成は global.aspで行う。
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_iSyoriNen      	   ':処理年度
    Public  m_sGakunen             ':学年
    Public  m_sGakkaCD             ':学科
    Public  m_sClass               ':クラス
    Public  m_sName                ':名称
    Public  m_sGakusekiNo          ':学籍番号
    Public  m_sGakuseiNo           ':学生番号
    Public  m_TxtMode      	       ':動作モード

    Public	m_Rs				   'recordset
    Public	m_iDsp				   '一覧表示行数

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
    w_sMsgTitle="学生情報検索結果"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	'//動作モードの取得
	m_TxtMode=request("txtMode")

    Do
		if m_TxtMode = "" then
           	Call showPage()
			Exit Do
		End if

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

'2011.04.05 ins
    '// 画像データ取得用 oo4o セッション作成
    'Set Session("OraDatabasePh") = OraSession.GetDatabaseFromPool(100)		'2017/05/17 Del Kiyomoto

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        'データ抽出SQLを作成する
        Call s_MakeSQL(w_sSQL)

       'レコードセットの取得
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
		w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL,m_iDsp)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// ページを表示
        If m_Rs.EOF Then
            Call showPage_NoData()
        Else

	    '学生情報表示せっし
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
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

'2011.04.05 ins
		'** oo4o 接続プール廃棄
	   'Session("OraDatabasePh").DestroyDatabasePool	'2017/05/17 Del Kiyomoto


End Sub

Sub s_SetParam()
'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************

	m_iSyoriNen = cint(Session("Nendo"))			'処理年度

    m_sGakunen=request("txtGakunen")            	'学年
	'コンボ未選択時
	If m_sGakunen="@@@" Then
		m_sGakunen=""
	End If

    m_sGakkaCD=request("txtGakka")            		'学科
	'コンボ未選択時
	If m_sGakkaCD="@@@" Then
		m_sGakkaCD=""
	End If

	'学年が選択されていない場合はクラスは選択できません
	if m_sGakunen="" then
		m_sClass=""
	else
    	m_sClass=request("txtClass")               	'クラス
		'コンボ未選択時
		If m_sClass="@@@" Then
			m_sClass=""
		End If
    end if

	m_sName = gf_Zen2Han(request("txtName"))        '名称(半角に変換)
	m_sGakusekiNo=request("txtGakusekiNo")          '学籍番号
	m_iDsp = cint(request("txtDisp"))				'検索リストの表示件数

End Sub


'********************************************************************************
'*  [機能]  学籍データ抽出SQL文字列の作成
'*  [引数]  p_sSql - SQL文字列
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Sub s_MakeSQL(p_sSql)

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " A.T13_GAKUSEKI_NO, "
    p_sSql = p_sSql & " A.T13_GAKUSEI_NO, "
    p_sSql = p_sSql & " B.T11_SIMEI "
    p_sSql = p_sSql & " FROM T13_GAKU_NEN A, T11_GAKUSEKI B, M02_GAKKA C, M05_CLASS D "
    p_sSql = p_sSql & " WHERE A.T13_NENDO = " & m_iSyoriNen & ""

    '検索条件のセット
    if m_sGakunen <> "" then        '学年
        p_sSql = p_sSql & " AND A.T13_GAKUNEN = " & cint(m_sGakunen)
    end if
    if m_sGakkaCD <> "" then         '学科
        p_sSql = p_sSql & " AND A.T13_GAKKA_CD = '" & m_sGakkaCD & "'"
    end if
    if m_sClass <> "" then           'クラス
        p_sSql = p_sSql & " AND A.T13_CLASS = '" & m_sClass & "'"
    end if

    '結合条件
    p_sSql = p_sSql & " AND A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO(+) "
    p_sSql = p_sSql & " AND A.T13_NENDO = C.M02_NENDO"
    p_sSql = p_sSql & " AND A.T13_NENDO = D.M05_NENDO"
    p_sSql = p_sSql & " AND A.T13_GAKUNEN = D.M05_GAKUNEN "
    p_sSql = p_sSql & " AND A.T13_CLASS = D.M05_CLASSNO "
    p_sSql = p_sSql & " AND A.T13_GAKKA_CD = C.M02_GAKKA_CD "

    p_sSql = p_sSql & " ORDER BY A.T13_GAKUNEN,A.T13_GAKUSEKI_NO "

End Sub

'********************************************************************************
'*  [機能]  写真があるか検索 (BLOB型対応の為 T09_IMAGE を　T09_GAKUSEI_NOに変更）
'*  [引数]  なし
'*  [戻値]  True: False
'*  [説明]
'********************************************************************************
Function f_Photoimg(pGAKUSEI_NO)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_Photoimg = False

	'// NULLなら抜ける(False)
	if trim(pGAKUSEI_NO) = "" then Exit Function

	Do
	    w_sSQL = ""
	    w_sSQL = w_sSQL & " SELECT "
	    w_sSQL = w_sSQL & " T09_GAKUSEI_NO "
	    w_sSQL = w_sSQL & " FROM T09_GAKU_IMG "
	    w_sSQL = w_sSQL & " WHERE T09_GAKUSEI_NO = '" & cstr(pGAKUSEI_NO) & "'"

		w_iRet = gf_GetRecordset(w_ImgRs, w_sSQL)

		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			Exit Do
		End If

		'// EOFなら抜ける(False)
		if w_ImgRs.Eof then	Exit Do

		'//正常終了
		f_Photoimg = True
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  担任名の取得
'*  [引数]  なし
'*  [戻値]  担任名
'*  [説明]
'********************************************************************************
Function sGetTannin()
	Dim w_iRet
	Dim w_sSQL
	Dim w_oRs

	sGetTannin = ""

	w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "  M04_KYOKANMEI_SEI ,"
    w_sSQL = w_sSQL & "  M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "  M04_KYOKAN ,"
    w_sSQL = w_sSQL & "  M05_CLASS "
    w_sSQL = w_sSQL & " WHERE "

    '結合条件
    w_sSQL = w_sSQL & "  M04_KYOKAN.M04_NENDO = M05_CLASS.M05_NENDO "
    w_sSQL = w_sSQL & "  AND M04_KYOKAN.M04_KYOKAN_CD = M05_CLASS.M05_TANNIN "

    '検索条件のセット
    w_sSQL = w_sSQL & "  AND M05_CLASS.M05_NENDO = " & m_iSyoriNen
    '学年
    If m_sGakunen <> "" Then
		w_sSQL = w_sSQL & "  AND M05_CLASS.M05_GAKUNEN = " & cint(m_sGakunen)
	Else
		w_sSQL = w_sSQL & "  AND M05_CLASS.M05_GAKUNEN = null"
	End If
	'学科
    If m_sGakkaCD <> "" Then
		w_sSQL = w_sSQL & "  AND ( M05_CLASS.M05_GAKKA_CD = '" & m_sGakkaCD & "'"
	Else
		w_sSQL = w_sSQL & "  AND ( M05_CLASS.M05_GAKKA_CD = ''"
	End If
	'クラス
    If m_sClass <> "" Then
		w_sSQL = w_sSQL & "  OR M05_CLASS.M05_CLASSNO = " & m_sClass & " )"
	Else
		w_sSQL = w_sSQL & "  OR M05_CLASS.M05_CLASSNO = null ) "
	End If

    w_iRet = gf_GetRecordset(w_oRs,w_sSQL)

	'データの獲得に失敗したら抜ける
    If w_iRet <> 0 Then Exit Function
    '教官名取得できない場合は抜ける
    If w_oRs.EOF Then Exit Function

	'教官名を関数の戻り値にセット
	sGetTannin = "担任："
	sGetTannin = sGetTannin & w_oRs.fields(0).value		'教官姓
	sGetTannin = sGetTannin & "　"						'姓名間のスペース
	sGetTannin = sGetTannin & w_oRs.fields(1).value		'教官名

end Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Sub showPage_NoData()

%>
	<html>
	<head>
	<title>学生情報検索</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
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
    //-->
    </SCRIPT>
    </head>

    <body>
	<% if m_TxtMode = "" then %>
		<center>
		<br><br><br>
		<span class="msg">項目を選んで表示ボタンを押してください</span>
		</center>
	<% Else %>
	    <div align="center">
	    <form action="" method="post" name="frm" target="">
		<% If sGetTannin <> "" Then %>
		<div align="left"><%= sGetTannin %></div>
		<% End If %>
		<table ID="Table1"><tr><td align="center" >
			<table border="0" width="100%" >
				<tr>
					<td align="center">
					<% if m_TxtMode = "" then %>
						<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
							<tr>
								<td width="60">&nbsp</td>
								<td valign="top"></td>
							</tr>
						</table>
					<% else %>
						<% dim w_cell %>

						<!--  学生写真表示　-->
							<table border="0" cellpadding="0" cellspacing="2">
								<%
									w_iCnt = 1
									Do Until m_Rs.Eof or w_iCnt > m_iDsp
										response.write 	"<tr>"
										i_TdLine = 1
										'// 横に５件表示ライン
										Do Until m_Rs.Eof or i_TdLine > 5 or w_iCnt > m_iDsp
										%>
											<td align="center" class=search width="150" valign="top">
												<%
												'// 顔写真があるか先に検索する
												w_bRet = ""
												w_bRet = f_Photoimg(m_Rs("T13_GAKUSEI_NO"))

												if w_bRet = True then
													' 2011.04.05 upd DispBinary => DispBinaryRec に変更
													' 2023.11.24 upd DispBinaryRec => DispBinary に変更
													%><IMG SRC="DispBinary.asp?gakuNo=<%= m_Rs("T13_GAKUSEI_NO") %>" width="90" height="120" border="0"><%

												Else
													%><IMG SRC="images/Img0000000000.gif" width="90" height="120" border="0"><%
												End if
												%></a><br>
												<table border="0" cellpadding="0" cellspacing="2" width="100%">
													<tr><td bgcolor="#666699" nowrap><font color="#FFFFFF" size="1"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></font></td><td><font size="1"><%= m_Rs("T13_GAKUSEKI_NO") %></font></td></tr>
													<tr><td bgcolor="#666699" nowrap><font color="#FFFFFF" size="1">氏名    </font></td><td><font size="1"><%= trim(m_Rs("T11_SIMEI")) %></font></td></tr>
												</table>
											</td>
											<%
											i_TdLine = i_TdLine + 1
											w_iCnt = w_iCnt + 1
											m_Rs.MoveNext
										Loop %>
										</tr>
									<% Loop	%>
								</tr>
							</table>
					<% end if %>
				</td>
			</tr>
		</table>


		</td></tr></table>

		</div>
	    <input type="hidden" name="txtMode">

		<%' 検索条件 %>
		<input type="hidden" name="txtGakunen"     value="<%=request("txtGakunen")%>">
		<input type="hidden" name="txtGakka"       value="<%=request("txtGakka")%>">
		<input type="hidden" name="txtClass"       value="<%=request("txtClass")%>">
		<input type="hidden" name="txtName"        value="<%=request("txtName")%>">
		<input type="hidden" name="txtGakusekiNo"  value="<%=request("txtGakusekiNo")%>">
		<input type="hidden" name="txtGakuseiNo"   value="<%=request("txtGakuseiNo")%>" ID="Hidden1">
		</form>
	<% End if %>
	</body>

    </html>

<%
    '---------- HTML END   ----------
End Sub

%>

