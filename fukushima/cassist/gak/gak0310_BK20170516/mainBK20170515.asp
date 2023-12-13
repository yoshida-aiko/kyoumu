<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索結果
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0310/main.asp
' 機      能: 下ページ 学籍データの検索結果を表示する
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtHyoujiNendo         :表示年度
'           txtGakunen             :学年
'           txtGakkaCD             :学科
'           txtClass               :クラス
'           txtName                :名称
'           txtGakusekiNo          :学籍番号
'           txtSeibetu             :性別
'           txtGakuseiNo           :学生番号
'           txtIdou                :異動
'           txtTyuClub             :中学校クラブ
'           txtClub                :現在クラブ
'           txtRyoseiKbn           :寮
'           CheckImage               :画像表示指定
'           txtMode                :動作モード
'                               BLANK   :初期表示
'                               SEARCH  :結果表示
' 説      明:
'           ■初期表示
'               タイトルのみ表示
'           ■結果表示
'               上ページで設定された検索条件にかなう学生情報を表示する
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
'           : 2002/05/06 BLOB型対応の為 T09_IMAGE を　T09_GAKUSEI_NOに変更
' 変      更: 2011/04/05 iwata 学生写真データを　Sessionからでなく、データベースから取得する。
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_TxtMode      	       ':動作モード
	Public  m_iSyoriNen      	   ':処理年度
    Public  m_iHyoujiNendo         ':表示年度
    Public  m_sGakunen             ':学年
    Public  m_sGakkaCD             ':学科
    Public  m_sClass               ':クラス
    Public  m_sName                ':名称
    Public  m_sGakusekiNo          ':学籍番号
    Public  m_sSeibetu             ':性別
    Public  m_sGakuseiNo           ':学生番号
    Public  m_sIdou                ':異動
    Public  m_sTyuClub             ':中学校クラブ
    Public  m_sClub                ':現在クラブ
    Public  m_sRyoseiKbn           ':寮
    Public  m_sCheckImage          ':画像表示指定
	Public  m_sTyugaku			   ':出身中学校

    Public	m_Rs					'recordset
    Public	m_iDsp					'一覧表示行数

    Public  m_iPageTyu      		':表示済表示頁数（自分自身から受け取る引数）

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

	'//セッション情報・動作モードの取得
	m_iSyoriNen = Session("NENDO")
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
    Set Session("OraDatabasePh") = OraSession.GetDatabaseFromPool(100)

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

	    '学生情報表示
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
	   Session("OraDatabasePh").DestroyDatabasePool

End Sub

Sub s_SetParam()
'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************

'    Session("HyoujiNendo") = request("txtHyoujiNendo")     	'表示年度
    Session("HyoujiNendo") = Session("NENDO")		'表示年度	'<-- 8/16修正	持
    m_sGakunen=request("txtGakunen")            	'学年
	'コンボ未選択時
	If m_sGakunen="@@@" Then
		m_sGakunen=""
	End If
    m_sGakkaCD=request("txtGakka")            	'学科
	'コンボ未選択時
	If m_sGakkaCD="@@@" Then
		m_sGakkaCD=""
	End If

	if m_sGakunen="" then	'学年が選択されていない場合はクラスは選択できません
		m_sClass=""
	else
    	m_sClass=request("txtClass")               	'クラス
		'コンボ未選択時
		If m_sClass="@@@" Then
			m_sClass=""
		End If
    end if

	m_sName = gf_Zen2Han(request("txtName"))                	'名称(半角に変換)

	m_sGakusekiNo=request("txtGakusekiNo")          '学籍番号
	m_sSeibetu=request("txtSeibetu")            	'性別
	'コンボ未選択時
	If m_sSeibetu="@@@" Then
		m_sSeibetu=""
	End If
	m_sGakuseiNo=request("txtGakuseiNo")           	'学生番号
	m_sIdou =request("TxtIdou")               	'異動
	'コンボ未選択時
	If m_sIdou="@@@" Then
		m_sIdou=""
	End If
	m_sTyuClub =request("txtTyuClub")            	'中学校クラブ
	'コンボ未選択時
	If m_sTyuClub="@@@" Then
		m_sTyuClub=""
	End If
	m_sClub=request("txtClub")                	'現在クラブ
	'コンボ未選択時
	If m_sClub="@@@" Then
		m_sClub=""
	End If
	m_sRyoseiKbn=request("txtRyoseiKbn")           	'寮
	'コンボ未選択時
	If m_sRyoseiKbn="@@@" Then
		m_sRyoseiKbn=""
	End If

	m_iDsp = cint(request("txtDisp"))						':検索リストの表示件数

    '// BLANKの場合は行数ｸﾘｱ
    If m_TxtMode = "Search" Then
        m_iPageTyu = 1
    Else
        m_iPageTyu = int(Request("txtPageTyu"))     ':表示済表示頁数（自分自身から受け取る引数）
    End If

	m_sCheckImage=request("CheckImage")           	'画像表示指定

	m_sTyugaku = request("txtTyugaku")

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
    p_sSql = p_sSql & " A.T13_GAKUNEN, "
    p_sSql = p_sSql & " E.M05_CLASSMEI, "
    p_sSql = p_sSql & " B.T11_SIMEI, "
    p_sSql = p_sSql & " B.T11_SEIBETU, "
    p_sSql = p_sSql & " D.M01_SYOBUNRUIMEI, "
    p_sSql = p_sSql & " C.M02_GAKKARYAKSYO "
    p_sSql = p_sSql & " FROM T13_GAKU_NEN A, T11_GAKUSEKI B, M02_GAKKA C, M01_KUBUN D, M05_CLASS E "
    p_sSql = p_sSql & " WHERE A.T13_NENDO = " & cint(Session("HyoujiNendo")) & ""

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
    if m_sName <> "" then            '名称
        p_sSql = p_sSql & " AND B.T11_SIMEI_KD LIKE '%" & m_sName & "%'"
    end if
    if m_sGakusekiNo <> "" then 	'学籍番号
    	p_sSql = p_sSql & " AND A.T13_GAKUSEKI_NO LIKE '%" & m_sGakusekiNo & "%'"
    end if
    if m_sSeibetu <> "" then 		'性別
    	p_sSql = p_sSql & " AND B.T11_SEIBETU = " & m_sSeibetu
    end if
    if m_sGakuseiNo <> "" then 		'学生番号
        p_sSql = p_sSql & " AND A.T13_GAKUSEI_NO LIKE '%" & m_sGakuseiNo & "%'"
    end if
    if m_sIdou <> "" then 		'異動
        p_sSql = p_sSql & " AND ( A.T13_IDOU_KBN_1 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_2 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_3 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_4 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_5 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_6 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_7 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_8 = '" & m_sIdou & "' )"
    end if

    if m_sTyuClub <> "" then 		'中学校クラブ
	    p_sSql = p_sSql & " AND B.T11_TYU_CLUB = '" & m_sTyuClub & "'"
    end if
    if m_sClub <> "" then 		'現在クラブ
	    p_sSql = p_sSql & " AND ( A.T13_CLUB_1 = '" & m_sClub & "'"
	    p_sSql = p_sSql & " OR A.T13_CLUB_2 = '" & m_sClub & "' ) "
    end if
    if m_sRyoseiKbn <> "" then 		'寮
        p_sSql = p_sSql & " AND A.T13_RYOSEI_KBN = '" & m_sRyoseiKbn & "'"
    end if

    if m_sTyugaku <> "" then 		'出身中学校
        p_sSql = p_sSql & " AND B.T11_TYUGAKKO_CD IN ("
        p_sSql = p_sSql & " SELECT M13_TYUGAKKO_CD "
        p_sSql = p_sSql & " FROM M13_TYUGAKKO "
        p_sSql = p_sSql & " WHERE M13_TYUGAKKOMEI like '%" & m_sTyugaku & "%') "
    end if

    '結合条件
'    p_sSql = p_sSql & " AND A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO "
'    p_sSql = p_sSql & " AND M02_NENDO(+) = '" & cstr(Session("HyoujiNendo")) & "'"
'    p_sSql = p_sSql & " AND M02_GAKKA_CD(+) = T13_GAKKA_CD "
'    p_sSql = p_sSql & " AND M01_NENDO(+) = '" & cstr(Session("HyoujiNendo")) & "'"
'    p_sSql = p_sSql & " AND M01_DAIBUNRUI_CD(+) = 1 "
'    p_sSql = p_sSql & " AND M01_SYOBUNRUI_CD(+) = T11_SEIBETU "
'    p_sSql = p_sSql & " AND M05_CLASSNO(+) = T13_CLASS "


    p_sSql = p_sSql & " AND A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO(+) "
    p_sSql = p_sSql & " AND A.T13_NENDO = C.M02_NENDO"
    p_sSql = p_sSql & " AND A.T13_NENDO = D.M01_NENDO"
    p_sSql = p_sSql & " AND A.T13_NENDO = E.M05_NENDO"
    p_sSql = p_sSql & " AND A.T13_GAKUNEN = E.M05_GAKUNEN "
    p_sSql = p_sSql & " AND A.T13_CLASS = E.M05_CLASSNO "
    p_sSql = p_sSql & " AND A.T13_GAKKA_CD = C.M02_GAKKA_CD "
    p_sSql = p_sSql & " AND D.M01_DAIBUNRUI_CD = 1 "
    p_sSql = p_sSql & " AND B.T11_SEIBETU = D.M01_SYOBUNRUI_CD "

    p_sSql = p_sSql & " ORDER BY A.T13_GAKUNEN,A.T13_GAKUSEKI_NO "
'    p_sSql = p_sSql & " ORDER BY A.T13_GAKUNEN, D.M01_SYOBUNRUIMEI,"
'    p_sSql = p_sSql & " A.T13_CLASS, A.T13_GAKUSEKI_NO, A.T13_GAKUSEI_NO "
'response.write " p_sSql=" & p_sSql & "<BR>"

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

		iRet = gf_GetRecordset(w_ImgRs, w_sSQL)
		If iRet <> 0 Then
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
	Dim w_pageBar			'ページBAR表示用
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
    //  [機能]  詳細ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_detail(pGAKUSEI_NO){

			url = "kojin.asp?hidGAKUSEI_NO=" + pGAKUSEI_NO;
			w   = 800;
			h   = 600;

			wn  = "SubWindow";
			opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
			if (w > 0)
				opt = opt + ",width=" + w;
			if (h > 0)
				opt = opt + ",height=" + h;
			newWin = window.open(url, wn, opt);

//		document.frm.hidGAKUSEI_NO.value = pGAKUSEI_NO;
//		document.forms[0].submit();
    }

    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="main.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageTyu.value = p_iPage;
        document.frm.submit();

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
	    <form action="kojin.asp" method="post" name="frm" target="_detail">

		<BR>
		<table><tr><td align="center">
		<%
			'ページBAR表示
			Call gs_pageBar(m_Rs,m_iPageTyu,m_iDsp,w_pageBar)
		%>
		<%=w_pageBar %>

			<table border="0" width="100%">
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

					    <!--  学生情報表示　-->
						<% if m_sCheckImage = "" then %>
								<table border="1" width="600" class=hyo>
									<tr>
										<th height=16 class=header><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
										<th height=16 class=header>学年</th>
										<th height=16 class=header>学科</th>
										<th height=16 class=header>クラス</th>
										<th height=16 class=header>氏　　名</th>
										<th height=16 class=header>性別</th>
										<th height=16 class=header><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_5NEN)%></th>
										<th height=16 class=header>詳細</th>
									</tr>

						        	<%
	'									m_Rs.Movefirst
										w_iCnt = 0
										Do Until m_Rs.EOF or w_iCnt >= m_iDsp
											call gs_cellPtn(w_cell)
											%>
											<tr>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEKI_NO")) %>&nbsp</td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T13_GAKUNEN")) %>&nbsp</td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M02_GAKKARYAKSYO")) %>&nbsp</td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M05_CLASSMEI")) %>&nbsp</td>
												<td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T11_SIMEI")) %></td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M01_SYOBUNRUIMEI")) %></td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEI_NO")) %>&nbsp</td>
												<td align="center" height="16" class=<%=w_cell%>><input type=button class=button value="詳細" onclick="f_detail('<%= gf_HTMLTableSTR(m_Rs("T13_GAKUSEI_NO")) %>');"></td>
											</tr>
											<%
											w_iCnt = w_iCnt + 1
											m_Rs.MoveNext
										Loop
									%>

								</table>
						<% else %>
						<!--  学生写真表示　-->

							<table border="0" cellpadding="0" cellspacing="2">
								<%
									w_iCnt = 1
									Do Until m_Rs.Eof or w_iCnt > m_iDsp
										response.write 	"<tr>"
										i_TdLine = 1					'// 横に４件表示ライン
										Do Until m_Rs.Eof or i_TdLine > 4 or w_iCnt > m_iDsp
										%>
											<td align="center" class=search width="150" valign="top">
												<a href="javascript:f_detail('<%= gf_HTMLTableSTR(m_Rs("T13_GAKUSEI_NO")) %>');">
												<%
												'// 顔写真があるか先に検索する
												w_bRet = ""
												w_bRet = f_Photoimg(m_Rs("T13_GAKUSEI_NO"))

												if w_bRet = True then
													' 2011.04.05 upd DispBinary => DispBinaryRec に変更
													%><IMG SRC="DispBinaryRec.asp?gakuNo=<%= m_Rs("T13_GAKUSEI_NO") %>" width="88" height="136" border="0"><%
												Else
													%><IMG SRC="images/Img0000000000.gif" width="100" height="120" border="0"><%
												End if
												%></a><br>
												<table border="0" cellpadding="0" cellspacing="2" width="100%">
													<tr><td bgcolor="#666699" nowrap><font color="#FFFFFF"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></font></td><td><%= m_Rs("T13_GAKUSEKI_NO") %></td></tr>
													<tr><td bgcolor="#666699" nowrap><font color="#FFFFFF">氏名    </font></td><td><%= trim(m_Rs("T11_SIMEI")) %></td></tr>
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

					<% end if %>
				</td>
			</tr>
		</table>

		<%=w_pageBar %>
		</td></tr></table>

		</div>
	    <input type="hidden" name="txtMode">
	    <input type="hidden" name="txtPageTyu" value="<%=m_iPageTyu%>">
	    <input type="hidden" name="hidGAKUSEI_NO">

		<%' 検索条件 %>
		<input type="hidden" name="txtHyoujiNendo" value="<%=request("txtHyoujiNendo")%>">
		<input type="hidden" name="txtGakunen"     value="<%=request("txtGakunen")%>">
		<input type="hidden" name="txtGakka"       value="<%=request("txtGakka")%>">
		<input type="hidden" name="txtClass"       value="<%=request("txtClass")%>">
		<input type="hidden" name="txtName"        value="<%=request("txtName")%>">
		<input type="hidden" name="txtGakusekiNo"  value="<%=request("txtGakusekiNo")%>">
		<input type="hidden" name="txtSeibetu"     value="<%=request("txtSeibetu")%>">
		<input type="hidden" name="txtGakuseiNo"   value="<%=request("txtGakuseiNo")%>">
		<input type="hidden" name="TxtIdou"        value="<%=request("TxtIdou")%>">
		<input type="hidden" name="txtTyuClub"     value="<%=request("txtTyuClub")%>">
		<input type="hidden" name="txtClub"        value="<%=request("txtClub")%>">
		<input type="hidden" name="txtRyoseiKbn"   value="<%=request("txtRyoseiKbn")%>">
		<input type="hidden" name="CheckImage"     value="<%=request("CheckImage")%>">
		<input type="hidden" name="txtTyugaku"     value="<%=request("txtTyugaku")%>">
		<input type="hidden" name="txtDisp"        value="<%=request("txtDisp")%>">
		</form>
	<% End if %>
	</body>

    </html>

<%
    '---------- HTML END   ----------
End Sub

%>

