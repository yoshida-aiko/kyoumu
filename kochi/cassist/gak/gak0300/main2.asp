<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学籍データ検索結果(画像表示
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/main2.asp
' 機      能: 下ページ 学籍データの検索結果を画像表示する
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
'           txtMode                :動作モード
'                               BLANK   :初期表示
'                               SEARCH  :結果表示
' 説      明:
'           ■初期表示
'               タイトルのみ表示
'           ■結果表示
'               上ページで設定された検索条件にかなう学生情報を画像表示する
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '取得したデータを持つ変数
    Public  m_TxtMode      	       ':動作モード
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
    
    Public	m_Rs					'recordset
    Public	m_iDsp					'// 一覧表示行数

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
    w_sMsgTitle="学籍データ検索結果"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	'//セッション情報・動作モードの取得
    m_TxtMode=request("txtMode")
    
    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
		w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

        '//初期表示
        if m_TxtMode = "" then
            Call showPage()
            Exit Do
        end if

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        'データ抽出SQLを作成する
        Call s_MakeSQL(w_sSQL)

        'レコードセットの取得
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// ページを表示
        If m_Rs.EOF Then
            Call showPage_NoData()
        Else
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

End Sub

Sub s_SetParam()
'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_iHyoujiNendo =request("txtHyoujiNendo")     	'表示年度
    m_sGakunen=request("txtGakunen")            	'学年
	'コンボ未選択時
	If m_sGakunen="@@@" Then
		m_sGakunen=""
	End If
    m_sGakkaCD=request("txtGakka")             		'学科
	'コンボ未選択時
	If m_sGakkaCD="@@@" Then
		m_sGakkaCD=""
	End If

	if m_sGakunen="" then		'学年が選択されていない場合はクラスは選択できません
		m_sClass=""
	else 
    	m_sClass=request("txtClass")               		'クラス
		'コンボ未選択時
		If m_sClass="@@@" Then
			m_sClass=""
		End If
    end if 
	m_sName=request("txtName")                		'名称
	m_sGakusekiNo=request("txtGakusekiNo")          '学籍番号
	m_sSeibetu=request("txtSeibetu")            	'性別
	'コンボ未選択時
	If m_sSeibetu="@@@" Then
		m_sSeibetu=""
	End If
	m_sGakuseiNo=request("txtGakuseiNo")           	'学生番号
	m_sIdou =request("TxtIdou")               		'異動
	'コンボ未選択時
	If m_sIdou="@@@" Then
		m_sIdou=""
	End If
	m_sTyuClub =request("txtTyuClub")            	'中学校クラブ
	'コンボ未選択時
	If m_sTyuClub="@@@" Then
		m_sTyuClub=""
	End If
	m_sClub=request("txtClub")                		'現在クラブ
	'コンボ未選択時
	If m_sClub="@@@" Then
		m_sClub=""
	End If
	m_sRyoseiKbn=request("txtRyoseiKbn")           	'寮
	'コンボ未選択時
	If m_sRyoseiKbn="@@@" Then
		m_sRyoseiKbn=""
	End If

End Sub

Sub s_MakeSQL(p_sSql)
'********************************************************************************
'*  [機能]  学籍データ抽出SQL文字列の作成
'*  [引数]  p_sSql - SQL文字列
'*  [戻値]  なし 
'*  [説明]  
'********************************************************************************

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " T13_GAKUSEKI_NO, "
    p_sSql = p_sSql & " T13_GAKUSEI_NO, "
    p_sSql = p_sSql & " T13_GAKUNEN, "
    p_sSql = p_sSql & " T13_CLASS, "
    p_sSql = p_sSql & " T11_SIMEI, "
    p_sSql = p_sSql & " T11_SEIBETU, "
    p_sSql = p_sSql & " M01_SYOBUNRUIMEI, "
    p_sSql = p_sSql & " M02_GAKKARYAKSYO "
    p_sSql = p_sSql & " FROM T13_GAKU_NEN, T11_GAKUSEKI, M02_GAKKA, M01_KUBUN "

    p_sSql = p_sSql & " WHERE T13_NENDO = '" & cstr(m_iHyoujiNendo) & "'"

    '検索条件のセット
    if m_sGakunen <> "" then        '学年
        p_sSql = p_sSql & " AND T13_GAKUNEN = " & cint(m_sGakunen)
    end if
    if m_sGakkaCD <> "" then         '学科
        p_sSql = p_sSql & " AND T13_GAKKA_CD = '" & m_sGakkaCD & "'"
    end if
    if m_sClass <> "" then           'クラス
        p_sSql = p_sSql & " AND T13_CLASS = '" & m_sClass & "'"
    end if
    if m_sName <> "" then            '名称
        p_sSql = p_sSql & " AND T11_SIMEI_KD LIKE '%" & m_sName & "%'"
    end if
    if m_sGakusekiNo <> "" then 	'学籍番号
    	p_sSql = p_sSql & " AND T13_GAKUSEKI_NO LIKE '%" & m_sGakusekiNo & "%'"
    end if
    if m_sSeibetu <> "" then 		'性別
    	p_sSql = p_sSql & " AND T11_SEIBETU = " & m_sSeibetu
    end if
    if m_sGakuseiNo <> "" then 		'学生番号
        p_sSql = p_sSql & " AND T13_GAKUSEI_NO LIKE '%" & m_sGakuseiNo & "%'"
    end if
    if m_sIdou <> "" then 		'異動
        p_sSql = p_sSql & " AND ( T13_IDOU_KBN_1 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_2 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_3 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_4 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_5 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_6 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_7 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_8 = '" & m_sIdou & "' )"
    end if

    if m_sTyuClub <> "" then 		'中学校クラブ
	    p_sSql = p_sSql & " AND T11_TYU_CLUB = '" & m_sTyuClub & "'"
    end if
    if m_sClub <> "" then 		'現在クラブ
	    p_sSql = p_sSql & " AND ( T13_CLUB_1 = '" & m_sClub & "'"
	    p_sSql = p_sSql & " OR T13_CLUB_2 = '" & m_sClub & "' ) "
    end if
    if m_sRyoseiKbn <> "" then 		'寮
        p_sSql = p_sSql & " AND T13_RYOSEI_KBN = '" & m_sRyoseiKbn & "'"
    end if
    
    '結合条件
    p_sSql = p_sSql & " AND T13_GAKUSEI_NO = T11_GAKUSEI_NO "
    p_sSql = p_sSql & " AND M02_NENDO(+) = '" & cstr(m_iHyoujiNendo) & "'"
    p_sSql = p_sSql & " AND M02_GAKKA_CD(+) = T13_GAKKA_CD "
    p_sSql = p_sSql & " AND M01_NENDO(+) = '" & cstr(m_iHyoujiNendo) & "'"
    p_sSql = p_sSql & " AND M01_DAIBUNRUI_CD(+) = 1 "
    p_sSql = p_sSql & " AND M01_SYOBUNRUI_CD(+) = T11_SEIBETU "

    p_sSql = p_sSql & " ORDER BY T13_GAKUNEN, M01_SYOBUNRUIMEI,"
    p_sSql = p_sSql & " T13_CLASS, T13_GAKUSEKI_NO, T13_GAKUSEI_NO "

'response.write " p_sSql=" & p_sSql & "<BR>"

End Sub

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
    function f_detail(){

        window.alert("詳細");
		document.forms[0].submit();
    }

    //-->
    </SCRIPT>

    </head>

    <body>

    <center>
    <form action="kojin.asp" method="post" name="frm" target="_detail">

	<% if m_TxtMode = "" then %>
	<table border="0" width="100%">
	
		<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
	<tr>
		<td width="60">&nbsp</td>
		<td valign="top"></td>
	</tr>
	</table>
	
	<% else %>
	
	<tr>
		<td align="center">

		<table border="1" width="600" class=hyo>
		<tr><!-- １年間番号  -->
		<th height=16 class=header><%=gf_GetGakuNomei(m_iHyoujiNendo,C_K_KOJIN_1NEN)%></th>
		<th height=16 class=header>学生写真</th>
		</tr>
		
        <% Do While not m_Rs.EOF %>
		<% dim w_cell
		   call gs_cellPtn(w_cell)%>
		<tr>
		<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEKI_NO")) %>&nbsp</td>
		<td >
<!--
                   <IMG SRC="./DispBinary.asp?gakuNo=0000000001"> -->
                   <!-- <IMG SRC="./DispBinary.asp?gakuNo=<%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEKI_NO")) %>"> -->
                </td>
		</tr>
		
		<%   m_Rs.MoveNext
		Loop %>
		
	    </table>
		</td>
		</tr>
		
		<% end if %>

	</center>

	</body>
    </html>


<%
    '---------- HTML END   ----------
End Sub

%>



















































