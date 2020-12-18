<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 調査書所見等登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0460/gak0460_main.asp
' 機      能: 下ページ 調査書所見等登録の検索を行う
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
' 変      更：2001/08/30 伊藤 公子     検索条件を2重に表示しないように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '氏名選択用のWhere条件
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官コード
    Public m_sNendo         '年度コンボボックスに入る値
    Public m_sGakuNo        '氏名コンボボックスに入る値
    Public m_sBeforGakuNo   '氏名コンボボックスに入る値の一人前
    Public m_sAfterGakuNo   '氏名コンボボックスに入る値の一人後
    Public m_sGakunen       '学年
    Public m_sClass         'クラス
    Public m_sClassNm       'クラス名
    Public m_sGakusei()     '学生の配列

    Public  m_TRs           
    Public  m_GRs           
    Public  m_URs
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
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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

        '//データ取得
        w_iRet = f_Gakusei()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		If m_GRs.EOF Then
			Call NO_Showpage()
			Exit Do
		End If

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

Sub s_SetParam()
'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_sNendo    = request("txtNendo")
    m_sGakuNo   = request("txtGakuNo")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm  = request("txtClassNm")
    m_iDsp      = C_PAGE_LINE

	'//前へOR次へボタンが押された時
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

Function f_Gakusei()
'********************************************************************************
'*  [機能]  生徒情報を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim i
i = 1


    w_iNyuNendo = Cint(m_sNendo) - Cint(m_sGakunen) + 1

	'//学生の情報収集
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
'    w_sSQL = w_sSQL & " AND T11_NYUNENDO = " & w_iNyuNendo & " "

    Set m_GRs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_GRs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
    End If


    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     A.T11_GAKUSEI_NO "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI A,T13_GAKU_NEN B "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     B.T13_NENDO = " & m_sNendo & " "
    w_sSQL = w_sSQL & " AND B.T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND B.T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
'    w_sSQL = w_sSQL & " AND A.T11_NYUNENDO = B.T13_NENDO - B.T13_GAKUNEN + 1"
    w_sSQL = w_sSQL & " ORDER BY B.T13_GAKUSEKI_NO "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
    End If
	w_rCnt=cint(gf_GetRsCount(w_Rs))

	'//配列の作成

		w_Rs.MoveFirst

       Do Until w_Rs.EOF

            ReDim Preserve m_sGakusei(i)
            m_sGakusei(i) = w_Rs("T11_GAKUSEI_NO")
            i = i + 1
            
            w_Rs.MoveNext
            
        Loop

		For i = 1 to w_rCnt

			If m_sGakusei(i) = m_sGakuNo Then

				If i <= 1 Then
					m_sGakuNo      = m_sGakusei(i)
	                m_sAfterGakuNo = m_sGakusei(i+1)
					Exit For
				End If

				If i = w_rCnt Then
					m_sGakuNo      = m_sGakusei(i)
	                m_sBeforGakuNo = m_sGakusei(i-1)
					Exit For
				End If

				m_sGakuNo      = m_sGakusei(i)
                m_sAfterGakuNo = m_sGakusei(i+1)
                m_sBeforGakuNo = m_sGakusei(i-1)
				
				Exit For
			End If

		Next

End Function

Function f_getGakuseki_No()
'********************************************************************************
'*  [機能]  学生の学籍NOを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	Dim rs
	Dim w_sSQL

    On Error Resume Next
    Err.Clear

    f_getGakuseki_No = ""

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     T13_GAKUSEKI_NO"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_NENDO = " & m_sNendo
        w_sSQL = w_sSQL & "     AND T13_GAKUSEI_NO = '" & m_sGakuNo & "' "

        w_iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            Exit Do 
        End If

		If rs.EOF = False Then
			w_iGakusekiNo = rs("T13_GAKUSEKI_NO")
		End If

        Exit Do
    Loop

	'//戻り値セット
    f_getGakuseki_No = w_iGakusekiNo

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

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
    <title>調査書所見等登録</title>
<link rel=stylesheet href="../../common/style.css" type=text/css>

<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="JavaScript">
<!--

//-->
</SCRIPT>

</head>
<body>
<form name="frm" method="post" onClick="return false;">
<center>
<%call gs_title("調査書所見等登録","登　録")%>

<BR>
<table border="0" width="570" class=hyo align="center">
	<tr>
		<th width="50" class="header">年度</th>
		<td width="100" align="center" class="detail"><%=m_sNendo%>年度</td>
		<th width="50" class="header">クラス</th>
		<td width="120" align="center" class="detail"><%=m_sGakunen%>年<%=m_sClassNm%></td>
		<th width="50" class="header">氏　名</th>
		<!--<td width="150" align="left" class="detail"><%=m_GRs("T11_SIMEI")%></td>-->
		<td width="200" align="left" class="detail">　( <%=f_getGakuseki_No() & " )　" & m_GRs("T11_SIMEI")%></td>
	</tr>
</table>
<br>
<div align="center"><span class=CAUTION>※ ｢前へ｣｢次へ｣のボタンをクリックした場合、入力されたものが保存され、<br>
										現在入力されている学生の前または、後の学生の情報入力に移ります。
</span></div>
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
