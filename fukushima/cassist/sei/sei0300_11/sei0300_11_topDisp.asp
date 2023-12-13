<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験毎所見登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/sei0300_11/sei0300_11_topDisp.asp
' 機      能: 
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:
'           ■初期表示
'               上部画面表示のみ
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2006/01/30　西村 綾和子 福島高専用に新規作成
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
	Public CONST C_KENGEN_SEI0300_FULL = "FULL"	'//アクセス権限FULL
	Public CONST C_KENGEN_SEI0300_TAN = "TAN"	'//アクセス権限担任
	Public CONST C_KENGEN_SEI0300_GAK = "GAK"	'//アクセス権限学科

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系

    '市町村選択用のWhere条件
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官コード
    Public m_sGakuNo        '氏名コンボボックスに入る値
    Public m_sBeforGakuNo   '氏名コンボボックスに入る値の一人前
    Public m_sAfterGakuNo   '氏名コンボボックスに入る値の一人後
    Public m_sSsyoken       '総合所見
    Public m_sBikou         '個人備考
    Public m_sSinro         '進路名
    Public m_sSotudai       '卒研課題
    Public m_sSkyokan1      '卒官1
    Public m_sSkyokan2      '卒官2
    Public m_sSkyokan3      '卒官3
    Public m_sGakunen       '学年
    Public m_sClass         'クラス
    Public m_sClassNm       'クラス名
    Public m_sGakusei()     '学生の配列
    Public m_sGakka     '学生の所属学科
    Public m_sShiken
	Public m_sGakkaNo
    public m_sKengen
    Public  m_GRs
    Public  m_Rs
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
    w_sMsgTitle="試験毎所見登録"
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

		'//権限チェック
		w_iRet = f_CheckKengen(w_sKengen)
		If w_iRet <> 0 Then
            m_bErrFlg = True
			m_sErrMsg = "参照権限がありません。"
			Exit Do
		End If

		'//権限が担任の場合は担任クラス情報を取得する
		If w_sKengen = C_KENGEN_SEI0300_TAN Then

			'//担任クラス情報取得
			'//情報が取得できない場合は担任クラスが無い為、参照不可とする
			w_iRet = f_GetClassInfo(m_sKengen)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				m_sErrMsg = "参照権限がありません。"
				Exit Do
			End If

		ElseIf w_sKengen = C_KENGEN_SEI0300_GAK Then

			'//学科情報取得
			'//情報が取得できない場合は学科が無い為、参照不可とする
			w_iRet = f_GetGakkaInfo(m_sKengen)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				m_sErrMsg = "参照権限がありません。"
				Exit Do
			End If

		End If


	  '//試験名を取得
            If f_GetSiken(m_sShiken) <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If

		Call f_Gakusei()

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

    m_iNendo    = cint(session("NENDO"))
    m_sGakuNo   = request("txtGakusei")
    m_iDsp      = C_PAGE_LINE
	m_sGakunen  = Cint(request("txtGakuNo"))
	m_sClass    = Cint(request("txtClassNo"))
	m_sShiken    = request("txtSikenKBN")
	m_sGakkaNo  = Request("txtGakkaNo")

	'//前へOR次へボタンが押された時
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

'********************************************************************************
'*	[機能]	権限チェック
'*	[引数]	なし
'*	[戻値]	w_sKengen
'*	[説明]	ログインUSERの処理レベルにより、参照可不可の判断をする
'*			①FULLアクセス権限保持者は、全ての生徒の成績情報を参照できる
'*			②担任アクセス権限保持者は、受け持ちクラス生徒の成績情報を参照できる
'*			③上記以外のUSERは参照権限なし
'********************************************************************************
Function f_CheckKengen(p_sKengen)
    Dim w_iRet
    Dim w_sSQL
	 Dim rs

	 On Error Resume Next
	 Err.Clear

	 f_CheckKengen = 1

	 Do

		'T51より権限情報取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL.T51_ID "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL.T51_ID IN ('SEI0300','SEI0301','SEI0302')"
		w_sSql = w_sSql & vbCrLf & "  AND T51_SYORI_LEVEL.T51_LEVEL" & Session("LEVEL") & " = 1"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			m_sErrMsg = Err.description
			f_CheckKengen = 99
			Exit Do
		End If

		If rs.EOF Then
			m_sErrMsg = "参照権限がありません。"
			Exit Do
		Else

			Select Case cstr(rs("T51_ID"))
				Case "SEI0300"	'//フルアクセス権限あり
					p_sKengen = C_KENGEN_SEI0300_FULL
				Case "SEI0301"	'//担任権限有り
					p_sKengen = C_KENGEN_SEI0300_TAN
				Case "SEI0302"	'//学科権限有り
					p_sKengen = C_KENGEN_SEI0300_GAK
			End Select

		End If

		f_CheckKengen = 0
		Exit Do
	 Loop


	Call gf_closeObject(rs)

End Function
'********************************************************************************
'*  [機能]  権限チェック（担任クラス情報取得）
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  ○担任アクセス権限が設定されているUSERでも、実際に担任クラスを
'*			受け持っていない場合には参照不可とする
'********************************************************************************
Function f_GetClassInfo(p_sKengen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassInfo = 1
	p_sKengen = ""

	Do 

		'// 担任クラス情報
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS.M05_CLASSNO "
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & session("KYOKAN_CD") & "'"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetClassInfo = 99
			Exit Do
		End If

		If rs.EOF Then
			'//クラス情報が取得できないとき
            m_sErrMsg = "参照権限がありません。"
			Exit Do
		End If

		f_GetClassInfo = 0
		p_sKengen = C_KENGEN_SEI0300_TAN
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  権限チェック（ユーザ学科情報取得）
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  ○担任アクセス権限が設定されているUSERでも、実際に担任クラスを
'*			受け持っていない場合には参照不可とする
'********************************************************************************
Function f_GetGakkaInfo(p_sKengen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetGakkaInfo = 1
	p_sKengen = ""

	Do 

		'// 担任クラス情報
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M04_GAKKA_CD "
		w_sSQL = w_sSQL & vbCrLf & " FROM M04_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M04_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M04_KYOKAN_CD='" & session("KYOKAN_CD") & "'"
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetGakkaInfo = 99
			Exit Do
		End If
		If rs.EOF Then
			'//クラス情報が取得できないとき
            m_sErrMsg = "参照権限がありません。"
			Exit Do
		Else
			p_sKengen = C_KENGEN_SEI0300_GAK 
'			m_sGakkaNo  = rs("M04_GAKKA_CD")
'			m_sGakkaMei = rs("M02_GAKKAMEI")

			'//権限が担任の場合は、担任クラス以外は選択できない
'			m_sGakuNoOption = " DISABLED "
'			m_sClassNoOption = " DISABLED "
		End If

		f_GetGakkaInfo = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function


Function f_Gakusei()
'********************************************************************************
'*  [機能]  学生データを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim i
i = 1


    w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

	'//学生の情報収集
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "

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
    w_sSQL = w_sSQL & "     B.T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND B.T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND B.T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
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
        w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo
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

'********************************************************************************
'*  [機能]  試験コンボを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSiken(p_sShiken)
    Dim w_sSQL,w_Rs

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & " M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_SYOBUNRUI_CD = " & cint(p_sShiken)

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If
	p_sShiken = w_Rs("M01_SYOBUNRUIMEI")

        f_GetSiken = 0
        Exit Do
    Loop
	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  学科の略称を取得
'*  [引数]  p_sGakkaCd : 学科CD
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetGakkaNm(p_iGakunen,p_iClass)
    Dim w_sSQL              '// SQL文
    Dim w_iRet              '// 戻り値
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakkaNm = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKAMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKA_CD = M05_CLASS.M05_GAKKA_CD "
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_NENDO = M05_CLASS.M05_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN=" & p_iGakunen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_CLASSNO=" & p_iClass

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M02_GAKKAMEI")
		End If 

		Exit do 
	Loop

	'//戻り値をセット
	f_GetGakkaNm = w_sName

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

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

%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type="text/css">

<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="JavaScript">
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
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">
<center>
	<% call gs_title(" 個人別成績一覧 "," 一　覧 ") %>
<BR>

<table border="0" width="500" class=hyo align="center">
	<tr>
		<th width="500" class="header2" colspan="4"><%=m_sShiken%></th>
	</tr>
	<tr>
		<th width="50" class="header">クラス</th>

		<%If m_sKengen <> C_KENGEN_SEI0300_GAK then%>
			<td width="150" align="center" class="detail"><%=m_sGakunen%>-<%=m_sClass%> [<%=f_GetGakkaNm(m_sGakunen,m_sClass)%>]</td>
		<%Else%>
			<td width="150" align="center" class="detail"><%=m_sGakunen%>年　<%=gf_GetGakkaNm(m_iNendo,m_sGakkaNo)%></td>
		<%End If%>


		<th width="50" class="header">氏　名</th>
		<td width="250" align="left" class="detail">　( <%=f_getGakuseki_No() & " )　" & m_GRs("T11_SIMEI")%></td>

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
