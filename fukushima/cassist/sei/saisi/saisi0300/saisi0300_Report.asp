<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 再試受講者一覧
' ﾌﾟﾛｸﾞﾗﾑID : saisi/saisi0300/saisi0300_Report.asp
' 機      能: 再試受講者一覧 科目一覧
'-------------------------------------------------------------------------
' 引      数    
'               
' 変      数
' 引      渡
'           
'           
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2003/02/20  松尾
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
Dim m_Rs		'recordset

Dim m_iNendo				'年度
Dim m_sKyokanCd				'教官コード

dim m_sKamokuCD				'科目CD
dim m_iGakunen				'学年
dim m_sKamokuMei			'科目名

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

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="再試受講者一覧"
	w_sMsg=""
	w_sRetURL="../default.asp"
	w_sTarget="_parent"

	Do
		'// ﾃﾞｰﾀﾍﾞｰｽ接続
		w_iRet = gf_OpenDatabase()
		If w_iRet <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If

		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//値を取得
		call s_SetParam()

		'// 未修得科目取得
		if wf_GetStudent() = false then
			m_bErrFlg = True
			m_sErrMsg = "再試科目の取得に失敗しました。"
			Exit Do
		end if

		Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If

	'//初期表示
	Call showPage()

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
	
	gDisabled = ""
	
    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
	
	m_sKamokuCD = request("hidKAMOKU_CD")
	m_iGakunen = cint(request("hidMISYU_GAKUNEN"))
	m_sKamokuMei = request("hidKAMOKU_MEI")
	
End Sub


function wf_GetStudent()
'********************************************************************************
'*  [機能]  未修得科目取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	'変数の宣言
	Dim w_sSql
	Dim w_iRet

	wf_GetStudent = false

	'今年度データ
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T13_GAKUNEN,"
	w_sSql = w_sSql & "		M05_CLASSMEI,"
	w_sSql = w_sSql & "		T13_GAKUSEKI_NO,"
	w_sSql = w_sSql & "		T11_SIMEI,"
	w_sSql = w_sSql & "		T120_NENDO, "
	w_sSql = w_sSql & "		T120_JYUKO_FLG, "
	w_sSql = w_sSql & "		T120_JYUKOKAISU, "
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1, "
	w_sSql = w_sSql & "		MAX(M08_MAX) "			'不可評価の点数が高い方
	
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "		,T16_RISYU_KOJIN "
	w_sSql = w_sSql & "		,M08_HYOKAKEISIKI "
	w_sSql = w_sSql & "		,T13_GAKU_NEN "
	w_sSql = w_sSql & "		,M05_CLASS "
	w_sSql = w_sSql & "		,T11_GAKUSEKI "
	
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'教官
	w_sSql = w_sSql & "		AND T120_KAMOKU_CD = '" & m_sKamokuCD & "' "	'科目CD
	w_sSql = w_sSql & "		AND	T120_SYUTOKU_NENDO Is Null "				'未修得
	w_sSql = w_sSql & " 	AND T120_GAKUSEI_NO  = T16_GAKUSEI_NO"			
	w_sSql = w_sSql & "		AND T120_KAMOKU_CD = T16_KAMOKU_CD"
	w_sSql = w_sSql & "		AND T120_NENDO = T16_NENDO "
	w_sSql = w_sSql & "		AND T16_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'成績不可（評価=*）
	'開設時期によって評価点を区分(前期開設はT16_HTEN_KIMATU_Z,他はT16_HTEN_KIMATU_K)
	w_sSql = w_sSql & " 	AND DECODE(T16_KAISETU ," & C_KAI_ZENKI & ",T16_HTEN_KIMATU_Z,T16_HTEN_KIMATU_K)  IS NOT NULL"	
	w_sSql = w_sSql & "		AND DECODE(T16_KAISETU ," & C_KAI_ZENKI & ",T16_HTEN_KIMATU_Z,T16_HTEN_KIMATU_K) <= M08_MAX"
	w_sSql = w_sSql & "		AND DECODE(T16_KAISETU ," & C_KAI_ZENKI & ",T16_HTEN_KIMATU_Z,T16_HTEN_KIMATU_K) >= M08_MIN"
	w_sSql = w_sSql & "		AND M08_NENDO = T120_NENDO"			'履修年度の評価
	w_sSql = w_sSql & " 	AND M08_HYOUKA_NO = 2"
	w_sSql = w_sSql & "		AND M08_HYOKA_TAISYO_KBN = 0"
	w_sSql = w_sSql & "		AND M08_HYOKA_SYOBUNRUI_RYAKU = 1"		'不可の科目

	'処理年度のT13とT120結合
	w_sSql = w_sSql & "		AND T13_NENDO = " & m_iNendo
	w_sSql = w_sSql & "		AND T13_GAKUSEI_NO = T120_GAKUSEI_NO"
	'クラスM結合
	w_sSql = w_sSql & "		AND M05_NENDO = T13_NENDO"	
	w_sSql = w_sSql & "		AND M05_GAKUNEN = T13_GAKUNEN"	
	w_sSql = w_sSql & "		AND M05_CLASSNO = T13_CLASS"	
	'T11とT13結合
	w_sSql = w_sSql & "		AND T11_GAKUSEI_NO = T13_GAKUSEI_NO"

	
	w_sSql = w_sSql & "	GROUP BY"
	w_sSql = w_sSql & "		T13_GAKUNEN,"
	w_sSql = w_sSql & "		M05_CLASSMEI,"
	w_sSql = w_sSql & "		T13_GAKUSEKI_NO,"
	w_sSql = w_sSql & "		T11_SIMEI,"
	w_sSql = w_sSql & "		T120_NENDO, "
	w_sSql = w_sSql & "		T120_JYUKO_FLG, "
	w_sSql = w_sSql & "		T120_JYUKOKAISU, "
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "

	
	w_sSql = w_sSql & "	UNION"
	
	'過去データ
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T13_GAKUNEN,"
	w_sSql = w_sSql & "		M05_CLASSMEI,"
	w_sSql = w_sSql & "		T13_GAKUSEKI_NO,"
	w_sSql = w_sSql & "		T11_SIMEI,"
	w_sSql = w_sSql & "		T120_NENDO, "
	w_sSql = w_sSql & "		T120_JYUKOKAISU, "
	w_sSql = w_sSql & "		T120_JYUKO_FLG, "
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1, "
	w_sSql = w_sSql & "		MAX(M08_MAX) "			'不可評価の点数が高い方
	
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "		,T17_RISYUKAKO_KOJIN "
	w_sSql = w_sSql & "		,M08_HYOKAKEISIKI "
	w_sSql = w_sSql & "		,T13_GAKU_NEN "
	w_sSql = w_sSql & "		,M05_CLASS "
	w_sSql = w_sSql & "		,T11_GAKUSEKI "
	
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'教官
	w_sSql = w_sSql & "		AND T120_KAMOKU_CD = '" & m_sKamokuCD & "' "	'科目CD
	w_sSql = w_sSql & "		AND	T120_SYUTOKU_NENDO Is Null "				'未修得
	w_sSql = w_sSql & "		AND T120_GAKUSEI_NO  = T17_GAKUSEI_NO"			
	w_sSql = w_sSql & "		AND T120_KAMOKU_CD = T17_KAMOKU_CD"
	w_sSql = w_sSql & "		AND T120_NENDO = T17_NENDO "
	w_sSql = w_sSql & "		AND T17_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'成績不可（評価=*）
	'過去は後期期末評価点
	w_sSql = w_sSql & " 	AND T17_HTEN_KIMATU_K  IS NOT NULL"	
	w_sSql = w_sSql & "		AND T17_HTEN_KIMATU_K <= M08_MAX"
	w_sSql = w_sSql & "		AND T17_HTEN_KIMATU_K >= M08_MIN"
	w_sSql = w_sSql & "		AND M08_NENDO = T120_NENDO"			'履修年度の評価
	w_sSql = w_sSql & " 	AND M08_HYOUKA_NO = 2"
	w_sSql = w_sSql & "		AND M08_HYOKA_TAISYO_KBN = 0"
	w_sSql = w_sSql & "		AND M08_HYOKA_SYOBUNRUI_RYAKU = 1"		'不可の科目
	
	'処理年度のT13とT120結合
	w_sSql = w_sSql & "		AND T13_NENDO = " & m_iNendo
	w_sSql = w_sSql & "		AND T13_GAKUSEI_NO = T120_GAKUSEI_NO"
	'クラスM結合
	w_sSql = w_sSql & "		AND M05_NENDO = T13_NENDO"	
	w_sSql = w_sSql & "		AND M05_GAKUNEN = T13_GAKUNEN"	
	w_sSql = w_sSql & "		AND M05_CLASSNO = T13_CLASS"	
	'T11とT13結合
	w_sSql = w_sSql & "		AND T11_GAKUSEI_NO = T13_GAKUSEI_NO"

	w_sSql = w_sSql & "	GROUP BY"
	w_sSql = w_sSql & "		T13_GAKUNEN,"
	w_sSql = w_sSql & "		M05_CLASSMEI,"
	w_sSql = w_sSql & "		T13_GAKUSEKI_NO,"
	w_sSql = w_sSql & "		T11_SIMEI,"
	w_sSql = w_sSql & "		T120_NENDO, "
	w_sSql = w_sSql & "		T120_JYUKOKAISU, "
	w_sSql = w_sSql & "		T120_JYUKO_FLG, "
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "

	w_sSql = w_sSql & "	ORDER BY"
	w_sSql = w_sSql & "		T13_GAKUNEN,"	
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "

		
    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function 
    End If


	wf_GetStudent = true

end function

sub showPage()
'********************************************************************************
'*  [機能]  画面の表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

'変数の宣言
	Dim w_iJukoFlg
	Dim w_sCellClass
	Dim w_sJyuko 
	
	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="ja">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../common/style.css" type="text/css">
<title>再試受講者一覧</title>

<script language="JavaScript">
<!--



//================================================
//	戻る処理
//================================================
function jf_Back() {

	location.href = "saisi0300_show.asp";
	return;

}

//-->
</script>

</head>

<body>

<form name="frm">
<center>
<br>
<table border="1" class="hyo">
	<tr>
		<td width="70"  class="header3" align="center"  height="16"><font color="#FFFFFF">履修学年</font></td>
		<td width="70"  class="CELL2"   height="16" align="center"><%=m_iGakunen%></td>
		<td width="70"  class="header3" align="center"  height="16"><font color="#FFFFFF">科　　目</font></td>
		<td width="200" class="CELL2"   height="16" align="center"><%=m_sKamokuMei%></td>
	</tr>
</table>

<br>
<br>
<table border="1" class="hyo" >

	<!-- TABLEヘッダ部 -->
	<tr>
		<td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">学年</font></td>
        <td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">クラス</font></td>
        <td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">学籍番号</font></td>
        <td width="200" class="header3" align="center" height="24"><font color="#FFFFFF">氏　　　名</font></td>
        <td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">履修年度</font></td>
        <td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">受験回数</font></td>
        <td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">受験届出</font></td>
    </tr>

      
	<!-- TABLEリスト部 -->      
<%

	'TDのCLASSの初期化
	w_sCellClass = "CELL2"

	do until m_Rs.EOF
	
	'受講フラグチェック
	w_iJukoFlg = cint(gf_SetNull2Zero(m_Rs("T120_JYUKO_FLG")))		'cintがないとエラーになる
	IF w_iJukoFlg = 1 then
		w_sJyuko = "○"
	ELSE
		w_sJyuko = "　"
	END IF
	
%>
   <tr>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=gf_HTMLTableSTR(m_Rs("T13_GAKUNEN"))%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=gf_HTMLTableSTR(m_Rs("M05_CLASSMEI"))%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEKI_NO"))%></font></td>
        <td width="200" class="<%=w_sCellClass%>" align="left"   height="24">　<%=gf_HTMLTableSTR(m_Rs("T11_SIMEI"))%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_NENDO")%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=gf_SetNull2Zero(m_Rs("T120_JYUKOKAISU"))%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=w_sJyuko%></font></td>    </tr>
	</tr>
<%
		m_Rs.MoveNext
		
		if w_sCellClass = "CELL2" then
			w_sCellClass = "CELL1"
		else
			w_sCellClass = "CELL2"
		end if
		
	loop
%>
</table>

<table>
	<tr>
		<td><input type="button" value=" 戻　る " onclick="jf_Back();"></td>
	</tr>
</table>

</center>

</form>

</body>

</html>
<%
end sub
%>