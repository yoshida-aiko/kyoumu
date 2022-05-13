<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 追試受講者一覧
' ﾌﾟﾛｸﾞﾗﾑID : saisi/saisi0400/saisi0400_show.asp
' 機      能: 追試受講者一覧 科目一覧
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
' 変      更: 2003/02/27  矢野
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
'エラー系
Public  m_bErrFlg          'ｴﾗｰﾌﾗｸﾞ

Dim m_Rs		'recordset

Dim m_iNendo             '年度
Dim m_sKyokanCd          '教官コード

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
    w_sMsgTitle="追試受講者一覧"
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
            m_sErrMsg = "追試科目の取得に失敗しました。"
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
'	w_sSql = ""
'	w_sSql = w_sSql & " SELECT DISTINCT"
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, "
'	w_sSql = w_sSql & "	T16_KAMOKU_KBN,"
'	w_sSql = w_sSql & "	T16_HISSEN_KBN,"
'	w_sSql = w_sSql & "	T16_COURSE_CD,"
'	w_sSql = w_sSql & "	T16_SEQ_NO,"
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, "
'	w_sSql = w_sSql & "	T120_KAMOKUMEI "
'	
'	w_sSql = w_sSql & " FROM "
'	w_sSql = w_sSql & "	T120_SAISIKEN "
'	w_sSql = w_sSql & "	,T16_RISYU_KOJIN "
'	
'	w_sSql = w_sSql & "	WHERE "
'	w_sSql = w_sSql & "	T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'教官
'	w_sSql = w_sSql & "	AND	T120_SYUTOKU_NENDO Is Null "				'未修得
'	w_sSql = w_sSql & " AND T120_GAKUSEI_NO  = T16_GAKUSEI_NO"			
'	w_sSql = w_sSql & "	AND T120_KAMOKU_CD = T16_KAMOKU_CD"
'	w_sSql = w_sSql & "	AND T120_NENDO = T16_NENDO "
''	w_sSql = w_sSql & "	AND T16_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'成績不可（評価=*）
'	w_sSql = w_sSql & "	AND T120_SEISEKI Is Null "
'	
'	w_sSql = w_sSql & "	GROUP BY"
'	w_sSql = w_sSql & "	T16_KAMOKU_KBN,"
'	w_sSql = w_sSql & "	T16_HISSEN_KBN,"
'	w_sSql = w_sSql & "	T16_COURSE_CD,"
'	w_sSql = w_sSql & "	T16_SEQ_NO,"
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, "
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, "
'	w_sSql = w_sSql & "	T120_KAMOKUMEI "
'
'	
'	w_sSql = w_sSql & "	UNION"
'	
'	'過去データ
'	w_sSql = w_sSql & " SELECT DISTINCT"
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, "
'	w_sSql = w_sSql & "	T17_KAMOKU_KBN,"
'	w_sSql = w_sSql & "	T17_HISSEN_KBN,"
'	w_sSql = w_sSql & "	T17_COURSE_CD,"
'	w_sSql = w_sSql & "	T17_SEQ_NO,"
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, "
'	w_sSql = w_sSql & "	T120_KAMOKUMEI "
'	
'	w_sSql = w_sSql & " FROM "
'	w_sSql = w_sSql & "	T120_SAISIKEN "
'	w_sSql = w_sSql & "	,T17_RISYUKAKO_KOJIN "
'	
'	w_sSql = w_sSql & "	WHERE "
'	w_sSql = w_sSql & "	T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'教官
'	w_sSql = w_sSql & "	AND	T120_SYUTOKU_NENDO Is Null "				'未修得
'	w_sSql = w_sSql & " AND T120_GAKUSEI_NO  = T17_GAKUSEI_NO"			
'	w_sSql = w_sSql & "	AND T120_KAMOKU_CD = T17_KAMOKU_CD"
'	w_sSql = w_sSql & "	AND T120_NENDO = T17_NENDO "
''	w_sSql = w_sSql & "	AND T17_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'成績不可（評価=*）
'	w_sSql = w_sSql & "	AND T120_SEISEKI Is Null "
'
'	w_sSql = w_sSql & "	GROUP BY"
'	w_sSql = w_sSql & "	T17_KAMOKU_KBN,"
'	w_sSql = w_sSql & "	T17_HISSEN_KBN,"
'	w_sSql = w_sSql & "	T17_COURSE_CD,"
'	w_sSql = w_sSql & "	T17_SEQ_NO,"
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, "
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, "
'	w_sSql = w_sSql & "	T120_KAMOKUMEI "
'	
'	w_sSql = w_sSql & "	ORDER BY 1,2,3,4,5"		'学年、科目区分、必選区分、コースコード、SEQNO

	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T120_MISYU_GAKUNEN, "
	w_sSql = w_sSql & "		T120_KAMOKU_CD, "
	w_sSql = w_sSql & "		T120_KAMOKUMEI "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		    T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'教官
	w_sSql = w_sSql & "		AND	T120_SYUTOKU_NENDO Is Null "					'未修得
	w_sSql = w_sSql & "		AND	T120_NENDO = " & Session("NENDO")				'年度
	w_sSql = w_sSql & "		AND	T120_SEISEKI Is Null "					'期末点数
	w_sSql = w_sSql & " GROUP BY "
	w_sSql = w_sSql & "		T120_MISYU_GAKUNEN, "
	w_sSql = w_sSql & "		T120_KAMOKU_CD, "
	w_sSql = w_sSql & "		T120_KAMOKUMEI "
		
    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

'Response.write gf_GetRsCount(m_Rs) & "<br>"
'Response.Write session("KYOKAN_CD") & "<br>"

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
	Dim w_iTableKbn
	Dim w_sCellClass

%>
<html>

<head>
<meta http-equiv="Content-Language" content="ja">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../common/style.css" type="text/css">
<title>追試受講者一覧</title>

<script language="JavaScript">
<!--

//===============================================
//	送信処理
//===============================================
function jf_Submit(p_iNumber,p_iGakunen,p_sName) {

	document.frm.hidKAMOKU_CD.value = p_iNumber;
	document.frm.hidMISYU_GAKUNEN.value = p_iGakunen;
	document.frm.hidKAMOKU_MEI.value = p_sName;
	document.frm.action = "saisi0400_Report.asp";
	document.frm.submit();
	return

}

//-->
</script>

</head>

<body>

<form name="frm">
<center>
<br>
<br>
<br>
<table border="1" class="hyo" >

	<!-- TABLEヘッダ部 -->
	<tr>
		<th width="70"  class="header3" align="center" height="24">履修学年</th>
		<th width="200" class="header3" align="center" height="24">科　　　目</th>
		<th width="70"  class="header3" align="center" height="24">一　　　覧</th>
	</tr>
      
	<!-- TABLEリスト部 -->      
<%

	'TDのCLASSの初期化
	w_sCellClass = "CELL2"

	do until m_Rs.EOF
%>
    <tr>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_MISYU_GAKUNEN")%></td>
		<td width="200" class="<%=w_sCellClass%>" align="left" height="24">　<%=m_Rs("T120_KAMOKUMEI")%></td>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="button" value=" 表　示 " onclick="jf_Submit('<%=m_Rs("T120_KAMOKU_CD")%>','<%=m_Rs("T120_MISYU_GAKUNEN")%>','<%=m_Rs("T120_KAMOKUMEI")%>')">
		</td>
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

<!-- 引数格納エリア -->
<input type="hidden" name="hidKAMOKU_CD">
<input type="hidden" name="hidKAMOKU_MEI">
<input type="hidden" name="hidMISYU_GAKUNEN">

</form>

</body>

</html>
<%
end sub
%>