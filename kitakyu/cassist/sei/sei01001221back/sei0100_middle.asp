<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0100_middle.asp
' 機      能: 下ページ 成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 説      明:
'           ■初期表示
'				コンボボックスは空白で表示
'			■表示ボタンクリック時
'				下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/26 前田 智史
' 変      更: 2001/08/21 伊藤 公子
' 変      更: 2001/08/21 伊藤 公子 ヘッダ部切り離し
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bErrMsg           'ｴﾗｰﾒｯｾｰｼﾞ

	'氏名選択用のWhere条件
    Public m_iNendo			'年度
    Public m_sKyokanCd		'教官コード
    Public m_sSikenKBN		'試験区分
    Public m_sGakuNo		'学年
    Public m_sClassNo		'学科
    Public m_sKamokuCd		'科目コード
    Public m_sSikenNm		'試験名
    Public m_sSikenbi		'試験日
    Public m_sKaisiT		'試験実施開始時間
    Public m_sSyuryoT		'試験実施終了時間
    Public m_sKamokuNo		'科目名
    Public m_sTKyokanCd		'担当科目の教官
	Dim		m_rCnt			'レコードカウント
    Public m_sGakkaCd
	Public m_TUKU_FLG		'通常授業フラグ

	Public m_sGetTable			'科目コンボを作成したテーブル
	
    Public m_iKamoku_Kbn
    Public m_iHissen_Kbn

	Public	m_Rs
	Public	m_TRs
	Public	m_DRs
	Public	m_SRs
	Public	m_iMax			'最大ページ
	Public	m_iNKaishi		'入力開始日
	Public	m_iNSyuryo		'入力終了日

	Public	m_iKikan		'入力期間フラグ

	m_sKaisiT = ""
	m_sSyuryoT = "-"
	m_sSikenbi = ""

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
	w_sMsgTitle="成績登録"
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

'//デバッグ
'Call s_DebugPrint

		'===============================
		'//期間データの取得
		'===============================
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			'// ページを表示
		'	Call No_showPage()
		'	Exit Do
			m_iKikan = "NO"	'成績入力期間外の場合は、表示のみ
		End If
		'If w_iRet <> 0 Then 
		'	m_bErrFlg = True
		'	Exit Do
		'End If

		'===============================
		'//試験時間等の取得
		'===============================
		'w_iRet = f_GetSikenJikan()
		'If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

	   '// ページを表示
	   Call showPage()
	   Exit Do
	Loop

	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
    
    '// 終了処理
    Call gs_CloseDatabase()
End Sub

Sub s_SetParam()
'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSikenKBN	= Cint(request("txtSikenKBN"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sGakkaCd	= request("txtGakkaCd")
	m_TUKU_FLG	= request("txtTUKU_FLG")

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub
    response.write "m_iNendo	=" & m_iNendo	 & "<br>"
    response.write "m_sKyokanCd	=" & m_sKyokanCd & "<br>"
    response.write "m_sSikenKBN	=" & m_sSikenKBN & "<br>"
    response.write "m_sGakuNo	=" & m_sGakuNo	 & "<br>"
    response.write "m_sClassNo	=" & m_sClassNo	 & "<br>"
    response.write "m_sKamokuCd	=" & m_sKamokuCd & "<br>"
    response.write "m_sGakkaCd	=" & m_sGakkaCd  & "<br>"
    response.write "m_TUKU_FLG	=" & m_TUKU_FLG  & "<br>"

End Sub

Function f_Nyuryokudate()
'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	Dim w_sSysDate

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI "
		w_sSQL = w_sSQL & vbCrLf & "  ,SYSDATE "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakuNo)
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
'//成績入力期間テスト用

'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '2002/04/30'"
'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '2000/03/01'"

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_

'response.write w_sSQL & "<<<BR>"

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If m_DRs.EOF Then
			Exit Do
		Else
			m_sSikenNm = gf_SetNull2String(m_DRs("M01_SYOBUNRUIMEI"))		'試験名称
			m_iNKaishi = gf_SetNull2String(m_DRs("T24_SEISEKI_KAISI"))		'成績入力開始日
			m_iNSyuryo = gf_SetNull2String(m_DRs("T24_SEISEKI_SYURYO"))		'成績入力終了日
			w_sSysDate = Left(gf_SetNull2String(m_DRs("SYSDATE")),10)		'システム日付

		End If

		'入力期間内なら正常
		If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			f_Nyuryokudate = 0
		End If

		'f_Nyuryokudate = 0

		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  履修テーブルより科目名称を取得
'*  [引数]  なし
'*  [戻値]  p_KamokuName
'*  [説明]  
'********************************************************************************
Function f_GetKamokuName(p_Gakunen,p_GakkaCd,p_KamokuCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_GakkaCd
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_GetKamokuName = ""
	p_KamokuName = ""

    Do 

	If m_TUKU_FLG = C_TUKU_FLG_TUJO Then '通常授業と特別活動で取り先を変える。
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI AS KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(p_Gakunen) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & p_GakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_KamokuCd & "'"
	Else 
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M41_MEISYO AS KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM M41_TOKUKATU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M41_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M41_TOKUKATU_CD='" & p_KamokuCd & "'"
	End If

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		If w_Rs.EOF = False Then
			p_KamokuName = w_Rs("KAMOKUMEI")
		End If

        Exit Do
    Loop

    f_GetKamokuName = p_KamokuName

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  試験時間等を取得
'*  [引数]  なし
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Function f_GetSikenJikan()

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_GetSikenJikan = ""
	p_KamokuName = ""

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_KAMOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_KAISI_JIKOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_SYURYO_JIKOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_SIKENBI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO = T26_SIKEN_JIKANWARI.T26_CLASS "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_GAKUNEN = M05_CLASS.M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_NENDO = M05_CLASS.M05_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_KBN=" & Cint(m_sSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_CD='0' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_GAKUNEN=" & cint(m_sGakuNo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKKA_CD='" & m_sGakkaCd & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_KAMOKU='" & m_sKamokuCd & "'"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
			f_GetSikenJikan = 99
            Exit Do
        End If

		If w_Rs.EOF = False Then
			m_sKaisiT = w_Rs("T26_KAISI_JIKOKU") & " 〜 "
			m_sSyuryoT = w_Rs("T26_SYURYO_JIKOKU")
			m_sSikenbi = w_Rs("T26_SIKENBI")
		End If

		f_GetSikenJikan = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  レベル別かどうかを調べる。
'*  [引数]  なし
'*  [戻値]  レベル別：true
'*  [説明]  
'********************************************************************************
Function f_LevelChk(p_Gakunen,p_KamokuCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_GakkaCd
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_LevelChk = false
	p_KamokuName = ""
    Do 

		'//引数不足のとき
		If trim(p_Gakunen)="" Or  trim(p_KamokuCd) = "" Then
            Exit Do
		End If

		'//学科CDを取得
'		w_iRet = f_GetGakkaCd(p_Gakunen,p_Class,w_GakkaCd)
		If w_iRet<> 0 Then
            Exit Do
		End If

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(T15_LEVEL_FLG) "
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_NYUNENDO = " & cint(m_iNendo) - cint(p_Gakunen) + 1
'		w_sSQL = w_sSQL & vbCrLf & "  AND T15_GAKKA_CD='" & w_GakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_KAMOKU_CD = '" & p_KamokuCd & "'"


        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		If w_Rs.EOF = False and cint(w_Rs("MAX(T15_LEVEL_FLG)")) = 1 Then
			f_LevelChk = true
		End If

        Exit Do
    Loop
    Call gf_closeObject(w_Rs)
End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_sGakusekiCd
Dim w_sSeiseki
Dim w_sHyoka
Dim w_sKekka
Dim w_sChikai
Dim w_sKekkasu
Dim w_sChikaisu

Dim w_ihalf
Dim i

i = 0

%>
<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

		//スクロール同期制御
		parent.init();
    }

   //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){
        parent.main.f_Touroku();
    }

	//************************************************************
	//	[機能]	キャンセルボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//************************************************************
	function f_Cansel(){

        //初期ページを表示
        parent.document.location.href="default.asp"
	
	}

	//-->
	</SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	<%If m_iKikan <> "NO" Then%>
		<% call gs_title(" 成績登録 "," 登　録 ") %>
	<%Else %>
		<% call gs_title(" 成績登録 "," 表　示 ") %>
	<%End If%>
	<center>
	<table border=1 class=hyo >
		<tr>
			<th class=header align="center" colspan="3">成績入力期間・<%=m_sSikenNm%></th>
		</tr>
		<tr>
			<th class=header width="96"  align="center">成績入力期間</th>
			<td class=detail width="360"  align="center" colspan="2"><%=m_DRs("T24_SEISEKI_KAISI")%> 〜 <%=m_DRs("T24_SEISEKI_SYURYO")%></td>
		</tr>
		<tr>
			<th class=header width="96"  align="center">実施科目</th>
			<!--td class=detail width="200"  align="center"><%=m_sSikenbi%>　<%=m_sKaisiT%><%=m_sSyuryoT%></td-->
			<!--<td class=detail width="104"  align="center"><%=f_GetKamokuName(m_sGakuNo,m_sGakkaCd,m_sKamokuCd)%></td>-->
<%
	If f_LevelChk(m_sGakuNo,m_sKamokuCd) = true then 
		w_str = m_sGakuNo & "年　" & gf_GetClassName(m_iNendo,m_sGakuNo,m_sClassNo) & "　" & f_GetKamokuName(m_sGakuNo,m_sGakkaCd,m_sKamokuCd)
	Else
		w_str = m_sGakuNo & "年　" & gf_GetClassName(m_iNendo,m_sGakuNo,m_sClassNo) & "　" & f_GetKamokuName(m_sGakuNo,m_sGakkaCd,m_sKamokuCd)
	End If

%>
			<td class=detail width="360"  align="center"><%=w_str%></td>

		</tr>
	</table>
			<br>
			<span class=msg>※「日々計」は、授業出欠入力メニューより日々入力された上記試験までの出欠状況です。<br>
			<span class=msg>※「対象外」は、公欠などの累計を入力してください。
	<%
 If m_TUKU_FLG = C_TUKU_FLG_TUJO Then '通常授業と特別活動で表示を変える
    %><br><%
	Select Case m_sSikenKBN
		Case C_SIKEN_ZEN_TYU%>
							※ 評価欄をクリックすると、評価の入力ができます。（○→・の順で表示されます）
		<%Case C_SIKEN_KOU_TYU%>
							※ 評価欄をクリックすると、評価の入力ができます。（○→◎→・の順で表示されます）
		<%Case Else%>
							<br>
	<%End Select%>
 <%Else%>
		<br>
 <%End If %>
			</span><br>
	<table width="550"%>
	<tr>
		<td align=center>
		<%If m_iKikan <> "NO" Then%>
			<input type=button class=button value="　登　録　" onclick="javascript:f_Touroku()">　
		<%End If%>
		<input type=button class=button value="キャンセル" onclick="javascript:f_Cansel()"></td>
	</tr>
	</table>
<!--
	<table >
	<tr>
-->
<% If m_TUKU_FLG = C_TUKU_FLG_TUJO Then '通常授業と特別活動で表示を変える。%>
<!--
	<td valign="top">
-->
		<table class="hyo" border=1 align="center" width="550">
		<tr>
			<th class="header" rowspan="2" width="40"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
			<th class="header" rowspan="2" width="250">氏　名</th>
			<th class="header" rowspan="2" width="35">成績</th>
			<th class="header" rowspan="2" width="35">評価</th>
			<th class="header" colspan="2" >遅刻</th>
			<th class="header" colspan="3" >欠課</th>
			
		</tr>
		<tr>
			<th class="header2" width="35"><span style="font-size:10px;">入力</span></th>
			<th class="header2" width="35"><span style="font-size:10px;">日々計</span></th>
			<th class="header2" width="40"><span style="font-size:10px;">対象</span></th>
			<th class="header2" width="40"><span style="font-size:10px;">対象外</span></th>
			<th class="header2" width="45"><span style="font-size:10px;">日々計</span></th>
		</tr>
		</table>

<!--
	</td>
	<td valign="top">
		<table class="hyo" border=1 align="center" width="383">
		<tr>
			<th class="header" rowspan="2" width="45">学籍<br>番号</th>
			<th class="header" rowspan="2" width="130">氏　名</th>
			<th class="header" rowspan="2" width="30">成<br>績</th>
			<th class="header" rowspan="2" width="30">評<br>価</th>
			<th class="header" colspan="2" width="74">遅刻</th>
			<th class="header" colspan="2" width="74">欠課</th>
		</tr>
		<tr>
			<th class="header2" width="37"><span style="font-size:10px;">入力</span></th>
			<th class="header2" width="37"><span style="font-size:10px;"日々計</span></th>
			<th class="header2" width="37"><span style="font-size:10px;">入力</span></th>
			<th class="header2" width="37"><span style="font-size:10px;">日々計</span></th>
		</tr>
		</table>
	</td>
-->
<% else %>
<!--
	<td valign="top">
-->
		<table class="hyo" border=1 align="center" width="550">
		<tr>
			<th class="header" rowspan="2" width="40"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
			<th class="header" rowspan="2" width="250">氏　名</th>
			<th class="header" rowspan="2" width="35">成績</th>
			<th class="header" rowspan="2" width="35">評価</th>
			<th class="header" rowspan="2" width="70">遅刻</th>
			<th class="header" colspan="2" width="120">欠課</th>
		</tr>
		<tr>
			<th class="header2" width="60"><span style="font-size:10px;">対象</span></th>
			<th class="header2" width="60"><span style="font-size:10px;">対象外</span></th>
		</tr>
		</table>
<!--
	</td>

	<td valign="top">
		<table class="hyo" border=1 align="center" width="383">
		<tr>
			<th class="header" width="45">学籍<br>番号</th>
			<th class="header" width="130">氏　名</th>
			<th class="header" width="30">成<br>績</th>
			<th class="header" width="30">評<br>価</th>
			<th class="header" width="54">遅刻</th>
			<th class="header" width="54">欠課</th>
		</tr>
		</table>
	</td>
-->
<% end if %>
<!--
	</tr>
	</table>
-->
	</FORM>
	</center>
	</body>
	</html>
<%
End sub

Sub No_showPage()
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
	<SCRIPT language="javascript">
	<!--

	    //************************************************************
	    //  [機能]  ページロード時処理
	    //  [引数]
	    //  [戻値]
	    //  [説明]
	    //************************************************************
	    function window_onload() {

	        //submit
			parent.location.href = "white.asp?txtMsg=成績入力期間外です。"
	        return;
	    }

	//-->
	</SCRIPT>
	</head>

    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">

	<input type="hidden" name="txtMsg" value="成績入力期間外です。">

	</form>
	</body>
	</html>

<%
End Sub

Sub showPage_No()
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
	<SCRIPT language="javascript">
	<!--

	    //************************************************************
	    //  [機能]  ページロード時処理
	    //  [引数]
	    //  [戻値]
	    //  [説明]
	    //************************************************************
	    function window_onload() {

	        //submit
			parent.location.href = "white.asp?txtMsg=個人履修データが存在しません。"
	        return;
	    }

	//-->
	</SCRIPT>
	</head>

    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	</head>

	<body>
	<br><br><br>
	<center>
		<span class="msg">個人履修データが存在しません。</span>
	</center>

	<input type="hidden" name="txtMsg" value="データが存在しません。">

	</form>
	</body>
	</html>

<%
End Sub
%>