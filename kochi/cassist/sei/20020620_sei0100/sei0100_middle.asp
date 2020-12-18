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
' 変      更: 2002/05/02 進   浩人 特別活動の遅刻もエクセル貼り付け対応に
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
	
    Public m_sGakuNo_s		'学年
    Public m_sGakkaCd_s		'学科
    Public m_sKamokuCd_s	'科目コード

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
	Public	m_iKekkaKaishi		'欠席入力開始日
	Public	m_iKekkaSyuryo		'欠席入力終了日


	Public	m_iKikan		'入力期間フラグ
	Public	m_bKekkaNyuryokuFlg		'欠課入力可能ﾌﾗｸﾞ(True:入力可 / False:入力不可)
	
	Public m_UpdateDate
	Public m_sFirstGakusekiNo
	
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

	    '// ﾊﾟﾗﾒｰﾀSET
	    Call s_SetParam()

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'===============================
		'//期間データの取得
		'===============================
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			m_iKikan = "NO"	'成績入力期間外の場合は、表示のみ
		End If
		
		if not f_GetUpdateDate(m_iNendo,m_sKamokuCd,m_sSikenKBN,m_TUKU_FLG,m_sFirstGakusekiNo,m_UpdateDate) then
			m_bErrFlg = True
			Exit Do
		end if
		
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

	m_sGakuNo_s	= Cint(request("txtGakuNo"))
	m_sGakkaCd_s	= request("txtGakkaCd")
	m_sKamokuCd_s	= request("txtKamokuCd")
	
	m_UpdateDate = ""
	m_sFirstGakusekiNo	= request("hidFirstGakusekiNo")
	
End Sub

'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_Nyuryokudate()

	Dim w_sSysDate

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1
	m_bKekkaNyuryokuFlg = False

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_SYURYO "
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

'response.write w_sSQL & "<BR>"

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If m_DRs.EOF Then
			m_iNKaishi="          "
			m_iNSyuryo="          "
			Exit Do
		Else
			m_sSikenNm = gf_SetNull2String(m_DRs("M01_SYOBUNRUIMEI"))		'試験名称
			m_iNKaishi = gf_SetNull2String(m_DRs("T24_SEISEKI_KAISI"))		'成績入力開始日
			m_iNSyuryo = gf_SetNull2String(m_DRs("T24_SEISEKI_SYURYO"))		'成績入力終了日
			m_iKekkaKaishi = gf_SetNull2String(m_DRs("T24_KEKKA_KAISI"))	'欠課入力開始
			m_iKekkaSyuryo = gf_SetNull2String(m_DRs("T24_KEKKA_SYURYO"))	'欠課入力終了
			w_sSysDate = Left(gf_SetNull2String(m_DRs("SYSDATE")),10)		'システム日付
		End If

		'入力期間内なら正常
		If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			f_Nyuryokudate = 0
		End If

		'欠課入力可能ﾌﾗｸﾞ
		If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			m_bKekkaNyuryokuFlg = True
		End If

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

'*******************************************************************************
' 機　　能：選んだ試験区分,科目の更新日を取得する
' 
' 返　　値：
' 　　　　　(True)成功, (False)失敗
' 引　　数：p_sSikenKbn			- 試験区分
' 　　　　　p_Nendo				- 年度
' 　　　　　p_KamokuCd			- 科目コード
'			p_FirstGakusekiNo	- 学籍NO
'			p_TUKU_FLG			- 科目区分(0:通常、1:特活)
'			(戻り値)p_UpdateDate - 科目実績登録の更新日
' 機能詳細：
' 備　　考：
' old_ver → gf_GetT16UpdDate(m_iNendo,m_sGakuNo_s,m_sGakkaCd_s,m_sKamokuCd_s,"")
'*******************************************************************************
function f_GetUpdateDate(p_Nendo,p_KamokuCd,p_ShikenKbn,p_TUKU_FLG,p_FirstGakusekiNo,p_UpdateDate)
	
	Dim w_Sql,w_Rs
	Dim w_ShikenType
	Dim w_Table
	Dim w_TableName
	Dim w_KamokuName
	
	On Error Resume Next
	Err.Clear
	
	f_GetUpdateDate = false
	
	if cint(p_TUKU_FLG) = cint(C_TUKU_FLG_TUJO) then
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	else
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_KamokuName = "T34_TOKUKATU_CD"
	end if
	
	select case cint(p_ShikenKbn)
		
		case C_SIKEN_ZEN_TYU '前期中間試験
			w_ShikenType = w_Table & "_KOUSINBI_TYUKAN_Z"
			
		case C_SIKEN_ZEN_KIM '前期期末試験
			w_ShikenType = w_Table & "_KOUSINBI_KIMATU_Z"
			
		case C_SIKEN_KOU_TYU '後期中間試験
			w_ShikenType = w_Table & "_KOUSINBI_TYUKAN_K"
			
		case C_SIKEN_KOU_KIM '後期期末試験
			w_ShikenType = w_Table & "_KOUSINBI_KIMATU_K"
			
		case else
			exit function
			
	end select
	
	w_Sql = ""
	w_Sql = w_Sql & " select "
	w_Sql = w_Sql & " 		" & w_ShikenType
	w_Sql = w_Sql & " from "
	w_Sql = w_Sql & " 		" & w_TableName
	w_Sql = w_Sql & " where "
	w_Sql = w_Sql & " 		" & w_Table & "_GAKUSEKI_NO= '"   & p_FirstGakusekiNo   & "' "
	w_Sql = w_Sql & " and	" & w_Table & "_NENDO = " & p_Nendo
	w_Sql = w_Sql & " and	" & w_KamokuName & "= '"   & p_KamokuCd   & "' "
	
	If gf_GetRecordset(w_Rs,w_Sql) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit function
	End If
	
	p_UpdateDate = gf_YYYY_MM_DD(w_Rs(0),"/")
	
	f_GetUpdateDate = true
	
end function


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
        parent.document.location.href="default.asp";
	
	}

	//************************************************************
	//	[機能]	ペーストボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//************************************************************
	function f_Paste(pType){
		
		parent.main.document.frm.PasteType.value=pType;
		
		//submitで画面を開くとウィンドウのステータスが設定できないため､
		//一旦空ページを開いてから、新ウィンドウに対してsubmitする。
		nWin=open("","Paste","location=no,menubar=no,resizable=yes,scrollbars=no,scrolling=no,status=no,toolbar=no,width=300,height=600,top=0,left=0");
		parent.main.document.frm.target="Paste";
		parent.main.document.frm.action="sei0100_paste.asp";
		parent.main.document.frm.submit();
	
	}
	//-->
	</SCRIPT>
	</head>
    <body onload="return window_onload()">
	<table border="0" cellpadding="0" cellspacing="0" height="245" width="100%">
		<tr>
			<td>
				<%
				If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then
					call gs_title(" 成績登録 "," 登　録 ")
				Else
					call gs_title(" 成績登録 "," 表　示 ")
				End If
				%>
			</td>
		</tr>
		<tr>
			<td align="center" nowrap><form name="frm" method="post">
				<table border=1 class=hyo width=670>
					<tr>
						<th class="header3" colspan="6" nowrap align="center">
						成績入力期間　<%=m_sSikenNm%>　　　更新日：<%=m_UpdateDate%>
						</th>
					</tr>
					<tr>
						<th class=header3 width="96"  align="center">成績入力期間</th><td class=detail width="239"  align="center" colspan="2"><%=m_iNKaishi%> ～ <%=m_iNSyuryo%></td>
						<th class=header3 width="96"  align="center">欠課入力期間</th><td class=detail width="239"  align="center" colspan="2"><%=m_iKekkaKaishi%> ～ <%=m_iKekkaSyuryo%></td>
					</tr>
					<tr>
						<th class=header3 width="96"  align="center">実施科目</th>
						<%
							If f_LevelChk(m_sGakuNo,m_sKamokuCd) = true then 
								w_str = m_sGakuNo & "年　" & gf_GetClassName(m_iNendo,m_sGakuNo,m_sClassNo) & "　" & f_GetKamokuName(m_sGakuNo,m_sGakkaCd,m_sKamokuCd)
							Else
								w_str = m_sGakuNo & "年　" & gf_GetClassName(m_iNendo,m_sGakuNo,m_sClassNo) & "　" & f_GetKamokuName(m_sGakuNo,m_sGakkaCd,m_sKamokuCd)
							End If
						%>
						<td class=detail colspan="5" align="center"><%=w_str%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">

				<span class=msg2>
				※「日々計」は、授業出欠入力メニューより日々入力された上記試験までの出欠状況です。<br>
				※「対象外」は、公欠などの累計を入力してください。<br>
				※ヘッダの文字色が「<FONT COLOR="#99CCFF">成績</FONT>」のようになっている部分をクリックすると、Excel貼り付け用の画面が開きます。<br>
				<%
				'通常授業と特別活動で表示を変える
				If m_TUKU_FLG = C_TUKU_FLG_TUJO Then
					Select Case m_sSikenKBN
						Case C_SIKEN_ZEN_TYU
							%>※ 評価欄をクリックすると、評価の入力ができます。（○→・の順で表示されます）<br><%
						Case C_SIKEN_KOU_TYU
							%>※ 評価欄をクリックすると、評価の入力ができます。（○→◎→・の順で表示されます）<br><%
						Case Else
							response.write "<BR>"
					End Select
				End If
				%>
				</spanc>

				<% If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then %>
					<input type=button class=button value="　登　録　" onclick="javascript:f_Touroku()">　
				<%End If%>
				<input type=button class=button value="キャンセル" onclick="javascript:f_Cansel()">

			</td>
		</tr>
		<tr>
			<td align="center" valign="bottom" nowrap>
				
				<!--通常授業と特別活動で表示を変える。-->
				<% If m_TUKU_FLG = C_TUKU_FLG_TUJO Then %>
					
					<table class="hyo" border="1" align="center" width="710">
						<tr><th class="header3" colspan="13" nowrap align="center">
								総授業数&nbsp;<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%><input type="text" class="NUM" maxlength="3" style="width:30px" name="txtSouJyugyou" value="<%= Request("hidSouJyugyou") %>"><% Else %><%= Request("hidSouJyugyou") %><% End if%>　
								純授業数&nbsp;<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%><input type="text" class="NUM" maxlength="3" style="width:30px" name="txtJunJyugyou" value="<%= Request("hidJunJyugyou") %>"><% Else %><%= Request("hidJunJyugyou") %><% End if%>　
							</th></tr>                                                                                                                                                 
						<tr>
							<th class="header3" rowspan="2" width="65" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
							<th class="header3" rowspan="2" width="150" nowrap>氏　名</th>
							<th class="header3" colspan="4" width="120" nowrap>成績履歴</th>
							<th class="header3" rowspan="2" width="50" nowrap onClick="f_Paste('Seiseki')"><FONT COLOR="#99CCFF">成績</FONT></th>
							<th class="header3" rowspan="2" width="50" nowrap>評価</th>
							<th class="header3" colspan="2" width="110" nowrap>遅刻</th>
							<th class="header3" colspan="3" width="165" nowrap">欠課</th>
						</tr>
						
						<tr>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">前中</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">前末</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">後中</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">学末</span></th>
							<th class="header2" width="55" nowrap onClick="f_Paste('Chikai')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">入力</FONT></span></th>
							<th class="header2" width="55" nowrap><span style="font-size:10px;">日々計</span></th>
							<th class="header2" width="55" nowrap onClick="f_Paste('Kekka')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">対象</FONT></span></th>
							<th class="header2" width="55" nowrap onClick="f_Paste('KekkaGai')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">対象外</FONT></span></th>
							<th class="header2" width="55" nowrap><span style="font-size:10px;">日々計</span></th>
						</tr>
					</table>
				<% else %>
					<table class="hyo" border=1 align="center" width="710" nowrap>
						<tr>
							<th class="header3" colspan="13" nowrap align="center">
								総授業数&nbsp;<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%><input type="text" class="NUM" maxlength="5" style="width:30px" name="txtSouJyugyou" value="<%= Request("hidSouJyugyou") %>"><% Else %><%= Request("hidSouJyugyou") %><% End if%>　
								純授業数&nbsp;<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%><input type="text" class="NUM" maxlength="5" style="width:30px" name="txtJunJyugyou" value="<%= Request("hidJunJyugyou") %>"><% Else %><%= Request("hidJunJyugyou") %><% End if%>　
							</th>
						</tr>
						<tr>
							<th class="header3" rowspan="2" width="65" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
							<th class="header3" rowspan="2" width="150" nowrap>氏　名</th>
							<th class="header3" colspan="4" width="120" nowrap>成績履歴</th>
							<th class="header3" rowspan="2" width="50" nowrap onClick="f_Paste('Seiseki')"><FONT COLOR="#99CCFF">成績</FONT></th>
							<th class="header3" rowspan="2" width="50" nowrap>評価</th>
							<th class="header3" rowspan="2" width="100" nowrap onClick="f_Paste('Chikai')"><FONT COLOR="#99CCFF">遅刻</FONT></th>
							<th class="header3" colspan="2" width="165" nowrap>欠課</th>
						</tr>
						<tr>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">前中</FONT></span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">前末</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">後中</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">学末</span></th>
							<th class="header2" width="80" nowrap onClick="f_Paste('Kekka')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">対象</FONT></span></th>
							<th class="header2" width="85" nowrap onClick="f_Paste('KekkaGai')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">対象外</FONT></span></th>
						</tr>
					</table>
				<% end if %>
			</td>
		</tr>
	</table>

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