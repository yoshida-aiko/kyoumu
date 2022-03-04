<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 仮進級者成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0900/sei0900_middle.asp
' 機      能: 下ページ 仮進級者成績登録の検索を行う
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
' 作      成: 2022/2/1 吉田　再試験成績登録画面を流用し作成
' 変      更: 
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
    Public m_sKamokuNM		'科目名			INS 2017/12/26 Nishimura
    Public m_sSikenNm		'試験名
    Public m_sSikenbi		'試験日
    Public m_sKaisiT		'試験実施開始時間
    Public m_sSyuryoT		'試験実施終了時間
    Public m_sKamokuNo		'科目名
    Public m_sTKyokanCd		'担当科目の教官
	Dim		m_rCnt			'レコードカウント
    Public m_sGakkaCd
	Public m_TUKU_FLG		'通常授業フラグ
	Public m_iRisyuKakoNendo'過年度
	
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
	
	Dim m_iCount
	Dim m_sMiHyoka
	Dim m_Checked
	Dim m_Disabled
	Dim m_SchoolFlg
	
	m_Checked  = ""
	m_Disabled = ""
	
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
'response.write "middle START" & "<BR>"
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="仮進級者成績登録"
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

		'===============================
		'//最新更新日を取得
		'===============================
		'//クラスの学年末成績を最後に更新した日を取得
		If f_GetUpdateDate(m_UpdateDate) <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		
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
	' m_sGakuNo	= Cint(request("txtGakuNo"))
	' m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sKamokuNM	= request("txtKamokuNM")
	m_iRisyuKakoNendo = request("txtRisyuKakoNendo")
	m_sGakkaCd	= request("txtGakkaCd")
	m_TUKU_FLG	= request("txtTUKU_FLG")

	m_sGakuNo_s	= Cint(request("txtGakuNo"))
	m_sGakkaCd_s	= request("txtGakkaCd")
	m_sKamokuCd_s	= request("txtKamokuCd")
	
	m_UpdateDate = ""
	m_sFirstGakusekiNo	= request("hidFirstGakusekiNo")
	
	m_iCount = cint(request("i_Max"))
	m_sMiHyoka = request("hidMihyoka")
	m_SchoolFlg = cbool(request("hidSchoolFlg"))
	
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
	' m_bKekkaNyuryokuFlg = False

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		' w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_KAISI "
		' w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_SYURYO "
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
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & C_SIKEN_KARISINKYU
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND rownum <= 1 "

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
			' m_iKekkaKaishi = gf_SetNull2String(m_DRs("T24_KEKKA_KAISI"))	'欠課入力開始
			' m_iKekkaSyuryo = gf_SetNull2String(m_DRs("T24_KEKKA_SYURYO"))	'欠課入力終了
			w_sSysDate = Left(gf_SetNull2String(m_DRs("SYSDATE")),10)		'システム日付
		End If

		'入力期間内なら正常
		If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			f_Nyuryokudate = 0
		End If

		' '欠課入力可能ﾌﾗｸﾞ
		' If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
		' 	m_bKekkaNyuryokuFlg = True
		' End If

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

	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI AS KAMOKUMEI"
	w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(p_Gakunen) + 1
	w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & p_GakkaCd & "'"
	w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_KamokuCd & "'"
	

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

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(T15_LEVEL_FLG) "
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_NYUNENDO = " & cint(m_iNendo) - cint(p_Gakunen) + 1
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

'********************************************************************************
'*  [機能]  最終更新日の取得
'*  [引数]  なし
'*  [戻値]  最終更新日
'*  [説明]  
'********************************************************************************
Function f_GetUpdateDate(p_UpdateDate)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetUpdateDate = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(T17_KOUSINBI_KIMATU_K)"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T17_RISYUKAKO_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T17_NENDO = " & Cint(m_iRisyuKakoNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T17_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T17_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T17_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T17_NENDO = C.T13_NENDO "

		'//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
		w_sSQL = w_sSQL & " AND	A.T17_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO

		'//必修か選択科目のうち選択している学生のみを取得する		'INS 2019/03/06 藤林
		w_sSQL = w_sSQL & " AND	( T17_HISSEN_KBN = " & C_HISSEN_HIS
		w_sSQL = w_sSQL & "       OR (T17_HISSEN_KBN = " & C_HISSEN_SEN & " AND T17_SELECT_FLG = 1) "
		w_sSQL = w_sSQL & " 	) "
		w_sSQL = w_sSQL & " AND T17_HYOKA_FUKA_KBN NOT IN(" & C_HYOKA_FUKA_KEKKA &  "," & C_HYOKA_FUKA_BOTH & ") "
	
        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetUpdateDate = 99
            Exit Do
        End If
		
		' response.write "w_sSQL:" & w_sSQL & "<BR>"
		' response.end
		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_UpdateDate = w_Rs("MAX(T17_KOUSINBI_KIMATU_K)")
		End If
		' response.write "p_UpdateDate" & p_UpdateDate & "<BR>"

        f_GetUpdateDate = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function


'********************************************************************************
'*  [機能]  未評価の設定
'********************************************************************************
Sub setHyokaType()
	
	'科目が未評価
	if cint(gf_SetNull2Zero(m_sMiHyoka)) = cint(C_MIHYOKA) then
		m_Checked = "checked"
	end if
	
	'入力期間外
	if m_iKikan = "NO" then
		m_Disabled = "disabled"
	end if
	
End Sub

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
Dim w_sInputClass

Dim w_ihalf
Dim i

i = 0

'//NN対応
If session("browser") = "IE" Then
	w_sInputClass = "class='num'"
Else
	w_sInputClass = ""
End If

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
		parent.main.document.frm.action="sei0900_paste.asp";
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
					call gs_title(" 仮進級者成績登録 "," 登　録 ")
				Else
					call gs_title(" 仮進級者成績登録 "," 表　示 ")
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
						<th class=header3 width="96"  align="center">仮進級成績入力期間</th><td class=detail width="239"  align="center" colspan="2"><%=m_iNKaishi%> ～ <%=m_iNSyuryo%></td>
					</tr>
					<tr>

						<th class=header3 width="96"  align="center">実施科目</th>
						<td class=detail colspan="5" align="center"><%=m_sKamokuNM%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<span class=msg2>
				<%
				'通常授業
				Select Case m_sSikenKBN
					Case C_SIKEN_ZEN_TYU
						%>※ 評価欄をクリックすると、評価の入力ができます。（○→・の順で表示されます）<br><%
					Case C_SIKEN_KOU_TYU
						%>※ 評価欄をクリックすると、評価の入力ができます。（○→◎→・の順で表示されます）<br><%
					Case Else
						response.write "<BR>"
				End Select
				
				%>
				</span>
				
				<%If m_iKikan <> "NO" Then%>
					<input type=button class=button value="　登　録　" onclick="javascript:f_Touroku()">　
				<%End If%>
				<input type=button class=button value="キャンセル" onclick="javascript:f_Cansel()">
				
			</td>
		</tr>
		<tr>
			<td align="center" valign="bottom" nowrap>
				<table class="hyo" border="1" align="center" width="<%= gf_IIF(m_SchoolFlg,760,710) %>">
					<tr>
						<th class="header3" colspan="14" nowrap align="center">

						</th>
					</tr>                                                                                                                                                 
					
					<tr>
						<th class="header3" rowspan="2" width="65"  nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
						<th class="header3" rowspan="2" width="150" nowrap>氏　名</th>
						<th class="header3" rowspan="2" width="50"  nowrap >成績</th>
						<th class="header3" rowspan="2" width="50"  nowrap>評価</th>
						<% if m_SchoolFlg then %>
							<th class="header3" rowspan="2" width="50"  nowrap>評価<br>不能</th>
						<% end if %>
						
					</tr>
					
					<tr>
					</tr>
				</table>

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