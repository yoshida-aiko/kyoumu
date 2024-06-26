<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0100_bottom.asp
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
' 変      更: 2002/03/20 松尾		NULLを""で表示する
' 変      更: 2017/12/26 西村		科目名引き継ぎを追加
' 変      更: 2018/03/22 西村		開設時期項目追加(非表示)  '2019/12/06 kiyomto ソースからこの対応箇所が漏れており追加
' 修　    正: 2019/03/06 藤林		成績登録時に、必修か選択科目のうち選択している学生のみを取得する様に修正
' 変      更: 2020/03/09 清本		成績登録時に、開設時期を個人履修データから判断する
' 変      更: 2023/12/14 吉田		WEBアクセスログカスタマイズ
' 変      更: 2024/06/19 清本		清掃とSHRは遅刻回数の入力不可としてハイフン表示にする
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '氏名選択用のWhere条件
    Public m_iNendo			'年度
    Public m_sKyokanCd		'教官コード
    Public m_sSikenKBN		'試験区分
    Public m_sGakuNo		'学年
    Public m_sClassNo		'学科
    Public m_sKamokuCd		'科目コード
    Public m_sKamokuNM		'科目名 Ins 2017/12/26 Nishimura
    Public m_sSikenNm		'試験名
    Public m_rCnt			'レコードカウント
    Public m_sGakkaCd
    Public m_iSyubetu		'出欠値集計方法
    Public m_TUKU_FLG
	Public m_iJigenTani		'//１時限の単位数
	Public m_sZokuseiCd		'属性コード Ins 2022/10/7 Yoshida
    
    Public m_iKamoku_Kbn
    Public m_iHissen_Kbn
	Public m_ilevelFlg
	Public m_Rs
	Public m_DRs
	Public m_SRs
	Public m_Rs_Hyoka			'評価情報
	
	Dim m_iSouJyugyou		'総授業時間
	DIm m_iJunJyugyou		'純授業時間
	
	Public	m_iKikan		'入力期間フラグ
	Public	m_bKekkaNyuryokuFlg		'欠課入力可能ﾌﾗｸﾞ(True:入力可 / False:入力不可)
	
	Public	m_iShikenInsertType
	Public m_FirstGakusekiNo
	Public m_FirstSeisekiInp			'成績の入力方法
	
	m_iShikenInsertType = 0
	
	Public m_sSyubetu
	Dim m_SchoolFlg

	Public  m_sTaisyo				'対象		'add 2023/12/14 吉田
	Public  m_sSosa					'操作		'add 2023/12/14 吉田
	Public  m_sUserId				'ログインID	'add 2023/12/14 吉田
	
    Private Const C_TOKU_SEISO = "0002" '特別活動 清掃 科目コード   '2024.06.19 Add Kiyomoto
    Private Const C_TOKU_SHR = "0003"   '特別活動 SHR 科目コード    '2024.06.19 Add Kiyomoto

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
		If gf_OpenDatabase() <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If
		'//不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'// ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()
		
		'//評価不能チェックの処理は、熊本だけのため学校番号のチェック
		'//Trueなら熊本電波
		if not gf_ChkDisp(C_DATAKBN_DISP,m_SchoolFlg) then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//期間データの取得
		If f_Nyuryokudate() = 1 Then
			m_iKikan = "NO"
		else
			m_iKikan = ""
		End If
		
		'=================
		'//出欠欠課の取り方を取得
		'=================
		'//科目区分(0:試験毎,1:累積)
		If gf_GetKanriInfo(m_iNendo,m_iSyubetu) <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		'=================
		'//特別活動を取得
		'=================
		'//特別活動(0:通常授業,1:特別活動)
        If f_getTUKU(m_iNendo,m_sKamokuCd,m_sGakuNo,m_sClassNo,m_TUKU_FLG) <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If

		'//2022/10/7 Ins-->
		'=================
		'評価形式を取得
		'=================
		If f_GetHyokaInfo() <> 0 Then 
				m_bErrFlg = True
				Exit Do
		End If
		'//2022/10/7 Ins<--

		'**********************************************************
		'通常授業と特別活動で、とって来る場所が変わる。
		'**********************************************************
		If m_TUKU_FLG = C_TUKU_FLG_TUJO then  '通常授業の場合
			'=================
			'//科目情報を取得
			'=================
			'//科目区分(0:一般科目,1:専門科目)、及び、必修選択区分(1:必修,2:選択)を調べる
			'//レベル別区分(0:一般科目,1:レベル別科目)を調べる
			If f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) <> 0 Then 
				m_bErrFlg = True
				Exit Do
			End If
			
			'===============================
			'//成績、学生データ取得
			'===============================
			'//科目区分がC_KAMOKU_SENMON(0:一般科目)の場合はクラス別に生徒を表示
			'//科目区分がC_KAMOKU_SENMON(1:専門科目)の場合は学科別に生徒を表示
			If f_getdate(m_iKamoku_Kbn) <> 0 Then m_bErrFlg = True : Exit Do
			If m_rs.EOF Then
				Call ShowPage_No()
				Exit Do
			End If
			
			'===============================
			'//欠課数の取得
			'===============================
			If f_GetSyukketu() <> 0 Then m_bErrFlg = True : Exit Do
			
			'===============================
			'//TODO　留年生取得2022/03/01 
			'===============================

		Else
			'===============================
			'//成績、学生データ取得
			'===============================
			If f_getTUKUclass(m_iNendo,m_sKamokuCd,m_sGakuNo,m_sClassNo) <> 0 Then m_bErrFlg = True : Exit Do
			If m_rs.EOF Then
				Call ShowPage_No()
				Exit Do
			End If
			
	    End If
		
		'add start 2023/12/14 吉田
		'操作LOG出力
		If gf_InsertOpeLog(m_iNendo,"SEI0100","成績登録",m_sTaisyo,m_sSosa,m_sUserId) <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		'add end 2023/12/14 吉田

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

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_SetParam()

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSikenKBN	= Cint(request("txtSikenKBN"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sKamokuNM	= request("txtKamokuNM")	'Ins 2017/12/26 Nishimura
	m_sGakkaCd	= request("txtGakkaCd")
	m_iJigenTani = Session("JIKAN_TANI") '１時限の単位数
	m_sSyubetu	= trim(Request("SYUBETU"))
	m_sZokuseiCd	= request("txtZokuseiCd")	'Ins 2022/10/7 Yoshida
	m_FirstSeisekiInp	= request("hidSeisekiInp")	'Ins 2022/10/14 Yoshida

	'add start 2023/12/14 吉田
	m_sTaisyo = request("LOG_TAISYO")
	m_sSosa = request("LOG_SOSA")
	m_sUserId = Session("LOGIN_ID")
	'add end 　2023/12/14 吉田
	
End Sub

'********************************************************************************
'*  [機能]  欠課数、遅刻数を取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSyukketu()
	Dim w_iRet
	'Dim w_bZenkiOnly
	'Dim w_sSikenKBN
	Dim w_sTKyokanCd
	
    On Error Resume Next
    Err.Clear
	
    f_GetSyukketu = 1
	
	Do
		'==========================================
		'//科目担当教官の教官CDの取得
		'==========================================
        w_iRet = f_GetTantoKyokan2(w_sTKyokanCd)
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
		
		'==========================================
		'//試験科目が前期のみか通年かを調べる
		'==========================================
		'//前期のみの場合はT21より前記期末試験までの欠課数を取得する
		'w_iRet = f_SikenInfo(w_bZenkiOnly)
		'If w_iRet<> 0 Then
		'	Exit Do
		'End If 
		
		'If w_bZenkiOnly = True Then
		'	w_sSikenKBN = C_SIKEN_ZEN_KIM
		'Else
		'	w_sSikenKBN = m_sSikenKBN
		'End If
		
		'//最初の生徒の学籍番号を取得
		if not m_Rs.EOF then
			m_FirstGakusekiNo = m_Rs("GAKUSEKI_NO")
			m_Rs.movefirst
		end if

		
		'==========================================
		'//科目に対する結果,遅刻の値取得
		'==========================================
		'if not gf_GetSyukketuData(m_SRs,w_sSikenKBN,m_sGakuNo,w_sTKyokanCd,m_sClassNo,m_sKamokuCd,w_skaisibi,w_sSyuryobi,"") then
		if not gf_GetSyukketuData2(m_SRs,m_sSikenKBN,m_sGakuNo,w_sTKyokanCd,m_sClassNo,m_sKamokuCd,w_skaisibi,w_sSyuryobi,"",m_iNendo,m_iShikenInsertType,m_FirstGakusekiNo,m_sSyubetu) then
			Exit Do
		end if
		
		'//正常終了
	    f_GetSyukketu = 0
		Exit Do
	Loop

End Function 

'********************************************************************************
'*  [機能]  試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_SikenInfo(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_SikenInfo = 1
	p_bZenkiOnly = false

'    Do 		'2020.03.09 Del 不要なLoopを削除

		'//試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSQL = w_sSQL & vbCrLf & " T15_RISYU.T15_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & Cint(m_iNendo)-cint(m_sGakuNo)+1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & Trim(m_sKamokuCd) & "'" 
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAISETU" & m_sGakuNo & "=" & C_KAI_ZENKI

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_SikenInfo = 99
'            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_bZenkiOnly = True
		End If

        f_SikenInfo = 0
'        Exit Do
'    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  コンボで選択された科目の科目区分及び、必修選択区分を調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKamokuInfo(p_iKamoku_Kbn,p_iHissen_Kbn,p_ilevelFlg)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKamokuInfo = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKU_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  ,T15_RISYU.T15_HISSEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  ,T15_RISYU.T15_LEVEL_FLG"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(m_sGakuNo) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & m_sKamokuCd & "' "
		
        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKamokuInfo = 99
            Exit Do
        End If
		
		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_iKamoku_Kbn = w_Rs("T15_KAMOKU_KBN")
			p_iHissen_Kbn = w_Rs("T15_HISSEN_KBN")
			p_ilevelFlg = w_Rs("T15_LEVEL_FLG")
		End If

        f_GetKamokuInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'2022/10/7 Yoshida Ins -->
'********************************************************************************
'*  [機能]  コンボで選択された科目の評価形式の評価を取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetHyokaInfo()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetHyokaInfo = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "   M08_HYOKA_SYOBUNRUI_CD"
		w_sSQL = w_sSQL & vbCrLf & "  ,M08_HYOKA_SYOBUNRUI_MEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M100_SEISEKI_INP"
		w_sSQL = w_sSQL & vbCrLf & " FROM M08_HYOKAKEISIKI,M100_KAMOKU_ZOKUSEI "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M08_HYOKAKEISIKI.M08_NENDO =" & cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M08_HYOKAKEISIKI.M08_HYOUKA_NO ='" & m_sZokuseiCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND M08_HYOKAKEISIKI.M08_NENDO = M100_KAMOKU_ZOKUSEI.M100_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M08_HYOKAKEISIKI.M08_HYOUKA_NO = M100_KAMOKU_ZOKUSEI.M100_ZOKUSEI_CD"
		w_sSQL = w_sSQL & vbCrLf & "  ORDER BY M08_HYOKA_SYOBUNRUI_CD"

 'response.write "w_sSQL" & w_sSQL

        iRet = gf_GetRecordset(m_Rs_Hyoka, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetHyokaInfo = 99
            Exit Do
        End If

		'//最初の評価の「M100_SEISEKI_INP」を取得する
		if not m_Rs_Hyoka.EOF then
			m_FirstSeisekiInp = Cint(gf_SetNull2Zero(m_Rs_Hyoka("M100_SEISEKI_INP")))
			m_Rs_Hyoka.movefirst
		end if
'  response.write "m_FirstSeisekiInp" & m_FirstSeisekiInp

        f_GetHyokaInfo = 0
        Exit Do
    Loop

End Function
'2022/10/7 Yoshida Ins -->

'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_getdate(p_iKamoku_Kbn)
Dim w_iNyuNendo
	
	On Error Resume Next
	Err.Clear
	f_getdate = 1
	
	Do
		w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakuNo) + 1
		
		'//検索結果の値より一覧を表示
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " A.T16_SEI_TYUKAN_Z AS SEI1,A.T16_SEI_KIMATU_Z AS SEI2,A.T16_SEI_TYUKAN_K AS SEI3,A.T16_SEI_KIMATU_K AS SEI4, "
		
		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_Z AS SEI,A.T16_KEKA_TYUKAN_Z AS KEKA,A.T16_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_Z AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_TYUKAN_Z as SOUJI, A.T16_JUNJIKAN_TYUKAN_Z as JYUNJI, "
				
				w_sSQL = w_sSQL & "		A.T16_DATAKBN_TYUKAN_Z as DataKbn ,"
				w_sSQL = w_sSQL & "		A.T16_HYOKA_TYUKAN_Z as HYOKA ," '//INS 2022/10/07 吉田
				'2023.08.24 Del Kiyomoto  前年度成績は全試験で学年末を参照 -->
				'w_sSQL = w_sSQL & "		D.T19_SEI_TYUKAN_Z as ZenNendoSeiseki ,"	'//INS 2022/03/02 吉田　再履修対応
				''//INS 2022/03/09 吉田　再履修対応 ST
				'w_sSQL = w_sSQL & "		D.T19_CHIKAI_TYUKAN_Z as ZenNendoCHIKAI ,"
				'w_sSQL = w_sSQL & "		D.T19_KEKA_TYUKAN_Z as ZenNendoKEKA ,"
				'w_sSQL = w_sSQL & "		D.T19_KEKA_NASI_TYUKAN_Z as ZenNendoKEKA_NASI ,"
				''//INS 2022/03/09 吉田　再履修対応 ED
				'2023.08.24 Del 前年度成績は全試験で学年末を参照 <--
				
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI,A.T16_KEKA_KIMATU_Z AS KEKA,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T16_CHIKAI_KIMATU_Z AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_KIMATU_Z as SOUJI, A.T16_JUNJIKAN_KIMATU_Z as JYUNJI, "
				
				w_sSQL = w_sSQL & "		A.T16_DATAKBN_KIMATU_Z as DataKbn,"
				w_sSQL = w_sSQL & "		A.T16_HYOKA_KIMATU_Z as HYOKA ," '//INS 2022/10/07 吉田
				'2023.08.24 Del Kiyomoto  前年度成績は全試験で学年末を参照 -->
				'w_sSQL = w_sSQL & "		D.T19_SEI_KIMATU_Z as ZenNendoSeiseki ,"	'//INS 2022/03/02 吉田　再履修対応
				''//INS 2022/03/09 吉田　再履修対応 ST
				'w_sSQL = w_sSQL & "		D.T19_CHIKAI_KIMATU_Z as ZenNendoCHIKAI ,"
				'w_sSQL = w_sSQL & "		D.T19_KEKA_KIMATU_Z as ZenNendoKEKA ,"
				'w_sSQL = w_sSQL & "		D.T19_KEKA_NASI_KIMATU_Z as ZenNendoKEKA_NASI ,"
				''//INS 2022/03/09 吉田　再履修対応 ED
				'2023.08.24 Del 前年度成績は全試験で学年末を参照 <--
				
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI,A.T16_KEKA_TYUKAN_K AS KEKA,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_K AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_TYUKAN_K as SOUJI, A.T16_JUNJIKAN_TYUKAN_K as JYUNJI, "
				
				w_sSQL = w_sSQL & "		A.T16_DATAKBN_TYUKAN_K as DataKbn,"
				w_sSQL = w_sSQL & "		A.T16_HYOKA_TYUKAN_K as HYOKA ," '//INS 2022/10/07 吉田
				'2023.08.24 Del Kiyomoto  前年度成績は全試験で学年末を参照 -->
				'w_sSQL = w_sSQL & "		D.T19_SEI_TYUKAN_K as ZenNendoSeiseki ,"	'//INS 2022/03/02 吉田　再履修対応
				''//INS 2022/03/09 吉田　再履修対応 ST
				'w_sSQL = w_sSQL & "		D.T19_CHIKAI_TYUKAN_K as ZenNendoCHIKAI ,"
				'w_sSQL = w_sSQL & "		D.T19_KEKA_TYUKAN_K as ZenNendoKEKA ,"
				'w_sSQL = w_sSQL & "		D.T19_KEKA_NASI_TYUKAN_K as ZenNendoKEKA_NASI ,"
				'//INS 2022/03/09 吉田　再履修対応 ED
				'2023.08.24 Del 前年度成績は全試験で学年末を参照 <--
				
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_Z AS SEI_ZT,A.T16_KEKA_TYUKAN_Z AS KEKA_ZT,A.T16_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,A.T16_CHIKAI_TYUKAN_Z AS CHIKAI_ZT,A.T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI_ZT, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI_ZK,A.T16_KEKA_KIMATU_Z AS KEKA_ZK,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,A.T16_CHIKAI_KIMATU_Z AS CHIKAI_ZK,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI_ZK, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI_KT,A.T16_KEKA_TYUKAN_K AS KEKA_KT,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,A.T16_CHIKAI_TYUKAN_K AS CHIKAI_KT,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI_KT, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI_KK,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_KIMATU_K as SOUJI, A.T16_JUNJIKAN_KIMATU_K as JYUNJI, A.T16_SAITEI_JIKAN, A.T16_KYUSAITEI_JIKAN, "
				
				w_sSQL = w_sSQL & "		A.T16_KOUSINBI_TYUKAN_Z,"
				w_sSQL = w_sSQL & "		A.T16_KOUSINBI_KIMATU_Z,"
				w_sSQL = w_sSQL & "		A.T16_KOUSINBI_TYUKAN_K,"
				w_sSQL = w_sSQL & "		A.T16_KOUSINBI_KIMATU_K,"

				w_sSQL = w_sSQL & "		A.T16_DATAKBN_KIMATU_K as DataKbn,"
				w_sSQL = w_sSQL & "		A.T16_HYOKA_KIMATU_K as HYOKA ," '//INS 2022/10/07 吉田
				'2023.08.24 Del Kiyomoto  前年度成績は全試験で学年末を参照 -->
				'w_sSQL = w_sSQL & "		D.T19_SEI_KIMATU_K as ZenNendoSeiseki ,"	'//INS 2022/03/02 吉田　再履修対応

				''//INS 2022/03/09 吉田　再履修対応 ST
				'w_sSQL = w_sSQL & "		D.T19_CHIKAI_KIMATU_K as ZenNendoCHIKAI ,"
				'w_sSQL = w_sSQL & "		D.T19_KEKA_KIMATU_K as ZenNendoKEKA ,"
				'w_sSQL = w_sSQL & "		D.T19_KEKA_NASI_KIMATU_K as ZenNendoKEKA_NASI ,"
				''//INS 2022/03/09 吉田　再履修対応 ED
				'2023.08.24 Del 前年度成績は全試験で学年末を参照 <--
				
		End Select
		
		'2023.08.24 Add Kiyomoto  前年度成績は全試験で学年末を参照 -->
		w_sSQL = w_sSQL & "		D.T19_SEI_KIMATU_K as ZenNendoSeiseki ,"
		w_sSQL = w_sSQL & "		D.T19_CHIKAI_KIMATU_K as ZenNendoCHIKAI ,"
		w_sSQL = w_sSQL & "		D.T19_KEKA_KIMATU_K as ZenNendoKEKA ,"
		w_sSQL = w_sSQL & "		D.T19_KEKA_NASI_KIMATU_K as ZenNendoKEKA_NASI ,"
		'2023.08.24 Upd 前年度成績は全試験で学年末を参照 <--

		w_sSQL = w_sSQL & " 	A.T16_GAKUSEI_NO AS GAKUSEI_NO,A.T16_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_SELECT_FLG "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_LEVEL_KYOUKAN "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_OKIKAE_FLG "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_KAISETU  AS KAISETU "	'//開設時期追加 Ins 2018/03/22 Nishimura
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_MENJYO_FLG AS Menjo "	'//INS 2022/03/04 吉田　再履修対応
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		'INS 2022/03/02 吉田　再履修対応(留年者の前年度の成績を取得) ST 
		w_sSQL = w_sSQL & " 	 ,("
		w_sSQL = w_sSQL & " 	 SELECT T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " 	   ,T13_RYUNEN_FLG"
		w_sSQL = w_sSQL & " 	   ,T13_GAKUNEN"
		w_sSQL = w_sSQL & " 	   ,T13_CLASS"	
		w_sSQL = w_sSQL & " 	   ,T11_SIMEI"
		w_sSQL = w_sSQL & " 	   ,M05_CLASSMEI"
		w_sSQL = w_sSQL & " 	   ,T13_RYUNEN_FLG"
		'2023.08.24 Upd Kiyomoto 前年度成績は全試験で学年末を参照 -->
		'Select Case m_sSikenKBN
		'	Case C_SIKEN_ZEN_TYU
		'		w_sSQL = w_sSQL & " 	   ,T19_SEI_TYUKAN_Z"
		'		'//INS 2022/03/09 吉田　再履修対応 ST
		'		w_sSQL = w_sSQL & " 	   ,T19_CHIKAI_TYUKAN_Z"
		'		w_sSQL = w_sSQL & " 	   ,T19_KEKA_TYUKAN_Z"
		'		w_sSQL = w_sSQL & " 	   ,T19_KEKA_NASI_TYUKAN_Z"
		'		'//INS 2022/03/09 吉田　再履修対応 ED
		'	Case C_SIKEN_ZEN_KIM
		'		w_sSQL = w_sSQL & " 	   ,T19_SEI_KIMATU_Z"	
		'		'//INS 2022/03/09 吉田　再履修対応 ST
		'		w_sSQL = w_sSQL & " 	   ,T19_CHIKAI_KIMATU_Z"
		'		w_sSQL = w_sSQL & " 	   ,T19_KEKA_KIMATU_Z"
		'		w_sSQL = w_sSQL & " 	   ,T19_KEKA_NASI_KIMATU_Z"
		'		'//INS 2022/03/09 吉田　再履修対応 ED
		'	Case C_SIKEN_KOU_TYU
		'		w_sSQL = w_sSQL & " 	   ,T19_SEI_TYUKAN_K"
		'		'//INS 2022/03/09 吉田　再履修対応 ST
		'		w_sSQL = w_sSQL & " 	   ,T19_CHIKAI_TYUKAN_K"
		'		w_sSQL = w_sSQL & " 	   ,T19_KEKA_TYUKAN_K"
		'		w_sSQL = w_sSQL & " 	   ,T19_KEKA_NASI_TYUKAN_K"
		'		'//INS 2022/03/09 吉田　再履修対応 ED
		'	Case C_SIKEN_KOU_KIM
		'		w_sSQL = w_sSQL & " 	   ,T19_SEI_KIMATU_K"
		'		'//INS 2022/03/09 吉田　再履修対応 ST
		'		w_sSQL = w_sSQL & " 	   ,T19_CHIKAI_KIMATU_K"
		'		w_sSQL = w_sSQL & " 	   ,T19_KEKA_KIMATU_K"
		'		w_sSQL = w_sSQL & " 	   ,T19_KEKA_NASI_KIMATU_K"
		'		'//INS 2022/03/09 吉田　再履修対応 ED
		'End Select
		w_sSQL = w_sSQL & " 	   ,T19_SEI_KIMATU_K"
		w_sSQL = w_sSQL & " 	   ,T19_CHIKAI_KIMATU_K"
		w_sSQL = w_sSQL & " 	   ,T19_KEKA_KIMATU_K"
		w_sSQL = w_sSQL & " 	   ,T19_KEKA_NASI_KIMATU_K"
		'2023.08.24 Upd 前年度成績は全試験で学年末を参照 <--

		w_sSQL = w_sSQL & "	 	  FROM "
		w_sSQL = w_sSQL & " 	  	T11_GAKUSEKI "
		w_sSQL = w_sSQL & " 	   ,T13_GAKU_NEN "
		w_sSQL = w_sSQL & " 	   ,M05_CLASS "
		w_sSQL = w_sSQL & " 	   ,T19_RISYURYUNEN_KOJIN "
		w_sSQL = w_sSQL & "	 	  WHERE"
		w_sSQL = w_sSQL & " 		T13_NENDO = " & Cint(m_iNendo) - 1 & " "
		w_sSQL = w_sSQL & " 		AND	T13_RYUNEN_FLG = "  & C_RYUNEN_ON
		w_sSQL = w_sSQL & "  		AND T11_GAKUSEI_NO = T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & "  		AND M05_NENDO = T13_NENDO "
		w_sSQL = w_sSQL & "  		AND M05_GAKUNEN = T13_GAKUNEN "
		w_sSQL = w_sSQL & "  		AND M05_CLASSNO = T13_CLASS "
		w_sSQL = w_sSQL & "  		AND T13_NENDO = T19_NENDO "
		w_sSQL = w_sSQL & "  		AND T13_GAKUSEI_NO = T19_GAKUSEI_NO "
		w_sSQL = w_sSQL & "  		AND T19_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " 	 )D "
		'INS 2022/03/02 吉田　ED
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T16_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_NENDO = C.T13_NENDO "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = D.T13_GAKUSEI_NO(+) " 'INS 2022/03/02 吉田　再履修対応
		
		'//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
		w_sSQL = w_sSQL & " AND	A.T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO

		'//必修か選択科目のうち選択している学生のみを取得する		'INS 2019/03/06 藤林
		w_sSQL = w_sSQL & " AND	( T16_HISSEN_KBN = " & C_HISSEN_HIS
		w_sSQL = w_sSQL & "       OR (T16_HISSEN_KBN = " & C_HISSEN_SEN & " AND T16_SELECT_FLG = 1) "
		w_sSQL = w_sSQL & " 	) "

		w_sSQL = w_sSQL & " ORDER BY A.T16_GAKUSEKI_NO "
		'   response.write w_sSQL & "<BR>"
		'   response.end
		If gf_GetRecordset(m_Rs, w_sSQL) <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getdate = 99
			m_bErrFlg = True
			Exit Do 
		End If
		
		m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
		m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))
		
		'//ﾚｺｰﾄﾞカウント取得
		m_rCnt=gf_GetRsCount(m_Rs)
		
		f_getdate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	特別活動受講学生取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_getTUKUclass(p_iNendo,p_sKamokuCd,p_iGakunen,p_iClass)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet
	
    On Error Resume Next
    Err.Clear
    
    f_getTUKUclass = 1
	p_sTKyokanCd = ""
	
	Do
		
        '//検索結果の値より一覧を表示
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "

		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_Z AS KEKA,A.T34_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,A.T34_CHIKAI_TYUKAN_Z AS CHIKAI, "
				w_sSQL = w_sSQL & "		A.T34_SOJIKAN_TYUKAN_Z as SOUJI, A.T34_JUNJIKAN_TYUKAN_Z as JYUNJI, "
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_Z AS KEKA,A.T34_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T34_CHIKAI_KIMATU_Z AS CHIKAI, "
				w_sSQL = w_sSQL & "		A.T34_SOJIKAN_KIMATU_Z as SOUJI, A.T34_JUNJIKAN_KIMATU_Z as JYUNJI, "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_K AS KEKA,A.T34_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T34_CHIKAI_TYUKAN_K AS CHIKAI, "
				w_sSQL = w_sSQL & "		A.T34_SOJIKAN_TYUKAN_K as SOUJI, A.T34_JUNJIKAN_TYUKAN_K as JYUNJI, "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_Z AS KEKA_ZT,A.T34_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,A.T34_CHIKAI_TYUKAN_Z AS CHIKAI_ZT, "
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_Z AS KEKA_ZK,A.T34_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,A.T34_CHIKAI_KIMATU_Z AS CHIKAI_ZK, "
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_K AS KEKA_KT,A.T34_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,A.T34_CHIKAI_TYUKAN_K AS CHIKAI_KT, "
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_K AS KEKA,A.T34_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T34_CHIKAI_KIMATU_K AS CHIKAI, "
				w_sSQL = w_sSQL & "		A.T34_SOJIKAN_KIMATU_K as SOUJI, A.T34_JUNJIKAN_KIMATU_K as JYUNJI, A.T34_SAITEI_JIKAN, A.T34_KYUSAITEI_JIKAN, "
		End Select

		w_sSQL = w_sSQL & " 	A.T34_GAKUSEI_NO AS GAKUSEI_NO,A.T34_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI"
		w_sSQL = w_sSQL & "     ,0  AS KAISETU "	'//開設時期追加 Ins 2018/03/22 Nishimura
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T34_RISYU_TOKU A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T34_NENDO = " & Cint(p_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T34_TOKUKATU_CD = '" & p_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T34_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T34_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(p_iGakunen) & " "
		w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(p_iClass) & " "
		
		w_sSQL = w_sSQL & " AND	A.T34_NENDO = C.T13_NENDO "
		
		w_sSQL = w_sSQL & " ORDER BY A.T34_GAKUSEKI_NO "
		
		'response.write w_sSQL & "<BR>"
		
		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getTUKUclass = 99
			m_bErrFlg = True
			Exit Do 
		End If
		
		'//最初の生徒の学籍番号を取得
		if not m_Rs.EOF then
			m_FirstGakusekiNo = m_Rs("GAKUSEKI_NO")
			m_Rs.movefirst
		end if
		
		m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
		m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))
		
		'//ﾚｺｰﾄﾞカウント取得
		m_rCnt=gf_GetRsCount(m_Rs)

		f_getTUKUclass = 0
		Exit Do
	Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[機能]	科目担当教官の教官CDの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_GetTantoKyokan(p_sTKyokanCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetTantoKyokan = 1
	p_sTKyokanCd = ""

    Do 
		'//科目担当教官の教官CDの取得
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & "  T20_KYOKAN "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & "  T20_JIKANWARI "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & "  T20_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND T20_KAMOKU = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND T20_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND T20_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " GROUP BY T20_KYOKAN "

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetTantoKyokan = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_sTKyokanCd = w_Rs("T20_KYOKAN")
		End If

        f_GetTantoKyokan = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[機能]	科目担当教官の教官CDの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_GetTantoKyokan2(p_sTKyokanCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetTantoKyokan = 1
	p_sTKyokanCd = ""

    Do 
		'//科目担当教官の教官CDの取得
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & "  T27_KYOKAN_CD "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & "  T27_TANTO_KYOKAN "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & "  T27_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND T27_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND T27_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND T27_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " GROUP BY T27_KYOKAN_CD "

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetTantoKyokan = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then

			p_sTKyokanCd = w_Rs("T27_KYOKAN_CD")
		End If

        f_GetTantoKyokan = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

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
	m_bKekkaNyuryokuFlg = False		'欠課入力ﾌﾗｸﾞ

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO "
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
			m_iKekkaKaishi = gf_SetNull2String(m_DRs("T24_KEKKA_KAISI"))	'欠課入力開始
			m_iKekkaSyuryo = gf_SetNull2String(m_DRs("T24_KEKKA_SYURYO"))	'欠課入力終了
			w_sSysDate = gf_SetNull2String(m_DRs("SYSDATE"))				'システム日付
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
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_getTUKU(p_iNendo,p_sKamoku,p_iGakunen,p_iClass,p_TUKU_FLG)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

	On Error Resume Next
	Err.Clear
	f_getTUKU = 0
	p_TUKU_FLG = C_TUKU_FLG_TUJO

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T20_TUKU_FLG "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T20_JIKANWARI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T20_NENDO=" & Cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_KAMOKU ='" & p_sKamoku & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_GAKUNEN =" & Cint(p_iGakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_CLASS =" & Cint(p_iClass)

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getTUKU = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If w_Rs.EOF = false Then
			p_TUKU_FLG = cStr(gf_SetNull2Zero(w_Rs("T20_TUKU_FLG")))
		End If

		Exit Do
	Loop
    Call gf_closeObject(w_Rs)

End Function


'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'********************************************************************************
Function f_Syukketu2(p_gaku,p_kbn)

	Dim w_GAKUSEI_NO
	Dim w_SYUKKETU_KBN

	f_Syukketu2 = 0
	w_GAKUSEI_NO = ""
	w_SYUKKETU_KBN = ""
	w_SKAISU = ""

	On Error Resume Next
	Err.Clear

	If m_SRs.EOF Then
		Exit Function
	Else
'		m_SRs.MoveFirst
		Do Until m_SRs.EOF

		w_GAKUSEI_NO = m_SRs("T21_GAKUSEKI_NO")
		w_SYUKKETU_KBN = m_SRs("T21_SYUKKETU_KBN")
		w_SKAISU = m_SRs("KAISU")

			If Cstr(w_GAKUSEI_NO) = Cstr(p_gaku) AND cstr(w_SYUKKETU_KBN) = cstr(p_kbn) Then
				f_Syukketu2 = w_SKAISU

				Exit Do
			End If
			m_SRs.MoveNext
		Loop
		
		m_SRs.MoveFirst
	End If

End Function


'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'********************************************************************************
Function f_Syukketu2New(p_gaku,p_kbn)

	Dim w_GAKUSEI_NO
	Dim w_SYUKKETU_KBN

	f_Syukketu2New = ""
	w_GAKUSEI_NO = ""
	w_SYUKKETU_KBN = ""
	w_SKAISU = ""

	If m_SRs.EOF Then
		Exit Function
	Else
		Do Until m_SRs.EOF

			w_GAKUSEI_NO = m_SRs("T21_GAKUSEKI_NO")
			w_SYUKKETU_KBN = m_SRs("T21_SYUKKETU_KBN")
			w_SKAISU = gf_SetNull2String(m_SRs("KAISU"))

			If Cstr(w_GAKUSEI_NO) = Cstr(p_gaku) AND cstr(w_SYUKKETU_KBN) = cstr(p_kbn) Then
				f_Syukketu2New = w_SKAISU
				Exit Do
			End If

			m_SRs.MoveNext
		Loop

		m_SRs.MoveFirst
	End If

End Function


'********************************************************************************
'*  [機能]  確定欠課数、遅刻数を取得。
'*  [引数]  p_iNendo　 　：　処理年度
'*          p_iSikenKBN　：　試験区分
'*          p_sKamokuCD　：　科目コード
'*          p_sGakusei 　：　５年間番号
'*  [戻値]  p_iKekka   　：　欠課数
'*          p_ichikoku 　：　遅刻回数
'*          0：正常修了
'*  [説明]  試験区分に入っている、欠課数、遅刻数を取得する。
'*			2002.03.20
'*			NULLを0に変換しないために、関数をモジュール内で作成（CACommon.aspからコピー）
'********************************************************************************
Function f_GetKekaChi(p_iNendo,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku,p_iKekkaGai)
	Dim w_sSQL
    Dim w_KekaChiRs
    Dim w_sKek,p_sChi
	Dim w_sSouG,w_sJyunG
	Dim w_Table,w_TableName
    Dim w_Kamoku
    
    On Error Resume Next
    Err.Clear
    
    p_iKekka = ""
    p_iChikoku = ""
	
	'特別授業、その他(通常など)の切り分け
	if trim(m_sSyubetu) = "TOKU" then
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_Kamoku = "T34_TOKUKATU_CD"
	else
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_Kamoku = "T16_KAMOKU_CD"
	end if
	
    f_GetKekaChi = 1
	
	'/試験区分によって取ってくる、フィールドを変える。
	Select Case p_iSikenKBN
		Case C_SIKEN_ZEN_TYU
			w_sKek   = w_Table & "_KEKA_TYUKAN_Z"
			w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_Z"
			p_sChi   = w_Table & "_CHIKAI_TYUKAN_Z"
			w_sSouG  = w_Table & "_SOJIKAN_TYUKAN_Z"
			w_sJyunG = w_Table & "_JUNJIKAN_TYUKAN_Z"
		Case C_SIKEN_ZEN_KIM
			w_sKek   = w_Table & "_KEKA_KIMATU_Z"
			w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_Z"
			p_sChi   = w_Table & "_CHIKAI_KIMATU_Z"
			w_sSouG  = w_Table & "_SOJIKAN_KIMATU_Z"
			w_sJyunG = w_Table & "_JUNJIKAN_KIMATU_Z"
		Case C_SIKEN_KOU_TYU
			w_sKek   = w_Table & "_KEKA_TYUKAN_K"
			w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_K"
			p_sChi   = w_Table & "_CHIKAI_TYUKAN_K"
			w_sSouG  = w_Table & "_SOJIKAN_TYUKAN_K"
			w_sJyunG = w_Table & "_JUNJIKAN_TYUKAN_K"
		Case C_SIKEN_KOU_KIM
			w_sKek   = w_Table & "_KEKA_KIMATU_K"
			w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_K"
			p_sChi   = w_Table & "_CHIKAI_KIMATU_K"
			w_sSouG  = w_Table & "_SOJIKAN_KIMATU_K"
			w_sJyunG = w_Table & "_JUNJIKAN_KIMATU_K"
	End Select
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & 	w_sKek   & " as KEKA, "
	w_sSQL = w_sSQL & 	w_sKekG  & " as KEKA_NASI, "
	w_sSQL = w_sSQL & 	p_sChi   & " as CHIKAI, "
	w_sSQL = w_sSQL & 	w_sSouG  & " as SOUJI, "
	w_sSQL = w_sSQL & 	w_sJyunG & " as JYUNJI "
	w_sSQL = w_sSQL & " FROM "   & w_TableName
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "      " & w_Table & "_NENDO =" & p_iNendo
	w_sSQL = w_sSQL & "  AND " & w_Table & "_GAKUSEI_NO= '" & p_sGakusei & "'"
	w_sSQL = w_sSQL & "  AND " & w_Kamoku & "= '" & p_sKamokuCD & "'"
	
	If gf_GetRecordset(w_KekaChiRs, w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		f_GetKekaChi = 99
	End If
	' response.write "w_sSQL:" & w_sSQL & "<BR>"
	'//戻り値ｾｯﾄ
	If w_KekaChiRs.EOF = False Then
		p_iKekka = gf_SetNull2String(w_KekaChiRs("KEKA"))
		p_iKekkaGai = gf_SetNull2String(w_KekaChiRs("KEKA_NASI"))
		p_iChikoku = gf_SetNull2String(w_KekaChiRs("CHIKAI"))
		
		m_iSouJyugyou = gf_SetNull2String(w_KekaChiRs("SOUJI"))
		m_iJunJyugyou = gf_SetNull2String(w_KekaChiRs("JYUNJI"))
	End If
	
	f_GetKekaChi = 0
	
	Call gf_closeObject(w_KekaChiRs)
	
End Function

'2022/10/11 INS -->
'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'[引数] pData1 : データ１
'       pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'****************************************************
Function f_Selected(pData1,pData2)

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else 
            f_Selected = "" 
        End If
    End If

End Function
'2022/10/11 INS <--

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	Dim w_sGakusekiCd
	Dim w_sSeiseki
	Dim w_sHyoka
	Dim w_sKekka,w_sKekkaGai
	Dim w_sChikai
	Dim w_sKekkasu
	Dim w_sChikaisu
	Dim w_sShikenKBN_RUI
	Dim w_iKekka_rui,w_iChikoku_rui

	Dim w_ihalf
	Dim i

	Dim w_lSeiTotal	'成績合計
	Dim w_lGakTotal	'学生人数

	'データがNULLの場合に0に変換しないために、一旦データを保存するワークで使用
	'2002.03.20
	Dim w_sData 	
	dim w_sData2
	Dim w_DataKbn
	Dim w_Checked

	Dim w_Padding
	Dim w_Padding2
	Dim w_Disabled
	Dim w_Disabled2
	Dim w_TableWidth
	
	on error resume next
	
	w_Padding = "style='padding:2px 0px;'"
	w_Padding2 = "style='padding:2px 0px;font-size:10px;'"

	w_lSeiTotal = 0
	w_lGakTotal = 0

	i = 1
	
	if m_SchoolFlg and m_TUKU_FLG = C_TUKU_FLG_TUJO then
		w_TableWidth = 760
		
		if cint(gf_SetNull2Zero(m_Rs("DataKbn"))) = cint(C_MIHYOKA) or m_iKikan = "NO" then
			w_Disabled = "disabled"
		end if
	else
	' 2022.03.08 UPD
		' w_TableWidth = 710
		if m_TUKU_FLG = C_TUKU_FLG_TUJO then 
			w_TableWidth = 875
		else 
			w_TableWidth = 710
		end if
	end if
	
%>
<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
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
		
		//成績合計値の取得
		f_GetTotalAvg();
		
		// 総時間と純時間をhiddenにセット
		document.frm.hidSouJyugyou.value = "<%= m_iSouJyugyou %>";
		document.frm.hidJunJyugyou.value = "<%= m_iJunJyugyou %>";
		
		var scrollbarWidth = window.innerWidth - document.body.clientWidth;
		// alert('window.innerWidth: ' + window.innerWidth +'\n document.body.clientWidth: ' + document.body.clientWidth
		// +'\n scrollbarWidth: ' + scrollbarWidth);

		document.frm.scrollbarWidth.value = scrollbarWidth;
        //submit
        document.frm.target = "topFrame";
        document.frm.action = "sei0100_middle.asp"
        document.frm.submit();
        
        return;
		
    }
	//************************************************************
    //  [機能]  評価ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_change(p_iS){
		w_sButton = eval("document.frm.button"+p_iS);
        w_sHyouka = eval("document.frm.Hyoka"+p_iS);

<%If m_sSikenKBN = C_SIKEN_ZEN_TYU Then%>

        if(w_sButton.value == "・") {
			w_sButton.value = "○";
			w_sHyouka.value = "○";
	        return;
		}
        if(w_sButton.value == "○") {
			w_sButton.value = "・";
			w_sHyouka.value = "";
	        return;
		}

<%Else%>

        if(w_sButton.value == "・") {
			w_sButton.value = "○";
			w_sHyouka.value = "○";
	        return;
		}
        if(w_sButton.value == "○") {
			w_sButton.value = "◎";
			w_sHyouka.value = "◎";
	        return;
		}
        if(w_sButton.value == "◎") {
			w_sButton.value = "・";
			w_sHyouka.value = "";
	        return;
		}
<%End If%>

    }
    
   //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){
		var ob,w_num;
		var i;
		var w_Hyoka;
		var w_hidHyoka;
		var indx;
		var HyokaArray =[];

		if(f_CheckData_All() == 1){
	            alert("入力値が不正です");
	            return 1;
		}else{
			if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return;}
			
			document.frm.hidSouJyugyou.value = parent.topFrame.document.frm.txtSouJyugyou.value;
			document.frm.hidJunJyugyou.value = parent.topFrame.document.frm.txtJunJyugyou.value;
			
			//ヘッダ部空白表示
			parent.topFrame.document.location.href="white.asp";
				
			//登録処理
		if(document.frm.hidSeisekiInp.value == "<%=C_SEISEKI_INP_TYPE_STRING%>"){	
			for (i = 1; i < document.frm.i_Max.value; i++) {
				w_Hyoka = eval("document.frm.Hyoka"+(i));
				indx = w_Hyoka.selectedIndex;
				//m_SeisekiIndex =  w_Hyoka.options[indx].value;
				HyokaArray[i-1] =  w_Hyoka.options[indx].value;
			}
			 //alert(HyokaArray);
		}
		<% if m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
			document.frm.hidHyoka.value = HyokaArray;
			document.frm.hidUpdMode.value = "TUJO";
			document.frm.action="sei0100_upd.asp";
		<% Else %>
			// document.frm.hidHyoka.value = HyokaArray;
			document.frm.hidUpdMode.value = "TOKU";
			document.frm.action="sei0100_upd_toku.asp";
		<% End if %>
	        document.frm.target="main";
			document.frm.LOG_SOSA.value = "登録";		//add 2023/12/14 吉田
			document.frm.LOG_TAISYO.value = document.frm.LOG_TAISYO.value;		//add 2023/12/14 吉田
	        document.frm.submit();
		}
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

    //************************************************************
    //  [機能]  入力値のﾁｪｯｸ(登録ボタン押下時)
    //  [引数]  なし
    //  [戻値]  0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
    //  [説明]  入力値のNULLﾁｪｯｸ、英数字ﾁｪｯｸ、桁数ﾁｪｯｸを行う
    //          引渡ﾃﾞｰﾀ用にﾃﾞｰﾀを加工する必要がある場合には加工を行う
    //************************************************************
    function f_CheckData_All() {
		var i
		var w_Seiseki
		var w_bFLG

		// 総時間・純時間入力チェック
		if(!f_CheckNum("parent.topFrame.document.frm.txtSouJyugyou")){ return 1; }
		if(!f_CheckNum("parent.topFrame.document.frm.txtJunJyugyou")){ return 1; }
		if(!f_CheckDaisyou()){ return 1; }

	<% if m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
		for (i = 1; i < document.frm.i_Max.value; i++) {
			
			w_Seiseki = eval("document.frm.Seiseki"+i);
			w_bFLG = true

			if (w_Seiseki){		//2001/12/17 Add
				if (isNaN(w_Seiseki.value)){
					w_bFLG = false;
					w_Seiseki.focus();
					return 1;
					break;
				}else{
					//上限値をチェック 2001/12/09 追加 伊藤
					//var wStr = new String(w_Seiseki.value)
					if (w_Seiseki.value > 100){
						w_bFLG = false;
						w_Seiseki.focus();
						return 1;
						break;
					};

					//マイナスをチェック
					var wStr = new String(w_Seiseki.value)
					if (wStr.match("-")!=null){
						w_bFLG = false;
						w_Seiseki.focus();
						return 1;
						break;
					};

					//小数点チェック
					w_decimal = new Array();
					w_decimal = wStr.split(".")
					if(w_decimal.length>1){
						w_bFLG = false;
						w_Seiseki.focus();
						return 1;
						break;
					};
				};
			};
		};
		if (w_bFLG == false){
			return 1;
		};
	<% End if %>

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_Chikai = eval("document.frm.Chikai"+i);
			w_bFLG = true
if (w_Chikai){		//2001/12/17 Add
			if (isNaN(w_Chikai.value)){
				w_bFLG = false;
				w_Chikai.focus();
				return 1;
				break;
			}else{

				//マイナスをチェック
				var wStr = new String(w_Chikai.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					w_Chikai.focus();
					return 1;
					break;
				};

				//小数点チェック
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					w_Chikai.focus();
					return 1;
					break;
				}

			};
};
		};

			if (w_bFLG == false){
				return 1;
			};

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {

			w_Kekka = eval("document.frm.Kekka"+i);
			w_bFLG = true
if (w_Kekka){		//2001/12/17 Add
			if (isNaN(w_Kekka.value)){
				w_bFLG = false;
				w_Kekka.focus();
				return 1;
				break;
			}else{

				//マイナスをチェック
				var wStr = new String(w_Kekka.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					w_Kekka.focus();
					return 1;
					break;
				};

				//小数点チェック
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					w_Kekka.focus();
					return 1;
					break;
				}

			};
};
		};
			if (w_bFLG == false){
				return 1;
			};

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_KekkaGai = eval("document.frm.KekkaGai"+i);
			w_bFLG = true
if (w_KekkaGai){		//2001/12/17 Add
			if (isNaN(w_KekkaGai.value)){
				w_bFLG = false;
				w_KekkaGai.focus();
				return 1;
				break;
			}else{

				//マイナスをチェック
				var wStr = new String(w_KekkaGai.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					w_KekkaGai.focus();
					return 1;
					break;
				};

				//小数点チェック
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					w_KekkaGai.focus();
					return 1;
					break;
				}

			};
};
		};
			if (w_bFLG == false){
				return 1;
			};

		return 0;
	};

    //************************************************************
    //  [機能]  簡易数値型チェック
    //************************************************************
	function f_CheckNum(pFromName){

		wFromName = eval(pFromName);
		if (isNaN(wFromName.value)){
			wFromName.focus();
			return false;
		}else{

			//マイナスをチェック
			var wStr = new String(wFromName.value)
			if (wStr.match("-")!=null){
				wFromName.focus();
				return false;
			};

			//小数点チェック
			w_decimal = new Array();
			w_decimal = wStr.split(".")
			if(w_decimal.length>1){
				wFromName.focus();
				return false;
			};
		}
		return true;
	}

    //************************************************************
    //  [機能]  大小チェック
    //************************************************************
	function f_CheckDaisyou(){

		wObj1 = eval("parent.topFrame.document.frm.txtSouJyugyou");
		wObj2 = eval("parent.topFrame.document.frm.txtJunJyugyou");

		if(wObj1.value != "" && wObj2.value != ""){
			if(wObj1.value-0 < wObj2.value-0){
				wObj1.focus();
				return false;
			}
		}
		return true;
	}

	//************************************************
	//Enter キーで下の入力フォームに動くようになる
	//引数：p_inpNm	対象入力フォーム名
	//    ：p_frm	対象フォーム
	//　　：i		現在の番号
	//戻値：なし
	//入力フォーム名が、xxxx1,xxxx2,xxxx3,…,xxxxn 
	//の名前のときに利用できます。
	//************************************************
	function f_MoveCur(p_inpNm,p_frm,i){
		if (event.keyCode == 13){		//押されたキーがEnter(13)の時に動く。
			i++;

			//入力可能のテキストボックスを探す。見つかったらフォーカスを移して処理を抜ける。
	        for (w_li = 1; w_li <= 99; w_li++) {
				
				if (i > <%=m_rCnt%>) i = 1; //iが最大値を超えると、はじめに戻る。
				inpForm = eval("p_frm."+p_inpNm+i);

				//入力可能領域ならフォーカスを移す。
				if (typeof(inpForm) != "undefined") {
					inpForm.focus();			//フォーカスを移す。
					inpForm.select();			//移ったテキストボックス内を選択状態にする。
					break;

				//入力付加なら次の項目へ
				} else{
					i++
				}
	        }
		}else{
	//		alert(event.keyCode);
			return false;
		}
		return true;
	}
	
	
	//************************************************
	//評価不能がクリックされたときの処理
	//************************************************
	function f_InpDisabled(p_num){
		var ob = new Array();
		
		ob[0] = eval("document.frm.chkHyokaFuno" + p_num);
		ob[1] = eval("document.frm.Seiseki" + p_num);
		//ob[2] = eval("document.frm.button" + p_num);
		//ob[3] = eval("document.frm.Hyoka" + p_num);
		
		if(ob[0].checked){
			ob[1].value = "";
			ob[1].disabled = true;
			
			//ob[2].value = "・";
			//ob[2].disabled = true;
			//ob[3].value = "";
		}else{
			ob[1].disabled = false;
			//ob[2].disabled = false;
		}
		
		f_GetTotalAvg();
		
	}
		
	//-->
	</SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<center>
	<form name="frm" method="post">
	
	<table width="<%=w_TableWidth%>">
	<td>
	
	<table class="hyo" align="center" width="<%=w_TableWidth%>" border="1">
	
	<%	m_Rs.MoveFirst
		Do Until m_Rs.EOF
			w_ihalf = gf_Round(m_rCnt / 2,0)
			j = j + 1 
			w_sSeiseki = ""
			w_sHyoka = ""
			w_sKekka = ""
			w_sChikai = ""
			w_sGakusekiCd = ""
			w_sKekkasu = ""
			w_sChikaisu = ""
			Call gs_cellPtn(w_cell)
%>
	<tr>
<%
			'//各データを取得する
			'** 一つ前の試験区分
			Select Case m_sSikenKBN
				Case C_SIKEN_ZEN_TYU								'//前期中間
					w_sShikenKBN_RUI = 99
					
				Case C_SIKEN_ZEN_KIM								'//前期期末
					w_sShikenKBN_RUI = C_SIKEN_ZEN_TYU
					
				Case C_SIKEN_KOU_TYU								'//後期中間
					w_sShikenKBN_RUI = C_SIKEN_ZEN_KIM
					
				Case C_SIKEN_KOU_KIM								'//後期期末
					w_sShikenKBN_RUI = C_SIKEN_KOU_TYU
			End Select
			
			'/**** 以下表示部分で成績、欠課の表示をNULL->0変換しないでNULLは""で表示する 2002.03.20 matsuo ****/
			
			w_sGakusekiCd = m_Rs("GAKUSEKI_NO")
			
			w_sKekka = gf_SetNull2String(m_Rs("KEKA"))
			w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI"))
			w_sChikai = gf_SetNull2String(m_Rs("CHIKAI"))

			'//前期で終わっている科目の欠課を取得して学期末成績にセットする。2002/02/21 ITO
			'//前期のみの場合はT21より前記期末試験までの欠課数を取得する
			
			'w_iRet = f_SikenInfo(w_bZenkiOnly)
			If cint(m_Rs("KAISETU")) = C_KAI_ZENKI Then	'2020.03.09 Upd Kiyomoto 開設時期をT16レコードから取得する
				w_bZenkiOnly = True
			End if 
			
			'学年末試験の場合のみ
			If m_sSikenKBN = C_SIKEN_KOU_KIM Then
				
				'前期開設だったら前期期末の欠課を学年末の成績にセットする
				If w_bZenkiOnly = True Then

					'学期末成績が更新されていない場合、前期の欠課、遅刻を学年末にセットする。
					If gf_SetNull2String(m_Rs("T16_KOUSINBI_KIMATU_K")) = "" Then 
						w_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))			'欠課数
						w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK"))	'欠課対象外
						w_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))		'遅刻回数
					End If

					'学期末成績が0
'					If gf_SetNull2String(m_Rs("KEKA")) = "" Then 
'						w_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))			'欠課数
'						w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK"))	'欠課対象外
'						w_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))		'遅刻回数
'					End If
					
				End If
			End If
			
			'前期で終わっている科目の欠課を取得して学期末成績にセットする。2002/02/21 ITO
			'値の初期化。
			w_bNoChange = False
			w_sKekkasu = ""
			w_sChikaisu = ""
			
			'---------------------------------------------------------------------------------------------
			'通常授業ときの処理
			if m_TUKU_FLG = C_TUKU_FLG_TUJO then 
				w_sSeiseki = gf_SetNull2String(m_Rs("SEI"))
				w_sHyoka = gf_HTMLTableSTR(m_Rs("HYOKAYOTEI"))
				
				'前期で終わっている科目の欠課を取得して学期末成績にセットする。2002/02/21 ITO
				
				'学年末試験の場合のみ
				If m_sSikenKBN = C_SIKEN_KOU_KIM Then
					
					'前期開設だったら前期期末の欠課を学年末の成績にセットする
					If w_bZenkiOnly = True Then
						'学期末成績が更新されていない場合、前期の成績を学年末にセットする。
						If gf_SetNull2String(m_Rs("T16_KOUSINBI_KIMATU_K")) = "" Then 
							w_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))			'前期期末成績
						End If
						'学期末成績が0
'						If gf_SetNull2String(m_Rs("SEI")) = "" Then 
'							w_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))			'前期期末成績
'						End If
					End If
				End If
				
				'前期で終わっている科目の欠課を取得して学期末成績にセットする。2002/02/21 ITO
				if w_sHyoka = "　" then w_sHyoka = "・"
				
				'//科目が選択科目の場合は、生徒が選択しているかどうかを判別する。選択しいない生徒は入力不可とする。
				w_bNoChange = False
				
				If cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then 
					If cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then
						w_bNoChange = True
					End If 
				Else
					if Cstr(m_iLevelFlg) = "1" then
						if isNull(m_Rs("T16_LEVEL_KYOUKAN")) = true then
							w_bNoChange = True
						else
							if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
								w_bNoChange = True
							End if
						End if
					End if
				End If
				
			End if
			
			'==異動ＣＨＫ（2001/12/19日バージョン:okada）================================
			Dim w_Date
			Dim w_SSSS
			Dim w_SSSR
			
			w_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")
			'//C_IDO_FUKUGAKU=3:復学、C_IDO_TEI_KAIJO=5:停学解除
			w_SSSS = ""
			w_SSSR = ""
			
			w_SSSS = gf_Get_IdouChk(w_sGakusekiCd,w_Date,m_iNendo,w_SSSR)
			
			IF CStr(w_SSSS) <> "" Then
				IF Cstr(w_SSSS) <> CStr(C_IDO_FUKUGAKU) AND Cstr(w_SSSS) <> Cstr(C_IDO_TEI_KAIJO) AND Cstr(w_SSSS) <> Cstr(C_IDO_TENKO) _
					AND Cstr(w_SSSS) <> Cstr(C_IDO_TENKA) AND Cstr(w_SSSS) <> Cstr(C_IDO_KOKUHI) AND Cstr(w_SSSS) <> Cstr(C_IDO_NYUGAKU) _
					AND  Cstr(w_SSSS) <> Cstr(C_IDO_TENNYU) Then
					w_SSSS = "[" & w_SSSR & "]"
					w_bNoChange = True
				Else
					w_SSSR = ""
					w_SSSS = ""
				End if
			End if
			
			'通常授業
			if Cstr(m_TUKU_FLG) = Cstr(C_TUKU_FLG_TUJO) then 
				
				'欠課遅刻数の取得
				
				'//欠課数×単位数の取得
				w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_KEKKA)		'戻り値はNULLの時は""
				
				'gf_IIFに渡すときにパラメータを計算するので、パラメータは0に変換
				w_sKekkasu = gf_IIF(w_sData = "", "", cint(gf_SetNull2Zero(w_sData)) * cint(m_iJigenTani))

'response.write(w_sGakusekiCd & ":" & " T21 欠課 w_sKekkasu = " & w_sKekkasu & "<BR>")
				
				'//１欠課の場合の欠課数の取得
				w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_KEKKA_1)
				if w_sKekkasu = "" and w_sData = "" then
					w_sKekkasu = ""
				else
					'どちらか一方""でなければ計算
					w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_sData))			'//１欠課の場合の欠課数の取得
				end if

'response.write(w_sGakusekiCd & ":" & " T21 1欠課 w_sKekkasu = " & w_sKekkasu & "<BR>")
				
				'//遅刻数の取得
				w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_TIKOKU)
				w_sChikaisu = gf_IIF(w_sData = "", "", cint(gf_SetNull2Zero(w_sData)))		
				
				'//早退数の取得
				w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_SOTAI)
				if w_sChikaisu = "" and w_sData = "" then
					'w_sKekkasuもw_sDataもどちらも""の時は""のまま
					w_sChikaisu = ""
				else
					'どちらか一方""でなければ計算
					w_sChikaisu = cint(gf_SetNull2Zero(w_sChikaisu)) + cint(gf_SetNull2Zero(w_sData))			'//１欠課の場合の欠課数の取得
				end if
			end if		
			'---------------------------------------------------------------------------------------------
			
			'「出欠欠課が累積」で「前期中間でない」の場合
			'欠課・欠席がNullだった場合、落ちるため関数追加 Add 2001.12.16 okada
			if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and w_sShikenKBN_RUI <> 99 then 
				'一つ前の試験の合計値を足す。
				call f_GetKekaChi(m_iNendo,m_iShikenInsertType,m_sKamokuCd,cstr(m_Rs("GAKUSEI_NO")),w_iKekka_rui,w_iChikoku_rui,w_iKekkaGai_rui)		
				'どちらも""の時は""
				if w_sKekkasu = "" and w_iKekka_rui = "" then
					w_sKekkasu = ""
				else
					w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))
				end if

'response.write(w_sGakusekiCd & ":" & " T16 w_sKekkasu = " & w_sKekkasu & "<BR><BR>")
				'どちらも""の時は""
				if w_sChikaisu = "" and w_iChikoku_rui = "" then
					w_sChikaisu = ""
				else
					w_sChikaisu = cint(gf_SetNull2Zero(w_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
				end if
			end if
			
			'If cint(gf_SetNull2Zero(w_sKekka)) = 0 and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then 		'//欠入が0で,欠計が0より大きい場合
			If gf_SetNull2String(w_sKekka) = "" and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then 		'//欠入が0で,欠計が0より大きい場合
				w_sKekka = cint(gf_SetNull2Zero(w_sKekkasu))								'//欠入＝欠計
			End If
			
			'If cint(gf_SetNull2Zero(w_sChikai)) = 0 AND cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then		'//遅入が0で,遅計が0より大きい場合
			If gf_SetNull2String(w_sChikai) = "" and cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then		'//遅入が0で,遅計が0より大きい場合
				w_sChikai = cint(gf_SetNull2Zero(w_sChikaisu))							'//遅入＝遅計
			End If				
			
			if m_SchoolFlg and m_TUKU_FLG = C_TUKU_FLG_TUJO then
				'//評価不能データ設定
				w_DataKbn = 0
				w_Checked = ""
				w_Disabled2 = ""
				
				w_DataKbn = cint(gf_SetNull2Zero(m_Rs("DataKbn")))
				
				if w_DataKbn = cint(C_HYOKA_FUNO) then
					w_Checked = "checked"
					w_Disabled2 = "disabled"
					
				elseif w_DataKbn = cint(C_MIHYOKA) then
					w_Disabled2 = "disabled"
					
				end if
				
				select case Cstr(w_SSSS)
					case Cstr(C_IDO_KYU_BYOKI),Cstr(C_IDO_KYU_HOKA)
						w_DataKbn = C_KYUGAKU
						
					case Cstr(C_IDO_TAI_2NEN),Cstr(C_IDO_TAI_HOKA),Cstr(C_IDO_TAI_SYURYO)
						w_DataKbn = C_TAIGAKU
				end select
			end if
			
			'========================================================================================
			'//科目が選択科目の時に科目を選択していない場合(入力不可)
			'========================================================================================
			If w_bNoChange = True Then
				if Cstr(m_TUKU_FLG) = Cstr(C_TUKU_FLG_TUJO) Then%>

					<input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>">
					<input type="hidden" name="txtKaisetu<%=i%>" value="<%=m_Rs("KAISETU")%>">
					<input type="hidden" name="hidUpdFlg<%=i%>" value="False">
					<td class="<%=w_cell%>" align="center" width="65"  nowrap <%=w_Padding%>><%=w_sGakusekiCd%></td>
					<td class="<%=w_cell%>" align="left"   width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
					<%' 2022.03.04 再履修対応 Ins ST%>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
					<%' 2022.03.04 再履修対応 Ins ED%>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>>-</td><%' 2022.03.08 再履修対応 Ins%>
					<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>>-</td>
					<%' 2022.03.08 再履修対応 Ins ST%>
					<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>>-</td>
					<%' 2022.03.08 再履修対応 Ins ED%>
				<%Else%>

					<input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>">
					<input type="hidden" name="txtKaisetu<%=i%>" value="<%=m_Rs("KAISETU")%>">
					<input type="hidden" name="hidUpdFlg<%=i%>" value="False">
					<td class="<%=w_cell%>" align="center" width="65"  nowrap <%=w_Padding%>><%=w_sGakusekiCd%></td>
					<td class="<%=w_cell%>" align="left"   width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="100" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="80"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="85"  nowrap <%=w_Padding%>>-</td>
				<%End if%>
			<%
			'=========================================================================
			'//科目が必修か、または選択科目の時に生徒が科目を選択している場合(入力可)
			'=========================================================================
			Else
				%>
					<td class="<%=w_cell%>" align="center" width="65" nowrap <%=w_Padding%>>
						<%=w_sGakusekiCd%>
						<input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>">
						<input type="hidden" name="txtKaisetu<%=i%>" value="<%=m_Rs("KAISETU")%>">
					</td>
						
					<input type="hidden" name="hidUpdFlg<%=i%>" value="True">
					<td class="<%=w_cell%>" align="left"  width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
					
					<!-- 2002.03.20 -->
					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI1"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI2"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI3"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI4"))%></td>
					<%Else%>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>&nbsp;&nbsp;</td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>&nbsp;&nbsp;</td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>&nbsp;&nbsp;</td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>&nbsp;&nbsp;</td>
					<%End If%>
					
					<%
					'//NN対応
					If session("browser") = "IE" Then
						w_sInputClass = "class='num'"
						w_sInputClass1 = "class='num'"
						w_sInputClass2 = "class='num'"
						w_sInputClass3 = "class='num'" '2022.03.16INS
					Else
						w_sInputClass = ""
						w_sInputClass1 = ""
						w_sInputClass2 = ""
						w_sInputClass3 = "" '2022.03.16INS
					End If
					
					if m_iKikan = "NO" or (m_sSikenKBN = C_SIKEN_KOU_KIM and w_bZenkiOnly = true) Then '2023.12.21 Yoshida　ADD 前期開設科目の場合の処理を追加
						w_sInputClass1 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
					End if
					
					'// 欠課入力可能ﾌﾗｸﾞ
					if Not m_bKekkaNyuryokuFlg or (m_sSikenKBN = C_SIKEN_KOU_KIM and w_bZenkiOnly = true) Then '2023.12.21 Yoshida　ADD 前期開設科目の場合の処理を追加
						w_sInputClass2 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
					End if
					w_sInputClass3 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'" '2022.03.16INS
					'=========================================================================
					'//通常授業の場合 
					'=========================================================================
					%>
				<%If m_iKikan <> "NO" Then%>
					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
						<%' 2022.03.04 再履修対応 UPD ST%>
						<%' 免除フラグが立っていれば、文字として表示--%>
						<% If CInt(gf_SetNull2Zero(m_Rs("Menjo"))) = 1 Then %>
							<td class="<%=w_cell%>" width="50"align="center" nowrap <%=w_Padding%>><font size="2"><%=w_sSeiseki%></font></td>
						<%Else%>
							<%' 2022.10.14 高専間単位互換対応 UPD-->%>
							<%' 点数(成績)入力と文字入力で処理を分岐--%>
							<% If m_FirstSeisekiInp = C_SEISEKI_INP_TYPE_NUM Then%>
								<td class="<%=w_cell%>" width="50"align="center" nowrap <%=w_Padding%>>
									<input type="text" <%=w_sInputClass1%> name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)" onChange="f_GetTotalAvg()" <%=w_Disabled2%>>
								</td>
							<% Else%>
								<td class="<%=w_cell%>" width="50"align="center" nowrap <%=w_Padding%>><font size="2"><%=w_sSeiseki%></font></td>
							<% End If%>
							<%' 2022.10.14 高専間単位互換対応 UPD<--%>
						<%End If%>
						<%' 2022.03.04 再履修対応 UPD ED%>

						<%' 2022.03.04 再履修対応 Ins ST%>
						<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(m_Rs("ZenNendoSeiseki"))%></td>
						<%' 2022.03.04 再履修対応 Ins ED%>
						<%' 2022.10.14 高専間単位互換対応 UPD-->%>
						<%' 点数(成績)入力と文字入力で処理を分岐%>
						<% If m_FirstSeisekiInp = C_SEISEKI_INP_TYPE_NUM Then%>
							<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>	
									<td class="<%=w_cell%>"  width="50" align="center" nowrap <%=w_Padding%>>
										<input type="button" size="2" name="button<%=i%>" value="<%=w_sHyoka%>" onClick="return f_change(<%=i%>)" class="<%=w_cell%>" style="text-align:center">
									</td>
									<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>">
							<%Else%>
									<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>><%=w_sHyoka%></td>
									<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>">
							<%End If%>
						<% Else%>
							<%' 成績が点数表示でない場合%>
							<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>>
								<select name="Hyoka<%=i%>" style='width:50px;' onchange="">
									<option value="@@@" selected >  </option>
									<%Do Until m_Rs_Hyoka.EOF%>
										<option value='<%=m_Rs_Hyoka("M08_HYOKA_SYOBUNRUI_MEI")%>'  <%=f_Selected(cstr(m_Rs_Hyoka("M08_HYOKA_SYOBUNRUI_MEI")),cstr(m_Rs("HYOKA")))%>><%=m_Rs_Hyoka("M08_HYOKA_SYOBUNRUI_MEI")%></option>									
										<%m_Rs_Hyoka.MoveNext%>
									<%Loop%>
									<%m_Rs_Hyoka.MoveFirst%>
							</td>
						<% End If%>
						<%' 2022.10.14 高専間単位互換対応 UPD<--%>	

							<td class="<%=w_cell%>" width="55" align="center" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="55" align="right"  nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sChikaisu)%></td>
							<%' 2022.03.08 再履修対応 Ins ST%>
							<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass3%> name=ZenNendoCHIKAI value="<%=(m_Rs("ZenNendoCHIKAI"))%>" size=2 maxlength=3 ></td>
							<%' 2022.03.08 再履修対応 Ins ED%>
							<td class="<%=w_cell%>" width="55" align="center" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="55" align="center" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="55" align="right"  nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sKekkasu)%></td>
							<%' 2022.03.08 再履修対応 Ins ST%>
							<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass3%>  name=ZenNendoKEKA value="<%=(m_Rs("ZenNendoKEKA"))%>" size=2 maxlength=3 ></td>
							<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass3%>  name=ZenNendoKEKA_NASI value="<%=(m_Rs("ZenNendoKEKA_NASI"))%>" size=2 maxlength=3 ></td>
							<%' 2022.03.08 再履修対応 Ins ED%>
					<%Else%>
							<td class="<%=w_cell%>" width="50"  nowrap align="center" <%=w_Padding%>>-</td>
							<!--<td class="<%=w_cell%>" width="50"  nowrap align="center" <%=w_Padding%>>-</td>	<%'-- 2022.03.04 再履修対応 Ins-- %>　2022.06.22 Del 特別活動の場合前年成績不要-->
							<td class="<%=w_cell%>" width="50"  nowrap align="center" <%=w_Padding%>>-</td>
							<%If m_sKamokuCd = C_TOKU_SEISO Or m_sKamokuCd = C_TOKU_SHR Then    '清掃とSHRは遅刻をハイフン表示 2024.06.19 Add Kiyomoto%>
								<td class="<%=w_cell%>" width="100"  nowrap align="center" <%=w_Padding%>>-</td>
							<%Else%>
								<td class="<%=w_cell%>" width="100" nowrap align="center" <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
							<%End If%>
							<td class="<%=w_cell%>" width="80"  nowrap align="center" <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="85"  nowrap align="center" <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
					<%End If%>
					
				<%Else%>
					
					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
						<%' 2022.03.04 再履修対応 UPD ST%>
						<%' 免除フラグが立っていれば、文字として表示--%>
						<% If CInt(gf_SetNull2Zero(m_Rs("Menjo"))) = 1 Then %>
							<td class="<%=w_cell%>" width="50"align="right" nowrap <%=w_Padding%>><font size="2"><%=w_sSeiseki%></font></td>
						<%Else%>
							<td class="<%=w_cell%>" width="50" align="right" nowrap <%=w_Padding%>>
								<input type="text" <%= w_sInputClass1 %> name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)" onChange="f_GetTotalAvg()" <%=w_Disabled2%>>
							</td>
						<%End If%>
						<%' 2022.03.04 再履修対応 UPD ED%>

						<%	'表示のみの場合の合計・平均値を求める
							If IsNull(w_sSeiseki) = False Then
								If IsNumeric(CStr(w_sSeiseki)) = True Then
									w_lSeiTotal = w_lSeiTotal + CLng(w_sSeiseki)
									w_lGakTotal = w_lGakTotal + 1
								End If
							End If
						%>
						<%' 2022.03.04 再履修対応 Ins ST%>
						<td class="<%=w_cell%>" align="right" width="50" nowrap <%=w_Padding%>><font size="2"><%=gf_HTMLTableSTR(m_Rs("ZenNendoSeiseki"))%></font></td>
						<%' 2022.03.04 再履修対応 Ins ED%>
						<%' 2022.10.11 高専間単位互換対応 UPD-->%>
						<% If m_FirstSeisekiInp = C_SEISEKI_INP_TYPE_NUM Then%>
							<%' 点数(成績)入力と文字入力で処理を分岐%>
							<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>
									<td class="<%=w_cell%>"  width="50" align="center" nowrap <%=w_Padding%>><%=trim(w_sHyoka)%></td>
							<%Else%>
									<td class="<%=w_cell%>"  width="50" align="center" nowrap <%=w_Padding%>><%=trim(w_sHyoka)%></td>
							<%End If%>
						<%Else%>
							<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(m_Rs("HYOKA"))%></td>
						<%End If%>
						<%' 2022.10.11 高専間単位互換対応 UPD<--%>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sChikaisu)%></td>
						<%' 2022.03.08 再履修対応 Ins ST%>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%> ><input type="text" <%=w_sInputClass3%> name=ZenNendoCHIKAI value="<%=(m_Rs("ZenNendoCHIKAI"))%>" size=2 maxlength=3 ></td>
						<%' 2022.03.08 再履修対応 Ins ED%>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sKekkasu)%></td>
						<%' 2022.03.08 再履修対応 Ins ST%>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass3%>  name=ZenNendoKEKA value="<%=(m_Rs("ZenNendoKEKA"))%>" size=2 maxlength=3 ></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass3%>  name=ZenNendoKEKA_NASI value="<%=(m_Rs("ZenNendoKEKA_NASI"))%>" size=2 maxlength=3 ></td>
						<%' 2022.03.08 再履修対応 Ins ED%>
					<%Else%>
						<td class="<%=w_cell%>" width="50"  align="center" nowrap  <%=w_Padding%>>-</td>
						<td class="<%=w_cell%>" width="50"  align="center" nowrap  <%=w_Padding%>>-</td>
						<%If m_sKamokuCd = C_TOKU_SEISO Or m_sKamokuCd = C_TOKU_SHR Then    '清掃とSHRは遅刻をハイフン表示 2024.06.19 Add Kiyomoto%>
							<td class="<%=w_cell%>" width="100"  nowrap align="center" <%=w_Padding%>>-</td>
						<%Else%>
							<td class="<%=w_cell%>" width="100" align="center" nowrap  <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<%End If%>
						<td class="<%=w_cell%>" width="80"  align="center" nowrap  <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="85"  align="center" nowrap  <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
					<%End If%>
				<%End If%>
			<%End If%>
				
				<% if m_SchoolFlg and m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
					
					<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>>
						<% if w_DataKbn = C_HYOKA_FUNO or w_DataKbn = C_MIHYOKA or w_DataKbn = 0 then %>
							<input type="checkbox" name="chkHyokaFuno<%=i%>" value="3" <%=w_Disabled%> <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);">
						<% else %>
							&nbsp;
							<input type="hidden" name="chkHyokaFuno<%=i%>" value="<%=w_DataKbn%>">
						<% end if %>
					</td>
					
				<% end if %>
			</tr>
			
			<%
				m_Rs.MoveNext
				i = i + 1
			Loop
			%>
			<tr>
				<td class="header" nowrap align="right" colspan="7">
					<FONT COLOR="#FFFFFF"><B>成績合計</B></FONT>
					<input type="text" name="txtTotal" size="5" <%=w_sInputClass%> readonly>
				</td>
				<td class="header" nowrap align="center" colspan="6">&nbsp;</td>
			</tr>
			
			<tr>
				<td class="header" nowrap align="right" colspan="7">
					<FONT COLOR="#FFFFFF"><B>　平均点</B></FONT>
					<input type="text" name="txtAvg" size="5" <%=w_sInputClass%> readonly>
				</td>
				<td class="header" nowrap align="center" colspan="6">&nbsp;</td>
			</tr>
		</table>
		
		</td>
		</tr>
		
		<tr>
		<td align="center">
		<table>
			<tr>
				<td align="center" align="center" colspan="13">
					<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%>
						<input type="button" class="button" value="　登　録　" onclick="javascript:f_Touroku()">
					<%End If%>
						<input type="button" class="button" value="キャンセル" onclick="javascript:f_Cansel()">
				</td>
			</tr>
		</table>
		</td>
		</tr>
	</table>
	
	<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="KamokuCd"    value="<%=m_sKamokuCd%>">
	<input type="hidden" name="i_Max"       value="<%=i%>">
	<input type="hidden" name="txtSikenKBN" value="<%=m_sSikenKBN%>">
	<input type="hidden" name="txtGakuNo"   value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakkaCd"  value="<%=m_sGakkaCd%>">
	<input type="hidden" name="txtClassNo"  value="<%=m_sClassNo%>">
	<input type="hidden" name="txtKamokuCd" value="<%=m_sKamokuCd%>">
	<input type="hidden" name="txtKamokuNM" value="<%=m_sKamokuNM%>">
	<input type="hidden" name="txtTUKU_FLG" value="<%=m_TUKU_FLG%>">
	<input type="hidden" name="PasteType"   value="">
	<!-- 02/03/27 追加 -->
	<input type="hidden" name="hidSouJyugyou">
	<input type="hidden" name="hidJunJyugyou">
	<input type="hidden" name="hidUpdMode">
	
	<input type="hidden" name="hidFirstGakusekiNo" value="<%=m_FirstGakusekiNo%>">
	<input type="hidden" name="hidMihyoka" value ="<%=w_DataKbn%>">
	<input type="hidden" name="hidSchoolFlg" value ="<%=m_SchoolFlg%>">
	<!-- 2022/03/17 追加 -->
	<input type="hidden" name="scrollbarWidth" value="<%=scrollbarWidth%>">
	<!-- 2022/10/11 追加 -->
	<input type="hidden" name="hidHyoka" value="">
	<input type="hidden" name="hidSeisekiInp" value="<%=m_FirstSeisekiInp%>">
	<input type="hidden" name="txtZokuseiCd" value="<%=m_sZokuseiCd%>">

	<!-- ADD START 2023/12/14 吉田 WEBアクセスログカスタマイズ -->
	<input type="hidden" name="LOG_TAISYO" value="<%=m_sTaisyo%>">
 	<input type="hidden" name="LOG_SOSA" value="<%=m_sSosa%>">
  	<!-- ADD END 2023/12/14 吉田 WEBアクセスログカスタマイズ -->
	</FORM>
	</center>
	</body>
	<SCRIPT>
	<!--
		//2002/02/05 佐野 追加
		//************************************************************
		//	[機能]	成績が変更されたとき
		//	[引数]	なし
		//	[戻値]	なし
		//	[説明]	成績の合計と平均を求める
		//	[備考]	学生の総数が分かるのは最後であるため、この位置に書く。
		//************************************************************
		function f_GetTotalAvg(){
			var i;
			var total;
			var avg;
			var cnt;
			
			total = 0;
			cnt = 0;
			avg = 0;
			
			<%If m_iKikan <> "NO" Then	'入力期間中%>
				//学生数でのループ
				for(i=0;i<<%=i%>;i++) {
					//存在するかどうか
					textbox = eval("document.frm.Seiseki" + (i+1));
					if (textbox) {
						//未入力チェック
						if (textbox.value != "") {
							//数字でないのは無視する
							if (!isNaN(textbox.value)) {
								total = total + parseInt(textbox.value);
							}
						}
						cnt = cnt + 1;
					}
				}
			
			<% Else	'入力期間中ではない%>
				total = <%=w_lSeiTotal%>;
				cnt   = <%=w_lGakTotal%>;
			<% End If%>
			
			document.frm.txtTotal.value=total;
			
			//四捨五入
			if (cnt!=0){
				avg = total/cnt;
				avg = avg * 10;
				avg = Math.round(avg);
				avg = avg / 10;
			}
			
			document.frm.txtAvg.value=avg;
		}
	//-->
	</SCRIPT>

	</html>
<%
End sub

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
	//-->
	</SCRIPT>
	</head>
	
    <body LANGUAGE="javascript">
	<form name="frm" method="post">
	</head>
	
	<body>
	<br><br><br>
	<center>
		<span class="msg">個人履修データが存在しません。</span>
	</center>
	
	<input type="hidden" name="txtMsg" value="個人履修データが存在しません。">
	
	</form>
	</body>
	</html>

<%
End Sub
%>