<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0150_bottom.asp
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
' 作      成: 2002/06/21 shin
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<!--#include file = "sei0150_bottom_tujo.asp"-->
<!--#include file = "sei0150_bottom_toku.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '氏名選択用のWhere条件
    Public m_iNendo			'年度
    Public m_sKyokanCd		'教官コード
    Public m_sSikenKBN		'試験区分
    Public m_sGakuNo		'学年
    Public m_sClassNo		'学科
    Public m_sKamokuCd		'科目コード
    Public m_sSikenNm		'試験名
    Public m_rCnt			'レコードカウント
    Public m_sGakkaCd
    Public m_iSyubetu		'出欠値集計方法
    Public m_iJigenTani		'//１時限の単位数
    
    Public m_iKamoku_Kbn
    Public m_iHissen_Kbn
	Public m_ilevelFlg
	Public m_Rs
	Public m_DRs
	Public m_SRs
	
	Dim m_iSouJyugyou		'総授業時間
	DIm m_iJunJyugyou		'純授業時間
	
	Public m_iKikan			'入力期間フラグ
	Public m_bKekkaNyuryokuFlg		'欠課入力可能ﾌﾗｸﾞ(True:入力可 / False:入力不可)
	
	Public m_iShikenInsertType
	Public m_FirstGakusekiNo
	
	m_iShikenInsertType = 0
	
	Public m_sSyubetu
	
	'2002/06/21
	Dim m_iKamokuKbn		'科目区分(0:通常授業、1:特別科目)
	Dim m_sKamokuBunrui		'科目分類(01:通常授業、02:認定科目、03:特別科目)
	
	Dim m_AryKamokuHyoka()	'科目評価セット配列
		'm_AryKamokuHyoka(0)　'評価,
	    'm_AryKamokuHyoka(1)　'評定,
	    'm_AryKamokuHyoka(2)　'欠点科目がセットされる
		
	Dim m_iDataCount
	Dim m_AryHyokaData()
	Dim m_iSeisekiInpType
	
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
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
	    '// ﾊﾟﾗﾒｰﾀSET
	    Call s_SetParam()
		
		'//成績入力方法の取得(0:点数[C_SEISEKI_INP_TYPE_NUM]、1:文字[C_SEISEKI_INP_TYPE_STRING])
		if not gf_GetKamokuSeisekiInp(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iSeisekiInpType) then 
			m_bErrFlg = True
			Exit Do
		end if
		
		'//科目評価取得
		'if not gf_GetKamokuTensuHyoka(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,60,m_AryKamokuHyoka) then 
		'	m_bErrFlg = True
		'	Exit Do
		'end if
		
		'//期間データの取得
        If f_Nyuryokudate() = 1 Then
			m_iKikan = "NO"
		else
			m_iKikan = ""
		End If
		
		'//出欠欠課の取り方を取得
		'//科目区分(0:試験毎,1:累積)
		If gf_GetKanriInfo(m_iNendo,m_iSyubetu) <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		
	    '**********************************************************
	    '通常授業と特別活動で、とって来る場所が変わる。
	    '**********************************************************
		If m_iKamokuKbn = C_JIK_JUGYO then  '通常授業の場合
			'//科目情報を取得
			'//科目区分(0:一般科目,1:専門科目)、及び、必修選択区分(1:必修,2:選択)を調べる
			'//レベル別区分(0:一般科目,1:レベル別科目)を調べる
			If f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) <> 0 Then 
				m_bErrFlg = True
				Exit Do
			End If
			
			'//成績、学生データ取得
			'//科目区分がC_KAMOKU_SENMON(0:一般科目)の場合はクラス別に生徒を表示
			'//科目区分がC_KAMOKU_SENMON(1:専門科目)の場合は学科別に生徒を表示
			If f_getdate(m_iKamoku_Kbn) <> 0 Then m_bErrFlg = True : Exit Do
			
			If m_rs.EOF Then
				Call gs_showWhitePage("個人履修データが存在しません。","成績登録")
				Exit Do
			End If
			
			'//欠課数の取得
			If f_GetSyukketu() <> 0 Then m_bErrFlg = True : Exit Do
			
			Call showPage_Tujo()
		Else
			'//成績、学生データ取得
			If f_getTUKUclass(m_iNendo,m_sKamokuCd,m_sGakuNo,m_sClassNo) <> 0 Then m_bErrFlg = True : Exit Do
			
			If m_rs.EOF Then
				Call gs_showWhitePage("個人履修データが存在しません。","成績登録")
				Exit Do
			End If
			
			Call showPage_Toku()
	    End If
		
		'// ページを表示
		'Call showPage()
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
	
	m_iNendo	 = request("txtNendo")
	m_sKyokanCd	 = request("txtKyokanCd")
	m_sSikenKBN	 = Cint(request("sltShikenKbn"))
	m_sGakuNo	 = Cint(request("txtGakuNo"))
	m_sClassNo	 = Cint(request("txtClassNo"))
	m_sKamokuCd	 = request("txtKamokuCd")
	m_sGakkaCd	 = request("txtGakkaCd")
	m_iJigenTani = Session("JIKAN_TANI") '１時限の単位数
	m_sSyubetu	 = trim(Request("SYUBETU"))
	
	m_iKamokuKbn = cint(Request("hidKamokuKbn"))
	
	if m_iKamokuKbn = C_JIK_JUGYO then
		'通常科目
		m_sKamokuBunrui = C_KAMOKUBUNRUI_TUJYO
	else
		'特別科目
		m_sKamokuBunrui = C_KAMOKUBUNRUI_TOKUBETU
	end if
	
End Sub

'********************************************************************************
'*  [機能]  欠課数、遅刻数を取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSyukketu()
	Dim w_sTKyokanCd
	
    On Error Resume Next
    Err.Clear
	
    f_GetSyukketu = 1
	
	Do
		'//科目担当教官の教官CDの取得
		If f_GetTantoKyokan2(w_sTKyokanCd) <> 0 Then m_bErrFlg = True : Exit Do
		
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
    
    On Error Resume Next
    Err.Clear
    
    f_SikenInfo = 1
	p_bZenkiOnly = false
	
    Do 
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
		
		If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_SikenInfo = 99
			Exit Do
		End If
		
		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_bZenkiOnly = True
		End If
		
        f_SikenInfo = 0
        Exit Do
    Loop
	
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
		
        If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
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

'********************************************************************************
'*	[機能]	データの取得
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
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI,A.T16_KEKA_KIMATU_Z AS KEKA,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T16_CHIKAI_KIMATU_Z AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_KIMATU_Z as SOUJI, A.T16_JUNJIKAN_KIMATU_Z as JYUNJI, "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI,A.T16_KEKA_TYUKAN_K AS KEKA,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_K AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_TYUKAN_K as SOUJI, A.T16_JUNJIKAN_TYUKAN_K as JYUNJI, "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_Z AS SEI_ZT,A.T16_KEKA_TYUKAN_Z AS KEKA_ZT,A.T16_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,A.T16_CHIKAI_TYUKAN_Z AS CHIKAI_ZT,A.T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI_ZT, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI_ZK,A.T16_KEKA_KIMATU_Z AS KEKA_ZK,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,A.T16_CHIKAI_KIMATU_Z AS CHIKAI_ZK,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI_ZK, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI_KT,A.T16_KEKA_TYUKAN_K AS KEKA_KT,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,A.T16_CHIKAI_TYUKAN_K AS CHIKAI_KT,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI_KT, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI_KK,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_KIMATU_K as SOUJI, A.T16_JUNJIKAN_KIMATU_K as JYUNJI, A.T16_SAITEI_JIKAN, A.T16_KYUSAITEI_JIKAN, "
		End Select

		w_sSQL = w_sSQL & " 	A.T16_GAKUSEI_NO AS GAKUSEI_NO,A.T16_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_SELECT_FLG "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_LEVEL_KYOUKAN "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_OKIKAE_FLG "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T16_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_NENDO = C.T13_NENDO "
		
		'//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
		w_sSQL = w_sSQL & " AND	A.T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
		w_sSQL = w_sSQL & " ORDER BY A.T16_GAKUSEKI_NO "
		
		If gf_GetRecordset(m_Rs, w_sSQL) <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getdate = 99
			m_bErrFlg = True
			Exit Do 
		End If
		
		m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
		m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))
		
		'//ﾚｺｰﾄﾞカウント取得
		m_rCnt = gf_GetRsCount(m_Rs)
		
		f_getdate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	特別活動受講学生取得
'********************************************************************************
Function f_getTUKUclass(p_iNendo,p_sKamokuCd,p_iGakunen,p_iClass)
	Dim w_sSQL
	Dim w_Rs
	
	On Error Resume Next
	Err.Clear
	
	f_getTUKUclass = 1
	p_sTKyokanCd = ""
	
	Do
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
		
		If gf_GetRecordset(m_Rs, w_sSQL) <> 0 Then
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
'********************************************************************************
Function f_GetTantoKyokan(p_sTKyokanCd)
	Dim w_sSQL
    Dim w_Rs
    
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
		
        If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
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
		
        If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetTantoKyokan = 99
            Exit Do
        End If
		
		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then p_sTKyokanCd = w_Rs("T27_KYOKAN_CD")
		
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
		
		If gf_GetRecordset(m_DRs, w_sSQL) <> 0 Then
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
'*	[機能]	データの取得
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

'********************************************************************************
'*  [機能]  HTMLを出力
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
	
	Dim i
	
	Dim w_lSeiTotal	'成績合計
	Dim w_lGakTotal	'学生人数
	
	Dim w_SSSS
	Dim w_SSSR
	Dim w_Date
	
	w_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")
	
	'データがNULLの場合に0に変換しないために、一旦データを保存するワークで使用
	Dim w_sData
	Dim w_sData2
	
	Dim w_Padding
	Dim w_Padding2
	
	w_Padding = "style='padding:2px 0px;'"
	w_Padding2 = "style='padding:2px 0px;font-size:10px;'"
	
	w_lSeiTotal = 0
	w_lGakTotal = 0
	
	'//NN対応
	If session("browser") = "IE" Then
		w_sInputClass  = "class='num'"
		w_sInputClass1 = "class='num'"
		w_sInputClass2 = "class='num'"
	Else
		w_sInputClass = ""
		w_sInputClass1 = ""
		w_sInputClass2 = ""
	End If
	
	if m_iKikan = "NO" Then
		w_sInputClass1 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
	End if
	
	'// 欠課入力可能ﾌﾗｸﾞ
	if Not m_bKekkaNyuryokuFlg then
		w_sInputClass2 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
	End if
	
	i = 1
	
%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--
	//************************************************************
    //  [機能]  ページロード時処理
    //************************************************************
    function window_onload() {
		//スクロール同期制御
		parent.init();
		
		//成績合計値の取得
		f_GetTotalAvg();
		
		//総時間と純時間をhiddenにセット
		document.frm.hidSouJyugyou.value = "<%= m_iSouJyugyou %>";
		document.frm.hidJunJyugyou.value = "<%= m_iJunJyugyou %>";
		
        document.frm.target = "topFrame";
        document.frm.action = "sei0150_middle.asp"
        document.frm.submit();
        return;
    }
	
	//************************************************************
    //  [機能]  評価ボタンが押されたとき
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
    //************************************************************
    function f_Touroku(){
		if(f_CheckData_All() == 1){
	            alert("入力値が不正です");
	            return 1;
		}else{
			if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return;}
			
			document.frm.hidSouJyugyou.value = parent.topFrame.document.frm.txtSouJyugyou.value;
			document.frm.hidJunJyugyou.value = parent.topFrame.document.frm.txtJunJyugyou.value;
			
			//ヘッダ部空白表示
			parent.topFrame.document.location.href="white.asp"
			
			//登録処理
		<% if m_iKamokuKbn = C_JIK_JUGYO then %>
			document.frm.hidUpdMode.value = "TUJO";
			document.frm.action="sei0150_upd.asp";
		<% Else %>
			document.frm.hidUpdMode.value = "TOKU";
			document.frm.action="sei0150_upd_toku.asp";
		<% End if %>
	        document.frm.target="main";
	        document.frm.submit();
		}
	}
	
	//************************************************************
	//	[機能]	キャンセルボタンが押されたとき
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
		
		<% if m_iKamokuKbn = C_JIK_JUGYO then %>
		
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
					}
				}
			}
		}
		if (w_bFLG == false){
			return 1;
		}
		
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
				}
			}
		}
		
		if (w_bFLG == false){ return 1; }
		
		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			
			w_Kekka = eval("document.frm.Kekka"+i);
			w_bFLG = true
			
			if (w_Kekka){
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
					}
					
					//小数点チェック
					w_decimal = new Array();
					w_decimal = wStr.split(".")
					if(w_decimal.length>1){
						w_bFLG = false;
						w_Kekka.focus();
						return 1;
						break;
					}
				}
			}
		}
		
		if (w_bFLG == false){ return 1; }
		
		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_KekkaGai = eval("document.frm.KekkaGai"+i);
			w_bFLG = true
			
			if (w_KekkaGai){
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
					}
					
					//小数点チェック
					w_decimal = new Array();
					w_decimal = wStr.split(".")
					if(w_decimal.length>1){
						w_bFLG = false;
						w_KekkaGai.focus();
						return 1;
						break;
					}
				}
			}
		}
		if (w_bFLG == false){ return 1; }
		
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
			}
			
			//小数点チェック
			w_decimal = new Array();
			w_decimal = wStr.split(".")
			
			if(w_decimal.length>1){
				wFromName.focus();
				return false;
			}
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
			if(wObj1.value < wObj2.value){
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
			return false;
		}
		return true;
	}
	
	//-->
	</SCRIPT>
	</head>
	<body LANGUAGE="javascript" onload="return window_onload()">
	<center>
	<form name="frm" method="post" onClick="return false;">
	
	<table width="710">
	<tr>
	<td>
	
	<table class="hyo" align="center" width="710" border="1">
	<%	
		m_Rs.MoveFirst
		
		Do Until m_Rs.EOF
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
		Call f_SikenInfo(w_bZenkiOnly)
		
		'学年末試験の場合のみ
		If m_sSikenKBN = C_SIKEN_KOU_KIM Then
			
			'前期開設だったら前期期末の欠課を学年末の成績にセットする
			If w_bZenkiOnly = True Then
				'学期末成績が0
				If gf_SetNull2String(m_Rs("KEKA")) = "" Then 
					w_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))			'欠課数
					w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK"))	'欠課対象外
					w_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))		'遅刻回数
				End If
			End If
		End If
		
	'//前期で終わっている科目の欠課を取得して学期末成績にセットする。2002/02/21 ITO
	'//値の初期化。
	w_bNoChange = False
	w_sKekkasu = ""
	w_sChikaisu = ""
	
	'---------------------------------------------------------------------------------------------
	'通常授業ときの処理
	if m_iKamokuKbn = C_JIK_JUGYO then 
		w_sSeiseki = gf_SetNull2String(m_Rs("SEI"))
		w_sHyoka = gf_HTMLTableSTR(m_Rs("HYOKAYOTEI"))
		
		'前期で終わっている科目の欠課を取得して学期末成績にセットする。2002/02/21 ITO
		
		'学年末試験の場合のみ
		If m_sSikenKBN = C_SIKEN_KOU_KIM Then
			
			'前期開設だったら前期期末の欠課を学年末の成績にセットする
			If w_bZenkiOnly = True Then
				'学期末成績が0
				If gf_SetNull2String(m_Rs("SEI")) = "" Then 
					w_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))			'前期期末成績
				End If
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
	'//C_IDO_FUKUGAKU=3:復学、C_IDO_TEI_KAIJO=5:停学解除
	w_SSSS = ""
	w_SSSR = ""
	
	w_SSSS = gf_Get_IdouChk(w_sGakusekiCd,w_Date,m_iNendo,w_SSSR)
	
	if CStr(w_SSSS) <> "" and Cstr(w_SSSS) <> CStr(C_IDO_FUKUGAKU) and Cstr(w_SSSS) <> Cstr(C_IDO_TEI_KAIJO) Then
		w_SSSS = "[" & w_SSSR & "]"
		w_bNoChange = True
	else
		w_SSSS = ""
	end if
	
	'通常授業
	if Cstr(m_iKamokuKbn) = Cstr(C_JIK_JUGYO) then 
		'//欠課遅刻数の取得
		'//欠課数×単位数の取得
		w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_KEKKA)		'戻り値がNULLの時は""
		
		'gf_IIFに渡すときにパラメータを計算するので、パラメータは0に変換
		w_sKekkasu = gf_IIF(w_sData = "", "", cint(gf_SetNull2Zero(w_sData)) * cint(m_iJigenTani))
		
		'//１欠課の場合の欠課数の取得
		w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_KEKKA_1)
		
		if w_sKekkasu = "" and w_sData = "" then
			w_sKekkasu = ""
		else
			'どちらか一方""でなければ計算
			w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_sData))			'//１欠課の場合の欠課数の取得
		end if
		
		'//遅刻数の取得
		w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_TIKOKU)
		w_sChikaisu = gf_IIF(w_sData = "", "", cint(gf_SetNull2Zero(w_sData)))
		
		'//早退数の取得
		w_sData = f_Syukketu2New(w_sGakusekiCd,C_KETU_SOTAI)
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
			'call f_GetKekaChi(m_iNendo,w_sShikenKBN_RUI,m_sKamokuCd,cstr(m_Rs("GAKUSEI_NO")),w_iKekka_rui,w_iChikoku_rui,w_iKekkaGai_rui)
			call f_GetKekaChi(m_iNendo,m_iShikenInsertType,m_sKamokuCd,cstr(m_Rs("GAKUSEI_NO")),w_iKekka_rui,w_iChikoku_rui,w_iKekkaGai_rui)
			
			'どちらも""の時は""
			if w_sKekkasu = "" and w_iKekka_rui = "" then
				w_sKekkasu = ""
			else
				w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))
			end if
			
			'どちらも""の時は""
			if w_sChikaisu = "" and w_iChikoku_rui = "" then
				w_sChikaisu = ""
			else
				w_sChikaisu = cint(gf_SetNull2Zero(w_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
			end if
		end if
		
		If cint(gf_SetNull2Zero(w_sKekka)) = 0 and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then 		'//欠入が0で,欠計が0より大きい場合
			w_sKekka = cint(gf_SetNull2Zero(w_sKekkasu))								'//欠入＝欠計
		End If
		
		If cint(gf_SetNull2Zero(w_sChikai)) = 0 AND cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then		'//遅入が0で,遅計が0より大きい場合
			w_sChikai = cint(gf_SetNull2Zero(w_sChikaisu))							'//遅入＝遅計
		End If
			
			'========================================================================================
			'//科目が選択科目の時に科目を選択していない場合(入力不可)
			'========================================================================================
			If w_bNoChange = True Then
				
				if Cstr(m_iKamokuKbn) = Cstr(C_JIK_JUGYO) Then%>
					<input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>">
					<input type="hidden" name="hidUpdFlg<%=i%>" value="False">
					<td class="<%=w_cell%>" width="65" nowrap ><%=w_sGakusekiCd%></td>
					<td class="<%=w_cell%>" align="left" width="150" nowrap  <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
					<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
				<%Else%>
					<input type="hidden" name=txtGseiNo<%=i%> value="<%=m_Rs("GAKUSEI_NO")%>">
					<input type="hidden" name="hidUpdFlg<%=i%>" value="False">
					<td class="<%=w_cell%>" width="65"  <%=w_Padding%>><%=w_sGakusekiCd%></td>
					<td class="<%=w_cell%>" align="left" width="150"   nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
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
						<td class="<%=w_cell%>"  width="65" nowrap <%=w_Padding%>><%=w_sGakusekiCd%><input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>"></td>
						<input type="hidden" name="hidUpdFlg<%=i%>" value="True">
						<td class="<%=w_cell%>" align="left"  width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
					
					<%If m_iKamokuKbn = C_JIK_JUGYO Then%>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI1"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI2"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI3"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI4"))%></td>
					<%Else%>
						<td class="<%=w_cell%>" width="30" nowrap <%=w_Padding%>>&nbsp;</td>
						<td class="<%=w_cell%>" width="30" nowrap <%=w_Padding%>>&nbsp;</td>
						<td class="<%=w_cell%>" width="30" nowrap <%=w_Padding%>>&nbsp;</td>
						<td class="<%=w_cell%>" width="30" nowrap <%=w_Padding%>>&nbsp;</td>
					<%End If%>
				
				<%If m_iKikan <> "NO" Then%>
					<% If m_iKamokuKbn = C_JIK_JUGYO Then '//通常授業の場合 %>
						
						<td class="<%=w_cell%>" width="50"align="center" nowrap <%=w_Padding%>><input type="text" <%= w_sInputClass1 %>  name=Seiseki<%=i%> value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)" onChange="f_GetTotalAvg()"></td>
						
						<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>
							<td class="<%=w_cell%>"  width="50" align="center" nowrap <%=w_Padding%>>
								<input type="button" size="2" name="button<%=i%>" value="<%=w_sHyoka%>" onClick="return f_change(<%=i%>)" class="<%=w_cell%>" style="text-align:center">
							</td>
							<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>">
						<%Else%>
							<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>><%=w_sHyoka%></td>
							<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>">
						<%End If%>
							
							<td class="<%=w_cell%>" width="55" align="center" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="55" align="right"  nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sChikaisu)%></td>
							<td class="<%=w_cell%>" width="55" align="center" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="55" align="center" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="55" align="right"  nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sKekkasu)%></td>
					<%Else%>
							
							<td class="<%=w_cell%>" width="50" nowrap align="center" <%=w_Padding%>>-</td>
							<td class="<%=w_cell%>" width="50" nowrap align="center" <%=w_Padding%>>-</td>
							<td class="<%=w_cell%>" width="100" nowrap align="center"<%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="80" nowrap align="center" <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="85" nowrap align="center" <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
					<%End If%>
				<%Else%>
					<%If m_iKamokuKbn = C_JIK_JUGYO Then%>
						<td class="<%=w_cell%>" width="50" align="right" nowrap <%=w_Padding%>><input type="text" <%= w_sInputClass1 %>  name=Seiseki<%=i%> value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)" onChange="f_GetTotalAvg()"></td>
						<%	'表示のみの場合の合計・平均値を求める
							If IsNull(w_sSeiseki) = False Then
								If IsNumeric(CStr(w_sSeiseki)) = True Then
									w_lSeiTotal = w_lSeiTotal + CLng(w_sSeiseki)
									w_lGakTotal = w_lGakTotal + 1
								End If
							End If
						%>
						
						<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>><%=trim(w_sHyoka)%></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sChikaisu)%></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="55" align="right" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sKekkasu)%></td>
					<%Else%>
						<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>>-</td>
						<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>>-</td>
						<td class="<%=w_cell%>" width="100" align="center" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="80" align="center" nowrap  <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="85" align="center" nowrap  <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
					<%End If%>
				<%End If%>
			<%End If%>
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
	<input type="hidden" name="sltShikenKbn" value="<%=m_sSikenKBN%>">
	<input type="hidden" name="txtGakuNo"   value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakkaCd"  value="<%=m_sGakkaCd%>">
	<input type="hidden" name="txtClassNo"  value="<%=m_sClassNo%>">
	<input type="hidden" name="txtKamokuCd" value="<%=m_sKamokuCd%>">
	<input type="hidden" name="txtTUKU_FLG" value="<%=m_iKamokuKbn%>">
	<input type="hidden" name="PasteType"   value="">
	
	<input type="hidden" name="hidSouJyugyou">
	<input type="hidden" name="hidJunJyugyou">
	<input type="hidden" name="hidUpdMode">
	
	
	<input type="hidden" name="hidKamokuKbn" value="<%=m_iKamokuKbn%>">
	<input type="hidden" name="hidKamokuBunrui" value="<%=m_sKamokuBunrui%>">
	<input type="hidden" name="hidSeisekiInpType" value="<%=m_iSeisekiInpType%>">
	<input type="hidden" name="hidKikan" value="<%=m_iKikan%>">
	
	
	<input type="hidden" name="hidFirstGakusekiNo" value="<%=m_FirstGakusekiNo%>">
	
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
%>