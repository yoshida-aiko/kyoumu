<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 欠席日数登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/sei0600/sei0600_main.asp
' 機      能: 下ページ 欠席日数の登録画面
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:
'               選択された試験区分の欠席日数を登録するための画面表示
'-------------------------------------------------------------------------
' 作      成: 2001/09/26 谷脇 良也
' 修      正: 2002/06/11 金澤 香織
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系

    '市町村選択用のWhere条件
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官コード
    Public m_sClass         'クラス
    Public m_sClassNm       'クラス名
    Public m_sGakka     '学生の所属学科
    Public m_iSikenKBN
    Public m_iSyubetu    
    Public m_sGakunen
	'//配列
    Public m_iKesseki()     '欠席数
    Public m_iKibiki()		'忌引等数
    Public m_iKessekiRui()  '欠席集計値
    Public m_iKibikiRui()	'忌引等集計値
    Public m_sGakuseiNo()	'学生番号（5年間番号）
    Public m_sGakusekiNo()	'学籍番号（1年間番号）
    Public m_sGakuSimei()	'学生氏名

    Public  m_GRs,m_DRs
    Public  m_Rs,m_KskRs
    Public  m_iMax          '最大ページ
    Public  m_iDsp          '一覧表示行数
	Public  m_rCnt
	
	
	'------------- 金澤 追加 2002/06/07 -----------------
	Public  m_iTokuketu() '特別欠席
	Public	m_iTokuCnt    '
	'----------------------------------------------------

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="欠席日数登録"
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

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
		'===============================
		'//期間データの取得
		'===============================
        If f_Nyuryokudate() = 1 Then
			'// ページを表示
			Call No_showPage("成績入力期間外です。")
			Exit Do
		End If
		
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		
		'=================
		'//出欠欠課の取り方を取得
		'=================
		'//科目区分(0:試験毎,1:累積)
        w_iRet = gf_GetKanriInfo(m_iNendo,m_iSyubetu)
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		
        '//学生データ取得
        If f_Gakusei() <> 0 Then m_bErrFlg = True : Exit Do
		
		'//日毎出欠集計値データ取得
        If f_GetKessekiData(m_KskRs, m_iSikenKBN, m_sGakunen, m_sClass, w_sKaisibi, w_sSyuryobi, "") <> 0 Then 
        	m_bErrFlg = True
        	Exit Do
        end if
        
		'//集計値データの加工取得
        If f_Kesseki(m_KskRs) <> 0 Then m_bErrFlg = True : Exit Do
		
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

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()
	m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp      = C_PAGE_LINE
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm    = request("txtClassNm")
	m_iSikenKBN    = request("txtSikenKBN")
End Sub

'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_Nyuryokudate()
	dim w_date
	
	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1
	
	w_date = gf_YYYY_MM_DD(date(),"/")
	
	Do
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_iSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & w_date & "' "
		
		If gf_GetRecordset(m_DRs, w_sSQL) <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If
		
		If m_DRs.EOF Then
			Exit Do
		Else
			m_sSikenNm = m_DRs("M01_SYOBUNRUIMEI")
		End If
		
		f_Nyuryokudate = 0
		Exit Do
	Loop
	
End Function

'********************************************************************************
'*  [機能]  学生データを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_Gakusei()
	dim w_Rs,w_sSQL,w_iRet,w_Rs2
	dim w_sSikenKBN,i,w_iCnt,w_vStr
	Dim w_Type
	
	On Error Resume Next
	
	Err.Clear
	f_Gakusei = 1
	
	'---------------------KANAZAWA 2002/6/10--------------------------------------------
	w_sSQL = ""	
	w_sSQL = w_sSQL & vbCrLf & "  SELECT "	
	w_sSQL = w_sSQL & vbCrLf & "  M01_SYOBUNRUIMEI "	
	w_sSQL = w_sSQL & vbCrLf & "  FROM "	
	w_sSQL = w_sSQL & vbCrLf & "   M01_KUBUN "	
	w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo	
	w_sSQL = w_sSQL & vbCrLf & "  AND M01_DAIBUNRUI_CD = " & C_M01_DAIBUNRUI150 '特別欠席	
	w_sSQL = w_sSQL & vbCrLf & "  ORDER BY M01_SYOBUNRUI_CD "	
	
	If gf_GetRecordset(w_Rs2, w_sSQL) <> 0 Then
        m_bErrFlg = True
		Exit function
    End If
    
    
	m_iTokuCnt = gf_GetRsCount(w_Rs2)
	
	'----------------------------------------------------------------------------------
	
  Do
	select case cint(m_iSikenKBN)
		case C_SIKEN_ZEN_TYU '前期中間
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_Z AS KIBIKI "
			w_Type = "_ZT"
		case C_SIKEN_ZEN_KIM '前期期末
			w_sSikenKBN = "T13_KESSEKI_KIMATU_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_KIMATU_Z AS KIBIKI "
			w_Type = "_ZK"
		case C_SIKEN_KOU_TYU '後期中間
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_K AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_K AS KIBIKI "
			w_Type = "_KT"
		case C_SIKEN_KOU_KIM '後期期末（学年末）
			w_sSikenKBN = "T13_SUMKESSEKI AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_SUMKIBTEI AS KIBIKI"
			w_Type = ""
	End select
	
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI,T11_GAKUSEI_NO,T13_GAKUSEKI_NO,"
    
    for w_num=1 to 10
	    w_sSql = w_sSql & " T13_TOKUKETU" & w_num & w_Type &" as T13_TOKUKETU" & w_num & " ,"
    next
    
    w_sSQL = w_sSQL & 		w_sSikenKBN
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI,T13_GAKU_NEN"
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND T11_GAKUSEI_NO = T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & " ORDER BY T13_GAKUSEKI_NO "
	
	If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
		Exit do
    End If
    
	m_rCnt = gf_GetRsCount(w_Rs)
	
	'//配列の作成
    Redim m_iKesseki(m_rCnt)        		'欠席数
    Redim m_iKibiki(m_rCnt)					'忌引等数
    Redim m_iKessekiRui(m_rCnt)  			'欠席集計値
    Redim m_iKibikiRui(m_rCnt)				'忌引等集計値
    Redim m_sGakuseiNo(m_rCnt)				'学生番号（5年間番号）
    Redim m_sGakusekiNo(m_rCnt)				'学籍番号（1年間番号）
    Redim m_sGakuSimei(m_rCnt)				'学生氏名
	Redim m_iTokuketu(m_iTokuCnt,m_rCnt)	'特別欠席
	w_Rs.MoveFirst
	
	i = 1
	w_iCnt = 0
	
	Do Until w_Rs.EOF
		m_iKesseki(i) = cint(gf_SetNull2Zero(w_Rs("KESSEKI")))
		m_iKibiki(i)	= cint(gf_SetNull2Zero(w_Rs("KIBIKI")))
		m_sGakuseiNo(i) = w_Rs("T11_GAKUSEI_NO")
		m_sGakusekiNo(i) = w_Rs("T13_GAKUSEKI_NO")
		m_sGakuSimei(i) = w_Rs("T11_SIMEI")
		m_iKessekiRui(i) = 0
		m_iKibikiRui(i) = 0
		'--------------------------------2002/6/10 kanazawa----------------------------
		For w_iCnt = 1 To m_iTokuCnt
			w_vStr = "T13_TOKUKETU" & w_iCnt
			m_iTokuketu(w_iCnt,i) = cint(gf_SetNull2Zero(w_Rs(w_vStr)))
		Next 
		'------------------------------------------------------------------------------
		i = i + 1
		w_Rs.MoveNext
	Loop

	f_Gakusei = 0 '正常終了
	exit do
  Loop

    Call gf_closeObject(w_Rs)
	Call gf_closeObject(w_Rs2)


End Function

Function f_GetKesskiSu(p_iSikenKBN,p_sGakuseiNo,p_iKessekiSu,p_iKibikiSu)
'********************************************************************************
'*  [機能]  学生データを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
	dim w_Rs,w_sSQL,w_iRet
	dim w_sSikenKBN,i

'	On Error Resume Next
	Err.Clear
	f_GetKesskiSu = 1
	
  Do
	select case cint(p_iSikenKBN)
		case C_SIKEN_ZEN_TYU '前期中間
			f_GetKesskiSu = 0
			Exit Do
		case C_SIKEN_ZEN_KIM '前期期末
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_Z AS KIBIKI"
		case C_SIKEN_KOU_TYU '後期中間
			w_sSikenKBN = "T13_KESSEKI_KIMATU_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_KIMATU_Z AS KIBIKI"
		case C_SIKEN_KOU_KIM '後期期末（学年末）
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_K AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_K AS KIBIKI"
	End select

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & 		w_sSikenKBN
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T13_GAKU_NEN"
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND T13_GAKUSEI_NO = '" & p_sGakuseiNo & "' "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
	    Call gf_closeObject(w_Rs)
		Exit do
    End If
		
		p_iKessekiSu = cint(gf_SetNull2Zero(w_Rs("KESSEKI")))
		p_iKibikiSu	= cint(gf_SetNull2Zero(w_Rs("KIBIKI")))

	f_GetKesskiSu = 0 '正常終了

    Call gf_closeObject(w_Rs)

	exit do
  Loop

End Function

Function F_GetKekkaKubun(p_KksKBN)
'*******************************************************************************
' 機　　能：出欠区分が欠席扱いになるのかならないのかを判定
' 返　　値：取得結果
' 　　　　　(1)欠席扱いする, (0)欠席扱いしない
' 引　　数：p_sKksKBN - 調べたい区分
' 機能詳細：指定された条件の出欠のデータを取得する
' 備　　考：なし
'*******************************************************************************
	dim w_sSQL,w_sRs,s_iRet
	F_GetKekkaKubun = 0
	
    On Error Resume Next
    Err.Clear
		
		w_sSQL = ""
		w_sSql = w_sSql & vbCrLf & "Select "
		w_sSql = w_sSql & vbCrLf & " M01_KEKKA_KBN "
		w_sSql = w_sSql & vbCrLf & "From "
		w_sSql = w_sSql & vbCrLf & " M01_KUBUN "
		w_sSql = w_sSql & vbCrLf & "Where "
		w_sSql = w_sSql & vbCrLf & "     M01_DAIBUNRUI_CD =" & C_KESSEKI 'No19 欠課区分
		w_sSql = w_sSql & vbCrLf & " AND M01_SYOBUNRUI_CD =" & cint(p_KksKBN)
		w_sSql = w_sSql & vbCrLf & " AND M01_NENDO =" & m_iNendo 

		w_iRet = gf_GetRecordset(w_sRs, w_sSQL)
		If w_iRet <> 0 then m_bErrFlg = True : Exit Function
		if w_sRs.EOF = true then m_bErrFlg = True : Exit Function
		
	F_GetKekkaKubun = cint(w_sRs("M01_KEKKA_KBN"))

    Call gf_closeObject(w_sRs)

End Function

Function f_Kesseki(m_KskRs)
'********************************************************************************
'*  [機能]  取得データを整理する。
'*  [引数]  なし
'*  [戻値]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
    Dim w_iKaisu,w_iSktKBN,w_sGakuNo

'    On Error Resume Next
    Err.Clear
   
	f_Kesseki = 1
	'// 集計結果でループ
    Do Until m_KskRs.EOF
    	
    	w_iKaisu = cint(m_KskRs("KAISU"))
    	w_iSktKBN = cint(m_KskRs("T30_SYUKKETU_KBN"))
    	w_sGakuNo = m_KskRs("T30_GAKUSEKI_NO")
    	
		'//学生情報をループ
		For i = 1 to m_rCnt 

			If w_sGakuNo = m_sGakusekiNo(i) then 
			
				If w_iSktKBN = 1 or w_iSktKBN > 3 then '遅刻と早退は除く
					'//欠席区分による集計値への割り振り
			    	If F_GetKekkaKubun(w_iSktKBN) = 1 then 
							m_iKessekiRui(i) = m_iKessekiRui(i) + w_iKaisu
					Else
							m_iKibikiRui(i) = m_iKibikiRui(i) + w_iKaisu
					End If
				End If
			End If
			
    	Next
	
		m_KskRs.MoveNext
    Loop

	f_Kesseki = 0

End Function

Function f_GetKessekiData(p_oRecordset, p_sSikenKbn, p_sGakunen, p_sClass, p_sKaisibi, p_sSyuryobi, p_s1NenBango)
'*******************************************************************************
' 機　　能：出欠データの取得
' 返　　値：取得結果
' 　　　　　(True)成功, (False)失敗
' 引　　数：p_oRecordset - レコードセット
' 　　　　　p_sSikenKbn - 試験区分
' 　　　　　p_sGakunen - 学年
' 　　　　　p_sTantoKyokan - 教官ＣＤ
' 　　　　　p_sClass - クラスNo
' 　　　　　p_sKaisibi - 開始日
' 　　　　　p_sSyuryobi - 終了日
' 　　　　　p_s1NenBango - １年間番号
' 機能詳細：指定された条件の出欠のデータを取得する
' 備　　考：なし
'*******************************************************************************
	Dim w_bRtn			'戻り値
	Dim w_sSql			'SQL
	
'	On Error Resume Next
	'== 初期化 ==
	gf_GetKessekiData = 1
	w_bRtn=False
	w_sSql=""
	'== 出欠を取得する開始日と終了日を取得する ==
	'//(試験間の期間)
	w_bRtn = gf_GetKaisiSyuryo(cint(p_sSikenKbn), p_sGakunen, p_sKaisibi, p_sSyuryobi)

	If w_bRtn <> True Then
		Exit Function
	End If

	'== 出欠を取得する ==
	'SQL作成
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & "SELECT "
	w_sSql = w_sSql & vbCrLf & "	Count(T30_GAKUSEKI_NO) as KAISU,"
	w_sSql = w_sSql & vbCrLf & "	T30_CLASS,"
	w_sSql = w_sSql & vbCrLf & "	T30_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & "	T30_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "FROM "
	w_sSql = w_sSql & vbCrLf & "	T30_KESSEKI "
	w_sSql = w_sSql & vbCrLf & "Where "
	w_sSql = w_sSql & vbCrLf & "	T30_NENDO = " & session("NENDO") & " "		'年度
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_GAKUNEN = " & p_sGakunen & " "					'学年
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_CLASS = " & p_sClass & " "					'クラス
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_HIDUKE >= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sKaisibi & "' "								'開始日
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_HIDUKE <= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sSyuryobi & "' "								'終了日
'	w_sSql = w_sSql & vbCrLf & "	And "
'	w_sSql = w_sSql & vbCrLf & "	T30_SYUKKETU_KBN IN ('" & C_KETU_KEKKA & "','" & C_KETU_TIKOKU & "','"& C_KETU_SOTAI &"')"
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_SYUKKETU_KBN >= " & C_KETU_KEKKA & " "

	'== １年間番号が指定されている場合 ==
	If p_s1NenBango <>"" Then
		w_sSql = w_sSql & vbCrLf & "And "
		w_sSql = w_sSql & vbCrLf & "T30_GAKUSEKI_NO = " & p_s1NenBango & " "			'クラス
	End If
	
	w_sSql = w_sSql & vbCrLf & "Group By "
	w_sSql = w_sSql & vbCrLf & "T30_CLASS,"
	w_sSql = w_sSql & vbCrLf & "T30_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & "T30_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "Order By "
	w_sSql = w_sSql & vbCrLf & "T30_CLASS, "
	w_sSql = w_sSql & vbCrLf & "T30_GAKUSEKI_NO "

	'== データの取得 ==
	Set p_oRecordset = Server.CreateObject("ADODB.Recordset")

	'== 失敗したとき ==
	    If gf_GetRecordset(p_oRecordset, w_sSql) <> 0 Then
		p_oRecordset.Close
		Set p_oRecordset = Nothing
		
		Exit Function
	End If
	gf_GetKessekiData = 0
	
End Function

Sub No_showPage(p_msg)
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
	</head>

	<body>
	<center>
	<br><br><br>
			<span class="msg"><%=p_msg%></span>
	</center>
	</body>

	</html>

<%
End Sub

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
	var chk_Flg;
	chk_Flg = false;
	//************************************************************
	//  [機能]  ページロード時処理
	//  [引数]
	//  [戻値]
	//  [説明]
	//************************************************************
	function window_onload(){
						
        document.frm.target="topFrame";
        document.frm.action="sei0600_topDisp.asp";
        document.frm.submit();
	return true;
	}

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){
        
        //------------------- kanazawa 2002/6/11 ------------------------------------------------------------------------------
        
		// ■■■入力項目の値ﾁｪｯｸ■■■
		var i;
		var w_iNum;
		var w_oKesseki;
		var w_oTokuketu;
		
		for (i = 1; i<= <%=m_rCnt%>; i++){
			w_oKesseki = eval("document.frm.txtKESSEKI_" + i );
			
			if (!f_CheckNum(w_oKesseki)) {
				alert("数値を入力して下さい。");
				return false;
				break;
			}	else {
					for (w_iNum = 1; w_iNum<= <%=m_iTokuCnt%>; w_iNum++) {
						w_oTokuketu = eval("document.frm.txtKIBIKI" + w_iNum + "_" + i );
							if (!f_CheckNum(w_oTokuketu)) {
								alert("半角整数を入力して下さい。");
								return false;
								break;
							}
					}
				}
			
		} //next i;
		
		//------------------------------------------------------------------------------------------------------------------------
		
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return false;
        };
        //'--------------------------------------- kanazawa 2002/6/12 ---------------------------------------------------------------
		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("登録しています・・・・　　しばらくお待ちください")%>"
		//'--------------------------------------------------------------------------------------------------------------------------
		
        document.frm.target="main";
        document.frm.action="sei0600_upd.asp";
        document.frm.submit();
    };
    
// ------------------ kanazawa 2002/6/10 ------------------------------    
//************************************************************
//  [機能]  簡易数値型チェック
//************************************************************
	function f_CheckNum(pFromName){
		var wFromName;
		
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
//----------------------------------------------------------------------	
	
    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Cansel(){

        //document.frm.action="default2.asp";
        //document.frm.target="main";
        document.frm.action="default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();
    
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
	function f_MoveCur(p_inpNm,p_frm,i) {
		if (event.keyCode == 13){		//押されたキーがEnter(13)の時に動く。
			i++;
			if (i > <%=m_rCnt%>) {i = 1;} //iが最大値を超えると、はじめに戻る。
			inpForm = eval("p_frm."+p_inpNm+i);
			inpForm.focus();			//フォーカスを移す。
			inpForm.select();			//移ったテキストボックス内を選択状態にする。
		}else{
	//		alert(event.keyCode);
			return false;
		}
		return true;
	}


//-->
</SCRIPT>
<%
	dim i
	i = 0
	
	'//NN対応
	If session("browser") = "IE" Then
		w_sInputClass = "class='num'"
	Else
		w_sInputClass = ""
	End If

%>
</head>
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">
<center>

<table border="1" cellpadding="1" cellspacing="1" class="hyo">

		<% For i = 1 to m_rCnt 
		        Call gs_cellPtn(w_cell)

				'//欠課累積情報区分が累積のときは、前試験の欠席数および忌引数を取ってくる。
				If cint(m_iSyubetu) = C_K_KEKKA_RUISEKI_KEI then 
					call f_GetKesskiSu(m_iSikenKBN,m_sGakuseiNo(i),w_iKessekiSu,w_iKibikiSu)
					m_iKessekiRui(i) = m_iKessekiRui(i) + w_iKessekiSu
					m_iKibikiRui(i)  = m_iKibikiRui(i)  + w_iKibikiSu
				End If

		        If m_iKesseki(i) = 0 then m_iKesseki(i) = m_iKessekiRui(i)
		        If m_iKibiki(i) = 0 then m_iKibiki(i) = m_iKibikiRui(i)

		%>
            <TR>
                    <TD CLASS="<%=w_cell%>" width="80"><%=m_sGakusekiNo(i)%><input type="hidden" name="txtGAKUSEINO_<%=i%>" value="<%=m_sGakuseiNo(i)%>"></TD>
                    <TD CLASS="<%=w_cell%>" width="150"><%=m_sGakuSimei(i)%></TD>
                    <TD CLASS="<%=w_cell%>" width="35" align="center"><input type="text" <%=w_sInputClass%> name="txtKESSEKI_<%=i%>" value='<%=m_iKesseki(i)%>' size=2 maxlength=3 onKeyDown="f_MoveCur('txtKESSEKI_',this.form,<%=i%>)"></TD>
                    <TD CLASS="<%=w_cell%>" width="35" align="right"><%=m_iKessekiRui(i)%></TD>
                    <!-- '------------------ Kanazawa 2002/6/10 ---------------------------- -->
                    <% dim w_iIndex %>
                    <% For w_iIndex = 1 To m_iTokuCnt %>
                    	<TD CLASS="<%=w_cell%>" width="35" align="center"><input type="text" <%=w_sInputClass%> name="txtKIBIKI<%=w_iIndex%>_<%=i%>" value='<%=m_iTokuketu(w_iIndex,i)%>' size=2 maxlength=3 onKeyDown="f_MoveCur('txtKIBIKI'+<%=w_iIndex%>+'_',this.form,<%=i%>)"></TD>
                    <% Next %>
					<!-- '----------------------------------------------------------------- -->
            </TR>
		<% Next %>
        </td>
    </TR>
</TABLE>

<br>
	<table width="50%">
	<tr>
		<td align="center"><input type="button" class="button" value="　登　録　" onclick="javascript:f_Touroku()">
		<input type="button" class="button" value="キャンセル" onclick="javascript:f_Cansel()"></td>
	</tr>
	</table>

	<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
	<input type="hidden" name="txtSikenKBN" value="<%=m_iSikenKBN%>">
	<input type="hidden" name="txtCnt" value="<%=m_rCnt%>">
	<input type="hidden" name="txtTokuCnt" value="<%=m_iTokuCnt%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
