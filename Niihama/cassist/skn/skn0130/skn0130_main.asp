<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 試験実施科目登録
'* ﾌﾟﾛｸﾞﾗﾑID : skn/skn0130/main.asp
'* 機      能: 下ページ 試験実施科目の一覧リスト表示を行う
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*          txtSikenKbn         :試験区分
'*          txtSikenCd          :試験コード
'*          txtMode             :動作モード
'*                              BLANK   :初期表示
'*                              DISP    :指定された区分のデータを表示
'*                              CHK     :指定された削除分のデータを表示
'*                              DEL     :削除処理を実行
'*          chkDelRenbanX   :削除連番（自分自身から受け取る引数）
'*          txtPageCD         :表示頁数
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*          txtSikenKbn      :選択された試験区分
'*          chkDelRenbanX   :削除連番（自分自身に渡す引数）
'*          txtPageCD         :表示頁数
'* 説      明:
'*           ■初期表示
'*               検索条件にかなう試験中予定を表示
'*           ■修正ボタンクリック時
'*               指定した条件にかなう試験予定を表示させて、修正させる
'*-------------------------------------------------------------------------
'* 作      成: 2001/06/18 高丘 知央
'* 変      更: 2001/06/26 根本
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_sMsg              'ﾒｯｾｰｼﾞ

    '取得したデータを持つ変数
    Public  m_iKyokanCd         ':教官コード
    Public  m_iSyoriNen         ':処理年度
    Public  m_iSikenKbn         ':試験区分
    Public  m_iSikenCode        ':試験ｺｰﾄﾞ
    Public  m_sMode             ':動作モード
    Public  m_sGetTable         ':動作モード、T26にデータがある場合は"26"、T27にデータがある場合は"27"
    Public  m_iCnt              'カウント件数
    Public  m_sPageCD           ':表示済表示頁数（自分自身から受け取る引数）
	Public  m_seisekiF    
    Public  m_Rs                'recordset
	Public  m_sSikenDate		'試験結果データ
	Public	m_iGakunen			'学年

    'ページ関係
    Public  m_iMax              ':最大ページ
    Public  m_iDsp              '// 一覧表示行数
	
	private const C_MAIN_FLG_YES = 1				'メイン教官（メイン教官フラグ）
	'private const C_SEISEKI_INP_FLG_YES = 1			'成績入力教官（成績入力教官フラグ）
	
	Dim m_AryGakunen()
	
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
    Dim w_sWHERE            '// WHERE文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle = "キャンパスアシスト"
    w_sMsgTitle = "試験実施科目登録"
    w_sMsg = ""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget = ""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

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

        if m_sMode <> "" then
			if m_sMode = "no" then
				'// ページを表示
				Call No_showPage("試験準備期間ではありません。")
				Exit Do
			Else
				'===============================
				'//期間データの取得
				'===============================
		        w_iRet = f_Nyuryokudate()
				
				If w_iRet = 1 Then
					    '// 終了処理
					    Call gf_closeObject(m_Rs)
					    Call gs_CloseDatabase()
						response.Redirect "default.asp?txtMode=no&txtSikenKbn="&m_iSikenKbn&""
						response.end
					Exit Do
				End If
				
				If w_iRet <> 0 Then 
					m_bErrFlg = True
					Exit Do
				End If
				
				'//一覧に表示する科目の学年を取得
				If not f_GetGakunen()  then Exit Do
				
				'// 表示用ﾃﾞｰﾀを取得する
				If f_GetKamoku() = False then Exit Do
				
				'// ページを表示
				Call showPage()
			'Exit Do
			End if
		Else
            '// 空白ページを表示
            Call showBrankPage()
        end if
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")          ':教官コード
    m_iSyoriNen = Session("NENDO")              ':処理年度
    m_iSikenKbn = Request("txtSikenKbn")            ':試験区分
    m_iSikenCode = Request("txtSikenCd")            ':試験ｺｰﾄﾞ
    
    if m_iSikenCode = "" then
        m_iSikenCode = 0 
    end if
    
    m_sMode = Request("txtMode")                ':動作モード
    
    '// BLANKの場合は行数ｸﾘｱ
    If Request("txtMode") = "Search" Then
        m_sPageCD = 1
    Else
        m_sPageCD = INT(Request("txtPageCD"))   ':表示済表示頁数（自分自身から受け取る引数）
    End If
	
End Sub

'********************************************************************************
'*  [機能]  クラス数を取得する
'*  [引数]  p_iNendo  ：処理年度
'*          p_iGakuNen：学年
'*  [戻値]  f_GetClassMax：クラス数
'*  [説明]  
'********************************************************************************
Function f_GetClassMax(p_iNendo,p_iGakuNen)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassMax = 0

	Do

		'//クラス名称取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  COUNT(M05_CLASSNO) as ClassMax"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M05_NENDO=" & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND M05_GAKUNEN=" & p_iGakuNen

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If
		
		'//データが取得できたとき
		If rs.EOF = False Then
			'//クラス名
			f_GetClassMax = cint(rs("ClassMax"))
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
'	f_GetClassMax = rs("ClassMax")

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

Function f_Nyuryokudate()
'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
	dim w_date

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1


	w_date = gf_YYYY_MM_DD(date(),"/")
'	w_date = "2000/06/18"
	w_Syuryo = "T24_SIKEN_SYURYO"
	w_kyokan = Session("KYOKAN_CD")
	
	if w_kyokan = NULL or w_kyokan = "" then w_kyokan = "@@@"

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MIN(T24_SIKEN_NITTEI.T24_SIKEN_KAISI) as KAISI"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SIKEN_SYURYO) as SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO) as SEI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iSyoriNen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_iSikenKbn)
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KAISI <= '" & w_date & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI." & w_Syuryo & " >= '" & w_date & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_SYURYO >= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  Group By M01_SYOBUNRUIMEI"

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

			m_sSikenDate = m_DRs("SYURYO") '//okada 2001.12.25

			if m_DRs("SYURYO") > w_date then m_seisekiF = 0 '準備期間内の場合は、成績入力教官のみモード解除
				
				m_sSikenNm = m_DRs("M01_SYOBUNRUIMEI")
		End If
		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  学年・クラス・科目コンボを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKamoku()
	Dim w_sSQL
    Dim w_num
	
    On Error Resume Next
    Err.Clear
    
    f_GetKamoku = False
	w_iCnt = 0
	
	Do 
		m_sGetTable = "T27"
		
		Select Case cint(m_iSikenKbn) '選んだ試験によって、取得科目の開設期間を変える
			Case C_SIKEN_ZEN_TYU
				w_sJiki = C_KAI_ZENKI
			Case C_SIKEN_ZEN_KIM
				w_sJiki = C_KAI_ZENKI
			Case C_SIKEN_KOU_TYU
				w_sJiki = C_KAI_KOUKI
			Case C_SIKEN_KOU_KIM
				w_sJiki = C_KAI_KOUKI
		End Select
		
		w_sGakunenWhere = ""		'//個人履修を見るときのWhereに使用
		w_sSQL = ""
		
		for w_num = 1 to ubound(m_AryGakunen)
			w_sSql = w_sSql & "  SELECT "
			w_sSql = w_sSql & "  T27_GAKUNEN AS GAKUNEN ,"
			w_sSql = w_sSql & "  T27_KAMOKU_CD AS KAMOKU ,"
			w_sSql = w_sSql & "  T15_KAMOKUMEI AS KAMOKUMEI　"
			w_sSql = w_sSql & "  FROM "
			w_sSql = w_sSql & "  T27_TANTO_KYOKAN ,"
			w_sSql = w_sSql & "  M05_CLASS, "
			w_sSql = w_sSql & "  T15_RISYU "
			w_sSql = w_sSql & "  WHERE "
			w_sSql = w_sSql & "      M05_NENDO = T27_NENDO "
			w_sSql = w_sSql & "  AND M05_GAKUNEN =T27_GAKUNEN "
			w_sSql = w_sSql & "  AND M05_CLASSNO = T27_CLASS "
			w_sSql = w_sSql & "  AND M05_GAKKA_CD = T15_GAKKA_CD "
			w_sSql = w_sSql & "  AND T27_KAMOKU_CD = T15_KAMOKU_CD(+) "
			w_sSql = w_sSql & "  AND T15_NYUNENDO(+) = T27_NENDO - T27_GAKUNEN + 1"
			w_sSql = w_sSql & "  AND T27_NENDO = " & m_iSyoriNen
			w_sSql = w_sSql & "  AND T27_KYOKAN_CD ='" & m_iKyokanCd & "' "
			w_sSql = w_sSql & "  AND T27_MAIN_FLG = " & C_MAIN_FLG_YES
			w_sSql = w_sSql & "  and T27_GAKUNEN = " & m_AryGakunen(w_num-1)
			w_sSql = w_sSql & "  AND (T15_KAISETU" & m_AryGakunen(w_num-1) & " =" & w_sJiki & " OR T15_KAISETU" & m_AryGakunen(w_num-1) & " =" & C_KAI_TUNEN & " )"
			w_sSql = w_sSql & "  AND (T27_KAMOKU_CD Not IN (" & f_SubQuery(m_AryGakunen(w_num-1)) & "))"
			
			w_sSql = w_sSql & "  Union "
			
			w_sGakunenWhere = w_sGakunenWhere & m_AryGakunen(w_num-1)
			if w_num <> ubound(m_AryGakunen) then w_sGakunenWhere = w_sGakunenWhere & ","
			
		next
		
		w_sSQL = w_sSQL & " SELECT DISTINCT "
		w_sSQL = w_sSQL & " 	T27_GAKUNEN AS GAKUNEN,"
		w_sSQL = w_sSQL & " 	T27_KAMOKU_CD AS KAMOKU,"
		w_sSQL = w_sSQL & " 	T16_KAMOKUMEI AS KAMOKUMEI"
		w_sSQL = w_sSQL & " FROM"
		w_sSQL = w_sSQL & " 	T27_TANTO_KYOKAN,"
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 	T27_GAKUNEN in (" & w_sGakunenWhere & ") and "
		w_sSQL = w_sSQL & " 	T27_KAMOKU_CD = T16_KAMOKU_CD(+) and "
		w_sSQL = w_sSQL & " 	T27_NENDO = T16_NENDO(+) and "
		w_sSQL = w_sSQL & " 	T27_NENDO = " & m_iSyoriNen & " and "
		w_sSQL = w_sSQL & " 	T27_KYOKAN_CD ='" & m_iKyokanCd & "' and "
		w_sSql = w_sSql & " 	T27_MAIN_FLG = " & C_MAIN_FLG_YES & " and "
		w_sSQL = w_sSQL & " 	T16_OKIKAE_FLG >= " & C_TIKAN_KAMOKU_SAKI 
		w_sSQL = w_sSQL & " GROUP BY "
		w_sSQL = w_sSQL & " 	T27_NENDO"
		w_sSQL = w_sSQL & " 	,T27_GAKUNEN"
		w_sSQL = w_sSQL & " 	,T27_CLASS"
		w_sSQL = w_sSQL & " 	,T27_KAMOKU_CD"
		w_sSQL = w_sSQL & " 	,T16_KAMOKUMEI"
		
		If gf_GetRecordset(m_Rs, w_sSQL) <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKamoku = 99
            Exit Do
        End If

        f_GetKamoku = True
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  置換先科目コードを取ってくるｻﾌﾞｸｴﾘｰ
'*  [引数]  pGakunen = 学年ＣＤ
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_SubQuery(pGakunen)
Dim w_sSubSql

	On Error Resume Next
	Err.Clear

	w_sSubSql = ""
	w_sSubSql = w_sSubSql & " SELECT "
	w_sSubSql = w_sSubSql & " 	  T65_KAMOKU_CD_SAKI "
	w_sSubSql = w_sSubSql & " FROM "
	w_sSubSql = w_sSubSql & " 	  T65_RISYU_SENOKIKAE "
	w_sSubSql = w_sSubSql & " WHERE "
	w_sSubSql = w_sSubSql & " 	  T65_NENDO    = " & m_iSyoriNen
'	w_sSubSql = w_sSubSql & " AND T65_GAKKA_CD = '06' "
	w_sSubSql = w_sSubSql & " AND T65_GAKUNEN  = " & pGakunen

	f_SubQuery = w_sSubSql

End Function


'********************************************************************************
'*  [機能]  試験時間割の中にログインユーザの担当(成績)する科目数を出す
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_SikenJWariCnt(p_iCnt)

    Dim w_sSQL,iRet,w_Rs

    On Error Resume Next
    Err.Clear
    
    f_SikenJWariCnt = 1

		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & " T26_GAKUNEN AS GAKUNEN"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_CLASS AS CLASS"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_KAMOKU AS KAMOKU"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE T26_NENDO = " & m_iSyoriNen

If m_iSikenKbn < C_SIKEN_KOU_KIM then '年度末試験の場合は、すべてが対象
		w_sSQL = w_sSQL & vbCrLf & "    AND T26_SIKEN_KBN =" & m_iSikenKbn
End If

		w_sSQL = w_sSQL & vbCrLf & "    AND T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "'"
		w_sSQL = w_sSQL & vbCrLf & "    AND T26_JISSI_KYOKAN ='" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "  T26_NENDO"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_GAKUNEN"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_KAMOKU"
		
'response.write w_sSQL  & "<BR>"
'response.end
        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_SikenJWariCnt = 99
			Exit Function
        End If

		p_iCnt = gf_GetRsCount(w_Rs)

        f_SikenJWariCnt = 0

End Function

'********************************************************************************
'*  [機能]  クラス名を取得する
'*  [引数]  p_iNendo  ：処理年度
'*          p_iGakuNen：学年
'*          p_Kamoku  ：科目
'*  [戻値]  p_Class	      ：クラスNO
'* 　　　　 f_GetClassName：クラス名
'*  [説明]  
'********************************************************************************
Function f_GetClassName(p_iNendo,p_iGakuNen,p_Kamoku,p_iKyositu,p_ClassNo)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassName = ""
	w_sClassName = ""
    p_ClassNo = ""
	w_clsMax = f_GetClassMax(m_iSyoriNen,m_Rs("GAKUNEN"))

	Do

		'//クラス名称取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT DISTINCT"
'		w_sSql = w_sSql & vbCrLf & "  M05.M05_CLASSMEI AS CLASSMEI,"
		w_sSql = w_sSql & vbCrLf & "   M05.M05_CLASSRYAKU AS CLASSMEI"
		w_sSql = w_sSql & vbCrLf & "  ,M05.M05_CLASSNO "
		w_sSql = w_sSql & vbCrLf & "  ,T27.T27_CLASS"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS M05 , T27_TANTO_KYOKAN T27"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T27.T27_NENDO=" & p_iNendo
'		w_sSQL = w_sSQL & vbCrLf & "  AND T27.T27_KYOSITU_CD = " & p_iKyositu
		w_sSql = w_sSql & vbCrLf & "  AND T27.T27_GAKUNEN=" & p_iGakuNen
		w_sSql = w_sSql & vbCrLf & "  AND T27.T27_KYOKAN_CD ='" & m_iKyokanCd & "' "
		w_sSql = w_sSql & vbCrLf & "  AND T27.T27_KAMOKU_CD ='" & p_Kamoku & "' "
	    w_sSql = w_sSql & vbCrLf & "  AND T27_MAIN_FLG  = " & C_MAIN_FLG_YES 
		w_sSql = w_sSql & vbCrLf & "  AND M05.M05_NENDO = T27.T27_NENDO"
		w_sSql = w_sSql & vbCrLf & "  AND M05.M05_GAKUNEN = T27.T27_GAKUNEN"
		w_sSql = w_sSql & vbCrLf & "  AND M05.M05_CLASSNO = T27.T27_CLASS "
		w_sSql = w_sSql & vbCrLf & " ORDER BY  "
		'w_sSql = w_sSql & vbCrLf & " M05.M05_CLASSNO "
		w_sSql = w_sSql & vbCrLf & " T27.T27_CLASS "

'response.write w_sSql
		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		i = 0
		Do Until rs.EOF
			w_sClassName = w_sClassName & "," & rs("CLASSMEI")
			p_ClassNo = p_ClassNo & "#" & rs("T27_CLASS")
			
			i = i + 1
			rs.MoveNext
		Loop

		If p_ClassNo <> "" then p_ClassNo = Mid(p_ClassNo,2)
		If w_sClassName <> "" then w_sClassName = Mid(w_sClassName,2)

		If i >= w_clsMax then w_sClassName = "全"

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	f_GetClassName = w_sClassName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  試験時間割データより実施時間を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSikenJikan(p_iNendo,p_iGakuNen,p_Kamoku,p_sSikenJikan,p_sSJissi,p_sKyositu)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetSikenJikan = False
    w_sSikenJikan = 0

    Do
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  T26_NENDO,"					'年度
        w_sSql = w_sSql & vbCrLf & "  T26_SIKEN_KBN,"               '試験区分
        w_sSql = w_sSql & vbCrLf & "  T26_SIKEN_CD,"                '試験コード
        w_sSql = w_sSql & vbCrLf & "  T26_GAKUNEN,"                 '学年
        w_sSql = w_sSql & vbCrLf & "  T26_CLASS,"                   'クラスＮＯ
        w_sSql = w_sSql & vbCrLf & "  T26_KAMOKU,"                  '科目コード
        w_sSql = w_sSql & vbCrLf & "  T26_JISSI_KYOKAN,"            '実施教官コード/成績入力教官
        w_sSql = w_sSql & vbCrLf & "  T26_JISSI_FLG,"               '実施フラグ
        w_sSql = w_sSql & vbCrLf & "  T26_SIKENBI,"                 '実施日付
        w_sSql = w_sSql & vbCrLf & "  T26_MAIN_FLG,"                'メイン教官フラグ 
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_INP_FLG,"         '成績入力教官フラグ 
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN1 ,"        '成績入力教官コード1  
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN2,"         '成績入力教官コード2  
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN3,"         '成績入力教官コード3  
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN4,"         '成績入力教官コード4  
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN5,"         '成績入力教官コード5  
        w_sSql = w_sSql & vbCrLf & "  T26_KANTOKU_KYOKAN,"          '監督教官コード
        w_sSql = w_sSql & vbCrLf & "  T26_KYOSITU,"                 '実施教室コード
        w_sSql = w_sSql & vbCrLf & "  T26_SIKEN_JIKAN,"             '試験時間
        w_sSql = w_sSql & vbCrLf & "  T26_KAISI_JIKOKU,"            '開始時刻
        w_sSql = w_sSql & vbCrLf & "  T26_SYURYO_JIKOKU,"           '終了時刻
        w_sSql = w_sSql & vbCrLf & "  T26_KYOKAN_RENMEI "           '教官連名
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  T26_SIKEN_JIKANWARI "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      T26_NENDO = " & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND T26_JISSI_KYOKAN='" & m_iKyokanCd & "'"
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_KBN = " & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_CD = '" & m_iSikenCode & "' "
        w_sSql = w_sSql & vbCrLf & "  AND T26_GAKUNEN = " & p_iGakuNen
        w_sSql = w_sSql & vbCrLf & "  AND T26_KAMOKU ='" & p_Kamoku & "' "
'response.write w_ssql & "<br>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		p_sSikenJikan = 0

        If rs.EOF = False Then
            p_sSikenJikan = rs("T26_SIKEN_JIKAN")
            p_sSJissi = rs("T26_JISSI_FLG")
            p_sKyositu = rs("T26_KYOSITU")
'response.write "aa(" & p_sKyositu & ")<br>"

        End If
        Exit Do
    Loop

	f_GetSikenJikan = True
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  表示項目(試験)を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSikenName()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetSikenName = ""
    w_sSikenName = ""

    Do
        '試験マスタよりデータを取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_iSikenKbn

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

        If rs.EOF = False Then
            w_sSikenName = rs("M01_SYOBUNRUIMEI")
        End If

        Exit Do
    Loop

	f_GetSikenName = w_sSikenName

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  一覧に表示する科目の学年を取得
'*  [引数]  なし
'*  [戻値]  True→成功、False→失敗
'*  [説明]  年度、試験区分、また実施開始日か実施終了日がNULLでないもので検索
'*          データが取れない学年は、表示しない
'********************************************************************************
Function f_GetGakunen()
	Dim w_sSQL
	Dim wRs
	Dim wCnt,w_num
	
	On Error Resume Next
	Err.Clear
	
	f_GetGakunen = false
	
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	* "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & " 	T24_SIKEN_NITTEI "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	T24_NENDO = " & m_iSyoriNen & " and "
	w_sSql = w_sSql & " 	T24_SIKEN_KBN = " & m_iSikenKbn & " and "
	w_sSql = w_sSql & " 	T24_JISSI_KAISI is not NULL and "
	w_sSql = w_sSql & " 	T24_JISSI_SYURYO is not NULL "
	
	w_sSql = w_sSql & " order by "
	w_sSql = w_sSql & " 	T24_GAKUNEN "
	
	If gf_GetRecordset(wRs, w_sSQL) <> 0 Then exit function
	
	If wRs.EOF Then
		ReDim m_AryGakunen(1)
		m_AryGakunen(0) = 9
		exit function
	end if
	
	wCnt = gf_GetRsCount(wRs)
	
	ReDim m_AryGakunen(wCnt)
	
	for w_num = 1 to wCnt
		m_AryGakunen(w_num-1) = wRs("T24_GAKUNEN")
		wRs.movenext
	next
	
	f_GetGakunen = true
	
	Call gf_closeObject(wRs)
	
End Function

'********************************************************************************
'*  [機能]  表示項目(教室)を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKyositu(p_lKyositu)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetKyositu = ""
    w_sKyositu = ""

    Do
        '教室マスタよりデータを取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "   M06_KYOSITUMEI "
        w_sSql = w_sSql & vbCrLf & "  ,M06_RYAKUSYO "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      M06_NENDO = " & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M06_KYOSITU_CD = " & p_lKyositu

'response.write w_sSql

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

        If rs.EOF = False Then
            w_sKyositu = rs("M06_KYOSITUMEI")
        End If

        Exit Do
    Loop

	f_GetKyositu = w_sKyositu

    Call gf_closeObject(rs)

End Function

    '---------- 関数 end ----------

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear

    '---------- HTML START ----------
    Dim w_lJikan	'試験実施時間
	Dim w_className
	Dim w_class
	Dim w_sSikenJikan
	Dim w_sSJissi
	Dim w_sKyosituCd
	Dim w_sKyositu
%>

<html>

<head>
    <title>試験実施科目登録</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

    }

    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  修正ボタン押下時の処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_regist(p_Gakunen,p_Class,p_Kamoku) {

        document.frm.txtGakunen.value=p_Gakunen;
        document.frm.txtClass.value=p_Class;
        document.frm.txtKamoku.value=p_Kamoku;
        
        document.frm.action="skn0130_regist.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();
    }

    //************************************************************
    //  [機能]  キャンセルボタン押下時の処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_Back() {
		location.href = "default.asp"
    }
    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>
	<body>
	    <center>
	    <form name=frm>
		<%call gs_title("試験実施科目登録","一　覧")%>
		<br>
		<table class="hyo" border="1" width="260" height="20">
		    <tr>
		        <th class="header" width="80"  align="center" nowrap>実施試験</th>
		        <td class="detail" width="180" align="center" nowrap><%=f_GetSikenName()%></td>
		    </tr>
		</table>
	<br>
	<hr size="1" color="#000000" >
	<input class="button" type="button" onclick="javascript:f_Back();" value="キャンセル">
	<br><br>
		<% if m_Rs.eof then %>
			<br><br><br>
			<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
		<% Else %>
	<span class=CAUTION>
		※ 修正の場合は｢>>｣をクリックしてください。<br>
	</span>

	    <table width="80%">
	        <tr>
	            <td align="center">
	                <%
	                    'ページBAR表示
	                    Call gs_pageBar(m_Rs,m_sPageCD,m_iDsp,w_pageBar)
	                %>
	                <%=w_pageBar %>

	                <table border="1" width="100%" class="hyo">
	                    <tr>
	                        <th class=header align="center" nowrap><font color="#ffffff">クラス</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">科目名称</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">実　施</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">時　間</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">実施教室</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">修　正</font></th>
	                    </tr>
			<%
				w_i = 1
				
				do until m_Rs.eof or w_i > C_PAGE_LINE
					m_JISSI_FLG = ""
					w_sSikenJikan = 0
					w_sSJissi = 0
					w_sKyosituCd = 0
					w_lJikan = f_GetSikenJikan(m_iSyoriNen,m_Rs("GAKUNEN"),m_Rs("KAMOKU"),w_sSikenJikan,w_sSJissi,w_sKyosituCd)				'試験時間取得
					m_JISSI_FLG = w_sSJissi
					m_sjissi_cls = "detail"
					
					Call gf_GetKubunName(C_SIKEN_KBN,m_JISSI_FLG,Session("NENDO"),m_JISSI_FLG)
					
					if m_JISSI_FLG = "0" Or m_JISSI_FLG = "" then 
						m_JISSI_FLG = "未入力" '実施区分が取れないときは未入力とみなす。
					End IF
					
					if m_JISSI_FLG = "未入力" then 
						m_sjissi_cls = "JISSHIMI"
					End If
					
					if isnull(w_sKyosituCd) = true Then
						w_sKyositu = "" 'w_sKyosituCd = 'm_Rs("KYOSITU") '//Add 2002.1.23
					Else
						 w_sKyositu = f_GetKyositu(w_sKyosituCd)	
					end if
					
					w_className = f_GetClassName(m_iSyoriNen,m_Rs("GAKUNEN"),m_Rs("KAMOKU"),w_sKyosituCd,w_class)'m_Rs("KYOSITU"),w_class)	'クラス名取得
										'教室名取得
					m_iGakunen = m_Rs("GAKUNEN") '// Add 2001.12.26
					
					%>
					<tr>
						<td class="detail" align="left" nowrap>　<%=gf_HTMLTableSTR(m_Rs("GAKUNEN"))%> - <%=w_className%></td>
						<td class="detail" align="center" nowrap><%=gf_HTMLTableSTR(m_Rs("KAMOKUMEI"))%></td>
						<td class="<%=m_sjissi_cls%>" align="center" nowrap><%=gf_HTMLTableSTR(m_JISSI_FLG)%></td>
						
						<td class=detail align="center" nowrap>
							<%If isnull(w_sSikenJikan) = True Or Trim(w_sSikenJikan) = "" Then w_sSikenJikan = "0"%>
							<%=gf_HTMLTableSTR(w_sSikenJikan)%>分
						</td>
						
						<td class=detail align="center" nowrap><%=gf_HTMLTableSTR(w_sKyositu)%></td>
						<td class=detail align="center" nowrap>
					<%
						'//2001/12/07 メイン教官でないなら詳細画面に遷移させない
						If CStr(m_Rs("T27_MAIN_FLG")) = CStr(C_MAIN_KYOKAN_YES) Then
					%>
						<input type="button" class="button" name="Change" value=">>" onclick="f_regist(<%=gf_HTMLTableSTR(m_Rs("GAKUNEN"))%>,'<%=w_class%>','<%=gf_HTMLTableSTR(m_Rs("KAMOKU"))%>')">
					<%
						End If
					%>
						</td>
					</tr>
			<%
					w_i = w_i + 1
					m_Rs.movenext
				loop
			%>
				<%=w_pageBar %>

	                </td>
	            </tr>
	        </table>

		<% End if %>

	        <input type="hidden" name="txtGetTable" value = "<%=m_sGetTable%>">
	        <input type="hidden" name="txtMode" value = "<%=m_sMode%>">
	        <input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
	        <input type="hidden" name="txtSikenKbn" value="<%= m_iSikenKbn %>">
	        <input type="hidden" name="txtSikenCd" value="<%= m_iSikenCode %>">
	        <input type="hidden" name="txtSeisekiFlg" value="<%= m_seisekiF %>">

	        <input type="hidden" name="txtGakunen" value="<%= m_iGakunen %>">
	        <input type="hidden" name="txtClass" value="">
	        <input type="hidden" name="txtKamoku" value="">

	        <input type="hidden" name="txtKikan" value="<%=m_sSikenDate%>">
<!--	        <input type="hidden" name="txtkikanEnd" value=""> -->
			

	    </form>

	    </table>

	    </center>

	</body>

	</html>

<%
    '---------- HTML END   ----------
End Sub

'********************************************************************************
'*  [機能]  空白HTMLを出力
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Sub showBrankPage()
%>
<html>
<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
</head>

<body>
<center>
<br><br><br>
<span class="msg"><%=C_BRANK_VIEW_MSG%></span>
</center>
</body>

</html>
<% End Sub 

Sub No_showPage(p_msg)
'********************************************************************************
'*  [機能]  空白HTMLを出力
'*  [引数]  
'*  [戻値]  
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
<% End Sub %>
