<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: （試験用）教官予定登録
'* ﾌﾟﾛｸﾞﾗﾑID : skn/skn0120/main.asp
'* 機      能: 下ページ 試験予定マスタの一覧リスト表示を行う
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
'*          txtPage         :表示頁数
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*          txtSikenKbn      :選択された試験区分
'*          chkDelRenbanX   :削除連番（自分自身に渡す引数）
'*          txtPage         :表示頁数
'* 説      明:
'*           ■初期表示
'*               検索条件にかなう試験中予定を表示
'*           ■修正ボタンクリック時
'*               指定した条件にかなう試験予定を表示させて、修正させる
'*           ■登録ボタンクリック時
'*               試験予定入力を表示させて、登録させる
'*           ■削除ボタンクリック時
'*               指定した条件にかなう試験を削除する
'*              本ページにて、削除の処理も行う
'*-------------------------------------------------------------------------
'* 作      成: 2001/06/18 高丘 知央
'* 変      更: 2001/06/26 根本
'* 変      更: 2001/08/03 伊藤公子 試験期間を表示するよう修正
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙｺﾝｽﾄ /////////////////////////////
	CONST C_MIN_TIME = "00:00"		'//最小時刻
	CONST C_MAX_TIME = "23:55"		'//最大時刻(時間)
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_sMsg              'ﾒｯｾｰｼﾞ

    '取得したデータを持つ変数
    Public  m_iKyokanCd         ':教官コード
    Public  m_iSyoriNen         ':処理年度
    Public  m_iSikenKbn         ':試験区分
    Public  m_iSikenCd          ':試験コード
    Public  m_sMode             ':動作モード
    Public  m_iRenban           ':連番
    Public  m_sYoteiMei         ':予定名称
    'Public  m_iRiyu            ':理由（予定コード）
    Public  m_dtYoteiKaisi      ':予定開始時間
    Public  m_dtYoteiSyuryo     ':予定終了時間
    Public  m_iMonth            '予定日（月）
    Public  m_iDay              '予定日（日）
    Public  m_sYobi             '予定日（曜日）
    Public  m_iRMax             '最大連番値
    Public  m_iCnt              'カウント件数
    Public  m_iPage             '表示済表示頁数（自分自身から受け取る引数）
    Public  m_iYoteiBi


    Public  m_Rs                'recordset

    'ページ関係
    Public  m_iMax              ':最大ページ
    Public  m_iDsp              '// 一覧表示行数

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
    w_sMsgTitle = "試験監督免除登録"
    w_sMsg = ""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"


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

'response.write "モード = " & Request("txtMode") & "<BR>"

		'===============================
		'//期間データの取得
		'===============================
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			'// ページを表示
			Call showPage_NoData("試験準備期間ではありません。")
			Exit Do
		End If
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If

        '// 削除処理
        If trim(Request("txtMode")) = "DEL" Then
			'//削除処理実行
			w_iRet = s_DelData()
			If w_iRet <> 0 Then
				m_bErrFlg = True
                Exit Do
			End If
        End If

		'削除確認画面を表示
		If Request("txtMode") = "CHK" Then

			'//削除選択されたデータを取得
			w_iRet = f_GetDeleteData()
			If w_iRet <> 0 Then
				m_bErrFlg = True
                Exit Do
			End If

		Else
	        '教官予定マスタを取得
	        w_sWHERE = ""

	        w_sSQL = ""
	        w_sSQL = w_sSQL & "SELECT "
	        w_sSQL = w_sSQL & vbCrLf & " T25_NENDO "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_SIKEN_KBN "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_SIKEN_CD "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_KYOKAN "        
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEIBI "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_RENBAN "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEI_KAISI "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEI_SYURYO "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_BIKO "
	        w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
	        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
	        w_sSQL = w_sSQL & vbCrLf & " T25_NENDO = " & m_iSyoriNen
	        w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"

	        '抽出条件の作成
	        If m_iSikenKbn <> "" Then
	           w_sSQL = w_sSQL & " AND T25_SIKEN_KBN = " & m_iSikenKbn
	           w_sSQL = w_sSQL & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
	        End If

	        'w_sSQL = w_sSQL & " ORDER BY T25_YOTEIBI ASC"
	        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T25_YOTEIBI ,T25_YOTEI_KAISI"

	        Set m_Rs = Server.CreateObject("ADODB.Recordset")
	        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
	        If w_iRet <> 0 Then
	            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
	            m_bErrFlg = True
	            m_sErrMsg = "レコードセットの取得に失敗しました"
	            Exit Do
	        Else
	            'ページ数の取得
	            m_iMax = gf_PageCount(m_Rs,m_iDsp)
	        End If
		End If

		If m_Rs.EOF Then
            '// ページを表示
            Call showPage_NoData("対象データは存在しません。条件を入力しなおして検索してください。")
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
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  DBから値を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_GetParam()

    Dim w_iDay
    Dim w_iMonth
    Dim w_sYobi
    
    '//日付と曜日表示
    w_iDay = ""
    w_iMonth = ""
    w_iYobi = ""
    w_iDay = f_GetDay(m_Rs("T25_YOTEIBI"))    
    w_iMonth = f_GetMonth(m_Rs("T25_YOTEIBI"))    
    w_sYobi = left(WeekdayName(Weekday(CDate(m_Rs("T25_YOTEIBI")))) ,1)
    m_iMonth = gf_fmtZero(w_iMonth,2)
    m_iDay = gf_fmtZero(w_iDay,2)
    m_sYobi = w_sYobi
    m_sYoteiMei = m_Rs("T25_BIKO")

    '//時刻表示
    m_dtYoteiKaisi = m_Rs("T25_YOTEI_KAISI")
    m_dtYoteiSyuryo = m_Rs("T25_YOTEI_SYURYO")

    '//連番表示
    m_iRenban = m_Rs("T25_RENBAN")

    m_iYoteiBi = m_Rs("T25_YOTEIBI")

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")         ':教官コード
    m_iSyoriNen = Session("NENDO")             ':処理年度
    m_iSikenKbn = Request("txtSikenKbn")       ':試験区分

    if Request("txtSikenCd") <> "" Then
        m_iSikenCd = Request("txtSikenCd")      ':試験コード
    else
        m_iSikenCd = 0
    end if

    m_sMode = Request("txtMode")                ':動作モード
    
    m_iRenban = Request("txtRenban")            ':連番    '//保留

    '// BLANKの場合は行数ｸﾘｱ
    If Request("txtMode") = "Search" Then
        m_iPage = 1
    Else
        m_iPage = INT(Request("txtPage"))   ':表示済表示頁数（自分自身から受け取る引数）
    End If

End Sub

'********************************************************************************
'*  [機能]  引き渡されてきた値を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ShowRequest()

    Response.write "txtMode=BLANK" & "&txtSikenKbn=" & m_iSikenKbn

End Sub

'********************************************************************************
'*  [機能]  表示項目(試験)を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDeleteData()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDeleteData = 1

    Do

		'//選択された情報を取得
		w_sDelData = replace(Request("chkDel")," ","")
		w_sDelData = split(w_sDelData,",")
		w_iCnt = UBound(w_sDelData)

        'マスタよりデータを取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "SELECT "
		w_sSQL = w_sSQL & vbCrLf & " T25_NENDO "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_SIKEN_KBN "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_SIKEN_CD "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_KYOKAN "        
		w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEIBI "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_RENBAN "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEI_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_BIKO "
		w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
		w_sSQL = w_sSQL & vbCrLf & " WHERE " 
		w_sSQL = w_sSQL & vbCrLf & " T25_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
		w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " AND ("

		For i = 0 To w_iCnt
			If i <> 0 Then
	            w_sSQL = w_sSQL & vbCrLf & " Or "
			End If

			w_Ary = split(w_sDelData(i),"_")
            w_sSQL = w_sSQL & vbCrLf & "  ( T25_YOTEIBI = '" & w_Ary(0) & "'"
            w_sSQL = w_sSQL & vbCrLf & "      AND T25_RENBAN = '" & w_Ary(1) & "'"
            w_sSQL = w_sSQL & vbCrLf & "   )"
		Next

            w_sSQL = w_sSQL & vbCrLf & " )"
	        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T25_YOTEIBI ,T25_YOTEI_KAISI"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetDeleteData = 99
            Exit Do
        End If

        '//正常終了
        f_GetDeleteData = 0

        Exit Do
    Loop

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

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MIN(T24_SIKEN_NITTEI.T24_SIKEN_KAISI) as KAISI"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SIKEN_SYURYO) as SYURYO"
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
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KAISI <= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_SYURYO >= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KAISI Is Not Null "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_SYURYO Is Not Null "
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
			m_sSikenNm = m_DRs("M01_SYOBUNRUIMEI")
		End If
		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function

'
''********************************************************************************
''*  [機能]  指定担当者を削除する
''*  [引数]  なし
''*  [戻値]  なし
''*  [説明]  
''********************************************************************************
'Sub s_DelData()
'
'    Dim w_sCD               '// 連番
'    Dim w_iRet              '// 戻り値
'    Dim w_sSQL              '// SQL文
'    
'    m_iCnt = CInt(Request("txtCnt"))
'    
'    For i = 1 to m_iCnt
'        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
'        Call gs_BeginTrans()
'
'        If Request("chkDelRenban" & i) <> "" Then
'
'            w_sCD = Request("chkDelRenban" & i)
'            '// 教官予定マスタﾚｺｰﾄﾞｾｯﾄを取得
'            w_sSQL = ""
'            w_sSQL = w_sSQL & "DELETE "
'            w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
'            w_sSQL = w_sSQL & vbCrLf & " WHERE " 
'            w_sSQL = w_sSQL & vbCrLf & " T25_NENDO = " & m_iSyoriNen
'            w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
'            w_sSQL = w_sSQL & vbCrLf & " AND T25_RENBAN = " & w_sCD
'            w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
'
'            w_iRet = gf_ExecuteSQL(w_sSQL)
'            If w_iRet <> 0 Then
'                '//ﾛｰﾙﾊﾞｯｸ
'                Call gs_RollbackTrans()
'                
'                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
'                m_bErrFlg = True
'                m_sErrMsg = "削除に失敗しました。"
'                Exit Sub 'GOTO MAIN
'            End If
'
'
'        End If
'
'    Next
'    
'    '//ｺﾐｯﾄ
'    Call gs_CommitTrans()
'
'End Sub

'********************************************************************************
'*  [機能]  予定を削除する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function  s_DelData()

    Dim w_sCD               '// 連番
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文

    On Error Resume Next
    Err.Clear

	s_DelData = 1

	Do

		'//削除情報を取得
		w_sDelData = replace(Request("chkDel")," ","")
		w_sDelData = split(w_sDelData,",")
		w_iCnt = UBound(w_sDelData)

		'//選択された教官予定を削除する
		w_sSQL = ""
		w_sSQL = w_sSQL & "DELETE "
		w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
		w_sSQL = w_sSQL & vbCrLf & " WHERE " 
		w_sSQL = w_sSQL & vbCrLf & " T25_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
		w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " AND ("

		For i = 0 To w_iCnt
			If i <> 0 Then
	            w_sSQL = w_sSQL & vbCrLf & " Or "
			End If

			w_Ary = split(w_sDelData(i),"_")
            w_sSQL = w_sSQL & vbCrLf & "  ( T25_YOTEIBI = '" & w_Ary(0) & "'"
            w_sSQL = w_sSQL & vbCrLf & "      AND T25_RENBAN = '" & w_Ary(1) & "'"
            w_sSQL = w_sSQL & vbCrLf & "   )"
		Next
        w_sSQL = w_sSQL & vbCrLf & " )"

		w_iRet = gf_ExecuteSQL(w_sSQL)
		If w_iRet <> 0 Then
		    '削除失敗
			s_DelData = 99
		    m_bErrFlg = True
		    m_sErrMsg = "削除に失敗しました。"
		    Exit Do
		End If

		'//正常終了時
		s_DelData = 0
		Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  YYYYMMDD形式の日付から月を抽出
'*  [引数]  YYYYMMDD形式の日付
'*  [戻値]  MM形式の月
'*  [説明]  
'********************************************************************************
Function f_GetMonth(p_sDate)

    f_GetMonth = ""

    If Trim(gf_SetNull2String(p_sDate)) = "" Then
        f_GetMonth = ""
        Exit Function
    End If
    
    f_GetMonth = Month(gf_FormatDate(p_sDate,"/"))

End Function

'********************************************************************************
'*  [機能]  YYYYMMDD形式の日付から日を抽出
'*  [引数]  YYYYMMDD形式の日付
'*  [戻値]  DD形式の日
'*  [説明]  
'********************************************************************************
Function f_GetDay(p_sDate)

    f_GetDay = ""

    If Trim(gf_SetNull2String(p_sDate)) = "" Then
        f_GetDay = ""
        Exit Function
    End If
    
    f_GetDay = Day(gf_FormatDate(p_sDate,"/"))

End Function

'********************************************************************************
'*  [機能]  時刻整形
'*  [引数]  数字のみ時刻(hhnn形式のもののみ)
'*  [戻値]  区切り文字付き時刻(エラー時、引数をそのまま)
'*  [説明]  数字のみの時分を区切り文字で分ける。
'*  [変更]  DB4桁→5桁
'*          hhnn→hh:nn
'********************************************************************************
Function gf_FormatTime(p_Time,p_Delimiter)
    Dim w_sTime 
    Dim w_sHour
    Dim w_sMinute

    '空白ならエラー
    If IsNull(p_Time)  Then
        gf_FormatTime = p_Time
        Exit Function
    End If
    If p_Time = "" Then 
        gf_FormatTime = p_Time
        Exit Function
    End If

    '数字でないならエラー
    If Not IsNumeric(p_Time) Then 
        gf_FormatTime = p_time
        Exit Function
    End If

    '4桁でないならエラー
    If Len(p_Time) <> 4 Then
        gf_FormatTime = p_Time
        Exit Function
    End If

    w_sHour = Mid(p_Time,1,2)
    w_sMinute  = Mid(p_Time,3,2)

    w_sTime = w_sHour & p_Delimiter 
    w_sTime = w_sTime & w_sMinute

    '最終的に日付でないならエラー
    'If Not IsDate(w_sDate) Then    
    '   gf_FormatDate = p_Date
    '   Exit Function
    'End If

    gf_FormatTime = w_sTime

End Function

'********************************************************************************
'*  [機能]  学年ごとの試験期間を取得
'*  [引数]  なし
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Function f_GetSikenKikan()

    Dim w_Rs2                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet2              '// 戻り値
    Dim w_sSQL2              '// SQL文

    On Error Resume Next
    Err.Clear
    f_GetSikenKikan = True

    Do

        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  T24.T24_SIKEN_KBN"
        w_sSql = w_sSql & vbCrLf & "  ,T24.T24_SIKEN_CD"
        w_sSql = w_sSql & vbCrLf & "  ,T24.T24_GAKUNEN"
        w_sSql = w_sSql & vbCrLf & "  ,T24.T24_JISSI_KAISI"
        w_sSql = w_sSql & vbCrLf & "  ,T24.T24_JISSI_SYURYO"
        w_sSql = w_sSql & vbCrLf & " FROM T24_SIKEN_NITTEI T24"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      T24.T24_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND T24.T24_SIKEN_KBN= " & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND T24.T24_SIKEN_CD='" & m_iSikenCd & "'"
        w_sSql = w_sSql & vbCrLf & " ORDER BY T24.T24_GAKUNEN"

        iRet = gf_GetRecordset(rs,w_sSQL)
        If iRet <> 0  Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_GetSikenKikan = False
            Exit Do
        End If

		'// 試験名称取得
		iRet = f_GetDisp_Data_Siken(w_sSikenName)
        If iRet <> 0  Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_GetSikenKikan = False
            Exit Do
        End If

		'===========================================
		'HTML書き出し
		'===========================================
		%>
		<table class=hyo border=1 width=420>
		<tr><th class="header" colspan="6"><%=w_sSikenName%>期間</th></tr>
		<tr>

		<%
		i = 1
		For i = 1 To 5
			If rs.EOF = False Then
				If i=cint(rs("T24_GAKUNEN")) Then%>
					<th class="header" width="33"  align="center"><font size=2><%=i%>年</font></th>
					<td class="detail" width="100" align="center"><font size=2><%=right(rs("T24_JISSI_KAISI"),5) & "～" & right(rs("T24_JISSI_SYURYO"),5) %></font></td>
					<%
					rs.MoveNext
				Else%>
					<th class="header" width="33" align="center"><font size=2><%=i%>年</font></th>
					<td class="detail" width="100" align="center" ><font size=2>―</font></td>
				<%
				End If

			Else%>
				<th class="header" width="33" align="center"><font size=2><%=i%>年</font></th>
				<td class="detail" width="100" align="center" ><font size=2>―</font></td>
				<%
			End If

			If i=3 Then%>
				</tr><tr>
			<%
			End If
		Next
		%>

		<td class="detail" width="100" align="center" colspan="2"></td>
		</tr>
		</table>
		<br>
		<%
		'===========================================

        Exit Do
    Loop

    gf_closeObject(rs)

'// LABEL_f_ChkDate_END
End Function

'********************************************************************************
'*  [機能]  表示項目(試験)を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDisp_Data_Siken(p_sSikenName)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Siken = 1

    Do
        '試験マスタよりデータを取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI, "
        w_sSql = w_sSql & vbCrLf & "  M27_SIKEN.M27_SIKENMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN ,M27_SIKEN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      M01_KUBUN.M01_SYOBUNRUI_CD = M27_SIKEN.M27_SIKEN_KBN(+)"
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_NENDO = M27_SIKEN.M27_NENDO(+)"
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND M27_SIKEN.M27_SIKEN_CD='" & m_iSikenCd & "'"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetDisp_Data_Siken = 99
            Exit Do
        End If

        p_sSikenName = ""
        If rs.EOF = False Then
            p_sSikenName = rs("M01_SYOBUNRUIMEI")

            '//実力試験または、追試が選択された場合試験詳細名も追加表示
            If cint(m_sSikenCd) <> 0  Then
                p_sSikenName = p_sSikenName & " (" 
                p_sSikenName = p_sSikenName & rs("M27_SIKENMEI")
                p_sSikenName = p_sSikenName & " )" 
            End If

        End If

        '//正常終了
        f_GetDisp_Data_Siken = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    Dim w_pageBar           'ページBAR表示用
    Dim w_iRecordCnt        '//レコードセットカウント
    Dim w_iCnt

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    m_iCnt  = 1

    'ページBAR表示
    Call gs_pageBar(m_Rs,m_iPage,m_iDsp,w_pageBar)

%>

<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    //************************************************************
    //  [機能]  登録画面を表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_NewClick(){
    
        document.frm.action="syousai.asp";
        document.frm.target = "main";
        document.frm.txtMode.value = "BLANK";
        document.frm.submit();
        
    }
    
    //************************************************************
    //  [機能]  選択された試験日の更新画面を表示する
    //  [引数]  p_sCode     :選択された連番（教官予定）
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ListClick(p_day,p_sCode){

        document.frm.action="syousai.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtRenban.value = p_sCode;
        //document.frm.cmbJissiDate.value = p_day;

        document.frm.txtKeyYoteibi.value = p_day;

        document.frm.txtMode.value = "DISP";
        document.frm.submit();
        
    }
    
    //************************************************************
    //  [機能]  削除ボタンが押されたとき（確認用）
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ChkDelClick(){
        if( confirm("予定を削除します。") == true ){
            document.frm.action="main.asp";
            document.frm.target = "main";
            document.frm.txtMode.value = "CHK";
            document.frm.submit();
        }
    }
    
    //************************************************************
    //  [機能]  削除ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_DelClick(){
        if( confirm("予定を削除します。") == true ){
            document.frm.action="main.asp";
            document.frm.target = "main";
            document.frm.txtMode.value = "DEL";
            document.frm.submit();
        }
    }

    
    //************************************************************
    //  [機能]  戻るボタンが押されたとき（確認用）
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //  [作成日] 
    //************************************************************
    function f_ChkBackClick(){
    
        document.frm.action="main.asp?<%Call s_ShowRequest()%>";
        document.frm.target="main";
        document.frm.submit();
        
        
    }
    
    //************************************************************
    //  [機能]  戻るボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //  [作成日] 
    //************************************************************
    function f_BackClick(){
    
        document.frm.action="default.asp";
        document.frm.target="_parent"
        document.frm.submit();
        
        
    }
    
    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="main.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPage.value = p_iPage;
        document.frm.submit();
    
    }
    

	//-->
	</SCRIPT>
	</head>

	<body>
	<center>
	<form name="frm" Method="POST">
	    <input type="hidden" name="txtBiko" value="<%=m_sYoteiMei%>">
	    <input type="hidden" name="txtMode" value="<%=m_sMode%>">
	    <input type="hidden" name="txtRenban">
	    <input type="hidden" name="txtSikenKbn" value="<%=m_iSikenKbn%>">
	    <input type="hidden" name="txtSikenCd" value="<%=m_iSikenCd%>">
	    <input type="hidden" name="txtPage" value="<%= m_iPage %>">
	<table border=0 width="<%=C_TABLE_WIDTH%>">
		<tr>
		<td align="center">
		<%
		if m_sMode = "CHK" Then
		    response.write "以下の内容を削除します。<br>" & chr(13)
		else

			'//試験期間を表示
			Call f_GetSikenKikan()

		    response.write w_pageBar
		End if
		%>


		<%Call showTableHead()%>
		<%
		    Do Until m_Rs.EOF
		%>
		<%Call s_GetParam()%>
		<%
		if m_sMode = "CHK" Then
		    Call ShowTableChk()
		else
		    Call ShowTable()
		End if
		%>
		<%m_iCnt = m_iCnt + 1%>
		<%
		            m_Rs.MoveNext

		            If w_iCnt >= C_PAGE_LINE Then
		                Exit Do
		            Else
		                w_iCnt = w_iCnt + 1
		            End If

		    Loop

		'//削除確認画面で、ない場合
		if m_sMode <> "CHK" Then
		%>
		    <tr>
		    <td colspan=6 align=right bgcolor=#9999BD><input class=button type=button value="×削除" Onclick="f_ChkDelClick()"></td>
		    </tr>
		<%
		Else%>
			<!--削除確認画面表示時値保持用-->
			<input type="hidden" name="chkDel" value="<%=Request("chkDel")%>">
		<%
		End If
		%>
	</table>

	<br>
	<%
	if m_sMode = "CHK" Then
	    response.write "実行しますか？<br>" & chr(13)
	else
	    response.write w_pageBar
	End if
	%>
	<%Call showButton()%>
	</td>
	</tr>
	</table>

	<input type="hidden" name="txtKeyYoteibi" value="">

	</form>

	</center>
	</body>

	</html>
<%
    '---------- HTML END   ----------
End Sub

Sub showTableHead()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

if m_sMode = "CHK" Then
%>
<table border="1" width="<%=C_TABLE_WIDTH%>" CLASS="hyo">
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="60%" ALIGN=center>
<tr>
    <th CLASS="Header">日付</th>
    <th CLASS="Header">予定時間</th>
    <th CLASS="Header">理由</th>
</tr>
<%
else
%>
<table border="1" width="<%=C_TABLE_WIDTH%>" CLASS="hyo">
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="46%" ALIGN=center>
    <COLGROUP WIDTH="6%" ALIGN=center>
    <COLGROUP WIDTH="8%" ALIGN=center>
<tr>
    <th CLASS="header">日付</th>
    <th CLASS="header">予定時間</th>
    <th CLASS="header">理由</th>
    <th CLASS="header">修正</th>
    <th CLASS="header">削除</th>
</tr>
<%
end if
    '---------- HTML END   ----------
End Sub


Sub showTable()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
<tr>

	<%

	'//入力されている時間より終日かどうかを判別
	'//C_MIN_TIME = "00:00"(最小時刻),C_MAX_TIME = "23:55"(最大時刻)
	If m_dtYoteiKaisi = C_MIN_TIME And m_dtYoteiSyuryo = C_MAX_TIME Then
		w_sStr = "終日"
	Else
		w_sStr = m_dtYoteiKaisi & "-" & m_dtYoteiSyuryo
	End If
	%>

    <td class=detail>
    <%=m_iMonth%>/<%=m_iDay%>(<%=m_sYobi%>)
    </td>
    <td class=detail>
    <%=w_sStr%>
    </td>
    <td class=detail align="left"><%=m_sYoteiMei%></td>
    <td class=detail align="center"><input type="button" value=">>" onClick="f_ListClick('<%=m_iYoteiBi%>','<%=m_iRenban%>');return false;" class=button></td>
    <td class=detail align="center"><input type="checkbox" name="chkDel" value="<%=m_iYoteiBi & "_" & m_iRenban%>"></td>

</tr>
<%
End Sub

'********************************************************************************
'*  [機能]  確認画面を表示する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub ShowTableChk()

	%>
	<tr>
	    <td class=detail>
	    <%=m_iMonth%>/<%=m_iDay%>(<%=m_sYobi%>)
	    </td>
	    <td class=detail>
	    <%=m_dtYoteiKaisi%>-<%=m_dtYoteiSyuryo%>
	    </td>
	    <td class=detail align="left"><%=m_sYoteiMei%></td>
	    <input type="hidden" name="chkDelRenban<%=m_iCnt%>" value="<%=m_iRenban%>">
	    
	</tr>
	<%

End Sub


Sub showButton()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
if m_sMode = "CHK" Then
%>
<table border="0" width="40%">
<COLGROUP span=2 WIDTH="50%" ALIGN=center>
<tr>
    <td><input type="button" value="　削　除　" onClick="javascript:f_DelClick();return false;" class=button></td>
    <td><input type="button" value="キャンセル" onClick="javascript:f_ChkBackClick();return false;" class=button></td>
</tr>
<input type="hidden" name="txtCnt" value="<%=m_iCnt%>">

</table>
<%
else
end if

End Sub

Sub showPage_NoData(p_msg)
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
    <html>
    <head>
	<link rel=stylesheet href="../../common/style.css" type=text/css>
    </head>
    <body>
    <center>
		<br><br><br>
		<span class="msg"><%=p_msg%></span>
    </center>
    </body>
    </html>

<%
    '---------- HTML END   ----------
End Sub
%>