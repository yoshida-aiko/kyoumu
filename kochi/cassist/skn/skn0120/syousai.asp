<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 教官予定マスタ
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0120/syossai.asp
' 機      能: 下ページ 教官予定マスタの詳細表示を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）

'           txtSikenKbn      :試験区分
'           txtSikenCd      :試験コード
'           txtMode         :動作モード
'                           BLANK   :登録のため全項目空白で表示
'                           INSERT  :登録処理を実行
'                           UPDATE  :更新処理を実行
'                           DISP    :指定された区分のデータを表示
'
'           cmbJissiDate    :日付
'           cmbRiyu         :理由
'           txtKaisiH       :開始時刻（時）
'           txtKaisiN       :開始時刻（分）
'           txtSyuRyoH      :終了時刻（時）
'           txtSyuryoN      :終了時刻（分）
'
'           txtRenban       :連番
'           txtPage         :表示頁数
'
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSikenKbn      :試験区分（戻るとき）
'           txtSikenCd      :試験コード
'           txtMode         :動作モード（戻るとき）
'                           BLANK   :全項目空白で表示
'
'           cmbJissiDate    :日付
'           cmbRiyu         :理由
'           txtKaisiH       :開始時刻（時）
'           txtKaisiN       :開始時刻（分）
'           txtSyuryoH      :終了時刻（時）
'           txtSyuryoN      :終了時刻（分）
'
'           txtRenban       :連番
'           txtPage         :表示頁数
'
' 説      明:
'           ■初期表示
'               検索条件にかなう試験中予定を表示
'           ■更新ボタンクリック時
'               指定した条件にかなう試験予定を更新させる
'           ■登録ボタンクリック時
'               試験予定入力を表示させて、登録させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/16 高丘 知央
' 変      更: 2001/06/26 根本
' 変      更: 2001/07/27 伊藤公子 M40_CALENDER削除の為変更
' 変      更: 2001/08/03 伊藤公子 試験期間を表示するよう修正
' 変      更: 2001/08/03 伊藤公子 日付の入力をFromToで範囲入力できる用に変更
'                                 全日チェックボックスを追加
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙｺﾝｽﾄ /////////////////////////////
	CONST C_MIN_TIME = "00:00"		'//最小時刻
	CONST C_MAX_TIME = "23:55"		'//最大時刻(時間)
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    '取得したデータを持つ変数
    Public  m_iSikenKbn     ':試験区分
    Public  m_iSikenCd      ':試験コード
    Public  m_sMode         ':動作モード
    Public  m_iKyokanCd     ':教官コード
    Public  m_iSyoriNen     ':処理年度
    Public  m_iRenban       ':連番
    Public  m_iJissiDate    ':日付
    Public  m_iJissiDateE    ':終了日付
    Public  m_iJissiKaisi   ':試験実施開始日
    Public  m_iJissiSyuryo  ':試験実施終了日
    Public  m_sBiko         ':備考
    Public  m_iKaisiH       ':開始時刻（時）
    Public  m_iKaisiN       ':開始時刻（分）
    Public  m_iSyuryoH      ':終了時刻（時）
    Public  m_iSyuryoN      ':終了時刻（分）
    Public  m_iKaisi       ':開始時刻
    Public  m_iSyuryo      ':終了時刻
    Public  m_Rs            'recordset
    Public  m_iRenbanCount  '連番数
    Public  m_iKikanWhere
    Public  m_iPage     ':表示済表示頁数（自分自身から受け取る引数）

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bBack             '// 正常終了時ﾌﾗｸﾞ
    Public  m_bMsgFlg           '// ﾒｯｾｰｼﾞﾌﾗｸﾞ（ｷｰ違反などのｴﾗｰの場合にﾀﾞｲｱﾛｸﾞを表示するためのﾌﾗｸﾞ）
    Public  m_sDebugStr         '// 以下ﾃﾞﾊﾞｯｸﾞ用
    Public  m_sMsg              '// ﾒｯｾｰｼﾞ用

    Public  m_sMinDate,m_sMaxDate	'//試験期間の最小日付、最大日付

    'ページ関係
    Public  m_iMax          ':最大ページ
    Public  m_iDsp                      '// 一覧表示行数

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
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="試験監督免除登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bBack = False
    m_bMsgFlg = False
    m_sMode = Request("txtMode")

    w_iRet = 0

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    Do
    
        '// 値の初期化
        Call s_SetBlank()

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

        Call s_setParam()
                
			'===============================
			'//期間データの取得
			'===============================
	        w_iRet = f_Nyuryokudate()
			If w_iRet = 1 Then
				    '// 終了処理
				    Call gs_CloseDatabase()
					response.Redirect "default.asp?txtMode=no&txtSikenKbn="&m_iSikenKbn&""
					response.end
				Exit Do
			End If
			If w_iRet <> 0 Then 
				m_bErrFlg = True
				Exit Do
			End If

        '理由コンボに関するWHEREを作成する（更新時に必要）
        Call c_MakeRiyuWhere() 

        '// 各ﾓｰﾄﾞにより処理を分ける
        Select Case m_sMode
            '// ■■■データ表示（登録用フォーム）
            Case "BLANK"
                m_sDebugStr = "BLANK"

                w_iRet = f_GetJissiDate()
                If w_iRet <> 0 Then
                    'エラー処理
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If
                
                '期間コンボに関するWHEREを作成する（更新時に必要）
                Call s_MakeKikanWhere()
                
                ' ページを表示
                Call showPage()

            '// ■■■データ表示（更新用フォーム）
            Case "DISP"
                '指定データ取得
                m_sDebugStr = "DISP"

'                Call s_setParam()
                
                w_iRet = f_GetData()
                If w_iRet <> 0 Then
                    'エラー処理
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If

                w_iRet = f_GetJissiDate()
                If w_iRet <> 0 Then
                    'エラー処理
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If
                
                Call s_MakeRiyuWhere()
                
                '期間コンボに関するWHEREを作成する（更新時に必要）
                Call s_MakeKikanWhere()
                
                'ページを表示
                Call showPage()

            '// ■■■データ追加
            Case "INSERT"
                m_sDebugStr = "INSERT"
                
'                Call s_SetParam()
                
                Call s_SetParamDate()

                w_iRet = f_GetJissiDate()
                If w_iRet <> 0 Then
                    'エラー処理
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If

				'//登録処理
				w_iRet = f_Insert()
                If w_iRet = 1 Then
                    '時刻が重複している場合
                    m_bMsgFlg = True
                    m_bBack = False
                    Call showPage()
                    Exit Do
				Else
                    If w_iRet <> 0 Then
                        'エラー処理
                        m_bErrFlg = True
                        m_sErrMsg = m_sMsg
                        Exit Do
                    Else
                        '正常終了で一覧表画面に戻る
                        m_bBack = True
                        Call showPage()

                    End If
                End If

            '// ■■■データ更新
            Case "UPDATE"
                m_sDebugStr = "UPDATE"
                
'                Call s_SetParam()
                
                Call s_SetParamDate()

                w_iRet = f_GetJissiDate()
                If w_iRet <> 0 Then
                    'エラー処理
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If


				'//入力された情報がすでに登録されてないかチェックする
				w_bRet = f_ChkDate(m_iJissiDate,m_iRenban)
                if w_bRet = True Then
                    w_iRet = f_UpdateData()
                    If w_iRet <> 0 Then
                        m_bErrFlg = True
                        m_sErrMsg = m_sMsg
                        Exit Do
                    Else
                        '正常終了で一覧表画面に戻る
                        m_bBack = True
                        
                        '期間コンボに関するWHEREを作成する（更新時に必要）
                        Call s_MakeKikanWhere()
                        
                        Call showPage()
                    End If
                else
                        'エラー処理
                        '時刻が重複している場合
                        m_bMsgFlg = True
                        m_bBack = False
                        Call showPage()
                        Exit Do
                
                end if
            '// ■■■その他（エラー）
            Case Else
                m_sDebugStr = "ETC"
                m_bErrFlg = True
                Call gs_SetErrMsg("処理モードが設定されていません(ｼｽﾃﾑｴﾗｰ)")
                Exit Do
        End Select
            
        '// 正常終了
        Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        if m_sErrMsg <> "" Then
            w_sMsg = m_sErrMsg
        else
            w_sMsg = gf_GetErrMsg()
        end if
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gs_CloseDatabase()

    On Error Goto 0
    Err.Clear
            
End Sub

'********************************************************************************
'*  [機能]  全項目を空白に初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetBlank()

    m_iSikenKbn = ""
    m_iSikenCd = ""
    m_iRenban = ""
    m_iJissiDate = ""
    m_sBiko = ""
    m_iKaisiH = ""
    m_iKaisiN = ""
    m_iSyuryoH = ""
    m_iSyuryoN = ""
    m_iKaisi = ""
    m_iSyuryo = ""
    m_iRenbanCount = ""
    m_iSyoriNen = ""
    m_iPage = ""

End Sub

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


'********************************************************************************
'*  [機能]  試験実施開始日・終了日をセット
'*  [引数]  なし
'*  [戻値]  0:情報取得成功、1:ﾚｺｰﾄﾞなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetJissiDate()

    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    f_GetJissiDate = 0
    m_iJissiKaisi = ""
    m_iJissiSyuryo = ""

    Do 
        '// 試験日程ﾚｺｰﾄﾞｾｯﾄを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & " T24_JISSI_KAISI"
        w_sSQL = w_sSQL & vbCrLf & " ,T24_JISSI_SYURYO"
        w_sSQL = w_sSQL & vbCrLf & " FROM T24_SIKEN_NITTEI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE T24_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND T24_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & vbCrLf & " AND T24_SIKEN_CD = '" & m_iSikenCd & "'"

'response.write w_sSQL & "<br>"

        Set w_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            'm_sMsg = "試験日程の取得に失敗しました"
            m_sMsg = Err.description
            f_GetJissiDate = 99
            Exit Do
        End If
    
        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            m_sMsg = "試験日程が登録されていません"
            f_GetJissiDate = 1
            Exit Do
        End If

        m_iJissiKaisi = w_Rs("T24_JISSI_KAISI")      ':試験実施開始日
        m_iJissiSyuryo = w_Rs("T24_JISSI_SYURYO")    ':試験実施終了日
        
        Exit Do

    Loop

    gf_closeObject(w_Rs)
End Function

'********************************************************************************
'*  [機能]  情報取得処理を行う（データ更新時表示に使用）
'*  [引数]  なし
'*  [戻値]  0:情報取得成功、1:ﾚｺｰﾄﾞなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetData()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim i                   '// ｶｳﾝﾀ
    
    On Error Resume Next
    Err.Clear
    f_GetData = 0

    Do 
        '// 教官予定ﾚｺｰﾄﾞｾｯﾄを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & " T25_NENDO"
        w_sSQL = w_sSQL & " ,T25_SIKEN_KBN"
        w_sSQL = w_sSQL & " ,T25_SIKEN_CD"
        w_sSQL = w_sSQL & " ,T25_KYOKAN"
        w_sSQL = w_sSQL & " ,T25_YOTEIBI"
        w_sSQL = w_sSQL & " ,T25_RENBAN"
        w_sSQL = w_sSQL & " ,T25_YOTEI_KAISI"
        w_sSQL = w_sSQL & " ,T25_YOTEI_SYURYO"
        w_sSQL = w_sSQL & " ,T25_BIKO"
        w_sSQL = w_sSQL & " FROM T25_KYOKAN_YOTEI "
        w_sSQL = w_sSQL & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        'w_sSQL = w_sSQL & " AND T25_YOTEIBI = '" & m_iJissiDate & "'"
        w_sSQL = w_sSQL & " AND T25_YOTEIBI = '" & gf_YYYY_MM_DD(Request("txtKeyYoteibi"),"/") & "'"
        w_sSQL = w_sSQL & " AND T25_RENBAN = " & m_iRenban

'response.write "w_sSQL = " &w_sSQL & "<BR>"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
       If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_sMsg = Err.description
            f_GetData = 99
            Exit Do 'GOTO LABEL_f_GetData_END
        End If

        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            f_GetData = 1
            m_sMsg = "教官予定の対象レコードがありません"
            Exit Do
        End If

        '// 取得した値をｸﾞﾛｰﾊﾞﾙ変数に格納
        m_iSikenKbn = w_Rs("T25_SIKEN_KBN")            ':試験区分
        m_iSikenCd = w_Rs("T25_SIKEN_CD")              ':試験コード
        m_iRenban = w_Rs("T25_RENBAN")                 ':連番
        m_iJissiDate = w_Rs("T25_YOTEIBI")             ':日付
        m_sBiko = w_Rs("T25_BIKO")                     ':理由
        m_iKaisiH = Left(w_Rs("T25_YOTEI_KAISI"),2)    ':開始時刻（時）
        m_iKaisiN = Mid(w_Rs("T25_YOTEI_KAISI"),4,2)   ':開始時刻（分）
        m_iSyuryoH = Left(w_Rs("T25_YOTEI_SYURYO"),2)  ':終了時刻（時）
        m_iSyuryoN = Mid(w_Rs("T25_YOTEI_SYURYO"),4,2) ':終了時刻（分）
        m_iKaisi = w_Rs("T25_YOTEI_KAISI")             ':開始時刻
        m_iSyuryo = w_Rs("T25_YOTEI_SYURYO")           ':終了時刻

        Exit Do

    Loop

    gf_closeObject(w_Rs)
    
'// LABEL_f_GetData_END
End Function

'********************************************************************************
'*  [機能]  登録処理
'*  [引数]  なし
'*  [戻値]  0:成功,1:重複あり、99:失敗
'*  [説明]  
'********************************************************************************
Function f_Insert()

	f_Insert = 0

	Do

		'//実施日付を取得
		m_iJissiDate = Request("cmbJissiDate")  ':日付
		m_iJissiDateE = Request("cmbJissiDateE")  ':終了日付

		If m_iJissiDateE <> "" Then
			iMax = DateDiff("d",m_iJissiDate,m_iJissiDateE)+1
		Else
			iMax = 1
		End If 

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

		'//開始日から終了日まで、いちにちずつINSERTする
		For i = 1 To imax
			w_sJissiBi = FormatDateTime(DateAdd("d",i-1,m_iJissiDate))

			'//連番取得
			w_iRet = f_GetCountRenban(w_sJissiBi,w_RenBan)
			If w_iRet <> 0 Then
			    'エラー処理
				f_Insert = 99
			    Exit Do
			End If

			'//重複チェック
			w_bRet = f_ChkDate(w_sJissiBi,w_RenBan)
			if w_bRet = True Then

				'//重複していない時登録処理
			    w_iRet = f_InsertData(w_sJissiBi,w_RenBan)
			    If w_iRet <> 0 Then
					'//ﾛｰﾙﾊﾞｯｸ
					Call gs_RollbackTrans()
			        'エラー処理
					f_Insert = 99
			        Exit Do
			    Else

			    End If
			else
		        '時刻が重複している場合
				f_Insert = 1
		        Exit Do
			end if

		Next

	   '//正常終了時、ｺﾐｯﾄ
	   Call gs_CommitTrans()

		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  連番をカウントする（新規登録時に使用）
'*  [引数]  p_JissiBi
'*  [戻値]  0:情報取得成功、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetCountRenban(p_JissiBi,p_RenBan)
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear

    f_GetCountRenban = 0
	p_RenBan = 0

    Do 
        '// 教官予定ﾚｺｰﾄﾞｾｯﾄを取得（連番Max値）
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & " MAX("
        w_sSQL = w_sSQL & vbCrLf & " T25_RENBAN "
        w_sSQL = w_sSQL & vbCrLf & ")"
        w_sSQL = w_sSQL & vbCrLf & " AS T25_MAXRENBAN"
        w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_YOTEIBI = '" & p_JissiBi & "'"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_sMsg = Err.description
            f_GetCountRenban = 99
            Exit Do
        End If

		If ISNULL(w_Rs("T25_MAXRENBAN")) = False Then
			p_RenBan = cint(w_Rs("T25_MAXRENBAN")) + 1
		Else
			p_RenBan = 0
		End If

    	Exit Do

    Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
    gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  実施日が、指定試験期間の期間内に設定されたか
'*  [引数]  なし
'*  [戻値]  
'*  [説明]  
'********************************************************************************
'Function f_ChkDate()
Function f_ChkDate(p_sJissiBi,p_RenBan)


    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    f_ChkDate = True

    Do

        '// 教官予定ﾚｺｰﾄﾞｾｯﾄを取得
        w_sSQL2 = ""
        w_sSQL2 = w_sSQL2 & "SELECT"
        w_sSQL2 = w_sSQL2 & " T25_YOTEI_KAISI"
        w_sSQL2 = w_sSQL2 & " ,T25_YOTEI_SYURYO"
        w_sSQL2 = w_sSQL2 & " ,T25_RENBAN"
        w_sSQL2 = w_sSQL2 & " FROM T25_KYOKAN_YOTEI "
        w_sSQL2 = w_sSQL2 & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL2 = w_sSQL2 & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL2 = w_sSQL2 & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL2 = w_sSQL2 & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL2 = w_sSQL2 & " AND T25_YOTEIBI = '" & p_sJissiBi & "'"
        w_sSQL2 = w_sSQL2 & " AND T25_RENBAN <> " & p_RenBan

        w_iRet2 = gf_GetRecordset(w_Rs2, w_sSQL2)
        If w_iRet2 <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_sMsg = Err.description
            f_ChkDate = False
            Exit Do
        End If

        If w_Rs2.EOF Then
            '対象ﾚｺｰﾄﾞなし
            f_ChkDate = True
            Exit Do
        End If

        Do Until w_Rs2.EOF
            If CDate(m_iKaisi) < CDate(w_Rs2("T25_YOTEI_SYURYO")) Then
                if CDate(m_iSyuryo) <= CDate(w_Rs2("T25_YOTEI_KAISI")) Then
                else
                    m_sMsg = "すでに登録されている予定時間と重複しています\n" & m_iKaisi & "-" & m_iSyuryo & "<" & w_Rs2("T25_YOTEI_KAISI") & "-" & w_Rs2("T25_YOTEI_SYURYO")
                    f_ChkDate = False
                    exit Do
                end if
            else
                If CDate(m_iSyuryo) > CDate(w_Rs2("T25_YOTEI_KAISI")) Then
                    if CDate(m_iKaisi) >= CDate(w_Rs2("T25_YOTEI_SYURYO")) Then
                    else
                        m_sMsg = "すでに登録されている予定時間と予定時間が重複しています\n" & m_iKaisi & "-" & m_iSyuryo & "<" & w_Rs2("T25_YOTEI_KAISI") & "-" & w_Rs2("T25_YOTEI_SYURYO")
                        f_ChkDate = False
                        exit Do
                    end if
                end if
            end if
	        w_Rs2.MoveNext
        Loop
    

        Exit Do
    Loop

    gf_closeObject(w_Rs2)

End Function

'********************************************************************************
'*  [機能]  登録処理を行う(教官予定)
'*  [引数]  なし
'*  [戻値]  0:更新成功、1:キー違反、99:失敗
'*  [説明]　
'********************************************************************************
'Function f_InsertData()
Function f_InsertData(p_sJissiBi,p_RenBan)

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文

    On Error Resume Next
    Err.Clear
    f_InsertData = 1
    
    Do

        '// 指定されたﾃﾞｰﾀ挿入
        w_sSQL = ""
        w_sSQL = w_sSQL & "INSERT INTO T25_KYOKAN_YOTEI"
        w_sSQL = w_sSQL & " ("
        w_sSQL = w_sSQL & "  T25_NENDO"
        w_sSQL = w_sSQL & ", T25_SIKEN_KBN"
        w_sSQL = w_sSQL & ", T25_SIKEN_CD"
        w_sSQL = w_sSQL & ", T25_KYOKAN"
        w_sSQL = w_sSQL & ", T25_YOTEIBI"
        w_sSQL = w_sSQL & ", T25_RENBAN"
        w_sSQL = w_sSQL & ", T25_YOTEI_KAISI"
        w_sSQL = w_sSQL & ", T25_YOTEI_SYURYO"
        w_sSQL = w_sSQL & ", T25_BIKO"
        w_sSQL = w_sSQL & ", T25_INS_DATE"
        w_sSQL = w_sSQL & ", T25_INS_USER"
        w_sSQL = w_sSQL & " ) VALUES ("
        w_sSQL = w_sSQL & m_iSyoriNen
        w_sSQL = w_sSQL & "," & m_iSikenKbn
        w_sSQL = w_sSQL & ", '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & ", '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & ", '" & gf_YYYY_MM_DD(p_sJissiBi,"/") & "'"
        w_sSQL = w_sSQL & "," & p_RenBan
        w_sSQL = w_sSQL & ", '" & m_iKaisi & "'"
        w_sSQL = w_sSQL & ", '" & m_iSyuryo & "'"
        w_sSQL = w_sSQL & ", '" & m_sBiko & "'"
        w_sSQL = w_sSQL & ", '" & gf_YYYY_MM_DD(date(),"/") & "'"
        w_sSQL = w_sSQL & ", '" & Session("LOGIN_ID") & "'"
        w_sSQL = w_sSQL & ")"

        w_iRet = gf_ExecuteSQL(w_sSQL)
        If w_iRet <> 0 Then
            '挿入処理失敗
            If w_iRet = C_ERR_DATA_EXIST or w_iRet = C_ERR_DATA_EXIST2 Then
                m_sMsg = "登録に失敗しました"
                Exit Do
            Else
                m_sMsg = "登録エラーです"
                f_InsertData = 99
                Exit Do
            End If
        End If

        '//正常終了
        f_InsertData = 0
        Exit Do
    Loop

End function

'********************************************************************************
'*  [機能]  更新処理を行う
'*  [引数]  なし
'*  [戻値]  0:更新成功、1:キー違反、99:失敗
'*  [説明]  
'********************************************************************************
Function f_UpdateData()

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    
    On Error Resume Next
    Err.Clear

    f_UpdateData = 1
    
    Do 

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

        '// 指定されたﾃﾞｰﾀの存在確認
        '// UPDATE実行時にﾃﾞｰﾀが存在しない場合、エラーが発生しないため事前に確認する
        '// 教官予定テーブルﾚｺｰﾄﾞｾｯﾄを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & "SELECT T25_NENDO "
        w_sSQL = w_sSQL & vbCrLf & ", T25_SIKEN_KBN "
        w_sSQL = w_sSQL & vbCrLf & ", T25_SIKEN_CD "
        w_sSQL = w_sSQL & vbCrLf & ", T25_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & ", T25_RENBAN "
        w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_YOTEIBI = '" & gf_YYYY_MM_DD(request("txtKeyYoteibi"),"/") & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_RENBAN = " & m_iRenban

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            '//ﾛｰﾙﾊﾞｯｸ
            Call gs_RollbackTrans()
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_UpdateData = 99
            Exit Do
        End If

        If w_Rs.EOF Then
            '//ﾛｰﾙﾊﾞｯｸ
            Call gs_RollbackTrans()
            '対象ﾚｺｰﾄﾞなし
            m_sMsg = "対象レコードがありません"
            'f_UpdateData = 1
            Exit Do
        End If

        '// 指定されたﾃﾞｰﾀ更新
        w_sSQL = ""
        w_sSQL = w_sSQL & "UPDATE T25_KYOKAN_YOTEI"
        w_sSQL = w_sSQL & " SET "
        w_sSQL = w_sSQL & " T25_YOTEIBI = '" & gf_YYYY_MM_DD(m_iJissiDate,"/") & "'"    '//保留
        w_sSQL = w_sSQL & ", T25_YOTEI_KAISI = '" & m_iKaisi & "'"  '//保留
        w_sSQL = w_sSQL & ", T25_YOTEI_SYURYO = '" & m_iSyuryo & "'" '//保留
        w_sSQL = w_sSQL & ", T25_BIKO = '" & m_sBiko & "'"
        w_sSQL = w_sSQL & ", T25_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "'"
        w_sSQL = w_sSQL & ", T25_UPD_USER = '" & Session("LOGIN_ID") & "'"
        w_sSQL = w_sSQL & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & " AND T25_YOTEIBI = '" & gf_YYYY_MM_DD(request("txtKeyYoteibi"),"/") & "'"
        w_sSQL = w_sSQL & " AND T25_RENBAN = " & m_iRenban

        w_iRet = gf_ExecuteSQL(w_sSQL)
        If w_iRet <> 0 Then
            '//ﾛｰﾙﾊﾞｯｸ
            Call gs_RollbackTrans()
            '更新処理失敗
            If w_iRet = C_ERR_DATA_EXIST Or w_iRet = C_ERR_DATA_EXIST2 Then
                m_sMsg = "更新処理に失敗しました"
                'f_UpdateData = 1
                Exit Do
            Else
                m_sMsg = "更新エラーです"
                f_UpdateData = 99
                Exit Do
            End If
        End If
        
        '//ﾚｺｰﾄﾞｾｯﾄCLOSE
        Call gf_closeObject(w_Rs)
        
        '//ｺﾐｯﾄ
        Call gs_CommitTrans()
        
        '//正常終了
        f_UpdateData = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSikenKbn = Request("txtSikenKbn")            ':試験区分
    
    if Request("txtSikenCd") <> "" Then
        m_iSikenCd = Request("txtSikenCd")      ':試験コード
    else
        m_iSikenCd = 0
    end if

    m_sMode = Request("txtMode")                    ':動作モード
    m_iSyoriNen = Session("NENDO")                  ':処理年度
    m_iKyokanCd = Session("KYOKAN_CD")              ':教官コード
    m_iJissiDate = Request("cmbJissiDate")			':日付
    m_iRenban = Request("txtRenban")                ':連番

    if m_sMode = "UPDATE" or m_sMode = "INSERT" Then
'        m_iJissiDate = Request("cmbJissiDate")  ':日付
        m_iJissiDateE = Request("cmbJissiDateE")  ':終了日付
        m_sBiko = Request("txtBiko")            ':備考
        m_iKaisiH = Request("txtKaisiH")        ':開始時刻（時）
        m_iKaisiN = Request("txtKaisiN")        ':開始時刻（分）
        m_iSyuryoH = Request("txtSyuRyoH")      ':終了時刻（時）
        m_iSyuryoN = Request("txtSyuryoN")      ':終了時刻（分）
    end if

    '// BLANKの場合は行数ｸﾘｱ
    If Request("txtMode") = "BLANK" Then
        m_iPage = 1
    Else
        m_iPage = INT(Request("txtPage"))   ':表示済表示頁数（自分自身から受け取る引数）
    End If

End Sub

'********************************************************************************
'*  [機能]  引き渡されてきた時刻の値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParamDate()
    
    if m_sMode = "UPDATE" or m_sMode = "INSERT" Then
    
    Dim w_sKaisiH
    Dim w_sKaisiN
    Dim w_sSyuryoH
    Dim w_sSyuryoN
    
    w_sKaisiH = ""
    w_sKaisiN = ""
    w_sSyuryoH = ""
    w_sSyuryoN = ""
    
    w_sKaisiH = Trim(m_iKaisiH)
    w_sKaisiN = Trim(m_iKaisiN)
    w_sSyuryoH = Trim(m_iSyuryoH)
    w_sSyuryoN = Trim(m_iSyuryoN)
    
    if Len(w_sKaisiH) > 0 Then
         w_sKaisiH = Right("0" & w_sKaisiH,2)
    else
    end if
    
    if Len(w_sKaisiN) > 0 Then
        w_sKaisiN = Right("0" & w_sKaisiN,2)
    else
    end if

    if Len(w_sSyuryoH) > 0 Then
         w_sSyuryoH = Right("0" & w_sSyuryoH,2)
    else
    end if

    if Len(w_sSyuryoN) > 0 Then
        w_sSyuryoN = Right("0" & w_sSyuryoN,2)
    else
    end if
    
    m_iKaisi = w_sKaisiH & ":" & w_sKaisiN
    m_iSyuryo = w_sSyuryoH & ":" & w_sSyuryoN
    
    end if

End Sub

Sub s_MakeKikanWhere()
'********************************************************************************
'*  [機能]  予定コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

'-----2001/07/27 ito M40_CALENDER削除の為変更

'    m_iKikanWhere= ""
'    m_iKikanWhere = m_iKikanWhere & " M40_DATE >= '" & m_iJissiKaisi & "' "
'    m_iKikanWhere = m_iKikanWhere & " AND M40_DATE <= '" & m_iJissiSyuryo & "' "

    m_iKikanWhere= ""
    m_iKikanWhere = m_iKikanWhere & " T32_HIDUKE >= '" & m_iJissiKaisi & "' "
    m_iKikanWhere = m_iKikanWhere & " AND T32_HIDUKE <= '" & m_iJissiSyuryo & "' "
    m_iKikanWhere = m_iKikanWhere & " GROUP BY T32_HIDUKE"

'response.write m_iKikanWhere & "<BR>"

End Sub

'********************************************************************************
'*  [機能]  学年ごとの試験期間を取得
'*  [引数]  なし
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Function f_GetSikenKikan()

    Dim rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim iRet              '// 戻り値
    Dim w_sSQL              '// SQL文

    On Error Resume Next
    Err.Clear
    f_GetSikenKikan = false

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

        iRet = gf_GetRecordset(rs,w_sSql)
        If iRet <> 0  Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSikenKikan = False
            Exit Do
        End If

		'// 試験名称取得
		iRet = f_GetDisp_Data_Siken(w_sSikenName)
        If iRet <> 0  Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSikenKikan = false
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

		f_GetSikenKikan = True
        Exit Do
    Loop

    gf_closeObject(rs)

'// LABEL_f_ChkDate_END
End Function

'********************************************************************************
'*  [機能]  学年ごとの試験期間を取得
'*  [引数]  なし
'*  [戻値]  p_sMinDate:試験期間の最小日付
'*          p_sMaxDate:試験期間の最大日付
'*  [説明]  
'********************************************************************************
Function f_GetKikanLimit(p_sMinDate,p_sMaxDate)

    Dim rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim iRet              '// 戻り値
    Dim w_sSQL              '// SQL文

    On Error Resume Next
    Err.Clear

    f_GetKikanLimit = True
	p_sMinDate = ""
	p_sMaxDate = ""

    Do

        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  MIN(T24.T24_JISSI_KAISI) AS MIN_JISSI_KAISI"
        w_sSql = w_sSql & vbCrLf & "  ,MAX(T24.T24_JISSI_SYURYO) AS MAX_JISSI_SYURYO"
        w_sSql = w_sSql & vbCrLf & " FROM T24_SIKEN_NITTEI T24"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      T24.T24_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND T24.T24_SIKEN_KBN= " & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND T24.T24_SIKEN_CD='" & m_iSikenCd & "'"
        w_sSql = w_sSql & vbCrLf & " ORDER BY T24.T24_GAKUNEN"

        iRet = gf_GetRecordset(rs,w_sSql)
        If iRet <> 0  Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_GetKikanLimit = False
            Exit Do
        End If

		If rs.EOF = false Then
			p_sMinDate = rs("MIN_JISSI_KAISI")
			p_sMaxDate = rs("MAX_JISSI_SYURYO")
		End If

        Exit Do
    Loop

	'//終了処理
    gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  表示項目(試験)を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDisp_Data_Siken(p_sSikenName)
    Dim iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Siken = 1

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


%>
<html>

<head>
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

        <%If m_bBack = True Then%>
           // ﾓｰﾄﾞにBLANKを設定し、一覧表ﾍﾟｰｼﾞに戻る
            document.frm.action="default.asp";
            document.frm.target="<%=C_MAIN_FRAME%>";
            document.frm.txtMode.value = "Search";
            document.frm.submit();

        <%Else%>
            // ｴﾗｰの場合、ﾒｯｾｰｼﾞを表示
            <%If m_bMsgFlg = True Then%>
                window.alert("<%=m_sMsg%>");
            <%End If%>
            // 初期ﾌｫｰｶｽ
        <%End If%>
    }

    //************************************************************
    //  [機能]  登録・更新処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_SaveClick() {
        var w_iRet;
        // 入力値のﾁｪｯｸ
        w_iRet = f_CheckData();

        if( w_iRet == 0 ){
            if( confirm("<%=C_TOUROKU_KAKUNIN%>") == true ){

				//全日の場合、時刻を変換
				if(document.frm.txtDayAll.checked==true){
					document.frm.txtKaisiH.readOnly = true;
					document.frm.txtKaisiH.value = "00";

					document.frm.txtKaisiN.readOnly = true;
					document.frm.txtKaisiN.value = "00";

					document.frm.txtSyuryoH.readOnly = true;
					document.frm.txtSyuryoH.value = "23";

					document.frm.txtSyuryoN.readOnly = true;
					document.frm.txtSyuryoN.value = "55";
				};

            <%If m_sMode = "BLANK" Or m_sMode = "INSERT" Then%>
                // ﾓｰﾄﾞにINSERTを設定し、本ﾍﾟｰｼﾞを呼ぶ
                document.frm.action="syousai.asp";
                document.frm.txtMode.value = "INSERT";
                document.frm.submit();
            <%ElseIf m_sMode = "DISP" Or m_sMode = "UPDATE" Then%>
                // ﾓｰﾄﾞにUPDATEを設定し、本ﾍﾟｰｼﾞを呼ぶ
                document.frm.action="syousai.asp";
                document.frm.txtMode.value = "UPDATE";
                document.frm.submit();
            <%End If%>
            }
        }
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
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "BLANK";
        document.frm.submit();

    }
    
    //************************************************************
    //  [機能]  入力値のﾁｪｯｸ
    //  [引数]  なし
    //  [戻値]  0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
    //  [説明]  入力値のNULLﾁｪｯｸ、数字ﾁｪｯｸ、桁数ﾁｪｯｸを行う
    //          引渡ﾃﾞｰﾀ用にﾃﾞｰﾀを加工する必要がある場合には加工を行う
    //************************************************************
    function f_CheckData() {

        // ■■■日付ﾁｪｯｸ■■■
        // ■ 実施日
        if(f_Trim(document.frm.cmbJissiDate.value) == "" ){
                window.alert("日付が入力されていません");
                document.frm.cmbJissiDate.focus();
                return 1;
		}else{
            if( IsDate(document.frm.cmbJissiDate.value) != 0 ){
                window.alert("日付の入力が不正です");
                document.frm.cmbJissiDate.focus();
                return 1;
            }
        }

		<%'//■■■試験期間内かどうかチェック■■■%>
		<%
		'//試験期間の最小日付と、最大日付を取得する
		w_bRet = f_GetKikanLimit(m_sMinDate,m_sMaxDate)
		If w_bRet = False Then
			m_bErrFlg = True
		End If
		%>

        var MinDate = new Date("<%=m_sMinDate%>");
        var MaxDate = new Date("<%=m_sMaxDate%>");
        var vl = new Date(document.frm.cmbJissiDate.value);
		if( vl< MinDate ){
            window.alert("試験期間外の日付は入力できません");
            document.frm.cmbJissiDate.focus();
            return 1;
		}else{
			if( vl > MaxDate ){
	            window.alert("試験期間外の日付は入力できません");
	            document.frm.cmbJissiDate.focus();
	            return 1;
			}
		}

		<%
		'//新規登録時は、予定終了日を追加表示する
		If m_sMode = "BLANK" Or m_sMode = "INSERT" Then%>
	        if(f_Trim(document.frm.cmbJissiDateE.value) != "" ){

		        <%'// ■■■予定終了日の日付チェック■■■%>
	            if( IsDate(document.frm.cmbJissiDateE.value) != 0 ){
	                window.alert("日付の入力が不正です");
	                document.frm.cmbJissiDateE.focus();
	                return 1;
	            }

		        <%'// ■■■期間の大小のﾁｪｯｸ■■■%>
		        if( DateParse(document.frm.cmbJissiDate.value,document.frm.cmbJissiDateE.value) < 0){
		            window.alert("開始日と終了日を正しく入力してください");
		            document.frm.cmbJissiDate.focus();
		            return 1;
		        }

				<%'//■■■試験期間内かどうかチェック■■■%>
		        var MinDate = new Date("<%=m_sMinDate%>");
		        var MaxDate = new Date("<%=m_sMaxDate%>");
		        var vl = new Date(document.frm.cmbJissiDateE.value);
				if( vl< MinDate ){
		            window.alert("試験期間外の日付は入力できません");
		            document.frm.cmbJissiDateE.focus();
		            return 1;
				}else{
					if( vl > MaxDate ){
			            window.alert("試験期間外の日付は入力できません");
			            document.frm.cmbJissiDateE.focus();
			            return 1;
					}
				}

	        }

		<%End If%>

        // ■■■備考の桁ﾁｪｯｸ■■■
       if(f_Trim(document.frm.txtBiko.value) == "" ){
            window.alert("理由が入力されていません");
            document.frm.txtBiko.focus();
            return 1;
		}else{
	        if( getLengthB(document.frm.txtBiko.value) > "200" ){
	            window.alert("理由の欄は全角100文字以内で入力してください");
	            document.frm.txtBiko.focus();
	            return 1;
	        }
		}


<%'==================================================%>
	//予定が終日の時は入力チェックはしない
	if (document.frm.txtDayAll.checked==true){
		return 0;
	}
<%'==================================================%>

<%' ■■■以下時刻チェック ■■■%>

        // ■■■NULLﾁｪｯｸ■■■
        // ■開始時刻
        if( f_Trim(document.frm.txtKaisiH.value) == "" ){
            window.alert("開始時刻が入力されていません");
            document.frm.txtKaisiH.focus();
            return 1;
        }
        if( f_Trim(document.frm.txtKaisiN.value) == "" ){
            window.alert("開始時刻が入力されていません");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        // ■終了時刻
        if( f_Trim(document.frm.txtSyuryoH.value) == "" ){
            window.alert("終了時刻が入力されていません");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
        if( f_Trim(document.frm.txtSyuryoN.value) == "" ){
            window.alert("終了時刻が入力されていません");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        // ■■■半角数ﾁｪｯｸ■■■
        // ■開始時刻
        //var str = new String(document.frm.txtKaisiH.value);
        var str = document.frm.txtKaisiH.value;
        if( isNaN(str) ){
            window.alert("開始時刻が半角数字ではありません");
            document.frm.txtKaisiH.focus();
            return 1;
        }
        if( str < 0 ){
            window.alert("開始時刻が不正です");
            document.frm.txtKaisiH.focus();
            return 1;
        }

		n = str.match(".");
		if (n == ".") {
			alert("開始時刻が不正です"); 
		    document.frm.txtKaisiH.focus();
		    return 1;
		}

        //var str = new String(document.frm.txtKaisiN.value);
        var str = document.frm.txtKaisiN.value;
        if( isNaN(str) ){
            window.alert("開始時刻が半角数字ではありません");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        if( str < 0 ){
            window.alert("開始時刻が不正です");
            document.frm.txtKaisiN.focus();
            return 1;
        }

		n = str.match(".");
		if (n == ".") {
			alert("開始時刻が不正です"); 
		    document.frm.txtKaisiN.focus();
		    return 1;
		}

        // ■終了時刻
        //var str = new String(document.frm.txtSyuryoH.value);
        var str = document.frm.txtSyuryoH.value;
        if( isNaN(str) ){
            window.alert("終了時刻が半角数字ではありません");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
        if( str < 0 ){
            window.alert("終了時刻が不正です");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
		n = str.match(".");
		if (n == ".") {
			alert("終了時刻が不正です"); 
		    document.frm.txtSyuryoH.focus();
		    return 1;
		}

        //var str = new String(document.frm.txtSyuryoN.value);
        var str = document.frm.txtSyuryoN.value;
        if( isNaN(str) ){
            window.alert("終了時刻が半角数字ではありません");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        if( str < 0 ){
            window.alert("終了時刻が不正です");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
		n = str.match(".");
		if (n == ".") {
			alert("終了時刻が不正です"); 
		    document.frm.txtSyuryoN.focus();
		    return 1;
		}

        // ■■■桁ﾁｪｯｸ■■■
        // ■開始時刻
        var str = new String(document.frm.txtKaisiH.value);
        if( str.length > 2 ){
            window.alert("開始時刻が2桁以内ではありません");
            document.frm.txtKaisiH.focus();
            return 1;
        }
        var str = new String(document.frm.txtKaisiN.value);
        if( str.length > 2 ){
            window.alert("開始時刻が2桁以内ではありません");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        // ■終了時刻
        var str = new String(document.frm.txtSyuryoH.value);
        if( str.length > 2 ){
            window.alert("終了時刻が2桁以内ではありません");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
        var str = new String(document.frm.txtSyuryoN.value);
        if( str.length > 2 ){
            window.alert("終了時刻が2桁以内ではありません");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        // ■■■時刻ﾁｪｯｸ■■■
        // ■開始時刻
<%
'//        if( f_Trim(document.frm.txtKaisiH.value) < 9 ){
'//            window.alert("予定時間の入力は、9：00から23:55までです");
'//            document.frm.txtKaisiH.focus();
'//            return 1;
'//        }
%>
        if( f_Trim(document.frm.txtKaisiH.value) >= 24 ){
            window.alert("予定時間の入力は、23:55までです");
            document.frm.txtKaisiH.focus();
            return 1;
        }

        if( f_Trim(document.frm.txtKaisiN.value) >= 60 ){
            window.alert("予定時間を正確に入力してください");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        if( f_Trim(document.frm.txtKaisiN.value) < 0 ){
            window.alert("予定時間を正確に入力してください");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        // ■終了時刻
<%
'//        if( f_Trim(document.frm.txtSyuryoH.value) < 9 ){
'//            window.alert("予定時間の入力は、9：00から23:55までです");
'//            document.frm.txtSyuryoH.focus();
'//            return 1;
'//        }
%>
        if( f_Trim(document.frm.txtSyuryoH.value) >= 24 ){
            window.alert("予定時間の入力は、23:55までです");
            document.frm.txtSyuryoH.focus();
            return 1;
        }


        if( f_Trim(document.frm.txtSyuryoN.value) >= 60 ){
            window.alert("予定時間を正確に入力してください");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        if( f_Trim(document.frm.txtSyuryoN.value) < 0 ){
            window.alert("予定時間を正確に入力してください");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        // ■終了時刻（時）が開始時刻（時）超えてないか
        if( Number(f_Trim(document.frm.txtKaisiH.value)) > Number(f_Trim(document.frm.txtSyuryoH.value)) ){
            window.alert("終了時刻は開始時刻以降にしてください");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
        // ■終了時刻（分）が開始時刻（分）超えてないか
        if( Number(f_Trim(document.frm.txtKaisiH.value)) == Number(f_Trim(document.frm.txtSyuryoH.value)) ){
            if( Number(f_Trim(document.frm.txtKaisiN.value)) > Number(f_Trim(document.frm.txtSyuryoN.value)) ){
                window.alert("終了時刻は開始時刻以降にしてください");
                document.frm.txtSyuryoN.focus();
                return 1;
            }
        }
        // ■開始時刻
        var str = new String(document.frm.txtKaisiN.value);
        if( str.length < 2 ){
            str = 0 + str;
        }
        if( f_Trim(str).substr(1,1) != 0 ){
            if( f_Trim(str).substr(1,1) != 5 ){
                window.alert("予定時間は5分単位で入力してください");
                document.frm.txtKaisiN.focus();
                return 1;
            }
        }
        
        // ■終了時刻
        var str = new String(document.frm.txtSyuryoN.value);
        if( str.length < 2 ){
            str = 0 + str;
        }
        if( f_Trim(str).substr(1,1) != 0 ){
            if( f_Trim(str).substr(1,1) != 5 ){
                window.alert("予定時間は5分単位で入力してください");
                document.frm.txtSyuryoN.focus();
                return 1;
            }
        }
        // ■開始時刻と終了時刻が同一でないか
        if( f_Trim(document.frm.txtKaisiH.value) == f_Trim(document.frm.txtSyuryoH.value) ){
            if( f_Trim(document.frm.txtKaisiN.value) == f_Trim(document.frm.txtSyuryoN.value) ){
                window.alert("開始時刻と終了時刻が同一です");
                document.frm.txtSyuryoN.focus();
                return 1;
            }
        }
        
        return 0;
    }
<%
'    //************************************************************
'    //  [機能]  終日チェック時
'    //  [引数]  なし
'    //  [戻値]  
'    //  [説明]  
'    //************************************************************%>
	function f_ZenCheck(obj){
		if(obj.checked==true){
			document.frm.txtKaisiH.value = " -";
			document.frm.txtKaisiH.readOnly = true;

			document.frm.txtKaisiN.value = " -";
			document.frm.txtKaisiN.readOnly = true;

			document.frm.txtSyuryoH.value = " -";
			document.frm.txtSyuryoH.readOnly = true;

			document.frm.txtSyuryoN.value = " -";
			document.frm.txtSyuryoN.readOnly = true;

		}else{
			document.frm.txtKaisiH.value = "";
			document.frm.txtKaisiH.readOnly = false;

			document.frm.txtKaisiN.value = "";
			document.frm.txtKaisiN.readOnly = false;

			document.frm.txtSyuryoH.value = "";
			document.frm.txtSyuryoH.readOnly = false;

			document.frm.txtSyuryoN.value = "";
			document.frm.txtSyuryoN.readOnly = false;

		}
	}

    //-->

</SCRIPT>
<link rel=stylesheet href="../../common/style.css" type=text/css>
</head>

<body LANGUAGE="javascript" onload="return window_onload()">
<form name="frm" Method="POST">
<div align="center">
<%
if m_sMode = "DISP" or m_sMode = "UPDATE" Then
    call gs_title("試験監督免除申請登録","修　正")
Else
    call gs_title("試験監督免除申請登録","新規登録")
End If
%>
<br>
<br>


<%
'//試験期間表示
Call f_GetSikenKikan()
%>

<table border="0">
	<tr>
	    <td>

			<table border="0" cellpadding="1" cellspacing="1">
		    <COLGROUP  ALIGN=center>
			    <tr>
			        <td align="center">

			            <table border=1 CLASS="hyo">
					        <tr>
					            <TH nowrap CLASS="header" width="120" align="center">日　　付</TH>
					            <TD CLASS="detail"  width="370">
								<input type="text" name="cmbJissiDate" value="<%=m_iJissiDate%>" maxlength="10" size="15">
								<input type="button" class="button" onclick="fcalender('cmbJissiDate')" value="選択">

							<%If m_sMode = "BLANK" Or m_sMode = "INSERT" Then%>
								～
								<input type="text" name="cmbJissiDateE" value="<%=m_iJissiDateE%>" maxlength="10" size="15">
								<input type="button" class="button" onclick="fcalender('cmbJissiDateE')" value="選択">
								<br>
							<%End If%>

								<font size=2>（入力例:<%=date()%>）</font>
								</td>
					        </tr>
					        <tr>
					            <TH nowrap CLASS="header" width="120" align="center">理　　由</TH>
					            <TD CLASS="detail" width="280">
							    <textarea rows=4 cols=50 class=text name="txtBiko"><%=m_sBiko%></textarea><BR>
							    <font size=2>（全角100文字以内）</font>
								</TD>
					        </tr>

							<%

							'//入力されている時間より終日かどうかを判別
							'//C_MIN_TIME = "00:00"(最小時刻),C_MAX_TIME = "23:55"(最大時刻)
							If m_iKaisi = C_MIN_TIME And m_iSyuryo = C_MAX_TIME Then
								w_sCheck="checked"
								w_iKaisiH = " -"
								w_iKaisiN = " -"
								w_iSyuryoH = " -"
								w_iSyuryoN = " -"
							Else
								w_sCheck=""
								w_iKaisiH  = m_iKaisiH 
								w_iKaisiN  = m_iKaisiN 
								w_iSyuryoH = m_iSyuryoH
								w_iSyuryoN = m_iSyuryoN

							End If
							%>

					        <tr>
					            <TH nowrap CLASS="header" width="120" align="center">予定</TH>
					            <TD CLASS="detail" width="280">
								<input type="checkbox" name="txtDayAll" onclick="javascript:f_ZenCheck(this);"  <%=w_sCheck%> >全日
								</TD>
					        </tr>


					        <tr>
					            <TH nowrap CLASS="header" width="120" align="center">予定時間</TH>

								
					            <TD CLASS="detail" width="280">
								<input type="text" name="txtKaisiH" size="2" maxlength="2" value="<%=w_iKaisiH%>">時　
								<input type="text" name="txtKaisiN" size="2" maxlength="2" value="<%=w_iKaisiN%>">分　
								～
								<input type="text" name="txtSyuryoH" size="2" maxlength="2" value="<%=w_iSyuryoH%>">時　
								<input type="text" name="txtSyuryoN" size="2" maxlength="2" value="<%=w_iSyuryoN%>">分　
								</TD>
					        </tr>
			            </TABLE>

			        </td>
			    </TR>
			</table>

		</td>
	</tr>
    <tr>
		<td align="center">

		    <table>
		        <tr>
		            <td align="center">
		                <%if m_sMode = "DISP" or m_sMode = "UPDATE" Then%>
		                <input type="button" value="　更　新　" onClick="javascript:f_SaveClick();return false;" class=button>
		                <%else%>
		                <input type="button" value="　登　録　" onClick="javascript:f_SaveClick();return false;" class=button>
		                <%end if%>
		                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value=" ク　リ　ア " class=button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		                <input type="button" value="キャンセル" onClick="javascript:f_BackClick();return false;" class=button>
		            </td>
		        </tr>
		    </table>

	    </td>
	</tr>
</table>

</div>


<input type="hidden" name="txtMode" value="<%=m_sMode%>">
<input type="hidden" name="txtSikenKbn" value="<%=m_iSikenKbn%>">
<input type="hidden" name="txtSikenCd" value="<%=m_iSikenCd%>">
<input type="hidden" name="txtRenban" value="<%=m_iRenban%>">
<input type="hidden" name="txtPage" value="<%=m_iPage%>">


<input type="hidden" name="txtKeyYoteibi" value="<%=Request("txtKeyYoteibi")%>">

</form>
</body>

</html>


<%
    '---------- HTML END   ----------
End Sub
%>

