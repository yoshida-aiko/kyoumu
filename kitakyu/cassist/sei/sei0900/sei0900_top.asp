<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 仮進級者成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0900/sei0900_top.asp
' 機      能: 上ページ 仮進級者成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう調査書の内容を表示させる
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

    Public m_iNendo             '年度
	Public m_iRisyuKakoNendo    '履修過去年度
    Public m_iGakki             '学期
    Public m_sKyokanCd          '教官コード
    Public m_iSikenKbn			'試験区分
	Public m_sTxtMode           '//動作モード

    Public m_iDispFlg			'更新日表示フラグ 0:表示、1:非表示

	Public m_sGetTable			'科目コンボを作成したテーブル
    
    Public m_Rs_Nendo			'年度情報を取得
    Public m_Rs					'学年、クラス、科目取得RS
	Public m_Rs_NendoCount			'年度情報の件数
	Public m_RsCnt			'レコードカウント 科目取得RS

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
    w_sMsgTitle="仮進級者成績登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"

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
		'//値を取得
		call s_SetParam()

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

		' 学年末試験の試験区分を取得
		m_iSikenKbn =  C_SIKEN_KOU_KIM
		'//年度コンボを取得
        w_iRet = f_GetRisyuKakoNendo()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//年度がNULLだったら、コンボの最後の年度を入れる
		If  gf_IsNull(m_iRisyuKakoNendo) Then
	        m_Rs_Nendo.MoveLast
			m_iRisyuKakoNendo  = m_Rs_Nendo("T17_NENDO")
			m_Rs_Nendo.MoveFirst
    	End If

		if Not gf_IsNull(m_iRisyuKakoNendo) then

			'//科目コンボを取得
			w_iRet = f_GetKamoku_Nenmatu()
			If w_iRet <> 0 Then m_bErrFlg = True : Exit Do	

		End if

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
    Call gf_closeObject(m_Rs_Nendo)
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

    m_iNendo    = session("NENDO")
    m_iGakki    = Session("GAKKI")
    m_sKyokanCd = session("KYOKAN_CD")
	m_sTxtMode  = Request("txtMode")
	m_iRisyuKakoNendo  = Request("txtRisyuKakoNendo")
	

End Sub

'********************************************************************************
'*  [機能]  年度コンボを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetRisyuKakoNendo()

    Dim w_sSQL
	Dim w_bRtn
	Dim w_oRecord
	Dim w_Nendo
	Dim w_Count

    On Error Resume Next
    Err.Clear
    
    f_GetRisyuKakoNendo = 1



		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & "  DISTINCT T17_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  TT17_RISYUKAKO_KOJIN"
		w_sSQL = w_sSQL & vbCrLf & "  ,TT27_TANTO_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE "
 		w_sSQL = w_sSQL & vbCrLf & "  T17_NENDO = T27_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_KAMOKU_CD = T17_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_GAKUNEN = T17_HAITOGAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO 
		w_sSQL = w_sSQL & vbCrLf & "    AND (T17_TANI_SUMI =NULL OR T17_TANI_SUMI = 0) "
		w_sSQL = w_sSQL & vbCrLf & "  ORDER BY T17_NENDO"
' response.write "w_sSQL:" & w_sSQL & "<BR>"
' response.end
        ' w_bRtn = gf_GetRecordset(m_Rs_Nendo, w_sSQL)
		w_bRtn = gf_GetRecordset_OpenStatic(m_Rs_Nendo, w_sSQL)

        If w_bRtn <> 0 Then
             Exit Function
        End If
			
        f_GetRisyuKakoNendo = 0


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
		w_sSQL = w_sSQL & vbCrLf & "  WHERE T26_NENDO = " & m_iNendo

If m_iSikenKbn < C_SIKEN_KOU_KIM then '年度末試験の場合は、すべてが対象
		w_sSQL = w_sSQL & vbCrLf & "    AND T26_SIKEN_KBN =" & m_iSikenKbn
End If

		w_sSQL = w_sSQL & vbCrLf & "    AND T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "'"
		w_sSQL = w_sSQL & vbCrLf & "    AND ("
		w_sSQL = w_sSQL & vbCrLf & "       T26_JISSI_KYOKAN    ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN1 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN2 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN3 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN4 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN5 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    )"
		'w_sSQL = w_sSQL & vbCrLf & "    AND T26_JISSI_FLG = " & Cint(C_SIKEN_KBN_JISSI)
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
'*  [機能]  後期期末の時、科目コンボを取得
'*          その年度に実施された試験を全て表示する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKamoku_Nenmatu()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetKamoku_Nenmatu = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
		w_sSQL = w_sSQL & vbCrLf & " 	T27_KAMOKU_CD AS KAMOKU"
		w_sSQL = w_sSQL & vbCrLf & " 	,MAX(T17_KAMOKUMEI) AS KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM"
		w_sSQL = w_sSQL & vbCrLf & " 	T27_TANTO_KYOKAN "
		w_sSQL = w_sSQL & vbCrLf & " 	,T17_RISYUKAKO_KOJIN "
		w_sSQL = w_sSQL & vbCrLf & " 	,M05_CLASS "
		w_sSQL = w_sSQL & vbCrLf & "	,("
		w_sSQL = w_sSQL & vbCrLf & " 		SELECT * FROM TT13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " 		WHERE  T13_NENDO = " & cInt(m_iRisyuKakoNendo) - 1
		w_sSQL = w_sSQL & vbCrLf & " 		 AND T13_KARI_SINKYU = 1) TT13_GAKU_NEN "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " 		T27_NENDO = M05_NENDO "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_GAKUNEN = M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_CLASS = M05_CLASSNO	"
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_KAMOKU_CD = T17_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_GAKUNEN = T17_HAITOGAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_NENDO = T27_NENDO "
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_GAKUSEI_NO = T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_NENDO = " & cInt(m_iRisyuKakoNendo)
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "
		w_sSQL = w_sSQL & vbCrLf & "    AND (T17_TANI_SUMI =NULL OR T17_TANI_SUMI = 0) " & " "
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO 
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_COURSE_CD IN ( '0' , CASE WHEN M05_GAKKA_CD = T17_GAKKA_CD THEN (CASE WHEN M05_COURSE_CD IS NOT NULL THEN M05_COURSE_CD ELSE T17_COURSE_CD END ) ELSE T17_COURSE_CD END ) " '2019.02.12 Upd Kiyomoto
		w_sSQL = w_sSQL & vbCrLf & "  GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & " 	T27_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY "
		w_sSQL = w_sSQL & vbCrLf & "  KAMOKU"

' response.write w_sSQL  & "<BR>"
' rensponse.end

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
		' iRet =gf_GetRecordset_OpenStatic(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKamoku_Nenmatu = 99
            Exit Do
        End If	
		'//ﾚｺｰﾄﾞカウント取得
		m_RsCnt=gf_GetRsCount(m_Rs)

        f_GetKamoku_Nenmatu = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  履修テーブルより科目名称を取得
'*  [引数]  なし
'*  [戻値]  p_KamokuName
'*  [説明]  
'********************************************************************************
Function f_GetKamokuName(p_Gakunen,p_Class,p_KamokuCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_GakkaCd
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_GetKamokuName = ""
	p_KamokuName = ""

    Do 

		'//引数不足のとき
		If trim(p_Gakunen)="" Or trim(p_Class) = "" Or  trim(p_KamokuCd) = "" Then
            Exit Do
		End If

		'//学科CDを取得
		w_iRet = f_GetGakkaCd(p_Gakunen,p_Class,w_GakkaCd)
		If w_iRet<> 0 Then
            Exit Do
		End If

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,T15_RISYU.T15_LEVEL_FLG"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(p_Gakunen) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & w_GakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD=" & p_KamokuCd

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		If w_Rs.EOF = False Then
			p_KamokuName = w_Rs("T15_KAMOKUMEI")
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

		If w_Rs.EOF = False and cint(gf_SetNull2Zero(w_Rs("MAX(T15_LEVEL_FLG)"))) = 1 Then
			f_LevelChk = true
		End If

        Exit Do
    Loop
    Call gf_closeObject(w_Rs)
End Function

'********************************************************************************
'*  [機能]  学科CDを取得
'*  [引数]  p_Gakunen:学年,p_Class:クラス
'*  [戻値]  p_GakkaCd:学科CD
'*  [説明]  
'********************************************************************************
Function f_GetGakkaCd(p_Gakunen,p_Class,p_GakkaCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetGakkaCd = 1
	p_GakkaCd = ""

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKKA_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO= " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN=" & cint(p_Gakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_CLASSNO=" & cint(p_Class)

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetGakkaCd = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_GakkaCd = w_Rs("M05_GAKKA_CD")
		End If

        f_GetGakkaCd = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'****************************************************
'[機能] 置き換えられた選択科目は表示しないための関数。
'[引数] 
'       
'[戻値] 
'****************************************************
Function f_chkOkikae(p_KamokuCd)
	Dim w_sSql
	Dim w_Rs
	Dim i_Ret

	On Error Resume Next
    Err.Clear

	f_chkOkikae = 1

Do

	w_sSql = "Select "
	w_sSql = w_sSql & "T65_KAMOKU_CD_SAKI "
	w_sSql = w_sSql & "From "
	w_sSql = w_sSql & "T65_RISYU_SENOKIKAE "
	w_sSql = w_sSql & "Where "
	w_sSql = w_sSql & "T65_KAMOKU_CD_SAKI = '" & p_KamokuCd & "'"

	iRet = gf_GetRecordset(w_Rs, w_sSql)
	If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_chkOkikae = 99
            Exit Do
    End If

	If w_Rs.EOF = False Then
		f_chkOkikae = 0
	End If
	
    Exit Do
Loop
	Call gf_closeObject(w_Rs)

End Function

'****************************************************
'[機能] 置き換えられた代替科目を表示するための関数。（留学生用）
'[引数] 
'       
'[戻値] 
'****************************************************
Function f_chkRyuOkikae(p_KamokuCd)
	Dim w_sSql
	Dim w_Rs
	Dim i_Ret

	On Error Resume Next
    Err.Clear

	f_chkRyuOkikae = 1

Do

	w_sSql = ""
	w_sSql = "Select "
	w_sSql = w_sSql & "T75_KAMOKU_CD_SAKI "
	w_sSql = w_sSql & "From "
	w_sSql = w_sSql & "T75_RYU_OKIKAE "
	w_sSql = w_sSql & "Where "
	w_sSql = w_sSql & "T75_KAMOKU_CD_SAKI = '" & p_KamokuCd & "'"
	w_sSql = w_sSql & "And "
	w_sSql = w_sSql & "T75_NENDO = " & m_iNendo

	iRet = gf_GetRecordset(w_Rs, w_sSql)
	If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_chkRyuOkikae = 99
            Exit Do
    End If

	If w_Rs.EOF = False Then
		f_chkRyuOkikae = 0
	End If
	
    Exit Do
Loop
	Call gf_closeObject(w_Rs)

End Function


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
	
	If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		f_getTUKU = 99
		m_bErrFlg = True
	End If
	
	If w_Rs.EOF = false Then
		p_TUKU_FLG = cStr(gf_SetNull2Zero(w_Rs("T20_TUKU_FLG")))
	End If
	
    Call gf_closeObject(w_Rs)
	
End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
	Dim i
	On Error Resume Next
    Err.Clear
	i = 1
%>
	<html>
	<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//  [機能]  試験が変更されたとき、再表示する
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_ReLoadMyPage(){

	    document.frm.action="sei0900_top.asp";
	    document.frm.target="topFrame";
	    document.frm.submit();

	}

	//************************************************************
	//  [機能]  表示ボタンクリック時の処理
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_Search(){

	    // ■■■NULLﾁｪｯｸ■■■
	    // ■年度
	    if( f_Trim(document.frm.txtRisyuKakoNendo.value) == "<%=C_CBO_NULL%>" ){
	        window.alert("年度の選択を行ってください");
	        document.frm.txtRisyuKakoNendo.focus();
	        return ;
	    }

	    // ■科目名
	    if( f_Trim(document.frm.txtKamokuCd.value) == "<%=C_CBO_NULL%>" ){

			if (document.frm.txtKamokuCd.length ==1){
		        window.alert("試験科目がありません");
		        return ;
			}else{
		        window.alert("科目の選択を行ってください");
		        document.frm.txtKamokuCd.focus();
		        return ;
			}
	    }

		// 選択されたコンボの値をｾｯﾄ
		iRet = f_SetData();
		if( iRet != 0 ){
	        window.alert("科目のデータがありません");
			return;
		}

	    document.frm.action="sei0900_bottom.asp";
	    document.frm.target="main";
	    document.frm.submit();

	}

	//************************************************************
	//  [機能]  年度コンボChangeイベント
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
			
		function f_changeKamoku(){
	    
		// 選択されたコンボの値をｾｯﾄ
		iRet = f_GetKamokuCombo();
		if( iRet != 0 ){
	        window.alert("年度の選択を行ってください");
			return;
		}
	
		document.frm.action="sei0900_top.asp";
	    document.frm.target="topFrame";
		 document.frm.txtMode.value = "Reload";
	    document.frm.submit();

	}

	//************************************************************
	//  [機能]  年度コンボのデータから科目を取得する処理
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_GetKamokuCombo(){

		if (document.frm.cboRisyuKakoNendo.value==""){
			return 1;
        };
		
		//データ取得
        m_iRisyuKakoNendo = document.frm.cboRisyuKakoNendo.value;
		document.frm.txtRisyuKakoNendo.value=m_iRisyuKakoNendo;
		// document.frm.cboRisyuKakoNendo.selected= true;

        return 0;
	}
	//************************************************************
	//  [機能]  表示ボタンクリック時に選択されたデータをｾｯﾄ
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_SetData(){
		if (document.frm.cboKamoku.value==""){
			return 1;
        };
		if (document.frm.cboRisyuKakoNendo.value==""){
			return 1;
        };

		//データ取得
        var vl = document.frm.cboKamoku.value.split('#@#');
		
        //選択されたデータをｾｯﾄ(科目CDを取得)
        document.frm.txtKamokuCd.value=vl[0];
        document.frm.txtKamokuNM.value=vl[1];
		
		m_iRisyuKakoNendo = document.frm.cboRisyuKakoNendo.value;
		document.frm.txtRisyuKakoNendo.value=m_iRisyuKakoNendo;

        return 0;
	}

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

		// 選択されたコンボの値をｾｯﾄ
		iRet = f_SetData();
		if( iRet != 0 ){
			return;
		}
		
    }

	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>

    <body LANGUAGE=javascript onload="return window_onload()">
	
	<center>
	<form name="frm" METHOD="post">

	<% 
		Dim w_iGakunen_s
		Dim w_sGakkaCd_s
		Dim w_sKamokuCd_s
		Dim w_sKamokuNM_s

		call gs_title(" 仮進級者成績登録 "," 登　録 ") %>
<br>
	<table border="0">
	    <tr><td valign="bottom">

	        <table border="0" width="100%">
	            <tr><td class="search">

	                <table border="0">
	                    <tr valign="middle">
							<td align="left" nowrap>年度</td>
	                        <td align="left" colspan="3">
								<%If m_Rs_Nendo.EOF Then%>
									<select name="cboRisyuKakoNendo" style='width:150px;' DISABLED>
										<option value="">データがありません
								<%Else%>
									<select name="cboRisyuKakoNendo" style='width:150px;' onchange = 'javascript:f_changeKamoku()'>
									<%Do Until m_Rs_Nendo.EOF%>
										<option value='<%=m_Rs_Nendo("T17_NENDO")%>'  <%=f_Selected(m_Rs_Nendo("T17_NENDO"),m_iRisyuKakoNendo)%>><%=m_Rs_Nendo("T17_NENDO")%>
										<%m_Rs_Nendo.MoveNext%>
									<%Loop%>
								<%End If%>
								</select>
							</td>
	                        <td>&nbsp;</td>
	                        <td align="left" nowrap>科目</td>
	                        <td align="left">
								<%If m_iSikenKbn = "" Then%>
									<select name="cboKamoku" style='width:230px;' DISABLED>
										<option value="">データがありません
								<%Else%>
									<%If m_Rs.EOF Then%>
										<select name="cboKamoku" style='width:230px;' DISABLED>
											<option value="">科目データがありません
									<%Else%>
										<select name="cboKamoku" style='width:230px;' onclick="javasript:f_SetData();">
										<%Do Until m_Rs.EOF%>
											<%
											
											'//選択科目が置き換えられてた場合の表示 Add 2001.12.17 岡田
											If f_chkOkikae(m_Rs("KAMOKU")) = 0 then
												m_Rs.MoveNext
											
											Else
												w_sKamokuCd_s = m_Rs("KAMOKU")
												w_sKamokuNM_s = m_Rs("KAMOKUMEI")
													'//表示内容を作成
													If f_LevelChk(m_Rs("GAKUNEN"),m_Rs("KAMOKU")) = true then 
														w_Str=""
														w_Str= w_Str & m_Rs("KAMOKUMEI") & "　"
														
													Else
															w_Str=""
															w_Str= w_Str & m_Rs("KAMOKUMEI") & "　"	
													End If
													
											%>

											<option value=<%=w_sKamokuCd_s  & "#@#" & w_sKamokuNM_s%> ><%=w_Str%>
											<%
											'2002/02/21 追加 ITO 成績データの更新日を取得するためのKEYを退避
											w_sKamokuCd_s = m_Rs("KAMOKU")
											w_sKamokuNM_s = m_Rs("KAMOKUMEI")
											%>

											<%m_Rs.MoveNext%>
											<% End IF %>
										<%Loop%>
									<%End If
								End If
								%>
								</select>
							</td>
	                    </tr>
						<tr>
					        <td colspan="7" align="right">
							<%If m_RsCnt = 0 Then%>
								<input type="button" class="button" value="　表　示　" DISABLED>
							<%Else%>
								<input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
							<% End IF %>
					        </td>
						</tr>
	                </table>
	            </td>
				</tr>
	        </table>
	        </td>
	    </tr>
	</table>

	<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtKamokuCd" value="<%=w_sKamokuCd_s%>">
	<input type="hidden" name="txtKamokuNM" value="<%=w_sKamokuNM_s%>">
	<input type="hidden" name="txtRisyuKakoNendo" value="<%=m_iRisyuKakoNendo%>">
	<input type="hidden" name="txtTable" value="<%=m_sGetTable%>">
	 <input type="hidden" name="txtMode"  value = "">
	<!--ADD ST-->  
	<input type="hidden" name="txtUpdDate" value="<%=gf_GetT16UpdDate(m_iNendo,w_iGakunen_s,w_sGakkaCd_s,w_sKamokuCd_s,"")%>">
	<!--ADD ED --> 
	<input type="hidden" name="SYUBETU" value="">
	
	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>