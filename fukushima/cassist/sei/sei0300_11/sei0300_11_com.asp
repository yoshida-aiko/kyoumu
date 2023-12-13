<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 個人別成績一覧　専用共通関数
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0300_11/sei0300_11_com.asp
' 機      能: 授業日数取得関数
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/09/03 伊藤公子
' 変      更: 2006/01/27 西村　彩和子　福島高専用に作成
'*************************************************************************/

'////////////////////////////////////////////////////////////////
'//コンスト定義
Private Const C_ZENKIKAISI = 10     '前期終了日
Private Const C_KOUKIKAISI = 11     '後期開始日
Private Const C_KOUKISYURYO = 12    '後期終了日

Private Const C_NULLGYOJI = 0       '行事なし
Private Const C_ALLNEN = 0          '全学年
Private Const C_ALLCLASS = 99       '全クラス

Private Const C_NULLJIGEN = 0       '空時限
Private Const C_NULLJIKAN = 0       '空総時間
Private Const C_NOZENJITU = 0       '空前日フラグ
Private Const C_FLG_TYOKIKYUKA = 1  '長期休暇フラグ

Private Const C_SOU_NISSU = True        '総授業日数
Private Const C_JUGYO_NISSU = False     '純授業日数

'////////////////////////////////////////////////////////////////

Public Function gf_SouJugyo(p_lJikan, p_sKCode, p_iNen, p_iClass, p_sKaisibi, p_sSyuryobibi, p_iNendo)
'*******************************************************************************
' 機　　能：総授業データの取得
' 返　　値：true: 成功　false: 失敗
' 引　　数：p_lJikan - 取得した時間数
' 　　　　　p_sKCode - 科目コード
' 　　　　　p_iNen - 学年
' 　　　　　p_iClass - クラス
' 　　　　　p_sKaisibi - 開始日
' 　　　　　p_sSyuryobibi - 終了日
' 機能詳細：総授業時間の取得
' 備　　考：なし
'*******************************************************************************
    Dim w_bRtn               '戻り値
    Dim w_sTyokiWhere         '長期休暇用のWhere
    Dim w_lJikan                '時間数
    
    Dim w_sStartDay           '開始日
    Dim w_sEndDay             '終了日
    
    'On Error GoTo Err_Func
	On Error Resume Next
	Err.Clear
    
    '== 変数の初期化 ==
    gf_SouJugyo = False
    
    w_sTyokiWhere = ""
    w_lJikan = 0
    
    w_sStartDay = ""
    w_sEndDay = ""
    
    '前期開始日より後の場合エラー
    If p_sKaisibi < f_GetGakkiDay(C_ZENKIKAISI, p_iNendo) Then
        'Call gf_iMsg(2129)
        Exit Function
    End If

    '後期終了日より後の場合エラー
    If p_sSyuryobibi > f_GetGakkiDay(C_KOUKISYURYO, p_iNendo) Then
        'Call gf_iMsg(2129)
        Exit Function
    End If

    '長期休暇の算出
    w_bRtn = f_GetTyokikyuka(w_sTyokiWhere, p_iNendo)
    If w_bRtn <> True Then: Exit Function

    '== 開始日が後期開始日より前(前期)の場合 ==
    If p_sKaisibi < f_GetGakkiDay(C_KOUKIKAISI, p_iNendo) Then
        '== 変数の設定 ==
        w_sStartDay = p_sKaisibi
        w_sEndDay = gf_YYYY_MM_DD(DateAdd("d", -1, f_GetGakkiDay(C_KOUKIKAISI, p_iNendo)),"/")

        '== 終了日が前期終了日より前の場合、終了日に変更 ==
        If p_sSyuryobibi < w_sEndDay Then
            w_sEndDay = p_sSyuryobibi
        End If

        '== 前期の時間割から取得 ==
        w_bRtn = f_GetJikanWari(C_GAKKI_ZENKI, p_sKCode, p_iNen, p_iClass, w_sStartDay, w_sEndDay, w_sTyokiWhere, C_SOU_NISSU, w_lJikan, p_iNendo)
        If w_bRtn <> True Then
            Exit Function
        End If

        '== なおかつ、終了日が後期開始日より後(学期またぎ)の場合 ==
        If p_sSyuryobibi > f_GetGakkiDay(C_KOUKIKAISI, p_iNendo) Then
            '== 変数の設定 ==
            w_sStartDay = f_GetGakkiDay(C_KOUKIKAISI, p_iNendo)
            w_sEndDay = p_sSyuryobibi

            '== 後期の時間割から取得 ==
            w_bRtn = f_GetJikanWari(C_GAKKI_KOUKI, p_sKCode, p_iNen, p_iClass, w_sStartDay, w_sEndDay, w_sTyokiWhere, C_SOU_NISSU, w_lJikan, p_iNendo)
            If w_bRtn <> True Then
                Exit Function
            End If

        End If

    '== 開始日が後期開始日より後(後期)の場合 ==
    Else
        '== 変数の設定 ==
        w_sStartDay = f_GetGakkiDay(C_KOUKIKAISI, p_iNendo)
        w_sEndDay = p_sSyuryobibi

        '== 後期の時間割から取得 ==
        w_bRtn = f_GetJikanWari(C_GAKKI_KOUKI, p_sKCode, p_iNen, p_iClass, w_sStartDay, w_sEndDay, w_sTyokiWhere, C_SOU_NISSU, w_lJikan, p_iNendo)
        If w_bRtn <> True Then
            Exit Function
        End If

    End If

    '== 時間数のセット ==
    p_lJikan = w_lJikan

    gf_SouJugyo = True

End Function

Private Function f_GetGakkiDay(p_iNo, p_iNendo)
'*******************************************************************************
' 機　　能：日付ゲット
' 返　　値：区分に適合する日付
' 引　　数：取得先のコード
' 機能詳細：マスタからデータを取得する　学期区分
' 備　　考：なし
'*******************************************************************************
Dim w_sSql
Dim w_oRecord
Dim w_bRtn

	On Error Resume Next
	Err.Clear

    f_GetGakkiDay = ""

    '重複チェック
    w_sSql = ""
    w_sSql = "Select M00_KANRI "
    w_sSql = w_sSql & "From "
    w_sSql = w_sSql & "M00_KANRI "
    w_sSql = w_sSql & "Where "
    w_sSql = w_sSql & "M00_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "M00_NO = " & p_iNo & " "

    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        '取得に失敗
        Exit Function
    End If
    f_GetGakkiDay = cstr(w_oRecord("M00_KANRI"))

    gf_closeObject(w_oRecord)
    
    Exit Function

'f_GetGakkiDay_E:


End Function

Private Function f_GetTyokikyuka(p_oTyokiWhere,p_iNendo)
'*******************************************************************************
' 機　　能：長期休暇の日付の取得
' 返　　値：true: 成功　false: 失敗
' 引　　数：長期休暇条件
' 　　　　　p_iNendo - 年度
' 機能詳細：長期休暇の日付の取得
' 備　　考：なし
'*******************************************************************************
Dim w_bRtn               '戻り値
Dim w_sSql                'SQL
Dim w_oRecKyuka
Dim w_sWhereTyoki

	'On Error GoTo Err_Func
	On Error Resume Next
	Err.Clear

    '== 初期化 ==
    f_GetTyokikyuka = False

    '== SQLの作成 ==
    w_sSql = ""
    w_sSql = w_sSql & "Select "
    w_sSql = w_sSql & "T31_KAISI_BI, "
    w_sSql = w_sSql & "T31_SYURYO_BI "
    w_sSql = w_sSql & "From "
    w_sSql = w_sSql & "T31_GYOJI_H "
    w_sSql = w_sSql & "Where "
    w_sSql = w_sSql & "T31_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "And "
    w_sSql = w_sSql & "T31_KYUKA_FLG = '" & C_FLG_TYOKIKYUKA & "' "
    w_sSql = w_sSql & "order by T31_KAISI_BI "

    '== データの取得 ==
    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecKyuka, w_sSql)
    If w_bRtn <> 0 Then
        Exit Function
    End If

    '== 長期休暇の日付を除くためのWhereを作成する ==
    Do Until w_oRecKyuka.EOF

        w_sWhereTyoki = w_sWhereTyoki & "And Not ("
        w_sWhereTyoki = w_sWhereTyoki & "T32_HIDUKE "
        w_sWhereTyoki = w_sWhereTyoki & "Between "
        w_sWhereTyoki = w_sWhereTyoki & "'" & cstr(w_oRecKyuka("T31_KAISI_BI")) & "' "
        w_sWhereTyoki = w_sWhereTyoki & "And "
        w_sWhereTyoki = w_sWhereTyoki & "'" & cstr(w_oRecKyuka("T31_SYURYO_BI")) & "') "
        w_oRecKyuka.MoveNext

    Loop

    '== レコードセットを閉じる ==
    Call gf_closeObject(w_oRecKyuka)

    p_oTyokiWhere = w_sWhereTyoki

    'データが無かったとしてもエラーではない
    f_GetTyokikyuka = True

    Exit Function

End Function

Private Function f_GetJikanWari(p_sGakki,p_sKCode,p_iNen, p_iClass, p_sStart, p_sEnd, p_sTyoki, p_bFlg, p_lJikan,p_iNendo)
'*******************************************************************************
' 機　　能：時間割データの取得
' 返　　値：true: 成功　false: 失敗
' 引　　数：p_sGakki - 学期フラグ
' 　　　　　p_sKCode - 科目コード
' 　　　　　p_iNen - 学年
' 　　　　　p_iClass - クラス
' 　　　　　p_sStart - 開始日
' 　　　　　p_sEnd - 終了日
' 　　　　　p_sTyoki - 長期休暇のWhere
' 　　　　　p_bFlg - 処理フラグ（true：総授業　false：純授業）
' 　　　　　p_lJikan - 結果格納変数
' 機能詳細：時間割の取得
' 備　　考：なし
'*******************************************************************************
    Dim w_sSql
    Dim w_bRtn
    Dim w_oRecord
    Dim w_lSojikan              '総時間数
    Dim w_lGyojiJikan           '行事時間数
    
    'On Error GoTo f_GetJikanWari_Err
	On Error Resume Next
	Err.Clear

    f_GetJikanWari = False
    w_lSojikan = 0

    '曜日　平日カウント
    w_sSql = ""
    w_sSql = w_sSql & "SELECT DISTINCT "
    w_sSql = w_sSql & "T20_YOUBI_CD, "
    w_sSql = w_sSql & "T20_JIGEN "
    w_sSql = w_sSql & "FROM "
    w_sSql = w_sSql & "T20_JIKANWARI "
    w_sSql = w_sSql & "WHERE "
    w_sSql = w_sSql & "T20_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T20_GAKKI_KBN = " & p_sGakki & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T20_KAMOKU = '" & p_sKCode & "' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T20_GAKUNEN = " & p_iNen & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T20_CLASS = " & p_iClass & " "

    '== データを取得する ==
    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        Exit Function
    End If

    w_oRecord.MoveFirst

    '== データの格納 ==
    Do Until w_oRecord.EOF = True
        '== 総授業日数の取得 ==
        w_bRtn = f_SouJugyoCnt(p_iNen, p_iClass, p_sStart, p_sEnd, w_oRecord("T20_YOUBI_CD"), w_oRecord("T20_JIGEN"), p_sTyoki, w_lSojikan, p_iNendo)
        If w_bRtn <> True Then
            '== 閉じる ==
            Call gf_closeObject(w_oRecord)

            Exit Function
        End If

        '== 時間の累計 ==
        p_lJikan = p_lJikan + w_lSojikan

        '== 純授業時間数を求める場合 ==
        If p_bFlg = C_JUGYO_NISSU Then
            '== 行事時間数の取得 ==
            w_bRtn = f_GyojiJugyoCnt(p_iNen, p_iClass, p_sStart, p_sEnd, w_oRecord("T20_YOUBI_CD"), w_oRecord("T20_JIGEN"), p_sTyoki, w_lGyojiJikan, p_iNendo)
            If w_bRtn <> True Then
                Exit Function
            End If

            '== 総時間数から行事時間数を引く ==
            p_lJikan = p_lJikan - w_lGyojiJikan * f_GetJigenTani(w_oRecord("T20_JIGEN"), p_iNendo)

        End If

        w_oRecord.MoveNext

    Loop

    '== 閉じる ==
    Call gf_closeObject(w_oRecord)

    f_GetJikanWari = True
    
    Exit Function

End Function

Public Function f_GetJigenTani(p_iJigen, p_iNendo)
'*******************************************************************************
' 機　　能：時限単位数の取得
' 返　　値：時限単位
' 引　　数：時限、年度
' 機能詳細：時限単位数の取得
' 備　　考：なし
'*******************************************************************************
Dim w_sSql
Dim w_oRecord
Dim w_bRtn

	On Error Resume Next
	Err.Clear

    f_GetJigenTani = 1

    '重複チェック
    w_sSql = ""
    w_sSql = w_sSql & "SELECT "
    w_sSql = w_sSql & "M07_TANISU "

    w_sSql = w_sSql & "FROM "
    w_sSql = w_sSql & "M07_JIGEN "

    w_sSql = w_sSql & "WHERE "
    w_sSql = w_sSql & "M07_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "M07_JIKAN = " & p_iJigen & " "

    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        '取得に失敗
        Exit Function
    End If

    If w_oRecord.EOF = True Then: Exit Function

    f_GetJigenTani = w_oRecord("M07_TANISU")

    Call gf_closeObject(w_oRecord)

    Exit Function

End Function

Private Function f_SouJugyoCnt(p_iNen, p_iClass, p_sStart, _
                               p_sEnd, p_iYoubi, p_iJigen, _
                               p_sTyoki, p_lJikan, p_iNendo)
'*******************************************************************************
' 機　　能：曜日毎時間データの取得
' 返　　値：true: 成功　false: 失敗
' 引　　数：学年、クラス、開始日、終了日、曜日、時限、長期休暇、結果格納変数
' 機能詳細：曜日毎時間データの取得
' 備　　考：なし
'*******************************************************************************
Dim w_sSql
Dim w_bRtn
Dim w_oRecord

	On Error Resume Next
	Err.Clear

    f_SouJugyoCnt = False
    p_lJikan = 0

    '曜日　平日カウント
    w_sSql = ""

    w_sSql = w_sSql & "SELECT DISTINCT "
    w_sSql = w_sSql & "T32_HIDUKE, "
    w_sSql = w_sSql & "T32_JIGEN "
    w_sSql = w_sSql & "FROM "
    w_sSql = w_sSql & "T32_GYOJI_M "

    w_sSql = w_sSql & "WHERE "
    w_sSql = w_sSql & "T32_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_KYUJITU_FLG = '0' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_GYOJI_CD = 0 "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_YOUBI_CD = " & p_iYoubi & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_JIGEN = " & p_iJigen & " "

    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_HIDUKE >= '" & p_sStart & "' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_HIDUKE <= '" & p_sEnd & "' "

    w_sSql = w_sSql & p_sTyoki

    '== データを取得する ==
    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        Exit Function
    End If

    '== データの格納 ==
    If w_oRecord.EOF = False Then

        w_oRecord.MoveLast

        'その曜日ごとの日数をカウントする
		p_lJikan = gf_GetRsCount(w_oRecord)

    End If

    Call gf_closeObject(w_oRecord)

    f_SouJugyoCnt = True

    Exit Function

End Function

Private Function f_GyojiJugyoCnt( p_iNen,  p_iClass,  p_sStart, _
                                p_sEnd,  p_iYoubi,  p_iJigen, _
                                p_sTyoki, p_lJikan,  p_iNendo)
'*******************************************************************************
' 機　　能：行事時間データの取得
' 返　　値：true: 成功　false: 失敗
' 引　　数：学年、クラス、開始日、終了日、曜日、時限、長期休暇、結果格納変数
' 機能詳細：行事時間データの取得
' 備　　考：なし
'*******************************************************************************
Dim w_sSql
Dim w_bRtn
Dim w_oRecord

	On Error Resume Next
	Err.Clear

    f_GyojiJugyoCnt = False

    '曜日　平日カウント
    w_sSql = ""

    w_sSql = w_sSql & "SELECT DISTINCT "
    w_sSql = w_sSql & "T32_HIDUKE, "
    w_sSql = w_sSql & "T32_JIGEN "
    
    w_sSql = w_sSql & "FROM "
    w_sSql = w_sSql & "T32_GYOJI_M "

    w_sSql = w_sSql & "WHERE "
    w_sSql = w_sSql & "T32_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_KYUJITU_FLG = '0' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_GYOJI_CD <> 0 "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_COUNT_KBN <> " & C_COUNT_KBN_JUGYO & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_YOUBI_CD = " & p_iYoubi & " "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_JIGEN = " & p_iJigen & " "

    If p_iNen <> C_ALLNEN Then

        w_sSql = w_sSql & "AND "
        w_sSql = w_sSql & "T32_GAKUNEN = " & p_iNen & " "

        If p_iClass <> C_ALLCLASS Then

            w_sSql = w_sSql & "AND "
            w_sSql = w_sSql & "T32_CLASS = " & C_ALLCLASS & " "

        End If

    End If

    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_HIDUKE >= '" & p_sStart & "' "
    w_sSql = w_sSql & "AND "
    w_sSql = w_sSql & "T32_HIDUKE <= '" & p_sEnd & "' "

    w_sSql = w_sSql & p_sTyoki

    '== データを取得する ==
    w_bRtn = gf_GetRecordset_OpenStatic(w_oRecord, w_sSql)
    If w_bRtn <> 0 Then
        Exit Function
    End If

    '== データの格納 ==
    If w_oRecord.EOF = False Then

        w_oRecord.MoveLast
        'その曜日ごとの日数をカウントする
		p_lJikan = gf_GetRsCount(w_oRecord)

    End If

    Call gf_closeObject(w_oRecord)

    f_GyojiJugyoCnt = True

    Exit Function

End Function
%>
