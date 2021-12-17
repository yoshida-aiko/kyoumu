<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 行事日程一覧
'* ﾌﾟﾛｸﾞﾗﾑID : gyo/gyo0200/main.asp
'* 機      能: 下ページ 行事日程マスタの一覧リスト表示を行う
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*           cboGyojiDate      :選択した行事日付
'*           chkGyojiCd      :行事コード
'*          txtMode             :動作モード
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'* 説      明:
'*           ■初期表示
'*               検索条件にかなう行事日程一覧を表示
'*           ■行事のみ表示チェックボックスON時
'*               行事のみ表示
'*-------------------------------------------------------------------------
'* 作      成: 2001/06/26 根本 直美
'* 変      更: 2001/07/27 伊藤公子  M40_CALENDERテーブル削除に対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    'Public  m_bErrMsg           'ｴﾗｰﾒｯｾｰｼﾞ
    Public  m_sMsg              'ﾒｯｾｰｼﾞ

    '取得したデータを持つ変数
    Public  m_iKyokanCd         ':教官コード
    Public  m_iSyoriNen         ':処理年度
    Public  m_iGyojiM          ':行事日程一覧月
    Public  m_iGyojiFlg         ':表示用
    
    Public  m_iDate            ':表示月日
    Public  m_iYear            ':表示年
    Public  m_iMonth            ':表示月
    Public  m_iDay              ':表示日
    Public  m_sYobi             ':表示曜日
    Public  m_iYobiCd           ':曜日コード
    Public  m_iKyujituFlg       ':休日コード（DB）
    Public  m_sColor            ':表示色（テーブル背景用）
    Public  m_iColorCd          ':表示色コード（テーブル背景用）
    Public  m_iGyojiCd          ':行事コード
    Public  m_sGyojiMei         ':行事名
    Public  m_sBiko             ':備考
    Public  m_iKaisibi          ':行事開始日
    Public  m_iSyuryobi         ':行事終了日
    Public  m_iHyojiFlg         ':表示フラグ
    Public  m_iNKaisiDate       ':年度開始日
    Public  m_iNKaisibi         ':年度開始日(日)

	'//学期関連情報
    Public  m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End

    Public  m_Rs                'recordset

    'ページ関係
    Public  m_iMax              ':最大ページ
    Public  m_iDsp              '// 一覧表示行数
    Public  m_iCount            '//1日当りの行事数
    Public  m_iCountN           '//1日当りの行事数（N番目）

    'データ取得用
'    Public  Const C_NENDO_KAISITUKI = 4             '年度開始月
    'Public  Const C_NENDO_KAISITUKI_MATUBI = 30     '年度開始月末日

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
    w_sMsgTitle="行事日程一覧"
    w_sMsg=""
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

        '// 値の初期化
        Call s_SetBlank()

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '// 年度開始日を取得
        Call f_GetNendoKaisibi()

		'//学期情報を取得
		w_iRet = gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'//行事明細テーブルより、選択された月のカレンダーを取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT T32_GYOJI_M.T32_HIDUKE"
		w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T32_GYOJI_M.T32_NENDO=" & cInt(m_iSyoriNen)
		w_sSQL = w_sSQL & vbCrLf & "  AND SUBSTR(T32_HIDUKE,6,2)='" & gf_fmtZero(m_iGyojiM,2) & "'"
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY T32_GYOJI_M.T32_HIDUKE"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY SUBSTR(T32_HIDUKE,9,2)"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            m_sErrMsg = "レコードセットの取得に失敗しました"
            Exit Do
        End If

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
    gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  全項目を空白に初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetBlank()

    m_iKyokanCd = ""
    m_iSyoriNen = ""
    m_iGyojiM = ""
    m_iGyojiFlg = ""

    m_iGyojiCd = ""
    m_sGyojiMei = ""
    m_sBiko = ""
    m_iHyojiFlg = ""
    
    m_sYobi = ""
    m_iKyujituFlg = ""
    m_iYobiCd = ""

    m_iDay = ""
    m_iMonth = ""
    m_iYear = ""
    m_iDate = ""
    m_sColor = ""
    
    m_iKaisibi = ""
    m_iSyuryobi = ""
    m_iNKaisiDate = ""
    m_iNKaisibi = ""
    
    m_iCount = ""
    m_iCountN = ""
    
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")         ':教官コード
    m_iSyoriNen = Session("NENDO")     ':処理年度
    m_iGyojiM = Request("cboGyojiDate")        ':表示月
    m_iGyojiFlg = 0                             ':行事表示用

	if Request("chkGyojiCd") = "on" Then
		m_iGyojiFlg = 1
		else
	end if

End Sub

'********************************************************************************
'*  [機能]  行事名の取得
'*  [引数]  
'*  [戻値]  0:情報取得成功、1:レコードなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetGyojiMei()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    f_GetGyojiMei = 0

    Do

        m_iCount = 0
        m_iCountN = 0

        '// 行事ヘッダﾚｺｰﾄﾞｾｯﾄを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & "T31_GYOJI_CD"
        w_sSQL = w_sSQL & ",T31_GYOJI_MEI"
        w_sSQL = w_sSQL & ",T31_BIKO"
        w_sSQL = w_sSQL & ",T31_KAISI_BI"
        w_sSQL = w_sSQL & ",T31_SYURYO_BI"
        w_sSQL = w_sSQL & ",T31_HYOJI_FLG"
        w_sSQL = w_sSQL & " FROM T31_GYOJI_H "
        w_sSQL = w_sSQL & " WHERE T31_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & " AND T31_KAISI_BI <= '" & m_iDate & "'"
        w_sSQL = w_sSQL & " AND T31_SYURYO_BI >= '" & m_iDate & "'"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
           'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_sGyojiMei = "　"
            m_sBiko = "　"
            m_sErrMsg = "ﾚｺｰﾄﾞｾｯﾄの取得失敗"
            m_bErrFlg = True
            f_GetGyojiMei = 99
            Exit Do
        Else
        End If

        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            m_sGyojiMei = "　"
            m_sBiko = "　"
            f_GetGyojiMei = 1
            Exit Do
        End If

        Do Until w_Rs.EOF
            '// 取得した値を格納
            m_iHyojiFlg = w_Rs("T31_HYOJI_FLG")     '//表示フラグを格納
            m_iKaisibi = w_Rs("T31_KAISI_BI")       '//行事開始日を格納
            m_iGyojiCd = w_Rs("T31_GYOJI_CD")       '//行事コードを格納

                if f_ChkKaisibi = 0 Then            '//開始日の場合表示
                    m_iCount = m_iCount + 1         '//表示件数をカウント
                else
                    if f_ChkHiduke = 0 Then         '//土日表示チェック
                        if f_ChkHyojibi = 0 Then        '//開始日以外でも全表示指定の場合は表示
                            m_iCount = m_iCount + 1     '//表示件数をカウント
                        else
                        end if
                    else
                        Exit Do
                    end if
                end if


            w_Rs.MoveNext
        Loop

        if m_iCount = 0 Then
            f_GetGyojiMei = 1
            Exit Do
        end if

        w_Rs.MoveFirst

        Do Until w_Rs.EOF
            '// 取得した値を格納

            m_iHyojiFlg = ""
            m_iKaisibi  = ""
            m_iGyojiCd  = ""
            m_iHyojiFlg = w_Rs("T31_HYOJI_FLG")     '//表示フラグを格納
            m_iKaisibi  = w_Rs("T31_KAISI_BI")       '//行事開始日を格納
            m_iGyojiCd  = w_Rs("T31_GYOJI_CD")       '//行事コードを格納
            m_sGyojiMei = w_Rs("T31_GYOJI_MEI")     '//行事名
            m_sBiko     = w_Rs("T31_BIKO")              '//備考

                if f_ChkKaisibi = 0 Then            '//開始日の場合表示
                    m_iCountN = m_iCountN + 1       '//表示件数(N番目)をカウント
                    Call show_Gyoji()
                else
                    if f_ChkHiduke = 0 Then                '//土日表示チェック
                        if f_ChkHyojibi = 0 Then          '//開始日以外でも全表示指定の場合は表示
                            m_iCountN = m_iCountN + 1      '//表示件数(N番目)をカウント
                            Call show_Gyoji()
                        else
                            'Call Show_NoGyoji()
                        end if
                    else
                        Exit Do
                    end if
                end if
                
            w_Rs.MoveNext
        Loop
        '// 正常終了
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs)

'// LABEL_f_GetGyojiMei_END
End Function

'********************************************************************************
'*  [機能]  日付CDチェック（土日祝日表示時使用）
'*  [引数]  なし
'*  [戻値]  0:情報取得成功、1:ﾚｺｰﾄﾞなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_ChkHiduke()

    Dim w_Rs2                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet2              '// 戻り値
    Dim w_sSQL2              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    f_ChkHiduke = 0

    Do
    
        '// 行事明細ﾚｺｰﾄﾞｾｯﾄを取得
        w_sSQL2 = ""
        w_sSQL2 = w_sSQL2 & "SELECT"
        w_sSQL2 = w_sSQL2 & " T32_GYOJI_CD"
        w_sSQL2 = w_sSQL2 & " FROM T32_GYOJI_M "
        w_sSQL2 = w_sSQL2 & " WHERE T32_NENDO = " & m_iSyoriNen
        w_sSQL2 = w_sSQL2 & " AND T32_HIDUKE = '" & m_iDate & "'"
        w_sSQL2 = w_sSQL2 & " AND T32_GYOJI_CD = " & m_iGyojiCd
        
        w_iRet2 = gf_GetRecordset(w_Rs2, w_sSQL2)
'response.write w_sSQL2 & "<br>"
        
        If w_iRet2 <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            'm_sErrMsg = "ﾚｺｰﾄﾞｾｯﾄの取得に失敗しました"
            f_ChkHiduke = 99
            Exit Do 'GOTO LABEL_f_ChkHiduke_END
        Else
        End If
        
        If w_Rs2.EOF Then
            '対象ﾚｺｰﾄﾞなし
            f_ChkHiduke = 1
            Exit Do 'GOTO LABEL_f_ChkHiduke_END
        End If

        '// 正常終了
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs2)

'// LABEL_f_ChkHiduke_END

End Function

'********************************************************************************
'*  [機能]  開始日チェック
'*  [引数]  なし
'*  [戻値]  0:開始日、1:開始日以外
'*  [説明]  
'********************************************************************************
Function f_ChkKaisibi()

    f_ChkKaisibi = 1
    
        if m_iDate = m_iKaisibi Then
            f_ChkKaisibi = 0
        end if

End Function

'********************************************************************************
'*  [機能]  表示日チェック
'*  [引数]  なし
'*  [戻値]  0:表示日、1:表示日以外
'*  [説明]  
'********************************************************************************
Function f_ChkHyojibi()

    f_ChkHyojibi = 1

    if m_iHyojiFlg = 0 Then
        f_ChkHyojibi = 0
    elseif m_iDay = 1 Then
        f_ChkHyojibi = 0
    end if

End Function

'********************************************************************************
'*  [機能]  DBから値を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParamD()

    m_sYobi = ""
    m_iKyujituFlg = ""
    m_iYobiCd = ""
    m_iDay   = ""
    m_iMonth = ""
    m_iYear  = ""
    m_iDate  = ""

    m_sYobi = left(WeekdayName(Weekday(CDate(m_Rs("T32_HIDUKE")))) ,1)
    'm_iKyujituFlg = m_Rs("M40_KYUJITU_FLG")
    m_iYobiCd = Weekday(m_Rs("T32_HIDUKE"))	'//曜日CD
    m_iDay = day(m_Rs("T32_HIDUKE"))		'//日
    m_iMonth = month(m_Rs("T32_HIDUKE"))	'//月
    m_iYear = year(m_Rs("T32_HIDUKE"))		'//年
    m_iDate = m_Rs("T32_HIDUKE")			'//日付

End Sub

'********************************************************************************
'*  [機能]  テーブルの背景色を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetColor()

	'//日付が祝日か休暇かどうか
	w_bHoliday = f_GetdateInfo()

	'//祝日,休暇
	If w_bHoliday = True Then
        m_sColor = "Holiday"
	Else
		'//祝日でない場合

        'If m_iYobiCd = "1" Then
        If m_iYobiCd = vbSunday Then
			'//日曜日
            m_sColor = "Holiday"

        'ElseIf  m_iYobiCd = "7" Then
        ElseIf  m_iYobiCd = vbSaturday Then
			'//土曜日
            m_sColor = "Saturday"
		Else
			'//平日
	        m_sColor = "Weekday"
        End If

	End If
    
End Sub

'********************************************************************************
'*  [機能]  日付が祝日かどうかを調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDateInfo()
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_bHoliday

    On Error Resume Next
    Err.Clear

	w_bHoliday = False

    Do

        '// 行事明細ﾚｺｰﾄﾞｾｯﾄ(休日データ)を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & " T32_GYOJI_CD"
        w_sSQL = w_sSQL & " FROM T32_GYOJI_M "
        w_sSQL = w_sSQL & " WHERE T32_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & " AND T32_HIDUKE = '" & m_iDate & "'"
        w_sSQL = w_sSQL & " AND T32_KYUJITU_FLG = '" & C_SYUKUJITU & "'"	'//T32_KYUJITU_FLG = C_SYUKUJITU …休日

'response.write w_sSQL & "<br>"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_sErrMsg = ""
            Exit Do
        End If

        If w_Rs.EOF Then
            '//休日ではない
			w_bHoliday = False
		Else
			'//休日
			w_bHoliday =True
        End If

        Exit Do
    Loop

	'//戻り値ｾｯﾄ
	f_GetDateInfo = w_bHoliday

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
    gf_closeObject(w_Rs)

End Function 

Sub show_Gyoji()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  行事予定が1件ある場合の出力
'********************************************************************************
if m_iCount = 1 Then
%>
    <TR>
        <TD ALIGN="right" class="<%=m_sColor%>"><%=m_iDay%><BR></TD>
        <TD ALIGN="center" class="<%=m_sColor%>"><%=m_sYobi%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>"><%=m_sGyojiMei%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">
<%

        '//4月始表示
        'if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) < CInt(m_iNKaisibi) Then
'        if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) <= CInt(day(m_sKouki_End)) Then
'            response.write m_sBiko & "※" & m_iYear & "年" & chr(13)
'        else
            response.write m_sBiko & chr(13)
'        end if

%><BR>
        </TD>
    </TR>
<%
else
    if m_iCountN = 1 Then
        Call show_GyojiS()
    else
        Call show_GyojiSTd()
    end if
end if

End Sub

Sub show_GyojiS()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  行事予定が複数ある場合の出力
'********************************************************************************
%>
    <TR>
        <TD ALIGN="right" class="<%=m_sColor%>" rowspan="<%=m_iCount%>"><%=m_iDay%><BR></TD>
        <TD ALIGN="center" class="<%=m_sColor%>" rowspan="<%=m_iCount%>"><%=m_sYobi%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>"><%=m_sGyojiMei%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">
<%
        '//4月始表示
        'if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) < CInt(m_iNKaisibi) Then
'        if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) <= CInt(day(m_sKouki_End)) Then
'            response.write m_sBiko & "※" & m_iYear & "年" & chr(13)
'        else
            response.write m_sBiko & chr(13)
'        end if
%><BR>
        </TD>
    </TR>
<%
End Sub

Sub show_GyojiSTd()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  行事予定が複数ある場合の出力
'********************************************************************************
%>
    <TR>
        <TD ALIGN="left" class="<%=m_sColor%>"><%=m_sGyojiMei%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">
<%
        '//4月始表示
        'if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) < CInt(m_iNKaisibi) Then
'        if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) <= CInt(day(m_sKouki_End)) Then
'            response.write m_sBiko & "※" & m_iYear & "年" & chr(13)
'        else
            response.write m_sBiko & chr(13)
'        end if
%><BR>
        </TD>
    </TR>
<%
End Sub

Sub show_NoGyoji()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  行事予定が無い場合の出力
'********************************************************************************
%>
    <TR>
        <TD ALIGN="right" class="<%=m_sColor%>"><%=m_iDay%><BR></TD>
        <TD ALIGN="center" class="<%=m_sColor%>"><%=m_sYobi%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">　<BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">
<%
        '//4月始表示
        'if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) < CInt(m_iNKaisibi) Then
'        if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) <= CInt(day(m_sKouki_End)) Then
'            response.write "※" & m_iYear & "年" & chr(13)
'        else
            response.write "　" & chr(13)
'        end if
%><BR>
        </TD>
    </TR>
<%
End Sub

Sub showPage_NoData()
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
		<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
    </center>

    </body>

    </html>


<%
    '---------- HTML END   ----------
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    Dim w_bFlg              '// ﾃﾞｰﾀ有無
    Dim w_bNxt              '// NEXT表示有無
    Dim w_bBfr              '// BEFORE表示有無
    Dim w_iNxt              '// NEXT表示頁数
    Dim w_iBfr              '// BEFORE表示頁数
    Dim w_iCnt              '// ﾃﾞｰﾀ表示ｶｳﾝﾀ

    Dim w_iRecordCnt        '//レコードセットカウント

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

%>

<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    //************************************************************
    //  [機能]  戻るボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //  [作成日] 
    //************************************************************
    function f_BackClick(){
    
        document.frm.action="../../menu/sansyo.asp";
        document.frm.target="_parent"
        document.frm.submit();
        
        
    }
//-->
</SCRIPT>
</head>

<body>
<center>
<form name="frm" Method="POST">
    <input type="hidden" name="txtMode" value="<%=m_sMode%>">
    <input type="hidden" name="chkGyojiCd" value="<%=m_iGyojiFlg%>">
<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr><td align="center">
    <table border="1" width="100%" CLASS="hyo">
        <colgroup width="10%" valign="top">
        <colgroup width="10%" valign="top">
        <colgroup width="40%" valign="top">
        <colgroup width="40%" valign="top">
        <TR>
            <TH CLASS="header">日</TH>
            <TH CLASS="header">曜日</TH>
            <TH CLASS="header">行事名</TH>
            <TH CLASS="header">備考</TH>
        </TR>

<%Do Until m_Rs.EOF%>

	<%
	'//グローバル変数に格納
	Call s_SetParamD()

	'//テーブル背景色設定
	Call s_SetColor()

	'//行事のみ表示が選択された場合
    if m_iGyojiFlg = 1 Then
        Call f_GetGyojiMei()
    else
        if f_GetGyojiMei() = 1 Then
            Call show_NoGyoji()
        end if
    end if

	m_Rs.MoveNext

    If m_Rs.EOF Then
        w_bFlg = False
    ElseIf w_iCnt >= m_iDsp Then
        w_iNxt = m_iPageT + 1
        w_bNxt = True
        w_bFlg = False
    Else
        w_iCnt = w_iCnt + 1
    End If
    if m_Rs.EOF Then
        Exit Do
    end if

Loop
%>
    </table>
</td>
</tr>
</table>

</form>
</center>
</body>

</html>
<%
    '---------- HTML END   ----------
End Sub

%>
