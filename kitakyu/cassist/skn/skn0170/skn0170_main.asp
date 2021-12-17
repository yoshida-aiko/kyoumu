<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験時間割(クラス別)
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0170/skn0170_main.asp
' 機      能: MAINページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:   NENDO           '//年度
'               KYOKAN_CD       '//教官CD
'               cboGakunenCd    '//学年
'               cboClassCd      '//クラス
'               cboSikenKbn     '//試験区分
'               cboSikenCd      '//試験CD
' 引      渡:
' 説      明:
'           ■初期表示
'               空白ページを表示
'           ■表示ボタンが押された場合
'               検索条件にかなった試験時間割を表示
'-------------------------------------------------------------------------
' 作      成: 2001/07/19 伊藤公子
' 変      更: 2001/08/10 根本 直美     NN対応に伴うソース変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////

    Public Const C_TIMES_1COL = 5   '//1COLSPANあたりの時間(分)
    Public Const C_WIDTH_1COL = 9   '//1COLSPANあたりのTDのWIDTH
    Public Const C_TD_PADDING = 5   '//TDの余白 '2001/08/10 追加

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iSyoriNen          '//年度
    'Public m_iKyokanCd         '//教官ｺｰﾄﾞ
    Public m_iGakunen           '//学年
    Public m_iClassNo           '//クラスNO
    Public m_iSikenKbn          '//試験区分
    Public m_sSikenCd           '//試験CD
    Public m_sClassName         '//クラス名称
    Public m_sGakkaCd           '//学科CD

    Public m_sSikenName         '//試験名称
    Public m_sJiWari_Syuryo_Max '//試験終了時刻の最大時間
    Public m_sJiGen_Syuryo_Max  '//時限終了時刻の最大時間
    Public m_sJiGen_Kaisi_Min   '//時限開始時刻の最小時間

    'ﾚｺｰﾄﾞセット
    Public m_Rs_Jigen           '//時限ﾚｺｰﾄﾞｾｯﾄ
    Public m_Rs_Jiwari          '//時間割ﾚｺｰﾄﾞｾｯﾄ

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="試験時間割(クラス別)"
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
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '//値の初期化
        Call s_ClearParam()

        '//変数セット
        Call s_SetParam()

'//デバッグ
'Call s_DebugPrint()

        '//表示項目(クラス)を取得
        w_iRet = f_GetDisp_Data_Class()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//表示項目(試験)を取得
        w_iRet = f_GetDisp_Data_Siken()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//時限情報の取得
        w_iRet = f_GetJigen()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//時限情報のうち、最も遅く終わる時間を取得
        w_iRet = f_GetJigen_Max()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '// 試験時間割の取得 
        w_iRet = f_GetSikenJkanwari()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//試験時間割データのうち、最も遅く終わる試験時間を取得
        w_iRet = f_GetSiken_Max()
        If w_iRet <> 0 Then
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

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_Rs_Jigen)
    Call gf_closeObject(m_Rs_Jiwari)

    '// 終了処理
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    'm_iKyokanCd = ""
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_sClassMei = ""
    m_iSikenKbn = ""
    m_sSikenCd  = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")
    'm_iKyokanCd = Session("KYOKAN_CD")
    m_iGakunen  = Request("cboGakunenCd")   '//学年
    m_iClassNo  = Request("cboClassCd")     '//クラス
    m_iSikenKbn = Request("cboSikenKbn")    '//試験区分
    m_sSikenCd  = Request("cboSikenCd")     '//試験CD

    If trim(m_sSikenCd) = "" Or trim(m_sSikenCd) = "@@@" Then
        m_sSikenCd = "0"
    End If

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    'response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_iSikenKbn = " & m_iSikenKbn & "<br>"
    response.write "m_sSikenCd =  " & m_sSikenCd & "<br>"

End Sub

'********************************************************************************
'*  [機能]  クラス情報を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDisp_Data_Class()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Class = 1

    Do
        'クラスマスタよりデータを取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_CLASSMEI"
        w_sSql = w_sSql & vbCrLf & "  ,M05_CLASS.M05_GAKKA_CD"
        w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN= " & m_iGakunen
        w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_CLASSNO= "   & m_iClassNo

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetDisp_Data_Class = 99
            Exit Do
        End If

        m_sClassName = ""
        If rs.EOF = False Then
            m_sClassName = rs("M05_CLASSMEI")
            m_sGakkaCd = rs("M05_GAKKA_CD")
        End If

        '//正常終了
        f_GetDisp_Data_Class = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  表示項目(試験)を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDisp_Data_Siken()
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

        m_sSikenName = ""
        If rs.EOF = False Then
            m_sSikenName = rs("M01_SYOBUNRUIMEI")

            '//実力試験または、追試が選択された場合試験詳細名も追加表示
            If cint(m_sSikenCd) <> 0  Then
                m_sSikenName = m_sSikenName & " (" 
                m_sSikenName = m_sSikenName & rs("M27_SIKENMEI")
                m_sSikenName = m_sSikenName & " )" 
            End If

        End If

        '//正常終了
        f_GetDisp_Data_Siken = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  時限情報の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetJigen()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen = 1

    Do
        '試験時限マスタより本年度の試験時限を取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M26_JIGEN,"
        w_sSql = w_sSql & vbCrLf & "  M26_KAISI_JIKOKU,"
        w_sSql = w_sSql & vbCrLf & "  M26_SYURYO_JIKOKU"
        w_sSql = w_sSql & vbCrLf & " FROM M26_SIKEN_JIGEN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M26_NENDO = " & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & " ORDER BY "
        w_sSql = w_sSql & vbCrLf & "  M26_JIGEN "

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs_Jigen, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetJigen = 99
            Exit Do
        End If

        '//正常終了
        f_GetJigen = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  本年度の試験時限の最終時間と最小時間を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetJigen_Max()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen_Max = 1

    Do
        '試験時限マスタより本年度の試験時限を取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  MIN(M26_KAISI_JIKOKU) AS MIN_KAISI_JIKOKU,"
        w_sSql = w_sSql & vbCrLf & "  MAX(M26_SYURYO_JIKOKU) AS MAX_SYURYO_JIKOKU"
        w_sSql = w_sSql & vbCrLf & " FROM M26_SIKEN_JIGEN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M26_NENDO = " & m_iSyoriNen

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetJigen_Max = 99
            Exit Do
        End If

        m_sJiGen_Syuryo_Max = ""

        If rs.EOF = False Then
            m_sJiGen_Kaisi_Min =  rs("MIN_KAISI_JIKOKU")
            m_sJiGen_Syuryo_Max = rs("MAX_SYURYO_JIKOKU")
        End If

        '//正常終了
        f_GetJigen_Max = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  試験時間割の取得 
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSikenJkanwari()
    Dim w_iRet
    Dim w_sSQL

    On Error Resume Next
    Err.Clear

    f_GetSikenJkanwari = 1

    Do

        '試験時間割テーブルより時間割情報を取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  A.T26_SIKENBI, "
        w_sSql = w_sSql & vbCrLf & "  A.T26_KAMOKU, "
        w_sSql = w_sSql & vbCrLf & "  B.M04_KYOKANMEI_SEI AS JISSI_KYOKAN, "
        w_sSql = w_sSql & vbCrLf & "  C.M04_KYOKANMEI_SEI AS KANTOKU_KYOKAN," 
        'w_sSql = w_sSql & vbCrLf & "  D.M03_KAMOKUMEI, "
        w_sSql = w_sSql & vbCrLf & "  E.M06_KYOSITUMEI, "
        w_sSql = w_sSql & vbCrLf & "  A.T26_SIKEN_JIKAN, "
        w_sSql = w_sSql & vbCrLf & "  A.T26_KAISI_JIKOKU, "
        w_sSql = w_sSql & vbCrLf & "  A.T26_SYURYO_JIKOKU"
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "   T26_SIKEN_JIKANWARI A"
        w_sSql = w_sSql & vbCrLf & "  ,M04_KYOKAN B"
        w_sSql = w_sSql & vbCrLf & "  ,M04_KYOKAN C"
        'w_sSql = w_sSql & vbCrLf & "  ,M03_KAMOKU D"
        w_sSql = w_sSql & vbCrLf & "  ,M06_KYOSITU E"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      A.T26_JISSI_KYOKAN = B.M04_KYOKAN_CD(+) "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_KANTOKU_KYOKAN = C.M04_KYOKAN_CD(+)"
        'w_sSql = w_sSql & vbCrLf & "  AND A.T26_KAMOKU = D.M03_KAMOKU_CD(+)"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_KYOSITU = E.M06_KYOSITU_CD(+)"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=B.M04_NENDO(+) "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=C.M04_NENDO(+) "
        'w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=D.M03_NENDO(+) "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=E.M06_NENDO(+) "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_SIKEN_KBN=" & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_SIKEN_CD='" & m_sSikenCd & "' "
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_GAKUNEN=" & m_iGakunen
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_CLASS=" & m_iClassNo
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_JISSI_FLG=" & C_SIKEN_KBN_JISSI
        '//データが不完全なものは取得しない(実施日付・実施時間・開始時間・実施教官・監督教官のどれかひとつでも入ってないものは表示しない)
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_SIKENBI IS NOT NULL"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_KAISI_JIKOKU IS NOT NULL"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_SYURYO_JIKOKU IS NOT NULL"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_JISSI_KYOKAN IS NOT NULL"
        w_sSql = w_sSql & vbCrLf & "  AND A.T26_KANTOKU_KYOKAN IS NOT NULL "
        w_sSql = w_sSql & vbCrLf & " ORDER BY "
        w_sSql = w_sSql & vbCrLf & "  T26_SIKENBI,T26_KAISI_JIKOKU "

'response.write w_sSQL & "<br>"
'response.end

        iRet = gf_GetRecordset_OpenStatic(m_Rs_Jiwari,w_sSQL)
        If iRet <> 0  Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSikenJkanwari = 99
            Exit Do
        End If
        '//正常終了
        f_GetSikenJkanwari = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  試験時間割データのうち、最も遅く終わる試験時間を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSiken_Max()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetSiken_Max = 1

    Do

        '//最も遅く終わる試験時間を取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  Max(T26_SYURYO_JIKOKU) AS MAX_SYURYO_JIKOKU"
        w_sSql = w_sSql & vbCrLf & " FROM T26_SIKEN_JIKANWARI"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      T26_SIKEN_JIKANWARI.T26_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_KBN=" & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_CD='" & m_sSikenCd & "' "
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_GAKUNEN=" & m_iGakunen
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_CLASS=" & m_iClassNo
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_JISSI_FLG=" & C_SIKEN_KBN_JISSI

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSiken_Max = 99
            Exit Do
        End If

        m_sJiWari_Syuryo_Max = ""

        If rs.EOF = False Then
            m_sJiWari_Syuryo_Max = rs("MAX_SYURYO_JIKOKU")
        End If

        '//正常終了
        f_GetSiken_Max = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  科目名を取得
'*  [引数]  p_sKamokuCd
'*  [戻値]  f_GetKamokName
'*  [説明]  
'********************************************************************************
Function f_GetKamokuName(p_sKamokuCd)
    Dim w_iRet
    Dim w_sSQL
    Dim rs
    Dim w_sKamokuName

    On Error Resume Next
    Err.Clear

    w_sKamokuName = ""

    Do

        '//科目名を取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI"
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  T15_RISYU"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & m_iSyoriNen - m_iGakunen + 1
        w_sSql = w_sSql & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "' "
        w_sSql = w_sSql & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_sKamokuCd & "'"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            Exit Do
        End If

        If rs.EOF = False Then
            w_sKamokuName = rs("T15_KAMOKUMEI")
        End If

        '//戻値ｾｯﾄ
        f_GetKamokuName = w_sKamokuName

        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  時間よりCOLSPANを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  1Colspan/5分とする
'********************************************************************************
Function f_Get_Colspan(p_sStartTime,p_sEndTime)
    Dim w_iTime
    Dim w_iColSpan
    On Error Resume Next
    Err.Clear

    w_iTime = 0
    w_iColSpan = 0

    Do
        w_iTime = DateDiff("n", p_sStartTime, p_sEndTime)
        w_iColSpan = w_iTime\C_TIMES_1COL   '//C_TIMES_1COL = 5 (1colspan/5分)5分単位

        Exit Do
    Loop

    Err.Clear
    f_Get_Colspan = w_iColSpan

End Function

'********************************************************************************
'*  [機能]  時間割内容をセット
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_SetNaiyo(p_Naiyo,p_Add)
    Dim w_sNaiyo

    w_sNaiyo = ""
    If Trim(gf_SetNull2String(p_Naiyo)) <> "" Then
        w_sNaiyo = "<br>" & p_Naiyo & p_Add
    End If

    f_SetNaiyo = w_sNaiyo

End Function

'********************************************************************************
'*  [機能]  日付を"M月D日(曜日)"の形にする
'*  [引数]  p_Date
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Function f_fmtDate(p_Date)
    Dim w_sDate

    w_sDate = ""

    If gf_SetNull2String(p_Date) <> "" Then
        w_sDate = month(p_Date) & "月"
        w_sDate = w_sDate & day(p_Date) & "日"
        w_sDate = w_sDate & "("
        w_sDate = w_sDate & gf_GetYoubi(Weekday(p_Date))
        w_sDate = w_sDate & ")"
    End If

    f_fmtDate = w_sDate

End Function

'********************************************************************************
'*  [機能]  空白TDを表示する
'*  [引数]  p_STime:時間(小)
'*          p_BTime:時間(大)
'*          p_Class:TDのclass
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetBrankTD(p_STime,p_BTime,p_Class)
Dim w_iColSpan

    '//Colspanを取得
    w_iColSpan = f_Get_Colspan(p_STime, p_BTime)
    If w_iColSpan > 0 Then
        %>
        <!--<td class="<%=p_Class%>" align="center" width="<%=w_iColSpan*C_WIDTH_1COL%>" colspan="<%'=w_iColSpan%>"><font ><br></font></td>-->
        <td class="<%=p_Class%>" align="center" colspan="<%=w_iColSpan%>" ><font ><br></font></td>
        <%
    End If
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    Dim w_sNaiyo    '//表示内容
    Dim w_MaxTime   '//試験期間最終時刻
    Dim w_iColSpan  '//COLSPAN
    Dim w_sEndTime  '//試験終了時間
    Dim w_sDate     '//試験日
    Dim w_sKaisi    '//試験開始時刻

%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>試験時間割(クラス別)</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">

<%
'//デバッグ
'Call s_DebugPrint()
%>
<center>
<br>
    <%Do%>
        <%
        '//試験データがない または時間割りデータがない場合
        If m_Rs_Jiwari.EOF = True or m_Rs_Jigen.EOF = True Then 
        %>
        <br><br><span class="msg">試験時間割情報がありません</span>
        <%
            Exit Do
        End If
        %>


        <table class="hyo" border="1" width="400">
            <tr>
                <th class="header" width="80"  align="center" nowrap><font size="2">クラス</font></th>
                <td class="detail" width="120" align="center" nowrap><font size="2"><%=m_iGakunen & "年　" & m_sClassName%></font></td>
                <th class="header" width="80"  align="center" nowrap><font size="2">試験</font></th>
                <td class="detail" width="180" align="center" nowrap><font size="2"><%=m_sSikenName%></font></td>
            </tr>
        </table>
        <br>

        <table border="0">
        <tr><td width="10"><br></td></tr>
        <tr><td align="center">

        <!--ヘッダ部-->
        <table class="hyo" border="1" width="100%">
        <%If m_Rs_Jigen.EOF=False Then%>
            <%
            '//===============================
            '//試験時間の最大の最終時間を取得
            '//===============================
            If m_sJiWari_Syuryo_Max >= m_sJiGen_Syuryo_Max Then
                w_MaxTime = m_sJiWari_Syuryo_Max
            Else
                w_MaxTime = m_sJiGen_Syuryo_Max
            End If

            %>
            <tr>
            <td class="header" align="center" colspan="1" nowrap><font color="#ffffff" size="2">時　限</font></td>
            <%

            '//=============
            '//時限を表示
            '//=============
            Do Until m_Rs_Jigen.EOF
                '===時限===
                '//Colspanを取得
                w_iColSpan = f_Get_Colspan(m_Rs_Jigen("M26_KAISI_JIKOKU"), m_Rs_Jigen("M26_SYURYO_JIKOKU"))
                w_sEndTime = m_Rs_Jigen("M26_SYURYO_JIKOKU")
                %>
                <td class="header2" align="center" width="<%=w_iColSpan*C_WIDTH_1COL%>" colspan="<%=w_iColSpan%>" nowrap><img src="../../image/sp.gif" width="<%=w_iColSpan*C_WIDTH_1COL-C_TD_PADDING*2%>" height="1"><br><font color="#ffffff" size="2"><%=m_Rs_Jigen("M26_JIGEN")%></font></td>
                <%
                m_Rs_Jigen.MoveNext
                If m_Rs_Jigen.EOF = False Then
                    '//空白TDをセット
                    Call s_SetBrankTD(w_sEndTime, m_Rs_Jigen("M26_KAISI_JIKOKU"),"header2")
                Else
                    '//空白TDをセット
                    Call s_SetBrankTD(w_sEndTime, w_MaxTime,"header2")
                End If%>

            <%Loop%>
            </tr>
            <tr>
            <td class="header" align="center" colspan="1" nowrap><font size="2" color="#ffffff">時　間</font></td>
            <%

            '//=================
            '//試験時間を表示
            '//=================
            m_Rs_Jigen.MoveFirst
            Do Until m_Rs_Jigen.EOF
                '//Colspanを取得
                w_iColSpan = f_Get_Colspan(m_Rs_Jigen("M26_KAISI_JIKOKU"), m_Rs_Jigen("M26_SYURYO_JIKOKU"))
                w_sEndTime = ""
                w_sEndTime = m_Rs_Jigen("M26_SYURYO_JIKOKU")
                '===時間===
                %>
                <td class="header2" align="center" width="<%=w_iColSpan*C_WIDTH_1COL%>" colspan="<%=w_iColSpan%>" nowrap>
                <font size="2" color="#ffffff"><%=gf_SetNull2String(m_Rs_Jigen("M26_KAISI_JIKOKU"))%>～<%=gf_SetNull2String(m_Rs_Jigen("M26_SYURYO_JIKOKU"))%></font></td>
                <%

                m_Rs_Jigen.MoveNext
                If m_Rs_Jigen.EOF = False Then
                    '//空白TDをセット
                    Call s_SetBrankTD(w_sEndTime, m_Rs_Jigen("M26_KAISI_JIKOKU"),"header2")
                Else
                    '//空白TDをセット
                    Call s_SetBrankTD(w_sEndTime, w_MaxTime,"header2")
                End If
            Loop%>
            <!--</tr>-->
        <%End If%>


        <!--明細部-->
        <%If m_Rs_Jiwari.EOF = False Then%>

            <%
            Do Until m_Rs_Jiwari.EOF

                '//=================
                '//試験日を表示
                '//=================
                If w_sDate <> m_Rs_Jiwari("T26_SIKENBI") Then
                    w_sDate = m_Rs_Jiwari("T26_SIKENBI")%>
                    </tr>
                    <tr>
                        <td class="header" align="center" height="35" colspan="1" nowrap><font size="2" color="#ffffff"><%=f_fmtDate(m_Rs_Jiwari("T26_SIKENBI"))%></font></td>
                    <%
                    '//時限時間の最小時間より、試験時間が遅い場合
                    If m_sJiGen_Kaisi_Min < m_Rs_Jiwari("T26_KAISI_JIKOKU") Then
                        '//空白TDをセット
                        Call s_SetBrankTD(m_sJiGen_Kaisi_Min, m_Rs_Jiwari("T26_KAISI_JIKOKU"),"CELL2")
                    End If
                End If

                '//=================
                '//試験内容を表示
                '//=================
                '//表示する内容を取得
                w_sNaiyo = f_GetKamokuName(m_Rs_Jiwari("T26_KAMOKU"))
                w_sNaiyo = w_sNaiyo & "(" & m_Rs_Jiwari("T26_SIKEN_JIKAN") & ")"
                w_sNaiyo = w_sNaiyo & f_SetNaiyo(m_Rs_Jiwari("JISSI_KYOKAN"),"(試)") & f_SetNaiyo(m_Rs_Jiwari("KANTOKU_KYOKAN"),"(監)") & f_SetNaiyo(m_Rs_Jiwari("M06_KYOSITUMEI"),"")

                '===============================================
                '//同じ時刻に別のテスト科目が入っていた場合の考慮
                w_sKaisi = m_Rs_Jiwari("T26_KAISI_JIKOKU")
                w_iMax_Time = m_Rs_Jiwari("T26_SIKEN_JIKAN")
                Do Until m_Rs_Jiwari.EOF
                    m_Rs_Jiwari.MoveNext
                    '//次のレコードがEOFでない場合
                    If m_Rs_Jiwari.EOF = False Then
                        '//日付が変わってないかどうか
                        If w_sDate <> m_Rs_Jiwari("T26_SIKENBI") Then
                            m_Rs_Jiwari.MovePrevious
                            Exit Do
                        Else
                            '//前のレコードの開始時間と、次のﾚｺｰﾄﾞの開始時間が同じ場合は同じ時刻に別のテストが入っている
                            If w_sKaisi = m_Rs_Jiwari("T26_KAISI_JIKOKU") Then

                                '//最大時間を取得
                                If cint(w_iMax_Time) < cint(m_Rs_Jiwari("T26_SIKEN_JIKAN")) Then
                                    w_iMax_Time = m_Rs_Jiwari("T26_SIKEN_JIKAN")
                                End If

                                '//内容をｾｯﾄ
                                w_sNaiyo = w_sNaiyo & "<br>-------<br>"
                                w_sNaiyo = w_sNaiyo & f_GetKamokuName(m_Rs_Jiwari("T26_KAMOKU")) & "(" & m_Rs_Jiwari("T26_SIKEN_JIKAN") & ")"
                                w_sNaiyo = w_sNaiyo & f_SetNaiyo(m_Rs_Jiwari("JISSI_KYOKAN"),"(試)") & f_SetNaiyo(m_Rs_Jiwari("KANTOKU_KYOKAN"),"(監)") & f_SetNaiyo(m_Rs_Jiwari("M06_KYOSITUMEI"),"")
                            Else
                                m_Rs_Jiwari.MovePrevious
                                Exit Do
                            End If
                        End If
                    Else
                        m_Rs_Jiwari.MovePrevious
                        Exit Do
                    End If
                Loop
                '===============================================

                '//COLSPANを取得
                'w_iColSpan = f_Get_Colspan(m_Rs_Jiwari("T26_KAISI_JIKOKU"), m_Rs_Jiwari("T26_SYURYO_JIKOKU"))
                w_iColSpan = cint(w_iMax_Time)\C_TIMES_1COL

                '//試験終了時刻を取得(空白TDに必要)
                w_sEndTime = ""
                w_sEndTime = m_Rs_Jiwari("T26_SYURYO_JIKOKU")
                %>
                <td class="CELL1" width="<%=w_iColSpan*C_WIDTH_1COL%>" colspan="<%=w_iColSpan%>" valign="top"><font size="2"><%=w_sNaiyo%></font></td>

                <%m_Rs_Jiwari.MoveNext
                If m_Rs_Jiwari.EOF = False Then
                    '//次のレコードの実施日が変わった場合、残りのTDを追加する
                    If w_sDate <> m_Rs_Jiwari("T26_SIKENBI") Then
                        '//空白TDをセット
                        Call s_SetBrankTD(w_sEndTime, w_MaxTime,"CELL2")
                    Else
                        '//空白TDをセット
                        Call s_SetBrankTD(w_sEndTime, m_Rs_Jiwari("T26_KAISI_JIKOKU"),"CELL2")
                    End If
                Else
                    '//空白TDをセット
                    Call s_SetBrankTD(w_sEndTime, w_MaxTime,"CELL2")
                End If
            Loop
        End If%>

                </tr>
    </table>
    </td></tr>
    </table>

    <%
        Exit Do
    Loop%>

</center>
</body>
</html>
<%
End Sub
%>