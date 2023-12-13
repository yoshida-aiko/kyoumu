<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0112/kks0112_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:txtGakunen     :学年
'            txtClass       :学科
'            txtTuki        :月
' 説      明:
'           ■初期表示
'               前後期のコンボボックスは当期を表示
'               月のコンボボックスは当月を表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう行事一覧を表示させる
'           ■登録ボタンクリック時
'               入力された情報を登録する
'-------------------------------------------------------------------------
' 作      成: 2001/07/03 伊藤公子
' 変      更: 2001/12/20 谷脇良也　留学生代替科目対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 0

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iSyoriNen          '教官ｺｰﾄﾞ
    Public m_iKyokanCd          '年度
    Public m_sGakki             '//学期
    Public m_sGakunen           '//学年
    Public m_sClassNo           '//クラスNO
    Public m_sClassMei          '//クラス名
    Public m_sTuki_Zenki_Start  '//前期開始日
    Public m_sTuki_Kouki_Start  '//後期開始日
    Public m_sTuki_Kouki_End    '//後期終了日
    Public m_Rs_Month           '//月
    Public m_Rs_Sbj             '//授業
    Public m_Rs_Daigae          '//代替授業
    Public m_Rs_RyuDaigae       '//留学生の代替授業
	Public m_JigenCount			'//時限数
    Public m_Rs_Jigen			'//レコードセット
    
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
    w_sMsgTitle="授業出欠入力"
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
		
        '//学期を取得
        m_sGakki = Request("GAKKI")
		
        If trim(m_sGakki) <> "" Then
            '//前期・後期情報を取得
            m_sTuki_Zenki_Start = Request("Tuki_Zenki_Start")
            m_sTuki_Kouki_Start = Request("Tuki_Kouki_Start")
            m_sTuki_Kouki_End   = Request("Tuki_Kouki_End")
        Else
            '//前期・後期情報を取得
            w_iRet = f_GetGakkiInfo()
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If
        End If
		
        '//ログイン教官の受持教科を取得(年度、教官CD、学期より)
        w_iRet = f_GetSubject()
        If w_iRet <> 0 Then
           m_bErrFlg = True
            Exit Do
        End If
		
        '//代替授業を取得
        w_iRet = f_GetDaigae()
        If w_iRet <> 0 Then
           m_bErrFlg = True
            Exit Do
        End If
		
        '//留学生の代替授業を取得
        w_iRet = f_GetRyuDaigae()
        If w_iRet <> 0 Then
           m_bErrFlg = True
           Exit Do
        End If
		
		'//時限数の取得
		if not f_Get_JigenData() then
			m_bErrFlg = True
			Exit Do
		end if
		
        '// ページを表示
        Call showPage()
        Exit Do
    Loop
	
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        'Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gf_closeObject(m_Rs_Month)
    Call gf_closeObject(m_Rs_Sbj)
    Call gf_closeObject(m_Rs_Daigae)

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
    m_iKyokanCd = ""
    m_sGakki    = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub

'********************************************************************************
'*  [機能]  前期・後期情報を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetGakkiInfo()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetGakkiInfo = 1

    Do
        '管理マスタから学期情報を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NO, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_KANRI, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_BIKO"
        w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO=" & cInt(m_iSyoriNen) & " AND "
        'w_sSQL = w_sSQL & vbCrLf & "   (M00_KANRI.M00_NO=10 Or M00_KANRI.M00_NO=11 Or M00_KANRI.M00_NO=12) "   '//[M00_NO]10:前期開始 11:後期開始
        w_sSQL = w_sSQL & vbCrLf & "   (M00_KANRI.M00_NO=" & C_K_ZEN_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_SYURYO & ") "  '//[M00_NO]10:前期開始 11:後期開始

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetGakkiInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            Do Until rs.EOF

                If cInt(rs("M00_NO")) = C_K_ZEN_KAISI Then
                    m_sTuki_Zenki_Start = rs("M00_KANRI")
                ElseIf cInt(rs("M00_NO")) = C_K_KOU_KAISI Then
                    m_sTuki_Kouki_Start = rs("M00_KANRI")
                ElseIf cInt(rs("M00_NO")) = C_K_KOU_SYURYO Then
                    m_sTuki_Kouki_End = rs("M00_KANRI")
                End If
                rs.MoveNext
            Loop

            '//現在の前期後期判定
            If gf_YYYY_MM_DD(date(),"/") < m_sTuki_Kouki_Start Then
                m_sGakki = "ZENKI"
            Else
                m_sGakki = "KOUKI"
            End If

        End If

        '//正常終了
        f_GetGakkiInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  コンボ月を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetTuki(p_sGakki)
    Dim w_iRet
    Dim w_sSQL


    On Error Resume Next
    Err.Clear


    If p_sGakki ="ZENKI" Then

        '//学期開始月
        w_iStartTuki = Month(m_sTuki_Zenki_Start)

        '//学期終了月
        If day(m_sTuki_Kouki_Start) <> 1 Then
            w_iEndTuki = Month(m_sTuki_Kouki_Start)
        Else
            w_iEndTuki = Month(m_sTuki_Kouki_Start) - 1
        End If

        w_iCnt = w_iEndTuki-w_iStartTuki

        For i = 0 To w_iCnt
            w_iMonth = w_iStartTuki + i
            %>
            <option value="<%=w_iMonth%>"  <%=f_Selected(cint(w_iMonth),cint(Month(date())))%>><%=w_iMonth%>
            <%
        Next

    Else
        '//学期開始月
        w_iStartTuki = Month(m_sTuki_Kouki_Start)

        '//学期終了月
        w_iEndTuki = Month(m_sTuki_Kouki_End)

        w_iCnt = (12+w_iEndTuki) - w_iStartTuki

        For i = 0 To w_iCnt
            w_iMonth = w_iStartTuki + i
            If w_iMonth > 12 Then
                w_iMonth = w_iMonth - 12
            End If
            %>
            <option value="<%=w_iMonth%>"  <%=f_Selected(cint(w_iMonth),cint(Month(date())))%>><%=w_iMonth%>
            <%
        Next

    End If

End Sub

'********************************************************************************
'*  [機能]  ログイン教官の受持教科を取得(年度、教官CD、学期より)
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSubject()

    Dim w_iRet
    Dim w_sSQL
    Dim w_sGakkiKbn '//学期区分

    On Error Resume Next
    Err.Clear

    f_GetSubject = 1

    Do

        '//前後期区分を取得
        If m_sGakki = "ZENKI" Then
            w_sGakkiKbn = cstr(C_GAKKI_ZENKI)   '//1:前期
        Else
            w_sGakkiKbn = cstr(C_GAKKI_KOUKI)   '//2:後期
        End If

        '//受持授業を取得
		'//通常授業と特別活動をUNIONでつないで、抽出する
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT DISTINCT "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_GAKUNEN, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_CLASS, "
        w_sSQL = w_sSQL & "  M05_CLASS.M05_CLASSMEI, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KAMOKU, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KYOKAN, "
        w_sSQL = w_sSQL & "  M03_KAMOKU.M03_KAMOKUMEI , "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_TUKU_FLG"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "  T20_JIKANWARI ,M05_CLASS,M03_KAMOKU"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_NENDO = M05_CLASS.M05_NENDO AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KAMOKU = M03_KAMOKU.M03_KAMOKU_CD AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_NENDO = M03_KAMOKU.M03_NENDO AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_GAKKI_KBN='" & w_sGakkiKbn & "' AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KYOKAN='" & m_iKyokanCd & "' AND "
        'w_sSQL = w_sSQL & "  (T20_JIKANWARI.T20_TUKU_FLG='0' Or T20_JIKANWARI.T20_TUKU_FLG Is Null)"
        '//C_TUKU_FLG_TUJO = "1"(0:通常授業,1:特別活動(HR等))
        w_sSQL = w_sSQL & "  (T20_JIKANWARI.T20_TUKU_FLG='" & C_TUKU_FLG_TUJO & "' Or T20_JIKANWARI.T20_TUKU_FLG Is Null)"
        w_sSQL = w_sSQL & " UNION ALL "
        w_sSQL = w_sSQL & " SELECT  DISTINCT "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_GAKUNEN, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_CLASS, "
        w_sSQL = w_sSQL & "  M05_CLASS.M05_CLASSMEI, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KAMOKU, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KYOKAN, "
        w_sSQL = w_sSQL & "  M41_TOKUKATU.M41_MEISYO, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_TUKU_FLG "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "  T20_JIKANWARI ,M05_CLASS,M41_TOKUKATU"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO "
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN"
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_NENDO = M05_CLASS.M05_NENDO "
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_KAMOKU = M41_TOKUKATU.M41_TOKUKATU_CD "
		w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_NENDO = M41_TOKUKATU.M41_NENDO "
    	w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_NENDO=" & cInt(m_iSyoriNen) & " "
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_GAKKI_KBN='" & w_sGakkiKbn & "' "
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_KYOKAN='" & m_iKyokanCd & "' "
        'w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_TUKU_FLG='1'"   '//0:通常授業,1:特別活動(HR等)
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_TUKU_FLG='" & C_TUKU_FLG_TOKU & "'"
		'//授業区分(C_JUGYO_KBN_JUHYO = 0：授業とみなす, C_JUGYO_KBN_NOT_JUGYO = 1:授業とみなさない)
        w_sSQL = w_sSQL & "  AND M41_TOKUKATU.M41_JUGYO_KBN=" & C_JUGYO_KBN_JUHYO
        w_sSQL = w_sSQL & " ORDER BY T20_GAKUNEN,T20_CLASS"

        iRet = gf_GetRecordset(m_Rs_Sbj, w_sSQL)

'response.write w_sSQL & "<br>"

        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSubject = 99
            Exit Do
        End If

        '//正常終了
        f_GetSubject = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  留学生代替時間割情報を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetRyuDaigae()

    Dim w_iRet
    Dim w_sSQL
    Dim w_sGakkiKbn '//学期区分

    On Error Resume Next
    Err.Clear

    f_GetRyuDaigae = 1

    Do

        '//前後期区分を取得
        If m_sGakki = "ZENKI" Then
            w_sGakkiKbn = cstr(C_GAKKI_ZENKI)   '//1:前期
        Else
            w_sGakkiKbn = cstr(C_GAKKI_KOUKI)   '//2:後期
        End If

        '//受持授業を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & " T27.T27_GAKUNEN, "
        w_sSQL = w_sSQL & " T27.T27_CLASS, "
        w_sSQL = w_sSQL & " M05.M05_CLASSMEI, "
        w_sSQL = w_sSQL & " T27.T27_KAMOKU_CD, "
        w_sSQL = w_sSQL & " T27.T27_KYOKAN_CD, "
        w_sSQL = w_sSQL & " T16.T16_KAMOKUMEI, "
        w_sSQL = w_sSQL & " T16.T16_OKIKAE_FLG "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & " T16_RISYU_KOJIN T16,T27_TANTO_KYOKAN T27 ,M05_CLASS M05"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & " T27.T27_KYOKAN_CD 	= '"	&	m_iKyokanCd			&	"' AND "
        w_sSQL = w_sSQL & " T16.T16_OKIKAE_FLG 	= "		&	C_TIKAN_KAMOKU_SAKI	&	"  AND "
        w_sSQL = w_sSQL & " T27.T27_NENDO		= "		& 	cInt(m_iSyoriNen) 	& 	"  AND "
'        w_sSQL = w_sSQL & " T27.T27_GAKKI_KBN	= '" 	& 	w_sGakkiKbn 		& 	"' AND "
        w_sSQL = w_sSQL & " T16.T16_KAMOKU_CD 	= T27.T27_KAMOKU_CD AND "
        w_sSQL = w_sSQL & " T16.T16_NENDO 		= T27.T27_NENDO AND"
        w_sSQL = w_sSQL & " T27.T27_NENDO 		= M05.M05_NENDO AND "
        w_sSQL = w_sSQL & " T27.T27_GAKUNEN		= M05.M05_GAKUNEN AND "
        w_sSQL = w_sSQL & " T27.T27_CLASS 		= M05.M05_CLASSNO "
'response.write w_ssql
        iRet = gf_GetRecordset(m_Rs_RyuDaigae, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetRyuDaigae = 99
            Exit Do
        End If

        '//正常終了
        f_GetRyuDaigae = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  代替時間割情報を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDaigae()

    Dim w_iRet
    Dim w_sSQL
    Dim w_sGakkiKbn '//学期区分

    On Error Resume Next
    Err.Clear

    f_GetDaigae = 1

    Do

        '//前後期区分を取得
        If m_sGakki = "ZENKI" Then
            w_sGakkiKbn = cstr(C_GAKKI_ZENKI)   '//1:前期
        Else
            w_sGakkiKbn = cstr(C_GAKKI_KOUKI)   '//2:後期
        End If

        '//受持授業を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_NENDO, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_YOUBI_CD, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_JIGEN, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_KAMOKU, "
        w_sSQL = w_sSQL & "  M03_KAMOKU.M03_KAMOKUMEI, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_KYOKAN"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN ,"
        w_sSQL = w_sSQL & "  M03_KAMOKU"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_KAMOKU = M03_KAMOKU.M03_KAMOKU_CD(+) AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_NENDO  = M03_KAMOKU.M03_NENDO(+) AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_NENDO  = " & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_KYOKAN ='" & m_iKyokanCd & "' AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_GAKUNEN Is Null AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_CLASS Is Null AND"
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_GAKKI_KBN='" & w_sGakkiKbn & "'"
'response.write w_ssql
        iRet = gf_GetRecordset(m_Rs_Daigae, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetDaigae = 99
            Exit Do
        End If

        '//正常終了
        f_GetDaigae = 0
        Exit Do
    Loop

End Function

'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'       (リストダウンボックス選択表示用)
'[引数] pData1 : データ１
'       pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'                   
'****************************************************
Function f_Selected(pData1,pData2)

    f_Selected = ""

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else
        End If
    End If

End Function

'********************************************************************************
'*	[機能]	処理年度の時限数の取得
'*	[引数]	
'*	[戻値]	true:成功 false:失敗
'*	[説明]	
'********************************************************************************
function f_Get_JigenData()
	Dim w_sSQL
	Dim w_iRet
	
	On Error Resume Next
	Err.Clear
	
	f_Get_JigenData = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  MAX(M07_JIKAN) "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  MM07_JIGEN "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  M07_NENDO = " & m_iSyoriNen
	
	w_iRet = gf_GetRecordset(m_Rs_Jigen,w_sSQL)
	
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	'//データなし
	if m_Rs_Jigen.EOF then
		f_Get_JigenData = true
		exit function
	end if
	
	m_JigenCount = cInt(m_Rs_Jigen(0))
	
	f_Get_JigenData = true
	
end function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	Dim w_num
%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>授業出欠入力</title>
	
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
    
    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        if (document.frm.KYOUKA.value==""){
            alert("授業データがありません。")
            return ;
        };

        var vl = document.frm.KYOUKA.value.split('#@#');

//        if (vl[0]=='KBTU'){
//            //個別授業(種別、課目ｺｰﾄﾞを取得)
//            document.frm.SYUBETU.value=vl[0];
//            document.frm.KAMOKU_CD.value=vl[1];
//
//            document.frm.GAKUNEN.value=vl[2];
//            document.frm.KAMOKU_NAME.value=vl[3];
//
//            'document.frm.KAMOKU_NAME.value=vl[2];
//
//        }else{
            //通常・特別授業(種別、課目ｺｰﾄﾞ、学年、ｸﾗｽNOを取得)
            document.frm.SYUBETU.value=vl[0];
            document.frm.KAMOKU_CD.value=vl[1];
            document.frm.GAKUNEN.value=vl[2];
            document.frm.CLASSNO.value=vl[3];

            document.frm.CLASS_NAME.value=vl[4];
            document.frm.KAMOKU_NAME.value=vl[5];
//        }

        //document.frm.action = "./kks0110_main.asp";
        document.frm.action="./WaitAction.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  学期を変更した時
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ChangeGakki(){

        //本画面をsubmit
        document.frm.target = "topFrame";
        document.frm.action = "./kks0112_top.asp"
        document.frm.submit();
        return;
    }
	
	//************************************************************
	//	[機能]	カレンダー表示
	//	[引数]	テキストの中身(日付)
	//	[戻値]	
	//	[説明]
	//************************************************************
	function _g_popCalender(w_txtname){
		var wWidth = 250 ;
		var wHeigh = 230 ;
		var wStr = "calender.asp?txtName="  + w_txtname;
		var PositionX,PositionY;
		
		var wCondition = "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,titlebar=0,width=" + wWidth + ",height=" + wHeigh;
		
		PositionX = window.screen.availWidth  / 2 - wWidth / 2;
		PositionY = window.screen.availHeight / 2 - wHeigh / 2;
		
		wCondition = wCondition + ",left=" + PositionX + ",top=" + PositionY;
		
		var popup = window.open(wStr,"calender",wCondition);
		
		if(!popup.opener)popup.opener=self;
		
		if(navigator.appName=="Netscape"&&navigator.appVersion.charAt(0)=="2"){
	 		popup.location.href = location.href.substring(0,location.href.lastIndexOf("/")+1)+"../../common/com_select/SEL_CAL/default.asp";
	 	}
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript" onload="return window_onload();">
    <% call gs_title("授業出欠入力","一　覧") %>
    <form name="frm" method="post">
    <center>
    <table border="0">
	    <tr>
		    <td align="right" class="search" nowrap>

			    <table border="0">
			        <tr>
				        <td nowrap>科目</td>
						<td colspan="7" nowrap>
				            <%
				            '//授業データがない場合
				            If m_Rs_Sbj.EOF And m_Rs_RyuDaigae.EOF Then
				            %>
				            <select name="KYOUKA" style="width:200px;" DISABLED>
				                <option value="">授業データがありません
				            <%Else%>
				            <select name="KYOUKA" style="width:200px;">
				            <%
				            '//========================
				            '//授業時間割データを表示
				            '//========================
				                If m_Rs_Sbj.EOF = False Then
				                    Do Until m_Rs_Sbj.EOF 
				                    If m_Rs_Sbj("T20_TUKU_FLG")="1" Then
				                        '//特別活動の場合
				                        w_Kamoku = m_Rs_Sbj("M03_KAMOKUMEI")
				                        w_Kamoku_CD = m_Rs_Sbj("T20_KAMOKU")
				                        w_Syubetu = "TOKU"  '//特別活動
				                    Else
				                        w_Kamoku = m_Rs_Sbj("M03_KAMOKUMEI")
				                        w_Kamoku_CD = m_Rs_Sbj("T20_KAMOKU")
				                        w_Syubetu = "TUJO"  '//通常授業
				                    End If
				            %>
				                <!--<option value="<%=CStr(w_Syubetu & "#@#" & w_Kamoku_CD & "#@#" & m_Rs_Sbj("T20_GAKUNEN") & "#@#" & m_Rs_Sbj("T20_CLASS"))%>"><%=m_Rs_Sbj("T20_GAKUNEN") & "年&nbsp;&nbsp;" & m_Rs_Sbj("M05_CLASSMEI") & "&nbsp;&nbsp;&nbsp;" & w_Kamoku %>-->
				                <option value="<%=CStr(w_Syubetu & "#@#" & w_Kamoku_CD & "#@#" & m_Rs_Sbj("T20_GAKUNEN") & "#@#" & m_Rs_Sbj("T20_CLASS")) & "#@#" &  m_Rs_Sbj("M05_CLASSMEI") & "#@#" & w_Kamoku%>"><%=m_Rs_Sbj("T20_GAKUNEN") & "年&nbsp;&nbsp;" & m_Rs_Sbj("M05_CLASSMEI") & "&nbsp;&nbsp;&nbsp;" & w_Kamoku %>
				            <%
				                    m_Rs_Sbj.MoveNext
				                    Loop
				                End If

				                '//===========================
				                '//留学生の代替時間割データを追加表示
				                '//===========================
				                If m_Rs_RyuDaigae.EOF = False Then
				                    Do Until m_Rs_RyuDaigae.EOF 
				                    w_Syubetu = "KBTU"  '//個別授業
				            %>
							<option value="<%=CStr(w_Syubetu & "#@#" & m_Rs_RyuDaigae("T27_KAMOKU_CD") & "#@#" & m_Rs_RyuDaigae("T27_GAKUNEN") & "#@#" & m_Rs_RyuDaigae("T27_CLASS")) & "#@#" &  m_Rs_RyuDaigae("M05_CLASSMEI") & "#@#" & m_Rs_RyuDaigae("T16_KAMOKUMEI")%>">留学生代替&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=m_Rs_RyuDaigae("T16_KAMOKUMEI")%>
				            <%
				                    m_Rs_RyuDaigae.MoveNext
				                    Loop
				                End If
				            End If

				                '//===========================
				                '//代替時間割データを追加表示
				                '//===========================
'				                If m_Rs_Daigae.EOF = False Then
'				                    Do Until m_Rs_Daigae.EOF 
'				                    w_Syubetu = "KBTU"  '//個別授業
'				            %>

				            <!--option Value="<=CStr(w_Syubetu & "#@#" & w_Kamoku_CD) & "#@#" & m_Rs_Daigae("M03_KAMOKUMEI")>">個別授業&nbsp;&nbsp;&nbsp;<=m_Rs_Daigae("M03_KAMOKUMEI")-->






				            <%
				                    'm_Rs_Daigae.MoveNext
				                    'Loop
'				                End If
'				            End If
				            %>
				            </select>
				        </td>
				    </tr>
				    
				    <tr><td colspan="8"><hr></td></tr>
				    
				    <tr>
				    	<th class="header" align="center">欠課入力</th>
				    	<td>日付　
				    		<input type="text" name="txtDate" value="<%=gf_YYYY_MM_DD(date(),"/")%>">
				    		<input type="button" class="button" onClick="fcalender('txtDate')" value="選択">
				    	</td>
				    	
				    	<td>時限
				    		<select name="sltJigen">
				    		
				    		<% for w_num=1 to m_JigenCount %>
				    			<option value="<%=w_num%>"><%=w_num%>
				    		<% next %>
				    	</td>
				    	
				    	<td align="right" nowrap>
							<input class="button" type="button" onclick="javascript:f_Search();" value="　入　力　">
						</td>
				    </tr>
				    
				    <tr>
				    	<th class="header" colspan="" align="center">欠課一覧参照</th>
				    	
				    	<td>指定月
				    		<select name="sltMonth">
				    			<option value="4">4
				    			<option value="5">5
				    			<option value="6">6
				    			<option value="7">7
				    			<option value="8">8
				    			<option value="9">9
				    			<option value="10">10
				    			<option value="11">11
				    			<option value="12">12
				    			<option value="1">1
				    			<option value="2">2
				    			<option value="3">3
				    		</select>
				    	</td>
				    	
				    	<td colspan="2" align="right" nowrap>
							<input class="button" type="button" onclick="javascript:f_Search();" value="　表　示　">
						</td>
				    </tr>
				    
				    
			    </table>

		    </td>
	    </tr>
    </table>

    <!--値渡し用-->
    <input type="hidden" name="Tuki_Zenki_Start" value="<%=m_sTuki_Zenki_Start%>">
    <input type="hidden" name="Tuki_Kouki_Start" value="<%=m_sTuki_Kouki_Start%>">
    <input type="hidden" name="Tuki_Kouki_End"   value="<%=m_sTuki_Kouki_End%>">
    <INPUT TYPE=HIDDEN NAME="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE=HIDDEN NAME="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
    <INPUT TYPE=HIDDEN NAME="GAKUNEN"   value = "">
    <INPUT TYPE=HIDDEN NAME="CLASSNO"   value = "">
    <INPUT TYPE=HIDDEN NAME="KAMOKU_CD" value = "">
    <INPUT TYPE=HIDDEN NAME="SYUBETU"   value = "">

    <INPUT TYPE=HIDDEN NAME="KAMOKU_NAME"   value = "">
    <INPUT TYPE=HIDDEN NAME="CLASS_NAME" value = "">

    <input TYPE="HIDDEN" NAME="txtURL" VALUE="kks0112_bottom.asp">
    <input TYPE="HIDDEN" NAME="txtMsg" VALUE="しばらくお待ちください">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>