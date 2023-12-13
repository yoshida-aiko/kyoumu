<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0110/kks0110_main.asp
' 機      能: 下ページ 授業出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             TUKI           '//月
' 変      数:
' 引      渡: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             TUKI           '//月
' 説      明:
'           ■初期表示
'               検索条件にかなう行事出欠入力を表示
'           ■登録ボタンクリック時
'               入力情報を登録する
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bDaigae            '代替留学生取得ﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public m_iSyoriNen      '//処理年度
    Public m_iKyokanCd      '//教官CD
    Public m_sGakunen       '//学年
    Public m_sClassNo       '//ｸﾗｽNO
    Public m_sTuki          '//月
    Public m_sZenki_Start   '//前期開始日
    Public m_sKouki_Start   '//後期開始日
    Public m_sKouki_End     '//後期終了日

    Public m_sGakki         '//学期
    Public m_sGakki_Kbn     '//学期区分
    Public m_sKamokuCd      '//課目CD
    Public m_sSyubetu       '//授業種別(TUJO:通常授業,TOKU:特別活動,KBTU:個別授業)
    Public m_sHissenKbn     '//必選区分
	Public m_iTani			'//１時限の単位数
	Public m_bEndFLG		'//すべて登録不可の場合TRUE

    Public m_AryHead()      '//ヘッダ情報格納配列
    Public m_iRsCnt         '//ヘッダﾚｺｰﾄﾞ数
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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

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
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '//変数初期化
        Call s_ClearParam()

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

'//デバッグ
'Call s_DebugPrint()

        '// ヘッダリスト情報取得
        w_iRet = f_Get_HeadData()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'//授業データがない場合
        'If m_iRsCnt < 0 Then
        If trim(request("txtMsg")) <> "" Then
			'//空白ページ表示
			Call showWhitePage(trim(request("txtMsg")))
            Exit Do
		Else
	        '// ページを表示
	        Call showPage()
        End If

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
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sTuki     = ""
    m_sGakki    = ""
    m_sKamokuCd = ""
    m_sSyubetu  = ""
	m_iTani		= ""
	m_bEndFLG	= true

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_sZenki_Start = trim(Request("Tuki_Zenki_Start"))
    m_sKouki_Start = trim(Request("Tuki_Kouki_Start"))
    m_sKouki_End   = trim(Request("Tuki_Kouki_End"))
	m_iTani = Session("JIKAN_TANI") '１時限の単位数

    m_iSyoriNen = trim(Request("NENDO"))
    m_iKyokanCd = trim(Request("KYOKAN_CD"))

    m_sTuki     = trim(Request("TUKI"))
    m_sGakki    = trim(Request("GAKKI"))

    m_sSyubetu  = trim(Request("SYUBETU"))
    m_sGakunen  = trim(Request("GAKUNEN"))
    m_sClassNo  = trim(Request("CLASSNO"))
    m_sKamokuCd = trim(Request("KAMOKU_CD"))

	m_bEndFLG	= Cbool(Request("EndFLG"))

    If m_sGakki = "ZENKI" Then
        m_sGakki_Kbn = cstr(C_GAKKI_ZENKI)
    Else
        m_sGakki_Kbn = cstr(C_GAKKI_KOUKI)
    End If

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
Exit Sub
    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"
    response.write "m_sTuki     = " & m_sTuki     & "<br>"
    response.write "m_sGakki    = " & m_sGakki    & "<br>"
    response.write "m_sKamokuCd = " & m_sKamokuCd & "<br>"
    response.write "m_sSyubetu  = " & m_sSyubetu  & "<br>"
    response.write "m_sZenki_Start = " & m_sZenki_Start & "<br>"
    response.write "m_sKouki_Start = " & m_sKouki_Start & "<br>"
    response.write "m_sKouki_End   = " & m_sKouki_End   & "<br>"

End Sub

'********************************************************************************
'*  [機能]  日付・曜日・時間のヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_HeadData()

    Dim w_sSQL
    Dim w_Rs

    On Error Resume Next
    Err.Clear
    
    f_Get_HeadData = 1

    Do 

        '//日付の範囲をセット
        Call f_GetTukiRange(w_sSDate,w_sEDate)
		
		'// 授業日付、時間データ

		'// 授業種別が個人授業（KBTU）の時は代替時間割から取得する。
		'// 2001/12/18 add
		If m_sSyubetu <> "KBTU" Then 

        w_sSQL = ""
		'// 通常、特別授業の場合
        w_sSQL = w_sSQL & vbCrLf & " SELECT"
        w_sSQL = w_sSQL & vbCrLf & "  A.T32_HIDUKE,"
        w_sSQL = w_sSQL & vbCrLf & "  B.T20_JIGEN AS JIGEN,"
        w_sSQL = w_sSQL & vbCrLf & "  B.T20_YOUBI_CD AS YOUBI_CD"
        w_sSQL = w_sSQL & vbCrLf & " FROM"
        w_sSQL = w_sSQL & vbCrLf & " T32_GYOJI_M A"
        w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI B"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & " B.T20_YOUBI_CD = A.T32_YOUBI_CD "

		'//特別活動の場合は、すぐ次の授業の行事情報を見て、行事かどうかを判断する
		If m_sSyubetu = "TOKU"Then
	        w_sSQL = w_sSQL & vbCrLf & " AND TRUNC(B.T20_JIGEN+0.5) = A.T32_JIGEN"
		Else
	        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_JIGEN = A.T32_JIGEN"
		End If

        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_NENDO = A.T32_NENDO"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE>='" & w_sSDate & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE<'"  & w_sEDate & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_NENDO="      & cInt(m_iSyoriNen)
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_GAKKI_KBN='" & m_sGakki_Kbn & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_GAKUNEN= "   & cInt(m_sGakunen)
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_CLASS= "     & cInt(m_sClassNo)
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_KAMOKU='"    & trim(m_sKamokuCd) & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_KYOKAN='"    & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_GYOJI_CD=0"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='0' "
        w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE,B.T20_YOUBI_CD,B.T20_JIGEN "
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE,B.T20_JIGEN"
		Else
			'// 留学生の代替科目の場合
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " SELECT"
			w_sSQL = w_sSQL & vbCrLf & "  A.T32_HIDUKE,"
			w_sSQL = w_sSQL & vbCrLf & "  B.T23_JIGEN AS JIGEN,"
			w_sSQL = w_sSQL & vbCrLf & "  B.T23_YOUBI_CD AS YOUBI_CD"
			w_sSQL = w_sSQL & vbCrLf & " FROM"
			w_sSQL = w_sSQL & vbCrLf & " T32_GYOJI_M A"
			w_sSQL = w_sSQL & vbCrLf & " ,T23_DAIGAE_JIKAN B"
			w_sSQL = w_sSQL & vbCrLf & " WHERE "
			w_sSQL = w_sSQL & vbCrLf & " B.T23_YOUBI_CD = A.T32_YOUBI_CD "
	'		 w_sSQL = w_sSQL & vbCrLf & " AND B.T23_JIGEN = A.T32_JIGEN"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_NENDO = A.T32_NENDO"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE>='" & w_sSDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE<'"  & w_sEDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_NENDO="		& cInt(m_iSyoriNen)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_GAKKI_KBN=" & m_sGakki_Kbn & " "
'			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_GAKUNEN= "	& cInt(m_sGakunen)
'			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_CLASS= " 	& cInt(m_sClassNo)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_KAMOKU='"	& trim(m_sKamokuCd) & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_KYOKAN='"	& m_iKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_GYOJI_CD=0"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='0' "
			w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE,B.T23_YOUBI_CD,B.T23_JIGEN "
			w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE,B.T23_JIGEN"
		End If

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_HeadData = 99
            Exit Do
        End If

        m_iRsCnt = 0

        '=======================
        '//時間割を配列にセット
        '=======================
        If w_Rs.EOF = false Then

            i = 0
            w_sHi = ""
            w_Rs.MoveFirst
            Do Until w_Rs.EOF
				
                '//取得した日付の時限が休日または、行事の場合(w_bGyoji=True)ははじく
                iRet = f_Get_DateInfo(w_Rs("T32_HIDUKE"),w_Rs("JIGEN"),w_bGyoji)
                If iRet <> 0 Then
                    msMsg = Err.description
                    f_Get_HeadData = 99
                    Exit Do
                End If

                '//休日・行事以外のみデータをセット
                If w_bGyoji <> True Then

                    '//配列を設定
                    ReDim Preserve m_AryHead(4,i)

                    '//データ格納
                    If w_sHi = gf_SetNull2String(w_Rs("T32_HIDUKE")) Then
                        m_AryHead(0,i) = ""     '//月
                        m_AryHead(1,i) = ""     '//日
                        m_AryHead(2,i) = ""     '//曜日CD
                    Else
                        m_AryHead(0,i) = month(gf_SetNull2String(w_Rs("T32_HIDUKE")))     '//月
                        m_AryHead(1,i) = day(gf_SetNull2String(w_Rs("T32_HIDUKE")))       '//日
                        m_AryHead(2,i) = gf_SetNull2String(w_Rs("YOUBI_CD"))          '//曜日CD
                    End If

                    m_AryHead(3,i) = gf_SetNull2String(w_Rs("JIGEN"))    '//時限
                    m_AryHead(4,i) = gf_SetNull2String(w_Rs("T32_HIDUKE"))   '//日付

                    w_sHi = gf_SetNull2String(w_Rs("T32_HIDUKE"))
                    i = i + 1

                End If

                w_Rs.MoveNext
            Loop

        End If

        '//取得したデータ数をセット
        m_iRsCnt = i-1

        '//正常終了
        f_Get_HeadData = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
   Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  取得した日付・時限が、休日または行事でないか
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_DateInfo(p_Hiduke,p_Jigen,p_bGyoji)

    Dim w_sSQL
    Dim w_Rs
    Dim w_bGyoujiFlg

    On Error Resume Next
    Err.Clear
    
    f_Get_DateInfo = 1
    w_bGyojiFlg = False

    Do 

        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT"
        w_sSQL = w_sSQL & vbCrLf & " A.T32_GYOJI_CD"
        w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M A"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  A.T32_NENDO=2001 "
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_GAKUNEN IN (" & cInt(m_sGakunen) & "," & C_GAKUNEN_ALL & ")"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_CLASS IN ("   & cInt(m_sClassNo) & "," & C_CLASS_ALL   & ")"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_HIDUKE='" & p_Hiduke & "'"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_JIGEN=" & p_Jigen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_COUNT_KBN<>" & C_COUNT_KBN_JUGYO
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_KYUJITU_FLG<>'" & C_HEIJITU & "'"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_DateInfo = 99
            Exit Do
        End If

        If w_Rs.EOF = False Then
            '//ﾚｺｰﾄﾞがある場合は休日か、行事の日
            w_bGyojiFlg = True
        End If

        f_Get_DateInfo = 0
        Exit Do
    Loop

        '//戻り値をセット
        p_bGyoji = w_bGyojiFlg

        '//ﾚｺｰﾄﾞｾｯﾄCLOSE
       Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  月の検索条件を作成(7月…　"MONTH>=2001/07/01 AND MONTH<2001/08/01" として使用)
'*  [引数]  なし
'*  [戻値]  p_sSDate
'*          p_sEDate
'*  [説明]  
'********************************************************************************
Function f_GetTukiRange(p_sSDate,p_sEDate)

    p_sSDate = ""
    p_sEDate = ""

    If m_sGakki = "ZENKI" Then
        w_iNen = cint(m_iSyoriNen)

	    '//開始日
		If cint(month(m_sZenki_Start)) = Cint(m_sTuki) Then
			p_sSDate = m_sZenki_Start
		Else
			p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki),2) & "/01"
		End If

	    '//終了日
		If cint(month(m_sKouki_Start)) = Cint(m_sTuki) Then
			p_sEDate = m_sKouki_Start
		Else 
		    If Cint(m_sTuki) = 12 Then
		        p_sEDate = cstr(w_iNen+1) & "/01/01"
		    Else
		        p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki+1),2) & "/01"
		    End If
		End If

    Else
		'//後期の年
        If cint(m_sTuki) <=4 Then
            w_iNen = cint(m_iSyoriNen) + 1
        Else
            w_iNen = cint(m_iSyoriNen)
        End If

	    '//開始日
		If cint(month(m_sKouki_Start)) = Cint(m_sTuki) Then
		    p_sSDate = m_sKouki_Start
		Else
		    p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki),2) & "/01"
		End If

	    '//終了日
		If cint(month(m_sKouki_End)) = Cint(m_sTuki) Then
			'p_sEDate = m_sKouki_End
			p_sEDate = DateAdd("d",1,m_sKouki_End)
		Else 
		    If Cint(m_sTuki) = 12 Then
		        p_sEDate = cstr(w_iNen+1) & "/01/01"
		    Else
		        p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki+1),2) & "/01"
		    End If
		End If

    End If

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	dim w_str	'表示メッセージ

    On Error Resume Next
    Err.Clear

	if m_bEndFLG = False then 'まだ登録可能
		If m_iTani > 1 then '１時限の単位数が２以上の時
			w_str = "<span class=CAUTION>※ 氏名の横の空白欄をクリックして、出欠状況を入力してください。（欠→遅→早→１→空欄(出席)の順で表示されます）<BR></span>" & vbCrLf
			w_str = w_str & "【 欠：欠課(" & m_iTani & "欠課分)　１：欠課(１欠課分)　遅：遅刻　早：早退 】" & vbCrLf
		Else				'ノーマル状態
			w_str = "<span class=CAUTION>※ 氏名の横の空白欄をクリックして、出欠状況を入力してください。（欠→遅→早→空欄(出席)の順で表示されます）</span>" & VbCrLf
			w_str = w_str & "【 欠：欠課　遅：遅刻　早：早退 】" & vbCrLf
		End If

	Else '登録期間を過ぎて登録不可能の場合は、参照のみ
		If m_iTani > 1 then 
			w_str = "【 欠：欠課(" & m_iTani & "欠課分)　１：欠課(１欠課分)　遅：遅刻　早：早退 】" & vbCrLf
		Else 
			w_str = "【 欠：欠課　遅：遅刻　早：早退 】" & vbCrLf
		End If 
	End If

%>
    <html>
    <head>
    <title>授業出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

		//スクロール同期制御
		parent.init();

    }

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Touroku(){

		parent.frames["main"].f_Touroku();
		return;
    }

    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Cancel(){
        //空白ページを表示
        parent.document.location.href="default.asp"
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    <%call gs_title("授業出欠入力","一　覧")%>
    <%Do %>
        <%If m_iRsCnt < 0 Then%>
            <br><br>
            <span class="msg">授業日程がありません</span>
            <%Exit Do%>
        <%End If%>

<!-------------------------追加分ヘッダ----------------------------------->

        <table>
		<tr><td>
	        <table class="hyo" border="1" width="545">
	            <tr>

					<%
					'//個別授業の場合は年、クラスがない
					If trim(m_sSyubetu) = "KBTU" Then
						w_sClass_Name = "-"
					Else
						w_sClass_Name = Request("CLASS_NAME")
					End If

					If m_sGakki_Kbn = cstr(C_GAKKI_ZENKI) Then
						w_sGakki = "前期"
					Else
						w_sGakki = "後期"
					End If

					%>
	                <th nowrap class="header" width="65"  align="center"><%=w_sGakki%></th>
	                <td nowrap class="detail" width="50"  align="center"><%=m_sTuki%>月</td>
	                <th nowrap class="header" width="65"  align="center">クラス</th>
	                <td nowrap class="detail" width="150"  align="center"><%=m_sGakunen%>年　<%=w_sClass_Name%></td>
	                <th nowrap class="header" width="65" align="center">授業</th>
	                <td nowrap class="detail" width="150" align="center"><%=Request("KAMOKU_NAME")%></td>

	            </tr>
	        </table>
		</td></tr><tr>
		<td align="center">
			<table>
				<tr>
<%	if m_bEndFLG = False then 'まだ登録可能 %>
		        <td valign="bottom"align="center">
		            <input class="button" type="button" onclick="javascript:f_Touroku();" value="　登　録　">
		            &nbsp;&nbsp;&nbsp;
		            <input class="button" type="button" onclick="javascript:f_Cancel();" value="キャンセル">
		        </td>
<% Else %>
		        <td valign="bottom"align="center">
		            <input class="button" type="button" onclick="javascript:f_Cancel();" value=" 戻　る ">
		        </td>
<% End If %>
				</tr>
	        </table>
		</td></tr>
        </table>

<!-------------------------追加分ヘッダ----------------------------------->

		<table>
        <tr>
<% '表示メッセージ　パターンにより表示文字可変 %>
            <td align="center" colspan=3><font size="2" color="#222268"><%=w_str%></font>
            </td>
        </tr>

		</table>

        <!--明細ヘッダ部(月・曜日・時限等を表示)-->
        <table >
        <tr>
            <td align="center" valign="top">
            <table class="hyo"  border="1" >

            <tr>
                <th class="header" height="100" rowspan="4" width="50" align="center"  nowrap><font >
                    <table ><tr><th width="10" class="header" nowrap><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th></tr></table></font>
                </th>
                <th class="header" width="150" align="center" nowrap><font color="#ffffff">月</font></th>
                <%for i = 0 to m_iRsCnt%>
                    <th class="header" align="center"  nowrap><font ><%=m_AryHead(0,i)%><br></font></th>

                <%Next%>
            </tr>

            <tr>
                <th class="header" align="center" nowrap ><font color="#ffffff">日</font></th>
                <%for i = 0 to m_iRsCnt%>
                    <th class="header" align="center" width="30"  nowrap ><font color="#ffffff"><%=m_AryHead(1,i)%></font></th>
                <%Next%>
            </tr>

            <tr>
                <th class="header" align="center" nowrap><font color="#ffffff">曜日</font></th>
                <%for i = 0 to m_iRsCnt%>
                    <th class="header" align="center" width="30"  nowrap ><font color="#ffffff"><%=gf_GetYoubi(m_AryHead(2,i))%></font></th>
                <%Next%>
            </tr>

            <tr>
                <th class="header" align="center" nowrap><font color="#ffffff">時限</font></th>
                <%for i = 0 to m_iRsCnt%>
					<%
					'//授業外活動の場合は、時限を表示しない
					If m_sSyubetu = "TOKU"Then
						w_iDispJigen = "-"
					Else
						w_iDispJigen = m_AryHead(3,i)
					End If
					%>
                    <th class="header" align="center" width="30"  nowrap ><font color="#ffffff"><%=w_iDispJigen%></font>
                    </th>
                <%Next%>
            </tr>
            </table>

        </td>
        <td width="10"><br></td>
        <td align="center" valign="top" width="120" nowrap>

            <!--月・学期の欠席及び遅刻数累計-->
            <table width="120" class="hyo" border="1">
            <tr>
                <th height="20" colspan="2" class="header" nowrap align="center" width="60"><font color="#ffffff">月計</font></th>
                <th height="20" colspan="2" class="header" nowrap align="center" width="60"><font color="#ffffff">計</font></th>
            </tr>
            <tr>
                <th height="80" class="header" width="30" align="center" nowrap><font color="#ffffff">遅<br>刻</font></th>
                <th height="80" class="header" width="30" align="center" nowrap><font color="#ffffff">欠<br>課</font></th>
                <th height="80" class="header" width="30" align="center" nowrap><font color="#ffffff">遅<br>刻</font></th>
                <th height="80" class="header" width="30" align="center" nowrap><font color="#ffffff">欠<br>課</font></th>
            </tr>
            </table>

        </td>
        </tr>
        </table>

        <%Exit Do%>
    <%Loop%>

    </form>
    </center>
    </body>
    </html>
<%
End Sub

'********************************************************************************
'*  [機能]  空白HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
    <html>
    <head>
    <title>授業出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		parent.document.location.href="default2.asp?txtMsg=<%=Server.URLEncode(p_Msg)%>"
		return;
    }
    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">

    </body>
    </html>
<%
End Sub
%>

