<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績一覧
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0200/sei0200_main.asp
' 機      能: 下ページ 成績一覧の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 説      明:
'           ■初期表示
'				コンボボックスは空白で表示
'			■表示ボタンクリック時
'				下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/08/08 前田 智史
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bErrMsg           'ｴﾗｰﾒｯｾｰｼﾞ

	'氏名選択用のWhere条件
    Public m_iNendo			'年度
    Public m_sKyokanCd		'教官コード
    Public m_sSikenKBN		'試験区分
    Public m_sGakuNo		'学年
    Public m_sClassNo		'学科
    Public m_sKBN			'区分コード
    Public m_sSeiseki		
	Dim	   m_rCnt			'レコードカウント
	Dim	   m_SrCnt			'レコードカウント

	'配列用
	Dim	   m_iKamokuCd()	'm_Rsの科目コードの配列
	Dim	   m_sKamokuNm()	'm_Rsの科目名の配列
	Dim	   m_iHTani()		'm_Rsの配当単位の配列
	Dim	   m_iTensuu()		'各科目点数の配列
	Dim	   m_iGakusei()		'm_SRsの学生コードの配列
	Dim	   m_iGakuseki()	'm_SRsの学籍コードの配列
	Dim	   m_sSimei()		'm_SRsの氏名の配列

	Public	m_Rs
	Public	m_KRs
	Public	m_TRs
	Public	m_SRs
	Public	m_iMax			'最大ページ

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
	w_sMsgTitle="成績一覧"
	w_sMsg=""
	w_sRetURL="default.asp"     
	w_sTarget="_parent"


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

	    '// ﾊﾟﾗﾒｰﾀSET
	    Call s_SetParam()

		'//科目データ取得
        Call f_getdate()

		If m_rCnt = 0 Then
			Call ShowPage_No()
			Exit Do
		End If

		'//学生データ取得
        Call f_getGaku()

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
    Call gf_closeObject(m_KRs)
    Call gs_CloseDatabase()
End Sub

Sub s_SetParam()
'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSikenKBN	= Cint(request("txtSikenKBN"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKBN	= cint(request("txtKBN"))

End Sub

Function f_getdate()
'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Dim i
i = 1
'***
Dim w_sKaisetu
w_sKaisetu = "T15_KAISETU"& Cint(m_sGakuNo) '開設区分フィールド名作成
'***
	On Error Resume Next
	Err.Clear
	f_getdate = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "    A.T16_KAMOKU_CD,A.T16_KAMOKUMEI,A.T16_HAITOTANI,A.T16_KAMOKU_KBN,A.T16_SEQ_NO "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T16_RISYU_KOJIN A,T13_GAKU_NEN B, M05_CLASS C"
		w_sSQL = w_sSQL & vbCrLf & " WHERE"
		w_sSQL = w_sSQL & vbCrLf & " A.T16_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & " AND A.T16_HISSEN_KBN=" & Cint(m_sKBN)
		w_sSQL = w_sSQL & vbCrLf & " AND C.M05_GAKUNEN=" & m_sGakuNo 
		w_sSQL = w_sSQL & vbCrLf & " AND C.M05_CLASSNO=" & m_sClassNo
		w_sSQL = w_sSQL & vbCrLf & " AND A.T16_NENDO = B.T13_NENDO "
		w_sSQL = w_sSQL & vbCrLf & " AND A.T16_GAKUSEI_NO = B.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & vbCrLf & " AND A.T16_HAITOGAKUNEN = B.T13_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & " AND C.M05_GAKKA_CD = B.T13_GAKKA_CD "
		w_sSQL = w_sSQL & vbCrLf & " AND C.M05_CLASSNO = B.T13_CLASS "
		w_sSQL = w_sSQL & vbCrLf & " AND C.M05_GAKUNEN = B.T13_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T16_KAMOKU_CD,A.T16_KAMOKUMEI,A.T16_HAITOTANI,A.T16_KAMOKU_KBN,A.T16_SEQ_NO "
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T16_KAMOKU_KBN,A.T16_SEQ_NO "

'response.write w_sSql

		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getdate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		'//学科を取得
		w_iRet = f_Get_Gakka(m_sGakuNo,m_sClassNo,w_iGakkaCd)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getdate = 99
			m_bErrFlg = True
			Exit Do 
		End If

'		m_rCnt=cint(gf_GetRsCount(m_Rs))

		m_Rs.MoveFirst

        Do Until m_Rs.EOF

	        ReDim Preserve m_iKamokuCd(i)
	        ReDim Preserve m_sKamokuNm(i)
	        ReDim Preserve m_iHTani(i)

			'//科目の開設情報を見る(開設しない場合は、科目を表示しない)
			w_iRet = f_Get_KaisetuInfo(m_sGakuNo,w_iGakkaCd,m_Rs("T16_KAMOKU_CD"),w_iKaisetu)
			If w_iRet <> 0 then
				f_getdate = 99
				m_bErrFlg = True
				Exit Do
			End If

			'//科目の担当教官が設定されているか(担当教官が設定されていない場合は、科目を表示しない)
			w_iRet = f_Get_KamokuTantoInfo(m_sGakuNo,m_sClassNo,m_Rs("T16_KAMOKU_CD"),w_bTanto)
			If w_iRet <> 0 then
				f_getdate = 99
				m_bErrFlg = True
				Exit Do
			End If

			'//開設区分が開設しないデータ(C_KAI_NASI=3 : 開設しない),及び科目担当教官が設定されていないデータは表示しない
			If cint(gf_SetNull2Zero(w_iKaisetu)) <> C_KAI_NASI AND w_bTanto = True then 
			'If cint(gf_SetNull2Zero(w_iKaisetu)) <> C_KAI_NASI then 
	            m_iKamokuCd(i) = m_Rs("T16_KAMOKU_CD")
	            m_sKamokuNm(i) = m_Rs("T16_KAMOKUMEI")
	            m_iHTani(i) = m_Rs("T16_HAITOTANI")
	            i = i + 1
			End If

            m_Rs.MoveNext
	
        Loop

		'//エラー時
		If m_bErrFlg = True Then
			Exit Do 
		End If

		'//データ数をセット
		m_rCnt = i-1

		f_getdate = 0
		Exit Do

	Loop

    Call gf_closeObject(m_Rs)

End Function

'********************************************************************************
'*  [機能]  クラスの学科を取得
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_Gakka(p_iGakuNen,p_iClassNo,p_iGakkaCd)

	Dim w_sSQL
	Dim w_Rs

	On Error Resume Next
	Err.Clear
	
	f_Get_Gakka = 1
	p_iGakkaCd = ""

	Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKKA_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN=" & p_iGakuNen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_CLASSNO=" & p_iClassNo

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_Gakka = 99
			Exit Do
		End If

		If w_Rs.EOF = False Then
			'//ﾚｺｰﾄﾞがある場合は休日か、行事の日
			p_iGakkaCd = w_Rs("M05_GAKKA_CD")
		End If

		f_Get_Gakka = 0
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
Function f_Get_KaisetuInfo(p_iGakuNen,p_iGakkaCd,p_sKamokuCd,p_iKaisetu)

	Dim w_sSQL
	Dim w_Rs
	Dim w_bGyoujiFlg

	On Error Resume Next
	Err.Clear
	
	f_Get_KaisetuInfo = 1
	w_iKaisetu = ""

	Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAISETU" & p_iGakuNen
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & Cint(m_iNendo) - cint(p_iGakuNen) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & p_iGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_sKamokuCd & "'"

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_KaisetuInfo = 99
			Exit Do
		End If

		If w_Rs.EOF = False Then
			'//学年に対応した、開設区分を取得
			w_iKaisetu = w_Rs("T15_KAISETU" & p_iGakuNen)
		End If

		f_Get_KaisetuInfo = 0
		Exit Do
	Loop

	'//戻り値をセット
	p_iKaisetu = w_iKaisetu

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  取得した科目の担当教官が設定されているか
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_KamokuTantoInfo(p_iGakuNen,p_iClassNo,p_sKamokuCd,p_bTanto)

	Dim w_sSQL
	Dim w_Rs
	Dim w_bGyoujiFlg

	On Error Resume Next
	Err.Clear
	
	f_Get_KamokuTantoInfo = 1
	p_bTanto = False

	Do 
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN.T27_KYOKAN_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN.T27_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_GAKUNEN=" & p_iGakuNen
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_CLASS=" & p_iClassNo
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_KAMOKU_CD='" & p_sKamokuCd & "'"

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_Get_KamokuTantoInfo = 99
			Exit Do
		End If

		If w_Rs.EOF = False Then
			p_bTanto = True
		End If

		f_Get_KamokuTantoInfo = 0
		Exit Do
	Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_Rs)

End Function

Function f_getGaku()
'********************************************************************************
'*	[機能]	学生の取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Dim w_iWkTensuu
Dim w_iGakuseiNo
Dim w_iKamokuIdx
Dim w_iDspFlg
Dim w_rs
Dim w_rCnt

	On Error Resume Next
	Err.Clear
	f_getGaku = 1

	Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakuNo) + 1

		'//検索結果の値より一覧を表示
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	A.T16_GAKUSEI_NO,A.T16_GAKUSEKI_NO,B.T11_SIMEI "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T16_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_NENDO = C.T13_NENDO "
'		w_sSQL = w_sSQL & " AND	B.T11_NYUNENDO = " & w_iNyuNendo & " "
		w_sSQL = w_sSQL & " GROUP BY A.T16_GAKUSEI_NO,A.T16_GAKUSEKI_NO,B.T11_SIMEI "
		w_sSQL = w_sSQL & " ORDER BY A.T16_GAKUSEKI_NO "

		Set w_rs = Server.CreateObject("ADODB.Recordset")
		w_iRet = gf_GetRecordset(w_rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getGaku = 99
			m_bErrFlg = True
			Exit Do 
		End If
		w_rCnt=cint(gf_GetRsCount(w_rs))

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	A.T16_GAKUSEI_NO,A.T16_GAKUSEKI_NO,B.T11_SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & " 	A.T16_KAMOKU_CD,A.T16_SEI_TYUKAN_Z,A.T16_SEI_KIMATU_Z,A.T16_SEI_TYUKAN_K,A.T16_SEI_KIMATU_K, "
		w_sSQL = w_sSQL & vbCrLf & " 	A.T16_SELECT_FLG,A.T16_HISSEN_KBN "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T16_RISYU_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & vbCrLf & " WHERE"
		w_sSQL = w_sSQL & vbCrLf & " 	A.T16_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & vbCrLf & " AND	A.T16_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & vbCrLf & " AND	A.T16_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & vbCrLf & " AND	C.T13_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & vbCrLf & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & vbCrLf & " AND	A.T16_NENDO = C.T13_NENDO "
'		w_sSQL = w_sSQL & vbCrLf & " AND	B.T11_NYUNENDO = " & w_iNyuNendo & " "

		'If m_sKBN <> Cint(C_HISSEN_HIS) Then
		'	w_sSQL = w_sSQL & "	AND T16_SELECT_FLG = " & C_SENTAKU_YES & " "
		'End If
		w_sSQL = w_sSQL & " ORDER BY A.T16_GAKUSEKI_NO "

'response.write w_sSQL & "<BR>"

		Set m_SRs = Server.CreateObject("ADODB.Recordset")
		w_iRet = gf_GetRecordset(m_SRs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getGaku = 99
			m_bErrFlg = True
			Exit Do 
		End If
		'm_SrCnt=cint(gf_GetRsCount(m_SRs))

	'//配列の作成

		m_SRs.MoveFirst

		'//点数配列の初期化
        ReDim Preserve m_iTensuu(m_rCnt,w_rCnt)
		For j=1 to w_rCnt
			For i=1 to m_rCnt
				m_iTensuu(i,j) = "-"
			Next
		Next

		m_SrCnt = 0
		w_iGakuseiNo = 0
        Do Until m_SRs.EOF

			'// 学生番号が変わったら
			If w_iGakuseiNo <> m_SRs("T16_GAKUSEI_NO") Then
				m_SrCnt = m_SrCnt + 1

'Response.write "m_SrCnt=[" & m_SrCnt & "] "
'Response.write "GAKUSEI_NO=[" & m_SRs("T16_GAKUSEI_NO") & "] "
'Response.write "SIMEI=[" & m_SRs("T11_SIMEI") & "] "
'Response.write "KAMOKU_CD=[" & m_SRs("T16_KAMOKU_CD") & "] "
'Response.write "w_iKamokuIdx=[" & w_iKamokuIdx & "] "
'Response.write "TYUKAN_Z=[" & m_SRs("T16_SEI_TYUKAN_Z") & "] "
'Response.write "w_iWkTensuu=[" & w_iWkTensuu & "] "
'Response.write "m_iTensuu(" & w_iKamokuIdx & "," & m_SrCnt & ")=[" & m_iTensuu(w_iKamokuIdx,m_SrCnt) & "]"
'Response.write "m_iTensuu()=[" & m_iTensuu(w_iKamokuIdx,m_SrCnt) & "]"
'Response.write "<BR>"
		        ReDim Preserve m_iGakusei(m_SrCnt)
		        ReDim Preserve m_iGakuseki(m_SrCnt)
		        ReDim Preserve m_sSimei(m_SrCnt)
		        'ReDim Preserve m_iTensuu(m_rCnt,m_SrCnt)

	            m_iGakusei(m_SrCnt) = m_SRs("T16_GAKUSEI_NO")
	            m_iGakuseki(m_SrCnt) = m_SRs("T16_GAKUSEKI_NO")
	            m_sSimei(m_SrCnt) = m_SRs("T11_SIMEI")

				w_iGakuseiNo = m_SRs("T16_GAKUSEI_NO")
			End if

			w_iDspFlg = 0
			If m_sKBN = C_HISSEN_HIS Then
				'// 必須が選択された場合の抽出
				If m_SRs("T16_HISSEN_KBN") = C_HISSEN_HIS Then
					w_iDspFlg = 1
				End If
			Else
				'// 選択が選択された場合の抽出
				If cint(m_SRs("T16_HISSEN_KBN")) = C_HISSEN_SEN and cint(m_SRs("T16_SELECT_FLG")) = C_SENTAKU_YES then 
					w_iDspFlg = 1
				End If
			End If

			If w_iDspFlg = 1 Then
				'//試験区分毎に点数の項目を求めて配列にセットする
				w_iWkTensuu = 0
				Select Case m_sSikenKBN
					Case C_SIKEN_ZEN_TYU
						w_iWkTensuu = m_SRs("T16_SEI_TYUKAN_Z")
					Case C_SIKEN_ZEN_KIM
						w_iWkTensuu = m_SRs("T16_SEI_KIMATU_Z")
					Case C_SIKEN_KOU_TYU
						w_iWkTensuu = m_SRs("T16_SEI_TYUKAN_K")
					Case C_SIKEN_KOU_KIM
						w_iWkTensuu = m_SRs("T16_SEI_KIMATU_K")
				End Select
				w_iKamokuIdx = f_GetKamokuIdx(m_SRs("T16_KAMOKU_CD"))
				if w_iKamokuIdx > 0 Then
					m_iTensuu(w_iKamokuIdx,m_SrCnt) = w_iWkTensuu
				End If

			End If

            m_SRs.MoveNext
            
        Loop

		f_getGaku = 0
		Exit Do

	Loop


    Call gf_closeObject(w_rs)
    Call gf_closeObject(m_SRs)

End Function

Function f_GetKamokuIdx(p_sKamokuCd)

	f_GetKamokuIdx = 0
	For i=1 to m_rCnt
		If m_iKamokuCd(i) = p_sKamokuCd Then
			f_GetKamokuIdx = i
			Exit For
		End If
	Next

End Function

Function f_TantoKyokan(p_sKamoku)
'********************************************************************************
'*	[機能]	担当教官名称取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Dim w_sTKyokan


	w_sTKyokan = ""

	w_sSQL = ""
	w_sSQL = w_sSQL & "	SELECT "
	w_sSQL = w_sSQL & " 	B.M04_KYOKANMEI_SEI,B.M04_KYOKANMEI_MEI "
	w_sSQL = w_sSQL & "	FROM "
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN A,M04_KYOKAN B "
	w_sSQL = w_sSQL & "	WHERE "
	w_sSQL = w_sSQL & "		A.T27_NENDO = " & m_iNendo & " "
	w_sSQL = w_sSQL & "	AND A.T27_GAKUNEN = " & Cint(m_sGakuNo) & " "
	w_sSQL = w_sSQL & "	AND A.T27_KAMOKU_CD = '" & p_sKamoku & "' "
	w_sSQL = w_sSQL & "	AND A.T27_CLASS = " & Cint(m_sClassNo) & " "
	w_sSQL = w_sSQL & " AND	A.T27_NENDO = B.M04_NENDO(+) "
	w_sSQL = w_sSQL & " AND	A.T27_KYOKAN_CD = B.M04_KYOKAN_CD(+) "

	Set m_TRs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(m_TRs, w_sSQL)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function 
	End If

	If m_TRs.EOF = False Then
		w_sTKyokan = m_TRs("M04_KYOKANMEI_SEI")&"　"&m_TRs("M04_KYOKANMEI_MEI")
	End If

	f_TantoKyokan = w_sTKyokan

    Call gf_closeObject(m_TRs)

    Err.Clear

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_sKBN
Dim i
Dim j

%>
	<html>
	<head>
	<link rel=stylesheet href="../../common/style.css" type=text/css>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT language="JavaScript">
	<!--
	//-->
	</SCRIPT>
	</head>
	<body>
	<form name="frm" method="post">
	<center>
	<table width="100%">
	<tr>
	<td width=100% valign=top>
		<table class=hyo border=1 align=center>
		<tr>
			<th class=header colspan=2 width="180">科　目　名</th>
	<%	For i = 1 to m_rCnt %>
			<td class=detail width="16" align=center valign=top><%=m_sKamokuNm(i)%></td>
	<%	Next%>
		</tr>
		<tr>
			<th class=header colspan=2>科目分類</th>
	<%	For i = 1 to m_rCnt
			If m_sKBN = Cint(C_HISSEN_HIS) Then
				w_sKBN = "必修"
			Else 
				w_sKBN = "選択"
			End If
	%>
			<td class=detail width="16" align=center valign=top><%=w_sKBN%></td>
	<%	Next %>
		</tr>
		<tr>
			<th class=header colspan=2>単　位　数</th>
	<%	For i = 1 to m_rCnt %>
			<td class=detail width="16" align=center valign=top><%=m_iHTani(i)%></td>
	<%	Next%>
		</tr>
		<tr>
			<th class=header colspan=2>担当教官</th>
	<%	For i = 1 to m_rCnt%> 
			<td class=detail width="16" rowspan=2 align=center valign=top><%=f_TantoKyokan(m_iKamokuCd(i))%></td>
	<%	Next%>
		</tr>
		<tr>
			<th class=header2><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
			<th class=header2>氏　名</th>
		</tr>

	<%	For j = 1 to m_SrCnt 
			Call gs_cellPtn(w_cell)%>
		<tr>
			<td class=<%=w_cell%>><%=m_iGakuseki(j)%></td>
			<td class=<%=w_cell%>><%=m_sSimei(j)%></td>
		<%	For i = 1 to m_rCnt%>
				<td class=<%=w_cell%> width="16" align=right><%=m_iTensuu(i,j) %></td>
		<%	Next%>
			</tr>
	<%	Next%>
		</table>
	</td>
	</tr>
	</table>
	</FORM>
	</center>
	</body>
	</html>
<%
End sub

Sub showPage_No()
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
End Sub
%>