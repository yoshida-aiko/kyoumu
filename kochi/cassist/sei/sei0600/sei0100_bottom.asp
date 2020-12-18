<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0100_bottom.asp
' 機      能: 下ページ 成績登録の検索を行う
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
' 作      成: 2001/07/26 前田 智史
' 変      更: 2001/08/21 伊藤 公子
' 変      更: 2001/08/21 伊藤 公子 ヘッダ部切り離し
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
    Public m_sKamokuCd		'科目コード
    Public m_sSikenNm		'試験名
    Public m_sSikenbi		'試験日
    Public m_sKaisiT		'試験実施開始時間
    Public m_sSyuryoT		'試験実施終了時間
    Public m_sKamokuNo		'科目名
    Public m_sTKyokanCd		'担当科目の教官
	Dim		m_rCnt			'レコードカウント
    Public m_sGakkaCd
    Public m_iSyubetu		'出欠値集計方法
    Public m_TUKU_FLG
    
    Public m_iKamoku_Kbn
    Public m_iHissen_Kbn

	Public	m_Rs
	Public	m_TRs
	Public	m_DRs
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
	w_sMsgTitle="成績登録"
	w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""


'    On Error Resume Next
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

'//デバッグ
'Call s_DebugPrint

		'//期間データの取得
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			'// ページを表示
			Call No_showPage()
			Exit Do
		End If
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If

		'=================
		'//出欠欠課の取り方を取得
		'=================
		'//科目区分(0:試験毎,1:累積)
        w_iRet = gf_GetKanriInfo(m_iNendo,m_iSyubetu)
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		'=================
		'//特別活動を取得
		'=================
		'//特別活動(0:通常授業,1:特別活動)
        w_iRet = f_getTUKU(m_iNendo,m_sKamokuCd,m_sGakuNo,m_sClassNo,m_TUKU_FLG)
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		
    '**********************************************************
    '通常授業と特別活動で、とって来る場所が変わる。
    '**********************************************************
	If m_TUKU_FLG = C_TUKU_FLG_TUJO then  '通常授業の場合
		'=================
		'//科目情報を取得
		'=================
		'//科目区分(0:一般課目,1:専門科目)、及び、必修選択区分(1:必修,2:選択)を調べる
        w_iRet = f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn)
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If


		'===============================
		'//成績、学生データ取得
		'===============================
		'//科目区分がC_KAMOKU_SENMON(0:一般科目)の場合はクラス別に生徒を表示
		'//科目区分がC_KAMOKU_SENMON(1:専門科目)の場合は学科別に生徒を表示
        'w_iRet = f_getdate()
        w_iRet = f_getdate(m_iKamoku_Kbn)
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
		If m_rs.EOF Then
			Call ShowPage_No()
			Exit Do
		End If

		'===============================
		'//欠課数の取得
		'===============================
		w_iRet = f_GetSyukketu()
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

	Else 
		'===============================
		'//成績、学生データ取得
		'===============================
        w_iRet = f_getTUKUclass(m_iNendo,m_sKamokuCd,m_sGakuNo,m_sClassNo)
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
		If m_rs.EOF Then
			Call ShowPage_No()
			Exit Do
		End If


    End If
		'===============================
		'//試験時間等の取得
		'===============================
'		w_iRet = f_GetSikenJikan()
'		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

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
	m_sKamokuCd	= request("txtKamokuCd")
	m_sGakkaCd	= request("txtGakkaCd")

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub
    response.write "m_iNendo	=" & m_iNendo	 & "<br>"
    response.write "m_sKyokanCd	=" & m_sKyokanCd & "<br>"
    response.write "m_sSikenKBN	=" & m_sSikenKBN & "<br>"
    response.write "m_sGakuNo	=" & m_sGakuNo	 & "<br>"
    response.write "m_sClassNo	=" & m_sClassNo	 & "<br>"
    response.write "m_sKamokuCd	=" & m_sKamokuCd & "<br>"
    response.write "m_sGakkaCd	=" & m_sGakkaCd  & "<br>"

End Sub

'********************************************************************************
'*  [機能]  欠課数、遅刻数を取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSyukketu()

    Dim w_iRet
	Dim w_iSyubetu
	Dim w_bZenkiOnly
	Dim w_sSikenKBN
	Dim w_sTKyokanCd

    On Error Resume Next
    Err.Clear

    f_GetSyukketu = 1

	Do
		'==========================================
		'//科目担当教官の教官CDの取得
		'==========================================
        'w_iRet = f_GetTantoKyokan()
        w_iRet = f_GetTantoKyokan(w_sTKyokanCd)
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'==========================================
		'//管理マスタより、出欠欠課の取り方を取得
		'==========================================
		w_iRet = f_GetKanriInfo(w_iSyubetu)
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'==========================================
		'//試験科目が前期のみか通年かを調べる
		'==========================================
		'//前期のみの場合はT21より前記期末試験までの欠課数を取得する
		w_iRet = f_SikenInfo(w_bZenkiOnly)
		If w_iRet<> 0 Then
			Exit Do
		End If 

		If w_bZenkiOnly = True Then
			w_sSikenKBN = C_SIKEN_ZEN_KIM
		Else
			w_sSikenKBN = m_sSikenKBN
		End If

		'==========================================
		'//科目に対する結果,遅刻の値取得
		'==========================================
		'Call gf_GetSyukketuData(m_SRs,m_sSikenKBN,m_sGakuNo,m_sTKyokanCd,m_sClassNo,m_sKamokuCd,w_skaisibi,w_sSyuryobi,"",w_iSyubetu)
		Call gf_GetSyukketuData(m_SRs,w_sSikenKBN,m_sGakuNo,w_sTKyokanCd,m_sClassNo,m_sKamokuCd,w_skaisibi,w_sSyuryobi,"")
		if m_SRs.EOF = false then m_SRs.MoveFirst
		'//正常終了
	    f_GetSyukketu = 0
		Exit Do
	Loop

End Function 

'********************************************************************************
'*  [機能]  管理マスタより出欠欠課の取り方を取得
'*  [引数]  なし
'*  [戻値]  p_sSyubetu = C_K_KEKKA_RUISEKI_SIKEN : 試験毎(=0)
'*  [戻値]  p_sSyubetu = C_K_KEKKA_RUISEKI_KEI   ：累積(=1)
'*  [説明]  
'********************************************************************************
Function f_GetKanriInfo(p_iSyubetu)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKanriInfo = 1

    Do 

		'//管理マスタより欠課累積情報区分を取得
		'//欠課累積情報区分(C_K_KEKKA_RUISEKI = 32)
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_SYUBETU"
		w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M00_KANRI.M00_NO=" & C_K_KEKKA_RUISEKI	'欠課累積情報区分(=32)

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKanriInfo = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			'//Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '試験毎
			'//Public Const C_K_KEKKA_RUISEKI_KEI = 1      '累積
			p_iSyubetu = w_Rs("M00_SYUBETU")

		End If

        f_GetKanriInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_SikenInfo(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_SikenInfo = 1
	p_bZenkiOnly = false

    Do 

		'//試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSQL = w_sSQL & vbCrLf & " T15_RISYU.T15_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & Cint(m_iNendo)-cint(m_sGakuNo)+1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & Trim(m_sKamokuCd) & "'" 
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAISETU" & m_sGakuNo & "=" & C_KAI_ZENKI

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_SikenInfo = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_bZenkiOnly = True
		End If

        f_SikenInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  コンボで選択された科目の科目区分及び、必修選択区分を調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKamokuInfo(p_iKamoku_Kbn,p_iHissen_Kbn)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKamokuInfo = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKU_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  ,T15_RISYU.T15_HISSEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(m_sGakuNo) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & m_sKamokuCd & "' "

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKamokuInfo = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_iKamoku_Kbn = w_Rs("T15_KAMOKU_KBN")
			p_iHissen_Kbn = w_Rs("T15_HISSEN_KBN")
		End If

        f_GetKamokuInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'Function f_getdate()
Function f_getdate(p_iKamoku_Kbn)
'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Dim w_iNyuNendo


	On Error Resume Next
	Err.Clear
	f_getdate = 1

	Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakuNo) + 1

		'//検索結果の値より一覧を表示
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "

		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_Z AS SEI,A.T16_KEKA_TYUKAN_Z AS KEKA,A.T16_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_Z AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI, "
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI,A.T16_KEKA_KIMATU_Z AS KEKA,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T16_CHIKAI_KIMATU_Z AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI,A.T16_KEKA_TYUKAN_K AS KEKA,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_K AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
		End Select

		w_sSQL = w_sSQL & " 	A.T16_GAKUSEI_NO AS GAKUSEI_NO,A.T16_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_SELECT_FLG"
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_OKIKAE_FLG"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T16_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(m_sGakuNo) & " "

		'//科目区分がC_KAMOKU_SENMON(1:専門科目)の場合は学科別に生徒を表示
		If cint(p_iKamoku_Kbn) = cint(C_KAMOKU_SENMON) Then
			w_sSQL = w_sSQL & vbCrLf & " AND	C.T13_GAKKA_CD = '" & m_sGakkaCd & "' "
		Else
			w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
		End If

		w_sSQL = w_sSQL & " AND	A.T16_NENDO = C.T13_NENDO "

		'//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
		w_sSQL = w_sSQL & " AND	A.T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
'		w_sSQL = w_sSQL & " AND	B.T11_NYUNENDO = " & w_iNyuNendo & " "
		w_sSQL = w_sSQL & " ORDER BY A.T16_GAKUSEKI_NO "

'response.write w_sSQL &"<<br>"

		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getdate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		'//ﾚｺｰﾄﾞカウント取得
		m_rCnt=gf_GetRsCount(m_Rs)

		f_getdate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[機能]	特別活動受講学生取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Function f_getTUKUclass(p_iNendo,p_sKamokuCd,p_iGakunen,p_iClass)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet
    Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear
    
    f_getTUKUclass = 1
	p_sTKyokanCd = ""

	Do

        w_iNyuNendo = Cint(p_iNendo) - Cint(p_iGakunen) + 1

		'//検索結果の値より一覧を表示
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "

		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_Z AS KEKA,A.T34_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,A.T34_CHIKAI_TYUKAN_Z AS CHIKAI, "
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_Z AS KEKA,A.T34_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T34_CHIKAI_KIMATU_Z AS CHIKAI, "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_K AS KEKA,A.T34_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T34_CHIKAI_TYUKAN_K AS CHIKAI, "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_K AS KEKA,A.T34_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T34_CHIKAI_KIMATU_K AS CHIKAI, "
		End Select

		w_sSQL = w_sSQL & " 	A.T34_GAKUSEI_NO AS GAKUSEI_NO,A.T34_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T34_RISYU_TOKU A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T34_NENDO = " & Cint(p_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T34_TOKUKATU_CD = '" & p_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T34_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T34_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(p_iGakunen) & " "
		w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(p_iClass) & " "

		w_sSQL = w_sSQL & " AND	A.T34_NENDO = C.T13_NENDO "

		w_sSQL = w_sSQL & " ORDER BY A.T34_GAKUSEKI_NO "

'response.write w_sSQL &"<<br>"

		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getTUKUclass = 99
			m_bErrFlg = True
			Exit Do 
		End If

		'//ﾚｺｰﾄﾞカウント取得
		m_rCnt=gf_GetRsCount(m_Rs)

		f_getTUKUclass = 0
		Exit Do
	Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[機能]	科目担当教官の教官CDの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
'Function f_GetTantoKyokan()
Function f_GetTantoKyokan(p_sTKyokanCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetTantoKyokan = 1
	p_sTKyokanCd = ""

    Do 
		'//科目担当教官の教官CDの取得
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & "  T20_KYOKAN "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & "  T20_JIKANWARI "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & "  T20_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND T20_KAMOKU = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND T20_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND T20_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " GROUP BY T20_KYOKAN "

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetTantoKyokan = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			'm_sTKyokanCd = w_Rs("T20_KYOKAN")
			p_sTKyokanCd = w_Rs("T20_KYOKAN")
		End If

        f_GetTantoKyokan = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

Function f_Nyuryokudate()
'********************************************************************************
'*	[機能]	成績入力期間データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakuNo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
'//成績入力期間テスト用

'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '2003/04/30'"
'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '1999/03/01'"

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

Function f_getTUKU(p_iNendo,p_sKamoku,p_iGakunen,p_iClass,p_TUKU_FLG)
'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

	On Error Resume Next
	Err.Clear
	f_getTUKU = 0
	p_TUKU_FLG = "0"

	Do

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

'response.write w_sSQL & "<<<BR>"

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_getTUKU = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If w_Rs.EOF = false Then
			p_TUKU_FLG = w_Rs("T20_TUKU_FLG")
		End If

		Exit Do
	Loop
	
    Call gf_closeObject(w_Rs)

End Function

Function f_Syukketu(p_gaku,p_kbn)
'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	f_Syukketu = 0

	If m_SRs.EOF Then
		Exit Function
	Else
		m_SRs.MoveFirst
		Do Until m_SRs.EOF
			If m_SRs("T21_GAKUSEKI_NO") = p_gaku AND cstr(m_SRs("T21_SYUKKETU_KBN")) = cstr(p_kbn) Then
				f_Syukketu = m_SRs("KAISU")
				Exit Do
			End If
		m_SRs.MoveNext
		Loop
	End If

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_sGakusekiCd
Dim w_sSeiseki
Dim w_sHyoka
Dim w_sKekka,w_sKekkaGai
Dim w_sChikai
Dim w_sKekkasu
Dim w_sChikaisu
Dim w_sShikenKBN_RUI
Dim w_iKekka_rui,w_iChikoku_rui

Dim w_ihalf
Dim i

i = 0

%>
<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
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

        //submit
        document.frm.target = "topFrame";
        document.frm.action = "sei0100_middle.asp"
        document.frm.submit();
        return;

    }

   //************************************************************
    //  [機能]  評価ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_change(p_iS){

        w_sButton = eval("document.frm.button"+p_iS);
        w_sHyouka = eval("document.frm.Hyoka"+p_iS);

<%If m_sSikenKBN = C_SIKEN_ZEN_TYU Then%>

        if(w_sButton.value == "・") {
			w_sButton.value = "○";
			w_sHyouka.value = "○";
	        return;
		}
        if(w_sButton.value == "○") {
			w_sButton.value = "・";
			w_sHyouka.value = "";
	        return;
		}

<%Else%>

        if(w_sButton.value == "・") {
			w_sButton.value = "○";
			w_sHyouka.value = "○";
	        return;
		}
        if(w_sButton.value == "○") {
			w_sButton.value = "◎";
			w_sHyouka.value = "◎";
	        return;
		}
        if(w_sButton.value == "◎") {
			w_sButton.value = "・";
			w_sHyouka.value = "";
	        return;
		}
<%End If%>

    }
    
   //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){

		if(f_CheckData_All() == 1){
            alert("入力値が不正です");
            return 1;
        }else{

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		//ヘッダ部空白表示
		parent.topFrame.document.location.href="white.asp"

		//登録処理
<% if m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
        document.frm.action="sei0100_upd.asp";
<% Else %>
        document.frm.action="sei0100_upd_toku.asp";
<% End if %>
        document.frm.target="main";
        document.frm.submit();
    	}
    }

	//************************************************************
	//	[機能]	キャンセルボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//************************************************************
	function f_Cansel(){

        //初期ページを表示
        parent.document.location.href="default.asp"
	
	}
    //************************************************************
    //  [機能]  入力値のﾁｪｯｸ(登録ボタン押下時)
    //  [引数]  なし
    //  [戻値]  0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
    //  [説明]  入力値のNULLﾁｪｯｸ、英数字ﾁｪｯｸ、桁数ﾁｪｯｸを行う
    //          引渡ﾃﾞｰﾀ用にﾃﾞｰﾀを加工する必要がある場合には加工を行う
    //************************************************************
    function f_CheckData_All() {

		var i
		var w_Seiseki
		var w_bFLG
<% if m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_Seiseki = eval("document.frm.Seiseki"+i);
			w_bFLG = true
			if (isNaN(w_Seiseki.value)){
				w_bFLG = false;
				return 1;
				break;
			}else{

				//マイナスをチェック
				var wStr = new String(w_Seiseki.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					return 1;
					break;
				};

				//小数点チェック
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					return 1;
					break;
				}

			};
		};
			if (w_bFLG == false){
				return 1;
			};
<% End if %>
		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_Kekka = eval("document.frm.Kekka"+i);
			w_bFLG = true
			if (isNaN(w_Kekka.value)){
				w_bFLG = false;
				return 1;
				break;
			}else{

				//マイナスをチェック
				var wStr = new String(w_Kekka.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					return 1;
					break;
				};

				//小数点チェック
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					return 1;
					break;
				}

			};
		};
			if (w_bFLG == false){
				return 1;
			};

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_KekkaGai = eval("document.frm.KekkaGai"+i);
			w_bFLG = true
			if (isNaN(w_KekkaGai.value)){
				w_bFLG = false;
				return 1;
				break;
			}else{

				//マイナスをチェック
				var wStr = new String(w_KekkaGai.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					return 1;
					break;
				};

				//小数点チェック
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					return 1;
					break;
				}

			};
		};
			if (w_bFLG == false){
				return 1;
			};

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_Chikai = eval("document.frm.Chikai"+i);
			w_bFLG = true
			if (isNaN(w_Chikai.value)){
				w_bFLG = false;
				return 1;
				break;
			}else{

				//マイナスをチェック
				var wStr = new String(w_Chikai.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					return 1;
					break;
				};

				//小数点チェック
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					return 1;
					break;
				}

			};
		};

			if (w_bFLG == false){
				return 1;
			};
		return 0;
	};

//************************************************
//Enter キーで下の入力フォームに動くようになる
//引数：p_inpNm	対象入力フォーム名
//    ：p_frm	対象フォーム
//　　：i		現在の番号
//戻値：なし
//入力フォーム名が、xxxx1,xxxx2,xxxx3,…,xxxxn 
//の名前のときに利用できます。
//************************************************
function f_MoveCur(p_inpNm,p_frm,i){
	if (event.keyCode == 13){		//押されたキーがEnter(13)の時に動く。
		i++;
		if (i > <%=m_rCnt%>) i = 1; //iが最大値を超えると、はじめに戻る。
		inpForm = eval("p_frm."+p_inpNm+i);
		inpForm.focus();			//フォーカスを移す。
		inpForm.select();			//移ったテキストボックス内を選択状態にする。
	}else{
//		alert(event.keyCode);
		return false;
	}
	return true;
}


	//-->
	</SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post" onClick="return false;">
	<center>

<!--
	<table border=1>
	<tr>
	<td valign="top">
-->
		<table class="hyo" border="1" align="center" width="550">
	<%	m_Rs.MoveFirst
		Do Until m_Rs.EOF
			w_ihalf = gf_Round(m_rCnt / 2,0)
			'i = i + 1 
			j = j + 1 
			w_sSeiseki = ""
			w_sHyoka = ""
			w_sKekka = ""
			w_sChikai = ""
			w_sGakusekiCd = ""
			w_sKekkasu = ""
			w_sChikaisu = ""
				Call gs_cellPtn(w_cell)

				'If w_ihalf + 1 = i then
'				If w_ihalf + 1 = j then
'				w_cell = ""
'				Call gs_cellPtn(w_cell)%>
<!--		</table>
	</td>
	<td valign="top" width="50%">
		<table class="hyo" border="1" align="center" width="98%">
-->
	<%
	'	 		End If 

		'//各データを取得する
		'** 一つ前の試験区分
		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU								'//前期中間
				w_sShikenKBN_RUI = 99
				
			Case C_SIKEN_ZEN_KIM								'//前期期末
				w_sShikenKBN_RUI = C_SIKEN_ZEN_TYU
				
			Case C_SIKEN_KOU_TYU								'//後期中間
				w_sShikenKBN_RUI = C_SIKEN_ZEN_KIM
				
			Case C_SIKEN_KOU_KIM								'//後期期末
				w_sShikenKBN_RUI = C_SIKEN_KOU_TYU
		End Select

		w_sGakusekiCd = m_Rs("GAKUSEKI_NO")
		w_sKekka = gf_SetNull2Zero(m_Rs("KEKA"))
		w_sKekkaGai = gf_SetNull2Zero(m_Rs("KEKA_NASI"))
		w_sChikai = gf_SetNull2Zero(m_Rs("CHIKAI"))
	'値の初期化。
	w_bNoChange = False
	w_sKekkasu = 0
	w_sChikaisu = 0
	'---------------------------------------------------------------------------------------------
	'通常授業ときの処理
	if m_TUKU_FLG = C_TUKU_FLG_TUJO then 
		w_sSeiseki = m_Rs("SEI")
		w_sHyoka = gf_HTMLTableSTR(m_Rs("HYOKAYOTEI"))
if w_sHyoka = "　" then w_sHyoka = "・"
			'//科目が選択科目の場合は、生徒が選択しているかどうかを判別する。選択しいない生徒は入力不可とする。
			w_bNoChange = False

			If cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then 
				If cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then
					w_bNoChange = True
				End If 
			End If
		'欠課遅刻数の取得　現在通常授業のみ
		w_sKekkasu = cint(f_Syukketu(w_sGakusekiCd,C_KETU_KEKKA))			'//欠課数の取得

		w_sChikaisu = cint(f_Syukketu(w_sGakusekiCd,C_KETU_TIKOKU))		'//遅刻数の取得
		w_sChikaisu = w_sChikaisu + cint(f_Syukketu(w_sGakusekiCd,C_KETU_SOTAI))		'//早退数の取得

	end if
	'---------------------------------------------------------------------------------------------

		'「出欠欠課が累積」で「前期中間でない」の場合
		if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and w_sShikenKBN_RUI <> 99 then 
	 		call gf_GetKekaChi(m_iNendo,w_sShikenKBN_RUI,m_sKamokuCd,cstr(m_Rs("GAKUSEI_NO")),w_iKekka_rui,w_iChikoku_rui,w_iKekkaGai_rui) '一つ前の試験の合計値を足す。
			w_sKekkasu = cint(w_sKekkasu) + cint(w_iKekka_rui)
			w_sChikaisu = cint(w_sChikaisu) + cint(w_iChikoku_rui)
		end if
		
		If cint(w_sKekka) = 0 and cint(w_sKekkasu) > 0 Then 		'//欠入が0で,欠計が0より大きい場合
			w_sKekka = cint(w_sKekkasu)								'//欠入＝欠計
		End If
		If cint(w_sChikai) = 0 AND cint(w_sChikaisu) > 0 Then		'//遅入が0で,遅計が0より大きい場合
			w_sChikai = cint(w_sChikaisu)							'//遅入＝遅計
		End If
	%>
		<tr>
	<%

			'========================================================================================
			'//科目が選択科目の時に科目を選択していない場合(入力不可)
			'========================================================================================
			If w_bNoChange = True Then%>

						<td class="<%=w_cell%>" width="40" ><%=w_sGakusekiCd%></td>
						<td class="<%=w_cell%>" align="left"   width="260"><%=m_Rs("SIMEI")%></td>
						<td class="<%=w_cell%>" align="center" width="35" >-</td>
						<td class="<%=w_cell%>" align="center" width="35" >-</td>
						<td class="<%=w_cell%>" align="center" width="40" >-</td>
						<td class="<%=w_cell%>" align="center" width="40" >-</td>
						<td class="<%=w_cell%>" align="center" width="40" >-</td>
						<td class="<%=w_cell%>" align="center" width="35" >-</td>
						<td class="<%=w_cell%>" align="center" width="35" >-</td>
			<%
			'=========================================================================
			'//科目が必修か、または選択科目の時に生徒が科目を選択している場合(入力可)
			'=========================================================================
			Else
				i = i+1
				%>
						<td class="<%=w_cell%>"  width="40"><%=w_sGakusekiCd%>
						<input type="hidden" name=txtGseiNo<%=i%> value="<%=m_Rs("GAKUSEI_NO")%>"></td>
						<td class="<%=w_cell%>" align="left"  width="210"><%=m_Rs("SIMEI")%></td>

						<%
						'//NN対応
						If session("browser") = "IE" Then
							w_sInputClass = "class='num'"
						Else
							w_sInputClass = ""
						End If
				'=========================================================================
				'//通常授業の場合
				'=========================================================================
						%>
				<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
						
							<td class="<%=w_cell%>" width="30"><input type="text" <%=w_sInputClass%>  name=Seiseki<%=i%> value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)"></td>
					<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>
							<td class="<%=w_cell%>"  width="30"><input type="button" size="2" name="button<%=i%>" value="<%=w_sHyoka%>" onClick="return f_change(<%=i%>)" style="text-align:center" class="<%=w_cell%>"><!-- class="<%=w_cell%>"-->
							<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>"></td>
					<%Else%>
							<td class="<%=w_cell%>"  width="30"><%=w_sHyoka%><input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>"></td>
					<%End If%>
						<td class="<%=w_cell%>" width="20"><input type="text" <%=w_sInputClass%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="20"><input type="text" <%=w_sInputClass%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="30"align="right"  ><%=w_sKekkasu%></td>
						<td class="<%=w_cell%>" width="20"><input type="text" <%=w_sInputClass%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=1 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="25"align="right"  ><%=w_sChikaisu%></td>
				<%Else%>
						<td class="<%=w_cell%>" align="center" width="30" >-</td>
						<td class="<%=w_cell%>" align="center" width="30" >-</td>
						<td class="<%=w_cell%>" width="45" align="center"><input type="text" <%=w_sInputClass%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="45" align="center"><input type="text" <%=w_sInputClass%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="60" align="center"><input type="text" <%=w_sInputClass%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
			
				<%End If%>
			<%End If%>
					</tr>
			<%
			m_Rs.MoveNext
			Loop%>
		</table>
<!--
	</td>
	</tr>
	</table>
-->
	<table width="50%">
	<tr>
		<td align="center"><input type="button" class="button" value="　登　録　" onclick="javascript:f_Touroku()">　
		<input type="button" class="button" value="キャンセル" onclick="javascript:f_Cansel()"></td>
	</tr>
	</table>

		<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
		<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
		<input type="hidden" name="KamokuCd"    value="<%=m_sKamokuCd%>">
		<input type="hidden" name="i_Max"       value="<%=i%>">
		<input type="hidden" name="txtSikenKBN" value="<%=m_sSikenKBN%>">
		<input type="hidden" name="txtGakuNo"   value="<%=m_sGakuNo%>">
		<input type="hidden" name="txtGakkaCd"  value="<%=m_sGakkaCd%>">
		<input type="hidden" name="txtClassNo"  value="<%=m_sClassNo%>">
		<input type="hidden" name="txtKamokuCd" value="<%=m_sKamokuCd%>">
		<input type="hidden" name="txtTUKU_FLG" value="<%=m_TUKU_FLG%>">

	</FORM>
	</center>
	</body>
	</html>
<%
End sub

Sub No_showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT language="javascript">
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

    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	<center>
	<br><br><br>
		<span class="msg">成績入力期間外です。</span>
	</center>

	<input type="hidden" name="txtMsg" value="成績入力期間外です。">

	</form>
	</body>
	</html>

<%
End Sub
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
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT language="javascript">
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

    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	</head>

	<body>
	<br><br><br>
	<center>
		<span class="msg">データが存在しません。</span>
	</center>

	<input type="hidden" name="txtMsg" value="データが存在しません。">

	</form>
	</body>
	</html>

<%
End Sub
%>