<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0150_upd.asp
' 機      能: 下ページ 成績登録の登録、更新
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
' 変      数:
' 引      渡:
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 廣田
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Dim m_bErrFlg		'//ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Dim m_sKyokanCd			'//教官CD
    Dim m_iNendo
    Dim m_sSikenKBN
    Dim m_sKamokuCd
    Dim i_max
    Dim m_sGakuNo			'//学年
    Dim m_sGakkaCd			'//学科
	Dim m_iSeisekiInpType
	Dim m_bZenkiOnly		'前期開設フラグ(True→前期開設、False→前期開設でない)
	Dim m_SchoolFlg
	Dim m_bSeiInpFlg
	Dim m_HyokaDispFlg
	Dim m_KekkaGaiDispFlg

	Dim m_bNiteiFlg

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Sub Main()
	Dim w_iRet              '// 戻り値
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="成績登録"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		if gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		end If

		'データ取得
		Call s_SetParam()

		'//不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'//トランザクション開始
		Call gs_BeginTrans()

		'//認定前後情報取得
'		if not gf_GetNintei(m_iNendo,m_bNiteiFlg) then
		if not gf_GetGakunenNintei(m_iNendo,cint(m_sGakuNo),m_bNiteiFlg) then	'2003.04.11 hirota
			m_bErrFlg = True
			Exit Do
		end if

		'//成績登録
		If f_Update(m_sSikenKBN) <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

'response.end

		'//試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
		'//前期のみの場合は、取得したデータを後期期末試験にも登録する
		If cint(m_sSikenKBN) = cint(C_SIKEN_ZEN_KIM) Then
			If m_bZenkiOnly = True Then
				'//成績登録(前期のみの試験科目の場合)
				If f_Update(C_SIKEN_KOU_KIM) <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If
			End If
		End If

		'// ページを表示
		Call showPage()

		Exit Do
	Loop

    '//ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        '//ロールバック
        Call gs_RollbackTrans()
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    else
		'//コミット
		Call gs_CommitTrans()
    End If

    '// 終了処理
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()

	m_sKyokanCd       = request("txtKyokanCd")
	m_iNendo          = request("txtNendo")
	m_sSikenKBN       = Cint(request("sltShikenKbn"))
	m_sKamokuCd       = request("KamokuCd")
	i_max             = request("i_Max")
	m_sGakuNo	      = Cint(request("txtGakuNo"))			'//学年
	m_sGakkaCd	      = request("txtGakkaCd")				'//学科
	m_iSeisekiInpType = cint(request("hidSeisekiInpType"))
	m_bZenkiOnly      = cbool(request("hidZenkiOnly"))
	m_SchoolFlg       = cbool(request("hidSchoolFlg"))
	m_HyokaDispFlg    = cbool(request("hidHyokaDispFlg"))
	m_KekkaGaiDispFlg = cbool(request("hidKekkaGaiDispFlg"))
	m_bSeiInpFlg      = cbool(request("hidKikan"))

End Sub

'********************************************************************************
'*  [機能]  ヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]
'********************************************************************************
Function f_Update(p_sSikenKBN)
	Dim i
	Dim w_Today
	Dim w_FieldName
	Dim w_DataKbnFlg
	Dim w_DataKbn
	Dim w_Time

	On Error Resume Next
	Err.Clear

	f_Update = 99
	w_DataKbnFlg = false
	w_DataKbn = 0

	Do
		w_Today = gf_YYYY_MM_DD(year(date())  & "/" & month(date()) & "/" & day(date()),"/")
		w_Time  = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)

		'//減算区分取得(sei0150_upd_func.asp内関数)
		'If Not Incf_SelGenzanKbn() Then Exit Function

		'//欠課・欠席設定取得(sei0150_upd_func.asp内関数)
		'If Not Incf_SelM15_KEKKA_KESSEKI() then Exit Function

		'//累積区分取得(sei0150_upd_func.asp内関数)
		'If Not Incf_SelKanriMst(m_iNendo,C_K_KEKKA_RUISEKI) then Exit Function

		For i=1 to i_max
			'//実授業時間取得(sei0150_upd_func.asp内関数)
			'Call Incs_GetJituJyugyou(i)

			'//学期末の場合、最低時間を取得する
			'if Cint(m_sSikenKBN) = C_SIKEN_KOU_KIM then
			'	'//最低時間取得(sei0150_upd_func.asp内関数)
			'	If Not Incf_GetSaiteiJikan(i) then Exit Function
			'End if

			'//評価不能チェック(熊本電波のみ)
			If m_SchoolFlg = true then
				w_DataKbn = 0
				w_DataKbnFlg = false

				'//未評価、評価不能の設定
				if cint(gf_SetNull2Zero(request("hidMihyoka"))) <> 0 then
					w_DataKbn = cint(gf_SetNull2Zero(request("hidMihyoka")))
					w_DataKbnFlg = true
				else
					w_DataKbn = cint(gf_SetNull2Zero(request("chkHyokaFuno" & i)))

					if w_DataKbn = cint(C_HYOKA_FUNO) then
						w_DataKbnFlg = true
					end if
				end if
			End If

			w_sSQL = ""
			w_sSQL = w_sSQL & " UPDATE T16_RISYU_KOJIN SET "

			Select Case p_sSikenKBN

				Case C_SIKEN_ZEN_TYU
					If Not request("hidNoChange" & i ) Then
						'//数値入力
						If m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & "	T16_HTEN_TYUKAN_Z    = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							w_sSQL = w_sSQL & "	T16_HYOKA_TYUKAN_Z   = '', "
							'//認定前か後か
							If not m_bNiteiFlg then
								w_sSQL = w_sSQL & "	T16_SEI_TYUKAN_Z = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							End If
						'//文字入力
						ElseIf m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & "	T16_HYOKA_TYUKAN_Z ='" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & "	T16_SEI_TYUKAN_Z   = NULL, "
							w_sSQL = w_sSQL & "	T16_HTEN_TYUKAN_Z  = NULL, "
						End If
						'欠課・停止・派遣・忌引
						w_sSQL = w_sSQL & " T16_KEKA_TYUKAN_Z		= " & f_CnvNumNull(request("txtKekka"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KEKA_NASI_TYUKAN_Z	= " & f_CnvNumNull(request("txtTeisi"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KOUKETSU_TYUKAN_Z	= " & f_CnvNumNull(request("txtHaken"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KIBI_TYUKAN_Z		= " & f_CnvNumNull(request("txtKibi"&i)) & ", "
					End if

					If m_bSeiInpFlg then
						w_sSQL = w_sSQL & " T16_SOJIKAN_TYUKAN_Z    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
						w_sSQL = w_sSQL & " T16_JUNJIKAN_TYUKAN_Z   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
						w_sSQL = w_sSQL & " T16_J_JUNJIKAN_TYUKAN_Z = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
					End If

					'更新日
					w_sSQL = w_sSQL & " T16_KOUSINBI_TYUKAN_Z   = '" & w_Today & "',"
					w_sSQL = w_sSQL & " T16_KOUSINTIME_TYUKAN_Z = '" & w_Time & "',"

				Case C_SIKEN_ZEN_KIM
					If Not request("hidNoChange" & i ) Then
						'//数値入力
						If m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T16_HTEN_KIMATU_Z  = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							w_sSQL = w_sSQL & " T16_HYOKA_KIMATU_Z = '', "
							'//認定前か後か
							If Not m_bNiteiFlg then
								w_sSQL = w_sSQL & " T16_SEI_KIMATU_Z   = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
								'前期開設科目ならば学年末成績にも登録する
								If m_bZenkiOnly then
									w_sSQL = w_sSQL & "	T16_SEI_KIMATU_K = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
								End If
							End If
							'前期開設科目ならば学年末成績にも登録する
							If m_bZenkiOnly then
								w_sSQL = w_sSQL & "	T16_HTEN_KIMATU_K    = " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							End If

						'//文字入力
						ElseIf m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & " T16_HYOKA_KIMATU_Z  ='" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T16_SEI_KIMATU_Z    = NULL, "
							w_sSQL = w_sSQL & " T16_HTEN_KIMATU_Z   = NULL, "
						End If

						w_sSQL = w_sSQL & " T16_KEKA_KIMATU_Z		= " & f_CnvNumNull(request("txtKekka"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KEKA_NASI_KIMATU_Z	= " & f_CnvNumNull(request("txtTeisi"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KOUKETSU_KIMATU_Z	= " & f_CnvNumNull(request("txtHaken"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KIBI_KIMATU_Z		= " & f_CnvNumNull(request("txtKibi"&i)) & ", "

						If m_bZenkiOnly then
							'前期開設科目ならば学年末欠課にも登録する
							w_sSQL = w_sSQL & " T16_KEKA_KIMATU_K		= " & f_CnvNumNull(request("txtKekka"&i)) & ", "
							w_sSQL = w_sSQL & " T16_KEKA_NASI_KIMATU_K	= " & f_CnvNumNull(request("txtTeisi"&i)) & ", "
							w_sSQL = w_sSQL & " T16_KOUKETSU_KIMATU_K	= " & f_CnvNumNull(request("txtHaken"&i)) & ", "
							w_sSQL = w_sSQL & " T16_KIBI_KIMATU_K		= " & f_CnvNumNull(request("txtKibi"&i)) & ", "
						End If
						'前期開設科目以外ならば（前期期末 + 後期中間）を学年末に登録する INS 2004/02/16
						If Not m_bZenkiOnly then
							'欠課
							If Not gf_IsNull(request("hidKeka_ZK"&i)) or Not gf_IsNull(request("txtKekka"&i)) then

								w_sSQL = w_sSQL & "	T16_KEKA_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidKeka_ZK"&i))) + Cint(gf_SetNull2Zero(request("txtKekka"&i))) & ", "
							Else
								w_sSQL = w_sSQL & "	T16_KEKA_KIMATU_K = NULL, "
							End If
							'停止
							If Not gf_IsNull(request("hidTeisi_ZK"&i)) or Not gf_IsNull(request("txtTeisi"&i)) then
								w_sSQL = w_sSQL & "	T16_KEKA_NASI_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidTeisi_ZK"&i))) + Cint(gf_SetNull2Zero(request("txtTeisi"&i))) & ", "
							Else
								w_sSQL = w_sSQL & "	T16_KEKA_NASI_KIMATU_K = NULL, "
							End If
							'派遣
							If Not gf_IsNull(request("hidHaken_ZK"&i)) or Not gf_IsNull(request("txtHaken"&i)) then
								w_sSQL = w_sSQL & "	T16_KOUKETSU_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidHaken_ZK"&i))) + Cint(gf_SetNull2Zero(request("txtHaken"&i))) & ", "
							Else
								w_sSQL = w_sSQL & "	T16_KOUKETSU_KIMATU_K = NULL, "
							End If
							'忌引
							If Not gf_IsNull(request("hidKibi_ZK"&i)) or Not gf_IsNull(request("txtKibi"&i)) then
								w_sSQL = w_sSQL & "	T16_KIBI_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidKibi_ZK"&i))) + Cint(gf_SetNull2Zero(request("txtKibi"&i))) & ", "
							Else
								w_sSQL = w_sSQL & "	T16_KIBI_KIMATU_K = NULL, "
							End If
						End If
						'前期開設科目以外ならば（前期期末 + 後期期末）を学年末に登録する INS END 2004/02/16




					End If
					If m_bSeiInpFlg then
						w_sSQL = w_sSQL & " T16_SOJIKAN_KIMATU_Z     = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
						w_sSQL = w_sSQL & " T16_JUNJIKAN_KIMATU_Z    = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
						w_sSQL = w_sSQL & " T16_J_JUNJIKAN_KIMATU_Z  = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
						'前期開設科目ならば学年末成績にも登録する
						If m_bZenkiOnly then
							w_sSQL = w_sSQL & " T16_SOJIKAN_KIMATU_K    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
							w_sSQL = w_sSQL & " T16_JUNJIKAN_KIMATU_K   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
							w_sSQL = w_sSQL & " T16_J_JUNJIKAN_KIMATU_K = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
						End If
					End If
					w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_Z   = '" & w_Today & "',"
					w_sSQL = w_sSQL & " T16_KOUSINTIME_KIMATU_Z = '" & w_Time & "',"

					If m_bZenkiOnly then
						w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_K   = '" & w_Today & "',"
						w_sSQL = w_sSQL & " T16_KOUSINTIME_KIMATU_K = '" & w_Time & "',"
					End If

				Case C_SIKEN_KOU_TYU
					If Not request("hidNoChange" & i ) Then
						'//数値入力
						If m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T16_HYOKA_TYUKAN_K = '', "
							w_sSQL = w_sSQL & " T16_HTEN_TYUKAN_K  =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							'//認定前か後か
							If not m_bNiteiFlg then
								w_sSQL = w_sSQL & " T16_SEI_TYUKAN_K   =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							End If
						'//文字入力
						ElseIf m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & " T16_HYOKA_TYUKAN_K = '" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T16_SEI_TYUKAN_K   =  NULL, "
							w_sSQL = w_sSQL & " T16_HTEN_TYUKAN_K  =  NULL, "
						End If
						w_sSQL = w_sSQL & " T16_KEKA_TYUKAN_K		= " & f_CnvNumNull(request("txtKekka"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KEKA_NASI_TYUKAN_K	= " & f_CnvNumNull(request("txtTeisi"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KOUKETSU_TYUKAN_K	= " & f_CnvNumNull(request("txtHaken"&i)) & ", "
						w_sSQL = w_sSQL & " T16_KIBI_TYUKAN_K		= " & f_CnvNumNull(request("txtKibi"&i)) & ", "
					End If
					If m_bSeiInpFlg then
						w_sSQL = w_sSQL & " T16_SOJIKAN_TYUKAN_K    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
						w_sSQL = w_sSQL & " T16_JUNJIKAN_TYUKAN_K   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
						w_sSQL = w_sSQL & " T16_J_JUNJIKAN_TYUKAN_K = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					End If

					'前期開設科目以外ならば（前期期末 + 後期中間）を学年末に登録する INS 2004/02/16
					If Not m_bZenkiOnly then
						'欠課
						If Not gf_IsNull(request("hidKeka_ZK"&i)) or Not gf_IsNull(request("txtKekka"&i)) then

							w_sSQL = w_sSQL & "	T16_KEKA_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidKeka_ZK"&i))) + Cint(gf_SetNull2Zero(request("txtKekka"&i))) & ", "
						Else
							w_sSQL = w_sSQL & "	T16_KEKA_KIMATU_K = NULL, "
						End If
						'停止
						If Not gf_IsNull(request("hidTeisi_ZK"&i)) or Not gf_IsNull(request("txtTeisi"&i)) then
							w_sSQL = w_sSQL & "	T16_KEKA_NASI_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidTeisi_ZK"&i))) + Cint(gf_SetNull2Zero(request("txtTeisi"&i))) & ", "
						Else
							w_sSQL = w_sSQL & "	T16_KEKA_NASI_KIMATU_K = NULL, "
						End If
						'派遣
						If Not gf_IsNull(request("hidHaken_ZK"&i)) or Not gf_IsNull(request("txtHaken"&i)) then
							w_sSQL = w_sSQL & "	T16_KOUKETSU_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidHaken_ZK"&i))) + Cint(gf_SetNull2Zero(request("txtHaken"&i))) & ", "
						Else
							w_sSQL = w_sSQL & "	T16_KOUKETSU_KIMATU_K = NULL, "
						End If
						'忌引
						If Not gf_IsNull(request("hidKibi_ZK"&i)) or Not gf_IsNull(request("txtKibi"&i)) then
							w_sSQL = w_sSQL & "	T16_KIBI_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidKibi_ZK"&i))) + Cint(gf_SetNull2Zero(request("txtKibi"&i))) & ", "
						Else
							w_sSQL = w_sSQL & "	T16_KIBI_KIMATU_K = NULL, "
						End If
					End If
					'前期開設科目以外ならば（前期期末 + 後期期末）を学年末に登録する INS END 2004/02/16


					w_sSQL = w_sSQL & " T16_KOUSINBI_TYUKAN_K   = '" & w_Today & "',"
					w_sSQL = w_sSQL & " T16_KOUSINTIME_TYUKAN_K = '" & w_Time & "',"

				Case C_SIKEN_KOU_KIM
					If Not request("hidNoChange" & i ) Then
						'//数値入力
						If m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
							w_sSQL = w_sSQL & " T16_HYOKA_KIMATU_K = '', "
							w_sSQL = w_sSQL & " T16_HTEN_KIMATU_K  =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							'//認定前か後か
							If Not m_bNiteiFlg then
								w_sSQL = w_sSQL & " T16_SEI_KIMATU_K   =  " & gf_IIF(w_DataKbnFlg,"NULL",f_CnvNumNull(request("Seiseki"&i))) & ", "
							End If
						'//文字入力
						ElseIf m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then
							w_sSQL = w_sSQL & " T16_HYOKA_KIMATU_K = '" & gf_IIF(w_DataKbnFlg,"",request("hidSeiseki"&i)) & "', "
							w_sSQL = w_sSQL & " T16_SEI_KIMATU_K   = NULL, "
							w_sSQL = w_sSQL & " T16_HTEN_KIMATU_K  = NULL, "
							w_sSQL = w_sSQL & " T16_HYOKA_FUKA_KBN = " & gf_SetNull2Zero(request("hidHyokaFukaKbn" & i)) & ", "
						End If
					End if
					If m_bSeiInpFlg then
						w_sSQL = w_sSQL & " T16_SOJIKAN_KIMATU_K    = " & f_CnvNumNull(request("hidSouJyugyou"))  & ","
						w_sSQL = w_sSQL & " T16_JUNJIKAN_KIMATU_K   = " & f_CnvNumNull(request("hidJunJyugyou"))  & ","
						w_sSQL = w_sSQL & " T16_J_JUNJIKAN_KIMATU_K = " & f_CnvNumNull(request("hidJunJyugyou")) & ","
					End If

					'前期開設科目以外ならば（前期期末 + 後期期末）を学年末に登録する
					If Not m_bZenkiOnly then
						'欠課
						If Not gf_IsNull(request("hidKeka_ZK"&i)) or Not gf_IsNull(request("hidKeka_KK"&i)) then
							w_sSQL = w_sSQL & "	T16_KEKA_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidKeka_ZK"&i))) + Cint(gf_SetNull2Zero(request("hidKeka_KK"&i))) & ", "
						Else
							w_sSQL = w_sSQL & "	T16_KEKA_KIMATU_K = NULL, "
						End If
						'停止
						If Not gf_IsNull(request("hidTeisi_ZK"&i)) or Not gf_IsNull(request("hidTeisi_KK"&i)) then
							w_sSQL = w_sSQL & "	T16_KEKA_NASI_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidTeisi_ZK"&i))) + Cint(gf_SetNull2Zero(request("hidTeisi_KK"&i))) & ", "
						Else
							w_sSQL = w_sSQL & "	T16_KEKA_NASI_KIMATU_K = NULL, "
						End If
						'派遣
						If Not gf_IsNull(request("hidHaken_ZK"&i)) or Not gf_IsNull(request("hidHaken_KK"&i)) then
							w_sSQL = w_sSQL & "	T16_KOUKETSU_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidHaken_ZK"&i))) + Cint(gf_SetNull2Zero(request("hidHaken_KK"&i))) & ", "
						Else
							w_sSQL = w_sSQL & "	T16_KOUKETSU_KIMATU_K = NULL, "
						End If
						'忌引
						If Not gf_IsNull(request("hidKibi_ZK"&i)) or Not gf_IsNull(request("hidKibi_KK"&i)) then
							w_sSQL = w_sSQL & "	T16_KIBI_KIMATU_K = " & Cint(gf_SetNull2Zero(request("hidKibi_ZK"&i))) + Cint(gf_SetNull2Zero(request("hidKibi_KK"&i))) & ", "
						Else
							w_sSQL = w_sSQL & "	T16_KIBI_KIMATU_K = NULL, "
						End If
					End If
					w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_K   = '" & w_Today & "',"
					w_sSQL = w_sSQL & " T16_KOUSINTIME_KIMATU_K = '" & w_Time & "',"
			End Select

			w_sSQL = w_sSQL & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
			w_sSQL = w_sSQL & "   T16_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
			w_sSQL = w_sSQL & " WHERE "
			w_sSQL = w_sSQL & "        T16_NENDO      = " & Cint(m_iNendo) & " "
			w_sSQL = w_sSQL & "    AND T16_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
			w_sSQL = w_sSQL & "    AND T16_KAMOKU_CD  = '" & Trim(m_sKamokuCd) & "'  "

'response.write " w_sSQL = " & w_sSQL & "<br>"
'response.end

			if gf_ExecuteSQL(w_sSQL) <> 0 then Exit Do

		Next

		'//正常終了
		f_Update = 0

		Exit Do
	Loop

End Function

'********************************************************************************
'*  [機能]  数値型項目の更新時の設定
'*  [引数]  値
'*  [戻値]  なし
'*  [説明]  数値が入っている場合は[値]、無い場合は"NULL"を返す
'********************************************************************************
Function f_CnvNumNull(p_vAtai)

	If Trim(p_vAtai) = "" Then
		f_CnvNumNull = "NULL"
	Else
		f_CnvNumNull = cInt(p_vAtai)
    End If

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
%>
    <html>
    <head>
    <title>成績登録</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //************************************************************
    function window_onload() {
		document.frm.target = "main";
		document.frm.action = "sei0150_23_print.asp"
		document.frm.submit();
	}
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE="javascript" onload="window_onload();">
    <form name="frm" method="post">

	<input type="hidden" name="txtNendo"     value="<%=trim(Request("txtNendo"))%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=trim(Request("txtKyokanCd"))%>">
	<input type="hidden" name="sltShikenKbn" value="<%=trim(Request("sltShikenKbn"))%>">
	<input type="hidden" name="txtGakuNo"    value="<%=trim(Request("txtGakuNo"))%>">
	<input type="hidden" name="txtClassNo"   value="<%=trim(Request("txtClassNo"))%>">
	<input type="hidden" name="txtKamokuCd"  value="<%=trim(Request("txtKamokuCd"))%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=trim(Request("txtGakkaCd"))%>">
	<input type="hidden" name="hidKamokuKbn" value="<%=request("hidKamokuKbn")%>">
	<input type="hidden" name="hidSyubetu"   value="<%=request("hidSyubetu")%>">
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>