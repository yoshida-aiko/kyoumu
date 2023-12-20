<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0100_upd.asp
' 機      能: 下ページ 成績登録の登録、更新
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
' 変      数:
' 引      渡:
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2001/07/27 前田 智史
' 変      更: 2018/03/22 西村 開設時期を前画面から取得し前期開設の場合は後期期末にも成績を更新する
' 変      更: 2023/12/14 吉田          WEBアクセスログカスタマイズ
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<!--#include file="sei0100_upd_func.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
    Const DebugPrint = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	
    '取得したデータを持つ変数
    Dim     m_sKyokanCd     '//教官CD
    Dim     m_iNendo 
    Dim     m_sSikenKBN
    Dim     m_sKamokuCd
    Dim     i_max 
    Dim     m_sGakuNo	'//学年
    Dim     m_sGakkaCd	'//学科
    Dim     m_SchoolFlg
    Dim     m_SQL
	Dim     hidHyoka	'//評価

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
    w_sMsgTitle="成績登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sKyokanCd     = request("txtKyokanCd")
    m_iNendo        = request("txtNendo")
	m_sSikenKBN     = Cint(request("txtSikenKBN"))
	m_sKamokuCd     = request("KamokuCd")
	i_max           = request("i_Max")
	m_sGakuNo	    = Cint(request("txtGakuNo"))	'//学年
	m_sGakkaCd	    = request("txtGakkaCd")			'//学科

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
		'// 成績登録
        w_iRet = f_Update(m_sSikenKBN)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If
	 
'// Del_s 2018/03/22 Nishimura
'//		'//試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
'//		'//前期のみの場合は、取得したデータを後期期末試験にも登録する
'//		If cint(m_sSikenKBN) = cint(C_SIKEN_ZEN_KIM) Then    'C_SIKEN_ZEN_KIM :前期期末試験(=2)
'//
'//			'//試験科目が前期のみか通年かを調べる
'//			w_iRet = f_SikenInfo(w_bZenkiOnly)
'//			If w_iRet<> 0 Then
'//				Exit Do
'//			End If 
'//
'//			If w_bZenkiOnly = True Then
'//		        '// 成績登録(前期のみの試験科目の場合)
'//		        w_iRet = f_Update(C_SIKEN_KOU_KIM)
'//		        If w_iRet <> 0 Then
'//		            m_bErrFlg = True
'//		            Exit Do
'//		        End If
'//
'//			End If
'//
'//		End If
'// Del_e 2018/03/22 Nishimura


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

Function f_Update(p_sSikenKBN)
'********************************************************************************
'*  [機能]  ヘッダ情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Dim i
Dim w_Today
Dim w_DataKbnFlg
Dim w_DataKbn
Dim w_HyokaArray
Dim w_iSeisekiInp
	
    On Error Resume Next
    Err.Clear
	
    f_Update = 99
	w_DataKbnFlg = false
	w_DataKbn = 0
	w_HyokaArray = split(Trim(request("hidHyoka")),",")
	'response.write  "w_HyokaArray:" & w_HyokaArray(0)
	w_iSeisekiInp = request("hidSeisekiInp")

    Do 
		w_Today = gf_YYYY_MM_DD(m_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		
		m_SchoolFlg = cbool(request("hidSchoolFlg"))
		
		'// 減算区分取得(sei0100_upd_func.asp内関数)
		If Not Incf_SelGenzanKbn() Then Exit Function
		
		'// 欠課・欠席設定取得(sei0100_upd_func.asp内関数)
		If Not Incf_SelM15_KEKKA_KESSEKI() then Exit Function
		
		'// 累積区分取得(sei0100_upd_func.asp内関数)
		If Not Incf_SelKanriMst(m_iNendo,C_K_KEKKA_RUISEKI) then Exit Function
		
		For i=1 to i_max
			'//実授業時間取得(sei0100_upd_func.asp内関数)
			Call Incs_GetJituJyugyou(i)
			
			'//学期末の場合、最低時間を取得する
			if cInt(m_sSikenKBN) = C_SIKEN_KOU_KIM then
				'//最低時間取得(sei0100_upd_func.asp内関数)
				If Not Incf_GetSaiteiJikan(i) then Exit Function
			End if
			
			if m_SchoolFlg = true then
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
			end if
			
			'//T16_RISYU_KOJINにUPDATE
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " UPDATE T16_RISYU_KOJIN SET "
			
			Select Case p_sSikenKBN
				
				Case C_SIKEN_ZEN_TYU
					if request("hidUpdFlg" & i ) Then
						
						if w_DataKbnFlg then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_TYUKAN_Z		= NULL , "
						else
							w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_TYUKAN_Z		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
						end if
						'2022/10/14 UPD 成績入力方法で処理を分岐　-->
						if w_iSeisekiInp = C_SEISEKI_INP_TYPE_NUM Then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_TYUKAN_Z	= '" & request("Hyoka"&i) & "', "
						else
							if  w_HyokaArray(i-1) = "@@@" then
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_TYUKAN_Z		= NULL, "
							else
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_TYUKAN_Z		= '" & w_HyokaArray(i-1) & "', "
							end if
						end if
						'2022/10/14 UPD　<--
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_TYUKAN_Z		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_TYUKAN_Z	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_TYUKAN_Z		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
						
					End if
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & vbCrLf & " 	T16_DATAKBN_TYUKAN_Z = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_TYUKAN_Z    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_TYUKAN_Z   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_TYUKAN_Z = " & f_CnvNumNull(m_iJituJyugyou) & ","
					
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_TYUKAN_Z = '" & gf_YYYY_MM_DD(date(),"/") & "',"
					
				Case C_SIKEN_ZEN_KIM
					if request("hidUpdFlg" & i ) Then
						
						if w_DataKbnFlg then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_KIMATU_Z		= NULL, "
						else
							w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_KIMATU_Z		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
						end if
						'response.write "w_HyokaArray(i-1):" &  w_HyokaArray(i-1)
						'2022/10/14 UPD 成績入力方法で処理を分岐　-->
						if w_iSeisekiInp = C_SEISEKI_INP_TYPE_NUM Then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_KIMATU_Z	= '" & request("Hyoka"&i) & "', "
						else	
							if  w_HyokaArray(i-1) = "@@@" then
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_KIMATU_Z		= NULL, "
							else
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_KIMATU_Z		= '" & w_HyokaArray(i-1) & "', "
							end if
						end if
						'2022/10/14 UPD　<--
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_KIMATU_Z		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_KIMATU_Z	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_KIMATU_Z		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
						
					End if
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & vbCrLf & " 	T16_DATAKBN_KIMATU_Z = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_KIMATU_Z    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_KIMATU_Z   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_KIMATU_Z = " & f_CnvNumNull(m_iJituJyugyou) & ","
					
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_KIMATU_Z = '" & gf_YYYY_MM_DD(date(),"/") & "',"
					
					'// Ins_S 2018/03/22 Nishimura
					'//開設時期取得 前期開設の場合は後期期末にも更新
					w_bZenki = Trim(request("txtKaisetu"&i))
					if CInt(w_bZenki) = C_KAI_ZENKI Then
						'//後期期末
						if request("hidUpdFlg" & i ) Then
							
							if w_DataKbnFlg then
								w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_KIMATU_K		= NULL, "
							else
								w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_KIMATU_K		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
							end if
							'2022/10/14 UPD 成績入力方法で処理を分岐　-->
							if w_iSeisekiInp = C_SEISEKI_INP_TYPE_NUM Then
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_KIMATU_K	= '" & request("Hyoka"&i) & "', "
							else
								if  w_HyokaArray(i-1) = "@@@" then
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_KIMATU_K		= NULL, "
								else
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_KIMATU_K		= '" & w_HyokaArray(i-1) & "', "
								end if
							end if
							'2022/10/14 UPD　<--
							w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_KIMATU_K		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
							w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_KIMATU_K	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
							w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_KIMATU_K		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
							
						End if
						
						if m_SchoolFlg = true then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_DATAKBN_KIMATU_K = " & gf_SetNull2Zero(w_DataKbn) & ","
						end if
						
						w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_KIMATU_K    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
						w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_KIMATU_K   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
						w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_KIMATU_K = " & f_CnvNumNull(m_iJituJyugyou) & ","
						w_sSQL = w_sSQL & vbCrLf & " 	T16_SAITEI_JIKAN        = " & f_CnvNumNull(m_iSaiteiJikan) & ","
						
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_KIMATU_K = '" & gf_YYYY_MM_DD(date(),"/") & "',"
						
						if Not gf_IsNull(m_iKyuSaiteiJikan) Then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_KYUSAITEI_JIKAN = " & f_CnvNumNull(m_iKyuSaiteiJikan) & ","
						End if
					End IF
					'// Ins_E 2018/03/22 Nishimura



				Case C_SIKEN_KOU_TYU
					if request("hidUpdFlg" & i ) Then
						
						if w_DataKbnFlg then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_TYUKAN_K		= NULL, "
						else
							w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_TYUKAN_K		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
						end if
						'2022/10/14 UPD 成績入力方法で処理を分岐　-->
						if w_iSeisekiInp = C_SEISEKI_INP_TYPE_NUM Then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_TYUKAN_K	= '" & request("Hyoka"&i) & "', "
						else
							if  w_HyokaArray(i-1) = "@@@" then
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_TYUKAN_K		= NULL, "
							else
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_TYUKAN_K		= '" & w_HyokaArray(i-1) & "', "
							end if
						end if
						'2022/10/14 UPD <--
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_TYUKAN_K		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_TYUKAN_K	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_TYUKAN_K		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
						
					End if
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & vbCrLf & " 	T16_DATAKBN_TYUKAN_K = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_TYUKAN_K    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_TYUKAN_K   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_TYUKAN_K = " & f_CnvNumNull(m_iJituJyugyou) & ","
					
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_TYUKAN_K = '" & gf_YYYY_MM_DD(date(),"/") & "',"
					
				Case C_SIKEN_KOU_KIM
					if request("hidUpdFlg" & i ) Then
						
						if w_DataKbnFlg then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_KIMATU_K		= NULL, "
						else
							w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_KIMATU_K		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
						end if
						'2022/10/14 UPD 成績入力方法で処理を分岐　-->
						if w_iSeisekiInp = C_SEISEKI_INP_TYPE_NUM Then
							w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_KIMATU_K	= '" & request("Hyoka"&i) & "', "
						else
							if  w_HyokaArray(i-1) = "@@@" then
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_KIMATU_K		= NULL, "
							else
								w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKA_KIMATU_K		= '" & w_HyokaArray(i-1) & "', "
							end if
						end if
						'2022/10/14 UPD <--
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_KIMATU_K		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_KIMATU_K	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
						w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_KIMATU_K		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
						
					End if
					
					if m_SchoolFlg = true then
						w_sSQL = w_sSQL & vbCrLf & " 	T16_DATAKBN_KIMATU_K = " & gf_SetNull2Zero(w_DataKbn) & ","
					end if
					
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_KIMATU_K    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_KIMATU_K   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_KIMATU_K = " & f_CnvNumNull(m_iJituJyugyou) & ","
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SAITEI_JIKAN        = " & f_CnvNumNull(m_iSaiteiJikan) & ","
					
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_KIMATU_K = '" & gf_YYYY_MM_DD(date(),"/") & "',"
					
					if Not gf_IsNull(m_iKyuSaiteiJikan) Then
						w_sSQL = w_sSQL & vbCrLf & " 	T16_KYUSAITEI_JIKAN = " & f_CnvNumNull(m_iKyuSaiteiJikan) & ","
					End if
					
			End Select
			
            w_sSQL = w_sSQL & vbCrLf & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
            w_sSQL = w_sSQL & vbCrLf & "   T16_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T16_NENDO = " & Cint(m_iNendo) & " "
            w_sSQL = w_sSQL & vbCrLf & "    AND T16_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "    AND T16_KAMOKU_CD = '" & Trim(m_sKamokuCd) & "'  "

			 'response.write  "txtGseiNo:" & Trim(request("txtGseiNo"&i)) & "<BR>"
            '  response.write w_sSQL & "<BR>"
			'  response.end
            If gf_ExecuteSQL(w_sSQL) <> 0 Then
                '//ﾛｰﾙﾊﾞｯｸ
                msMsg = Err.description
                Exit Do
            End If
             
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

'		'//試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSQL = w_sSQL & vbCrLf & " T15_RISYU.T15_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & Cint(m_iNendo)-cint(m_sGakuNo)+1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & Trim(m_sKamokuCd) & "'" 
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAISETU" & m_sGakuNo & "=" & C_KAI_ZENKI	'//前期開設

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
    <title>成績登録</title>
    <link rel=stylesheet href="../../common/style.css" type=text/css>

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

	   //alert("<%=m_SQL%>");
	    alert("<%=C_TOUROKU_OK_MSG%>");

	    document.frm.target = "main";
	    document.frm.action = "./sei0100_bottom.asp"
		document.frm.LOG_SOSA.value = "登録";		//add 2023/12/14 吉田
		document.frm.LOG_TAISYO.value = document.frm.LOG_TAISYO.value;		//add 2023/12/14 吉田
	    document.frm.submit();
	    return;

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type=hidden name=txtNendo    value="<%=trim(Request("txtNendo"))%>">
	<input type=hidden name=txtKyokanCd value="<%=trim(Request("txtKyokanCd"))%>">
	<input type=hidden name=txtSikenKBN value="<%=trim(Request("txtSikenKBN"))%>">
	<input type=hidden name=txtGakuNo   value="<%=trim(Request("txtGakuNo"))%>">
	<input type=hidden name=txtClassNo  value="<%=trim(Request("txtClassNo"))%>">
	<input type=hidden name=txtKamokuCd value="<%=trim(Request("txtKamokuCd"))%>">
	<input type=hidden name=txtGakkaCd  value="<%=trim(Request("txtGakkaCd"))%>">
	<input type="hidden" name="hidSeisekiInp" value="<%=trim(request("hidSeisekiInp"))%>">
	<input type="hidden" name="txtZokuseiCd" value="<%=trim(Request("txtZokuseiCd"))%>">
	<!-- ADD START 2023/12/14 吉田 WEBアクセスログカスタマイズ -->
	<input type="hidden" name="LOG_TAISYO" value="<%=request("LOG_TAISYO")%>">
	<input type="hidden" name="LOG_SOSA" value="<%=request("LOG_SOSA")%>">
	<!-- ADD END 2023/12/14 吉田 WEBアクセスログカスタマイズ -->
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>