<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 教官別授業時間一覧
'* ﾌﾟﾛｸﾞﾗﾑID : jik/jik0200/main.asp
'* 機      能: 下ページ 時間割マスタの一覧リスト表示を行う
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*           cboKyokaKeiCd      :科目系列コード
'*           cboKyokanCd      :教官コード
'*           txtMode         :動作モード
'           :session("PRJ_No")      '権限ﾁｪｯｸのキー
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'* 説      明:
'*           選択された教官の授業時間一覧を表示
'*-------------------------------------------------------------------------
'* 作      成: 2001/07/03 根本 直美
'* 変      更: 2001/07/30 根本 直美 戻り先URL変更
'*                                  変数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '取得したデータを持つ変数
    Public  m_iSyoriNen         ':処理年度
    Public  m_iKyokanCd         ':教官コード
    Public  m_iSKyokanCd        ':選択教官コード
    
    Public  m_Rs                'recordset
	Public  gRs					'代替時間割ﾚｺｰﾄﾞｾｯﾄ
    
    Public  m_iGakunen          ':学年
    Public  m_sClass            ':クラス
    Public  m_sYobi             ':表示曜日
    Public  m_iYobiCd           ':曜日コード
    Public  m_iJigen            ':時限
    Public  m_iKamokuCd         ':科目コード
    Public  m_sKamoku           ':科目名
    Public  m_iKyosituCd        ':教室コード
    Public  m_sKyositu          ':教室名
    
    Public  m_sCellD             ':テーブルセル色（曜日）'//2001/07/30変更
    Public  m_iJMax             ':最大時限数
    Public  m_Flg			'時間割１限目確認フラグ
    
    'ページ関係
    Public  m_iMax              ':最大ページ
    Public  m_iDsp              '// 一覧表示行数

    'データ取得用
    Public  m_iYobiCnt          ':カウント（曜日）
    Public  m_iJgnCnt           ':カウント（時限）
    Public  m_iYobiCCnt         ':カウント（曜日・テーブル色表示用）
    
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
    w_sMsgTitle="教官別授業時間一覧"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    Do

'response.write "AAA"
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
        
        '授業時間割テーブルマスタを取得
		w_sSQL = ""
		w_sSQL = w_sSQL & "SELECT"
		w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_GAKUNEN"
		w_sSQL = w_sSQL & vbCrLf & " ,M05_CLASS.M05_CLASSRYAKU"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_YOUBI_CD"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_JIGEN"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_KAMOKU"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_KYOSITU"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_TUKU_FLG"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " FROM T20_JIKANWARI"
		w_sSQL = w_sSQL & vbCrLf & ", M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE " 
		w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND T20_GAKKI_KBN = " & Session("GAKKI") & " "
		w_sSQL = w_sSQL & vbCrLf & " AND M05_CLASS.M05_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_KYOKAN = '" & m_iSKyokanCd & "' "
		w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN(+) "
		w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO(+) "
		w_sSQL = w_sSQL & vbCrLf & " Order By "
		w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_CLASS "

'response.write w_sSQL
'response.end

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            m_sErrMsg = "レコードセットの取得に失敗しました"
            Exit Do 'GOTO LABEL_MAIN_END
        End If

        '//最大時限数を取得
        Call gf_GetJigenMax(m_iJMax)
		if m_iJMax = "" Then
			m_bErrFlg = True
			m_sErrMsg = Err.description
			Exit Do
		end if

        '// ページを表示
        Call showPage()
        Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  代替時間割取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  2001/12/20追加
'********************************************************************************
Function f_GetDaigae(p_iYoubiCD,p_iJigen,p_sKamoku,p_flg)

    On Error Resume Next
    Err.Clear

	f_GetDaigae = False

	Dim wSQL

	'??????????????????????????????????????????????????????
	'              SQLを書く（要デバック）
	'??????????????????????????????????????????????????????
	wSQL = ""
	wSQL = wSQL & " SELECT "
	wSQL = wSQL & " T16_KAMOKUMEI "
	wSQL = wSQL & " FROM T23_DAIGAE_JIKAN ,T16_RISYU_KOJIN "
	wSQL = wSQL & " WHERE "
	wSQL = wSQL & " T23_KYOKAN = '" & m_iKyokanCd & "' AND "
	wSQL = wSQL & " T23_NENDO = " & m_iSyoriNen & " AND "
	wSQL = wSQL & " T23_GAKKI_KBN = " & Session("GAKKI") & " AND "
	wSQL = wSQL & " T23_YOUBI_CD = " & p_iYoubiCD & " AND "
	wSQL = wSQL & " T23_JIGEN = " & p_iJigen & " AND "
	wSQL = wSQL & " T16_NENDO = T23_NENDO AND "
	wSQL = wSQL & " T16_KAMOKU_CD = T23_KAMOKU "

'response.write wSQL
	Set gRs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(gRs, wSQL)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		m_sErrMsg = "1レコードセットの取得に失敗しました"
		Exit Function
	End If

	'科目コードがあれば、返値として持たせる。
	if gRs.EOF then 
		p_sKamoku = p_sKamoku
		p_flg = 0
	Else
		p_sKamoku = gRs("T16_KAMOKUMEI")
		p_flg = 1
	End if	

	'オブジェクトクローズ
	gf_closeObject(gRs)

f_GetDaigae = True

End Function

'********************************************************************************
'*  [機能]  全項目を空白に初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetBlank()

    m_iKyokanCd = ""
    m_iSyoriNen = ""
    
    m_iKyokaKeiKbn = ""
    m_iSKyokanCd = ""
    
    m_sYobi = ""
    m_iYobiCd = ""
    m_iJigen = ""
    m_iKamokuCd = ""
    
    m_iYobiCnt = ""
    m_iJgnCnt = ""
    m_iYobiCCnt = ""
    
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")          ':教官コード
    m_iSyoriNen = Session("NENDO")              ':処理年度
    m_iSKyokanCd = Request("SKyokanCd1")        ':選択教官コード
    
End Sub

Sub s_ShowYobi(p_iJigenMax)    '//2001/07/30変更
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  曜日を表示（テーブル用）
'********************************************************************************
    On Error Resume Next
    Err.Clear


	if m_iYobiCCnt Mod 2 <> 0 Then
	    m_sCellD = ""
	end if

	call gs_cellPtn(m_sCellD)
    if m_iJgnCnt <= 1  And m_Flg = 0 Then
	m_Flg = 1
        'response.write "<td rowspan=8 class="
        response.write "<td rowspan=" & p_iJigenMax & " class="
'        response.write "<td rowspan=" & m_iJMax & " class="
        'call showYobiColor()
        response.write m_sCellD
        response.write ">" & WeekdayName(m_iYobiCnt,True) & "</td>"
    else
    end if
    
End Sub

Function f_ShowClass()   '//2001/07/30変更
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  学年・クラスを表示
'********************************************************************************
    On Error Resume Next
    Err.Clear


    Dim w_sClass
    Dim w_clsMax
    Dim w_clsCnt

    m_iGakunen = ""
    m_sClass = ""
    
    w_sClass = ""
    f_ShowClass = ""
	w_clsCnt = 0
    
	w_clsMax = f_GetClassMax(m_iSyoriNen,m_Rs("T20_GAKUNEN"))

    Do Until m_Rs.EOF

		'曜日と時限が同じ間は文字列作成
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then

			'初期文字列編集
			If m_iGakunen = "" Then
				f_ShowClass =  f_ShowClass & cstr(m_Rs("T20_GAKUNEN")) & "-" & m_Rs("M05_CLASSRYAKU")
	            m_iGakunen = m_Rs("T20_GAKUNEN")
	            m_sClass = m_Rs("M05_CLASSRYAKU")
				w_clsCnt = 1
			End If

			'学年が変わったら文字列編集
			if cstr(m_iGakunen) <> cstr(m_Rs("T20_GAKUNEN")) then
				If w_clsCnt = w_clsMax Then
					f_ShowClass =  m_iGakunen & "-全"
				End If

				f_ShowClass =  f_ShowClass & "<BR>" & cstr(m_Rs("T20_GAKUNEN")) & "-" & m_Rs("M05_CLASSRYAKU")
	            m_iGakunen = m_Rs("T20_GAKUNEN")
			Else
				
				'クラスが変わったら文字列編集
				if cstr(m_sClass) <> cstr(m_Rs("M05_CLASSRYAKU")) then
					w_clsCnt = w_clsCnt + 1

'response.write w_clsCnt & "=" & w_clsMax
					If w_clsCnt = w_clsMax Then
						f_ShowClass =  m_iGakunen & "-全"
						w_clsCnt = 1
					Else
						f_ShowClass =  f_ShowClass & "," & cstr(m_Rs("M05_CLASSRYAKU"))
					End If
		            m_sClass = m_Rs("M05_CLASSRYAKU")
				End if
			End if
			
        End if
        
        m_Rs.MoveNext
    Loop

    m_Rs.MoveFirst
    
    'if m_iGakunen <> "" and m_sClass <> "" Then
	'	If f_ShowClass = "" Then
	'        f_ShowClass = f_ShowClass & m_iGakunen & "-" & m_sClass
	'	Else
	'        f_ShowClass = f_ShowClass & "," & m_sClass
	'	End If
    'End if

End Function

Function f_ShowKamokuMei()   '//2001/07/30変更
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  科目名を表示
'********************************************************************************
    On Error Resume Next
    Err.Clear

    dim w_iCourseCD
	dim w_sKamoku

    w_sKamoku =""
    
    m_sKamoku = ""
    f_ShowKamokuMei = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then

'            m_sKamoku = f_getKamokumei(m_iSyoriNen,m_Rs("T20_KAMOKU"),m_Rs("T20_GAKUNEN"),m_Rs("T20_TUKU_FLG"),w_iCourseCD) 
		
            m_sKamoku = f_getKamokumei(m_iSyoriNen,m_Rs("T20_KAMOKU"),m_Rs("T20_GAKUNEN"),m_Rs("T20_TUKU_FLG"),w_iCourseCD)

            if CStr(w_sKamoku) <> "" And w_sKamoku <> m_sKamoku then
            	w_sKamoku = w_sKamoku & "<BR>" & m_sKamoku
            Else
            	w_sKamoku = m_sKamoku
            End If

        Else
        End If
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_iGakunen <> "" and m_sClass <> "" Then
'        f_ShowKamokuMei = m_sKamoku		こっちをコメントアウトすれば、表示される科目は1つになります
        f_ShowKamokuMei = w_sKamoku
    end if

End Function

Function f_KamokuSu(p_iYobiCnt)   '//2001/09/06 add
'********************************************************************************
'*  [機能]  科目数を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  科目名を表示
'********************************************************************************
    On Error Resume Next
    Err.Clear

    f_KamokuSu = cint(m_iJMax)
    m_Rs.MoveFirst
    Do Until m_Rs.EOF

        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(p_iYobiCnt) and right(cstr(cDbl(m_Rs("T20_JIGEN"))*10),1) <> "0" Then
            f_KamokuSu = f_KamokuSu + 1
        end if
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

'    if m_iGakunen <> "" and m_sClass <> "" Then
'        response.write m_sKamoku
'    end if

End Function

Sub s_SetKyositu()  '//2001/07/30変更
'********************************************************************************
'*  [機能]  値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  教室コードを設定
'********************************************************************************

    On Error Resume Next
    Err.Clear

    m_iKyosituCd = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then
            m_iKyosituCd = m_Rs("T20_KYOSITU")
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

End Sub

'********************************************************************************
'*  [機能]  教室名の取得
'*  [引数]  
'*  [戻値]  0:情報取得成功、1:レコードなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetKyosituMei()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    Call s_SetKyositu()
    f_GetKyosituMei = 0
    m_sKyositu = ""
    
    if m_iKyosituCd <> "" Then
        Do
            
            '// 教室名マスタﾚｺｰﾄﾞｾｯﾄを取得
            w_sSQL = ""
            w_sSQL = w_sSQL & "SELECT"
            w_sSQL = w_sSQL & " M06_KYOSITUMEI"
            w_sSQL = w_sSQL & " FROM M06_KYOSITU "
            w_sSQL = w_sSQL & " WHERE M06_NENDO = " & m_iSyoriNen
            w_sSQL = w_sSQL & " AND M06_KYOSITU_CD = " & m_iKyosituCd
            
            w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
            
            If w_iRet <> 0 Then
                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                'response.write w_iRet & "<br>"
                'm_sErrMsg = "ﾚｺｰﾄﾞｾｯﾄの取得失敗"
                'm_bErrFlg = True
                f_GetKyosituMei = 99
                Exit Do 'GOTO LABEL_f_GetKyosituMei_END
            Else
            End If
            
            If w_Rs.EOF Then
                '対象ﾚｺｰﾄﾞなし
                'm_sErrMsg = "対象ﾚｺｰﾄﾞなし"
                f_GetKyosituMei = 1
                Exit Do 'GOTO LABEL_f_GetKyosituMei_END
            End If
            
                '// 取得した値を格納
                    m_sKyositu = w_Rs("M06_KYOSITUMEI")    '//教室名を格納
            '// 正常終了
            Exit Do
        
        Loop
        
        gf_closeObject(w_Rs)
    
    end if

'// LABEL_f_GetKyosituMei_END
End Function

Function f_getKamokumei(p_iNendo,p_sKamokuCD,p_iGaknen,p_iTUKU,p_iCourseCD) 
'********************************************************************************
'*  [機能]  科目名の取得(ついでにコースも取ってきます。)
'*  [引数]  
'*  [戻値]  科目名
'*  [説明]  2001/9/15
'********************************************************************************
    dim w_sSQL,w_Rs,w_iRet
    
    On Error Resume Next
    Err.Clear
    
    f_getKamokumei = "-"
    p_iCourseCD = 0 
    
  Do

   if p_iTUKU = C_TUKU_FLG_TOKU then '特別活動のときは、M41特別活動マスタから名称取得
    	w_sSQL = ""
    	w_sSQL = w_sSQL & vbCrLf & "SELECT "
    	w_sSQL = w_sSQL & vbCrLf & "M41_MEISYO "
    	w_sSQL = w_sSQL & vbCrLf & "FROM  "
    	w_sSQL = w_sSQL & vbCrLf & "M41_TOKUKATU "
    	w_sSQL = w_sSQL & vbCrLf & "WHERE "
    	w_sSQL = w_sSQL & vbCrLf & "M41_NENDO = " & p_iNendo & " AND "
    	w_sSQL = w_sSQL & vbCrLf & "M41_TOKUKATU_CD = '" & p_sKamokuCD & "' "
    	w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
    	w_sSQL = w_sSQL & vbCrLf & "M41_MEISYO "

   	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
   	If w_iRet <> 0 OR w_Rs.EOF = true Then Exit Do 
		
   	f_getKamokumei = w_Rs("M41_MEISYO")

    Else '普通の授業のときは、T15履修から名称取得
    w_i = p_iNendo - p_iGaknen + 1
    	w_sSQL = ""
    	w_sSQL = w_sSQL & vbCrLf & "SELECT "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKUMEI, "
    	w_sSQL = w_sSQL & vbCrLf & "T15_COURSE_CD"
    	w_sSQL = w_sSQL & vbCrLf & "FROM "
    	w_sSQL = w_sSQL & vbCrLf & " T15_RISYU "
    	w_sSQL = w_sSQL & vbCrLf & "WHERE "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKU_CD = '" & p_sKamokuCD & "' AND "
    	w_sSQL = w_sSQL & vbCrLf & "T15_NYUNENDO = "& p_iNendo - cint(p_iGaknen) + 1 &" "
    	w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKUMEI, T15_COURSE_CD"

   	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    	If w_iRet <> 0 OR w_Rs.EOF = true Then Exit Do 

    	f_getKamokumei = w_Rs("T15_KAMOKUMEI")
    	p_iCourseCD = w_Rs("T15_COURSE_CD") 

    End if
    Exit Do

   Loop

end Function

'********************************************************************************
'*  [機能]  クラス数を取得する
'*  [引数]  p_iNendo  ：処理年度
'*          p_iGakuNen：学年
'*  [戻値]  f_GetClassMax：クラス数
'*  [説明]  
'********************************************************************************
Function f_GetClassMax(p_iNendo,p_iGakuNen)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassMax = 0

	Do

		'//クラス名称取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  COUNT(M05_CLASSNO) as ClassMax"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M05_NENDO=" & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND M05_GAKUNEN=" & p_iGakuNen

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
'response.write w_sSql&vbCrLf&"<BR>"
'response.write w_iRet
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//クラス名
			f_GetClassMax = cint(rs("ClassMax"))
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
'	f_GetClassMax = rs("ClassMax")

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

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

    Dim w_cellT             '//Tableセル色
    Dim w_sClass          '//クラス名
    Dim w_sKamoku       '//科目名
    Dim w_iJigenMax      '//曜日のrow用
    Dim w_iJgnNo          '//表示用時限
	Dim w_daigaeF         '//代替フラグ

    On Error Resume Next
    Err.Clear

%>


<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
</head>

<body>

<center>

<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr>
<td align="center">

    <table border=1 class=hyo width="100%">
    <COLGROUP WIDTH="5%" ALIGN=center>
    <COLGROUP WIDTH="5%" ALIGN=center>
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="35%" ALIGN=center>
    <COLGROUP WIDTH="35%" ALIGN=center>
    <tr>
        <th colspan="2" class=header><br></th>  
        <th class=header>クラス</th>
        <th class=header>教科名</th>
        <th class=header>教室</th>
    </tr>

<%
m_iYobiCCnt = 1

'For m_iYobiCnt = 2 to 6
For m_iYobiCnt = C_YOUBI_MIN to C_YOUBI_MAX

	m_Flg = 0
	w_iJigenMax =  f_KamokuSu(m_iYobiCnt)
    For m_iJgnCnt = 0.5 to m_iJMax step 0.5
		w_sClass =  f_ShowClass()
		w_sKamoku = f_ShowKamokuMei()
		w_iJgnNo = m_iJgnCnt
		w_daigaeF = 0	'代替フラグの初期化

		If right(cstr(w_iJgnNo*10),1) <> "0" then	'0.5時限の場合
			w_iJgnNo = " "
		Else										'普通の時間の場合

			'通常の時間割科目がないとき、代替科目の時間割取得
'			If w_sKamoku = "" and w_sClass = "" then 
				If f_GetDaigae(m_iYobiCnt,m_iJgnCnt,w_sKamoku,w_daigaeF) <> true then
					m_bErrFlg = True
					m_sErrMsg = "レコードセットの取得に失敗しました"
					Exit Sub
				End if

				if w_daigaeF = 1 then
					w_sClass = "代替科目"
					m_sKyositu = ""
				end if
'			End if

		End If

		If w_sKamoku <> "" OR w_sClass <> "" OR right(cstr(m_iJgnCnt*10),1) = "0" Then 
			call gs_cellPtn(w_cellT)

			'教室名の取得（代替科目を除く）
			if w_daigaeF <> 1 then 
				call f_GetKyosituMei()
'				    if f_GetKyosituMei = 0 Then
'				        response.write m_sKyositu
'				    else
'				    end if
			end if
			%>
			    <tr>
			<%call s_ShowYobi(w_iJigenMax)%>
			        <td class=<%=w_cellT%>><%=w_iJgnNo%></td>
			        <td class=<%=w_cellT%>><%=w_sClass%><br></td>
			        <td class=<%=w_cellT%>><%=w_sKamoku%><br></td>
			        <td class=<%=w_cellT%>><%=m_sKyositu%><br></td>
			    </tr>
			<%
		End If
    Next
m_iYobiCCnt = m_iYobiCCnt + 1   '//曜日カウント（テーブル背景色表示用）

Next
%>

    </table>

</td>
</tr>
</table>
</center>

</body>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>
