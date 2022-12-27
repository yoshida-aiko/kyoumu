<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索結果
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/main.asp
' 機      能: 下ページ 学籍データの検索結果を表示する
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtHyoujiNendo         :表示年度
'           txtGakunen             :学年
'           txtGakkaCD             :学科
'           txtClass               :クラス
'           txtName                :名称
'           txtGakusekiNo          :学籍番号
'           txtSeibetu             :性別
'           txtGakuseiNo           :学生番号
'           txtIdou                :異動
'           txtTyuClub             :中学校クラブ
'           txtClub                :現在クラブ
'           txtRyoseiKbn           :寮
'           CheckImage               :画像表示指定
'           txtMode                :動作モード
'                               BLANK   :初期表示
'                               SEARCH  :結果表示
' 説      明:
'           ■初期表示
'               タイトルのみ表示
'           ■結果表示
'               上ページで設定された検索条件にかなう学生情報を表示する
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	
    '取得したデータを持つ変数
    Public  m_TxtMode      	       ':動作モード
    Public  m_PgMode      	       ':プログラムモード
	Public  m_iSyoriNen      	   ':処理年度
    Public  m_iHyoujiNendo         ':表示年度
    Public  m_sGakunen             ':学年
    Public  m_sGakkaCD             ':学科
    Public  m_sClass               ':クラス
    Public  m_sName                ':名称
    Public  m_sGakusekiNo          ':学籍番号
    Public  m_sSeibetu             ':性別
    Public  m_sGakuseiNo           ':学生番号
    Public  m_sIdou                ':異動
    Public  m_sTyuClub             ':中学校クラブ
    Public  m_sClub                ':現在クラブ
    Public  m_sRyoseiKbn           ':寮
    Public  m_sCheckImage          ':画像表示指定
	Public  m_sTyugaku			   ':出身中学校
    Public  m_sMsgTitle            ':ﾀｲﾄﾙ
	  
    Public	m_Rs					'recordset
    Public	m_RsGakka
    
    Public	m_iDsp					'一覧表示行数

    Public  m_iPageTyu      		':表示済表示頁数（自分自身から受け取る引数）

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
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_iDsp=15

	'//セッション情報・動作モードの取得
	m_iSyoriNen = Session("NENDO")
    m_TxtMode=request("txtMode")
    m_PgMode=request("p_mode")
    
	Select Case m_PgMode
		Case "P_HAN0100"
		    m_sMsgTitle="成績一覧表"
		Case "P_KKS0200"
		    m_sMsgTitle="欠課一覧表"
		Case "P_KKS0210"
		    m_sMsgTitle="遅刻一覧表"
		Case "P_KKS0220"
		    m_sMsgTitle="行事欠課一覧表"
		Case Else
	End Select
	w_sMsgTitle = m_sMsgTitle


    Do
		if m_TxtMode = "" then
           	Call showPage()
			Exit Do
		End if

        '// ﾃﾞｰﾀﾍﾞｰｽ接続
		w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 権限チェックに使用
		session("PRJ_No") = C_LEVEL_NOCHK

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))


        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        'データ抽出SQLを作成する
		'クラス権限等でＳＱＬ文を作成する。
        Call s_MakeSQL(w_sSQL)
		
		'If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then
		'	m_bErrFlg = True
		'	exit do
		'End If
		
		
       'レコードセットの取得
        Set m_Rs = Server.CreateObject("ADODB.Recordset")

		w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// ページを表示
        If m_Rs.EOF Then
            Call showPage_NoData()
        Else
			'PDF情報表示
           	Call showPage()
        End If

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

Sub s_SetParam()
'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

'    Session("HyoujiNendo") = request("txtHyoujiNendo")     	'表示年度
'    Session("HyoujiNendo") = Session("NENDO")		'表示年度	'<-- 8/16修正	持
'
'	m_iDsp = cint(request("txtDisp"))						':検索リストの表示件数
'
    '// BLANKの場合は行数ｸﾘｱ
    If m_TxtMode = "Search" Then
        m_iPageTyu = 1
    Else
        m_iPageTyu = int(Request("txtPageTyu"))     ':表示済表示頁数（自分自身から受け取る引数）
    End If

End Sub


'********************************************************************************
'*  [機能]  PDF情報データ抽出SQL文字列の作成
'*  [引数]  p_sSql - SQL文字列
'*  [戻値]  なし 
'*  [説明]  
'********************************************************************************
Sub s_MakeSQL(p_sSql)

	Dim w_iRet
	Dim w_sId
	Dim w_sWhere
	Dim w_iGakunen
	Dim w_iClassNo
	Dim w_sGakka
	

	'処理権限のIDを取得
	w_iRet = f_GetSyori(w_sId)

'response.write "getid" & w_sId


	w_sWhere = ""
	Select Case w_sId
		'//FULL権限
        Case C_ID_SEI0200
			'Full権限の場合はWhere条件を指定しない。
		'学科別 //////////////////////////////////////////////////////////////////////////////////////////////////
		'Add 2002.1.13 okada
        Case C_ID_SEI0210
        
			'//教官が所属する学科を取得
			if Not f_GetKyokanGakka(m_iSyoriNen,session("KYOKAN_CD"),w_sGakka) Then
					m_bErrFlg = True
					m_sErrMsg = "所属学科の取得に失敗しました。"
					Exit sub        
			End if
 

			'// 学科Ｇｒｐからのm_RsGakka の取得
			if Not f_GetGakkaGrp(m_iSyoriNen,w_sGakka) Then
					m_bErrFlg = True
					m_sErrMsg = "学科Ｇｒｐの取得に失敗しました。"
					Exit sub				
			End if
			w_sWhere = w_sWhere & "( "
			Do until m_RsGakka.EOF'//戻り値ｾｯﾄ
	
				'// 学年とクラスを取得する。
				w_iClassNo = ""
				'w_iGakunen = ""
'response.write m_RsGakka("M23_GAKUNEN")
'Response.end				
				'学年とクラスを取得する
				if Not f_GetClass(m_iSyoriNen,m_RsGakka("M23_GAKUNEN"),w_iClassNo,m_RsGakka("M23_GAKKA_CD")) then
					m_bErrFlg = True
					m_sErrMsg = "クラスデータの取得に失敗しました。"
					Exit sub				
				end if
				
				w_sWhere = w_sWhere & "(T76_GAKUNEN = " & m_RsGakka("M23_GAKUNEN") & " And T76_CLASS = " & w_iClassNo & ") "				
								
				m_RsGakka.MoveNext
				if Not m_RsGakka.Eof then	'ＭｏｖｅＮｅｘｔする場合はＯＲを追加
					w_sWhere = w_sWhere & " or "
				Else
					w_sWhere = w_sWhere & " ) AND "
				End if
			Loop
			
			'//ﾚｺｰﾄﾞｾｯﾄCLOSE
			Call gf_closeObject(m_RsGakka)
		'/////////////////////////////////////////////////////////////////////////////////////////////////////////////	

		'1年生
        Case C_ID_SEI0221
			w_sWhere = " T76_GAKUNEN = 1 AND "
		'2年生
        Case C_ID_SEI0222
			w_sWhere = " T76_GAKUNEN = 2 AND "
		'3年生
        Case C_ID_SEI0223
			w_sWhere = " T76_GAKUNEN = 3 AND "
		'4年生
        Case C_ID_SEI0224
			w_sWhere = " T76_GAKUNEN = 4 AND "
		'5年生
        Case C_ID_SEI0225
			w_sWhere = " T76_GAKUNEN = 5 AND "

		'担任
        Case C_ID_SEI0230

			'//担任権限の場合は、そのユーザーが担任しているクラス
	        '
			If Not f_GetTanninClass(m_iSyoriNen,session("KYOKAN_CD"),w_iGakunen,w_iClassNo) Then
				m_bErrFlg = True
				m_sErrMsg = "クラスデータの取得に失敗しました。"
				Exit sub
			End If

			w_sWhere = " T76_GAKUNEN = " & w_iGakunen & " AND "
			w_sWhere = w_sWhere & " T76_CLASS = " & w_iClassNo & " AND "

'response.write " Where文 " & w_sWhere
		
		Case Else

	End Select

'response.write " Where文 " & w_sWhere

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " T76_NENDO      , "
    p_sSql = p_sSql & " T76_GAKKI_KBN  , "
    p_sSql = p_sSql & " T76_TYOHYO_ID  , "
    p_sSql = p_sSql & " T76_TYOHYO_NAME, "
    p_sSql = p_sSql & " T76_SIKEN_KBN  , "
    p_sSql = p_sSql & " T76_SIKENMEI   , "
    p_sSql = p_sSql & " T76_GAKUNEN    , "
    p_sSql = p_sSql & " T76_CLASS      , "
    p_sSql = p_sSql & " T76_CLASSMEI   , "
    p_sSql = p_sSql & " T76_PATH       , "
    p_sSql = p_sSql & " T76_FILENAME   , "
    p_sSql = p_sSql & " T76_INS_DATE     "
    p_sSql = p_sSql & " FROM T76_PDF "
    p_sSql = p_sSql & " WHERE "
    p_sSql = p_sSql & w_sWhere
    p_sSql = p_sSql & " T76_NENDO = " & m_iSyoriNen & " AND "
    p_sSql = p_sSql & " T76_TYOHYO_ID = '" & m_PgMode & "' "
    p_sSql = p_sSql & " ORDER BY "
    p_sSql = p_sSql & " T76_TYOHYO_ID, "
    p_sSql = p_sSql & " T76_SIKEN_KBN DESC, "
    p_sSql = p_sSql & " T76_GAKUNEN ,"
    p_sSql = p_sSql & " T76_CLASS "

'response.write p_sSql & "<BR>"

End Sub

'********************************************************************************
'*  [機能]  権限種別取得
'*  [引数]  p_iMenuID：項目のNO
'*  [戻値]  true/false、p_iMenuID:権限ID
'*  [説明]  許可されている権限の権限IDを取得
'********************************************************************************
Function f_GetSyori(p_iMenuID)
	Dim w_sLevel
	Dim w_iRet,w_Rs,w_sSq
	Dim w_Where

	Dim w_iCnt
	
	f_GetSyori = false

	'// Session("LEVEL")がNULLなら、ぬける
	if gf_IsNull(Trim(Session("LEVEL"))) then Exit Function

	'// Session("LEVEL")が"0"なら、ぬける
	if Cint(Session("LEVEL")) = Cint(0) then Exit Function

	w_sLevel = "T51_LEVEL" & Trim(Session("LEVEL"))

	'// WHERE文作成

    Do
		w_sSql = ""
		w_sSql = w_sSql & "Select "
		w_sSql = w_sSql & "T51_ID,"
		w_sSql = w_sSql &  w_sLevel & " "
		w_sSql = w_sSql & "From T51_SYORI_LEVEL "
		w_sSql = w_sSql & "Where "
		w_sSql = w_sSql & w_sLevel & " = 1 AND "
		w_sSql = w_sSql & "T51_ID in ("
		w_sSql = w_sSql & "'" & C_ID_SEI0200 & "',"	'FULL権限
		w_sSql = w_sSql & "'" & C_ID_SEI0210 & "',"	'学科別
		w_sSql = w_sSql & "'" & C_ID_SEI0221 & "',"	'1年生
		w_sSql = w_sSql & "'" & C_ID_SEI0222 & "',"	'2年生
		w_sSql = w_sSql & "'" & C_ID_SEI0223 & "',"	'3年生
		w_sSql = w_sSql & "'" & C_ID_SEI0224 & "',"	'4年生
		w_sSql = w_sSql & "'" & C_ID_SEI0225 & "',"	'5年生
		w_sSql = w_sSql & "'" & C_ID_SEI0230 & "' "	'担任
		w_sSql = w_sSql & ") "
		w_sSql = w_sSql & " ORDER BY T51_ID "

'response.write " 権限取得p_sSql=" & w_sSql & "<BR>"
		w_iRet = gf_GetRecordset(w_Rs, w_sSql)

		If w_iRet <> 0 Then
		    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		    Exit Do
		End If

		If w_Rs.EOF = true Then
		    '該当無し
		    Exit Do
		End If

		w_flg = false
		w_Rs.movefirst
		Do Until w_Rs.EOF
			If trim(gf_SetNull2String(w_Rs(w_sLevel))) = "1" then 

				p_iMenuID = trim(gf_SetNull2String(w_Rs("T51_ID")))

				w_flg = true
				Exit Do
			End If

'response.write "OK!" & w_iCnt & "<BR>"

			w_Rs.movenext
		Loop

		w_Rs.close
		Set w_Rs = Nothing

		If w_flg <> true Then
		    '対象ﾚｺｰﾄﾞなし
		    Exit Do
		End If

		f_GetSyori = true

		'// 正常終了
		Exit Do

    Loop

End Function

'********************************************************************************
'*  [機能]  クラスコードと学年を取得する
'*  [引数]  p_iNendo   ：処理年度
'*          p_sKyokanCd：教官コード
'*          p_iGakunen ：学年
'*          p_iClassNo ：クラスNO
'*  [戻値]  gf_GetClassName：クラス名
'*  [説明]  
'********************************************************************************
Function f_GetTanninClass(p_iNendo,p_sKyokanCd,p_iGakunen,p_iClassNo)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetTanninClass = False

	p_iGakunen = 0
	p_iClassNo = 0
	
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " SELECT "
	w_sSql = w_sSql & vbCrLf & "  M05_GAKUNEN,"
	w_sSql = w_sSql & vbCrLf & "  M05_CLASSNO "
	w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
	w_sSql = w_sSql & vbCrLf & " WHERE "
	w_sSql = w_sSql & vbCrLf & "      M05_NENDO=" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "  AND M05_TANNIN=" & p_sKyokanCd

	'//ﾚｺｰﾄﾞｾｯﾄ取得
	w_iRet = gf_GetRecordset(rs, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit Function
	End If

	'//データが取得できたとき
	If rs.EOF = False Then
		p_iGakunen = rs("M05_GAKUNEN")	'学年
		p_iClassNo = rs("M05_CLASSNO")	'クラスNO
	End If

	'//戻り値ｾｯﾄ
	f_GetTanninClass = True

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  教官の所属する学科を取得
'*  [引数]  p_iNendo   ：処理年度
'*          p_sKyokanCd：教官コード
'*          p_iGakkaCd ：学科コード
'*  [戻値]  gf_GetClassName：p_iGakkaCd ：学科コード
'*  [説明]  
'********************************************************************************
Function f_GetKyokanGakka(p_iNendo,p_sKyokanCd,p_sGakkaCd)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetKyokanGakka = False

	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " SELECT "
	w_sSql = w_sSql & vbCrLf & "  M04_GAKKA_CD "
	w_sSql = w_sSql & vbCrLf & " FROM M04_KYOKAN"
	w_sSql = w_sSql & vbCrLf & " WHERE "
	w_sSql = w_sSql & vbCrLf & "      M04_NENDO=" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "  AND M04_KYOKAN_CD =" & p_sKyokanCd

	'//ﾚｺｰﾄﾞｾｯﾄ取得
	w_iRet = gf_GetRecordset(rs, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit Function
	End If

	'//データが取得できたとき
	If rs.EOF = False Then
		p_sGakkaCd = rs("M04_GAKKA_CD")	'学科コード
	End If

	'//戻り値ｾｯﾄ
	f_GetKyokanGakka = True

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  学科ＣＤから学科Ｇｒｐを取得
'*  [引数]  p_iNendo   ：処理年度
'*          p_sKyokanCd：教官コード
'*          p_iGakkaCd ：学科コード
'*  [戻値]  gf_GetClassName：p_iGakkaCd ：学科コード
'*  [説明]  
'********************************************************************************
Function f_GetGakkaGrp(p_iNendo,p_sGakkaCd)
	Dim w_iRet
	Dim w_sSQL
	Dim rs
	Dim w_sGakaGrp	

	On Error Resume Next
	Err.Clear

	f_GetGakkaGrp = False

	'学科ＣＤからＧｒｐを取得
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " Select "
	w_sSql = w_sSql & vbCrLf & "	M23_GROUP,"
	w_sSql = w_sSql & vbCrLf & "	M23_GAKKA_CD "
	w_sSql = w_sSql & vbCrLf & " From "
	w_sSql = w_sSql & vbCrLf & "	M23_GAKKA_GRP "
	w_sSql = w_sSql & vbCrLf & " Where "
	w_sSql = w_sSql & vbCrLf & "	M23_NENDO =" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "	AND M23_GAKKA_CD =" & p_sGakkaCd
	w_sSql = w_sSql & vbCrLf & " Order By M23_GROUP "

'response.write w_sSql
 
	'//ﾚｺｰﾄﾞｾｯﾄ取得
	w_iRet = gf_GetRecordset(rs, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit Function
	End If

	'//データが取得できたとき
	If rs.EOF = False Then
		w_sGakaGrp = rs("M23_GROUP")	'学科コード
	End If
	
'response.write "OK!" & w_sGakaGrp 
'Response.End 	
	
	'そのＧｒｐに所属している学科を取得しなおす。
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " Select "
	w_sSql = w_sSql & vbCrLf & "	M23_GROUP,"
	w_sSql = w_sSql & vbCrLf & "	M23_GAKUNEN,"
	w_sSql = w_sSql & vbCrLf & "	M23_GAKKA_CD "
	w_sSql = w_sSql & vbCrLf & " From "
	w_sSql = w_sSql & vbCrLf & "	M23_GAKKA_GRP "
	w_sSql = w_sSql & vbCrLf & " Where "
	w_sSql = w_sSql & vbCrLf & "	M23_NENDO =" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "	AND M23_GROUP =" & w_sGakaGrp
	w_sSql = w_sSql & vbCrLf & " Order By M23_GROUP "

'response.write "OK!" & w_sSql 	
'Response.End 		
	sSet m_RsGakka = Server.CreateObject("ADODB.Recordset")
	'//ﾚｺｰﾄﾞｾｯﾄ取得
	w_iRet = gf_GetRecordset(m_RsGakka, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit Function
	End If

	'//データが取得できたとき aaaa
	If m_RsGakka.EOF = False Then
	
		'//戻り値ｾｯﾄ
		f_GetGakkaGrp = True
		
	End If	
'response.write "OK!" & "A"	
'Response.End 			
 			
	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)
'response.write "OK!" & "B"

End Function

'********************************************************************************
'*  [機能]  学科ＣＤからクラスコードを取得する
'*  [引数]  p_iNendo   ：処理年度
'*          p_sKyokanCd：教官コード
'*          p_iGakunen ：学年
'*          p_iClassNo ：クラスNO
'*  [戻値]  gf_GetClassName：クラス名
'*  [説明]  
'********************************************************************************
Function f_GetClass(p_iNendo,p_iGakunen,p_iClassNo,p_iGakkaCD)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClass = False

	'p_iGakunen = 0
	'p_iClassNo = 0
	
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " SELECT "
	w_sSql = w_sSql & vbCrLf & "  M05_CLASSNO "
	w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
	w_sSql = w_sSql & vbCrLf & " WHERE "
	w_sSql = w_sSql & vbCrLf & "      M05_NENDO=" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "  AND M05_GAKUNEN=" & p_iGakunen
	w_sSql = w_sSql & vbCrLf & "  AND M05_GAKKA_CD=" & p_iGakkaCD

'response.write w_sSql
	
	'//ﾚｺｰﾄﾞｾｯﾄ取得
	w_iRet = gf_GetRecordset(rs, w_sSQL)
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit Function
	End If

	'//データが取得できたとき
	If rs.EOF = False Then
		p_iClassNo = rs("M05_CLASSNO")	'クラスNO
	End If

	'//戻り値ｾｯﾄ
	f_GetClass = True

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage_NoData()

%>
	<html>
	<head>
	<title><%=m_sMsgTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
	<link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>

	<body>

	<center>
		<br><br><br>
		<span class="msg">対象データは存在しません</span>
	</center>

	</body>
    </html>

<%
    '---------- HTML END   ----------
End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	Dim w_pageBar			'ページBAR表示用
%>

<html>

<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
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

    }

    //************************************************************
    //  [機能]  詳細ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_detail(p_sUrl){

			url = p_sUrl;
			w   = 1015;
			h   = 710;
			wn  = "SubWindow";
			//opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
			opt = "left=0,top=0,directoris=no,location=no,menubar=no,scrollbars=yes,status=no,toolbar=no,resizable=yes";
			if (w > 0)
				opt = opt + ",width=" + w;
			if (h > 0)
				opt = opt + ",height=" + h;
			newWin = window.open(url, wn, opt);

    }

    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="main.asp?p_mode=<%=m_PgMode%>";
        document.frm.target="_self";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageTyu.value = p_iPage;
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
    </head>

    <body>
	<% if m_TxtMode = "" then %>
		<center>
		<br><br><br>
		<span class="msg">項目を選んで表示ボタンを押してください</span>
		</center>
	<% Else %>
	    <div align="center">
	    <form action="kojin.asp" method="post" name="frm" target="_detail">

		<BR>
		<table><tr><td align="center">
		<%
			'ページBAR表示
			Call gs_pageBar(m_Rs,m_iPageTyu,m_iDsp,w_pageBar)
		%>
		<%=w_pageBar %>

			<table border="0" width="100%">
				<tr>
					<td align="center">
					<% if m_TxtMode = "" then %>
						<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
							<tr>
								<td width="60">&nbsp</td>
								<td valign="top"></td>
							</tr>
						</table>
					<% else %>
						<% dim w_cell %>

					    <!--  PDF情報表示　-->

                        <font color="red">「詳細」ボタンをクリックして、成績一覧が表示されない、又はエラーメッセージが表示される場合は、<br></font>
                        <font color="#3366CC">Download</font><font color="red">をマウスで右クリックして「対象をファイルに保存」を選択してください。<br><br></font>

						<table border="1" width="610" class=hyo>
							<tr>
								<th align="center" height=16 class=header width="200"nowrap>試　験</th>
								<th align="center" height=16 class=header width="60"nowrap>学年</th>
								<th align="center" height=16 class=header width="80"nowrap>クラス</th>
								<th align="center" height=16 class=header width="125"nowrap>作　成　日</th>
								<th align="center" height=16 class=header width="45"nowrap>詳細</th>
								<th align="center" height=16 class=header width="100"nowrap >ダウンロード</th>
							</tr>

				        	<%
							w_iCnt = 1
							Do Until m_Rs.EOF or w_iCnt > m_iDsp
								call gs_cellPtn(w_cell)
							%>
								<tr>
									<td align="center" height="16" class=<%=w_cell%> width="200"nowrap><%=gf_HTMLTableSTR(m_Rs("T76_SIKENMEI")) %>&nbsp</td>
									<td align="center" height="16" class=<%=w_cell%> width="60"nowrap><%=gf_HTMLTableSTR(m_Rs("T76_GAKUNEN")) %>&nbsp</td>
									<td align="center" height="16" class=<%=w_cell%> width="80"nowrap><%=gf_HTMLTableSTR(m_Rs("T76_CLASSMEI")) %>&nbsp</td>
									<td align="center" height="16" class=<%=w_cell%> width="125"nowrap><%=gf_HTMLTableSTR(m_Rs("T76_INS_DATE")) %>&nbsp</td>
									<td align="center" height="16" class=<%=w_cell%> width="45"nowrap><input type=button class=button value="詳細" onclick="f_detail('../..<%=gf_HTMLTableSTR(m_Rs("T76_PATH")) %><%=gf_HTMLTableSTR(m_Rs("T76_FILENAME")) %>');"></td>
									<td align="center" height="16" class=<%=w_cell%> width="100"nowrap><a href='../..<%=gf_HTMLTableSTR(m_Rs("T76_PATH")) %><%=gf_HTMLTableSTR(m_Rs("T76_FILENAME")) %>'>Download</a></td>
								</tr>
							<%
								w_iCnt = w_iCnt + 1
								m_Rs.MoveNext
							Loop
							%>

						</table>

					<% end if %>
				</td>
			</tr>
		</table>

		<%=w_pageBar %>
		</td></tr></table>

		</div>
	    <input type="hidden" name="txtMode">
	    <input type="hidden" name="txtPageTyu" value="<%=m_iPageTyu%>">
	    <input type="hidden" name="hidGAKUSEI_NO">

		<%' 検索条件 %>
		</form>
	<% End if %>
	</body>
    </html>


<%
    '---------- HTML END   ----------
End Sub

%>

