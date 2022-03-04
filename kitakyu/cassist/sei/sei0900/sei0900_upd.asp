<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 仮進級者成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0900/sei0900_upd.asp
' 機      能: 下ページ 仮進級者成績登録の登録、更新
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
' 変      数:
' 引      渡:
' 説      明:
'           ■入力データの登録、更新を行う
'-------------------------------------------------------------------------
' 作      成: 2022/2/1 吉田　再試験成績登録画面を流用し作成
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<!--#include file="sei0900_upd_func.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
    Const DebugPrint = 0
    Public Const C_GOUKAKUTEN = 60  '合格点
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_Rs_Hyoka			'評価情報
	
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
    Dim     hidSeiseki
    Dim     m_iRisyuKakoNendo   '//過年度 
    Dim     m_iHaitotani   '//配当単位
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
    w_sMsgTitle="仮進級者成績登録"
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
    m_iRisyuKakoNendo =  request("txtRisyuKakoNendo")
    m_iHaitotani =  request("txtHaitoTani")

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
Dim w_Sisekiarray

    On Error Resume Next
    Err.Clear
	
    f_Update = 99
	w_DataKbnFlg = false
	w_DataKbn = 0
    w_Sisekiarray = split(Trim(request("hidSeiseki")),",")
    ' response.write  "w_Sisekiarray0:" & w_Sisekiarray(0)

    Do 
		w_Today = gf_YYYY_MM_DD(m_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		
		m_SchoolFlg = cbool(request("hidSchoolFlg"))
		
        '// 科目評価取得
        w_iRet = f_GetKamokuTensuHyoka(m_iRisyuKakoNendo,m_sKamokuCd)
        If w_iRet<> 0 Then
            Exit Do
        End If

		For i=1 to i_max : Do
            '成績が否の場合は、次の学生へ
            if w_Sisekiarray(i-1) = 2 then Exit Do
			       
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

			'//T17_RISYUKAKO_KOJINにUPDATE
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " UPDATE T17_RISYUKAKO_KOJIN SET "
			w_sSQL = w_sSQL & vbCrLf & "   T17_SEI_KIMATU_K = " & C_GOUKAKUTEN  & ","
            w_sSQL = w_sSQL & vbCrLf & "   T17_HYOKA_KIMATU_K = '" &  m_Rs_Hyoka("M08_HYOKA_SYOBUNRUI_MEI")  & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T17_GPA_KIMATU_K = " & m_Rs_Hyoka("M08_HYOTEN_GPA")   & ","
            w_sSQL = w_sSQL & vbCrLf & "   T17_TANI_SUMI = " & m_iHaitotani  & ","
            w_sSQL = w_sSQL & vbCrLf & "   T17_KOUSINBI_KIMATU_K = '" & gf_YYYY_MM_DD(date(),"/") & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T17_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
            w_sSQL = w_sSQL & vbCrLf & "   T17_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T17_NENDO = " & Cint(m_iRisyuKakoNendo) & " "
            w_sSQL = w_sSQL & vbCrLf & "    AND T17_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "    AND T17_KAMOKU_CD = '" & Trim(m_sKamokuCd) & "'  "

            If gf_ExecuteSQL(w_sSQL) <> 0 Then
                '//ﾛｰﾙﾊﾞｯｸ
                msMsg = Err.description
                Exit Do
            End If
        ' response.write w_sSQL & "<BR>"
        '   response.end
	    Loop Until 1: Next
		' response.end
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
'*  [機能]  科目評価取得
'*  [引数]  値
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKamokuTensuHyoka(p_iNendo,p_sKamokuCD)

    f_GetKamokuTensuHyoka = 1
    Dim w_iZokuseiCD         '科目属性
    Dim w_iHyokaNo
    Dim w_iRet

    ' 科目属性取得
	w_iRet = f_GetKamokuZokusei(p_iNendo,p_sKamokuCD,w_iZokuseiCD)
    If w_iRet<> 0 Then
            Exit Function
    End If
    
    '科目属性から評価NO取得
    w_iRet = f_iGetHyokaNo(p_iNendo,w_iZokuseiCD,w_iHyokaNo) 
    If w_iRet<> 0 Then
            Exit Function
    End If
    
    '評価NOから評価データ取得
    w_iRet = f_GetTensuHyoka(p_iNendo,w_iZokuseiCD,C_GOUKAKUTEN) 
    If w_iRet<> 0 Then
            Exit Function
    End If

    f_GetKamokuTensuHyoka = 0
End Function

'********************************************************************************
'*  [機能]  科目属性取得(通常時)
'*  [引数]  p_iNendo - 年度(IN)
'        　 p_sKamokuCD - 科目コード(IN)
'           p_iZokuseiCD - 属性コード(OUT)
'*  [戻値]   
'*  [説明]  
'********************************************************************************
Function f_GetKamokuZokusei(p_iNendo,p_sKamokuCD, p_iZokuseiCD)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKamokuZokusei = 1
	p_bZenkiOnly = false

    Do 

'		'//科目属性取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSql = w_sSql & vbCrLf & " M03_ZOKUSEI_CD"
        w_sSql = w_sSql & vbCrLf & " FROM"
        w_sSql = w_sSql & vbCrLf & " M03_KAMOKU"
        w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " M03_NENDO=" & Cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & " AND M03_KAMOKU_CD='" & Trim(m_sKamokuCd) & "'" 

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKamokuZokusei = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_iZokuseiCD = w_Rs("M03_ZOKUSEI_CD")
		End If

        f_GetKamokuZokusei = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  評価形式Noを取得する
'*  [引数]  p_iNendo - 年度(IN)
'           p_iKamokuZokusei_CD - 科目属性コード(IN)
'           p_iHYOKA_NO - 評価形式No(OUT)
'*  [戻値]   
'*  [説明]  
'********************************************************************************
Function f_iGetHyokaNo(p_iNendo,p_iKamokuZokusei_CD,p_iHYOKA_NO)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_iGetHyokaNo = 1
	p_bZenkiOnly = false

    Do 

'		'//評価形式Noを取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSql = w_sSql & vbCrLf & " M100_HYOUKA_NO"
        w_sSql = w_sSql & vbCrLf & " FROM"
        w_sSql = w_sSql & vbCrLf & " M100_KAMOKU_ZOKUSEI"
        w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " M100_NENDO=" & Cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & " AND M100_ZOKUSEI_CD='" & Trim(p_iKamokuZokusei_CD) & "'" 

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_iGetHyokaNo = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			p_iHYOKA_NO = w_Rs("M100_HYOUKA_NO")
		End If

        f_iGetHyokaNo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  科目評価取得(評価=可のデータ)
'*  [引数]  p_iNendo - 年度(IN)
'           p_iKamokuZokusei_CD - 科目属性コード(IN)
'           p_iHYOKA_NO - 評価形式No(OUT)
'*  [戻値]   
'*  [説明]  
'********************************************************************************
Function f_GetTensuHyoka(p_iNendo,p_iHYOKA_NO,p_iTensu)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetTensuHyoka = 1
	p_bZenkiOnly = false

    Do 

'		'//評価形式Noを取得
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSql = w_sSql & vbCrLf & " M08_HYOKA_SYOBUNRUI_MEI,"
        w_sSql = w_sSql & vbCrLf & " M08_HYOTEI,"
        w_sSql = w_sSql & vbCrLf & " M08_HYOKA_SYOBUNRUI_RYAKU,"
        w_sSql = w_sSql & vbCrLf & " M08_HYOTEN_GPA"
        w_sSql = w_sSql & vbCrLf & " FROM"
        w_sSql = w_sSql & vbCrLf & " M08_HYOKAKEISIKI"
        w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " M08_NENDO=" & Cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & " AND M08_HYOUKA_NO='" & p_iHYOKA_NO & "'" 
        w_sSQL = w_sSQL & vbCrLf & " AND M08_MIN <= '" & p_iTensu & "'" 
        w_sSQL = w_sSQL & vbCrLf & " AND M08_MAX >= '" & p_iTensu & "'" 
        w_sSQL = w_sSQL & vbCrLf & " AND M08_HYOKA_TAISYO_KBN ='" & C_HYOKA_TAISHO_IPPAN & "'" 

' response.write w_sSQL
'response.end
        iRet = gf_GetRecordset(m_Rs_Hyoka, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetTensuHyoka = 99
            Exit Do
        End If

        '//正常終了
        f_GetTensuHyoka = 0
        Exit Do
    Loop

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
    <title>仮進級者成績登録</title>
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
	    document.frm.action = "./sei0900_bottom.asp"
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
	<input type=hidden name=txtKamokuCd value="<%=trim(Request("txtKamokuCd"))%>">
    <input type="hidden" name="txtKamokuNM" value="<%=trim(Request("txtKamokuNM"))%>"">
    <input type="hidden" name="txtRisyuKakoNendo" value="<%=m_iRisyuKakoNendo%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>