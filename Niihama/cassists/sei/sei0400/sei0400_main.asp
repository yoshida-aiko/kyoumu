<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 調査書所見等登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/sei0400/sei0400_main.asp
' 機      能: 下ページ 調査書所見等登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/18 前田 智史
' 変      更: 2001/08/09 根本 直美     NN対応に伴うソース変更
' 変      更：2001/08/30 伊藤 公子     検索条件を2重に表示しないように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系

    '市町村選択用のWhere条件
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官コード
    Public m_sGakuNo        '氏名コンボボックスに入る値
    Public m_sBeforGakuNo   '氏名コンボボックスに入る値の一人前
    Public m_sAfterGakuNo   '氏名コンボボックスに入る値の一人後
    Public m_sSsyoken       '総合所見
    Public m_sBikou         '個人備考
    Public m_sSyoken
    Public m_sSinro         '進路名
    Public m_sSotudai       '卒研課題
    Public m_sSkyokan1      '卒官1
    Public m_sSkyokan2      '卒官2
    Public m_sSkyokan3      '卒官3
    Public m_sGakunen       '学年
    Public m_sClass         'クラス
    Public m_sClassNm       'クラス名
    Public m_sGakusei()     '学生の配列
    Public m_sGakka     '学生の所属学科
    Public m_iSikenKBN
    
    Public  m_GRs,m_DRs
    Public  m_Rs
    Public  m_iMax          '最大ページ
    Public  m_iDsp          '一覧表示行数

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
    w_sMsgTitle="調査書所見等登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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


'2001/12/05 ITO 修正 入力期限は設定しなくてよい
		'===============================
		'//期間データの取得
		'===============================
        'w_iRet = f_Nyuryokudate()
		'If w_iRet = 1 Then
			'// ページを表示
		'	Call No_showPage("成績入力期間外です。")
		'	Exit Do
		'End If
		'If w_iRet <> 0 Then 
		'	m_bErrFlg = True
		'	Exit Do
		'End If


		Call f_Gakusei()

        '//データ取得
        w_iRet = f_getdate()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
         '//学科ＣＤ取得
        w_iRet = f_getGakka()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

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
    Call gs_CloseDatabase()
End Sub

Sub s_SetParam()
'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakuNo   = request("txtGakuNo")
    m_iDsp      = C_PAGE_LINE
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm    = request("txtClassNm")
	m_iSikenKBN    = request("txtSikenKBN")
	'//前へOR次へボタンが押された時
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

Function f_Nyuryokudate()
'********************************************************************************
'*	[機能]	データの取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
	dim w_date

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1


	w_date = gf_YYYY_MM_DD(date(),"/")
'	w_date = "2000/06/18"

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
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_iSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & w_date & "' "

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
'//成績入力期間テスト用

'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '2002/04/30'"
'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '2000/03/01'"

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

Function f_Gakusei()
'********************************************************************************
'*  [機能]  教官の氏名を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim i
i = 1


    w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

	'//学生の情報収集
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "      T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
'    w_sSQL = w_sSQL & " AND T11_NYUNENDO = " & w_iNyuNendo & " "

    Set m_GRs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_GRs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
    End If


    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     A.T11_GAKUSEI_NO "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI A,T13_GAKU_NEN B "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     B.T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND B.T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND B.T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & " ORDER BY B.T13_GAKUSEKI_NO "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
    End If
	w_rCnt=cint(gf_GetRsCount(w_Rs))

	'//配列の作成

		w_Rs.MoveFirst

       Do Until w_Rs.EOF

            ReDim Preserve m_sGakusei(i)
            m_sGakusei(i) = w_Rs("T11_GAKUSEI_NO")
            i = i + 1
            
            w_Rs.MoveNext
            
        Loop

		For i = 1 to w_rCnt

			If m_sGakusei(i) = m_sGakuNo Then

				If i <= 1 Then
					m_sGakuNo      = m_sGakusei(i)
	                m_sAfterGakuNo = m_sGakusei(i+1)
					Exit For
				End If

				If i = w_rCnt Then
					m_sGakuNo      = m_sGakusei(i)
	                m_sBeforGakuNo = m_sGakusei(i-1)
					Exit For
				End If

				m_sGakuNo      = m_sGakusei(i)
                m_sAfterGakuNo = m_sGakusei(i+1)
                m_sBeforGakuNo = m_sGakusei(i-1)
				
				Exit For
			End If

		Next

End Function


Function f_KYO_MEI(p_sCD,p_iNENDO)
'********************************************************************************
'*  [機能]  教官の氏名を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_Rs

    If Isnull(p_sCD) Then 
        f_KYO_MEI = "" 
        Exit Function
    End If

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     M04_KYOKAN "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     M04_KYOKAN_CD = '" & p_sCD & "' "
    w_sSQL = w_sSQL & " AND M04_NENDO = " & p_iNENDO & " "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(w_Rs, w_sSQL, m_iDsp)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
    End If

    'f_KYO_MEI = w_Rs("M04_KYOKANMEI_SEI")&"　"&w_Rs("M04_KYOKANMEI_MEI")
    response.write w_Rs("M04_KYOKANMEI_SEI")&"　"&w_Rs("M04_KYOKANMEI_MEI")

End Function

Function f_SINRO(p_sCD,p_iNENDO)
'********************************************************************************
'*  [機能]  進路先を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_Rs

    If Isnull(p_sCD) Then 
        f_SINRO = "" 
        Exit Function
    End If

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     M32_SINROMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     M32_SINRO "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     M32_SINRO_CD = '" & p_sCD & "' "
    w_sSQL = w_sSQL & " AND M32_NENDO = " & p_iNENDO & " "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(w_Rs, w_sSQL, m_iDsp)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
    End If

    'f_SINRO = w_Rs("M32_SINROMEI")
    response.write w_Rs("M32_SINROMEI")

End Function

Function f_getdate()
'********************************************************************************
'*  [機能]  データの取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
	dim w_sSikenKBN

    On Error Resume Next
    Err.Clear
    f_getdate = 1

	select case cint(m_iSikenKBN)
		case C_SIKEN_ZEN_TYU
			w_sSikenKBN = "T13_SYOKEN_TYUKAN_Z"
		case C_SIKEN_ZEN_KIM
			w_sSikenKBN = "T13_SYOKEN_KIMATU_Z"
		case C_SIKEN_KOU_TYU
			w_sSikenKBN = "T13_SYOKEN_TYUKAN_K"
		case C_SIKEN_KOU_KIM
			w_sSikenKBN = "T13_SYOKEN_KIMATU_K"
	End select

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     " & w_sSikenKBN & " as Shoken "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo & " AND "
        w_sSQL = w_sSQL & "     T13_GAKUSEI_NO = '" & m_sGakuNo & "' "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_getdate = 99
            m_bErrFlg = True
            Exit Do 
        End If
        m_sSyoken  = m_Rs("Shoken")
        f_getdate = 0
        Exit Do
    Loop

End Function

Function f_getGakka()
'********************************************************************************
'*  [機能]  学生の所属学科を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_getGakka = 1

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     T13_GAKKA_CD"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_GAKUSEI_NO = '" & m_sGakuNo & "' "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_getGakka = 99
            m_bErrFlg = True
            Exit Do 
        End If

	m_sGakka = m_Rs("T13_GAKKA_CD")
        f_getGakka = 0
        Exit Do
    Loop

End Function

Function f_getGakuseki_No()
'********************************************************************************
'*  [機能]  学生の学籍NOを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	Dim rs
	Dim w_sSQL

    On Error Resume Next
    Err.Clear

    f_getGakuseki_No = ""

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     T13_GAKUSEKI_NO"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & "     AND T13_GAKUSEI_NO = '" & m_sGakuNo & "' "

        w_iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            Exit Do 
        End If

		If rs.EOF = False Then
			w_iGakusekiNo = rs("T13_GAKUSEKI_NO")
		End If

        Exit Do
    Loop

	'//戻り値セット
    f_getGakuseki_No = w_iGakusekiNo

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

Sub No_showPage(p_msg)
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
	</head>

	<body>
	<center>
	<br><br><br>
			<span class="msg"><%=p_msg%></span>
	</center>
	</body>

	</html>

<%
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    On Error Resume Next
    Err.Clear

%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type="text/css">

<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="JavaScript">
<!--
	var chk_Flg;
	chk_Flg = false;
	//************************************************************
	//  [機能]  ページロード時処理
	//  [引数]
	//  [戻値]
	//  [説明]
	//************************************************************
	function window_onload() {

        document.frm.target="topFrame";
        document.frm.action="sei0400_topDisp.asp";
        document.frm.submit();

	}

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(p_FLG){

        // ■■■担任所見の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.Syoken.value) > "200" ){
            window.alert("担任所見の欄は全角100文字以内で入力してください");
            document.frm.Syoken.focus();
            return ;
        }
        
	if (chk_Flg == false && p_FLG != 0) {f_Button(p_FLG);return false;} //変更がない場合はそのまま次へ

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

        document.frm.action="sei0400_upd.asp";
        document.frm.target="main";
		if( p_FLG == 1){
			document.frm.GakuseiNo.value = document.frm.txtBeforGakuNo.value;
		}
		if( p_FLG == 2){
        	document.frm.GakuseiNo.value = document.frm.txtAfterGakuNo.value;
        }
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Cansel(){

        //document.frm.action="default2.asp";
        //document.frm.target="main";
        document.frm.action="default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  前へ,次へボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Button(p_FLG){

        //document.frm.action="default.asp";
        document.frm.action="sei0400_main.asp";
        document.frm.target="main";

		if( p_FLG == 1){
			document.frm.GakuseiNo.value = document.frm.txtBeforGakuNo.value;
		}else{
        	document.frm.GakuseiNo.value = document.frm.txtAfterGakuNo.value;
        }
		document.frm.submit();
    
    }

//-->
</SCRIPT>

</head>
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">
<center>

<br>
<table border="0" width="250">
    <tr>
<%If m_sBeforGakuNo <> "" Then%>
        <td valign="top" align="center">
            <input type="button" value="　前　へ　" class="button" onclick="javascript:f_Touroku(1)">
        </td>
<%Else%>
        <td valign="top" align="center">
            <input type="button" value="　前　へ　" class="button" DISABLED>
        </td>
<%End If%>
        <td valign="top" align="center">
            <input type="button" value="　登　録　" class="button" onclick="javascript:f_Touroku(0)">
        </td>
        <td valign="top" align="center">
            <input type="button" value="キャンセル" class="button" onclick="javascript:f_Cansel()">
        </td>
<%If m_sAfterGakuNo <> "" Then%>
        <td valign="top" align="center">
            <input type="button" value="　次　へ　" class="button" onclick="javascript:f_Touroku(2)">
        </td>
<%Else%>
        <td valign="top" align="center">
            <input type="button" value="　次　へ　" class="button" DISABLED>
        </td>
<%End If%>
    </tr>
</table>
<br>
<table border="0" cellpadding="1" cellspacing="1" width="520" >
    <tr>
        <td align="left">
            <table width="500" border=1 CLASS="hyo">
                <TR>
                    <TH CLASS="header" width="120">担任所見</TH>
                    <TD CLASS="detail"><textarea rows="4" cols="50" class="text" name="Syoken" onChange="chk_Flg=true;"><%=m_sSyoken%></textarea><br>
                    <font size="2">（全角100文字以内）</font></TD>
                </TR>
            </TABLE>
        </td>
    </TR>
</TABLE>

<br>

<table border="0" width="250">
    <tr>
<%If m_sBeforGakuNo <> "" Then%>
        <td valign="top" align="center">
            <input type="button" value="　前　へ　" class="button" onclick="javascript:f_Touroku(1)">
        </td>
<%Else%>
        <td valign="top" align="center">
            <input type="button" value="　前　へ　" class="button" DISABLED>
        </td>
<%End If%>
        <td valign="top" align="center">
            <input type="button" value="　登　録　" class="button" onclick="javascript:f_Touroku(0)">
        </td>
        <td valign="top" align="center">
            <input type="button" value="キャンセル" class="button" onclick="javascript:f_Cansel()">
        </td>
<%If m_sAfterGakuNo <> "" Then%>
        <td valign="top" align="center">
            <input type="button" value="　次　へ　" class="button" onclick="javascript:f_Touroku(2)">
        </td>
<%Else%>
        <td valign="top" align="center">
            <input type="button" value="　次　へ　" class="button" DISABLED>
        </td>
<%End If%>
    </tr>
</table>
	<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">
	<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="txtBeforGakuNo" value="<%=m_sBeforGakuNo%>">
	<input type="hidden" name="txtAfterGakuNo" value="<%=m_sAfterGakuNo%>">
	<input type="hidden" name="GakuseiNo" value="">
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
	<input type="hidden" name="txtSikenKBN" value="<%=m_iSikenKBN%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
