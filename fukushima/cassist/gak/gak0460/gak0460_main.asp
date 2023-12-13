<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 調査書所見等登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0460/gak0460_main.asp
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
' 変      更：2002/10/08 廣田 耕一郎   担任所見、資格等の項目を追加
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
    Public m_sTanninSyoken  '担任所見
    Public m_sTanninBikou   '資格等
    Public m_sSsyoken       '総合所見
    Public m_sBikou         '個人備考
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
	
    Public  m_GRs
    Public  m_Rs
    Public  m_iMax          '最大ページ
    Public  m_iDsp          '一覧表示行数
	
	Public m_sNendo         '年度コンボボックスに入る値
	Public m_sGakkoNO       '学校番号
	
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
        If gf_OpenDatabase() <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

		Call f_Gakusei()
		
        '//データ取得
        If f_getdate() <> 0 Then m_bErrFlg = True : Exit Do
        
        '//学科ＣＤ取得
        If f_getGakka() <> 0 Then m_bErrFlg = True : Exit Do

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

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
    m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakuNo   = request("txtGakuNo")
    m_iDsp      = C_PAGE_LINE
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm  = request("txtClassNm")
	
	m_sNendo    = request("txtNendo")
	
	'//前へOR次へボタンが押された時
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

'********************************************************************************
'*  [機能]  教官の氏名を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_Gakusei()

Dim i
i = 1

    w_iNyuNendo = Cint(m_sNendo) - Cint(m_sGakunen) + 1
    'w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

	'//学生の情報収集
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & " T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
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
    w_sSQL = w_sSQL & "     B.T13_NENDO = " & m_sNendo & " "
    'w_sSQL = w_sSQL & "     B.T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND B.T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND B.T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & " ORDER BY B.T13_GAKUSEKI_NO "
	
	If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
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

'********************************************************************************
'*  [機能]  データの取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_getdate()

    On Error Resume Next
    Err.Clear
    f_getdate = 1
	
	if Not gf_GetGakkoNO(m_sGakkoNO) then
        m_bErrFlg = True
		exit function
	end if
	
    Do
		
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     T11_SOGOSYOKEN,T11_KOJIN_BIK,T11_SINRO,T11_SOTUKEN_DAI, "
        w_sSQL = w_sSQL & "     T11_SOTU_KYOKAN_CD1,T11_SOTU_KYOKAN_CD2,T11_SOTU_KYOKAN_CD3 "
        
        if m_sGakkoNO = cstr(C_NCT_KUMAMOTO) then
	        w_sSQL = w_sSQL & "    ,T13_TANNINSYOKEN "
        	w_sSQL = w_sSQL & "    ,T13_TANNIN_BIK"
        end if
        
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T11_GAKUSEKI, "
		w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
		
		if m_sGakkoNO = cstr(C_NCT_KUMAMOTO) then
			w_sSQL = w_sSQL & "     AND T13_NENDO = " & m_sNendo
			w_sSQL = w_sSQL & "     AND T13_GAKUSEI_NO = T11_GAKUSEI_NO "
		end if
		
		Set m_Rs = Server.CreateObject("ADODB.Recordset")
        
        If gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp) <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_getdate = 99
            m_bErrFlg = True
            Exit Do 
        End If

        m_sSsyoken  = m_Rs("T11_SOGOSYOKEN")
        m_sBikou    = m_Rs("T11_KOJIN_BIK")
        m_sSinro    = m_Rs("T11_SINRO")
        m_sSotudai  = m_Rs("T11_SOTUKEN_DAI")
        m_sSkyokan1 = m_Rs("T11_SOTU_KYOKAN_CD1")
        m_sSkyokan2 = m_Rs("T11_SOTU_KYOKAN_CD2")
        m_sSkyokan3 = m_Rs("T11_SOTU_KYOKAN_CD3")

		if m_sGakkoNO = cstr(C_NCT_KUMAMOTO) then
			m_sTanninSyoken = m_Rs("T13_TANNINSYOKEN")
	        m_sTanninBikou  = m_Rs("T13_TANNIN_BIK")
		end if

        f_getdate = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  学生の所属学科を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_getGakka()

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
        document.frm.action="gak0460_topDisp.asp";
        document.frm.submit();

	}

    //************************************************************
    //  [機能]  進路先選択画面ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function NewWin(p_iFLG,p_sSNm) {

		var obj=eval("document.frm."+p_sSNm)
        URL = "../../mst/mst0133/default.asp?txtFLG="+p_iFLG+"&txtSNm="+escape(obj.value)+"";
        //URL = "../../mst/mst0133/default.asp?txtFLG="+p_iFLG+"&txtSNm="+p_sSNm+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,toolbar=no,width=560,height=600,top=0,left=0");
        nWin.focus();
        return true;    
    }

    //************************************************************
    //  [機能] クリアボタンを押されたとき
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function jf_Clear(pTextName,pHiddenName){
        eval("document.frm."+pTextName).value = "";
        eval("document.frm."+pHiddenName).value = "";
    }

    //************************************************************
    //  [機能]  卒研教官参照選択画面ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function KyokanWin(p_iInt,p_sKNm) {
		
		var obj=eval("document.frm."+p_sKNm)
        URL = "../../Common/com_select/SEL_KYOKAN/default.asp";
        URL = URL + "?txtI="+p_iInt;
        URL = URL + "&txtKNm="+escape(obj.value);
        URL = URL + "&txtGakka=<%=m_sGakka%>";
        //URL = URL + "&hidNendo=<%=m_sNendo%>";
        
        //URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+p_sKNm+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,toolbar=no,width=550,height=650,top=0,left=0");
        nWin.focus();
        return true;    
    }
    
    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(p_FLG){
	
	if (chk_Flg == false && p_FLG != 0) {f_Button(p_FLG);return false;} //変更がない場合はそのまま次へ

        // ■■■総合所見の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.SGSyoken.value) > "200" ){
            window.alert("総合所見の欄は全角100文字以内で入力してください");
            document.frm.SGSyoken.focus();
            return ;
        }
        // ■■■個人備考の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.Bikou.value) > "80" ){
            window.alert("個人備考の欄は全角40文字以内で入力してください");
            document.frm.Bikou.focus();
            return ;
        }
<%If m_sGakunen = 5 Then%>
        // ■■■卒研論題の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.SRondai.value) > "80" ){
            window.alert("卒研論題の欄は全角40文字以内で入力してください");
            document.frm.SRondai.focus();
            return ;
        }
<%End If%>
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

        document.frm.action="gak0460_upd.asp";
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
        document.frm.action="gak0460_main.asp";
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
<form name="frm" method="post" onClick="return false;">
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
				<% if m_sGakkoNO = cstr(C_NCT_KUMAMOTO) then %>
	                <TR>
	                    <TH CLASS="header" width="120">担任所見</TH>
	                    <TD CLASS="detail"><textarea rows="4" cols="50" class="text" name="TanninSyoken" onChange="chk_Flg=true;"><%=m_sTanninSyoken%></textarea><br>
	                    <font size="2">（全角100文字以内）</font></TD>
	                </TR>
	                <TR>
	                    <TH CLASS="header" width="120">資格等</TH>
	                    <TD CLASS="detail"><textarea rows="2" cols="50" class="text" name="TanninBikou" onChange="chk_Flg=true;"><%=m_sTanninBikou%></textarea><br>
	                    <font size="2">（全角40文字以内）</font></TD>
	                </TR>
				<% end if %>

                <TR>
                    <TH CLASS="header" width="120">総合所見</TH>
                    <TD CLASS="detail"><textarea rows="4" cols="50" class="text" name="SGSyoken" onChange="chk_Flg=true;"><%=m_sSsyoken%></textarea><br>
                    <font size="2">（全角100文字以内）</font></TD>
                </TR>
                <TR>
                    <TH CLASS="header" width="120">備　考</TH>
                    <TD CLASS="detail"><textarea rows="2" cols="50" class="text" name="Bikou" onChange="chk_Flg=true;"><%=m_sBikou%></textarea><br>
                    <font size="2">（全角40文字以内）</font></TD>
                </TR>
<%If m_sGakunen = 5 Then%>
                <!--TR>
                    <TH CLASS="header" width="120">卒業後の進路</TH>
                    <TD CLASS="detail">
                    <input type="text" class="text" name="SinroNm" VALUE='<%Call f_SINRO(m_sSinro,m_iNendo)%>' size="50" readonly style="width:260px;" onChange="chk_Flg=true;">
                    <input type="hidden" name="SinroCd" VALUE='<%=m_sSinro%>'>
                    <input type="button" class="button" value="選択" onclick="NewWin(1,'SinroNm')">
                    <input type="button" class="button" value="クリア" onclick="jf_Clear('SinroNm','SinroCd')">
                </TR-->
                <TR>
                    <TH CLASS="header" width="120">卒研論題</TH>
                    <TD CLASS="detail"><textarea rows="2" cols="50" class="text" name="SRondai" onChange="chk_Flg=true;"><%=m_sSotudai%></textarea><br>
                    <font size="2">（全角40文字以内）</font></TD>
                </TR>
                <TR>
                    <TH CLASS="header" nowrap width="120">卒研教官1</TH>
                    <TD CLASS="detail">
                    <!--input type="text" class="text" name="SKyokanNm1" VALUE='<%Call f_KYO_MEI(m_sSkyokan1,m_iNendo)%>' size="24" readonly onChange="chk_Flg=true;"-->
                    <input type="text" class="text" name="SKyokanNm1" VALUE='<%Call f_KYO_MEI(m_sSkyokan1,m_sNendo)%>' size="24" readonly onChange="chk_Flg=true;">
                    <input type="hidden" name="SKyokanCd1" VALUE='<%=m_sSkyokan1%>'>
                    <input type="button" class="button" value="選択" onclick="KyokanWin(1,'SKyokanNm1')">
                    <input type="button" class="button" value="クリア" onclick="jf_Clear('SKyokanNm1','SKyokanCd1')"></td>
                </TR>
                <TR>
                    <TH CLASS="header" nowrap width="120">卒研教官2</TH>
                    <TD CLASS="detail">
                    <!--input type="text" class="text" name="SKyokanNm2" VALUE='<%Call f_KYO_MEI(m_sSkyokan2,m_iNendo)%>' size="24" readonly onChange="chk_Flg=true;"-->
                    <input type="text" class="text" name="SKyokanNm2" VALUE='<%Call f_KYO_MEI(m_sSkyokan2,m_sNendo)%>' size="24" readonly onChange="chk_Flg=true;">
                    <input type="hidden" name="SKyokanCd2" VALUE='<%=m_sSkyokan2%>'>
                    <input type="button" class="button" value="選択" onclick="KyokanWin(2,'SKyokanNm2')">
                    <input type="button" class="button" value="クリア" onclick="jf_Clear('SKyokanNm2','SKyokanCd2')"></td>
                </TR>
                <TR>
                    <TH CLASS="header" nowrap width="120">卒研教官3</TH>
                    <TD CLASS="detail">
                    <!--input type="text" class="text" name="SKyokanNm3" VALUE='<%Call f_KYO_MEI(m_sSkyokan3,m_iNendo)%>' size="24" readonly onChange="chk_Flg=true;"-->
                    <input type="text" class="text" name="SKyokanNm3" VALUE='<%Call f_KYO_MEI(m_sSkyokan3,m_sNendo)%>' size="24" readonly onChange="chk_Flg=true;">
                    <input type="hidden" name="SKyokanCd3" VALUE='<%=m_sSkyokan3%>'>
                    <input type="button" class="button" value="選択" onclick="KyokanWin(3,'SKyokanNm3')">
                    <input type="button" class="button" value="クリア" onclick="jf_Clear('SKyokanNm3','SKyokanCd3')"></td>
                </TR>
<%End If%>
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
	<input type="hidden" name="txtNendo" value="<%=m_sNendo%>">
	<!--input type="hidden" name="txtNendo" value="<%=m_iNendo%>"-->
	<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="txtBeforGakuNo" value="<%=m_sBeforGakuNo%>">
	<input type="hidden" name="txtAfterGakuNo" value="<%=m_sAfterGakuNo%>">
	<input type="hidden" name="GakuseiNo" value="">
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
