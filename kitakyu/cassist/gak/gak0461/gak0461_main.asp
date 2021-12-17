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
' 変      更：2001/08/30 伊藤 公子     検索条件を2重に表示しないように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '氏名選択用のWhere条件
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官コード
    Public m_sNendo         '年度コンボボックスに入る値
    Public m_sGakuNo        '氏名コンボボックスに入る値
    Public m_sBeforGakuNo   '氏名コンボボックスに入る値の一人前
    Public m_sAfterGakuNo   '氏名コンボボックスに入る値の一人後
    Public m_sGakunen       '学年
    Public m_sClass         'クラス
    Public m_sClassNm       'クラス名
    Public m_sGakusei()     '学生の配列

    Public  m_TRs           
    Public  m_GRs           
    Public  m_URs
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

        '//データ取得
        w_iRet = f_Gakusei()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		If m_GRs.EOF Then
			Call NO_Showpage()
			Exit Do
		End If

        '//データ取得
        w_iRet = f_getdate()
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

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_sNendo    = request("txtNendo")
    m_sGakuNo   = request("txtGakuNo")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm  = request("txtClassNm")
    m_iDsp      = C_PAGE_LINE

	'//前へOR次へボタンが押された時
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

Function f_Gakusei()
'********************************************************************************
'*  [機能]  教官の氏名を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim i
i = 1


    w_iNyuNendo = Cint(m_sNendo) - Cint(m_sGakunen) + 1

	'//学生の情報収集
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
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
    w_sSQL = w_sSQL & " AND B.T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND B.T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
'    w_sSQL = w_sSQL & " AND A.T11_NYUNENDO = B.T13_NENDO - B.T13_GAKUNEN + 1"
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

Function f_getdate()
'********************************************************************************
'*  [機能]  データの取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_getdate = 1

    Do
        '//行動所見,趣味･特技･取得資格等,個人備考のデータ取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
'        w_sSQL = w_sSQL & "     T11_KODOSYOKEN,T11_SYUMITOKUGI,T11_TYOSA_BIK "
        w_sSQL = w_sSQL & "     T11_TYOSA_BIK "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T11_GAKUSEKI "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "

        Set m_TRs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_TRs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_getdate = 99
            m_bErrFlg = True
            Exit Do 
        End If

        '//特別活動の記録,年毎所見のデータ取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     T13_TOKUKATU_DET,T13_NENSYOKEN,T13_NENSYOKEN2,T13_NENSYOKEN3 "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_GAKUSEI_NO = '" & m_sGakuNo & "' "
        w_sSQL = w_sSQL & " AND T13_NENDO = " & m_sNendo & " "
        Set m_URs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_URs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_getdate = 99
            m_bErrFlg = True
            Exit Do 
        End If

        f_getdate = 0
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

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
        w_sSQL = w_sSQL & "     T13_NENDO = " & m_sNendo
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
    <title>調査書所見等登録</title>
<link rel=stylesheet href="../../common/style.css" type=text/css>

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
        document.frm.action="gak0461_topDisp.asp";
        document.frm.submit();

	}

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(p_FLG){


        // ■■■行動所見の桁ﾁｪｯｸ■■■
//        if( getLengthB(document.frm.KSyoken.value) > "100" ){
//            window.alert("行動所見の欄は全角50文字以内で入力してください");
//            document.frm.KSyoken.focus();
//            return ;
//        }
//        // ■■■趣味等の桁ﾁｪｯｸ■■■
//        if( getLengthB(document.frm.SyumiTokugi.value) > "200" ){
//            window.alert("趣味等の欄は全角100文字以内で入力してください");
//            document.frm.SyumiTokugi.focus();
//            return ;
//        }
        // ■■■指導上参考となる諸事項の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.NSyoken.value) > "100" ){
            window.alert("指導上参考となる諸事項の欄は全角50文字以内で入力してください");
            document.frm.NSyoken.focus();
            return ;
        }
        // ■■■指導上参考となる諸事項の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.NSyoken2.value) > "100" ){
            window.alert("指導上参考となる諸事項の欄は全角50文字以内で入力してください");
            document.frm.NSyoken2.focus();
            return ;
        }
        // ■■■指導上参考となる諸事項の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.NSyoken3.value) > "100" ){
            window.alert("指導上参考となる諸事項の欄は全角50文字以内で入力してください");
            document.frm.NSyoken3.focus();
            return ;
        }
        // ■■■特別活動の桁ﾁｪｯｸ■■■
        if( getLengthB(document.frm.Tokukatu.value) > "100" ){
            window.alert("特別活動の欄は全角50文字以内で入力してください");
            document.frm.Tokukatu.focus();
            return ;
        }
        // ■■■備考の桁ﾁｪｯｸ■■■
//        if( getLengthB(document.frm.Bikou.value) > "200" ){
//            window.alert("備考の欄は全角100文字以内で入力してください");
//            document.frm.Bikou.focus();
//            return ;
//        }

	if (chk_Flg == false && p_FLG != 0) {f_Button(p_FLG);return false;} //変更がない場合はそのまま次へ

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        document.frm.action="gak0461_upd.asp";
        document.frm.target="main";
        //document.frm.target="<%=C_MAIN_FRAME%>";
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
        document.frm.action="gak0461_main.asp";
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
<table border="0" cellpadding="1" cellspacing="1" width="520">
    <tr>
        <td align="left">
            <table width="100%" border=1 CLASS="hyo">
<%
'--------------　行動所見　趣味特技　削除　--------------
'                <TR>
'                    <TH CLASS="header" width="120" nowrap>行動所見</TH>
'                    <TD CLASS="detail"><textarea rows=2 cols=50 class=text name="KSyoken"><%=m_TRs("T11_KODOSYOKEN")%***></textarea><br>
'                    <font size=2>（全角50文字以内）</font></TD>
'                </TR>
'                <TR>
'                    <TH CLASS="header" width="120" nowrap>趣味･特技<BR>取得資格等</TH>
'                    <TD CLASS="detail"><textarea rows=4 cols=50 class=text name="SyumiTokugi"><%='m_TRs("T11_SYUMITOKUGI")%****></textarea><br>
'                    <font size=2>（全角100文字以内）</font></TD>
'                </TR>
%>
                <TR>
                    <TH CLASS="header" width="120" nowrap>指導上参考<BR>となる諸事項</TH>
                    <TD CLASS="detail">
                        <table>
                            <TR align="center">
                                <TH CLASS="header" width="120" align="left" nowrap>(1)学習における特徴等<BR>(2)行動の特徴、特技等</TH>
                            </TR>
                            <TR>
                                <TD CLASS="detail">
<!--2015/03/18 UPDATE URAKAWA-->
<!--<textarea rows=2 cols=50 class=text name="NSyoken" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN")%></textarea><br>-->
<textarea rows=4 cols=50 class=text name="NSyoken" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN")%></textarea><br>
                    <font size=2>（全角50文字以内）</font>
                                </TD>
                            </TR>
                            <TR>
                                <TH CLASS="header" width="120" align="left" nowrap>(3)部活動、ボランティア活動等<BR>(4)取得資格、検定等</TH>
                            </TR>
                            <TR>
                                <TD CLASS="detail">
<!--2015/03/18 UPDATE URAKAWA-->
<!--<textarea rows=2 cols=50 class=text name="NSyoken2" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN2")%></textarea><br>-->
<textarea rows=4 cols=50 class=text name="NSyoken2" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN2")%></textarea><br>
                    <font size=2>（全角50文字以内）</font>
                                </TD>
                            </TR>
                            <TR>
                                <TH CLASS="header" width="120" align="left" nowrap>(5)その他</TH>
                            </TR>
                            <TR>
                                <TD CLASS="detail">
<!--2015/03/18 UPDATE URAKAWA-->
<!--<textarea rows=2 cols=50 class=text name="NSyoken3" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN3")%></textarea><br>-->
<textarea rows=4 cols=50 class=text name="NSyoken3" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN3")%></textarea><br>
                    <font size=2>（全角50文字以内）</font>
                                </TD>
                            </TR>
                        </table>
                    </TD>
                </TR>
                <TR>
                    <TH CLASS="header" width="120" nowrap>特別活動<BR>の記録</TH>
<!--2015/03/18 UPDATE URAKAWA-->
<!--                    <TD CLASS="detail"><textarea rows=2 cols=50 class=text name="Tokukatu" onChange="chk_Flg=true;"><%=m_URs("T13_TOKUKATU_DET")%></textarea><br> -->
                    <TD CLASS="detail"><textarea rows=8 cols=50 class=text name="Tokukatu" onChange="chk_Flg=true;"><%=m_URs("T13_TOKUKATU_DET")%></textarea><br>
                    <font size=2>（全角50文字以内）</font></TD>
                </TR>

<!--2015/03/18 DELETE URAKAWA-->
<!--            <TR>
                    <TH CLASS="header" width="120" nowrap>備　考</TH>
                    <TD CLASS="detail"><textarea rows=4 cols=50 class=text name="Bikou" onChange="chk_Flg=true;"><%=m_TRs("T11_TYOSA_BIK")%></textarea><br>
                    <font size=2>（全角100文字以内）</font></TD>
                </TR>
-->
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
Sub NO_Showpage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
    <html>
    <head>
    <title>調査書所見等登録</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <script language="javascript">
    </script>
    </head>
    <body>
    <center>
<br><br><br><br><br>
        <span class="msg">選択された学生の調査書所見等登録のデータがありません。</span>


    </center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub
%>
