<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: レベル別科目決定
' ﾌﾟﾛｸﾞﾗﾑID : web/web0390/web0390_main.asp
' 機      能: 下ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/10/26 谷脇　良也
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_iMax          ':最大ページ
    Public  m_iDsp          '// 一覧表示行数
    Public  m_sPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_Grs           '学生用レコードセット
    Public  m_KSrs          '科目数のレコードセット
    Public  m_GrsCntMax     '学生用レコード数
'    Public  m_rs            'レコードセット
    Dim     m_iNendo        '//年度
    Dim     m_sKyokanCd     '//教官コード
    Dim     m_sGakunen      '//学年
    Dim     m_sClass        '//クラス
	Dim		m_sKamokuCD		'//科目コード
    Dim     m_GrCnt         '//学生のレコードカウント
    Dim     m_cell          '配色の設定
	Dim 	m_sKengen		'//権限
    Dim     m_iSTani        
	Dim		m_sRisyuJotai	'履修状態フラグ add 2001/10/25
	Dim 	m_sLKyokan()	'選択されたレベル別科目の担当教官
	Dim 	m_iLKyokanCnt()	'担当教官を選んでいる人の数
    Dim     i               
    Dim     j               

    'エラー系
    Public  m_bErrFlg       'ｴﾗｰﾌﾗｸﾞ
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

    Dim	 w_iRet              '// 戻り値
    Dim  w_Krs           '科目用レコードセット
    Dim  w_KrCnt         '//科目のレコードカウント

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="レベル別科目登録"
    w_sMsg=""
    w_sRetURL=C_RetURL & C_ERR_RETURL
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
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()


		'//権限を取得
		w_iRet = gf_GetKengen_web0390(m_sKengen)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'権限の定数
		'C_WEB0390_ACCESS_FULL  
		'C_WEB0390_ACCESS_SENMON
		'C_WEB0390_ACCESS_TANNIN

		'履修状態区分を取得(履修が決定してるかどうか）
		'C_K_RIS_MAE = 0        '確定処理前
		'C_K_RIS_ATO = 1        '確定処理後
		if f_GetKanriM(m_iNendo,C_K_RIS_JOUTAI,m_sRisyuJotai) <> 0 then 
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
	        m_bErrFlg = True
	        Call w_sMsg("管理マスタの履修状態区分がありません。")
	        Exit Do
		end if

'-----------------------------------------------------
'm_sRisyuJotai = "1" 'test用
'-----------------------------------------------------

        '//教官の情報取得
        w_iRet = f_KyokanData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

		If Ubound(m_sLKyokan) = 0 Then
			Call showPage_NoData()
	        Exit Do
		End If

        '//学生の情報取得
        w_iRet = f_GakuseiData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

        '// ページを表示
        Call showPage()

        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_Grs)
    '// 終了処理
    Call gs_CloseDatabase()
    
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
response.write w_sMsg
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sGakunen  = request("txtGakunen")
    m_sClass    = request("txtClass")
    m_iDsp      = C_PAGE_LINE
	m_sKamokuCD = request("cboKamokuCode")

End Sub

Function f_KyokanData()
'******************************************************************
'機　　能：教官のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear
    f_KyokanData = 1
	i = 0
	m_bErrFlg = false

    Do

        '//科目のデータ取得
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT "
        m_sSQL = m_sSQL & vbCrLf & "     T27_KYOKAN_CD"
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T27_TANTO_KYOKAN T27"
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T27_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T27_GAKUNEN = " & m_sGakunen & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T27_KAMOKU_CD = '" & m_sKamokuCD & "' "
        m_sSQL = m_sSQL & vbCrLf & " GROUP BY T27_KYOKAN_CD"
        m_sSQL = m_sSQL & vbCrLf & " ORDER BY T27_KYOKAN_CD"

'response.write m_sSQL & "<BR>"
'response.end
        Set w_Krs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset_OpenStatic(w_Krs, m_sSQL)

'response.write "w_iRet = " & w_iRet & "<BR>"
'response.write w_Krs.EOF & "<BR>"

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
			w_sMsg = "教官データを取得できません。"
            Exit Do 
        End If

    w_KrCnt=cint(gf_GetRsCount(w_Krs))
'response.write w_KrCnt

	Redim m_sLKyokan(w_KrCnt)
	Redim m_iLKyokanCnt(w_KrCnt)

	w_Krs.MoveFirst
	Do Until w_Krs.EOF
		m_sLKyokan(i) = w_Krs("T27_KYOKAN_CD")
		m_iLKyokanCnt(i) = 0

	if m_sKengen = C_WEB0390_ACCESS_SENMON then
		If m_sKyokanCd = w_Krs("T27_KYOKAN_CD") then
			m_sMain = i
		End If
	End If

		i = i + 1
		w_Krs.MoveNext
	Loop
'response.end
    f_KyokanData = 0

    Exit Do

    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(w_Krs)

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If

End Function

Function f_GakuseiData()
'******************************************************************
'機　　能：学生のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
	Dim w_sSQL

    On Error Resume Next
    Err.Clear
    f_GakuseiData = 1

    Do
        '//学生のデータ取得
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_LEVEL_KYOUKAN AS L_KYOKAN, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO AS GAKUSEI,"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO AS GAKUSEKI,"
        w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI AS SIMEI"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN,T13_GAKU_NEN,T11_GAKUSEKI"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_NENDO = "          & cInt(m_iNendo) 		& " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_HAITOGAKUNEN = "   & cInt(m_sGakunen)  	& " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_KAMOKU_CD = '"     & m_sKamokuCd       	& "' AND "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLASS = "     	    &  m_sClass       		& " AND"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_ZAISEKI_KBN < "   	& C_ZAI_SOTUGYO		  	& " AND"
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_GAKUSEI_NO = T13_GAKU_NEN.T13_GAKUSEI_NO AND"
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_HAITOGAKUNEN = T13_GAKU_NEN.T13_GAKUNEN AND"
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_NENDO = T13_GAKU_NEN.T13_NENDO AND"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO "
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI "

'response.write m_sSQL & "<BR>"

        Set m_Grs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset_OpenStatic(m_Grs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
    m_GrCnt=gf_GetRsCount(m_Grs)

    f_GakuseiData = 0

    Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
'    If m_bErrFlg = True Then
'        w_sMsg = gf_GetErrMsg()
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
'		response.end
'    End If

End Function

'********************************************************************************
'*  [機能]  管理マスタよりデータを取得
'*  [引数]  p_iNendo	年度
'*  　　　  p_iNo		処理番号
'*  [戻値]  p_iKanri	管理データ
'*  [説明]  管理マスタよりデータを取得する。
'********************************************************************************
Function f_GetKanriM(p_iNendo,p_iNo,p_sKanri)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKanriM = 0
    p_sKanri = ""

    Do 

		'//管理マスタより履修状態区分を取得
		'//履修状態区分(C_K_RIS_JOUTAI = 28)
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_NENDO=" & cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M00_KANRI.M00_NO=" & C_K_RIS_JOUTAI	'履修状態区分(=28)

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_GetKanriM = iRet
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			'//Public Const C_K_RIS_MAE = 0    '決定前
			'//Public Const C_K_RIS_ATO = 1    '決定後
			p_sKanri = w_Rs("M00_KANRI")
		End If

        f_GetKanriM = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

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
	<SCRIPT language="javascript">
	<!--
    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		parent.location.href = "white.asp?txtMsg=レベル別科目のデータがありません。"
        return;
    }
	//-->
	</SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>
    </center>
	<form name="frm" method="post">

	<input type="hidden" name="txtMsg" value="レベル別科目のデータがありません。">

	</form>
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
Dim n
Dim w_sChg

    On Error Resume Next
    Err.Clear

i = 0
n = 0
%>
<HTML>


<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>レベル別科目決定</title>

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

        //ヘッダ部submit
        document.frm.target = "middle";
        document.frm.action = "web0390_middle.asp";
        document.frm.submit();
        return;

    }

    //************************************************************
    //  [機能]  ボタンのVALUEの変更
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Chenge(p_iS,p_iK){
		var w_sBtn;
		var w_sTmp;
		var w_sBNm ;		//押されたボタンの列の名前
		var w_sKYOCntNm;	//middleの選択学生数のフォームの名前
		var w_sKYOCnt;		//middleの選択学生数のフォーム
		var w_sKYONm; 		//選択した教官の教官ＣＤ
		var w_sLKYONm;		//学生の選択した教官の教官ＣＤいれるところ
		var w_sLKYO_OLDNm;	//学生の選択していた教官の教官ＣＤいれるところ
		
        w_sBNm = "document.frm.K"+p_iS+"_";
		w_sLKYO_OLDNm = eval("document.frm.L_KYOKAN_OLD"+p_iS);
        w_sKYOCntNm = "parent.middle.document.frm.KYOKAN";
		w_sKYONm = eval("document.frm.KyokanuCd"+p_iK);
		w_sLKYONm = eval("document.frm.L_KYOKAN"+p_iS);
		//教官数
		w_sCnt = <%=UBound(m_sLKyokan)%>;
		w_sBtn = eval(w_sBNm+p_iK);

		//今まで選択していたもののカウントを減らす(選択していた場合）
		if (w_sLKYO_OLDNm.value != 999) {
			w_sKYOCnt = eval(w_sKYOCntNm + w_sLKYO_OLDNm.value);
			w_sKYOCnt.value = parseInt(w_sKYOCnt.value) - 1;
		}
		//選択したものを取り消すとき
		if (w_sBtn.value == "○") {
			w_sBtn.value = "　";
			w_sLKYONm.value = "";
			eval(w_sLKYO_OLDNm).value = 999;
		} else {

		//選択した時
			//一旦、全ての○を削除
			for ( i=0;i<=w_sCnt-1;i++) {

				w_sTmp = eval(w_sBNm+i)
				w_sTmp.value = "　";
			}

			//選択したものに丸をつける
			w_sBtn.value = "○";

			//今、選択していたもののカウントを増やす
			w_sKYOCnt = eval(w_sKYOCntNm + p_iK);
			w_sKYOCnt.value = parseInt(w_sKYOCnt.value) + 1;

			//選択した教官のコードをいれる
			eval(w_sLKYO_OLDNm).value = p_iK;
			w_sLKYONm.value = w_sKYONm.value;
		}
        return;
    }
    
    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Cansel(){
        //空白ページを表示
        parent.document.location.href="default2.asp";

    
    }
    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        document.frm.action="web0390_upd.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>

	<center>

	<body onload="return window_onload()">
	<FORM NAME="frm" method="post">

	    <%
		'//隠しフィールドに科目CDと各科目の単位数を格納(登録時に使用する)

        Do Until n > UBound(m_sLKyokan)
		    %>
	        <input type=hidden name=KyokanuCd<%=n%> value="<%=m_sLKyokan(n)%>">
		    <%
	        n = n + 1
        Loop%>
	<table class=hyo border=1>

	    <%
	        m_Grs.MoveFirst
	        Do Until m_Grs.EOF
	            Call gs_cellPtn(m_cell)
		        i = i + 1
		        j = 0
				w_iChkNo = 999
			    %>
			    <tr>
			        <td class=<%=m_cell%> width="50"><%=m_Grs("GAKUSEKI")%>
			        <input type=hidden name=gakuNo<%=i%> value="<%=m_Grs("GAKUSEI")%>"></td>
			        <td class=<%=m_cell%> width="120"><%=m_Grs("SIMEI")%>
			        <input type=hidden name=gakuNm<%=i%> value="<%=m_Grs("SIMEI")%>">
			        <input type=hidden name=L_KYOKAN<%=i%> value="<%=m_Grs("L_KYOKAN")%>"></td>
			    <%

				For n = 0 to UBound(m_sLKyokan)-1

					'対応する教官を選択していれば、○をつける
					If m_sLKyokan(n) = m_Grs("L_KYOKAN") then
						w_sChk = "○" 
						w_iChkNo = n 
						m_iLKyokanCnt(n) = m_iLKyokanCnt(n) + 1
					else 
						w_sChk = "　"
					End If

			If cint(m_sRisyuJotai) = C_K_RIS_ATO then 
				'確定処理後----------------------------------------------------- 
				w_sChg = ""
				
			Else
				'確定処理前----------------------------------------------------- 
				If m_sKengen <> C_WEB0390_ACCESS_SENMON then
					'権限が担当教官のみモードでない----------------------------------------------------- 
					w_sChg = "onclick='f_Chenge(""" & i & """,""" & n & """)'"
				Else
					'権限が担当教官のみモード----------------------------------------------------- 
					'm_sLKyokan(i)と教官ＣＤが一致(変更できる)----------------------------------------------------- 
					If m_sLKyokan(n) = m_sKyokanCd then 
						w_sChg = "onclick='f_Chenge(""" & i & """,""" & n & """)'"
					Else 
						w_sChg = ""
					End If
				End If
			End If

			%>
			        <td class=<%=m_cell%>   width="90">
			        <input type="button" class="<%=m_cell%>" name="K<%=i%>_<%=n%>" value="<%=w_sChk%>" <%=w_sChg%> style="text-align:center" >
					</td>
			<% 
		        Next
				%>
			        <input type=hidden name=L_KYOKAN_OLD<%=i%> value="<%=w_iChkNo%>">
				    </tr>
				<%
				m_Grs.MoveNext
	        Loop%>
	</table>
	<% If cint(m_sRisyuJotai) = C_K_RIS_MAE then %>
	<table>
	    <tr>
	        <td align=center><input type=button class=button value="　登　録　" onclick="javascript:f_Touroku()"></td>
	        <td align=center><input type=button class=button value="キャンセル" onclick="javascript:f_Cansel()"></td>
	    </tr>
	</table>
	<% End If %>

	<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="cboKamokuCode"      value="<%=m_sKamokuCD%>">

	<input type="hidden" name="txtGakunen"  value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass"    value="<%=m_sClass%>">
	<input type="hidden" name="txtRisyu"      value="<%=m_sRisyuJotai%>">
	<input type="hidden" name="txtGakuMax"      value="<%=m_GrCnt%>">

<% '教官を選択した学生数を隠して持たす
	For n=0 to UBound(m_sLKyokan)
%>
	<input type="hidden" name="txtLKCnt<%=n%>"      value="<%=m_iLKyokanCnt(n)%>">
<%
    Next
%>

	</FORM>
	</center>
	</BODY>
	</HTML>
<%
End Sub
%>