<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 実力試験成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0500/sei0500_middle.asp
' 機      能: 検索内容を表示する
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
'           :年度			＞		SESSIONより（保留）
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/09/07 モチナガ
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_iNendo			'年度
    Public  m_sKyokanCd	        '教官コード
    Public  m_sSiKenCd	        '試験コード
    Public  m_sGakuNo	        '学年
    Public  m_sClassNo	        '学科
    Public  m_sKamokuCd	        '科目コード
                                
    Public  m_Kaisi 			'成績入力期間（はじめ）
    Public  m_Syuryo			'成績入力期間（おわり）
    Public  m_Kamokumei			'科目名

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
	w_sMsgTitle="実力試験成績登録"
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

		'//期間データの取得
        w_iRet = f_Nyuryokudate()
		If w_iRet <> 0 Then 
			Exit Do
		End If

		'//科目名を取得
		w_iRet = f_GetKamokumei()
		If w_iRet <> 0 Then 
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

Sub s_SetParam()
'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSiKenCd	= Cint(request("txtShikenCd"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")

End Sub

Function f_Nyuryokudate()
'********************************************************************************
'*	[機能]	入力期間取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1

	Do
'2001/12/17 Mod ---->
'        w_sSQL = ""
'        w_sSQL = w_sSQL & vbCrLf & "  SELECT "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SEISEKI_KAISI, "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SEISEKI_SYURYO "
'        w_sSQL = w_sSQL & vbCrLf & "  FROM  "
'        w_sSQL = w_sSQL & vbCrLf & "    M28_SIKEN_KAMOKU M28 "
'        w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND " '試験区分(実力試験のみ)
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_NENDO        =  " & m_iNendo & "  AND "         '処理年度
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SIKEN_CD     =  " & m_sSiKenCd & "  AND "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SIKEN_KAMOKU = '" & m_sKamokuCd & "' AND "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_GAKUNEN      =  " & m_sGakuNo & "  AND "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_CLASS        =  " & m_sClassNo & "  AND "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SEISEKI_KAISI  <= '" & gf_YYYY_MM_DD(Date, "/") & "' AND"
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(Date, "/") & "' "
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SEISEKI_KAISI, "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SEISEKI_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & "  FROM  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND "	'試験区分(実力試験のみ)
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO        =  " & m_iNendo	& "  AND "		'処理年度
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_CD     =  " & m_sSiKenCd  & "  AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KAMOKU = '" & m_sKamokuCd & "'  "
'		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_GAKUNEN      =  " & m_sGakuNo   & "  AND "
'		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_CLASS        =  " & m_sClassNo  & "  AND "
'		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_SEISEKI_KAISI  <= '" & gf_YYYY_MM_DD(date(),"/") & "' AND"
'		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "
'2001/12/17 Mod <----

'response.write w_sSQL & "<br>"

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If Not w_Rs.EOF Then
'			m_Kaisi  = w_Rs("M28_SEISEKI_KAISI")
'			m_Syuryo = w_Rs("M28_SEISEKI_SYURYO")
			m_Kaisi  = w_Rs("M27_SEISEKI_KAISI")
			m_Syuryo = w_Rs("M27_SEISEKI_SYURYO")
		End If

	    Call gf_closeObject(w_Rs)

		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function

Function f_GetKamokumei()
'********************************************************************************
'*	[機能]	科目名取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************

	On Error Resume Next
	Err.Clear
	f_GetKamokumei = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_KAMOKUMEI "
		w_sSQL = w_sSQL & vbCrLf & "  FROM "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO        =  " & m_iNendo	& " AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_CD     =  " & m_sSiKenCd  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KAMOKU = '" & m_sKamokuCd & "' "

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_GetKamokumei = 99
			m_bErrFlg = True
			Exit Do 
		End If

		if Not w_Rs.Eof then
			m_Kamokumei = w_Rs("M27_KAMOKUMEI")
		End if

	    Call gf_closeObject(w_Rs)

		f_GetKamokumei = 0
		Exit Do
	Loop

End Function


'********************************************************************************
'*  [機能]  試験時間等を取得
'*  [引数]  なし
'*  [戻値]  
'*  [説明]  
'********************************************************************************
Function f_GetSikenJikan()

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_GetSikenJikan = ""
	p_KamokuName = ""

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_KAMOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_KAISI_JIKOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_SYURYO_JIKOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_SIKENBI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO = T26_SIKEN_JIKANWARI.T26_CLASS "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_GAKUNEN = M05_CLASS.M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_NENDO = M05_CLASS.M05_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_KBN=" & Cint(m_sSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_CD='0' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_GAKUNEN=" & cint(m_sGakuNo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKKA_CD='" & m_sGakkaCd & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_KAMOKU='" & m_sKamokuCd & "'"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
			f_GetSikenJikan = 99
            Exit Do
        End If

		If Not w_Rs.EOF Then
			m_sKaisiT = w_Rs("T26_KAISI_JIKOKU")
			m_sSyuryoT = w_Rs("T26_SYURYO_JIKOKU")
			m_sSikenbi = w_Rs("T26_SIKENBI")
		End If

		f_GetSikenJikan = 0
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

    On Error Resume Next
    Err.Clear

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
    }

   //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){
        parent.main.f_Touroku();
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
	//	[機能]	ペーストボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//************************************************************
	function f_Paste(pType){

		parent.main.document.frm.PasteType.value=pType;

		//submitで画面を開くとウィンドウのステータスが設定できないため､
		//一旦空ページを開いてから、新ウィンドウに対してsubmitする。
		nWin=open("","Paste","location=no,menubar=no,resizable=yes,scrollbars=no,scrolling=no,status=no,toolbar=no,width=300,height=600,top=0,left=0");
		parent.main.document.frm.target="Paste";
		parent.main.document.frm.action="sei0500_paste.asp";
		parent.main.document.frm.submit();
	
	}
	//-->
	</SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	<% call gs_title(" 実力試験成績登録 "," 登　録 ") %>
	<center>
	<table border="0" cellpadding="0" cellspacing="0"><tr><td align="center">
		<table border=1 class=hyo>
			<tr>
				<th class=header align="center" colspan="2">実力試験成績入力期間</th>
			</tr>
			<tr>
				<th class=header width="96"  align="center">成績入力期間</th>
				<td class=detail width="360" align="center"><%=m_Kaisi%> ～ <%=m_Syuryo%></td>
			</tr>
			<tr>
				<th class=header width="96"  align="center">科目</th>
				<td class=detail width="360" align="center"><%=m_sGakuNo%>年　<%= gf_GetClassName(m_iNendo,m_sGakuNo,m_sClassNo) %>　<%= m_Kamokumei %></td>

			</tr>
		</table>
	</td></td>
	<tr><td align="center"><span class=msg>※ 成績を入力して、登録ボタンを押してください。</span><br>
	※ヘッダの文字色が「<FONT COLOR="#99CCFF">成績</FONT>」のようになっている部分をクリックすると、Excel貼り付け用の画面が開きます。</span></td></tr>
	</table>

	<table width=50%>
		<tr>
			<td align=center nowrap>
				<input type=button class=button value="　登　録　" onclick="javascript:f_Touroku()">　
				<input type=button class=button value="キャンセル" onclick="javascript:f_Cansel()"></td>
		</tr>
	</table>

	<table >
		<tr>
			<td valign="top">
				<table class="hyo" border=1 align="center" width="280">
					<tr>
						<th class="header" width="50"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
						<th class="header" width="200">氏　名</th>
						<th class="header" width="30" nowrap onClick="f_Paste('Seiseki')"><Font COLOR="#99CCFF">成績</Font></th>
					</tr>
				</table>
			</td>
			<td valign="top">
				<table class="hyo" border=1 align="center" width="280">
					<tr>
						<th class="header" width="50"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
						<th class="header" width="200">氏　名</th>
						<th class="header" width="30">成績</th>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	</FORM>
	</center>
	</body>
	</html>
<%
End sub

%>