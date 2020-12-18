<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 特別教室予約
' ﾌﾟﾛｸﾞﾗﾑID : web/web0300/web0300_lst.asp
' 機      能: 教室情報を表示
'-------------------------------------------------------------------------
' 引      数:   NENDO           '//年度
'				YoyakKyokanCd	:予約教官CD
'				txtMode			:処理モード
'				hidJigen		:時限
'				hidDay			:日にち
'				hidYear			:年
'				hidMonth		:月
'				hidKyositu		:教室CD
'				hidKyosituName	:教室名称
' 引      渡:
'				YoyakKyokanCd	:予約教官CD
'				txtMode			:処理モード
'				hidJigen		:時限
'				hidDay			:日にち
'				hidYear			:年
'				hidMonth		:月
'				hidKyositu		:教室CD
'				hidKyosituName	:教室名称
' 説      明:
'           ■初期表示
'               空白ページを表示
'           ■表示ボタンが押された場合
'               検索条件にかなった試験時間割を表示
'-------------------------------------------------------------------------
' 作      成: 2001/08/07 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	Public m_iSyoriNen          '//年度
	Public m_iKyokanCd          '//教官ｺｰﾄﾞ

	Public m_sYear   			'//年
	Public m_sMonth				'//月
	Public m_sDay    			'//日
	Public m_iKyosituCd			'//教室CD
	Public m_iKaijyoCnt			'//解除チェックボックスカウント
	Public m_sMode				'//処理モード
	Public m_iJigen				'//時限
	Public m_sMokuteki			'//目的
	Public m_sBiko				'//備考
	Public m_sKyosituName		'//教室名称

	Public m_sUserId

	Public m_iKyokanCdUpd

    'ﾚｺｰﾄﾞセット
    Public m_Rs_Jigen           '//時限ﾚｺｰﾄﾞｾｯﾄ
    Public m_Rs_Kyositu			'//教室予約情報

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="特別教室予約"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

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

        '//値の初期化
        Call s_ClearParam()

        '//変数セット
        Call s_SetParam()

'//デバッグ
Call s_DebugPrint()


		'//処理モードにより処理を振り分ける
		Select Case m_sMode

			'//新規登録入力用フォーム表示
			Case "BLANK"
				'//新規登録画面を表示
				Call showPage()

			'//修正ボタンクリック時,修正画面を表示
			Case "DETAIL"

				'//表示用データ取得
				w_iRet = f_GetDispData()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//画面を表示
				Call showPage()

			'//リンククリック時画面を表示
			Case "DISP"

				'//表示用データ取得
				w_iRet = f_GetDispData()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//画面を表示
				Call showPage()

			'//新規登録処理(リロード時)
			Case "INSERT"
				'//データINSERT
				w_iRet = f_DataInsert()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//登録正常終了時
				Call showWhitePage(C_TOUROKU_OK_MSG)

			'//更新処理(リロード時)
			Case "UPDATE"
				'//データUPDATE
				w_iRet = f_DataUpdate()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//更新正常終了時
				Call showWhitePage(C_UPDATE_OK_MSG)

			Case Else

		End Select

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_Rs_Jigen)

    '// 終了処理
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen  = ""
    m_iKyokanCd  = ""
    m_sYear      = ""
    m_sMonth     = ""
    m_sDay       = ""
	m_iKyosituCd = ""
	m_sMode      = ""
	m_iJigen     = ""
	m_sMokuteki  = ""
	m_sBiko      = ""

	m_sUserId = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen    = Session("NENDO")
	'm_iKyokanCdUpd = Request("YoyakKyokanCd")
	m_iKyokanCdUpd = trim(Request("SKyokanCd1"))
	m_iKyokanCd    = trim(session("KYOKAN_CD"))
    m_sYear        = Request("hidYear")
    m_sMonth       = Request("hidMonth")
    m_sDay         = Request("hidDay")
	m_iKyosituCd   = Request("hidKyositu")
	m_sMode        = Request("txtMode")
	m_iJigen       = Request("hidJigen")
	m_sKyosituName = Request("hidKyosituName")

	m_sUserId = trim(Session("LOGIN_ID"))

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_sMode        = " & m_sMode      & "<br>"
    response.write "m_iSyoriNen    = " & m_iSyoriNen  & "<br>"
    response.write "m_iKyokanCd    = " & m_iKyokanCd  & "<br>"
    response.write "m_sYear        = " & m_sYear      & "<br>"
    response.write "m_sMonth       = " & m_sMonth     & "<br>"
    response.write "m_sDay         = " & m_sDay       & "<br>"
    response.write "m_iKyosituCd   = " & m_iKyosituCd & "<br>"
    response.write "m_iJigen       = " & m_iJigen     & "<br>"
    response.write "m_sMokuteki    = " & m_sMokuteki  & "<br>"
    response.write "m_sBiko        = " & m_sBiko      & "<br>"
    response.write "m_sKyosituName = " & m_sKyosituName & "<br>"

End Sub

'********************************************************************************
'*  [機能]  表示データを取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetDispData()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen = 1

    Do
		'//日付を作成
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'//教室予約データ取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & " T58_MOKUTEKI, "
		w_sSql = w_sSql & vbCrLf & " T58_KYOKAN_CD, "
		w_sSql = w_sSql & vbCrLf & " T58_BIKO"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & " T58_KYOSITU_YOYAKU"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & " T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND T58_HIDUKE='" & w_sDate & "'"
		w_sSql = w_sSql & vbCrLf & " AND T58_JIGEN=" & cint(m_iJigen)
		w_sSql = w_sSql & vbCrLf & " AND T58_KYOSITU=" & m_iKyosituCd
		'w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_iKyokanCd & "'"

response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetDispData = 99
            Exit Do
        End If

		If rs.EOF = False Then
			m_sMokuteki = rs("T58_MOKUTEKI")
			m_sBiko     = rs("T58_BIKO")
			m_iKyokanCd = rs("T58_KYOKAN_CD")
		End If

        '//正常終了
        f_GetDispData = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  データINSERT
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_DataInsert()

    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_iJigen
	Dim w_iJikanCnt
	Dim w_sRiyosya
	Dim w_sMokuteki
	Dim w_sBiko
	Dim w_sInsUser

    On Error Resume Next
    Err.Clear

    f_DataInsert = 1

    Do
		'//日付を作成
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'w_sRiyosya = trim(Request("SKyokanCd1"))
		'w_sRiyosya = trim(session("KYOKAN_CD"))
		w_sMokuteki = trim(Request("txtMokuteki"))
		w_sBiko = trim(Request("txtBiko"))
		'w_sInsUser = trim(Session("LOGIN_ID"))

        '//時限情報を取得
        w_iJigen = split(replace(m_iJigen," ",""),",")
        w_iJikanCnt = UBound(w_iJigen)

        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()

		For i=0 To w_iJikanCnt

			'//INSERT
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & " INSERT INTO T58_KYOSITU_YOYAKU"
			w_sSql = w_sSql & vbCrLf & " ("
			w_sSql = w_sSql & vbCrLf & " T58_NENDO "
			w_sSql = w_sSql & vbCrLf & " ,T58_HIDUKE "
			w_sSql = w_sSql & vbCrLf & " ,T58_YOUBI_CD "
			w_sSql = w_sSql & vbCrLf & " ,T58_JIGEN "
			w_sSql = w_sSql & vbCrLf & " ,T58_KYOSITU "
			w_sSql = w_sSql & vbCrLf & " ,T58_KYOKAN_CD "
			w_sSql = w_sSql & vbCrLf & " ,T58_MOKUTEKI "
			w_sSql = w_sSql & vbCrLf & " ,T58_BIKO "
			w_sSql = w_sSql & vbCrLf & " ,T58_INS_DATE "
			w_sSql = w_sSql & vbCrLf & " ,T58_INS_USER"
			w_sSql = w_sSql & vbCrLf & " ) VALUES ("
			w_sSql = w_sSql & vbCrLf & " "   & m_iSyoriNen
			w_sSql = w_sSql & vbCrLf & " ,'" & w_sDate & "'"
			w_sSql = w_sSql & vbCrLf & " ,"  & Weekday(w_sDate)
			w_sSql = w_sSql & vbCrLf & " ,"  & cint(w_iJigen(i))
			w_sSql = w_sSql & vbCrLf & " ,"  & m_iKyosituCd
			'w_sSql = w_sSql & vbCrLf & " ,'" & w_sRiyosya  & "'"
	'//教官ＣＤがない場合はユーザーＩＤを入力する 2002.1.8
	'If m_iKyokanCd <> "" then
	'		w_sSql = w_sSql & vbCrLf & " ,'" & m_iKyokanCd & "'"
	'Else
			w_sSql = w_sSql & vbCrLf & " ,'" & m_sUserId & "'"
	'End If
			w_sSql = w_sSql & vbCrLf & " ,'" & w_sMokuteki & "'"
			w_sSql = w_sSql & vbCrLf & " ,'" & w_sBiko     & "'"
			w_sSql = w_sSql & vbCrLf & " ,'" & gf_YYYY_MM_DD(date(),"/") & "'"
			w_sSql = w_sSql & vbCrLf & " ,'" & w_sInsUser & "'"
			w_sSql = w_sSql & vbCrLf & " )"

response.write w_sSQL & "<br>"
response.end

			iRet = gf_ExecuteSQL(w_sSQL)
			If iRet <> 0 Then
                '//ﾛｰﾙﾊﾞｯｸ
                Call gs_RollbackTrans()
				'登録失敗
				f_DataInsert = 99
				Exit Do
			End If

		Next

        '//ｺﾐｯﾄ
        Call gs_CommitTrans()

        '//正常終了
        f_DataInsert = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  データUPDATE
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_DataUpdate()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_DataUpdate = 1

    Do
		'//日付を作成
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'//UPDATE
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " UPDATE T58_KYOSITU_YOYAKU SET"
		w_sSql = w_sSql & vbCrLf & "  T58_MOKUTEKI='" & trim(Request("txtMokuteki")) & "'"
		w_sSql = w_sSql & vbCrLf & " ,T58_BIKO='" & trim(Request("txtBiko")) & "'"
		w_sSql = w_sSql & vbCrLf & " ,T58_UPD_DATE='" & gf_YYYY_MM_DD(date(),"/") & "'"
		w_sSql = w_sSql & vbCrLf & " ,T58_UPD_USER='" & Session("LOGIN_ID") & "'"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & " T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND T58_HIDUKE='" & w_sDate & "'"
		w_sSql = w_sSql & vbCrLf & " AND T58_JIGEN=" & cint(m_iJigen)
		w_sSql = w_sSql & vbCrLf & " AND T58_KYOSITU=" & cint(m_iKyosituCd)
		
		If m_iKyokanCd <> "" then
			w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_iKyokanCd & "'"
		Else
			w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_sUserId & "'"
		End if

'response.write w_sSQL & "<br>"
'response.end

		iRet = gf_ExecuteSQL(w_sSQL)
		If iRet <> 0 Then
			'登録失敗
			msMsg = Err.description
			f_DataUpdate = 99
			Exit Do
		End If

		'//正常終了
		f_DataUpdate = 0
		Exit Do
	Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  利用者名を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  T58_KYOKAN_CDには、教官CDかUSERID(M10)のどちらかが入っているので、
'*          はじめに、教官マスタを検索し名称が取得できなかった場合はUSERマスタをみる
'********************************************************************************
Function f_GetName(p_sUserId)
    Dim w_iRet
	Dim w_sUserName

    On Error Resume Next
    Err.Clear

    f_GetName = ""
	w_sUserName = ""

    Do

		'//教官マスタより、教官名を取得する
		w_sUserName = gf_GetKyokanNm(m_iSyoriNen,p_sUserId)

		'//教官名称が取得できなかった場合
		If Trim(w_sUserName) = "" Then
			'//USERマスタより、USER名を取得する
			w_sUserName = gf_GetUserNm(m_iSyoriNen,p_sUserId)
		End If

        Exit Do
    Loop

    f_GetName = w_sUserName

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()

%>
    <html>
    <head>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>特別教室予約</title>

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
    //  [機能]  キャンセルボタンクリック
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_Cancel() {

		document.frm.action="web0300_lst.asp";
		document.frm.target="bottom";
		document.frm.submit();
    }

    //************************************************************
    //  [機能]  登録ボタンクリック
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_Touroku(){

		// 入力値のﾁｪｯｸ
		iRet = f_CheckData();
		if( iRet != 0 ){
			return;
		}

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		<%If m_sMode="BLANK" Then%>
			document.frm.txtMode.value="INSERT";
		<%Else%>
			document.frm.txtMode.value="UPDATE";
		<%End If%>

		document.frm.action="web0300_detail.asp";
		document.frm.target="bottom";
		document.frm.submit();

    }

    //************************************************************
    //  [機能]  入力値のﾁｪｯｸ
    //  [引数]  なし
    //  [戻値]  0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
    //  [説明]  入力値のNULLﾁｪｯｸ、英数字ﾁｪｯｸ、桁数ﾁｪｯｸを行う
    //          引渡ﾃﾞｰﾀ用にﾃﾞｰﾀを加工する必要がある場合には加工を行う
    //************************************************************
    function f_CheckData() {
    
		// ■■■NULLﾁｪｯｸ■■■
		// ■目的
		//if( f_Trim(document.frm.txtMokuteki.value) == "" ){
		//	window.alert("目的が入力されていません");
		//	document.frm.txtMokuteki.focus();
		//	return 1;
		//}

//		// ■備考
//		if( f_Trim(document.frm.txtBiko.value) == "" ){
//			window.alert("備考が入力されていません");
//			document.frm.txtBiko.focus();
//			return 1;
//		}

		// ■■■文字数ﾁｪｯｸ■■■
		// ■目的
		if( getLengthB(document.frm.txtMokuteki.value) > "50" ){
			window.alert("目的は全角25文字以内で入力してください");
			document.frm.txtMokuteki.focus();
			return 1;
		}

		// ■備考
		if( getLengthB(document.frm.txtBiko.value) > "200" ){
			window.alert("備考は全角100文字以内で入力してください");
			document.frm.txtBiko.focus();
			return 1;
		}

        return 0;
    }


    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post" onSubmit="return false">

<%
'//デバッグ
'Call s_DebugPrint()
%>
<br>
	<center>


		<table border="1" class="hyo" width="98%">

		<tr>
		<th CLASS="header" width="90"  nowrap>日付</th>
		<td class="detail" ><%=gf_fmtWareki(gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/"))%><BR></td>
		</tr>

		<tr>
		<th CLASS="header" width="90"  nowrap>時限</th>
		<%
		If m_sMode="BLANK" Then
			w_sJigen =replace(replace(m_iJigen," ",""),",","時限、") & "時限"
		Else
			w_sJigen = m_iJigen & "時限目"
		End If
		%>
		<td class="detail"><%=w_sJigen%></td>
		</tr>

		<%
		'//表示のみの時
		If m_sMode="DISP" Then%>
			<tr>
			<th CLASS="header" width="90"  nowrap>利用者</th>
			<td class="detail" ><%=f_GetName(m_iKyokanCd)%><BR></td>
			</tr>
		<%End If%>

		<tr>
		<th CLASS="header" width="90" nowrap>教室</th>
		<td class="detail"><%=m_sKyosituName%><BR></td>
		</tr>

		<tr>
		<th CLASS="header" width="90" nowrap>使用目的</th>
		<%
		'//表示のみの時
		If m_sMode="DISP" Then%>
			<td class="detail" height="20"><%=m_sMokuteki%><BR></td>
		<%Else%>
			<td class="detail"><input type="text" name="txtMokuteki" value="<%=m_sMokuteki%>" maxlength="50" size="70">
			</td>
		<%End If%>

		</tr>

		<tr>
		<th CLASS="header" width="90" nowrap>備考</th>
		<%
		'//表示のみの時
		If m_sMode="DISP" Then%>
				<td class="detail" height="40" valign="top"><%=m_sBiko%><BR></td>
			</tr>
			</table>

		<%Else%>
				<td class="detail"><textarea rows="4" cols="50" WRAP="soft" class="text" name="txtBiko" ><%=m_sBiko%></textarea>
				<br><font size=2>（全角100文字以内）</font>
				</td>
			</tr>
			</table>
		<%End If%>

		<br>

		<table width="250">
		<tr>
		<%If m_sMode="DISP" Then%>
			<td align="center"><input class="button" type="button" value="閉じる" onclick="javascript:f_Cancel()"></td>
		<%Else%>
			<td align="center"><input class="button" type="button" value="　登　録　" onclick="javascript:f_Touroku()"></td>
			<td align="center"><input class="button" type="button" value="キャンセル" onclick="javascript:f_Cancel()"></td>
		<%End If%>
		</tr>
		</table>

	<!--値渡し用-->
	<input type="hidden" name="txtMode"       value="">
	<input type="hidden" name="hidJigen"      value="<%=m_iJigen%>">
	<input type="hidden" name="YoyakKyokanCd" value="<%=m_iKyokanCd%>">
	<input type="hidden" name="SKyokanCd1"    value="<%=Request("SKyokanCd1")%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">

	<input type="hidden" name="hidDay"     value="<%=m_sDay%>">
	<input type="hidden" name="hidYear"    value="<%=m_sYear %>">
	<input type="hidden" name="hidMonth"   value="<%=m_sMonth%>">
	<input type="hidden" name="hidKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="hidKyosituName" value="<%=m_sKyosituName%>">

	</form>
	</center>
	</body>
	</html>

<%
End Sub

'********************************************************************************
'*  [機能]  空白ページ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showWhitePage(p_sOkMsg)
%>
    <html>
    <head>
    <meta>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>特別教室予約</title>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		alert("<%=p_sOkMsg%>")

		var wArg

		//カレンダーページを再表示
		wArg = ""
		wArg = wArg + "?TUKI=<%=m_sMonth%>"
		wArg = wArg + "&cboKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidDay=<%=m_sDay%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=Server.URLEncode(Request("SKyokanCd1"))%>"

		parent.middle.location.href="./calendar.asp"+wArg

		//リストページを再表示
		wArg = ""
		wArg = wArg + "?hidDay=<%=m_sDay%>"
		wArg = wArg + "&hidYear=<%=m_sYear%>"
		wArg = wArg + "&hidMonth=<%=m_sMonth%>"
		wArg = wArg + "&hidKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidKyosituName=<%=Server.URLEncode(m_sKyosituName)%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=Server.URLEncode(Request("SKyokanCd1"))%>"

		parent.bottom.location.href="./web0300_lst.asp"+wArg

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type="hidden" name="TUKI"       value="<%=m_sMonth%>">
	<input type="hidden" name="cboKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">
	<input type="hidden" name="SKyokanCd1" value="<%=Server.HTMLEncode(Request("SKyokanCd1"))%>">

	<input type="hidden" name="hidDay"     value="<%=m_sDay%>">
	<input type="hidden" name="hidYear"    value="<%=m_sYear %>">
	<input type="hidden" name="hidMonth"   value="<%=m_sMonth%>">
	<input type="hidden" name="hidKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="hidKyosituName" value="<%=m_sKyosituName%>">

	</form>
	</body>
	</html>
<%
End Sub
%>