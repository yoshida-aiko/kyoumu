<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 特別教室予約
' ﾌﾟﾛｸﾞﾗﾑID : web/web0300/web0300_lst.asp
' 機      能: 教室情報を表示
'-------------------------------------------------------------------------
' 引      数:   NENDO           '//年度
'               KYOKAN_CD       '//教官CD
'				hidDay    		:日にち
'				hidYear    		:年
'				hidMonth   		:月
'				hidKyositu 		:教室CD
'
' 引      渡:	txtMode			:処理モード
'				hidJigen		:時限
'				YoyakKyokanCd	:予約教官CD
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
' 作      成: 2001/07/19 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
'
'	Const C_ACCESS_FULL   = "FULL"		'//アクセス権限FULLアクセス可
'	Const C_ACCESS_NORMAL = "NORMAL"	'//アクセス権限一般
'	Const C_ACCESS_VIEW   = "VIEW"		'//アクセス権限参照のみ

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	Public m_iSyoriNen			'//年度
	Public m_iKyokanCd			'//教官ｺｰﾄﾞ

	Public m_sYear   			'//年
	Public m_sMonth			  	'//月
	Public m_sDay   			'//日
	Public m_iKyosituCd			'//教室CD
	Public m_iKaijyoCnt			'//解除チェックボックスカウント
	Public m_iYoyakCnt			'//予約チェックボックスカウント
	Public m_sKyosituName		'//教室名称

	Public m_sUserId

    'ﾚｺｰﾄﾞセット
    Public m_Rs_Jigen       	'//時限ﾚｺｰﾄﾞｾｯﾄ
    Public m_Rs_Kyositu			'//教室予約情報

    Public m_bUpdate_OK			'//予約、更新可不可判別ﾌﾗｸﾞ
    Public m_sKengen

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

	'2011/12/26 Add Start
	Public m_sReservFrom		'//教室予約開始日
	Public m_sReservTo			'//教室予約終了日
	Public m_bReservationFlg	'//予約可不可判別フラグ
	'2011/12/26 Add End

	'2015/05/22 Ins Start
	Public m_JigenCount			'//時限数(2つで1つ)
	Public m_JigenDivFlg		'//時限数が全て2つにまとまるか
	'2015/05/22 Ins End

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

		'//権限を取得
		w_iRet = gf_GetKengen_web0300(m_sKengen)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'//権限より、表示内容を変える
        Call s_SetViewInfo()

'//デバッグ
'Call s_DebugPrint()

		'//教室名取得
		w_iRet = f_GetKyousituName()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//時限情報の取得
        w_iRet = f_GetJigen()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '2015/05/22 Ins Start
        //時限数を1,2→1、2,3→2、3,4→3･･･のように2つセットに変更
        w_JigenCount = gf_GetRsCount(m_Rs_Jigen)

        w_amari = w_JigenCount Mod 2

        IF w_amari <> 0 then
			m_JigenCount = (w_JigenCount / 2) + 0.5
			m_JigenDivFlg = False
		Else
			m_JigenCount = w_JigenCount / 2
			m_JigenDivFlg = True
        End If
        
        '2015/05/22 Ins End

        '// 教室予約状況の取得
        w_iRet = f_GetKyosituInfo()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'2011/12/26 Add Start
		'//教室予約日の取得
		w_iRet = f_GetReservationDate()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'//教室予約日より、予約、解除可能かどうかを判断する
		Call s_SetPreservation()
		'2011/12/26 Add End

        '// ページを表示
        Call showPage()
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
	
	m_sUserId = ""

	'2011/12/26 Add Start
	m_sReservFrom = ""
	m_sReservTo = ""
	'2011/12/26 Add End

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen  = Session("NENDO")
    'm_iKyokanCd  = Request("SKyokanCd1")
   'm_iKyokanCd  = SESSION("KYOKAN_CD")
    m_sYear      = Request("hidYear")
    m_sMonth     = Request("hidMonth")
    m_sDay       = Request("hidDay")
	m_iKyosituCd = Request("hidKyositu")

'	m_sUserId    = SESSION("LOGIN_ID")
	m_iKyokanCd  = SESSION("LOGIN_ID")

End Sub

'********************************************************************************
'*  [機能]  権限より、表示内容を変更する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetViewInfo()

	m_bUpdate_OK = False

	'//参照のみ可能な場合
	If m_sKengen = C_ACCESS_VIEW Then
		m_bUpdate_OK = False
	Else
		'//権限がFULLアクセスまたは、一般の場合
		m_bUpdate_OK = True
	End If

End Sub

'2011/12/26 Add
'********************************************************************************
'*  [機能]  教室予約日より、予約、解除可能かを判断する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetPreservation()
	Dim w_sToday

	w_sToday = gf_YYYY_MM_DD(date(),"/")
	If m_sReservFrom = "" Or m_sReservTo = "" Then
		m_bReservationFlg = False
	ElseIf m_sReservFrom <= w_sToday And m_sReservTo >= w_sToday Then
		m_bReservationFlg = True
	Else
		m_bReservationFlg = False
	End If
End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen  = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd  = " & m_iKyokanCd & "<br>"
    response.write "m_sYear      = " & m_sYear     & "<br>"
    response.write "m_sMonth     = " & m_sMonth    & "<br>"
    response.write "m_sDay       = " & m_sDay      & "<br>"
    response.write "m_iKyosituCd = " & m_iKyosituCd      & "<br>"

End Sub

'********************************************************************************
'*  [機能]  教室名取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKyousituName()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetKyousituName = 1

    Do
		'//教室名取得
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU.M06_KYOSITUMEI"
		w_sSql = w_sSql & vbCrLf & " FROM M06_KYOSITU"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU.M06_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M06_KYOSITU.M06_KYOSITU_CD=" & m_iKyosituCd

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKyousituName = 99
            Exit Do
        End If

		If rs.EOF = False Then
			m_sKyosituName = rs("M06_KYOSITUMEI")
		End If

        '//正常終了
        f_GetKyousituName = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  時限情報の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetJigen()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen = 1

    Do

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M07_JIKAN"
		w_sSql = w_sSql & vbCrLf & " FROM M07_JIGEN"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "      M07_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " GROUP BY M07_JIKAN"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs_Jigen, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetJigen = 99
            Exit Do
        End If

        '//正常終了
        f_GetJigen = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  教室予約状況の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKyosituInfo()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetKyosituInfo = 1

    Do

		w_sDate = m_sYear & "/" & gf_fmtZero(m_sMonth,2) & "/" &  gf_fmtZero(m_sDay,2)

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "   T58.T58_JIGEN"
		w_sSql = w_sSql & vbCrLf & "  ,T58.T58_KYOKAN_CD"
		w_sSql = w_sSql & vbCrLf & "  ,T58.T58_MOKUTEKI"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T58_KYOSITU_YOYAKU T58"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T58.T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND T58.T58_HIDUKE='" & w_sDate & "' "
		w_sSql = w_sSql & vbCrLf & "  AND T58.T58_KYOSITU=" & m_iKyosituCd

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs_Kyositu, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKyosituInfo = 99
            Exit Do
        End If

        '//正常終了
        f_GetKyosituInfo = 0
        Exit Do
    Loop

End Function

'2011/12/26 Add Start
'********************************************************************************
'*  [機能]  教室予約日の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetReservationDate()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetReservationDate = 1

    Do

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M00_NO"
		w_sSql = w_sSql & vbCrLf & "  ,M00_KANRI"
		w_sSql = w_sSql & vbCrLf & " FROM M00_KANRI"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "      M00_NENDO =" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND "
		w_sSql = w_sSql & vbCrLf & "      M00_NO IN (" & C_K_KYOSHITUYOYAKU_FROM & "," & C_K_KYOSHITUYOYAKU_TO & ")"
		w_sSql = w_sSql & vbCrLf & " ORDER BY M00_NO"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetReservationDate = 99
            Exit Do
        End If

		Do Until rs.EOF
			If Clng(rs("M00_NO")) = Clng(C_K_KYOSHITUYOYAKU_FROM) THEN
				m_sReservFrom = rs("M00_KANRI")
			Else
				m_sReservTo = rs("M00_KANRI")
			End If
			rs.MoveNext
		Loop
        '//正常終了
        f_GetReservationDate = 0
        Exit Do
    Loop

End Function
'2011/12/26 Add End

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
'*  [機能]  予約教室データを表示する
'*  [引数]  p_Jigen	：時限
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_KyousituData(p_Jigen,p_sClass)

	Dim w_sJigen
	Dim w_sMokuteki
	Dim w_sTourokusya
	Dim w_sTourokuCD

	w_sMokuteki = ""
	w_sTourokusya = ""
	w_sTourokuCD = ""
	w_btnUpdate = "<br>"
	w_chkKaijyo ="<br>"

	w_bYoyak = False

	Do

		If m_Rs_Kyositu.EOF = false then

			Do Until m_Rs_Kyositu.EOF

				'//取得した時限に教室予約が入っているか
				If clng(p_Jigen) = clng(m_Rs_Kyositu("T58_JIGEN")) Then

					w_bYoyak = True
					w_sJigen      = p_Jigen
					w_sMokuteki   = "<A href='javascript:f_LinkClick(" & p_Jigen & ");'>" & m_Rs_Kyositu("T58_MOKUTEKI")  & "</A>"
					w_sTourokusya = f_GetName(m_Rs_Kyositu("T58_KYOKAN_CD"))

						'//アクセス権限が参照以外の場合、本人関連のデータの修正・削除が可能となる
						If m_sKengen <> C_ACCESS_VIEW Then

							'//現在の利用者と登録されている利用者が同じ場合は修正ボタン及び解除チェックボックスを表示
'Response.Write "kyoukan=[" & m_Rs_Kyositu("T58_KYOKAN_CD") & "]" & "[" & m_iKyokanCd & "]"
							If NOT ISNULL(m_Rs_Kyositu("T58_KYOKAN_CD")) AND NOT ISNULL(m_iKyokanCd) Then
								If cstr(m_Rs_Kyositu("T58_KYOKAN_CD")) = cstr(m_iKyokanCd) Or cstr(m_Rs_Kyositu("T58_KYOKAN_CD")) = m_sUserId Then
									'2011/12/26 Upd Start
'									w_chkKaijyo = "<input type='checkbox' name='chkKaijyo' value='" & p_Jigen & "' >"
'									w_btnUpdate = "<input type='button'   name='btnUpdate' value='修正' class='button' onclick='javascript:f_UpdClick(" & p_Jigen & ")'>"
									If m_bReservationFlg = True Then
										w_chkKaijyo = "<input type='checkbox' name='chkKaijyo' value='" & p_Jigen & "' >"
										w_btnUpdate = "<input type='button'   name='btnUpdate' value='修正' class='button' onclick='javascript:f_UpdClick(" & p_Jigen & ")'>"
									Else
										w_btnUpdate = "<input type='button'   name='btnUpdate' value='参照' class='button' onclick='javascript:f_LinkClick(" & p_Jigen & ")'>"
									End If
									'2011/12/26 Upd End
									m_iKaijyoCnt = m_iKaijyoCnt + 1
								End If
							End If

						End If

					Exit Do

				End If

				m_Rs_Kyositu.MoveNext
			Loop

			m_Rs_Kyositu.MoveFirst
		End If

		Exit Do
	Loop

	'//すでに予約されているか
	If w_bYoyak = True Then
		'//予約があるとき予約ボタン非表示
		w_btnYoyak = "<br>"
	Else 
		'//予約がない時は予約ボタン表示
		'2011/12/26 Upd Start
'		w_btnYoyak  = "<input type='checkbox' name='hidYoyak' & value='" & trim(p_Jigen) & "'>"
		If m_bReservationFlg = True Then
			w_btnYoyak  = "<input type='checkbox' name='hidYoyak' & value='" & trim(p_Jigen) & "'>"
		Else
			w_btnYoyak = "<br>"
		End If
		'2011/12/26 Upd End
		w_sMokuteki = "空き</font>"
		w_sTourokusya = "―"
		w_sJigen   = p_Jigen
		m_iYoyakCnt = m_iYoyakCnt + 1
	End If

	%>
	<td class="<%=p_sClass%>" align="left"><%=w_sMokuteki%></td>
	<td class="<%=p_sClass%>" align="center" nowrap><%=w_sTourokusya%></td>

	<%'//権限により表示を制御%>

	<%If m_bUpdate_OK = True then%><td class="<%=p_sClass%>" align="center" nowrap><%=w_btnYoyak%></td><%End If%>
	<%If m_bUpdate_OK = True then%><td class="<%=p_sClass%>" align="center" nowrap><%=w_btnUpdate%></td><%End If%>
	<%If m_bUpdate_OK = True then%><td class="<%=p_sClass%>" align="center" nowrap><%=w_chkKaijyo%></td><%End If%>

<%
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
    //  [機能]  解除ボタンクリック
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_KaijyoClick() {

		//チェック欄数を取得
		var iMax = document.frm.hidKaijyoCnt.value
		if (iMax==0){
			//alert("No Avairable")
			return;
		}

		if(iMax==1){
			if(document.frm.chkKaijyo.checked==false){
				alert("解除するデータが選択されていません")
				return;
			}
		}else{

			var i
			var w_bCheck = 1
			for (i = 0; i < iMax; i++) {
				if(document.frm.chkKaijyo[i].checked==true){
					w_bCheck = 0
					break;
				}
			};

			if(w_bCheck == 1){
				alert("解除するデータが選択されていません")
				return;
			};
		};

		document.frm.txtMode.value="DISP";
		document.frm.action="web0300_del.asp";
		document.frm.target="bottom";
		document.frm.submit();

    }

    //************************************************************
    //  [機能]  修正ボタンクリック
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_UpdClick(p_Jigen){

		// document.frm.YoyakKyokanCd.value="imawaka"

		document.frm.hidJigen.value=p_Jigen;
		document.frm.txtMode.value="DETAIL";
		document.frm.action="web0300_detail.asp";
		document.frm.target="bottom";
		document.frm.submit();

    }

    //************************************************************
    //  [機能]  リンククリック
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_LinkClick(p_Jigen){

		// document.frm.YoyakKyokanCd.value=p_sKyokanCd;

		document.frm.hidJigen.value=p_Jigen;
		document.frm.txtMode.value="DISP";
		document.frm.action="web0300_detail.asp";
		document.frm.target="bottom";
		document.frm.submit();
    }

    //************************************************************
    //  [機能]  予約ボタンクリック
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
	function f_btnYoyakClick(){

		//チェック欄数を取得
		var iMax = document.frm.hidYoyakCnt.value
		if (iMax==0){
			//alert("No Avairable")
			return;
		}

		//チェックボックスが選択されているかチェック
		//選択されていればhidJigenに格納
		if(iMax==1){
			if(document.frm.hidYoyak.checked==false){
				alert("予約する時限が選択されていません")
				return;
			}else{
				document.frm.hidJigen.value = document.frm.hidYoyak.value
			};
		}else{

			var i
			for (i = 0; i < iMax; i++) {
				if(document.frm.hidYoyak[i].checked==true){
					if(document.frm.hidJigen.value==""){
						document.frm.hidJigen.value = document.frm.hidYoyak[i].value
					}else{
						document.frm.hidJigen.value = document.frm.hidJigen.value+","+document.frm.hidYoyak[i].value
					};

				};
			};

			if(document.frm.hidJigen.value==""){
				alert("予約する時限が選択されていません")
				return;
			};
		};

		document.frm.txtMode.value="BLANK";
		document.frm.action="web0300_detail.asp";
		document.frm.target="bottom";
		document.frm.submit();

	}

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

<%
'//デバッグ
'Call s_DebugPrint()
%>

	<center>
	<img src="img/sp.gif" height="3">
	<%Do%>

		<table border="1" width="98%" class="hyo">
		<tr>
		<th class=header width="50">時限</th>
		<th class=header>使用目的</th>
		<th class=header>利用者</th>

		<%If m_bUpdate_OK = True then%><th class=header>予約</th><%End If%>
		<%If m_bUpdate_OK = True then%><th class=header>修正</th><%End If%>
		<%If m_bUpdate_OK = True then%><th class=header>解除</th><%End If%>

		</tr>

		<%

		'//解除チェックボックスカウント数初期化
		m_iKaijyoCnt = 0

		'//予約チェックボックスカウント数初期化
		m_iYoyakCnt = 0

		'//2015/05/22 Ins Start
		w_iRowCount = 1
		w_iShowCount = 1
		'//2015/05/22 Ins End

		Do Until m_Rs_Jigen.EOF%>

			<!--2015/05/22 Ins Start -->
			<%  If w_iShowCount Mod 2 <> 0 Then
					w_sClass = "CELL1"
				Else
					w_sClass = "CELL2"
				End If
			%>

			<% If w_iRowCount Mod 2 <> 0 Then %>
				<% If w_iShowCount <> m_JigenCount Then %>
					<td rowspan="3" class="<%=w_sClass%>" align="center" height="25" width="50" nowrap><%=w_iShowCount%></td>
				<% Else %><!--最終時限-->
					<% If m_JigenDivFlg Then %>	
						<td rowspan="3" class="<%=w_sClass%>" align="center" height="25" width="50" nowrap><%=w_iShowCount%></td>
					<% Else %>
						<td rowspan="2" class="<%=w_sClass%>" align="center" height="25" width="50" nowrap><%=w_iShowCount%></td>
					<%End If%>
				<%End If%>
				<%w_iShowCount = w_iShowCount + 1%>
				

			<%End If%>
			<!--2015/05/22 Ins End -->

			<%if f_LenB(m_Rs_Jigen("M07_JIKAN")) < 3 then %>

				<tr>
				<%
				'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
				Call gs_cellPtn(w_Class)

				'//詳細データ表示
				Call f_KyousituData(m_Rs_Jigen("M07_JIKAN"),w_Class)
				%>
				</tr>

			<%End If%>
			<%w_iRowCount = w_iRowCount + 1%> <!--2015/05/22 Ins-->
			<%m_Rs_Jigen.MoveNext%>
		<%Loop%>

	    <tr>
		<%If m_bUpdate_OK = True then%>
		    <td colspan="4" align=right bgcolor=#9999BD>
			<%If m_bReservationFlg = True Then%><!--2011/12/26 Add-->
				<input class=button type=button value="予約" onclick="javascript:f_btnYoyakClick()">
			<%Else%>
				<br>
			<%End If%><!--2011/12/26 Add-->
			</td>
		<%End If%>


		<%If m_bUpdate_OK = True then%>
		    <td colspan="2" align=right bgcolor=#9999BD>
			<%If m_bReservationFlg = True Then%><!--2011/12/26 Add-->
				<input class=button type=button value="解除" onclick="javascript:f_KaijyoClick()">
			<%Else%>
				<br>
			<%End If%><!--2011/12/26 Add-->
			</td>
		<%End If%>

	    </tr>

		</table>

		<table width="98%" border=0>
			<tr>
<!--2011/12/28 Upd Start 
			<td align="right">
				<span class="msg"><font size="2">
					※教室予約は&nbsp;<%= m_sReservFrom & "～" & m_sReservTo%>&nbsp;までです。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
				<%If m_sKengen = C_ACCESS_VIEW Then%>
					※予約情報の詳細は、使用目的をクリックすると確認できます。
				<%Else%>
					※すでに予約されている時限には予約できません。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
					予約情報の詳細は、使用目的をクリックすると確認できます。<br>
					※修正・解除は登録者のみ可能です。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%End If%>
				</font></span>
-->
			<td width="45%">
			</td>
			<td width="55%" align="left">
				<span class="msg"><font size="2">
					※教室予約は&nbsp;<%= m_sReservFrom & "～" & m_sReservTo%>&nbsp;までです。<br>
				<%If m_sKengen = C_ACCESS_VIEW Then%>
					※予約情報の詳細は、使用目的をクリックすると確認できます。
				<%Else%>
					※すでに予約されている時限には予約できません。<br>
					&nbsp;&nbsp;&nbsp;予約情報の詳細は、使用目的をクリックすると確認できます。<br>
					※修正・解除は登録者のみ可能です。
				<%End If%>
				</font></span>
			</td>
		<%Exit Do%>
	<%Loop%>

	<!--値渡用-->
	<input type="hidden" name="hidYoyakCnt"    value="<%=m_iYoyakCnt%>">
	<input type="hidden" name="hidKaijyoCnt"   value="<%=m_iKaijyoCnt%>">
	<input type="hidden" name="txtMode"        value="">
	<input type="hidden" name="hidJigen"       value="">
	<input type="hidden" name="YoyakKyokanCd"  value="">
	<input type="hidden" name="SKyokanNm1"     value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">
	<input type="hidden" name="SKyokanCd1"     value="<%=m_iKyokanCd%>">

	<input type="hidden" name="hidDay"         value="<%=m_sDay%>">
	<input type="hidden" name="hidYear"        value="<%=m_sYear %>">
	<input type="hidden" name="hidMonth"       value="<%=m_sMonth%>">
	<input type="hidden" name="hidKyositu"     value="<%=m_iKyosituCd%>">
	<input type="hidden" name="hidKyosituName" value="<%=m_sKyosituName%>">

	</form>
	</center>
	</body>
	</html>

<%
End Sub
%>
