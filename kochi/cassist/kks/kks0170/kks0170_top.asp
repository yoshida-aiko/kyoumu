<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 日毎出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0140/kks0140_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:cboDate        :日付
'            TUKI           :月
' 説      明:
'           ■初期表示
'               月、日のコンボボックスは本日を表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう担任クラス一覧を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 伊藤公子
' 変      更: 2001/07/30 伊藤公子　長期休暇を表示しないように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
	Const C_KYUKA_TYOUKI = 1	'//長期休暇ﾌﾗｸﾞ(長期ﾌﾗｸﾞ)
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iSyoriNen          '教官ｺｰﾄﾞ
    Public m_iKyokanCd          '年度
    Public m_iTuki              '//月
    Public m_sDate              '//日付
    Public m_sDateWhere

    Public m_sAryDay()			'//長期休暇と土日以外の日付
	Public m_iCnt				'//長期休暇と土日以外の日付数

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
    w_sMsgTitle="日毎出欠入力"
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

        '//コンボ日付を取得(長期休暇及び土日祝日を除く日付)
        w_iRet = f_SetDay()
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

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sClassMei = ""
    m_iTuki = ""
    m_sDate = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")    
    m_iKyokanCd = Session("KYOKAN_CD")

    '//月情報を取得
    If request("TUKI") <> "" Then
        m_iTuki = request("TUKI")
    Else
        m_iTuki = month(date())
        m_sDate = gf_YYYY_MM_DD(date(),"/")
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

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"
    response.write "m_sClassMei = " & m_sClassMei & "<br>"

End Sub

'********************************************************************************
'*  [機能]  日付コンボデータを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_SetDay()

    Dim w_iRet
    Dim w_sSQL
    Dim rs
    Dim w_sSDate
    Dim w_sEDate

    On Error Resume Next
    Err.Clear

    f_SetDay = 1

    Do

        '//1～3月
        If m_iTuki <= 3 Then
            w_iNen = cint(m_iSyoriNen)+1
        Else
            w_iNen = cint(m_iSyoriNen)
        End If

        '//月の検索条件を作成
        '//開始日
        w_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iTuki),2) & "/01"

        '//終了日
        If Cint(m_sTuki) = 12 Then
            w_sEDate = cstr(w_iNen+1) & "/01/01"
        Else
            w_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iTuki+1),2) & "/01"
        End If

        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " A.T32_HIDUKE"
        w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M A"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & " A.T32_NENDO=" & cInt(m_iSyoriNen)
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE>='" & w_sSDate & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE< '" & w_sEDate & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='" & C_HEIJITU & "'"
        w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE"
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE"

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_SetDay = 99
            Exit Do
        End If

		i = 0

		If rs.EOF = False Then

			Do Until rs.EOF

				w_bKyuka = False
				w_sDate = rs("T32_HIDUKE")

				'//取得した日付が長期休暇かどうか(長期休暇…w_bKyuka = True)
				w_iRet = f_GetKyukaInfo(w_sDate,w_bKyuka)
				If w_iRet <> 0 Then
		            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		            f_SetDay = 99
		            Exit Function
				End If

				'//長期休暇でない日付を取得
				If w_bKyuka = False Then
					ReDim Preserve m_sAryDay(i)
					m_sAryDay(i) = w_sDate
					i = i + 1
				End If

				rs.MoveNext
			Loop

		End If

		m_iCnt = i-1

        '//正常終了
        f_SetDay = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  取得した日付が長期休暇かどうか
'*  [引数]  p_sDate  : 日付
'*  [戻値]  p_bKyuka : 長期休暇=True 長期休暇でない = False
'*  [説明]  
'********************************************************************************
Function f_GetKyukaInfo(p_sDate,p_bKyuka)

    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_sGakunen,w_sClassNo

    On Error Resume Next
    Err.Clear

    f_GetKyukaInfo = 1
	p_bKyuka = False	'//長期休暇ﾌﾗｸﾞ

    Do

		'//担任クラス情報を取得
		iRet = f_GetClassInfo(w_sGakunen,w_sClassNo)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_GetKyukaInfo = 99
            Exit Do
        End If

		'//長期休暇かどうか
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  Count(T31_GYOJI_H.T31_GYOJI_CD) AS CNT"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H "
		w_sSQL = w_sSQL & vbCrLf & "  ,T32_GYOJI_M "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_CD = T32_GYOJI_M.T32_GYOJI_CD "
		w_sSQL = w_sSQL & vbCrLf & "  AND T31_GYOJI_H.T31_NENDO = T32_GYOJI_M.T32_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND T31_GYOJI_H.T31_KYUKA_FLG='" & cstr(C_KYUKA_TYOUKI) & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T31_GYOJI_H.T31_NENDO=" & cInt(m_iSyoriNen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T32_GYOJI_M.T32_GAKUNEN IN (" & cint(w_sGakunen) & "," & cint(C_GAKUNEN_ALL) & ")"
		w_sSQL = w_sSQL & vbCrLf & "  AND T32_GYOJI_M.T32_CLASS IN (" & cint(w_sClassNo) & "," & cint(C_CLASS_ALL) & ") "
		w_sSQL = w_sSQL & vbCrLf & "  AND T32_GYOJI_M.T32_HIDUKE='" & p_sDate & "'"

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKyukaInfo = 99
            Exit Do
        End If

		'//長期休暇データを取得した場合
		If cint(rs("CNT")) > 0 Then
			p_bKyuka = True
		End If

        '//正常終了
        f_GetKyukaInfo = 0
        Exit Do
    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  教官CDより、担任クラス情報を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetClassInfo(p_sGakunen,p_sClassNo)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassInfo = 1

	Do
		'クラスマスタからクラス情報を取得
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & "    M05_NENDO,"
		w_sSQL = w_sSQL & "    M05_GAKUNEN,"
		w_sSQL = w_sSQL & "    M05_CLASSNO,"
		w_sSQL = w_sSQL & "    M05_CLASSMEI"
		w_sSQL = w_sSQL & " FROM M05_CLASS"
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & "       M05_TANNIN = '" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & "   AND M05_NENDO = " & cInt(m_iSyoriNen)

'response.write w_sSQL & "<br>"
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			f_GetClassInfo = 99
			Exit Do
		End If

		If rs.EOF = False Then
			p_sGakunen  = rs("M05_GAKUNEN")
			p_sClassNo  = rs("M05_CLASSNO")
		End If

		'//正常終了
		f_GetClassInfo = 0
		Exit Do
	Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'       (リストダウンボックス選択表示用)
'[引数] pData1 : データ１
'       pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'                   
'****************************************************
Function f_Selected(pData1,pData2)

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else 
            f_Selected = "" 
        End If
    End If

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
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>日毎出欠入力</title>

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
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

		if(document.frm.cboDate.value==""){
			alert("対象日がありません");
			return;
		}

        document.frm.action="./kks0170_bottom.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  月を変更した時
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ChangeTuki(){

        //本画面をsubmit
        document.frm.target = "_self";
        document.frm.action = "./kks0170_top.asp"
        document.frm.submit();
        return;
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
    <%call gs_title("日毎出欠入力","一　覧")%>
    <table border="0" >
	    <tr>
		    <td colspan="2" align="center"><span class=CAUTION>※ 土日祝日及び長期休暇時は、登録できません</span></td>
	    </tr>
    <tr>
        <td class=search>
            <table border="0" cellpadding="0" cellspacing="0">
                <tr>
				<td>入力対象日</td>
                <td nowrap >　<select name="TUKI" onchange="javascript:f_ChangeTuki();" style="width:50px;">
                        <option value="4"  <%=f_Selected("4" ,cstr(m_iTuki))%> >4
                        <option value="5"  <%=f_Selected("5" ,cstr(m_iTuki))%> >5
                        <option value="6"  <%=f_Selected("6" ,cstr(m_iTuki))%> >6
                        <option value="7"  <%=f_Selected("7" ,cstr(m_iTuki))%> >7
                        <option value="8"  <%=f_Selected("8" ,cstr(m_iTuki))%> >8
                        <option value="9"  <%=f_Selected("9" ,cstr(m_iTuki))%> >9
                        <option value="10" <%=f_Selected("10",cstr(m_iTuki))%> >10
                        <option value="11" <%=f_Selected("11",cstr(m_iTuki))%> >11
                        <option value="12" <%=f_Selected("12",cstr(m_iTuki))%> >12
                        <option value="1"  <%=f_Selected("1" ,cstr(m_iTuki))%> >1
                        <option value="2"  <%=f_Selected("2" ,cstr(m_iTuki))%> >2
                        <option value="3"  <%=f_Selected("3" ,cstr(m_iTuki))%> >3
                    </select></td>
				<td>月</td>
                </td>
                <td nowrap >

                    <%If m_iCnt < 0 Then%>
	                    <select name="cboDate"  DISABLED style="width:50px;">
                        <option value="">
                    <%Else%>
	                    <select name="cboDate"  style="width:50px;">
						<%For i = 0 To m_iCnt
                            %>
                            <option value="<%=m_sAryDay(i)%>" <%=f_Selected(m_sAryDay(i) ,m_sDate)%>><%=Day(m_sAryDay(i))%>
                            <%
                        Next
                    End If
                    %>
                    </select></td>
				<td>日</td>
				<td valign="bottom" align="right">
	            <input class="button" type="button" onclick="javascript:f_Search();" value="　表　示　">
				</tr>
            </table>
        </td>
    </tr>
    </table>

    </center>
    </form>
    </body>
    </html>
<%
End Sub
%>
