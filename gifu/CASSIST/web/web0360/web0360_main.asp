<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 部活動部員一覧
' ﾌﾟﾛｸﾞﾗﾑID : web/web0360/web0360_main.asp
' 機      能: 部員を表示
'-------------------------------------------------------------------------
' 引      数:   txtClubCd		:部活CD
'               KYOKAN_CD       '//教官CD
'
' 引      渡:	txtMode			:処理モード
'               GAKUSEI_NO		:学生NO
'
' 説      明:
'           ■初期表示
'               空白ページを表示
'           ■表示ボタンが押された場合
'               ・選択された部活動部員を一覧表示する
'               ・ログイン者が顧問の場合は、登録、削除が可能となる
'               ・顧問以外の時は、参照のみとする
'-------------------------------------------------------------------------
' 作      成: 2001/08/22 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	Public m_iSyoriNen			'//年度
	Public m_iKyokanCd			'//教官ｺｰﾄﾞ
	Public m_sClubCd			'//クラブCD

    Public m_bKomon				'//顧問かどうかを判別するﾌﾗｸﾞ
    Public m_bUpdate_OK			'//更新制御ﾌﾗｸﾞ
    Public m_sKomonKyokanStr	'//顧問教官CD

    'ﾚｺｰﾄﾞセット
    Public m_Rs					'//部員一覧ﾚｺｰﾄﾞｾｯﾄ（入部者）
    Public m_Rs2				'//部員一覧ﾚｺｰﾄﾞｾｯﾄ（退部者）
    Public m_iRsCnt				'//ﾚｺｰﾄﾞカウント
    Public m_bGetMember			'//部員取得ﾌﾗｸﾞ

	Dim	gTaibuFlg				'// 退部者ﾌﾗｸﾞ
	Dim	gTaibubi				'// 退部日
	Dim gFieldName				'// 

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
    w_sMsgTitle="部活動部員一覧"
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
'Call s_DebugPrint()

		'//部活名、顧問教官名情報取得
		w_iRet = f_GetClubInfo()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//部員の取得
        w_iRet = f_GetMember()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'//顧問教官以外のUSERは参照のみとする
        Call s_SetViewInfo()

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
    Call gf_closeObject(m_Rs)

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

    m_iSyoriNen       = ""
    m_iKyokanCd       = ""
	m_sClubCd           = ""
	m_sKomonKyokanStr = ""
	m_bKomon          = False

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen  = Session("NENDO")
    m_iKyokanCd  = Session("KYOKAN_CD")
	m_sClubCd    = Request("txtClubCd")
	Session("HyoujiNendo") = m_iSyoriNen

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
    response.write "m_sClubCd    = " & m_sClubCd   & "<br>"

End Sub

'********************************************************************************
'*  [機能]  教官以外のUSERは参照のみとする
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetViewInfo()

	m_bUpdate_OK = False

	'//顧問の教官は登録・削除が可能
	If m_bKomon = True Then
		m_bUpdate_OK = True
	End If

End Sub

'********************************************************************************
'*  [機能]  クラブ情報取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  部活名、顧問教官名情報を取得
'********************************************************************************
Function f_GetClubInfo()

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubInfo = 1

	Do

		'//部活動情報取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD1, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD2, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD3, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD4, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD5, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUJYOKYO_KBN"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD='" &  m_sClubCd & "'"

'response.write w_sSQL & "<br>"
'response.end
		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_GetClubInfo = 99
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//部活名
			m_sClubName = rs("M17_BUKATUDOMEI")

			'//①ログイン者が顧問教官かどうかを判断し、
			'//②顧問教官CDをカンマ区切りで保存する
			For i = 1 To 5

				'//①ログイン者が顧問教官かどうかを判断
				If trim(gf_SetNull2String(rs("M17_KOMON_CD" & i))) = trim(m_iKyokanCd) Then
					m_bKomon = True
				End If

				'//②顧問教官CDをカンマ区切りで保存する
				If trim(gf_SetNull2String(rs("M17_KOMON_CD" & i))) <> "" Then
					If m_sKomonKyokanStr = "" Then
						m_sKomonKyokanStr = rs("M17_KOMON_CD" & i)
					Else
						m_sKomonKyokanStr = m_sKomonKyokanStr & "," & rs("M17_KOMON_CD" & i)
					End If
				End If

			Next

		End If

		'//正常終了
		f_GetClubInfo = 0
		Exit Do
	Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  部員一覧情報取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  入部者、退部者の順番に並べるために、分けてデータを取得
'********************************************************************************
Function f_GetMember()

	Dim w_iRet
	Dim w_sSQL
	Dim rs
	Dim w_lCnt1
	Dim w_lCnt2

	On Error Resume Next
	Err.Clear

	f_GetMember = 1

	Do

		'//部活動情報取得（入部者取得）
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_NYUBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_NYUBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI.T11_SIMEI"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO"
'		w_sSql = w_sSql & vbCrLf & "  AND T13_GAKU_NEN.T13_NENDO = T11_GAKUSEKI.T11_NYUNENDO + T13_GAKU_NEN.T13_GAKUNEN - 1"
		w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen

		w_sSql = w_sSql & vbCrLf & "  AND (  (T13_GAKU_NEN.T13_CLUB_1='" & m_sClubCd & "' AND T13_GAKU_NEN.T13_CLUB_1_FLG = 1)"
		w_sSql = w_sSql & vbCrLf & "      Or (T13_GAKU_NEN.T13_CLUB_2='" & m_sClubCd & "' AND T13_GAKU_NEN.T13_CLUB_2_FLG = 1)"
		w_sSql = w_sSql & vbCrLf & "      )"

		'w_sSql = w_sSql & vbCrLf & "  AND ((T13_GAKU_NEN.T13_CLUB_1=" & m_sClubCd & " AND T13_GAKU_NEN.T13_CLUB_1_FLG in (1,2))"
		'w_sSql = w_sSql & vbCrLf & "      Or (T13_GAKU_NEN.T13_CLUB_2=" & m_sClubCd & " AND T13_GAKU_NEN.T13_CLUB_2_FLG in (1,2))"
		'w_sSql = w_sSql & vbCrLf & "      )"
		w_sSql = w_sSql & vbCrLf & " ORDER BY "
		w_sSql = w_sSql & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write "111111111111<br>" & w_sSQL & "<br>"
		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
'response.end
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_GetMember = 99
			Exit Do
		End If

		'//部活動情報取得（退部者取得）
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_NYUBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_NYUBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI.T11_SIMEI"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO"
'		w_sSql = w_sSql & vbCrLf & "  AND T13_GAKU_NEN.T13_NENDO = T11_GAKUSEKI.T11_NYUNENDO + T13_GAKU_NEN.T13_GAKUNEN - 1"
		w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen

		w_sSql = w_sSql & vbCrLf & "  AND ((T13_GAKU_NEN.T13_CLUB_1='" & m_sClubCd & "' AND T13_GAKU_NEN.T13_CLUB_1_FLG = 2)"
		w_sSql = w_sSql & vbCrLf & "      Or (T13_GAKU_NEN.T13_CLUB_2='" & m_sClubCd & "' AND T13_GAKU_NEN.T13_CLUB_2_FLG = 2)"
		w_sSql = w_sSql & vbCrLf & "      )"

		'w_sSql = w_sSql & vbCrLf & "  AND ((T13_GAKU_NEN.T13_CLUB_1=" & m_sClubCd & " AND T13_GAKU_NEN.T13_CLUB_1_FLG in (1,2))"
		'w_sSql = w_sSql & vbCrLf & "      Or (T13_GAKU_NEN.T13_CLUB_2=" & m_sClubCd & " AND T13_GAKU_NEN.T13_CLUB_2_FLG in (1,2))"
		'w_sSql = w_sSql & vbCrLf & "      )"
		w_sSql = w_sSql & vbCrLf & " ORDER BY "
		w_sSql = w_sSql & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write "222222222222<br>" & w_sSQL & "<br>"
		w_iRet = gf_GetRecordset(m_Rs2, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
'response.end
			f_GetMember = 99
			Exit Do
		End If

        m_iRsCnt = 0

		w_lCnt1 = 0
		w_lCnt2 = 0

		'入部者数
        If m_Rs.EOF = False Then
            w_lCnt1 = gf_GetRsCount(m_Rs)
        End If

		'退部者数
        If m_Rs2.EOF = False Then
            w_lCnt2 = gf_GetRsCount(m_Rs2)
        End If

		'//ﾚｺｰﾄﾞカウント取得
        '//件数を取得
        m_iRsCnt = w_lCnt1 + w_lCnt2
		
		If m_iRsCnt > 0 Then
			m_bGetMember = True
		Else
			m_bGetMember = False
		End If

        'If m_Rs.EOF = False Then
        '    m_iRsCnt = gf_GetRsCount(m_Rs)
		'	m_bGetMember = True
		'Else
		'	m_bGetMember = False
        'End If

		'//正常終了
		f_GetMember = 0
		Exit Do
	Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function


'********************************************************************************
'*  [機能]  クラス情報を取得
'*  [引数]  p_iGakuNen:学年,p_iClassNo:クラスNO
'*  [戻値]  f_GetClassName:クラス略称
'*  [説明]  
'********************************************************************************
Function f_GetClassName(p_iGakuNen,p_iClassNo)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetClassName = ""
	w_sClassName = ""

    Do
        'クラスマスタよりデータを取得
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_CLASSRYAKU"
        w_sSql = w_sSql & vbCrLf & "  ,M05_CLASS.M05_GAKKA_CD"
        w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN= " & p_iGakuNen
        w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_CLASSNO= "   & p_iClassNo

'response.write w_sSQL & "<br>"

		'//データ取得
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

        If rs.EOF = False Then
            w_sClassName = rs("M05_CLASSRYAKU")
            'w_sGakkaCd = rs("M05_GAKKA_CD")
        End If

        Exit Do
    Loop

	'//戻り値ｾｯﾄ
	f_GetClassName = w_sClassName

	'//ﾚｺｰﾄﾞCLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  顧問教官を表示(HTML書き出し)
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ShowKomon()
	Dim i
	Dim w_Ary
	Dim w_sRowspan

	Do

		'//顧問教官が設定されていない場合
		If LenB(m_sKomonKyokanStr)=0 Then%>
			<table class="hyo" border="1">
				<tr>
					<th nowrap class="header" width="40"  align="center" >顧問</th>
					<td nowrap class="detail" width="120"  align="center">―</td>
				</tr>
			</table>
			<%
		Else

			'//顧問教官CD(CSV形式)取得
			w_Ary = split(m_sKomonKyokanStr,",")
			iMax = UBound(w_Ary)

			If iMax >= 3 Then
				w_sRowspan="rowspan=2"
			Else
				w_sRowspan="rowspan=1"
			End If

			'//ヘッダ書き出し
			%>
			<table class="hyo" border="1">
				<tr>
					<th nowrap class="header" width="50"  align="center" <%=w_sRowspan%> >顧問</th>
			<%For i = 0 To iMax%>

				<%If i = 3 Then%>
					</tr><tr>
				<%End If%>
					<td nowrap class="detail" width="120"  align="center"><%=gf_GetKyokanNm(m_iSyoriNen,w_Ary(i))%></td>
			<%Next

			'//改行ありの場合、のこりの空白行を表示
			If i-1 >= 3 Then
				For j = 1 To 6-i%>
					<td nowrap class="detail" width="100"  align="center"><br></td>
				<%
				Next
			End If

			%>
				</tr>
			</table>
		<%
		End If

		Exit Do
	Loop

End Sub


'********************************************************************************
'*  [機能]  退部者を取得
'********************************************************************************
Sub s_Taibu(s_NyuTai_Flg)

	Dim wTaibuFlg

	'// 退部者フラグ
	'入部者の場合
	if s_NyuTai_Flg = 1 Then		
		wC1  = m_Rs("T13_CLUB_1")
		wC2  = m_Rs("T13_CLUB_2")
		wC1F = m_Rs("T13_CLUB_1_FLG")
		wC2F = m_Rs("T13_CLUB_2_FLG")
	'退部者の場合
	else
		wC1  = m_Rs2("T13_CLUB_1")
		wC2  = m_Rs2("T13_CLUB_2")
		wC1F = m_Rs2("T13_CLUB_1_FLG")
		wC2F = m_Rs2("T13_CLUB_2_FLG")
	end if

	wTaibuFlg1 = False
	wTaibuFlg2 = False
	gTaibuFlg  = False
	gTaibubi   = ""

	if Not gf_IsNull(wC1) Then
'response.write "1111111  " & CStr(wC1) & " = " & CStr(m_sClubCd) & " = " & m_Rs("T13_GAKUSEI_NO") & "<br>"
		if (CStr(wC1) = CStr(m_sClubCd)) then

			if (Cint(wC1F) = 2) then
				wTaibuFlg1 = True
				
				'入部者の場合
				if s_NyuTai_Flg = 1 Then
					gTaibubi = m_Rs("T13_CLUB_1_TAIBI")
				'退部者の場合
				else
					gTaibubi = m_Rs2("T13_CLUB_1_TAIBI")
				end if
				gFieldName = 1		'// クラブ1ってこと
			End if

			If (Cint(wC1F) = 1) then
				gFieldName = 1		'// クラブ1ってこと
			End if

		End If

	End if

	if Not gf_IsNull(wC2) then
'response.write "2222222  " & CStr(wC2) & " = " & CStr(m_sClubCd) & " = " & m_Rs("T13_GAKUSEI_NO") & "<br>"
		if (CStr(wC2) = CStr(m_sClubCd)) then

			if (Cint(wC2F) = 2) then
				wTaibuFlg2 = True
				'入部者の場合
				if s_NyuTai_Flg = 1 Then
					gTaibubi = m_Rs("T13_CLUB_2_TAIBI")
				'退部者の場合
				else
					gTaibubi = m_Rs2("T13_CLUB_2_TAIBI")
				end if

				gFieldName = 2		'// クラブ2ってこと
			End if

			If (Cint(wC2F) = 1) then
				gFieldName = 2		'// クラブ2ってこと
			End if

		End if
	End if

	if wTaibuFlg1 OR wTaibuFlg2 then
		gTaibuFlg = True
	End if


End SUb



'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	dim w_NyuBi '入部日

%>

    <html>
    <head>
    <title>部活動部員一覧</title>
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
    //  [機能]  登録ボタンが押されたとき,登録画面を表示
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){

        //リスト情報をsubmit
		//上フレーム
		parent.topFrame.location.href="./web0360_insTop.asp?txtClubCd=<%=m_sClubCd%>"

		//下フレーム
		parent.main.location.href="./default3.asp?txtClubCd=<%=m_sClubCd%>"

        return;

    }

    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Cancel(){
        //空白ページを表示
        parent.document.location.href="default2.asp"
    }

	//************************************************************
	//  [機能]  退部ボタンが押されたとき
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//************************************************************
	function f_Taibu(){

		var i
		var w_bCheck = 1

		//チェック欄数を取得
		var iMax = document.frm.chkMax.value
		if (iMax==0){
			//alert("No Avairable")
			return 1;
		}

		// 入力値のﾁｪｯｸ
		iRet = f_CheckData();
		if( iRet != 0 ){
			return;
		}

		//削除が設定されている場合のメッセージ表示
		if(iMax==1){
			if(obj2.value == "<%=C_DELETE0%>"){
	            window.alert("データの削除が設定されています。" + "\n" + "実行すると入退部の履歴も削除されます。");
			}
		}else{
			for (i = 0; i < iMax; i++) {
				if(obj2[i].value == "<%=C_DELETE0%>"){
		            window.alert("データの削除が設定されています。" + "\n" + "実行すると入退部の履歴も削除されます。");
					break;
				}
			}
		}

		if (!confirm("更新してもよろしいですか？")) {
			document.frm.hidTaibubi.value = "";
			return ;
		}

		//リスト情報をsubmit
		document.frm.txtMode.value = "DELETE";
		document.frm.target = "main";
		document.frm.action = "./web0360_edt.asp"
		document.frm.submit();
		return;
	}

    //************************************************************
    //  [機能]  チェック欄がチェックされているか
    //  [引数]  なし
    //  [戻値]  0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
    //************************************************************
    function f_CheckData(p_bChk) {

		obj  = eval(document.frm.txtTaibubi);
		obj2 = eval(document.frm.txtNyububiC);
		objTaibu = eval(document.frm.hidTaibuFlg);

		//チェック欄数を取得
		var iMax = document.frm.chkMax.value
		if (iMax==0){
			//alert("No Avairable")
			return 1;
		}

		if(iMax==1){

			// 入部日チェック
			if(obj2.value == ""){
				obj2.value = "<%= gf_YYYY_MM_DD(date(),"/") %>";
			}else{
				if(obj2.value != "<%=C_DELETE0%>"){
					if( chk_dateSplit(obj2.value) == 1 ){
					    obj2.focus();
					    return 1;
					}
				}
			}

			// まだ入部中の人が対象
			if(objTaibu == "False") {
				// 削除指定がされていない場合
				if(obj2.value != "<%=C_DELETE0%>"){
					// 退部日が入力されていたらフラグにチェックがあるかチェックする
					if(obj.value != ""){
						if(document.frm.GAKUSEI_NO.checked==false){
							alert("退部日が入力されています。退部登録する場合は、退部欄にチェックを付けてください。")
							document.frm.GAKUSEI_NO.focus();
							return 1;
						}
					}
				}
			}

			if(document.frm.GAKUSEI_NO.checked==false){
//				alert("退部登録する生徒が選択されていません")
//				return 1;
			}else{
				if(obj.value == ""){
					obj.value = "<%= gf_YYYY_MM_DD(date(),"/") %>";
				}else{
					// 日付チェック
					if( chk_dateSplit(obj.value) == 1 ){
					    obj.focus();
					    return 1;
					}
				}

				// 日付大小チェック
		        if( DateParse(obj.value,obj2.value) >= 1){
		            window.alert("開始日と終了日を正しく入力してください");
		            obj.focus();
		            return 1;
		        }
				document.frm.hidTaibubi.value = obj.value;
			}

		}else{

			var i
			var w_bCheck = 1
			for (i = 0; i < iMax; i++) {

				// 入部日チェック
				if(obj2[i].value == ""){
					obj2[i].value = "<%= gf_YYYY_MM_DD(date(),"/") %>";
				}else{
					if(obj2[i].value != "<%=C_DELETE0%>"){
						if( chk_dateSplit(obj2[i].value) == 1 ){
						    obj2[i].focus();
						    return 1;
						}
					}
				}

				// まだ入部中の人が対象
				if(objTaibu[i].value == "False") {
					// 削除指定がされていない場合
					if(obj2[i].value != "<%=C_DELETE0%>"){
						// 退部日が入力されていたらフラグにチェックがあるかチェックする
						if(obj[i].value != ""){
							if(document.frm.GAKUSEI_NO[i].checked==false){
								alert("退部日が入力されています。退部登録する場合は、退部欄にチェックを付けてください。")
								document.frm.GAKUSEI_NO[i].focus();
								return 1;
							}
						}
					}
				}

				if(document.frm.GAKUSEI_NO[i].checked==true){
					w_bCheck = 0

					if(obj[i].value == ""){
						obj[i].value = "<%= gf_YYYY_MM_DD(date(),"/") %>";
					}else{
						// 日付チェック
						if( chk_dateSplit(obj[i].value) == 1 ){
						    obj[i].focus();
						    return 1;
						}
					}

					// 日付大小チェック
			        if( DateParse(obj[i].value,obj2[i].value) >= 1){
			            window.alert("開始日と終了日を正しく入力してください");
			            obj[i].focus();
			            return 1;
			        }

					if(document.frm.hidTaibubi.value != ""){
						document.frm.hidTaibubi.value = document.frm.hidTaibubi.value + ",";
					}
					document.frm.hidTaibubi.value = document.frm.hidTaibubi.value + obj[i].value;

				}
			};

			if(w_bCheck == 1){
//				alert("退部登録する生徒が選択されていません")
//				return 1;
			};
		};
        return 0;
    }
    
    //************************************************************
    //  [機能]  詳細ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_detail(pGAKUSEI_NO){

			url = "/cassist/gak/gak0300/kojin.asp?hidGAKUSEI_NO=" + pGAKUSEI_NO;
			w   = 700;
			h   = 630;

			wn  = "SubWindow";
			opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
			if (w > 0)
				opt = opt + ",width=" + w;
			if (h > 0)
				opt = opt + ",height=" + h;
			newWin = window.open(url, wn, opt);

//		document.frm.hidGAKUSEI_NO.value = pGAKUSEI_NO;
//		document.forms[0].submit();
    }

    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">
    <center>
    <form name="frm" method="post">
	<br>

	<%
	Do 

		'=====================
		'//顧問教官表示
		'=====================
		Call s_ShowKomon()
		%>

		<br>

		<%
		'=====================
		'//登録、削除ボタン
		'=====================
		'//顧問の教官の場合
		If m_bKomon = True Then

			'//部員がいない場合
			If m_bGetMember = false Then%>
				<span class="msg">部員が登録されていません</span>
				<br><br>
				<br>
			<%End If%>

			<table><tr><td>
				<span class="msg">
				＊入部者を登録する際は「入部登録」ボタンをクリックしてください。<br>
				<%If m_bGetMember Then%>
				＊退部者を登録する際は退部する学生の退部欄をチェック(複数可)し、<BR>
				　&nbsp;退部日を入力の上「更新」ボタンをクリックしてください。（空白の場合は処理日が入ります。）<BR>
				＊削除（履歴データから消す）する場合は、入部日欄に「<%=C_DELETE0%>」（10桁）を入力して<BR>
				　&nbsp;「更　新」ボタンをクリックしてください。<BR>
				＊日付を修正する場合は、入、退部日欄を変更して「更　新」ボタンをクリックしてください。<BR>
				<%End If%>
				</span>
			</td></tr></table>

	        <table>
				<tr>
				<td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="入部登録"></td>
				<%If m_bGetMember = True Then%><td ><input class="button" type="button" onclick="javascript:f_Taibu();" value="　更　新　"></td><%End If%>
				</tr>
	        </table>

		<%End If%>

		<%
		'=====================
		'//リスト部表示
		'=====================

		'//部員がいない場合
		If m_bGetMember = false Then
			If m_bKomon = false Then%>
				<br><br>
				<span class="msg">部員が登録されていません</span>
			<%End If
			Exit Do
		End If
		%>


		<table>
			<tr><td valign="top" align="right">

			<table><tr><td>
				<span class="msg">
				<%If m_bKomon = True Then%>
					入力例：（2001/01/01 又は <%=C_DELETE0%>）<BR>
					日付が未入力の場合、自動的に現在の日付が入ります。<BR>
				<%End If%>
				</span>
			</td></tr></table>

			<table class=hyo border="1" bgcolor="#FFFFFF">
				<!--ヘッダ-->
				<tr>
					<%If m_bKomon = True Then%><th nowrap class="header" width="45"  align="center">退部</th><%End If%>
					<th nowrap class="header" width="40"  align="center">クラス</th>
					<th nowrap class="header" width="40"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
					<th nowrap class="header" width="150" align="center">氏名</th>
					<th nowrap class="header" align="center">入部日</th>
					<% If m_bKomon = True Then%>
						<th nowrap class="header" align="center">退部日</th>
					<% End if %>
				</tr>
		<%
		'//改行カウント
		'w_iCnt = INT(m_iRsCnt/2 + 0.9)
'--- 入部者表示 ------------------------------------------------------------------------------------------------------------------------------
		Do Until m_Rs.EOF

			'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
			Call gs_cellPtn(w_Class)
			i = i + 1
			
			'// 入部日を変数に代入
			If m_sClubCd = m_Rs("T13_CLUB_1") then 
				w_NyuBi = m_Rs("T13_CLUB_1_NYUBI")
			Else
				w_NyuBi = m_Rs("T13_CLUB_2_NYUBI")
			End If

			'// 退部者を取得
			Call s_Taibu(1)
			%>
				<tr>
					<%If m_bKomon = True Then%>
						<td nowrap class="<%=w_Class%>" width="45"  align="center">
							<% If gTaibuFlg Then %>
								退部者<input type="hidden" name="GAKUSEI_NO" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<% Else %>
								<input type="checkbox" name="GAKUSEI_NO" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<% End if %>
							<input type="hidden" name="hidTaibuFlg"  value="<%=gTaibuFlg%>">
							<input type="hidden" name="hidGakuseiNo" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<input type="hidden" name="hidFieldName" value="<%=gFieldName%>">
						</td>
					<%End If%>
					<td nowrap class="<%=w_Class%>" align="center"><%=m_Rs("T13_GAKUNEN")%>-<%=f_GetClassName(m_Rs("T13_GAKUNEN"),m_Rs("T13_CLASS"))%><br></td>
					<td nowrap class="<%=w_Class%>" align="left"  ><%=m_Rs("T13_GAKUSEKI_NO")%><br></td>
					<td nowrap class="<%=w_Class%>" align="left"  ><a href="#" onClick="f_detail('<%=m_Rs("T13_GAKUSEI_NO")%>')"><%=m_Rs("T11_SIMEI")%></a><br></td>
					<%If m_bKomon = True Then%>
						<td nowrap class="<%=w_Class%>" align="center"  >
						<input type="text" style="width:80px;" name="txtNyububiC" maxlength="10" value="<%=w_NyuBi%>" id="id_Txt1<%=i-1%>">&nbsp;
						<input type="button" class="button" onclick="fcalender('id_Txt2<%=i-1%>')" value="選択">
					<%Else%>
						<td nowrap class="<%=w_Class%>" align="center"  >
						<%=w_NyuBi%>&nbsp;
					<%End If%>
					</td>

					<% If m_bKomon = True Then %>
						<td nowrap class="<%=w_Class%>" align="center"><input type="text" style="width:80px;" name="txtTaibubi"  maxlength="10" id="id_Txt2<%=i-1%>" value="<%=gTaibubi%>">&nbsp;<input type="button" class="button" onclick="fcalender('id_Txt2<%=i-1%>')" value="選択"></td>
					<% End if %>
				</tr>

				<%m_Rs.MoveNext%>
		<%Loop
'---------------------------------------------------------------------------------------------------------------------------------
		%>
		<%
'--- 退部者表示 ------------------------------------------------------------------------------------------------------------------------------
		'顧問教官の場合のみ退部者を表示
		If m_bKomon = True Then
			Do Until m_Rs2.EOF

				'//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
				Call gs_cellPtn(w_Class)
				i = i + 1
				'// 入部日を変数に代入
				If m_sClubCd = m_Rs2("T13_CLUB_1") then 
					w_NyuBi = m_Rs2("T13_CLUB_1_NYUBI")
				Else
					w_NyuBi = m_Rs2("T13_CLUB_2_NYUBI")
				End If

				'// 退部者を取得
				Call s_Taibu(2)
				%>
					<tr>
						<%If m_bKomon = True Then%>
							<td nowrap class="<%=w_Class%>" width="45"  align="center">
								<% If gTaibuFlg Then %>
									退部者<input type="hidden" name="GAKUSEI_NO" value="<%=m_Rs2("T13_GAKUSEI_NO")%>">
								<% Else %>
									<input type="checkbox" name="GAKUSEI_NO" value="<%=m_Rs2("T13_GAKUSEI_NO")%>">
								<% End if %>
								<input type="hidden" name="hidTaibuFlg"  value="<%=gTaibuFlg%>">
								<input type="hidden" name="hidGakuseiNo" value="<%=m_Rs2("T13_GAKUSEI_NO")%>">
								<input type="hidden" name="hidFieldName" value="<%=gFieldName%>">
							</td>
						<%End If%>
						<td nowrap class="<%=w_Class%>" align="center"><%=m_Rs2("T13_GAKUNEN")%>-<%=f_GetClassName(m_Rs2("T13_GAKUNEN"),m_Rs2("T13_CLASS"))%><br></td>
						<td nowrap class="<%=w_Class%>" align="left"  ><%=m_Rs2("T13_GAKUSEKI_NO")%><br></td>
						<td nowrap class="<%=w_Class%>" align="left"  ><a href="#" onClick="f_detail('<%=m_Rs2("T13_GAKUSEI_NO")%>')"><%=m_Rs2("T11_SIMEI")%></a><br></td>
					<%If m_bKomon = True Then%>
						<td nowrap class="<%=w_Class%>" align="center"  >
						<input type="text" style="width:80px;" name="txtNyububiC" maxlength="10" value="<%=w_NyuBi%>" id="id_Txt1<%=i-1%>">&nbsp;
						<input type="button" class="button" onclick="fcalender('id_Txt2<%=i-1%>')" value="選択">
					<%Else%>
						<td nowrap class="<%=w_Class%>" align="center"  >
						<%=w_NyuBi%>&nbsp;
					<%End If%>
					</td>
						<% If m_bKomon = True Then %>
							<td nowrap class="<%=w_Class%>" align="center"><input type="text" style="width:80px;" name="txtTaibubi"  maxlength="10" id="id_Txt2<%=i-1%>" value="<%=gTaibubi%>">&nbsp;<input type="button" class="button" onclick="fcalender('id_Txt2<%=i-1%>')" value="選択"></td>
						<% End if %>
					</tr>

					<%m_Rs2.MoveNext%>
			<%Loop
		End If
'---------------------------------------------------------------------------------------------------------------------------------
		%>

				</table>

				<table><tr><td>
					<span class="msg">
					<%If m_bKomon = True Then%>
						入力例：（2001/01/01 又は <%=C_DELETE0%>）<BR>
						日付が未入力の場合、自動的に現在の日付が入ります。<BR>
					<%End If%>
					</span>
				</td></tr></table>

			</td></tr>
		</table>
		<br>

		<%
		'//顧問の教官の場合
		If m_bKomon = True Then%>
			<table>
				<tr>
					<td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="入部登録"></td>
					<td ><input class="button" type="button" onclick="javascript:f_Taibu();" value="　更　新　"></td>
				</tr>
			</table>
		<%End If%>

		<%Set m_Rs  = Nothing%>
		<%Set m_Rs2 = Nothing%>
		<%Exit Do%>
	<%Loop%>

	<!--値渡し-->
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
	<input type="hidden" name="txtClubCd" value="<%=m_sClubCd%>">
	<input type="hidden" name="chkMax" value="<%=i%>">
	<input type="hidden" name="hidTaibubi">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

