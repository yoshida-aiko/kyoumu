<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 個人履修選択科目決定
' ﾌﾟﾛｸﾞﾗﾑID : web/web0340/web0340_main.asp
' 機      能: 下ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/25 前田
' 変      更: 2001/08/28 伊藤公子 ヘッダ部切り離し対応
' 変      更: 2015/08/19 清本 1年間番号の幅を50→70に変更
' 変      更: 2015/08/27 藤林 科目のデータ取得方法変更(T15_RISYU→T16_RISYU_KOJIN)
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
    Public  m_Krs           '科目用レコードセット
    Public  m_Grs           '学生用レコードセット
    Public  m_KSrs          '科目数のレコードセット
'    Public  m_rs            'レコードセット
    Dim     m_iNendo        '//年度
    Dim     m_sKyokanCd     '//教官コード
    Dim     m_sGakunen      '//学年
    Dim     m_sClass        '//クラス
    Dim     m_sKBN          '//区分
    Dim     m_sGRP          '//グループ区分
    Dim     m_KrCnt         '//科目のレコードカウント
    Dim     m_KSrCnt        '//科目数のレコードカウント
    Dim     m_GrCnt         '//学生のレコードカウント
    Dim     m_cell          '配色の設定
    Dim     m_iSTani        
	Dim		m_sRisyuJotai	'履修状態フラグ add 2001/10/25
    Dim     i               
    Dim     j               
    Dim     k               

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

    Dim w_iRet              '// 戻り値

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="連絡事項登録"
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

        '//科目の情報取得
        w_iRet = f_KamokuData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

		If m_Krs.EOF Then
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

        '//区分,選択種別の総合単位取得
        w_iRet = f_Tani()
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
    Call gf_closeObject(m_Krs)
    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_Grs)
    '// 終了処理
    Call gs_CloseDatabase()
    
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
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
    m_sKBN      = Cint(request("txtKBN"))
    m_sGRP      = Cint(request("txtGRP"))
    m_iDsp      = C_PAGE_LINE

End Sub

Function f_KamokuData()
'******************************************************************
'機　　能：科目のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear
    f_KamokuData = 1

    Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

        '//科目のデータ取得
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT DISTINCT "
        m_sSQL = m_sSQL & vbCrLf & "     T16_KAMOKUMEI,T16_KAMOKU_CD,T16_HAITOTANI"
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T16_RISYU_KOJIN "
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T16_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_HISSEN_KBN = " & C_HISSEN_SEN & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_HAITOTANI <> " & C_T15_HAITO & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_GRP = " & m_sGRP & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_KAMOKU_KBN = " & m_sKBN & " "
        m_sSQL = m_sSQL & vbCrLf & " AND EXISTS ( SELECT 'X' "
        m_sSQL = m_sSQL & vbCrLf & "              FROM  "
        m_sSQL = m_sSQL & vbCrLf & "                    T11_GAKUSEKI,T13_GAKU_NEN "
        m_sSQL = m_sSQL & vbCrLf & "              WHERE  "
        m_sSQL = m_sSQL & vbCrLf & "                    T13_NENDO = T16_NENDO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_GAKUSEI_NO = T16_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_GAKUSEI_NO = T11_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T11_NYUNENDO = " & w_iNyuNendo & " "
        m_sSQL = m_sSQL & vbCrLf & "             ) "

'response.write m_sSQL & "<BR>"

        Set m_Krs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Krs, m_sSQL,m_iDsp)

'response.write "w_iRet = " & w_iRet & "<BR>"
'response.write m_Krs.EOF & "<BR>"

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
    m_KrCnt=gf_GetRsCount(m_Krs)

    f_KamokuData = 0

    Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
'    If m_bErrFlg = True Then
'        w_sMsg = gf_GetErrMsg()
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
'		response.end
'    End If

End Function

Function f_GakuseiData()
'******************************************************************
'機　　能：学生のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_GakuseiData = 1

    Do
        '//学生のデータ取得
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT "
        m_sSQL = m_sSQL & vbCrLf & "     T13_GAKUSEKI_NO,T11_SIMEI,T13_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T13_GAKU_NEN,T11_GAKUSEKI "
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T13_GAKUSEI_NO = T11_GAKUSEI_NO(+) "
        m_sSQL = m_sSQL & vbCrLf & " AND T13_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T13_GAKUNEN = " & m_sGakunen & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO & " "
        m_sSQL = m_sSQL & vbCrLf & " ORDER BY T13_GAKUSEKI_NO "

'response.write m_sSQL & "<BR>"

        Set m_Grs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Grs, m_sSQL,m_iDsp)
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

Function f_Tani()
'******************************************************************
'機　　能：区分,選択種別の総合単位取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
Dim w_iNyuNendo,w_rs

    On Error Resume Next
    Err.Clear
    f_Tani = 1

    Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

        '//区分,選択種別の総合単位取得
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT "
        m_sSQL = m_sSQL & "     T18_GAKUNEN_SU"&Cint(m_sGakunen)&" "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T18_SELECTSYUBETU "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T18_NYUNENDO = " & w_iNyuNendo & " "
        m_sSQL = m_sSQL & " AND T18_GRP = " & m_sGRP & " "
        m_sSQL = m_sSQL & " AND T18_GAKKA_CD = " & m_sClass & " "
        m_sSQL = m_sSQL & " AND T18_KAMOKUSYU_CD = " & m_sKBN & " "
        m_sSQL = m_sSQL & " AND T18_GAKUNEN_SEL = " & C_T18_SEL_GAKU & " "

'response.write m_sSQL & "<BR>"
        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

        m_iSTani = w_rs("T18_GAKUNEN_SU"&Cint(m_sGakunen)&"")

	    f_Tani = 0

    	Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
'    If m_bErrFlg = True Then
'        w_sMsg = gf_GetErrMsg()
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
'		response.end
'    End If
    Call gf_closeObject(w_rs)

End Function

Function f_KibouData()
'******************************************************************
'機　　能：希望欄のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************
	Dim w_rs,w_sSQL
	
    On Error Resume Next
    Err.Clear
    
    f_KibouData = 1

    Do
        '//希望欄のデータ取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     T16_SELECT_FLG,T16_KIBOU_FLG "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T16_RISYU_KOJIN "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "     T16_NENDO = " & m_iNendo & " "
        w_sSQL = w_sSQL & " AND T16_GAKUSEI_NO = '" & m_Grs("T13_GAKUSEI_NO") & "' "
        w_sSQL = w_sSQL & " AND T16_KAMOKU_CD = '" & m_Krs("T16_KAMOKU_CD") & "' "

'        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

	If w_rs.EOF = false then
		If cint(m_sRisyuJotai) = C_K_RIS_MAE then 

			'確定処理前-----------------------------------------------------
			'決定している場合
	        If Cint(gf_SetNull2Zero(w_rs("T16_SELECT_FLG"))) = C_SENTAKU_YES Then%>
		        <td class=<%=m_cell%>  width="88">
		        <input type=button class=<%=m_cell%> name=button<%=k%>_<%=j%> value="○" onclick="javascript:f_Chenge(<%=k%>,<%=j%>)" style="text-align:center">
		        <input type=hidden name=MAE<%=k%>_<%=j%> value="○">
		        <input type=hidden name=ATO<%=k%>_<%=j%> value="○">
		        <input type=hidden name=KibouFLG<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
		        </td>
	        <%Else
				'希望している場合
				If Cint(gf_SetNull2Zero(w_rs("T16_KIBOU_FLG"))) = 0 Then%>
			        <td class=<%=m_cell%>   width="88">
			        <input type=button class=<%=m_cell%> name=button<%=k%>_<%=j%> value="" onclick="javascript:f_Chenge(<%=k%>,<%=j%>)" style="text-align:center">
			        <input type=hidden name=MAE<%=k%>_<%=j%> value="">
			        <input type=hidden name=ATO<%=k%>_<%=j%> value="">
			        <input type=hidden name=KibouFLG<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
			        </td>

				<%Else
					'何もない場合
				%>
			        <td class=<%=m_cell%>   width="88">
			        <input type=button class=<%=m_cell%> name=button<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>' onclick="javascript:f_Chenge(<%=k%>,<%=j%>)" style="text-align:center">
			        <input type=hidden name=MAE<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
			        <input type=hidden name=ATO<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
			        <input type=hidden name=KibouFLG<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
			        </td>
				<%End If
	        End If
		Else
			'確定処理後-----------------------------------------------------
	        If Cint(gf_SetNull2Zero(w_rs("T16_SELECT_FLG"))) = C_SENTAKU_YES Then%>
		        <td class=<%=m_cell%>  width="88" align="center">○</td>
	        <%Else
					'何もない場合
			%>
			        <td class=<%=m_cell%> width="88" align="center">　</td>
			<%
			End If
		End If
	Else 

	  If cint(m_sRisyuJotai) = C_K_RIS_MAE then 
		'確定処理前-----------------------------------------------------
%>
        <td class=<%=m_cell%>   width="88">
        <input type=button class=<%=m_cell%> name=button<%=k%>_<%=j%> value="" onclick="javascript:f_Chenge(<%=k%>,<%=j%>)" style="text-align:center">
        <input type=hidden name=MAE<%=k%>_<%=j%> value="">
        <input type=hidden name=ATO<%=k%>_<%=j%> value="">
        <input type=hidden name=KibouFLG<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
        </td>
<%
	  Else 
		'確定処理後-----------------------------------------------------
%>
        <td class=<%=m_cell%> width="88" align="center">　</td>
<%
	  End If
	End If
	    f_KibouData = 0
	    Exit Do

    Loop

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(w_rs)
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If

End Function

Function f_KamokusuData()
'******************************************************************
'機　　能：科目数のデータ取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_KamokusuData = 1

    m_KSrCnt=""

    Do
        '//科目数のデータ取得
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT T16_KAMOKU_CD "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T16_RISYU_KOJIN ,T13_GAKU_NEN "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T16_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & " AND T16_NENDO = T13_NENDO "
        m_sSQL = m_sSQL & " AND T16_GAKUSEI_NO = T13_GAKUSEI_NO "
        m_sSQL = m_sSQL & " AND T16_HAITOGAKUNEN = T13_GAKUNEN "
        m_sSQL = m_sSQL & " AND T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & " AND T16_SELECT_FLG = " & C_SENTAKU_YES & " "
        m_sSQL = m_sSQL & " AND T16_KAMOKU_CD = '" & m_Krs("T16_KAMOKU_CD") & "' "
        m_sSQL = m_sSQL & " AND T16_HAITOGAKUNEN = " & m_sGakunen & " "

        Set m_KSrs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_KSrs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

	    m_KSrCnt=gf_GetRsCount(m_KSrs)

        If m_KSrs.EOF Then
            m_KSrCnt = "0"%>
	        <td class=disph><%=m_Krs("T16_KAMOKUMEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=m_KSrCnt%>" class="CELL2" name=Kamoku<%=i%> readonly></td>
        <%Else%>
	        <td class=disph><%=m_Krs("T16_KAMOKUMEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=m_KSrCnt%>" class="CELL2" name=Kamoku<%=i%> readonly></td>
        <%End If

	    f_KamokusuData = 0

	    Exit Do

    Loop


    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_KSrs)

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If

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
		parent.location.href = "white.asp?txtMsg=個人履修選択科目のデータがありません。"
        return;
    }
	//-->
	</SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>
    </center>
	<form name="frm" method="post">

	<input type="hidden" name="txtMsg" value="個人履修選択科目のデータがありません。">

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
Dim w_iKhalf
Dim w_iGhalf
Dim n

    On Error Resume Next
    Err.Clear

i = 0
k = 0
n = 0
%>
<HTML>


<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>個人履修選択科目決定</title>

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
        document.frm.action = "web0340_middle.asp"
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

        w_sBNm = eval("document.frm.button"+p_iS+"_"+p_iK);
        w_sMAE = eval("document.frm.MAE"+p_iS+"_"+p_iK);
        w_sATO = eval("document.frm.ATO"+p_iS+"_"+p_iK);

        //w_sKNm = eval("document.frm.Kamoku"+p_iK);
        w_sKNm = eval("parent.middle.document.frm.Kamoku"+p_iK);
        w_sKFLG = eval("document.frm.KibouFLG"+p_iS+"_"+p_iK);

        if(w_sBNm.value == "○"){
			if (w_sMAE.value == "○"){
					if (w_sKFLG.value == "0"){
			            w_sBNm.value = "";
			            w_sATO.value = "";
			            w_sKNm.value--;
					}else{
			            w_sBNm.value = w_sKFLG.value;
			            w_sATO.value = w_sKFLG.value;
			            w_sKNm.value--;
					}
			}else{
	            w_sBNm.value = w_sMAE.value;
	            w_sATO.value = w_sMAE.value;
	            w_sKNm.value--;
			}
        }else{
			if (w_sBNm.value == ""){
				if (w_sMAE.value == "○"){
		            w_sBNm.value = "○";
		            w_sATO.value = "○";
		            w_sKNm.value++;
				}else{
					if (w_sMAE.value == ""){
			            w_sBNm.value = "○";
			            w_sATO.value = "○";
			            w_sKNm.value++;
					}else{
			            w_sBNm.value = w_sMAE.value;
			            w_sATO.value = w_sMAE.value;
			            w_sKNm.value--;
					}
				}
			}else{
				if (w_sMAE.value == "○"){
		            w_sBNm.value = "○";
		            w_sATO.value = "○";
		            w_sKNm.value++;
				}else{
					if (w_sMAE.value == ""){
			            w_sBNm.value = "○";
			            w_sATO.value = "○";
			            w_sKNm.value++;
					}else{
			            w_sBNm.value = "○";
			            w_sATO.value = "○";
			            w_sKNm.value++;
					}
				}
			}
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
        parent.document.location.href="default2.asp"

    
    }
    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){

        var i;
        var j;
        i = 1;

<%  If m_sKBN = C_KAMOKU_IPPAN AND m_sGRP <> C_SENTAKU_JIYU Then%>

        do{
            j = 1;
            w_sTTNI = 0;
			w_sFLG = true

            do{

                w_sATO = eval("document.frm.ATO"+i+"_"+j);
                w_sTsu = eval("document.frm.Tanisuu"+j);
                w_sTsuG = eval("document.frm.txtSTani");

                if(w_sATO.value =="○"){
                    w_sTTNI = w_sTTNI + Number(w_sTsu.value);
                }
                if(w_sTTNI >= w_sTsuG.value){
                    break;
                }

            j++; }  while(j<=document.frm.n_Max.value);

            if(w_sTTNI < w_sTsuG.value){
                if (!confirm("最低取得単位に達していない人がいますが登録しますか？")) {
                   return ;
                }
                document.frm.action="web0340_upd.asp";
                document.frm.target="main";
                document.frm.submit();
				w_sFLG = false
                break;
            }
        i++; }  while(i<=document.frm.k_Max.value);

		if(w_sFLG == true){
			if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
			   return ;
			}
			document.frm.action="web0340_upd.asp";
			document.frm.target="main";
			document.frm.submit();
		}
<%Else%>
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        document.frm.action="web0340_upd.asp";
        document.frm.target="main";
        document.frm.submit();
<%End If%>
    
    }
    //-->
    </SCRIPT>

	<center>

	<body onload="return window_onload()">
	<FORM NAME="frm" method="post">

	    <%
		'//隠しフィールドに科目CDと各科目の単位数を格納(登録時に使用する)
        m_Krs.MoveFirst
        Do Until m_Krs.EOF
	        n = n + 1
		    %>
	        <input type=hidden name=kamokuCd<%=n%> value="<%=m_Krs("T16_KAMOKU_CD")%>">
	        <input type=hidden name=Tanisuu<%=n%> value="<%=m_Krs("T16_HAITOTANI")%>">
		    <%
	        m_Krs.MoveNext
        Loop%>
	<table class=hyo border=1>

	    <%
	        m_Grs.MoveFirst
	        Do Until m_Grs.EOF
	            Call gs_cellPtn(m_cell)
		        k = k + 1
		        j = 0
			    %>
			    <tr>
			        <td class=<%=m_cell%> width="70"><%=m_Grs("T13_GAKUSEKI_NO")%>
			        <input type=hidden name=gakuNo<%=k%> value="<%=m_Grs("T13_GAKUSEI_NO")%>"></td>
			        <td class=<%=m_cell%>  width="120"><%=m_Grs("T11_SIMEI")%>
			        <input type=hidden name=gakuNm<%=k%> value="<%=m_Grs("T11_SIMEI")%>"></td>
			    <%
		        m_Krs.MoveFirst
		        Do Until m_Krs.EOF
			        j = j + 1
			        Call f_KibouData() 
			        m_Krs.MoveNext
		        Loop

		        m_Grs.MoveNext
	        Loop%>
	    </tr>
	</table>
	<% If cint(m_sRisyuJotai) = C_K_RIS_MAE then %>
	<table>
	    <tr>
	        <td align=center><input type=button class=button value="　登　録　" onclick="javascript:f_Touroku()"></td>
	        <td align=center><input type=button class=button value="キャンセル" onclick="javascript:f_Cansel()"></td>
	    </tr>
	</table>
	<% End If %>

	<input type="hidden" name="n_Max"       value="<%=n%>">
	<input type="hidden" name="k_Max"       value="<%=k%>">
	<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtSTani"    value="<%=m_iSTani%>">

	<input type="hidden" name="txtGakunen"  value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass"    value="<%=m_sClass%>">
	<input type="hidden" name="txtKBN"      value="<%=m_sKBN%>">
	<input type="hidden" name="txtGRP"      value="<%=m_sGRP%>">
	<input type="hidden" name="txtRisyu"      value="<%=m_sRisyuJotai%>">


	</FORM>
	</center>
	</BODY>
	</HTML>
<%
End Sub
%>