<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 教官別授業時間一覧
'* ﾌﾟﾛｸﾞﾗﾑID : jik/jik0210/main.asp
'* 機      能: 下ページ 時間割マスタの一覧リスト表示を行う
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*           cboGakunenCd      :学年コード
'*           cboClassCd      :クラスコード
'*           txtMode         :動作モード
'           :session("PRJ_No")      '権限ﾁｪｯｸのキー
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'* 説      明:
'*           選択されたクラスの授業時間一覧を表示
'*-------------------------------------------------------------------------
'* 作      成: 2001/07/06 根本 直美
'* 変      更: 2001/07/30 根本 直美  戻り先URL変更
'*                                  変数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_sMsg              'ﾒｯｾｰｼﾞ
    
    '取得したデータを持つ変数
    Public  m_iSyoriNen         ':処理年度
    Public  m_iKyokanCd         ':教官コード
    Public  m_iGakunen          ':学年コード
    Public  m_iClass            ':クラスコード
    
    Public  m_Rs                'recordset
    
    Public  m_sClass            ':クラス名
    Public  m_sYobi             ':表示曜日
    Public  m_iYobiCd           ':曜日コード
    Public  m_iJigen            ':時限
    Public  m_iKamokuCd         ':科目コード
    Public  m_sKamoku           ':科目名
    Public  m_iKyosituCd        ':教室コード
    Public  m_sKyositu          ':教室名
    Public  m_sKyokan           ':教官名
    Public  m_iNyuNen           ':入年度
    Public  m_iCourseCd         ':コースコード
    
    Public  m_iJMax             ':最大時限数
    Public  m_Flg			'時間割１限目確認フラグ
    
    Public m_iCourse            ':コースコード
    
    'ページ関係
    Public  m_iMax              ':最大ページ
    Public  m_iDsp              '// 一覧表示行数

    'データ取得用
    Public  m_iYobiCnt          ':カウント（曜日）
    Public  m_iJgnCnt           ':カウント（時限）
    Public  m_iYobiCCnt         ':カウント（曜日・テーブル色表示用）
    
    Public  m_sCellD             ':テーブルセル色（曜日）'//2001/07/30変更
    
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
    Dim w_sWHERE            '// WHERE文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="クラス別授業時間一覧"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

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

        '// 値の初期化
        Call s_SetBlank()

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '// クラス名取得
        Call f_GetClassMei()
            if f_GetClassMei <> 0 Then
                Exit Do
            end if
        
        '授業時間割テーブルマスタを取得
        
            w_sSQL = ""
            w_sSQL = w_sSQL & "SELECT"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_YOUBI_CD,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_JIGEN,"
            w_sSQL = w_sSQL & vbCrLf & " M04.M04_KYOKANMEI_SEI,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_KAMOKU, "
            w_sSQL = w_sSQL & vbCrLf & " M06.M06_KYOSITUMEI,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_DOJI_JISSI_FLG,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_TUKU_FLG "
            w_sSQL = w_sSQL & vbCrLf & "FROM "
            w_sSQL = w_sSQL & vbCrLf & "T20_JIKANWARI T20,"
            w_sSQL = w_sSQL & vbCrLf & "M04_KYOKAN M04,"
            w_sSQL = w_sSQL & vbCrLf & "M06_KYOSITU M06 "
            w_sSQL = w_sSQL & vbCrLf & "WHERE "
            w_sSQL = w_sSQL & vbCrLf & "T20.T20_KYOKAN = M04.M04_KYOKAN_CD(+) "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_NENDO = M04.M04_NENDO(+) "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_KYOSITU = M06.M06_KYOSITU_CD(+) "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_NENDO = M06.M06_NENDO(+) "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_DOJI_JISSI_FLG IS NULL "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_NENDO=" & m_iSyoriNen & " "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_GAKUNEN=" & m_iGakunen & " "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_CLASS=" & m_iClass & " "
            w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_YOUBI_CD,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_JIGEN,"
            w_sSQL = w_sSQL & vbCrLf & " M04.M04_KYOKANMEI_SEI,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_KAMOKU, "
            w_sSQL = w_sSQL & vbCrLf & " M06.M06_KYOSITUMEI,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_DOJI_JISSI_FLG,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_TUKU_FLG "
            w_sSQL = w_sSQL & vbCrLf & "ORDER BY "
'            w_sSQL = w_sSQL & vbCrLf & "T20.T20_GAKKI_KBN, "
            w_sSQL = w_sSQL & vbCrLf & "T20.T20_YOUBI_CD, "
            w_sSQL = w_sSQL & vbCrLf & "T20.T20_JIGEN "
            
'        Response.Write w_sSQL & vbCrLf &"<br>"
'response.end
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        'w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            'm_sErrMsg = Err.description
            m_sErrMsg = "レコードセットの取得に失敗しました"
            Exit Do 'GOTO LABEL_MAIN_END
        Else
            'ページ数の取得
            'm_iMax = gf_PageCount(m_Rs,m_iDsp)
        End If
        
        '//最大時限数を取得
        Call gf_GetJigenMax(m_iJMax)
            if m_iJMax = "" Then
                m_bErrFlg = True
                m_sErrMsg = Err.description
                Exit Do
            end if
        
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
    gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub
'********************************************************************************
'*  [機能]  全項目を空白に初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetBlank()

    m_sYobi = ""
    m_iYobiCd = ""
    m_iJigen = ""
    m_iKamokuCd = ""
    
    m_iYobiCnt = ""
    m_iJgnCnt = ""
    m_iYobiCCnt = ""
    
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()    '//2001/07/30変更

    m_iKyokanCd = Session("KYOKAN_CD")          ':教官コード
    'm_iKyokanCd = 000000                       ':教官コード'//テスト用
    m_iSyoriNen = Session("NENDO")              ':処理年度
    'm_iSyoriNen = 2002                         ':処理年度'//テスト用
    
    m_iGakunen = Request("cboGakunenCd")   ':学年コード
    m_iClass = Request("cboClassCd")       ':クラスコード
    
    m_iNyuNen = m_iSyoriNen - m_iGakunen + 1
    
End Sub

'********************************************************************************
'*  [機能]  クラス名の取得
'*  [引数]  
'*  [戻値]  0:情報取得成功、1:レコードなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetClassMei()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    f_GetClassMei = 0
    m_sClass = ""

    Do

        '// クラスマスタを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & "M05_CLASSMEI "
        w_sSQL = w_sSQL & vbCrLf & "FROM "
        w_sSQL = w_sSQL & vbCrLf & "M05_CLASS "
        w_sSQL = w_sSQL & vbCrLf & "WHERE "
        w_sSQL = w_sSQL & vbCrLf & "M05_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "AND M05_GAKUNEN = " & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "AND M05_CLASSNO = " & m_iClass
        
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
        
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            f_GetClassMei = 99
            Exit Do 'GOTO LABEL_f_GetClassMei_END
        Else
        End If
        
        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            f_GetClassMei = 1
            Exit Do 'GOTO LABEL_f_GetClassMei_END
        End If

            '// 取得した値を格納
            m_sClass = w_Rs("M05_CLASSMEI")
        '// 正常終了
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs)

'// LABEL_f_GetClassMei_END
End Function

Function f_getKamokumei(p_iNendo,p_sKamokuCD,p_iGaknen,p_iTUKU,p_iCourseCD) 
'********************************************************************************
'*  [機能]  科目名の取得(ついでにコースも取ってきます。)
'*  [引数]  
'*  [戻値]  科目名
'*  [説明]  2001/9/15
'********************************************************************************
    dim w_sSQL,w_Rs,w_iRet
    
    On Error Resume Next
    Err.Clear
    
    f_getKamokumei = "-"
    p_iCourseCD = 0 
    
  Do

   if p_iTUKU = C_TUKU_FLG_TOKU then '特別活動のときは、M41特別活動マスタから名称取得
    	w_sSQL = ""
    	w_sSQL = w_sSQL & vbCrLf & "SELECT "
    	w_sSQL = w_sSQL & vbCrLf & "M41_MEISYO "
    	w_sSQL = w_sSQL & vbCrLf & "FROM  "
    	w_sSQL = w_sSQL & vbCrLf & "M41_TOKUKATU "
    	w_sSQL = w_sSQL & vbCrLf & "WHERE "
    	w_sSQL = w_sSQL & vbCrLf & "M41_NENDO = " & p_iNendo & " AND "
    	w_sSQL = w_sSQL & vbCrLf & "M41_TOKUKATU_CD = '" & p_sKamokuCD & "' "
    	w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
    	w_sSQL = w_sSQL & vbCrLf & "M41_MEISYO "

   	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
   	If w_iRet <> 0 OR w_Rs.EOF = true Then Exit Do 
		
   	f_getKamokumei = w_Rs("M41_MEISYO")

    Else '普通の授業のときは、T15履修から名称取得
    	w_sSQL = ""
    	w_sSQL = w_sSQL & vbCrLf & "SELECT "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKUMEI, "
    	w_sSQL = w_sSQL & vbCrLf & "T15_COURSE_CD"
    	w_sSQL = w_sSQL & vbCrLf & "FROM "
    	w_sSQL = w_sSQL & vbCrLf & " T15_RISYU "
    	w_sSQL = w_sSQL & vbCrLf & "WHERE "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKU_CD = '"&p_sKamokuCD&"' AND "
    	w_sSQL = w_sSQL & vbCrLf & "T15_NYUNENDO = "& p_iNendo - p_iGaknen + 1 &" "
    	w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKUMEI, T15_COURSE_CD"

    	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    	If w_iRet <> 0 OR w_Rs.EOF = true Then Exit Do 

    	f_getKamokumei = w_Rs("T15_KAMOKUMEI")
    	p_iCourseCD = w_Rs("T15_COURSE_CD") 

    End if
    Exit Do

   Loop

end Function

Sub s_ShowYobi(p_iJigenMax)    '//2001/07/30変更
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  曜日を表示（テーブル用）
'********************************************************************************

if m_iYobiCCnt Mod 2 <> 0 Then
    m_sCellD = ""
end if

call gs_cellPtn(m_sCellD)

    if m_iJgnCnt <= 1  And m_Flg = 0 Then
	m_Flg = 1
        'response.write "<td rowspan=8 class="
        response.write "<td rowspan=" & p_iJigenMax & " class="
        'call showYobiColor()
        response.write m_sCellD
        response.write ">" & WeekdayName(m_iYobiCnt,True) & "</td>"
    else
    end if
    
End Sub

Function f_ShowKamokuMei()   '//2001/07/30変更
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  科目名を表示
'********************************************************************************
dim w_iCourseCD
    m_sKamoku = ""
    f_ShowKamokuMei = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then

		m_sKamoku = f_getKamokumei(m_iSyoriNen,m_Rs("T20_KAMOKU"),m_iGakunen,m_Rs("T20_TUKU_FLG"),w_iCourseCD) 

            if CInt(w_iCourseCD) = 0 Then

'                m_iCourseCd = m_Rs("T15_COURSE_CD")
                Exit Do
            else
                m_sKamoku = "-"
                'm_iCourseCd = m_Rs("T15_COURSE_CD")
            end if
        else
            m_sKamoku = "-"
            'm_iCourseCd = m_Rs("T15_COURSE_CD")
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst
    
    if m_sKamoku = "-" Then
        Call f_GetSentaku()
            'if m_sKamoku = "@@" Then
            if f_GetSentaku = 1 or m_sKamoku = "" Then
                Call s_GetCourse()
            end if
    end if

    if m_iGakunen <> "" and m_iClass <> "" Then
        f_ShowKamokuMei = m_sKamoku
    else
        f_ShowKamokuMei =  "-"
    end if

End Function

'********************************************************************************
'*  [機能]  選択科目名の取得
'*  [引数]  
'*  [戻値]  0:情報取得成功、1:レコードなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetSentaku()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    f_GetSentaku = 0
    m_sKamoku = ""
    m_sKyokan = ""

    Do

        '// 授業時間割テーブルマスタを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"

w_sSQL = w_sSQL & vbCrLf & "T20.T20_YOUBI_CD, "
w_sSQL = w_sSQL & vbCrLf & "T20.T20_JIGEN,"
w_sSQL = w_sSQL & vbCrLf & "T20.T20_HYOJI_KYOKAN,"
w_sSQL = w_sSQL & vbCrLf & "T18.T18_SYUBETU_MEI,"
w_sSQL = w_sSQL & vbCrLf & "T20.T20_DOJI_JISSI_FLG,"
w_sSQL = w_sSQL & vbCrLf & "T15.T15_COURSE_CD"
w_sSQL = w_sSQL & vbCrLf & "FROM "
w_sSQL = w_sSQL & vbCrLf & "T20_JIKANWARI T20, "
w_sSQL = w_sSQL & vbCrLf & "T15_RISYU T15,"
w_sSQL = w_sSQL & vbCrLf & "T18_SELECTSYUBETU T18 "
w_sSQL = w_sSQL & vbCrLf & "WHERE "
w_sSQL = w_sSQL & vbCrLf & "T20.T20_KAMOKU = T15.T15_KAMOKU_CD(+) "
w_sSQL = w_sSQL & vbCrLf & "AND T15.T15_NYUNENDO = T18.T18_NYUNENDO(+) "
w_sSQL = w_sSQL & vbCrLf & "AND T15.T15_GRP = T18.T18_GRP(+) "
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_GAKKI_KBN = '1' "
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_NENDO = " & m_iSyoriNen
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_GAKUNEN = " & m_iGakunen
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_CLASS = " & m_iClass
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_YOUBI_CD = " & m_iYobiCnt
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_JIGEN = " & m_iJgnCnt
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_DOJI_JISSI_FLG Is Not Null "
w_sSQL = w_sSQL & vbCrLf & "AND T15.T15_NYUNENDO = " & m_iNyunen
w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
w_sSQL = w_sSQL & vbCrLf & "T20.T20_YOUBI_CD, "
w_sSQL = w_sSQL & vbCrLf & "T20.T20_JIGEN,"
w_sSQL = w_sSQL & vbCrLf & "T20.T20_HYOJI_KYOKAN,"
w_sSQL = w_sSQL & vbCrLf & "T18.T18_SYUBETU_MEI,"
w_sSQL = w_sSQL & vbCrLf & "T20.T20_DOJI_JISSI_FLG,"
w_sSQL = w_sSQL & vbCrLf & "T15.T15_COURSE_CD"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
        
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            'response.write w_iRet & "<br>"
            'm_sErrMsg = "ﾚｺｰﾄﾞｾｯﾄの取得失敗"
            'm_bErrFlg = True
            f_GetSentaku = 99
            Exit Do 'GOTO LABEL_f_GetSentaku_END
        Else
        End If
        
        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            'm_sErrMsg = "対象ﾚｺｰﾄﾞなし"
            f_GetSentaku = 1
            Exit Do 'GOTO LABEL_f_GetSentaku_END
        End If

            '// 取得した値を格納
            'm_sKamoku = w_Rs("T18_SYUBETU_MEI") & "@@"
            'm_sKyokan =  w_Rs("T20_HYOJI_KYOKAN") & "@@"
            'm_sKyositu =  "@@"
            'm_iCourseCd = w_Rs("T15_COURSE_CD") & "@@"
            if IsNull(w_Rs("T18_SYUBETU_MEI")) = False Then
                m_sKamoku = w_Rs("T18_SYUBETU_MEI")
            else
                m_sKamoku = ""
            end if
            if IsNull(w_Rs("T20_HYOJI_KYOKAN")) = False Then
                m_sKyokan = w_Rs("T20_HYOJI_KYOKAN")
            else
                m_sKyokan = ""
            end if
            m_sKyositu =  ""
            'm_iCourseCd = w_Rs("T15_COURSE_CD")
        '// 正常終了
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs)

'// LABEL_f_GetSentaku_END
End Function


'********************************************************************************
'*  [機能]  コース別の取得
'*  [引数]  
'*  [戻値]  0:情報取得成功、1:レコードなし、99:失敗
'*  [説明]  
'********************************************************************************
Sub s_GetCourse()
    
    Dim w_Rs2                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet2              '// 戻り値
    Dim w_sSQL2              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    m_iCourse = 0
    m_sKamoku = ""
    m_sKyokan = ""
    m_sKyositu = ""

    Do

        '// 授業時間割テーブルマスタを取得
        w_sSQL2 = ""
        w_sSQL2 = w_sSQL2 & "SELECT"
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_YOUBI_CD, "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_JIGEN,"
'w_sSQL2 = w_sSQL2 & vbCrLf & "T15.T15_KAMOKUMEI,"
'w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_KYOSITU,"
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_HYOJI_KYOKAN as M04_KYOKANMEI_SEI"
w_sSQL2 = w_sSQL2 & vbCrLf & "FROM "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20_JIKANWARI T20, "
w_sSQL2 = w_sSQL2 & vbCrLf & "T15_RISYU T15"
w_sSQL2 = w_sSQL2 & vbCrLf & "WHERE "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_KAMOKU = T15.T15_KAMOKU_CD(+) "
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_GAKKI_KBN = '1' "
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_NENDO = " & m_iSyoriNen
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_GAKUNEN = " & m_iGakunen
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_CLASS = " & m_iClass
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T15.T15_COURSE_CD != '0' "
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T15.T15_NYUNENDO = " & m_iNyunen
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_YOUBI_CD = " & m_iYobiCnt
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_JIGEN = " & m_iJgnCnt
w_sSQL2 = w_sSQL2 & vbCrLf & "GROUP BY "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_YOUBI_CD, "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_JIGEN,"
'w_sSQL2 = w_sSQL2 & vbCrLf & "T15.T15_KAMOKUMEI,"
'w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_KYOSITU,"
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_HYOJI_KYOKAN"

        w_iRet2 = gf_GetRecordset(w_Rs2, w_sSQL2)
'response.write w_sSQL2 & "<br>"
        
        If w_iRet2 <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            'response.write w_iRet2 & "<br>"
            'm_sErrMsg = "ﾚｺｰﾄﾞｾｯﾄの取得失敗"
            'm_bErrFlg = True
            'response.write "?"
            m_iCourse = 99
            Exit Do
        Else
        End If
        
        If w_Rs2.EOF Then
            '対象ﾚｺｰﾄﾞなし
            'm_sErrMsg = "対象ﾚｺｰﾄﾞなし"
            m_iCourse = 1
            'response.write "---"
            Exit Do
        End If

            '// 取得した値を格納
            'm_sKamoku = w_Rs2("T15_KAMOKUMEI") & "*"
            if m_iCourse = 0 Then
                m_sKamoku = "コース別"
                if IsNull(w_Rs2("T20_HYOJI_KYOKAN")) = False Then
                    'm_sKyokan = m_iCourse
                    m_sKyokan = w_Rs2("T20_HYOJI_KYOKAN")
                else
                    m_sKyokan = ""
                end if
                m_sKyositu = w_Rs2("T20_KYOSITU")
            else
                m_sKamoku = "?"
                m_sKyokan = "?"
                m_sKyositu = "?"
            end if
        '// 正常終了
        Exit Do
    
    
    Loop
    
    gf_closeObject(w_Rs2)

End Sub

Function f_ShowKyokanMei()   '//2001/07/30変更
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  教官名を表示
'********************************************************************************

    m_sKyokan = ""
    f_ShowKyokanMei = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then
            m_sKyokan = m_Rs("M04_KYOKANMEI_SEI")
            Exit Do

        else
            'm_sKyokan = ""
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_sKyokan = "" Then
        Call f_GetSentaku()
            if f_GetSentaku = 1 or m_sKyokan = "" Then
            'if f_GetSentaku = 1 or m_sKyokan = "@@" Then
                Call s_GetCourse()
            end if
    end if

    if m_iGakunen <> "" and m_iClass <> "" Then
        f_ShowKyokanMei = m_sKyokan
    else
        f_ShowKyokanMei = "-"
    end if
    
End Function

Sub s_ShowKyosituMei()  '//2001/07/30変更
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  教室名を表示
'********************************************************************************

    m_sKyositu = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then
            m_sKyositu = m_Rs("M06_KYOSITUMEI")
            Exit Do
        else
            'm_sKyositu = ""
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_sKyositu = "" Then
        Call f_GetSentaku()
            if f_GetSentaku = 1 or m_sKyositu = "" Then
            'if f_GetSentaku = 1 or m_sKyositu = "@@" Then
                Call s_GetCourse()
            end if
    end if

    if m_iGakunen <> "" and m_iClass <> "" Then
        response.write m_sKyositu
    else
        response.write "-"
    end if

End Sub

Function f_KamokuSu(p_iYobiCnt)   '//2001/09/06 add
'********************************************************************************
'*  [機能]  科目数を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  科目名を表示
'********************************************************************************
    f_KamokuSu = cint(m_iJMax)
    m_Rs.MoveFirst
    Do Until m_Rs.EOF

        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(p_iYobiCnt) and right(cstr(CDbl(m_Rs("T20_JIGEN"))*10),1) <> "0" Then
            f_KamokuSu = f_KamokuSu + 1
        end if
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

'    if m_iGakunen <> "" and m_sClass <> "" Then
'        response.write m_sKamoku
'    end if

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    Dim w_cellT             '//Tableセル色
    Dim w_sKyokan        '//教官名
    Dim w_sKamoku       '//科目名
    Dim w_iJigenMax      '//曜日のrow用
    Dim w_iJgnNo          '//表示用時限

   
    On Error Resume Next
    Err.Clear

%>


<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
</head>

<body>

<center>

<table border=1 class=hyo >
	<tr>
	<th class=header width="64"  align="center">クラス</th>
	<td class=detail width="50"  align="center"><%=m_iGakunen%>年</td>
	<td class=detail width="130" align="center"><%=m_sClass%></td>
	</tr>
</table>
<br>
<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr>
<td align="center">

    <table border=1 class=hyo width="100%">
        <COLGROUP WIDTH="5%" ALIGN=center>
        <COLGROUP WIDTH="5%" ALIGN=center>
        <COLGROUP WIDTH="30%" ALIGN=center>
        <COLGROUP WIDTH="30%" ALIGN=center>
        <COLGROUP WIDTH="30%" ALIGN=center>
        <tr>
            <th colspan="2" class=header><br></th>  
            <th class=header>科目</th>
            <th class=header>教官</th>
            <th class=header>教室</th>
        </tr>

<%
m_iYobiCCnt = 1

For m_iYobiCnt = C_YOUBI_MIN to C_YOUBI_MAX
 m_Flg = 0
 w_iJigenMax =  f_KamokuSu(m_iYobiCnt)
    For m_iJgnCnt = 0.5 to m_iJMax step 0.5

	w_sKyokan =  f_ShowKyokanMei()
	w_sKamoku = f_ShowKamokuMei()
	w_iJgnNo = m_iJgnCnt
	if right(cstr(w_iJgnNo*10),1) <> "0" then w_iJgnNo = " "

	If w_sKamoku <> "" OR w_sClass <> "" OR right(cstr(m_iJgnCnt*10),1) = "0" Then 

    call gs_cellPtn(w_cellT)
%>
    <tr>
<%call s_ShowYobi(w_iJigenMax)%>
        <td class=<%=w_cellT%>>
        <%=w_iJgnNo%>
        <br></td>
        <td class=<%=w_cellT%>>
        <%=w_sKamoku%>
        <br></td>
        <td class=<%=w_cellT%>>
        <%=w_sKyokan%>
        <br></td>
        <td class=<%=w_cellT%>>
        <%call s_ShowKyosituMei()%>
        <br></td>
    </tr>
<%

	End If
    Next
m_iYobiCCnt = m_iYobiCCnt + 1   '//曜日カウント（テーブル背景色表示用）
%>
<%
Next
%>

    </table>

</td>
</tr>
</table>

</center>

</body>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>
