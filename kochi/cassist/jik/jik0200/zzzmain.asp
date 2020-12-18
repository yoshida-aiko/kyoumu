<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 教官別授業時間一覧
'* ﾌﾟﾛｸﾞﾗﾑID : jik/jik0200/main.asp
'* 機      能: 下ページ 時間割マスタの一覧リスト表示を行う
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*           cboKyokaKeiCd      :科目系列コード
'*           cboKyokanCd      :教官コード
'*           txtMode         :動作モード
'           :session("PRJ_No")      '権限ﾁｪｯｸのキー
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'* 説      明:
'*           選択された教官の授業時間一覧を表示
'*-------------------------------------------------------------------------
'* 作      成: 2001/07/03 根本 直美
'* 変      更: 2001/07/30 根本 直美 戻り先URL変更
'*                                  変数名命名規則に基く変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '取得したデータを持つ変数
    Public  m_iSyoriNen         ':処理年度
    Public  m_iKyokanCd         ':教官コード
    Public  m_iSKyokanCd        ':選択教官コード
    
    Public  m_Rs                'recordset
    
    Public  m_iGakunen          ':学年
    Public  m_sClass            ':クラス
    Public  m_sYobi             ':表示曜日
    Public  m_iYobiCd           ':曜日コード
    Public  m_iJigen            ':時限
    Public  m_iKamokuCd         ':科目コード
    Public  m_sKamoku           ':科目名
    Public  m_iKyosituCd        ':教室コード
    Public  m_sKyositu          ':教室名
    
    Public  m_sCellD             ':テーブルセル色（曜日）'//2001/07/30変更
    Public  m_iJMax             ':最大時限数
    Public  m_Flg			'時間割１限目確認フラグ
    
    'ページ関係
    Public  m_iMax              ':最大ページ
    Public  m_iDsp              '// 一覧表示行数

    'データ取得用
    Public  m_iYobiCnt          ':カウント（曜日）
    Public  m_iJgnCnt           ':カウント（時限）
    Public  m_iYobiCCnt         ':カウント（曜日・テーブル色表示用）
    
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
    w_sMsgTitle="教官別授業時間一覧"
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
        
        '授業時間割テーブルマスタを取得
        
            w_sSQL = ""
            w_sSQL = w_sSQL & "SELECT"
            w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_GAKUNEN"
            w_sSQL = w_sSQL & vbCrLf & " ,M05_CLASS.M05_CLASSRYAKU"
            w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_YOUBI_CD"
            w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_JIGEN"
            w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_KAMOKU"
            w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_KYOSITU"
            w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_TUKU_FLG"
            w_sSQL = w_sSQL & vbCrLf & " FROM T20_JIKANWARI"
            w_sSQL = w_sSQL & vbCrLf & ", M05_CLASS"
            w_sSQL = w_sSQL & vbCrLf & " WHERE " 
            w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_NENDO = " & m_iSyoriNen
            w_sSQL = w_sSQL & vbCrLf & " AND M05_CLASS.M05_NENDO = " & m_iSyoriNen
            w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_KYOKAN = " & m_iSKyokanCd
            w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN(+) "
            w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO(+) "
            
            
'        Response.Write w_sSQL & "<br>"

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

    m_iKyokanCd = ""
    m_iSyoriNen = ""
    
    m_iKyokaKeiKbn = ""
    m_iSKyokanCd = ""
    
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
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")          ':教官コード
    'm_iKyokanCd = 000000                       ':教官コード'//テスト用
    m_iSyoriNen = Session("NENDO")              ':処理年度
    'm_iSyoriNen = 2002                         ':処理年度'//テスト用
    
    m_iSKyokanCd = Request("SKyokanCd1")       ':選択教官コード
    
End Sub

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
'        response.write "<td rowspan=" & m_iJMax & " class="
        'call showYobiColor()
        response.write m_sCellD
        response.write ">" & WeekdayName(m_iYobiCnt,True) & "</td>"
    else
    end if
    
End Sub

Function f_ShowClass()   '//2001/07/30変更
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  学年・クラスを表示
'********************************************************************************

    Dim w_sClass

    m_iGakunen = ""
    m_sClass = ""
    
    w_sClass = ""
    f_ShowClass = ""
    
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then
            m_iGakunen = m_Rs("T20_GAKUNEN")
            'w_sClass = Right(m_Rs("M05_CLASSRYAKU"),1)
            'm_sClass = m_sClass & w_sClass
            m_sClass = m_sClass & m_Rs("M05_CLASSRYAKU")
            
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst
    
    if m_iGakunen <> "" and m_sClass <> "" Then
        f_ShowClass =  m_iGakunen & "-" & m_sClass
    end if

End Function

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

            m_sKamoku = f_getKamokumei(m_iSyoriNen,m_Rs("T20_KAMOKU"),m_Rs("T20_GAKUNEN"),m_Rs("T20_TUKU_FLG"),w_iCourseCD) 
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_iGakunen <> "" and m_sClass <> "" Then
        f_ShowKamokuMei = m_sKamoku
    end if

End Function

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

        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(p_iYobiCnt) and right(cstr(cDbl(m_Rs("T20_JIGEN"))*10),1) <> "0" Then
            f_KamokuSu = f_KamokuSu + 1
        end if
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

'    if m_iGakunen <> "" and m_sClass <> "" Then
'        response.write m_sKamoku
'    end if

End Function

Sub s_SetKyositu()  '//2001/07/30変更
'********************************************************************************
'*  [機能]  値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  教室コードを設定
'********************************************************************************

    m_iKyosituCd = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then
            m_iKyosituCd = m_Rs("T20_KYOSITU")
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

End Sub

'********************************************************************************
'*  [機能]  教室名の取得
'*  [引数]  
'*  [戻値]  0:情報取得成功、1:レコードなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetKyosituMei()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    Call s_SetKyositu()
    f_GetKyosituMei = 0
    m_sKyositu = ""
    
    if m_iKyosituCd <> "" Then
        Do
            
            '// 教室名マスタﾚｺｰﾄﾞｾｯﾄを取得
            w_sSQL = ""
            w_sSQL = w_sSQL & "SELECT"
            w_sSQL = w_sSQL & " M06_KYOSITUMEI"
            w_sSQL = w_sSQL & " FROM M06_KYOSITU "
            w_sSQL = w_sSQL & " WHERE M06_NENDO = " & m_iSyoriNen
            w_sSQL = w_sSQL & " AND M06_KYOSITU_CD = " & m_iKyosituCd
            
            w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
            
            If w_iRet <> 0 Then
                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                'response.write w_iRet & "<br>"
                'm_sErrMsg = "ﾚｺｰﾄﾞｾｯﾄの取得失敗"
                'm_bErrFlg = True
                f_GetKyosituMei = 99
                Exit Do 'GOTO LABEL_f_GetKyosituMei_END
            Else
            End If
            
            If w_Rs.EOF Then
                '対象ﾚｺｰﾄﾞなし
                'm_sErrMsg = "対象ﾚｺｰﾄﾞなし"
                f_GetKyosituMei = 1
                Exit Do 'GOTO LABEL_f_GetKyosituMei_END
            End If
            
                '// 取得した値を格納
                    m_sKyositu = w_Rs("M06_KYOSITUMEI")    '//教室名を格納
            '// 正常終了
            Exit Do
        
        Loop
        
        gf_closeObject(w_Rs)
    
    end if

'// LABEL_f_GetKyosituMei_END
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
    w_i = p_iNendo - p_iGaknen + 1
    	w_sSQL = ""
    	w_sSQL = w_sSQL & vbCrLf & "SELECT "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKUMEI, "
    	w_sSQL = w_sSQL & vbCrLf & "T15_COURSE_CD"
    	w_sSQL = w_sSQL & vbCrLf & "FROM "
    	w_sSQL = w_sSQL & vbCrLf & " T15_RISYU "
    	w_sSQL = w_sSQL & vbCrLf & "WHERE "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKU_CD = '"&p_sKamokuCD&"' AND "
    	w_sSQL = w_sSQL & vbCrLf & "T15_NYUNENDO = "& p_iNendo - cint(p_iGaknen) + 1 &" "
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
   </head>

    <body>

    <center>
		<br><br><br>
		<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
    </center>

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

    Dim w_cellT             '//Tableセル色
    Dim w_sClass          '//クラス名
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

<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr>
<td align="center">

    <table border=1 class=hyo width="100%">
    <COLGROUP WIDTH="5%" ALIGN=center>
    <COLGROUP WIDTH="5%" ALIGN=center>
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="35%" ALIGN=center>
    <COLGROUP WIDTH="35%" ALIGN=center>
    <tr>
        <th colspan="2" class=header><br></th>  
        <th class=header>クラス</th>
        <th class=header>教科名</th>
        <th class=header>教室</th>
    </tr>

<%
m_iYobiCCnt = 1

'For m_iYobiCnt = 2 to 6
For m_iYobiCnt = C_YOUBI_MIN to C_YOUBI_MAX
 m_Flg = 0
 w_iJigenMax =  f_KamokuSu(m_iYobiCnt)
    For m_iJgnCnt = 0.5 to m_iJMax step 0.5
    
	w_sClass =  f_ShowClass()
	w_sKamoku = f_ShowKamokuMei()
	w_iJgnNo = m_iJgnCnt
	if right(cstr(w_iJgnNo*10),1) <> "0" then w_iJgnNo = " "

	If w_sKamoku <> "" OR w_sClass <> "" OR right(cstr(m_iJgnCnt*10),1) = "0" Then 
		call gs_cellPtn(w_cellT)
		%>
		    <tr>
		<%call s_ShowYobi(w_iJigenMax)%>
		        <td class=<%=w_cellT%>><%=w_iJgnNo%></td>
		        <td class=<%=w_cellT%>><%=w_sClass%><br></td>
		        <td class=<%=w_cellT%>><%=w_sKamoku%><br></td>
		        <td class=<%=w_cellT%>>
		<%
		call f_GetKyosituMei()
		    if f_GetKyosituMei = 0 Then
		        response.write m_sKyositu
		    else
		    end if
		%>
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
