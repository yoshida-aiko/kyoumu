<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 授業時間一覧
'* ﾌﾟﾛｸﾞﾗﾑID : login/jikanwari.asp
'* 機      能: 上ページ 時間割マスタの一覧リスト表示を行う
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'* 説      明:
'*           ログインした教官の授業時間一覧を表示
'*-------------------------------------------------------------------------
'* 作      成: 2001/07/19 根本 直美
'* 変      更: 2001/07/25 根本 直美
'*           : 2001/08/06 根本 直美     戻り先URL、target変更
'*           :                          変数名命名規則に基く変更
'*           : 2001/08/07 根本 直美     NN対応に伴うソース変更
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_sMsg              'ﾒｯｾｰｼﾞ
    
    '取得したデータを持つ変数
    Public  m_iSyoriNen         ':処理年度
    Public  m_iKyokanCd         ':教官コード
    
    Public  m_Rs                'recordset
    Public  m_Rds                'recordset
    
    Public  m_iGakunen          ':学年
    Public  m_sClass            ':クラス
    Public  m_sYobi             ':表示曜日
    Public  m_iYobiCd           ':曜日コード
    Public  m_iJigen            ':時限
    Public  m_iKamokuCd         ':科目コード
    Public  m_sKamoku           ':科目名
    Public  m_iKyosituCd        ':教室コード
    Public  m_sKyositu          ':教室名
    
    Public  m_sCellD             ':テーブルセル色（曜日）
    Public  m_iJMax             ':最大時限数
    
    'ページ関係
    Public  m_iMax              ':最大ページ
    Public  m_iDsp              '// 一覧表示行数

    'データ取得用
    Public  m_iYobiCnt          ':カウント（曜日）
    Public  m_iJgnCnt           ':カウント（時限）
    Public  m_iYobiCCnt         ':カウント（曜日・テーブル色表示用）
    Public  m_iDate             ':今日の日付
    Public  m_sYobiD            ':今日の曜日
    
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
    w_sMsgTitle="HOME"
    w_sMsg=""
    w_sRetURL="../default.asp"     
    w_sTarget="_top"


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

        '// 権限チェックに使用
        session("PRJ_No") = C_LEVEL_NOCHK

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// 値の初期化
        Call s_SetBlank()

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
        
        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetDate()
        
        '授業時間割テーブルマスタを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_GAKUNEN"
        w_sSQL = w_sSQL & vbCrLf & " ,M05_CLASS.M05_CLASSRYAKU"
        w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_YOUBI_CD"
        w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_JIGEN"
        w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_KYOSITU"
        w_sSQL = w_sSQL & vbCrLf & " ,M03_KAMOKU.M03_KAMOKUMEI"
        'w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_GODO_FLG"
        w_sSQL = w_sSQL & vbCrLf & " FROM T20_JIKANWARI"
        w_sSQL = w_sSQL & vbCrLf & ", M03_KAMOKU"
        w_sSQL = w_sSQL & vbCrLf & ", M05_CLASS"
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND M03_KAMOKU.M03_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND M05_CLASS.M05_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_KAMOKU = M03_KAMOKU.M03_KAMOKU_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN(+) "
        w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO(+) "



response.write "工事中" & "<BR>"
response.end
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
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
    
    m_sYobi = ""
    m_iYobiCd = ""
    m_iJigen = ""
    m_iKamokuCd = ""
    
    m_iYobiCnt = ""
    m_iJgnCnt = ""
    m_iYobiCCnt = ""
    
    m_iDate = ""
    m_sYobiD = ""
    
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")          ':教官コード
    m_iSyoriNen = Session("NENDO")              ':処理年度
    
End Sub

Sub s_SetDate()
'********************************************************************************
'*  [機能]  今日の日付を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_iDate = gf_YYYY_MM_DD(date(),"/")
    m_sYobiD = Weekdayname(Weekday(m_iDate),true)

End Sub

Sub s_SetYobi()
'********************************************************************************
'*  [機能]  曜日を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sYobi = Weekdayname(m_iYobiCnt,true)

End Sub


Sub showYobi()
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

    if m_iJgnCnt = 1 Then
        response.write "<td rowspan='" & m_iJMax & "' class='"
        response.write m_sCellD
        response.write "'>" & m_sYobi & "</td>"
    else
    end if
    
End Sub

Sub showClass()
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

    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CInt(m_Rs("T20_JIGEN")) = CInt(m_iJgnCnt) Then
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
        response.write m_iGakunen & "-" & m_sClass
    end if

End Sub

Sub showKamokuMei()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  科目名を表示
'********************************************************************************

    m_sKamoku = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CInt(m_Rs("T20_JIGEN")) = CInt(m_iJgnCnt) Then
            m_sKamoku = m_Rs("M03_KAMOKUMEI")
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_iGakunen <> "" and m_sClass <> "" Then
        response.write m_sKamoku
    end if

End Sub

Sub SetKyositu()
'********************************************************************************
'*  [機能]  値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  教室コードを設定
'********************************************************************************

    m_iKyosituCd = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CInt(m_Rs("T20_JIGEN")) = CInt(m_iJgnCnt) Then
            m_iKyosituCd = m_Rs("T20_KYOSITU")
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

End Sub

sub s_Jikanwari()
'********************************************************************************
'*  [機能]  時間割変更データの有無の確認
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT * "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T52_JYUGYO_HENKO "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & "     T52_KYOKAN_CD = '" & m_iKyokanCd & "' "
    w_sSQL = w_sSQL & " AND T52_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "'"
    w_sSQL = w_sSQL & " AND T52_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"

    Set m_Rds = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(m_Rds, w_sSQL,m_iDsp)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
    End If

    If m_Rds.EOF Then
        Exit Sub
    End If
%>
<a href="#" onclick=NewWin()> ※時間割の変更連絡があります</a>  
<%

End Sub

'********************************************************************************
'*  [機能]  教室名の取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetKyosituMei()
    
    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    Call SetKyositu()
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
                'm_sErrMsg = "レコードセットの取得に失敗しました"
                'm_bErrFlg = True
                's_SetKyosituMei = 99
                Exit Do 'GOTO LABEL_s_SetKyosituMei_END
            Else
            End If
            
            If w_Rs.EOF Then
                '対象ﾚｺｰﾄﾞなし
                'm_sErrMsg = "対象レコードがありません"
                's_SetKyosituMei = 1
                Exit Do 'GOTO LABEL_s_SetKyosituMei_END
            End If
            
                '// 取得した値を格納
                    m_sKyositu = w_Rs("M06_KYOSITUMEI")    '//教室名を格納
            '// 正常終了
            Exit Do
        
        Loop
        
        gf_closeObject(w_Rs)
    
    end if

response.write m_sKyositu

End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    Dim w_sCellT             '//Tableセル色

    On Error Resume Next
    Err.Clear

%>


<html>
<head>
<link rel="stylesheet" href="../common/style.css" type="text/css">
<!--#include file="../Common/jsCommon.htm"-->
<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
    //************************************************************
    //  [機能]  申請内容表示用ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function NewWin() {
        URL = "j_view.asp";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=450,height=450,top=0,left=0");
        return false;   
    }
//-->
</SCRIPT>
</head>
<body>
<center>
<BR>
<font size="3">今日の時間割</font>
<BR><BR>
<table border="0" width="90%">
<tr>
<td valign="top" align="center">
    <table border="1" class="hyo" width="100%">
    <COLGROUP WIDTH="5%" ALIGN="center">
    <COLGROUP WIDTH="5%" ALIGN="center">
    <COLGROUP WIDTH="20%" ALIGN="center">
    <COLGROUP WIDTH="35%" ALIGN="center">
    <COLGROUP WIDTH="35%" ALIGN="center">
    <tr>
        <th colspan="2" class="header"><br></th>  
        <th class="header">クラス</th>
        <th class="header">教科名</th>
        <th class="header">教室</th>
    </tr>
<%

    m_iYobiCCnt = 1
    
    'For m_iYobiCnt = C_YOUBI_MIN to C_YOUBI_MAX
    For m_iYobiCnt = 1 to 7

        For m_iJgnCnt = 1 to m_iJMax
        
        '//テーブルセル背景色
        call gs_cellPtn(w_sCellT)
        '//時間割テーブルから曜日を取得
        call s_SetYobi()

'response.write "aaaaaa" & "<BR>"
'response.end

            '//時間割の曜日と今日の曜日が同一の場合表示
            if m_sYobi = m_sYobiD Then

'response.write "aaaaaa" & "<BR>"
'response.end


%>
    <tr>
<%
            call showYobi()
%>
        <td class="<%=w_sCellT%>"><%=m_iJgnCnt%></td>


        <td class="<%=w_sCellT%>"><%call showClass()%><br></td>

        <td class="<%=w_sCellT%>"><%call showKamokuMei()%><br></td>
        <td class="<%=w_sCellT%>"><%call s_SetKyosituMei()%><br></td>
    </tr>
<%
            end if
        Next
    m_iYobiCCnt = m_iYobiCCnt + 1   '//曜日カウント（テーブル背景色表示用）
    Next
%>
    </table>
</td>
</tr>
</table>
<%
    Do Until m_Rs.EOF
%>
<%
m_Rs.MoveNext

    if m_Rs.EOF Then
        Exit Do
    end if
    Loop
%>
<br>
<%Call s_Jikanwari()%>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
