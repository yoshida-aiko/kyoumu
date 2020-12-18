<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0144/kakunin.asp
' 機      能: 下ページ 就職先マスタの詳細変更を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           txtSinroKBN     :進路コード
'           txtSingakuCd        :進学コード
'           txtSinroName        :進路名称（一部）
'           txtPageSinro        :表示済表示頁数（自分自身から受け取る引数）
'           Sinro_syuseiCD      :選択された進路コード
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           txtSinroKBN     :進路コード（戻るとき）
'           txtSingakuCd        :進学コード（戻るとき）
'           txtSinroName        :進路名称（戻るとき）
'           txtPageSinro        :表示済表示頁数（戻るとき）
' 説      明:
'           ■初期表示
'               指定された進学先・就職先の詳細データを表示
'           ■地図画像ボタンクリック時
'               指定した条件にかなう進学先・就職先を表示する（別ウィンドウ）
'-------------------------------------------------------------------------
' 作      成: 2001/07/12 岩下 幸一郎
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数

    Public  m_Rs            'recordset
    Public  m_sDBMode           'DBのﾓｰﾄﾞの設定
    Public  m_sDATE
    Public  m_sNendo
    Public  m_sGakkiCD
    Public  m_sGakunenCD
    Public  m_sGakkaCD
    Public  m_sCourseCD
    Public  m_sKamokuCD
    Public  m_sKyokanCD
    Public  m_sKyokanMei
    Public  m_sKyokasyoName
    Public  m_sSyuppansya
    Public  m_sTyosya
    Public  m_sKyokanyo
    Public  m_sSidousyo
    Public  m_sBiko
    Public  m_sMaxNO

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

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="就職先マスタ"
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
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '// DB登録
        if m_sDBMode = "Insert" then
            w_iRet = f_Insert
        else
            w_iRet = f_Update
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
    call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    Dim strErrMsg

    strErrMsg = ""

    m_sDBMode    = Request("txtMode")       'DBﾓｰﾄﾞの取得
    m_sNendo     = Request("txtNendo")      '年度の取得
    m_sGakkiCD   = Request("txtGakkiCD")    '学期の取得
    m_sGakunenCD     = Request("txtGakunenCD")  '学年の取得
    m_sGakkaCD   = Request("txtGakkaCD")    '学科の取得
    m_sCourseCD  = Request("txtCourseCD")   'コースの取得
    m_sKamokuCD  = Request("txtKamokuCD")   '科目の取得
    m_sKyokanMei     = Request("txtKyokanMei")  '教官名の取得
    m_sKyokasyoName  = Request("txtKyokasyoName")   '教科書名の取得
    m_sSyuppansya    = Request("txtSyuppansya") '出版社の取得
    m_sTyosya    = Request("txtTyosya")     '著者の取得
    m_sKyokanyo  = Request("txtKyokanyo")   '教官用の取得
    m_sSidousyo  = Request("txtSidousyo")   '指導書の取得
    m_sBiko      = Request("txtBiko")       '教官用の取得

    m_sDate = gf_YYYY_MM_DD(date(),"/")

    'm_sKyokanCD = Session("KYOKAN_CD")          ':ユーザーID
    m_sKyokanCD = Request("SKyokanCd1")          ':ユーザーID

    if m_sDBMode = "Insert" then
        Call f_Max()
        m_sMaxNO = Cint(m_sMaxNO) + 1
    else
        m_sMaxNO = Request("txtUpdNo")
    end if

    If strErrmsg <> "" Then
        ' エラーを表示するファンクション
        Call err_page(strErrMsg)
        response.end
    End If
'   call s_viewForm(request.form)   'デバッグ用　引数の内容を見る
End Sub


'********************************************************************************
'*  [機能]  新規連絡先コードを生成
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_Max()

Dim w_Rs
Dim w_sSQL
dim f_MaxNO 

    f_MaxNO = 0
    m_sMaxNO = 0

    '// 連絡先コードを取得（連絡先コードMax値）
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
    w_sSQL = w_sSQL & " MAX( "
    w_sSQL = w_sSQL & " T47.T47_NO"
    w_sSQL = w_sSQL & " ) "
    w_sSQL = w_sSQL & " AS MAXNO "
    w_sSQL = w_sSQL & " FROM T47_KYOKASYO T47 "
    w_sSQL = w_sSQL & " WHERE T47_NENDO = " & m_sNendo & " "

'response.write w_sSQL & "<<<BR>"

    w_sRet = gf_GetRecordset(w_Rs, w_sSQL)

    If w_sRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_sMsg = Err.description
        Exit function
    End If

    IF w_Rs.EOF THEN
        f_MaxNO = 0
    Else
        f_MaxNO = gf_SetNull2Zero(w_Rs("MAXNO"))
    End If
    
    m_sMaxNO = f_MaxNO

'    response.write("<BR>m_sMaxNO = " & m_sMaxNO)

End Function

'********************************************************************************
'*  [機能]  新規登録処理
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
function f_Insert
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文

    f_Insert = False

    w_sSQL = w_sSQL & vbCrLf & " Insert Into "
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKASYO"
    w_sSQL = w_sSQL & "(T47_NENDO,"
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKKI_KBN,"
    w_sSQL = w_sSQL & vbCrLf & " T47_NO,"
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKUNEN,"
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKKA_CD,"
    w_sSQL = w_sSQL & vbCrLf & " T47_COURSE_CD,"
    w_sSQL = w_sSQL & vbCrLf & " T47_KAMOKU,"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKAN,"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKASYO,"
    w_sSQL = w_sSQL & vbCrLf & " T47_SYUPPANSYA,"
    w_sSQL = w_sSQL & vbCrLf & " T47_TYOSYA,"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKANYOUSU,"
    w_sSQL = w_sSQL & vbCrLf & " T47_SIDOSYOSU,"
    w_sSQL = w_sSQL & vbCrLf & " T47_BIKOU,"
    w_sSQL = w_sSQL & vbCrLf & " T47_INS_DATE,"
    w_sSQL = w_sSQL & vbCrLf & " T47_INS_USER,"
    w_sSQL = w_sSQL & vbCrLf & " T47_UPD_DATE,"
    w_sSQL = w_sSQL & vbCrLf & " T47_UPD_USER)"
    w_sSQL = w_sSQL & vbCrLf & " Values"
    w_sSQL = w_sSQL & "(" & m_sNendo & ","
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sGakkiCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " " & m_sMaxNO & ","
    w_sSQL = w_sSQL & vbCrLf & " " & m_sGakunenCD & ","
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sGakkaCD & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sCourseCD & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sKamokuCD & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sKyokanCD & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sKyokasyoName & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sSyuppansya & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sTyosya & "',"
    w_sSQL = w_sSQL & vbCrLf & " " & m_sKyokanyo & ","
    w_sSQL = w_sSQL & vbCrLf & " " & m_sSidousyo & ","
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sBiko & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sDATE & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sDATE & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "')"

    w_iRet = gf_ExecuteSQL(w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    End If

    f_Insert = True

End Function

'********************************************************************************
'*  [機能]  更新処理
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
function f_Update
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文

    f_Update = False

    w_sSQL = w_sSQL & vbCrLf & " Update T47_KYOKASYO SET "
    w_sSQL = w_sSQL & vbCrLf & " T47_NENDO = " & m_sNendo & ","
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKKI_KBN = '" & m_sGakkiCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKUNEN = " & m_sGakunenCD &","
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKKA_CD = '" & m_sGakkaCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_COURSE_CD = '" & m_sCourseCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_KAMOKU = '" & m_sKamokuCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKAN = '" & m_sKyokanCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKASYO = '" & m_sKyokasyoName & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_SYUPPANSYA = '" & m_sSyuppansya & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_TYOSYA = '" & m_sTyosya & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKANYOUSU = " & m_sKyokanyo & ","
    w_sSQL = w_sSQL & vbCrLf & " T47_SIDOSYOSU = " & m_sSidousyo & ","
    w_sSQL = w_sSQL & vbCrLf & " T47_BIKOU = '" & m_sBiko & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_UPD_DATE = '" & m_sDATE & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_UPD_USER = '" & Session("LOGIN_ID") & "' "
    w_sSQL = w_sSQL & vbCrLf & " Where T47_NENDO = " & Request("KeyNendo")
'    w_sSQL = w_sSQL & vbCrLf & " and T47_GAKKI_KBN = '" & m_sGakkiCD & "'"
    w_sSQL = w_sSQL & vbCrLf & " and T47_NO = " & m_sMaxNO 

'Response.Write w_sSQL & "<br>"

    w_iRet = gf_ExecuteSQL(w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    End If

    f_Update = True

End Function

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
'Sub s_MapHTML()

'   If ISNULL(m_Rs("M13_TIZUFILENAME")) OR m_Rs("M13_TIZUFILENAME")="" Then
'       Response.Write("登録されていません")
'   Else
'       Response.Write("<a Href=""javascript:f_OpenWindow('" & Session("TYUGAKU_TIZU_PATH") & m_Rs("M13_TIZUFILENAME") & "')"">周辺地図</a>")
'   End If
    
'End Sub


Sub S_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_slink
Dim w_iCnt

w_iCnt = 0

Do While not m_Rs.EOF

w_slink = "　"

if m_Rs("M32_SINRO_URL") <> "" Then 
    w_sLink= "<a href='" & gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "'>" 
    w_sLink= w_sLink &  gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "</a>"
End if

        %>
        <%=w_slink%>
        <%
            m_Rs.MoveNext

        Loop

    'LABEL_showPage_OPTION_END
End sub


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

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    function gonext() {
        window.alert('<%=C_TOUROKU_OK_MSG%>');

            document.frm.action = "./default.asp";
            document.frm.target="fTopMain";
            document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>

<body bgcolor="#ffffff" onLoad="setTimeout('gonext()',0000)">

<center>

<Form Name ="frm" Action="">


<input type="hidden" Name="txtMode" Value="">
<input type="hidden" name="SKyokanCd1" value="<%=m_sKyokanCD%>">

</From>
</center>

</body>

</html>


<%
    '---------- HTML END   ----------
End Sub

Sub Nyuryokuzumi()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

%>

    <html>
    <head>
    </head>

    <body>

    <center>
    <font size="2">入力された連絡先コードはすでに使用済みです<br><br></font>
    <input type="button" onclick="javascript:history.back()" value="戻　る">
    </center>
    </body>

    </html>


<%
    '---------- HTML END   ----------
End Sub
%>