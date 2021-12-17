<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0144/syusei.asp
' 機      能: 下ページ 進路マスタの詳細変更を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSinroCD      :進路コード
'           txtSingakuCd        :進学コード
'           txtSyusyokuName     :進路名称（一部）
'           txtPageSinro        :表示済表示頁数（自分自身から受け取る引数）
'           RenrakusakiCD       :選択された進路コード
'           txtSchMode          :検索ﾓｰﾄﾞ
'                                   + JyusyoSch = 住所検索
'                                   + ZipSch    = 郵便番号検索
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSinroCD      :進路コード（戻るとき）
'           txtSingakuCd        :進学コード（戻るとき）
'           txtSyusyokuName     :進路名称（戻るとき）
'           txtPageSinro        :表示済表示頁数（戻るとき）
' 説      明:
'           ■初期表示
'               指定された進学先・就職先の詳細データを表示
'           ■地図画像ボタンクリック時
'               指定した条件にかなう進学先・就職先を表示する（別ウィンドウ）
'-------------------------------------------------------------------------
' 作      成: 2001/06/26 岩下 幸一郎
' 変      更: 2001/07/13 谷脇　良也
' 　      　: 2001/07/24 根本　直美（DB変更に伴う修正）
' 　　　　　: 2001/08/22 伊藤　公子　業種区分追加対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public m_sSinroCD           ':データベースから取得した進路区分
    Public m_sSinroCD3          ':同ページから取得した進路区分
    Public m_sSingakuCD         ':データベースから取得した進学区分
    Public m_sSinromei          ':データベースから取得した進路名称
    'Public m_sSinromei_Eig     ':データベースから取得した進路英語名称
    Public m_sSinromei_Kan      ':データベースから取得した進路名称カナ
    Public m_sSinromei_Rya      ':データベースから取得した進路略称
    Public m_sJusyo             ':データベースから取得した進路住所
    Public m_sJusyo1            ':データベースから取得した進路住所1
    Public m_sJusyo2            ':データベースから取得した進路住所2
    Public m_sJusyo3            ':データベースから取得した進路住所3
    Public m_sTel               ':データベースから取得した進路電話番号
    Public m_sSinro_URL         ':データベースから取得した進路URL
    Public m_Rs                 ':recordset
    Public m_sMode              ':モード
    Public m_sSchMode           ':検索ﾓｰﾄﾞ
    Public m_bReFlg             ':リロードされたかどうか
    Public m_sRenrakusakiCD     ':Mainから取得した連絡先CD
    Public m_sPageCD            ':ページ数
    Public m_sSinroCD2          ':Mainから取得した進路区分
    Public m_sSingakuCD2        ':Mainから取得した進学区分
    Public m_sSyusyokuName      ':Mainから取得した検索名称（一部)
    Public m_iNendo             ':年度
    Public m_sSubtitle          ':サブタイトル
    Public m_iKenCd             ':県コード
    Public m_iSityoCd           ':市町村コード
    Public m_sYubin             ':郵便番号
    'Public m_iGyosyu_Kbn        ':業種区分
    Public m_iSihonkin          ':資本金（単位：万円）
    Public m_iSihonkinY         ':資本金（単位：円）
    Public m_iJyugyoin_Suu      ':従業員
    Public m_iSyoninkyu         ':初任給
    Public m_sBiko              ':備考

    '進路のWhere条件
    Public m_sSinroWhere        ':進路の条件
    Public m_sSingakuWhere      ':進路の条件
    Public m_sSelected1         ':進路コンボの条件
    Public m_sSelected2         ':進学コンボの条件
    Public m_sSingakuOption     ':進学コンボのオプション

    Public m_sKenWhere          ':県の条件
    Public m_sSityoWhere        ':市町村コンボの条件
    Public m_sSityoOption       ':市町村コンボのオプション
    Public m_sKenSentakuWhere
    Public m_sSityoSentakuWhere

    'Public Const C_SYORYAKU_KETA=4'//表示時に省略する桁数（資本金）


'///////////////////////////メイン処理/////////////////////////////
    Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()

    '値の初期化
    Call s_Syokika

    '// ﾃﾞｰﾀﾍﾞｰｽ接続
    w_iRet = gf_OpenDatabase()
    If w_iRet <> 0 Then
        'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
        m_bErrFlg = True
        m_sErrMsg = "データベースとの接続に失敗しました。"
        Exit sub
    End If  

    '// 不正アクセスチェック
    Call gf_userChk(session("PRJ_No"))

    'パラメータセット
    Call s_ParaSet()

    '// 訂正ボタンを押されたときは、DBからデータ取得
    If m_sMode = "Syusei" and m_bReFlg = false Then
        call db_get()
    End If

    If m_sSchMode = "JyusyoSch" then
        Call f_SchJyusyo()
    End if

    '県に関するWHREを作成する
    Call f_MakeKenWhere()   

    '市町村に関するWHREを作成する
    Call f_MakeSityoWhere() 

    '// コンボ用where文作成
    Call f_MakeCommbo()

    'HTMLを作成する
    Call showPage()

    '// 終了処理
    call gf_closeObject(m_Rs)
    call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  DBでデータの取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub db_get()

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="進路マスタ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    Do

        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI_KANA "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_KEN_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SITYOSON_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_YUBIN_BANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINGAKU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_GYOSYU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SIHONKIN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JYUGYOIN_SUU "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SYONINKYU "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_BIKO "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & vbCrLf & "    AND M32_SINRO_CD = '" & m_sRenrakusakiCD & "' "

        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If

        '//レコードセットを変数に入れる
        Call s_Dataset()

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
        response.end
    End If
    

End Sub

'********************************************************************************
'*  [機能]  DBでデータの取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub f_SchJyusyo()

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="進路マスタ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & "     M12_SITYOSONMEI,  "
        w_sSQL = w_sSQL & "     M12_TYOIKIMEI "
        w_sSQL = w_sSQL & "FROM  "
        w_sSQL = w_sSQL & "     M12_SITYOSON "
        w_sSQL = w_sSQL & "WHERE "
        w_sSQL = w_sSQL & "     M12_YUBIN_BANGO = '" & Request("txtYUBINBANGO") & "'"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If

        Exit Do
    Loop

    if Not w_Rs.Eof then
        m_sJusyo1 = w_Rs("M12_SITYOSONMEI")
        m_sJusyo2 = w_Rs("M12_TYOIKIMEI")
        m_sJusyo3 = ""
    End if

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
        response.end
    End If
    
End Sub

'///////////////////////////値の初期化/////////////////////////////
Sub s_Syokika()

    m_sSinroCD          = ""
    m_sSingakuCD        = ""
    m_sSinromei         = ""
    'm_sSinromei_Eig    = ""
    m_sSinromei_Kan     = ""
    m_sSinromei_Rya     = ""
    'm_sJusyo           = ""
    m_sJusyo1           = ""
    m_sJusyo2           = ""
    m_sJusyo3           = ""
    m_sTel              = ""
    m_sSinro_URL        = ""
    m_sPageCD           = ""
    m_Rs                = ""
    m_sMode             = ""
    m_sSinroWhere       = ""
    m_sSingakuWhere     = ""
    m_sSelected1        = ""
    m_sSelected2        = ""
    m_sSingakuOption    = ""
    m_sSinroCD2         = ""
    m_sSingakuCD2       = ""
    m_iKenCd            = ""
    m_iSityoCd          = ""
    m_sYubin            = ""
    'm_iGyosyu_Kbn       = ""
    m_iSihonkin         = ""
    m_iSihonkinY        = ""
    m_iJyugyoin_Suu     = ""
    m_iSyoninkyu        = ""
    m_sBiko             = ""

End Sub


'********************************************************************************
'*  [機能]  引数を変数に代入
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ParaSet()

    m_sMode     = Request("txtMode")
    If m_sMode  = "" then m_sMode = "Sinki"

    m_sSchMode  = Request("txtSchMode")                     ':検索ﾓｰﾄﾞ

    m_bReFlg    = Request("txtReFlg")
    If m_bReFlg = "" then m_bReFlg = false

    m_sRenrakusakiCD = Request("txtRenrakusakiCD")          ':連絡先コード

    '/*仕様変更により使用しない
    'If m_sMode = "Sinki" Then
        'Call f_Max()
    'End If

    'm_sSinroCD3    = Request("txtSinroCD")                 ':進路コード
    m_sSinroCD      = gf_cboNull(Request("txtSinroCD"))     ':進路コード
    m_sSingakuCD    = gf_cboNull(Request("txtSingakuCD"))   ':進学コード
    m_sSyusyokuName = Request("txtSyusyokuName")            ':就職先名称（一部）

    m_sSinroCD2     = Request("txtSinroCD2")                ':「戻る」用の進路コード
    if m_sSinroCD   = "" then Request("txtSinroCD")

    m_sSingakuCD2   = Request("txtSingakuCD2")              ':「戻る」用の進学コード
    if m_sSingakuCD2= "" then Request("txtSingakuCD")

    m_sSINROMEI     = Request("txtSINROMEI")                ':進路名
    'm_sSinromei_Eig= Request("txtSINROMEI_EIGO")           ':進路名英語
    m_sSinromei_Kan = Request("txtSINROMEI_KANA")           ':進路名カナ
    m_sSinromei_Rya = Request("txtSINRORYAKSYO")            ':進路略称
    'm_sJusyo       = Request("txtJUSYO")                   ':住所

    If m_sSchMode = "" then
        m_sJusyo1       = Request("txtJUSYO1")              ':住所1
        m_sJusyo2       = Request("txtJUSYO2")              ':住所2
        m_sJusyo3       = Request("txtJUSYO3")              ':住所3
    End if

    m_sTel          = Request("txtDENWABANGO")              ':電話番号
    m_iNendo        = Session("NENDO")                      ':年度

    m_iKenCd        = Request("txtKenCd")                   ':県コード
    m_iSityoCd      = Request("txtSityoCd")                 ':市町村コード
    m_sYubin        = Request("txtYUBINBANGO")              ':郵便番号
    'm_iGyosyu_Kbn   = Request("txtGYOSYU_KBN")              ':業種区分
    m_iSihonkin     = Request("txtSIHONKIN")                ':資本金
    m_iJyugyoin_Suu = Request("txtJYUGYOIN_SUU")            ':従業員数
    m_iSyoninkyu    = Request("txtSYONINKYU")               ':初任給
    m_sBiko         = Request("txtBIKO")                    ':備考

    m_sSinro_URL    = Request("txtSINRO_URL")               ':URL
    if m_sSinro_URL = "" Then m_sSinro_URL = "http://"

    '//BLANKの場合は行数ｸﾘｱ
    If m_sMode = "Sinki" Then
        m_sPageCD = 1
    Else
        m_sPageCD = INT(Request("txtPageCD"))               ':表示済表示頁数（自分自身から受け取る引数）
    End If

End Sub

'********************************************************************************
'*  [機能]  DBの値を変数に代入
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_Dataset()

Dim w_iSihonkin

    m_sSinroCD      = m_Rs("M32_SINRO_KBN")

	If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
		'//進路先CDが1(進学)の場合、進学区分を取得
    	m_sSingakuCD    = gf_SetNull2String(m_Rs("M32_SINGAKU_KBN"))
	ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
		'//進路先CDが2(就職)の場合、業種区分を取得
    	m_sSingakuCD    = gf_SetNull2String(m_Rs("M32_GYOSYU_KBN"))
	End If

    m_sSinromei     = m_Rs("M32_SINROMEI")
    'm_sSinromei_Eig = m_Rs("M32_SINROMEI_EIGO")
    m_sSinromei_Kan = m_Rs("M32_SINROMEI_KANA")
    m_sSinromei_Rya = m_Rs("M32_SINRORYAKSYO")
    'm_sJusyo        = m_Rs("M32_JUSYO")
    m_sJusyo1        = m_Rs("M32_JUSYO1")
    m_sJusyo2        = m_Rs("M32_JUSYO2")
    if IsNull(m_Rs("M32_JUSYO3")) = False Then
        m_sJusyo3        = m_Rs("M32_JUSYO3")
    end if
    m_sTel          = m_Rs("M32_DENWABANGO")
    m_iKenCd        = m_Rs("M32_KEN_CD")
    m_iSityoCd        = m_Rs("M32_SITYOSON_CD")
    m_sYubin = m_Rs("M32_YUBIN_BANGO")

'    if IsNull(m_Rs("M32_GYOSYU_KBN")) = False Then
'        m_iGyosyu_Kbn = m_Rs("M32_GYOSYU_KBN")
'    end if

    if IsNull(m_Rs("M32_SIHONKIN")) = False Then
        m_iSihonkinY = m_Rs("M32_SIHONKIN")
        w_iSihonkin = CInt(Len(m_iSihonkinY)) - C_SYORYAKU_KETA
        m_iSihonkin = Mid(m_iSihonkinY,1,w_iSihonkin)
    end if
    if IsNull(m_Rs("M32_JYUGYOIN_SUU")) = False Then
        m_iJyugyoin_Suu = m_Rs("M32_JYUGYOIN_SUU")
    end if
    if IsNull(m_Rs("M32_SYONINKYU")) = False Then
        m_iSyoninkyu = m_Rs("M32_SYONINKYU")
    end if
    if IsNull(m_Rs("M32_BIKO")) = False Then
        m_sBiko = m_Rs("M32_BIKO")
    end if

    if IsNull(m_Rs("M32_SINRO_URL")) = False Then
        m_sSinro_URL = m_Rs("M32_SINRO_URL")
    else
        m_sSinro_URL = "http://"
    end if
    
    'if m_sSinro_URL = "" Then m_sSinro_URL = "http://"
    
End Sub


Sub f_MakeKenWhere()
'********************************************************************************
'*  [機能]  県コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sKenWhere=""
    m_sKenSentakuWhere=""
        m_sKenWhere = " M16_NENDO = '" & Session("NENDO") & "' "
        'm_sKenSentakuWhere = Request("txtKenCd")
        m_sKenSentakuWhere = m_iKenCd
End Sub

Sub f_MakeSityoWhere()
'********************************************************************************
'*  [機能]  市町村コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_sSityoWhere=""
    m_sSityoSentakuWhere = ""
    m_sSityoOption=""

    'If Request("txtKenCd") <> "" Then
    If m_iKenCd <> "" Then
        'm_sSityoWhere = "     M12_KEN_CD = '" & Request("txtKenCd") & "' "
        m_sSityoWhere = "     M12_KEN_CD = '" & m_iKenCd & "' "
        m_sSityoWhere = m_sSityoWhere & " GROUP BY M12_SITYOSON_CD,M12_SITYOSONMEI "
        m_sSityoSentakuWhere = m_iSityoCd
    Else
        m_sSityoOption = " DISABLED "
        m_sSityoWhere  = " M12_Ken_CD = '0' "
    End IF

End Sub

'********************************************************************************
'*  [機能]  進路コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub f_MakeCommbo()


    m_sSinroWhere = " M01_DAIBUNRUI_CD = "&C_SINRO&"  AND "
    m_sSinroWhere = m_sSinroWhere & " M01_NENDO = " & m_iNendo & ""

    If m_sMode = "Insert" Then
        m_sSingakuOption = "DISABLED"
        m_sSelected1      = ""
        m_sSelected2      = ""

    Else 

        m_sSelected1 = m_sSinroCD

		'// 進学
	    If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
	        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_SINGAKU & "  AND "
	        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
            m_sSelected2 = m_sSingakuCD

		'// 就職
		ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
	        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_GYOSYU_KBN & "  AND "
	        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
            m_sSelected2 = m_sSingakuCD

		'// その他
	    Else
	        m_sSingakuWhere= " M01_DAIBUNRUI_CD = 0 "
	        m_sSingakuOption = " DISABLED "
	    End IF


    End If

End Sub

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
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  進路が修正されたとき、再表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./syusei.asp";
        document.frm.target="";

        document.frm.txtReFlg.value ='true';
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  住所検索ボタン押されたとき
    //  [引数]  なし
    //  [戻値]  true:ﾁｪｯｸOK、false:ﾁｪｯｸｴﾗｰ
    //  [説明]  
    //************************************************************
    function jf_JyusyoSch(){

        if( f_Trim(document.frm.txtYUBINBANGO.value) == "" ){
            window.alert("郵便番号が入力されていません");
            document.frm.txtYUBINBANGO.focus();
            return false;
        }

        // ■■■文字数ﾁｪｯｸ■■■
        // ■進路先郵便番号
        var str = new String(document.frm.txtYUBINBANGO.value);
        if( getLengthB(str) != "8" ){
            window.alert("郵便番号は8文字入力してください");
            document.frm.txtYUBINBANGO.focus();
            return false;
        }

        var str = new String(document.frm.txtYUBINBANGO.value);
        if( f_Trim(str) != "" ){
          if( IsHankakuSujiHyphen(str) == false ){
              window.alert("郵便番号は半角数字とハイフンのみ入力してください");
              document.frm.txtYUBINBANGO.focus();
              return false;
          }
        }

        document.frm.txtSchMode.value = "JyusyoSch";
        document.frm.action="./syusei.asp";
        document.frm.target="fTopMain";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  郵便番号検索ボタン押されたとき
    //  [引数]  pMode = ﾓｰﾄﾞ
    //                  'SEARCH' = 検索
    //                  'DISPLAY'= 参照
    //  [戻値]  true:ﾁｪｯｸOK、false:ﾁｪｯｸｴﾗｰ
    //  [説明]  
    //************************************************************
    function jf_ZipCodeSch(pMode){
        var w_JUSYO1 = ""
        var w_JUSYO2 = ""


        // 検索ﾓｰﾄﾞの場合、住所が必要
        if( pMode == 'SEARCH' ){
            if( f_Trim(document.frm.txtJUSYO1.value) == "" && f_Trim(document.frm.txtJUSYO2.value) == "" ){
                window.alert("住所を入力してください");
                document.frm.txtJUSYO1.focus();
                return false;
            }
            w_JUSYO1 = document.frm.txtJUSYO1.value;
            w_JUSYO2 = document.frm.txtJUSYO2.value;
        }

        // サブウィンドーを開く
        w   = 520;
        h   = 520;
        url = "../../Common/com_select/SEL_JYUSYO/default.asp";
        wn  = "SubWindow";
        opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=yes";
        if (w > 0)
            opt = opt + ",width=" + w;
        if (h > 0)
            opt = opt + ",height=" + h;
        newWin = window.open(url, wn, opt);

		// 検索の場合は、住所を送る（検索条件）
		if ( pMode == 'SEARCH' ){
	        document.frm.action="../../Common/com_select/SEL_JYUSYO/default.asp";
	        document.frm.target="SubWindow";
	        document.frm.submit();
		}

        // window移動
        x   = (screen.availWidth - w) / 2;
        y   = (screen.availHeight - h) / 2;
        newWin.moveTo(x, y);

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
        // ■連絡先コード
        if( f_Trim(document.frm.txtRenrakusakiCD.value) == "" ){
            window.alert("連絡先コードが入力されていません");
            document.frm.txtRenrakusakiCD.focus();
            return false;
        }

        // ■■■数値妥当性ﾁｪｯｸ■■■
        // ■連絡先コード数値
        if( isNaN(document.frm.txtRenrakusakiCD.value) ){
            window.alert("連絡先コードには数値を入力してください");
            document.frm.txtRenrakusakiCD.focus();
            return false;
        }

        // ■■■NULLﾁｪｯｸ■■■
        // ■名称
        if( f_Trim(document.frm.txtSINROMEI.value) == "" ){
            window.alert("名称が入力されていません");
            document.frm.txtSINROMEI.focus();
            return false;
        }

        // ■■■文字数ﾁｪｯｸ■■■
        // ■名称
        var str = new String(document.frm.txtSINROMEI.value);
        if( getLengthB(str) > "60" ){
            window.alert("名称は全角30文字で入力してください");
            document.frm.txtSINROMEI.focus();
            return false;
        }

        // ■■■文字数ﾁｪｯｸ■■■
        // ■名称
        var str = new String(document.frm.txtSINROMEI_KANA.value);
        if( getLengthB(str) > "60" ){
            window.alert("名称は全角30文字以内で入力してください");
            document.frm.txtSINROMEI_KANA.focus();
            return false;
        }

        // ■■■文字数ﾁｪｯｸ■■■
        // ■略称
        var str = new String(document.frm.txtSINRORYAKSYO.value);
        if( getLengthB(str) > "10" ){
            window.alert("略称は全角5文字以内で入力してください");
            document.frm.txtSINRORYAKSYO.focus();
            return false;
        }

        // ■■■NULLﾁｪｯｸ■■■
        // ■進路区分
        if( f_Trim(document.frm.txtSinroCD.value) == "@@@" ){
            window.alert("進路区分が入力されていません");
            document.frm.txtSinroCD.focus();
            return false;
        }

        if(document.frm.txtSinroCD.value == '1') {

                // ■■■NULLﾁｪｯｸ■■■
                // ■進学区分
                if( f_Trim(document.frm.txtSingakuCD.value) == "@@@" ){
                    window.alert("進学区分が入力されていません");
                    document.frm.txtSingakuCD.focus();
                    return false;
                }
        };


        // ■■■文字数ﾁｪｯｸ■■■
        // ■進路先郵便番号
        var str = new String(document.frm.txtYUBINBANGO.value);
        if( getLengthB(str) != "8" ){
            window.alert("郵便番号は8文字入力してください");
            document.frm.txtYUBINBANGO.focus();
            return false;
        }

        var str = new String(document.frm.txtYUBINBANGO.value);
        if( f_Trim(str) != "" ){
          if( IsHankakuSujiHyphen(str) == false ){
              window.alert("郵便番号は半角数字とハイフンのみ入力してください");
              document.frm.txtYUBINBANGO.focus();
              return false;
          }
        }
        // ■■■NULLﾁｪｯｸ■■■
        // ■県コード
//        if( f_Trim(document.frm.txtKenCd.value) == "@@@" ){
//          window.alert("都道府県が選択されていません");
//          document.frm.txtKenCd.focus();
//          return false;
//        }

        // ■■■NULLﾁｪｯｸ■■■
        // ■住所（１）
        if( f_Trim(document.frm.txtJUSYO1.value) == "" ){
            window.alert("住所（１）が入力されていません");
            document.frm.txtJUSYO1.focus();
            return false;
        }

        // ■■■文字数ﾁｪｯｸ■■■
        // ■住所（１）
        var str = new String(document.frm.txtJUSYO1.value);
        if( getLengthB(str) > "40" ){
            window.alert("住所は全角20文字以内で入力してください");
            document.frm.txtJUSYO1.focus();
            return false;
        }

        // ■■■NULLﾁｪｯｸ■■■
        // ■住所（２）
        if( f_Trim(document.frm.txtJUSYO2.value) == "" ){
            window.alert("住所（２）が入力されていません");
            document.frm.txtJUSYO2.focus();
            return false;
        }
        // ■■■文字数ﾁｪｯｸ■■■
        // ■住所（２）
        var str = new String(document.frm.txtJUSYO2.value);
        if( getLengthB(str) > "40" ){
            window.alert("住所は全角20文字以内で入力してください");
            document.frm.txtJUSYO2.focus();
            return false;
        }

        // ■■■文字数ﾁｪｯｸ■■■
        // ■住所（３）
        var str = new String(document.frm.txtJUSYO3.value);
        if( getLengthB(str) > "40" ){
            window.alert("住所は全角20文字以内で入力してください");
            document.frm.txtJUSYO3.focus();
            return false;
        }

        // ■■■NULLﾁｪｯｸ■■■
        // ■進路先電話番号
        if( f_Trim(document.frm.txtDENWABANGO.value) == "" ){
            window.alert("電話番号が入力されていません");
            document.frm.txtDENWABANGO.focus();
            return false;
        }
        // ■■■文字数ﾁｪｯｸ■■■
        // ■進路先電話番号
        var str = new String(document.frm.txtDENWABANGO.value);
        if( getLengthB(str) > "15" ){
            window.alert("電話番号は15文字以内で入力してください");
            document.frm.txtDENWABANGO.focus();
            return false;
        }
        // ■■■文字ﾁｪｯｸ■■■
        var str = new String(document.frm.txtDENWABANGO.value);
        if( f_Trim(str) != "" ){
          if( IsHankakuSujiHyphen(str) == false ){
              window.alert("電話番号は半角数字とハイフンのみ入力してください");
              document.frm.txtDENWABANGO.focus();
              return false;
          }
        }
    
        // ■■■文字数ﾁｪｯｸ■■■
        // ■URL
        var str = new String(document.frm.txtSINRO_URL.value);
        if( f_Trim(str) != "" ){
            if( getLengthB(str) > "40" ){
                window.alert("URLは40文字以内で入力してください");
                document.frm.txtSINRO_URL.focus();
                return false;
            }
        }

<%
'//        // ■■■数値妥当性ﾁｪｯｸ■■■
'//        // ■業種区分
'//        if( f_Trim(document.frm.txtGYOSYU_KBN.value) != "" ){
'//            if( isNaN(document.frm.txtGYOSYU_KBN.value) ){
'//                window.alert("業種区分は数値を半角で入力してください");
'//                document.frm.txtGYOSYU_KBN.focus();
'//                return false;
'//            }
'//        }
'//
'//        // ■■■桁ﾁｪｯｸ■■■
'//        // ■業種区分
'//        if( f_Trim(document.frm.txtGYOSYU_KBN.value) != "" ){
'//            var str = new String(document.frm.txtGYOSYU_KBN.value);
'//            if( str.length > 2 ){
'//                window.alert("業種区分の入力値は2桁以内にしてください");
'//                document.frm.txtGYOSYU_KBN.focus();
'//                return false;
'//            }
'//        }
%>
        // ■■■数値妥当性ﾁｪｯｸ■■■
        // ■資本金
        if( f_Trim(document.frm.txtSIHONKIN.value) != "" ){
            if( isNaN(document.frm.txtSIHONKIN.value) ){
                window.alert("資本金は数値を半角で入力してください");
                document.frm.txtSIHONKIN.focus();
                return false;
            }
        }
        
        // ■■■文字数ﾁｪｯｸ■■■
        // ■資本金
        var str = new String(document.frm.txtSIHONKIN.value);
        if( getLengthB(str) > "7" ){
            window.alert("資本金は半角7桁以内で入力してください");
            document.frm.txtSIHONKIN.focus();
            return false;
        }
        
        // ■■■数値妥当性ﾁｪｯｸ■■■
        // ■従業員数
        if( f_Trim(document.frm.txtJYUGYOIN_SUU.value) != "" ){
            if( isNaN(document.frm.txtJYUGYOIN_SUU.value) ){
                window.alert("従業員数は数値を半角で入力してください");
                document.frm.txtJYUGYOIN_SUU.focus();
                return false;
            }
        }
        
        // ■■■桁ﾁｪｯｸ■■■
        // ■従業員数
        var str = new String(document.frm.txtJYUGYOIN_SUU.value);
        if( str.length > 7 ){
            window.alert("従業員数の入力値は7桁以内にしてください");
            document.frm.txtJYUGYOIN_SUU.focus();
            return false;
        }
        
        // ■■■数値妥当性ﾁｪｯｸ■■■
        // ■初任給
        if( f_Trim(document.frm.txtSYONINKYU.value) != "" ){
            if( isNaN(document.frm.txtSYONINKYU.value) ){
                window.alert("初任給は数値を半角で入力してください");
                document.frm.txtSYONINKYU.focus();
                return false;
            }
        }
        
        // ■■■桁ﾁｪｯｸ■■■
        // ■初任給
        var str = new String(document.frm.txtSYONINKYU.value);
        if( str.length > 7 ){
            window.alert("初任給の入力値は7桁以内にしてください");
            document.frm.txtSYONINKYU.focus();
            return false;
        }
        
        // ■■■桁ﾁｪｯｸ■■■
        // ■備考
        if( getLengthB(document.frm.txtBIKO.value) > "100" ){
            window.alert("備考の欄は全角50文字以内で入力してください");
            document.frm.txtBIKO.focus();
            return false;
        }

        document.frm.action="./kakunin.asp";
        document.frm.target="_self";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    </head>

    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <div align="center">
<!--    <form name="frm" action="kakunin.asp" target="_self" Method="POST">-->
    <form name="frm" Method="POST">
    <%
        If m_sMode = "Sinki" Then
          m_sSubtitle = "新規登録"
        else
          m_sSubtitle = "修　正"
        End If

        call gs_title("進路先情報登録",m_sSubtitle)
    %>

    <br>
    進　路　先　情　報
    <br><br>
<table border="0" cellpadding="1" cellspacing="1">
    <tr>
        <td align="left">
            <table width="100%" border=1 CLASS="hyo">
	        <colgroup valign="top">
	        <colgroup valign="top">
		        <tr>
		            <th class=header width="100">進路先コード</th>
		            <td nowrap class=detail align="left">
		            <% If m_sMode <> "Sinki" Then %>
		                <%=m_sRenrakusakiCD%>
		                <input type="hidden" name="txtRenrakusakiCD" value="<%= m_sRenrakusakiCD %>">
		            <% Else %>
		                <input type="text" name="txtRenrakusakiCD" value="<%= m_sRenrakusakiCD %>" MAXLENGTH=6 size=8>
		                <span class=hissu>*</span>（半角数字6桁以内）
		            <% End If %>
		            </td>
		        </tr>
		        <tr>
		            <th class=header>名　称</th>
		            <td nowrap class=detail><input type="text" size="64" name="txtSINROMEI" value="<%= m_sSinromei %>" MAXLENGTH=60 size=30><span class=hissu>*</span>（全角30文字以内）</td>
		        </tr>
		        <tr>
		            <th class=header>名　称（カナ）</th>
		            <td nowrap class=detail><input type="text" size="64" name="txtSINROMEI_KANA" value="<%= m_sSinromei_Kan %>" MAXLENGTH=60 size=30>（全角30文字以内）</td>
		        </tr>
		        <tr>
		            <th class=header>略　称</th>
		            <td nowrap class=detail><input type="text" size="20" name="txtSINRORYAKSYO" value="<%= m_sSinromei_Rya %>" MAXLENGTH=10 size=10>（全角5文字以内）</td>
		        </tr>
		        <tr>
		            <th class=header>進路区分</th>
		            <td nowrap class=detail>
		                <%  '共通関数から進路に関するコンボボックスを出力する
		                    call gf_ComboSet("txtSinroCD",C_CBO_M01_KUBUN,m_sSinroWhere,"onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sSelected1)
		                %><span class=hissu>*</span>
		            </td>
		        </tr>
		        <tr>
		            <th class=header>種別区分</th>
		            <td nowrap class=detail>
		                <% '共通関数から進学に関するコンボボックスを出力する（進路区分が条件）（進路区分が1ではないときは、DISABLEDとなる）
		                    call gf_ComboSet("txtSingakuCD",C_CBO_M01_KUBUN,m_sSingakuWhere,"style='width=100px' "&m_sSingakuOption,True,m_sSelected2)
		                %>

		                <% If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then %>
		                    <span class=hissu>*</span>
		                <% End If %>
		            </td>
		        </tr>
		        <tr>
		            <th class=header>郵便番号</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtYUBINBANGO" value="<%= m_sYubin %>" MAXLENGTH=8><span class=hissu>*</span>（例:000-0000）
		                <img src="../../image/sp.gif" width="100" height="1">
		                <input type="button" class="button" name="btnJyusyoSch" value="〒 → 住所" onClick="javascript:return jf_JyusyoSch();">
		                <input type="button" class="button" name="btnZipSch"    value="住所 → 〒" onClick="javascript:return jf_ZipCodeSch('SEARCH');">
		            </td>
		        </tr>
		        <tr>
		            <th class=header>住　所（１）</th>
		            <td nowrap class=detail><input type="text" size="44" name="txtJUSYO1" value="<%= m_sJusyo1 %>" MAXLENGTH=40 size=40><span class=hissu>*</span>（全角20文字以内）
		                <img src="../../image/sp.gif" width="10" height="1"><input type="button" class="button" name="btnJyusyoDsp" value="参照" onClick="javascript:return jf_ZipCodeSch('DISPLAY');"></td>
		        </tr>
		        <tr>
		            <th class=header>住　所（２）</th>
		            <td nowrap class=detail><input type="text" size="44" name="txtJUSYO2" value="<%= m_sJusyo2 %>" MAXLENGTH=40 size=40><span class=hissu>*</span>（全角20文字以内）</td>
		        </tr>
		        <tr>
		            <th class=header>住　所（３）</th>
		            <td nowrap class=detail><input type="text" size="44" name="txtJUSYO3" value="<%= m_sJusyo3 %>" MAXLENGTH=40 size=40>（建物名等を記入・全角20文字以内）</td>
		        </tr>
		        <tr>
		            <th class=header>電話番号</th>
		            <td nowrap class=detail><input type="text" size="20" name="txtDENWABANGO" value="<%= m_sTel %>" MAXLENGTH=15 size=15><span class=hissu>*</span>（例:000-000-0000）</td>
		        </tr>
		        <tr>
		            <th class=header>ＵＲＬ</th>
		            <td nowrap class=detail><input type="text" size="50" name="txtSINRO_URL"  value="<%= m_sSinro_URL %>" MAXLENGTH=40></td>
		        </tr>
<!--
		        <tr>
		            <th class=header>業種区分</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtGYOSYU_KBN"  value="<%= m_iGyosyu_Kbn %>" MAXLENGTH=2>（半角数字6桁以内）</td>
		        </tr>
-->
		        <tr>
		            <th class=header>資本金</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtSIHONKIN"  value="<%= m_iSihonkin %>" MAXLENGTH="7">万円</td>
		        </tr>
		        <tr>
		            <th class=header>従業員数</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtJYUGYOIN_SUU"  value="<%= m_iJyugyoin_Suu %>" MAXLENGTH=7>人</td>
		        </tr>
		        <tr>
		            <th class=header>初任給</th>
		            <td nowrap class=detail><input type="text" size="10" name="txtSYONINKYU"  value="<%= m_iSyoninkyu %>" MAXLENGTH=7>円</td>
		        </tr>
		        <tr>
		            <th class=header>備　考</th>
		            <td nowrap class=detail><textarea rows=3 cols=40 name="txtBIKO"><%= m_sBiko %></textarea>（全角50文字以内）</td>
                </TR>
            </TABLE>
		    <table width=100%><tr><td align=right><span class=hissu>*印は必須項目です。</span></td></tr></table>
		    <br>
        </td>
    </TR>
	</TABLE>
    <table border="0" >
        <tr>
            <td valign="top" align=left>
                <input type="button" class="button" value="　登　録　" Onclick="return f_CheckData()">
                <input type="hidden" name="txtKenCd"   value="<%= m_iKenCd %>">
                <input type="hidden" name="txtSityoCd" value="<%= m_iSityoCd %>">
                <input type="hidden" name="txtRenban"  value="">
                <input type="hidden" name="txtMode" value="<%= m_sMode %>">
                <input type="hidden" name="txtReFlg" value="<%= m_bReFlg %>">
                <input type="hidden" name="txtSinroCD2" value="<%= m_sSinroCD2 %>">
                <input type="hidden" name="txtSingakuCD2" value="<%= m_sSingakuCD2 %>">
                <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
                <!--<input type="hidden" name="txtNendo" value="<%= Session("SYORI_NENDO") %>">-->
                <input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
                <input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
                <input type="hidden" NAME="ButtonClick" value="">
                <input type="hidden" NAME="txtSchMode">
                </form>
            </td>
			<td><img src="../../image/sp.gif" width="20" height="1"></td>
            <td valign="top" align=right>
                <form action=default.asp name="cansel" method=post target="<%=C_MAIN_FRAME%>">
                    <input type="hidden" name="txtMode" value="search">
                    <input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD2 %>">
                    <input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD2 %>">
                    <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
                    <input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
                    <input class=button type='submit' value='キャンセル'>
                </form>
            </td>
        </tr>
    </table>

</div>
</body>
</html>


<%
    '---------- HTML END   ----------
End Sub
%>