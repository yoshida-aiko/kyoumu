<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0144/update.asp
' 機      能: 下ページ 就職先マスタの詳細変更を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSinroKBN     :進路コード
'           txtSingakuCd        :進学コード
'           txtSyusyokuName     :進路名称（一部）
'           txtPageSinro        :表示済表示頁数（自分自身から受け取る引数）
'           Sinro_syuseiCD      :選択された進路コード
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtSinroKBN     :進路コード（戻るとき）
'           txtSingakuCd        :進学コード（戻るとき）
'           txtSyusyokuName     :進路名称（戻るとき）
'           txtPageSinro        :表示済表示頁数（戻るとき）
' 説      明:
'           ■初期表示
'               指定された進学先・就職先の詳細データを表示
'           ■地図画像ボタンクリック時
'               指定した条件にかなう進学先・就職先を表示する（別ウィンドウ）
'-------------------------------------------------------------------------
' 作      成: 2001/06/22 岩下 幸一郎
' 変      更: 2001/07/13 谷脇　良也
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数

    Public  m_sRenrakusakiCD        ':連絡先コード
    Public  m_sSINROMEI             ':進路名
    'Public m_sSINROMEI_EIGO        ':進路名英語
    Public  m_sSINROMEI_KANA        ':進路名カナ
    Public  m_sSINRORYAKSYO         ':進路略称
    'Public m_sJUSYO                ':住所
    Public  m_sJUSYO1               ':住所1
    Public  m_sJUSYO2               ':住所2
    Public  m_sJUSYO3               ':住所3
    Public  m_iKenCd                ':県コード
    Public  m_iSityoCd              ':市町村コード
    Public  m_sDENWABANGO           ':電話番号
    Public  m_sSinro_syuseiCD       ':進路区分
    Public  m_sSINRO_URL            ':URL
    Public  m_iNendo        ':年度
    Public  m_sDATE
    Public  m_sKyokanCD
    Public  m_sMode
    Public  m_sYubin            ':郵便番号
    Public  m_iGyosyu_Kbn       ':業種区分
    Public  m_iSihonkin         ':資本金（単位：万円）
    Public  m_iSihonkinY         ':資本金（単位：円）
    Public  m_iJyugyoin_Suu     ':従業員
    Public  m_iSyoninkyu        ':初任給
    Public  m_sBiko             ':備考

    Public  m_sSinroCD      ':進路コード
    Public  m_sSingakuCD        ':進学コード
    Public  m_sSinroCD2     ':進路コード
    Public  m_sSingakuCD2       ':進学コード
    Public  m_sSyusyokuName     ':進路名称（一部）
    Public  m_iPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_Rs            'recordset


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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="就職先マスタ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
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
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
        
        '//ﾄﾗﾝｻﾞｸｼｮﾝ開始
        Call gs_BeginTrans()
        
        '// 処理の振り分け
        If m_sMode = "Sinki" then
            call s_ins(w_sSQL)
        Else
            call s_update(w_sSQL)
        End If

'Response.Write w_sSQL & "<br>"

		w_iRet = gf_ExecuteSQL(w_sSQL)
        If w_iRet <> 0 Then
            '失敗
            '//ﾛｰﾙﾊﾞｯｸ
            Call gs_RollbackTrans()
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If
        
        '//ｺﾐｯﾄ
        Call gs_CommitTrans()

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

'/* 新規登録
Sub s_ins(p_sSQL)

        p_sSQL = p_sSQL & vbCrLf & " Insert Into "
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO"
        p_sSQL = p_sSQL & "(M32_NENDO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_CD,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI,"
        'p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI_EIGO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI_KANA,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRORYAKSYO,"
        'p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO1,"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO2,"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO3,"
        p_sSQL = p_sSQL & vbCrLf & " M32_KEN_CD,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SITYOSON_CD,"
        p_sSQL = p_sSQL & vbCrLf & " M32_DENWABANGO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_YUBIN_BANGO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_KBN,"

		If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
			'//進路CDが1(進学)の場合
	        p_sSQL = p_sSQL & vbCrLf & " M32_SINGAKU_KBN,"
		ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
			'//進路CDが2(就職)の場合
	        p_sSQL = p_sSQL & vbCrLf & " M32_GYOSYU_KBN,"
		End If

        p_sSQL = p_sSQL & vbCrLf & " M32_SIHONKIN,"
        p_sSQL = p_sSQL & vbCrLf & " M32_JYUGYOIN_SUU,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SYONINKYU,"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_URL,"
        p_sSQL = p_sSQL & vbCrLf & " M32_BIKO,"
        p_sSQL = p_sSQL & vbCrLf & " M32_INS_DATE,"
        p_sSQL = p_sSQL & vbCrLf & " M32_INS_USER,"
        p_sSQL = p_sSQL & vbCrLf & " M32_UPD_DATE,"
        p_sSQL = p_sSQL & vbCrLf & " M32_UPD_USER)"
        p_sSQL = p_sSQL & vbCrLf & " Values"
        p_sSQL = p_sSQL & "(" & m_iNendo & ","
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sRenrakusakiCD) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sSINROMEI) & "',"
        'p_sSQL = p_sSQL & vbCrLf & "'" & m_sSINROMEI_EIGO & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sSINROMEI_KANA) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sSINRORYAKSYO) & "',"
        'p_sSQL = p_sSQL & vbCrLf & "'" & m_sJUSYO & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sJUSYO1) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sJUSYO2) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sJUSYO3) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_iKenCd) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_iSityoCd) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sDENWABANGO) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sYubin) & "',"
        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_sSinroCD) & ","

		If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
			'//進路CDが1(進学)の場合
	        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_sSingakuCD) & ","
		ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
			'//進路CDが2(就職)の場合
			If trim(replace(m_sSingakuCD,"@@@","")) <> "" Then
		        p_sSQL = p_sSQL & vbCrLf & "" & trim(m_sSingakuCD) & ","
			Else
		        p_sSQL = p_sSQL & vbCrLf & " NULL,"
			End If

		End If

        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_iSihonkinY) & ","
        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_iJyugyoin_Suu) & ","
        p_sSQL = p_sSQL & vbCrLf & " " & trim(m_iSyoninkyu) & ","
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sSINRO_URL) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sBiko) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & trim(m_sDATE) & "',"
        p_sSQL = p_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "',"
        p_sSQL = p_sSQL & vbCrLf & "'',"
        p_sSQL = p_sSQL & vbCrLf & "'')"
End Sub

Sub s_update(p_sSQL)
        p_sSQL = ""
        p_sSQL = p_sSQL & vbCrLf & " Update "
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO M32"
        p_sSQL = p_sSQL & vbCrLf & " Set "
        p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI         = '" & trim(m_sSINROMEI) & "',"
        'p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI_EIGO   = '" & m_sSINROMEI_EIGO & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINROMEI_KANA    = '" & trim(m_sSINROMEI_KANA) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRORYAKSYO     = '" & trim(m_sSINRORYAKSYO) & "',"
        'p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO           = '" & m_sJUSYO & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO1           = '" & trim(m_sJUSYO1) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO2           = '" & trim(m_sJUSYO2) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_JUSYO3           = '" & trim(m_sJUSYO3) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_KEN_CD           = '" & trim(m_iKenCd) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_SITYOSON_CD      = '" & trim(m_iSityoCd) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_DENWABANGO       = '" & trim(m_sDENWABANGO) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_YUBIN_BANGO      = '" & trim(m_sYubin) & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_KBN        =  " & trim(m_sSinroCD) & ","

		If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
			'//進路CDが1(進学)の場合
	        p_sSQL = p_sSQL & vbCrLf & " M32_SINGAKU_KBN     =  " & trim(m_sSingakuCD) & ","
		ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
			'//進路CDが2(就職)の場合
			If trim(replace(m_sSingakuCD,"@@@","")) <> "" Then
		        p_sSQL = p_sSQL & vbCrLf & " M32_GYOSYU_KBN      = " & trim(m_sSingakuCD) & ","
			Else
		        p_sSQL = p_sSQL & vbCrLf & " M32_GYOSYU_KBN      = NULL,"
			End If
		End If

        p_sSQL = p_sSQL & vbCrLf & " M32_SIHONKIN      =  " & trim(m_iSihonkinY)    & ","
        p_sSQL = p_sSQL & vbCrLf & " M32_JYUGYOIN_SUU  =  " & trim(m_iJyugyoin_Suu) & ","
        p_sSQL = p_sSQL & vbCrLf & " M32_SYONINKYU     =  " & trim(m_iSyoninkyu)    & ","
        p_sSQL = p_sSQL & vbCrLf & " M32_SINRO_URL     = '" & trim(m_sSINRO_URL)    & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_BIKO          = '" & trim(m_sBiko)         & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_UPD_DATE      = '" & trim(m_sDATE)         & "',"
        p_sSQL = p_sSQL & vbCrLf & " M32_UPD_USER      = '" & Session("LOGIN_ID")   & "'"
        p_sSQL = p_sSQL & vbCrLf & " WHERE "
        p_sSQL = p_sSQL & vbCrLf & "    M32_NENDO      =  " & m_iNendo              & " AND "
        p_sSQL = p_sSQL & vbCrLf & " M32.M32_SINRO_CD  = '" & m_sRenrakusakiCD      &"' "

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_sMode = Request("txtMode")

    m_sRenrakusakiCD = Request("txtRenrakusakiCD")          ':連絡先コード

    m_sSINROMEI      = Request("txtSINROMEI")       ':進路名
    'm_sSINROMEI_EIGO = Request("txtSINROMEI_EIGO")     ':進路名英語
    'If m_sSINROMEI_EIGO="　" Then m_sSINROMEI_EIGO=""
    m_sSINROMEI_KANA = Request("txtSINROMEI_KANA")      ':進路名英語
    m_sSINRORYAKSYO  = Request("txtSINRORYAKSYO")       ':進路略称
    If m_sSINRORYAKSYO="　" Then m_sSINRORYAKSYO=""
    'm_sJUSYO         = Request("txtJUSYO")         ':住所
    m_sJUSYO1         = Request("txtJUSYO1")            ':住所1
    m_sJUSYO2         = Request("txtJUSYO2")            ':住所2
    m_sJUSYO3         = Request("txtJUSYO3")            ':住所3
    m_iKenCd          = Request("txtKenCd")             ':県コード
    m_iSityoCd        = Request("txtSityoCd")           ':市町村コード
    m_sDENWABANGO    = Request("txtDENWABANGO")     ':電話番号
    m_sSINRO_URL     = Request("txtSINRO_URL")      ':URL
    If m_sSINRO_URL="　" Then m_sSINRO_URL=""
    m_sSinroCD = Request("txtSinroCD")          ':進路区分
    m_sSingakuCD = Request("txtSingakuCD")          ':進学区分

    If m_sSingakuCD ="" Then m_sSingakuCD = 0   'コンボ未選択時

    m_sDate = gf_YYYY_MM_DD(date(),"/")
    m_sKyokanCD = Session("KYOKAN_CD")          ':ユーザーID
    m_iNendo = Request("txtNendo")              ':年度
    m_sYubin = Request("txtYUBINBANGO")              ':郵便番号
    m_iGyosyu_Kbn = Request("txtGYOSYU_KBN")              ':業種区分
    m_iSihonkin = gf_SetNull2Zero(Request("txtSIHONKIN"))              ':資本金（単位：万円）
    if m_iSihonkin <> "" Then
        m_iSihonkinY = m_iSihonkin & "0000"              ':資本金（単位：円）
    end if
    m_iJyugyoin_Suu = gf_SetNull2Zero(Request("txtJYUGYOIN_SUU"))              ':従業員数
    m_iSyoninkyu = gf_SetNull2Zero(Request("txtSYONINKYU"))              ':初任給
    m_sBiko = Request("txtBIKO")              ':備考

    m_sSinroCD2 = Request("txtSinroCD2")        ':進路コード（「戻る」時に使用）
    m_sSingakuCD2 = Request("txtSingakuCD2")    ':進学コード（「戻る」時に使用）

    m_sSyusyokuName = Request("txtSyusyokuName")            ':就職先名称（一部）

End Sub

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
    
<%
If m_sMode = "Syusei" Then
    response.write "window.alert('" & C_TOUROKU_OK_MSG & "');"
Else
    response.write "window.alert('" & C_TOUROKU_OK_MSG & "');"
End If
%>
            document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>

<body bgcolor="#ffffff" onLoad="gonext()">
<center>
<form name="frm" action="./default.asp" target="<%=C_MAIN_FRAME%>" method=post>
<input type="hidden" name="txtMode" value="search">
<input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtPageCD" value="<%= m_iPageCD %>">
</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>