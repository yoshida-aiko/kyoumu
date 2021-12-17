<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験実施科目登録
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0130/skn0130_db.asp
' 機      能: 試験実施科目の登録・削除を行なう
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
'           ■DB部分のみ
'               全画面からわたってきたﾃﾞｰﾀの更新・削除を行なう
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 本村 文
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_sDBMode           'DBのﾓｰﾄﾞの設定
    Public  m_sMode             'モードの設定
    Public  m_iKyokanCd         ':教官コード
    Public  m_iSyoriNen         ':処理年度
    Public  m_iSikenKbn         ':試験区分
    Public  m_iSikenCode        ':試験ｺｰﾄﾞ
    Public  m_sGakunen          ':学年
    Public  m_sClass            ':ｸﾗｽNo
    Public  m_sKamoku           ':科目ｺｰﾄﾞ
'   Public  m_sJissiFLG         ':実施ﾌﾗｸﾞ
    Public  m_iMain_FLG         ':メイン教官
    Public  m_iSeiseki_FLG      ':成績入力教官フラグ
    Public  m_iJISSI_FLG        ':実施ﾌﾗｸﾞ
    Public  m_sJissiDate        ':実施日付
    Public  m_sJikan            ':実施時間
    Public  m_sKyositu          ':実施教室
    Public  m_sKokan1           ':成績入力教官１
    Public  m_sKokan2           ':成績入力教官２
    Public  m_sKokan3           ':成績入力教官３
    Public  m_sKokan4           ':成績入力教官４
    Public  m_sKokan5           ':成績入力教官５
    Public  m_sPageCD           ':ﾍﾟｰｼﾞナンバー

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

        '// DB登録
'        If m_sMode = "Delete" then
'            If Not f_Delete() then Exit do
'		End If

'        If m_sMode = "Update" then
            If Not f_Update() then Exit do
'		End If

'        If m_sDBMode = "T26" then
'            If Not f_Update() then Exit do
'        End If
'
'        If m_sDBMode = "T27" then
'            If Not f_Insert() then Exit do
'		End If

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

    m_sNendo     = Request("txtNendo")      '年度の取得
    m_sKyokanCD  = Session("KYOKAN_CD")     ':ユーザーID
    m_sPageCD    = Request("txtPageCD")     ': ﾍﾟｰｼﾞナンバー
    m_sMode      = Request("txtMode")       'ﾓｰﾄﾞの設定

    m_sDBMode    = Request("txtDBMode")         'DBのﾓｰﾄﾞの設定
    m_iKyokanCd = Session("KYOKAN_CD")         ':教官コード
    m_iSyoriNen     = Session("NENDO")         ':処理年度
    m_iSikenKbn  = Request("txtSikenKbn")           ':試験区分
    m_iSikenCode     = Request("txtSikenCode")      ':試験ｺｰﾄﾞ
    m_sGakunen   = Request("txtGakunen")            ':学年
    m_sClass     = Request("txtClass")          ':ｸﾗｽNo
    m_sKamoku    = Request("txtKamoku")         ':科目ｺｰﾄﾞ
    m_iMain_FLG    = Request("txtMainF")         ':メイン教官
    m_iSeiseki_FLG    = Request("txtSeisekiF")   ':成績入力教官フラグ
    m_sJissiFLG      = Request("txtJissiFLG")       ':実施ﾌﾗｸﾞ
    m_iJISSI_FLG = gf_SetNull2Zero(Request("chk1")) ':実施ﾌﾗｸﾞ
    m_sJikan     = Request("txtJikan")          ':実施時間
    m_sKyositu   = Request("txtKyositu")            ':実施教室
    m_sKokan1    = Request("SKyokanCd1")            ':成績入力教官１
    m_sKokan2    = Request("SKyokanCd2")            ':成績入力教官２
    m_sKokan3    = Request("SKyokanCd3")            ':成績入力教官３
    m_sKokan4    = Request("SKyokanCd4")            ':成績入力教官４
    m_sKokan5    = Request("SKyokanCd5")            ':成績入力教官５

    If strErrmsg <> "" Then
        ' エラーを表示するファンクション
        Call err_page(strErrMsg)
        response.end
    End If
'   call s_viewForm(request.form)   'デバッグ用　引数の内容を見る
End Sub


'********************************************************************************
'*  [機能]  更新処理
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
function f_Update()
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    f_Update = False

	'//実施しない場合、時間･教室にNULLを入れる
	if gf_IsNull(m_sJikan) then m_sJikan = "Null"
	if gf_IsNull(m_sKyositu) then m_sKyositu = "Null"
	if m_sKyositu = C_CBO_NULL then m_sKyositu = "Null"

	w_Cls = split(m_sClass,"#")

'Response.Write "UPD<br>" & w_Cls(0) & "<br><br>"
'Response.end

	i = 0
	For i = 0 to UBound(w_Cls) 

		'試験時間割に該当データがある場合は、時間割データの更新、無い場合は新規追加
		If f_GetSJikanKensu(w_Cls(i)) > 0 Then
		
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " Update T26_SIKEN_JIKANWARI SET "
	
			If m_iMain_FLG = "1" And m_sJissiFLG <> 1 then '//入力期間をみる。
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_JISSI_FLG = " & m_iJISSI_FLG &","
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKEN_JIKAN = " & m_sJikan & ","
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_KYOSITU = " & m_sKyositu & ","
			End If 

			If m_iSeiseki_FLG = "1" then
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN1 = '" & m_sKokan1 & "',"
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN2 = '" & m_sKokan2 & "',"
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN3 = '" & m_sKokan3 & "',"
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN4 = '" & m_sKokan4 & "',"
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN5 = '" & m_sKokan5 & "', "
			End If

			w_sSQL = w_sSQL & vbCrLf & " 	T26_UPD_DATE = TO_CHAR(SYSDATE,'YYYY/MM/DD'), "
			w_sSQL = w_sSQL & vbCrLf & " 	T26_UPD_USER = '" & Session("LOGIN_ID") & "' "
			w_sSQL = w_sSQL & vbCrLf & " WHERE  T26_NENDO = " & m_iSyoriNen
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_SIKEN_KBN = " & m_iSikenKbn
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_SIKEN_CD = '" & m_iSikenCode & "'"
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_GAKUNEN = " & m_sGakunen
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_CLASS = " & w_Cls(i)
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_KAMOKU = '" & m_sKamoku & "'"
'Response.Write "UPD<br>" & w_sSQL & "<br><br>"
'esponse.end
		'新規追加
		Else
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T26_SIKEN_JIKANWARI ("
			w_sSQL = w_sSQL & vbCrLf & " 	T26_NENDO,"				'年度
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKEN_KBN,"         '試験区分
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKEN_CD,"          '試験コード
			w_sSQL = w_sSQL & vbCrLf & " 	T26_GAKUNEN,"           '学年
			w_sSQL = w_sSQL & vbCrLf & " 	T26_CLASS,"             'クラスＮＯ
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KAMOKU,"            '科目コード
			w_sSQL = w_sSQL & vbCrLf & " 	T26_JISSI_KYOKAN,"      '実施教官コード/成績入力教官
			w_sSQL = w_sSQL & vbCrLf & " 	T26_JISSI_FLG,"         '実施フラグ
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKENBI,"           '実施日付
			w_sSQL = w_sSQL & vbCrLf & " 	T26_MAIN_FLG,"          'メイン教官フラグ 
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_INP_FLG,"   '成績入力教官フラグ 
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN1,"   '成績入力教官コード1
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN2,"   '成績入力教官コード2
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN3,"   '成績入力教官コード3
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN4,"   '成績入力教官コード4
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN5,"   '成績入力教官コード5
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KANTOKU_KYOKAN,"    '監督教官コード
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KYOSITU,"           '実施教室コード
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKEN_JIKAN,"       '試験時間
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KAISI_JIKOKU,"      '開始時刻
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SYURYO_JIKOKU,"     '終了時刻
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KYOKAN_RENMEI,"     '教官連名
			w_sSQL = w_sSQL & vbCrLf & " 	T26_INS_DATE,"          '登録日時
			w_sSQL = w_sSQL & vbCrLf & " 	T26_INS_USER,"          '登録者
			w_sSQL = w_sSQL & vbCrLf & " 	T26_UPD_DATE,"          '更新日時
			w_sSQL = w_sSQL & vbCrLf & " 	T26_UPD_USER "          '更新者
			w_sSQL = w_sSQL & vbCrLf & ") VALUES ("
			w_sSQL = w_sSQL & vbCrLf & " " & m_iSyoriNen & ","		'年度
			w_sSQL = w_sSQL & vbCrLf & " " & m_iSikenKbn & ","      '試験区分
			w_sSQL = w_sSQL & vbCrLf & "'" & m_iSikenCode & "',"    '試験コード
			w_sSQL = w_sSQL & vbCrLf & " " & m_sGakunen & ","       '学年
			w_sSQL = w_sSQL & vbCrLf & " " & w_Cls(i) & ","         'クラスＮＯ
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKamoku & "',"       '科目コード
			w_sSQL = w_sSQL & vbCrLf & "'" & m_iKyokanCd & "',"     '実施教官コード/成績入力教官
			w_sSQL = w_sSQL & vbCrLf & " " & m_iJISSI_FLG & ","     '実施フラグ
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '実施日付
			'w_sSQL = w_sSQL & vbCrLf & "'" & m_iMain_FLG & "',"     'メイン教官フラグ 
			'w_sSQL = w_sSQL & vbCrLf & "'" & m_iSeiseki_FLG & "',"  '成績入力教官フラグ 
			w_sSQL = w_sSQL & vbCrLf & "'" & "1" & "',"				'メイン教官フラグ 
			w_sSQL = w_sSQL & vbCrLf & "'" & "1" & "',"				'成績入力教官フラグ 
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan1 & "',"       '成績入力教官コード1
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan2 & "',"       '成績入力教官コード2
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan3 & "',"       '成績入力教官コード3
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan4 & "',"       '成績入力教官コード4
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan5 & "',"       '成績入力教官コード5
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '監督教官コード
			w_sSQL = w_sSQL & vbCrLf & "" & m_sKyositu & ","       '実施教室コード
			If Trim(m_sJikan) = "" Or IsNull(m_sJikan) Then
				m_sJikan = 0
			End If
			w_sSQL = w_sSQL & vbCrLf & " " & m_sJikan & ","         '試験時間
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '開始時刻
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '終了時刻
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '教官連名
			w_sSQL = w_sSQL & vbCrLf & " " & "TO_CHAR(SYSDATE,'YYYY/MM/DD')"  & ","    '登録日時
			w_sSQL = w_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "',"            	   '登録者
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '更新日時
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & " "           '更新者
			w_sSQL = w_sSQL & vbCrLf & ")"
'Response.Write "INS<br>" & w_sSQL & "<br><br>"
		End If		    
'Response.End

		w_iRet = gf_ExecuteSQL(w_sSQL)
		If w_iRet <> 0 Then
		    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		    m_bErrFlg = True
		    Exit Function
		End If
	Next

'response.end
    f_Update = True

End Function

'********************************************************************************
'*  [機能]  試験時間割データの該当レコード件数取得する
'*  [引数]  
'*  [戻値]  f_GetSJikanKensu：レコード件
'*  [説明]  
'********************************************************************************
Function f_GetSJikanKensu(p_sClass)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetSJikanKensu = 0

	Do

		'//クラス名称取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  COUNT(*) as JikanData"
		w_sSql = w_sSql & vbCrLf & " FROM T26_SIKEN_JIKANWARI"
		w_sSql = w_sSql & vbCrLf & " WHERE  T26_NENDO = " & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " and T26_SIKEN_KBN = " & m_iSikenKbn
		w_sSql = w_sSql & vbCrLf & " and T26_SIKEN_CD = '" & m_iSikenCode & "'"
		w_sSql = w_sSql & vbCrLf & " and T26_GAKUNEN = " & m_sGakunen
		w_sSql = w_sSql & vbCrLf & " and T26_CLASS = " & p_sClass
		w_sSql = w_sSql & vbCrLf & " and T26_KAMOKU = '" & m_sKamoku & "'"

'response.write w_sSql&vbCrLf&"<BR>ssssssssssssssss" & p_sClass
'response.end
		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)

		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'試験時間割データの該当レコード件数取得
			f_GetSJikanKensu = cint(rs("JikanData"))
		End If

		Exit Do
	Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  削除処理
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
function f_Delete
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文

    f_Delete = False

    w_sSQL = w_sSQL & vbCrLf & " delete "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & " T26_SIKEN_JIKANWARI "
    '抽出条件の作成
    w_sSQL = w_sSQL & vbCrLf & " WHERE T26_NENDO = " & m_iSyoriNen
    w_sSQL = w_sSQL & vbCrLf & " and T26_SIKEN_KBN = " & m_iSikenKbn
    w_sSQL = w_sSQL & vbCrLf & " and T26_SIKEN_CD = '" & m_iSikenCode & "'"
    w_sSQL = w_sSQL & vbCrLf & " and T26_GAKUNEN = " & m_sGakunen
    w_sSQL = w_sSQL & vbCrLf & " and T26_CLASS = " & m_sClass
    w_sSQL = w_sSQL & vbCrLf & " and T26_KAMOKU = '" & m_sKamoku & "'"
        
'       response.write ("<BR>w_sSQL = " & w_sSQL)

    w_iRet = gf_ExecuteSQL(w_sSQL)

    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    End If

    f_Delete = True

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

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function gonext() {
		window.alert("<%= C_TOUROKU_OK_MSG %>");
		//document.frm.action = "./default.asp";
		document.frm.action = "./skn0130_main.asp";
		document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>

<body bgcolor="#ffffff" onLoad="setTimeout('gonext()',0000)">

<center>

<Form Name="frm" method="post">

<input type="hidden" name="txtMode"     value = "<%=m_sMode%>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
<input type="hidden" name="txtSikenKbn" value="<%= m_iSikenKbn %>">
<input type="hidden" name="txtSikenCode" value="<%= m_iSikenCode %>">

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