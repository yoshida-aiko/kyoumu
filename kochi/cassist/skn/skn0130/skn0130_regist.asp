<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 試験実施科目登録
' ﾌﾟﾛｸﾞﾗﾑID : skn/skn0130/skn0130_regist.asp
' 機      能: 試験実施科目の登録を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 本村　文
' 変      更: 2001/12/07 佐野大悟 成績入力教官かどうか判断してUIを変えるのをやめる。
' 変      更: 2009/11/25 岩田     ｸﾗｽ単位で登録する。
' 変      更: 2019/06/24 藤林     クラスのメイン学科の受講生が一人もいない場合は、クラス名の欄に学科名を表示する
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系

    Public  m_iKyokanCd         ':教官コード
    Public  m_iSyoriNen         ':処理年度
    Public  m_iSikenKbn         ':試験区分
    Public  m_iSikenCode        ':試験ｺｰﾄﾞ
    Public  m_sGakunen          ':学年
    Public  m_sClass            ':ｸﾗｽ
    Public  m_sClassTmp         ':ｸﾗｽ作業用
    Public  m_sKamoku           ':科目ｺｰﾄﾞ
    Public  m_sKyokan_NAME      ':教官名
	Public  m_seisekiF

    Public m_sJikanWhere    '時間の条件
    Public m_sKyosituWhere  '教室コンボの条件


    Public  m_Rs                ':表示ﾃﾞｰﾀ

    Public m_bErrFlg        'ｴﾗｰﾌﾗｸﾞ
    Public m_iNendo         '年度
    Public m_sKyokan_CD     '教官CD
    Public m_iMax
    Public m_iDsp
    Public m_sPageCD
    Public m_sTitle         ''新規登録・修正の表示用
    Public m_sDBMode        ''DBへの更新ﾓｰﾄﾞ
    Public m_sMode          ''画面の表示のﾓｰﾄﾞ

    Public m_sGetTable    ''画面の表示のﾓｰﾄﾞ

	Public m_sSeisekiDate
	Public m_chekdate

	Public  m_sClassName            'クラス名(クラスのメイン学科の受講生が一人もいない場合は、学科名)		'//2019/06/24 Add Fujibayashi

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

    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="試験実施科目登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    m_bErrFlg = False

	m_chekdate = 0

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

        '// 値を変数に入れる
        Call s_SetParam()

        '// 教官名を表示する
        if f_GetData_Kyokan(m_iKyokanCd,m_sKyokan_NAME) = False then
            exit do
        end if

        '// ﾃﾞｰﾀを表示する
        if f_GetData() = False then
            exit do
        end if

        '時間に関するWHREを作成する
        Call f_MakeJikanWhere()
        '実施教室に関するWHREを作成する
        Call f_MakeKyosituWhere()

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
    m_Rs.close
    set m_Rs = nothing
    Call gs_CloseDatabase()

End Sub



'********************************************************************************
'*  [機能]  値を変数に入れる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Sub s_SetParam()
Dim w_clsTmp()
'''Session("SYORI_NENDO") = "999"

    On Error Resume Next
    Err.Clear

    m_iNendo     = Session("NENDO")
    m_sTitle = "修正"
    m_sPageCD    = Request("txtPageCD")
    m_sMode = Request("txtMode")
    m_iKyokanCd = Session("KYOKAN_CD")              ':教官コード
    m_iSyoriNen = Session("NENDO")                  ':処理年度
    m_iSikenKbn = Request("txtSikenKbn")            ':試験区分
    m_iSikenCode = Request("txtSikenCd")            ':試験ｺｰﾄﾞ
    m_sGakunen = Request("txtGakunen")              ':学年
    m_sClass = Request("txtClass")                  ':ｸﾗｽ		'//2009.11.25 skn0130_main.asp で選択されたｸﾗｽが渡される。
    m_sClassTmp = f_FstSplit(Request("txtClass"),"#")  ':ｸﾗｽ作業用
    m_sKamoku = Request("txtKamoku")                ':科目ｺｰﾄﾞ
    m_seisekiF = Request("txtSeisekiFlg")
	m_sGetTable = Request("txtGetTable")			'T26とT27のどちらのテーブルからデータを取ってくるか

	m_sSeisekiDate = Request("txtKikan")			'入力期間終了日

	m_sClassName = Request("txtClassName")			'クラス名(クラスのメイン学科の受講生が一人もいない場合は、学科名)		'//2019/06/24 Add Fujibayashi
    
End Sub

Function f_FstSplit(p_str,p_chr)
'********************************************************************************
'*  [機能]  最初の文字を取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
	dim w_num
	f_FstSplit = p_str

	w_num = InStr(p_str,p_chr)
	If w_num <> 0 then f_FstSplit = left(p_str,w_num-1)

End Function

'********************************************************************************
'*  [機能]  教官の名称を取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
function f_GetData_Kyokan(p_iKyokanCd,p_sKyokan_NAME)
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    dim w_Rs

    On Error Resume Next
    Err.Clear

    f_GetData_Kyokan = False
    p_sKyokan_NAME = ""

    if trim(cstr(p_iKyokanCd)) = "" then
        exit function
    end if

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " M04.M04_NENDO "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKAN_CD "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN M04 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    M04_NENDO = " & m_iNendo & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN_CD = '" & p_iKyokanCd & "'"

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    Else
        'ページ数の取得
        m_iMax = gf_PageCount(w_Rs,m_iDsp)
    End If

    p_sKyokan_NAME = w_Rs("M04_KYOKANMEI_SEI") & "  " & w_Rs("M04_KYOKANMEI_MEI")


    w_Rs.close

    f_GetData_Kyokan = True

end function

'********************************************************************************
'*  [機能]  科目の名称を取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
function f_GetKamoku(p_KamokuCd)
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    dim w_Rs

    On Error Resume Next
    Err.Clear

    f_GetKamoku = ""

    If Trim(cstr(p_KamokuCd)) = "" then
        Exit Function
    End If

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & "  M03_NENDO "
    w_sSQL = w_sSQL & vbCrLf & " ,M03_KAMOKU_CD "
    w_sSQL = w_sSQL & vbCrLf & " ,M03_KAMOKUMEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    M03_KAMOKU "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    M03_NENDO = " & m_iNendo & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M03_KAMOKU_CD = '" & p_KamokuCd & "'"

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    End If

    f_GetKamoku = gf_HTMLTableSTR(w_Rs("M03_KAMOKUMEI"))

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  更新時の表示ﾃﾞｰﾀを取得する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
function f_GetData()
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_Rs

    On Error Resume Next
    Err.Clear

    f_GetData = False

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " T26.T26_SIKENBI "          ''実施日付
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_JISSI_FLG "       ''実施ﾌﾗｸﾞ
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SIKEN_JIKAN "         ''試験時間
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_KYOSITU "         ''実施教室
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_MAIN_FLG "     		''メイン教官フラグ
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_INP_FLG "     ''成績入力教官フラグ
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN1 "         ''成績入力教官１
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN2 "     ''成績入力教官２
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN3 "         ''成績入力教官３
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN4 "     ''成績入力教官４
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN5 "     ''成績入力教官５
    w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI"                ''科目名
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    T26_SIKEN_JIKANWARI T26 "
    w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_NENDO  = M03.M03_NENDO(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_NENDO = " & m_iSyoriNen & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_SIKEN_KBN = " & m_iSikenKbn & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_SIKEN_CD = '" & m_iSikenCode & "' AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_GAKUNEN = " & m_sGakunen & " AND "
    'w_sSQL = w_sSQL & vbCrLf & "    T26.T26_CLASS = " & m_sClassTmp & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_CLASS = " & m_sClass & " AND "		'//2009.11.25 iwata ｸﾗｽ単位で登録する
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_KAMOKU = '" & m_sKamoku & "' AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_JISSI_KYOKAN = '" & m_iKyokanCd & "'"

    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        m_bErrFlg = True
        Exit Function
    Else
        'ページ数の取得
        m_iMax = gf_PageCount(m_Rs,m_iDsp)
    End If

    f_GetData = True

end function


'********************************************************************************
'*  [機能]  時間コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Sub f_MakeJikanWhere()

    m_sJikanWhere=""
    m_sJikanWhere = m_sJikanWhere & " M42_NENDO = " & m_iSyoriNen & ""

'response.write m_sJikanWhere

End Sub

'********************************************************************************
'*  [機能]  実施教室コンボに関するWHREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Sub f_MakekyosituWhere()

    m_sKyosituWhere=""
    m_sKyosituWhere = " M06_NENDO = " & m_iSyoriNen & ""

'response.write m_sKyosituWhere

End Sub

'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'       (リストダウンボックス選択表示用)
'[引数] pData1 : データ１
'       pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'
'****************************************************
Function f_Selected(pData1,pData2)

    On Error Resume Next
    Err.Clear

    f_Selected = ""

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected"
        Else
        End If
    End If

End Function

'****************************************************
'[機能] ﾗｼﾞｵﾎﾞﾀﾝ
'[引数] pData1 : データ１
'[戻値] f_Checked : "SELECTED" OR ""
'
'****************************************************
Function f_Checked(pData,p_Chk1,p_Chk2)

    On Error Resume Next
    Err.Clear

    p_Chk1=""
    p_Chk2=""

    if cstr(pData) = cstr(C_SIKEN_KBN_JISSI) then           ''実施
        p_Chk1=""
        p_Chk2="checked"
    elseif cstr(pData) = cstr(C_SIKEN_KBN_NOT_JISSI) then       ''実施しない
        p_Chk1="checked"
        p_Chk2=""
    else
        p_Chk1=""
        p_Chk2=""
    end if

End Function

'********************************************************************************
'*  [機能]  表示項目(試験)を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Function f_GetSikenName()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetSikenName = ""
    w_sSikenName = ""

    Do
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_NENDO=" & m_iNendo
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_iSikenKbn

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

        If rs.EOF = False Then
            w_sSikenName = rs("M01_SYOBUNRUIMEI")
        End If

        Exit Do
    Loop

	f_GetSikenName = w_sSikenName

    Call gf_closeObject(rs)

End Function

Function f_Nyuryokudate(p_sSikenDate,p_sGakunen)
'********************************************************************************
'*	[機能]	学年別試験期間を取得する。
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	Add 2001.12.26 岡田
'********************************************************************************
	dim w_date

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1


	w_date = gf_YYYY_MM_DD(date(),"/")
	w_Syuryo = "T24_SIKEN_SYURYO"
	w_kyokan = Session("KYOKAN_CD")

	if w_kyokan = NULL or w_kyokan = "" then w_kyokan = "@@@"

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MIN(T24_SIKEN_NITTEI.T24_SIKEN_KAISI) as KAISI"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SIKEN_SYURYO) as SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO) as SEI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_GAKUNEN =" & Cint(p_sGakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_iSikenKbn)
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KAISI <= '" & w_date & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_SYURYO >= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  Group By M01_SYOBUNRUIMEI"

'response.write w_sSQL & "<<<BR>"

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do
		End If

		If m_DRs.EOF Then
			Exit Do
		Else

			p_sSikenDate = m_DRs("SYURYO") '//okada 2001.12.25
'response.write " [ " & p_sSikenDate & " ] "
		End If
		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Dim w_Kyokan
Dim w_Chk1
Dim w_Chk2

    On Error Resume Next
    Err.Clear

	'// ﾌﾞﾗｳｻﾞｰによってｸﾗｽ指定を変える
	if session("browser") = "IE" then
		w_Class = "class='num'"
	Else
		w_Class = ""
	End if

%>

<html>

<head>

<title>使用教科書登録</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
         flag = false;

         function Lock() {
            if(frm.chk1[0].checked){
                fm = document.frm;
                flag = !flag;
                fm.date.disabled = true;
                fm.txtTime.disabled = true;
                fm.room.disabled = true;
                }
            }
         function unLock() {
            if(frm.chk1[1].checked){
                fm = document.frm;
                fm.date.disabled = false;
                fm.txtTime.disabled = false;
                fm.room.disabled = false;
                }
            }
        //************************************************************
        //  [機能]  メインページへ戻る
        //  [引数]  なし
        //  [戻値]  なし
        //  [説明]
        //
        //************************************************************
        function f_Back(){

            //document.frm.action="./default.asp";
            document.frm.action="./skn0130_main.asp";
            document.frm.target="";
            document.frm.submit();

        }

        //************************************************************
        //  [機能]  卒検教官参照選択画面ウィンドウオープン
        //  [引数]
        //  [戻値]
        //  [説明]
        //************************************************************
        function KyokanWin(p_iInt,p_sKNm) {

			var obj=eval("document.frm."+p_sKNm)

            URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
            //URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+p_sKNm+"";
            nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=650,top=0,left=0");
            nWin.focus();
            return true;
        }

	//************************************************************
	//  [機能] クリアボタンを押されたとき
	//  [引数]
	//  [戻値]
	//  [説明]
	//************************************************************
	function jf_Clear(pTextName,pHiddenName){
		eval("document.frm."+pTextName).value = "";
		eval("document.frm."+pHiddenName).value = "";
	}

	//************************************************************
	//  [機能] 実施する、しないボタンを押されたとき
	//  [引数]
	//  [戻値]
	//  [説明]
	//************************************************************
	function jf_Action(pMode){

		if(pMode == "False"){


			// 時間をNUllにする
			document.frm.txtJikan.value = "";

			// 教室をNULLにする
			document.frm.txtKyositu.options[0].selected = true;

			//時間をreadOnlyにする
			document.frm.txtJikan.readOnly = true;

			//教室をdisabledにする
			document.frm.txtKyositu.disabled = true;

		}else{
			//時間のreadOnlyをはずす
			document.frm.txtJikan.readOnly = false;

			//教室をdisabledにする
			document.frm.txtKyositu.disabled = false;

		}

	}

	//************************************************************
	//  [機能]入力準備期間外
	//  [引数]
	//  [戻値]
	//  [説明]
	//************************************************************
	function jf_Action2(){

			//時間をreadOnlyにする
			document.frm.txtJikan.readOnly = true;

			//実施するしない
			//document.frm.chk1.disabled = false;

			//教室をdisabledにする
			document.frm.txtKyositu.disabled = true;

	}


    //************************************************************
    //  [機能]  使用教科書登録
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_touroku(){

		// 実施するとき
		if(document.frm.chk1[1].checked == true && document.frm.chk1[1].disabled == false){

	        // ■■■NULLﾁｪｯｸ■■■
	        if( f_Trim(document.frm.txtJikan.value) == "" ){
	            window.alert("試験時間が入力されていません");
	            document.frm.txtJikan.focus();
	            return;
	        }
	        // ■■■半角英数ﾁｪｯｸ■■■
	        //var str = new String(document.frm.txtJikan.value);
	        var str = document.frm.txtJikan.value;
	        if( isNaN(str) ){
	            window.alert("試験時間が半角英数字ではありません");
	            document.frm.txtJikan.focus();
	            return ;
	        }

	        if( f_Trim(document.frm.txtJikan.value) < 0 ){
	            window.alert("試験時間が正しく入力されていません");
	            document.frm.txtJikan.focus();
	            return;
	        }

            if (f_chkNumber(f_Trim(document.frm.txtJikan.value))==1){
                alert("試験時間が正しく入力されていません")
	            document.frm.txtJikan.focus();
                return;
			}

	        // ■■■5分チェック■■■
	        var str = new String(document.frm.txtJikan.value);
			if( f_Trim(str) == 0 ){
	                window.alert("試験時間は5分単位で入力してください");
	                document.frm.txtJikan.focus();
	                return ;
			}else{
		        if( str.length < 2 ){
		            str = 0 + str;
		        }

		        if( f_Trim(str).substr(Number(str.length)-1,1) != 0 ){
		            if( f_Trim(str).substr(Number(str.length)-1,1) != 5 ){
		                window.alert("試験時間は5分単位で入力してください");
		                document.frm.txtJikan.focus();
		                return ;
		            }
				}
			}
		}
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		document.frm.chk1[0].disabled = false;
		document.frm.chk1[1].disabled = false;
        document.frm.action="./skn0130_db.asp";
        document.frm.target="";
        document.frm.txtDBMode.value = "Update";
        document.frm.submit();
    }

    //************************************************************
    //  [機能]  数字チェック
    //  [引数]  p_num
    //  [戻値]  成功：0   失敗：1
    //  [説明]	数字かどうかをチェック(マイナス値、小数点有の場合はエラーを返す)
    //************************************************************
	function f_chkNumber(p_num){

		//数値チェック
		if (isNaN(p_num)){
			return 1;
		}else{

			//マイナスをチェック
			var wStr = new String(p_num)
			if (wStr.match("-")!=null){
				return 1;
			};

			//小数点チェック
			w_decimal = new Array();
			w_decimal = wStr.split(".")
			if(w_decimal.length>1){
				return 1;
			}

		};
		return 0;
	}


    //************************************************************
    //  [機能]  使用教科書削除
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Delete(){

        document.frm.action="./skn0130_db.asp";
        document.frm.target="";
        document.frm.txtDBMode.value = "Delete";
        document.frm.submit();
    }

    //-->
    </script>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <style>
    .gray {background:gray;}
    .white {background:white;}
    </style>
</head>

<%
	'// 実施ﾌﾗｸﾞ
	if gf_HTMLTableSTR(m_Rs("T26_JISSI_FLG")) = "2" then
		w_onLadFcnc = "onLoad=jf_Action('False')"
	End if

	'// 入力準備期間外

	w_date = gf_YYYY_MM_DD(date(),"/")
'response.write m_sSeisekiDate & " " &  w_date  & " " & m_sGakunen
	'学年の試験入力準備期間を取得（m_sGakunen）
	Call f_Nyuryokudate(m_sSeisekiDate,m_sGakunen)

	if m_sSeisekiDate < w_date Then
'response.write m_sSeisekiDate & " " &  w_date
		w_onLadFcnc = "onLoad=jf_Action2()"
		m_chekdate = 1
	End if
%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%=w_onLadFcnc%>>
<%call gs_title("試験実施科目登録",m_sTitle)%>
<form name="frm" method="post">
<center>

<span class=CAUTION>
	※ 「実施する」の場合は、「時間」が必須入力となります<BR>
	※ 「実施しない」にすると、「時間」と「実施教室」の選択ができません
</span>
<BR>
<br>

<table border="0" cellpadding="1" cellspacing="1" width="400">
    <tr>
        <td align="left">
            <table width="100%" border=1 CLASS="hyo">
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">試験</th>
	                <TD CLASS="detail">　<%=f_GetSikenName%></td>
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">クラス</th>
<!--	                <TD CLASS="detail">　<%=m_sGakunen & "年　" & gf_GetClassName(m_iNendo,m_sGakunen,m_sClass) %></td>				'2019/06/24 Del Fujibayashi-->
						<TD CLASS="detail">　<%=m_sGakunen & "年　" & m_sClassName%></td>											<!--'2019/06/24 Add Fujibayashi-->
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">科目名称</th>
	                <TD CLASS="detail">　<%=f_GetKamoku(m_sKamoku)%>　</td>
	            </tr>
<!--
				<% if Not gf_IsNull(m_Rs("T26_SIKENBI")) then %>
		            <!tr>
						<TH nowrap CLASS="header" align="center" height="16">実施日付</th>
		                <TD CLASS="detail">　<%= gf_fmtWareki(m_Rs("T26_SIKENBI")) %></td>
		            </tr>
				<% End if %>
-->
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">実　　施</th>
	                <% call f_Checked(gf_HTMLTableSTR(m_Rs("T26_JISSI_FLG")),w_Chk1,w_Chk2) %>
					<%
					if gf_SetNull2Zero(trim(m_Rs("T26_JISSI_FLG"))) = 0 then
						wClass = "JISSHIMI"
						w_Chk2 = "checked"
					Else
						wClass = "detail"
					End if
					%>
	                <TD CLASS="<%=wClass%>">
					<% if m_chekdate = 1 then %>
						<input type="radio" disabled name="chk1" value="2" <%=w_Chk1%> onClick="javascript:jf_Action('False')"><font>実施しない
						<input type="radio" disabled name="chk1" value="1" <%=w_Chk2%> onClick="javascript:jf_Action('True')">実施する</font></td>
					<% else %>
						<input type="radio" name="chk1" value="2" <%=w_Chk1%> onClick="javascript:jf_Action('False')"><font>実施しない
						<input type="radio" name="chk1" value="1" <%=w_Chk2%> onClick="javascript:jf_Action('True')">実施する</font></td>
					<% end if %>
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">時　　間</th>
	                <TD CLASS="detail"><input type="text" name="txtJikan" size="4" value="<%=m_RS("T26_SIKEN_JIKAN")%>" <%=w_Class%>>&nbsp;分
					</td>
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">実施教室</th>
	                <TD CLASS="detail">
	    <%          '共通関数から実施教室に関するコンボボックスを出力する
	                call gf_ComboSet("txtKyositu",C_CBO_M06_KYOSITU,m_sKyosituWhere,"",True,gf_HTMLTableSTR(m_Rs("T26_KYOSITU")))
	    %>
	                </td>
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16" valign=top ROWSPAN="5">成績入力教官</th>
	                <TD CLASS="detail">&nbsp;1：<%=m_sKyokan_NAME%><br>
	                <input type=hidden name="SKyokanCd1" VALUE='<%=m_iKyokanCd%>'>
	                </td>
	            </tr>
	            <tr>
	                <TD CLASS="detail" nowrap>
	                <% call f_GetData_Kyokan(gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN2")),w_Kyokan)%>
	                &nbsp;2：<input type=text class=text name="SKyokanNm2" VALUE='<%=w_Kyokan%>' readonly>
	                <input type=hidden name="SKyokanCd2" VALUE='<%=gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN2"))%>'>
	                <!--<input type=button class=button value="選択" onclick="KyokanWin(2,'<%=w_Kyokan%>')">-->

	                <input type=button class=button value="選択" onclick="KyokanWin(2,'SKyokanNm2')">
					<input type=button class=button value="クリア" onclick="jf_Clear('SKyokanNm2','SKyokanCd2')">
	                </td>
	            </tr>
	            <tr>
	                <TD CLASS="detail" nowrap>
	                <% call f_GetData_Kyokan(gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN3")),w_Kyokan)%>
	                &nbsp;3：<input type=text class=text name="SKyokanNm3" VALUE='<%=w_Kyokan%>' readonly>
	                <input type=hidden name="SKyokanCd3" VALUE='<%=gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN3"))%>'>
	                <!--<input type=button class=button value="選択" onclick="KyokanWin(3,'<%=w_Kyokan%>')">-->
	                <input type=button class=button value="選択" onclick="KyokanWin(3,'SKyokanNm3')">
					<input type=button class=button value="クリア" onclick="jf_Clear('SKyokanNm3','SKyokanCd3')">
	                </td>
	            </tr>
	            <tr>
	                <TD CLASS="detail" nowrap>
	                <% call f_GetData_Kyokan(gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN4")),w_Kyokan)%>
	                &nbsp;4：<input type=text class=text name="SKyokanNm4" VALUE='<%=w_Kyokan%>' readonly>
	                <input type=hidden name="SKyokanCd4" VALUE='<%=gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN4"))%>'>
	                <!--<input type=button class=button value="選択" onclick="KyokanWin(4,'<%=w_Kyokan%>')">-->
	                <input type=button class=button value="選択" onclick="KyokanWin(4,'SKyokanNm4')">
					<input type=button class=button value="クリア" onclick="jf_Clear('SKyokanNm4','SKyokanCd4')">
	                </td>
	            </tr>
	            <tr>
	                <TD CLASS="detail" nowrap>
	                <% call f_GetData_Kyokan(gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN5")),w_Kyokan)%>
	                &nbsp;5：<input type=text class=text name="SKyokanNm5" VALUE='<%=w_Kyokan%>' readonly>
	                <input type=hidden name="SKyokanCd5" VALUE='<%=gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN5"))%>'>
	                <!--<input type=button class=button value="選択" onclick="KyokanWin(5,'<%=w_Kyokan%>')">-->
	                <input type=button class=button value="選択" onclick="KyokanWin(5,'SKyokanNm5')">
					<input type=button class=button value="クリア" onclick="jf_Clear('SKyokanNm5','SKyokanCd5')">
	                </td>
	            </tr>
            </TABLE>
        </td>
    </TR>
</TABLE>
<table width=40%>
    <tr>
        <td width="50%" align="left">
            <input type="button" class=button value="　登　録　" OnClick="f_touroku()">
        </td>
        <td width="50%" align="right">
            <input type="Button" class=button value="キャンセル" OnClick="f_Back()">
        </td>
    </tr>
</table>

    <input type="hidden" name="txtDBMode" value = "<%=m_sGetTable%>">
    <input type="hidden" name="txtMode"   value = "<%=m_sMode%>">
    <input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
    <input type="hidden" name="txtTitle"  value="<%= m_sTitle %>">

    <input type="hidden" name="txtSikenKbn"  value="<%= m_iSikenKbn %>">
    <input type="hidden" name="txtSikenCode" value="<%= m_iSikenCode %>">
    <input type="hidden" name="txtGakunen"   value="<%= m_sGakunen %>">
    <input type="hidden" name="txtClass"     value="<%= m_sClass %>">
    <input type="hidden" name="txtKamoku"    value="<%= m_sKamoku %>">
    <input type="hidden" name="txtMainF"    value="<%= m_Rs("T26_MAIN_FLG") %>">
    <input type="hidden" name="txtSeisekiF"    value="<%= m_Rs("T26_SEISEKI_INP_FLG") %>">
    <input type="hidden" name="txtJissiFLG"  value=<%=m_chekdate%>>
    <input type="hidden" name="txtClassName"     value="<%= m_sClassName %>">		<% '2019/06/24 Add Fujibayashi%>

</center>

</form>
</body>

</html>

<%
End Sub
%>

