<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0150_bottom.asp
' 機      能: 下ページ 成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード   ＞    SESSIONより（保留）
'           :年度     ＞    SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード   ＞    SESSIONより（保留）
'           :年度     ＞    SESSIONより（保留）
' 説      明:
' (パターン)
' ・通常授業、特別活動
' ・数値入力、文字入力(成績)
' ・評価不能処理(熊本電波のみ)
' ・科目区分(0:一般科目,1:専門科目)
' ・必修選択区分(1:必修,2:選択)
' ・レベル別区分(0:一般科目,1:レベル別科目)を調べる
'-------------------------------------------------------------------------
' 作      成: 2002/06/21 shin
' 変      更: 2003/04/11 hirota 認定状況を学年別にみるように変更
' 変      更: 2003/05/13 hirota 久留米高専用　成績入力時は受講時間を必須入力とする
' 変      更: 2005/12/16 西村　　福島高専の場合最大授業時間数を取得するf_GetJyugyoJIkan()を追加
' 変      更: 2018/10/16 清本 混合学級対応
' 変      更: 2019/06/17 藤林 80001：工学基礎、80002：情報リテラシーの特別処理をやめる
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
  'エラー系
    Dim m_bErrFlg       '//ｴﾗｰﾌﾗｸﾞ

    Const C_ERR_GETDATA = "データの取得に失敗しました"

    '氏名選択用のWhere条件
    Dim m_iNendo        '//年度
    Dim m_sKyokanCd       '//教官コード
    Dim m_sSikenKBN       '//試験区分
    Dim m_sGakuNo       '//学年
    Dim m_sClassNo        '//学科
    Dim m_sKamokuCd       '//科目コード
    Dim m_sSikenNm        '//試験名
    Dim m_rCnt          '//レコードカウント
    Dim m_sGakkaCd
    Dim m_iSyubetu        '//出欠値集計方法

    Dim m_iNKaishi
    Dim m_iNSyuryo
    Dim m_iKekkaKaishi
    Dim m_iKekkaSyuryo

    Dim m_iIdouEnd        '//異動対象期間終了日

    Dim m_iKamoku_Kbn
    Dim m_iHissen_Kbn
  Dim m_ilevelFlg
  Dim m_Rs
  Dim m_SRs

    Dim m_iKongou

  Dim m_iSouJyugyou     '//総授業時間
  DIm m_iJunJyugyou     '//純授業時間

  Dim m_iSouJyugyou1     '//総授業時間	INS 2005/06/13 西村　福島高専用
  DIm m_iJunJyugyou1     '//純授業時間
  Dim m_iSouJyugyou2     '//総授業時間
  DIm m_iJunJyugyou2     '//純授業時間
  Dim m_iSouJyugyou3     '//総授業時間
  DIm m_iJunJyugyou3     '//純授業時間
  Dim m_iSouJyugyou4     '//総授業時間
  DIm m_iJunJyugyou4     '//純授業時間


  Dim m_bSeiInpFlg      '//入力期間フラグ
  Dim m_bKekkaNyuryokuFlg   '//欠課入力可能ﾌﾗｸﾞ(True:入力可 / False:入力不可)

  Dim m_iShikenInsertType

  Dim m_sSyubetu

  '2002/06/21
  Dim m_iKamokuKbn        '//科目区分(0:通常授業、1:特別科目)
  Dim m_sKamokuBunrui       '//科目分類(01:通常授業、02:認定科目、03:特別科目)

  Dim m_iSeisekiInpType
  Dim m_Date
  Dim m_bZenkiOnly
  Dim m_SchoolFlg,m_KekkaGaiDispFlg,m_HyokaDispFlg

  Dim m_MiHyokaFlg

  Dim m_bNiteiFlg

  Dim m_sGakkoNO       '学校番号

  Dim m_lHaitoTani

'///////////////////////////メイン処理/////////////////////////////
  'ﾒｲﾝﾙｰﾁﾝ実行
  Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Sub Main()
  Dim w_iRet
  Dim w_sSQL
  Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

  'Message用の変数の初期化
  w_sWinTitle = "キャンパスアシスト"
  w_sMsgTitle = "成績登録"
  w_sMsg = ""
  w_sRetURL = C_RetURL & C_ERR_RETURL
  w_sTarget = ""

  On Error Resume Next
  Err.Clear

  m_bErrFlg = false
  m_MiHyokaFlg = false

  Do
    '//ﾃﾞｰﾀﾍﾞｰｽ接続
    If gf_OpenDatabase() <> 0 Then
      m_bErrFlg = True
      Exit Do
    End If

    '//不正アクセスチェック
    Call gf_userChk(session("PRJ_No"))

    '//ﾊﾟﾗﾒｰﾀSET
    Call s_SetParam()


    '2002.12.25 Ins
    '学校番号の取得
    if Not gf_GetGakkoNO(m_sGakkoNO) then Exit Do

'Response.Write "[1]"

    '//評価不能を表示するかチェック
    if not gf_ChkDisp(C_DATAKBN_DISP,m_SchoolFlg) then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[2]"

    '//評価不能チェックの処理が必要なら
    '//未評価フラグを調べる
    if m_SchoolFlg then
      if not f_GetMihyoka(m_MiHyokaFlg) then
        m_bErrFlg = True
        Exit Do
      end if
    end if

'Response.Write "[3]"

    '//欠課外を表示するかチェック
    if not gf_ChkDisp(C_KEKKAGAI_DISP,m_KekkaGaiDispFlg) then
      m_bErrFlg = True
      Exit Do
    End If


    '//評価予定を表示するかチェック
    if not gf_ChkDisp(C_HYOKAYOTEI_DISP,m_HyokaDispFlg) then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[5]"

    '//成績入力方法の取得(0:点数[C_SEISEKI_INP_TYPE_NUM]、1:文字[C_SEISEKI_INP_TYPE_STRING]、2:欠課、遅刻[C_SEISEKI_INP_TYPE_KEKKA])
    if not gf_GetKamokuSeisekiInp(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iSeisekiInpType) then
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[6]"

    '//前期のみ開設か通年か調べる
    if not f_SikenInfo(m_bZenkiOnly) then
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[7]"

    '//成績、欠課入力期間チェック
    If not f_Nyuryokudate() Then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[8]"

    '//出欠欠課の取り方を取得
    '//科目区分(0:試験毎,1:累積)
    If gf_GetKanriInfo(m_iNendo,m_iSyubetu) <> 0 Then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[9]"

    '//認定前後情報取得
'   if not gf_GetNintei(m_iNendo,m_bNiteiFlg) then
    if not gf_GetGakunenNintei(m_iNendo,cint(m_sGakuNo),m_bNiteiFlg) then '2003.04.11 hirota
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[10]"

    If m_iKamokuKbn = C_JIK_JUGYO then  '通常授業の場合
      '//科目情報を取得
      '//科目区分(0:一般科目,1:専門科目)、及び、必修選択区分(1:必修,2:選択)を調べる
      '//レベル別区分(0:一般科目,1:レベル別科目)を調べる
      If not f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) Then m_bErrFlg = True : Exit Do
    end if

'Response.Write "[11]"
    If not f_GetKongoClass(cint(m_sGakuNo),cint(m_sClassNo),m_iKongou) Then m_bErrFlg = True : Exit Do

    '//成績、学生データ取得
    If not f_GetStudent() Then m_bErrFlg = True : Exit Do

    If m_Rs.EOF Then
      Call gs_showWhitePage("個人履修データが存在しません。","成績登録")
      Exit Do
    End If

	IF m_sGakkoNO = cstr(C_NCT_FUKUSHIMA) THEN
		'//福島高専の場合,最大の授業時間数を取得 INS 2005/12/16西村
		IF NOT f_GetJyugyoJIkan()Then m_bErrFlg = True : Exit Do
	    If m_Rs.EOF Then
	      Call gs_showWhitePage("個人履修データが存在しません。","成績登録")
	      Exit Do
	    End If

	END IF

'Response.Write "[12]"
'response.end

    '//欠課数の取得
    if not gf_GetSyukketuData2(m_SRs,m_sSikenKBN,m_sGakuNo,m_sClassNo,m_sKamokuCd,m_iNendo,m_iShikenInsertType,m_sSyubetu) then
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[13]"
'response.end
    '// ページを表示
    Call showPage()
    Exit Do
  Loop

  '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
  If m_bErrFlg = True Then
    w_sMsg = gf_GetErrMsg()

    if w_sMsg = "" then w_sMsg = C_ERR_GETDATA

'    Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
  End If

  '// 終了処理
  Call gf_closeObject(m_Rs)
  Call gf_closeObject(m_SRs)

  Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()

  m_iNendo   = request("txtNendo")
  m_sKyokanCd  = request("txtKyokanCd")

  m_sSikenKBN  = cint(request("sltShikenKbn"))
  m_sGakuNo  = cint(request("txtGakuNo"))
  m_sClassNo   = cint(request("txtClassNo"))
  m_sKamokuCd  = request("txtKamokuCd")
  m_sGakkaCd   = request("txtGakkaCd")
  m_sSyubetu   = trim(Request("SYUBETU"))
  m_iShikenInsertType = 0

  m_iKamokuKbn = cint(Request("hidKamokuKbn"))

  if m_iKamokuKbn = C_JIK_JUGYO then
    '通常科目
    m_sKamokuBunrui = C_KAMOKUBUNRUI_TUJYO
  else
    '特別科目
    m_sKamokuBunrui = C_KAMOKUBUNRUI_TOKUBETU
  end if

  m_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")

End Sub

'********************************************************************************
'*  [機能]  前期開設かどうか調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Function f_SikenInfo(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs

    On Error Resume Next
    Err.Clear

    f_SikenInfo = false
  p_bZenkiOnly = false

  '//試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   T15_KAMOKU_CD, "
  w_sSQL = w_sSQL & "   T15_HAITO" & m_sGakuNo
  w_sSQL = w_sSQL & "   ,T15_KAISETU" & m_sGakuNo
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T15_RISYU "
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   T15_NYUNENDO = " & Cint(m_iNendo)-cint(m_sGakuNo)+1
  w_sSQL = w_sSQL & " AND T15_GAKKA_CD = '" & m_sGakkaCd & "'"
  w_sSQL = w_sSQL & " AND T15_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'"
  'w_sSQL = w_sSQL & " AND T15_KAISETU" & m_sGakuNo & "=" & C_KAI_ZENKI

  if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

  'Response.Write w_ssql & "<BR>"
  'response.end

  '//戻り値ｾｯﾄ
  If w_Rs.EOF = False Then

	if cint(w_Rs("T15_KAISETU" & m_sGakuNo)) = C_KAI_ZENKI then
  'Response.Write "C_KAI_ZENKI = " & C_KAI_ZENKI & "<BR>"
  'response.end
	    p_bZenkiOnly = True
	end if

	'配当単位の取得
	m_lHaitoTani = w_Rs("T15_HAITO" & m_sGakuNo)

  else
	Call f_SikenInfo_T16(p_bZenkiOnly)
  End If

  f_SikenInfo = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  前期開設かどうか調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Function f_SikenInfo_T16(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs

    On Error Resume Next
    Err.Clear

    f_SikenInfo_T16 = false
	p_bZenkiOnly = false

  '//試験区分が前期期末の時は、その科目が前期のみか通年かを調べる
  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   T16_HAITOTANI " 
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T16_RISYU_KOJIN "
  w_sSQL = w_sSQL & " 	,T13_GAKU_NEN "		'2018.10.16 Add Kiyomoto 混合学級対応
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   T16_NENDO = " & Cint(m_iNendo)
'  w_sSQL = w_sSQL & " AND T16_GAKKA_CD = '" & m_sGakkaCd & "'"		'2018.10.16 Del Kiyomoto 混合学級対応
  w_sSQL = w_sSQL & " AND T16_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'"
  'w_sSQL = w_sSQL & " AND T16_KAISETU = " & C_KAI_ZENKI
  '2018.10.16 Add Kiyomoto 混合学級対応 -->
  w_sSQL = w_sSQL & " AND T13_NENDO = T16_NENDO "
  w_sSQL = w_sSQL & " AND T13_GAKUNEN = T16_HAITOGAKUNEN "
  w_sSQL = w_sSQL & " AND T13_GAKUSEI_NO = T16_GAKUSEI_NO "
  w_sSQL = w_sSQL & " AND T13_CLASS = " & Cint(m_sClassNo)
  '2018.10.16 Add Kiyomoto 混合学級対応 <--
  w_sSQL = w_sSQL & " GROUP BY T16_HAITOTANI,T16_NENDO,T16_GAKKA_CD,T16_KAMOKU_CD,T16_KAISETU "

  if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

  'Response.Write w_ssql
  'response.end

  '//戻り値ｾｯﾄ
  If w_Rs.EOF = False Then

	if w_Rs("T16_KAISETU") = C_KAI_ZENKI then
	    p_bZenkiOnly = True
	end if

	'配当単位の取得
	m_lHaitoTani = w_Rs("T16_HAITOTANI")

  End If

  f_SikenInfo_T16 = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  未評価フラグがたっているか調べる
'********************************************************************************
function f_GetMihyoka(p_MiHyokaFlg)

  Dim w_sSQL,w_Rs
  Dim w_Table,w_FieldName,w_FromTable,w_KamokuCd

  On Error Resume Next
  Err.Clear

  f_GetMihyoka = false
  p_MiHyokaFlg = false

  if m_iKamokuKbn = C_JIK_JUGYO then
    w_Table = "T16"
    w_FromTable = "T16_RISYU_KOJIN"
    w_KamokuCd = "T16_KAMOKU_CD"
  else
    w_Table = "T34"
    w_FromTable = "T34_RISYU_TOKU"
    w_KamokuCd = "T34_TOKUKATU_CD"
  end if

  select case m_sSikenKBN
    case C_SIKEN_ZEN_TYU : w_FieldName = w_Table & "_DATAKBN_TYUKAN_Z"
    case C_SIKEN_ZEN_KIM : w_FieldName = w_Table & "_DATAKBN_KIMATU_Z"
    case C_SIKEN_KOU_TYU : w_FieldName = w_Table & "_DATAKBN_TYUKAN_K"
    case C_SIKEN_KOU_KIM : w_FieldName = w_Table & "_DATAKBN_KIMATU_K"
  end select

  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   " & w_FieldName & " as MIHYOKA "
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL &       w_FromTable
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   " & w_Table & "_NENDO = " & Cint(m_iNendo) & " and "
  w_sSQL = w_sSQL & "   " & w_KamokuCd & " = '" & m_sKamokuCd & "' and "
  w_sSQL = w_sSQL & "   " & w_Table & "_HAITOGAKUNEN = " & Cint(m_sGakuNo) & " and "
  w_sSQL = w_sSQL & "   " & w_Table & "_GAKKA_CD     = '" & m_sGakkaCd & "' and "
  w_sSQL = w_sSQL &     w_FieldName & "= 4 "


  If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function

  'Response.Write " 1"

  if not w_Rs.EOF then p_MiHyokaFlg = true

  f_GetMihyoka = true

end function


'********************************************************************************
'*  [機能]  コンボで選択された科目の科目区分及び、必修選択区分を調べる
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Function f_GetKamokuInfo(p_iKamoku_Kbn,p_iHissen_Kbn,p_ilevelFlg)
  Dim w_sSQL
  Dim w_Rs

  On Error Resume Next
  Err.Clear

  f_GetKamokuInfo = false

  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   T15_RISYU.T15_KAMOKU_KBN"
  w_sSQL = w_sSQL & "   ,T15_RISYU.T15_HISSEN_KBN"
  w_sSQL = w_sSQL & "   ,T15_RISYU.T15_LEVEL_FLG"
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T15_RISYU"
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(m_sGakuNo) + 1
  w_sSQL = w_sSQL & " AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
  w_sSQL = w_sSQL & " AND T15_RISYU.T15_KAMOKU_CD='" & m_sKamokuCd & "' "

'response.write w_sSQL
'response.end

  If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

  '//戻り値ｾｯﾄ
  If w_Rs.EOF = False Then
    p_iKamoku_Kbn = w_Rs("T15_KAMOKU_KBN")
    p_iHissen_Kbn = w_Rs("T15_HISSEN_KBN")
    p_ilevelFlg = w_Rs("T15_LEVEL_FLG")
  End If

  f_GetKamokuInfo = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  混合クラスかどうかを調べる
'*  [説明]
'********************************************************************************
Function f_GetKongoClass(p_iGakunen,p_iClass,p_iKongo)
  Dim w_sSQL
  Dim w_Rs

  On Error Resume Next
  Err.Clear

  f_GetKongoClass = false

    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M05_SYUBETU "
    w_sSQL = w_sSQL & "FROM M05_CLASS "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_GAKUNEN = " & p_iGakunen & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_CLASSNO = " & p_iClass & " "

  If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

  'Response.Write "2"

  '//戻り値ｾｯﾄ
  If w_Rs.EOF = False Then
    p_iKongo = Cint(w_Rs("M05_SYUBETU"))
  End If

  f_GetKongoClass = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  データの取得
'********************************************************************************
Function f_GetStudent()

  Dim w_sSQL
  Dim w_FieldName
  Dim w_Table
  Dim w_TableName
  Dim w_KamokuName

  On Error Resume Next
  Err.Clear

  f_GetStudent = false

  if m_iKamokuKbn = C_JIK_JUGYO then  '通常授業の場合
    w_Table = "T16"
    w_TableName = "T16_RISYU_KOJIN"
    w_KamokuName = "T16_KAMOKU_CD"
  else
    w_Table = "T34"
    w_TableName = "T34_RISYU_TOKU"
    w_KamokuName = "T34_TOKUKATU_CD"
  end if

  '//文字、数値入力により、取ってくるフィールドを変える
  if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then

    if m_bNiteiFlg and m_iKamokuKbn = C_JIK_JUGYO then
      w_FieldName = "HTEN"
    else
      w_FieldName = "SEI"
    end if

  else
    w_FieldName = "HYOKA"
  end if

  '//検索結果の値より一覧を表示
  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI1, "
  w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI2, "
  w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI3, "
  w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI4, "

  Select Case m_sSikenKBN
    Case C_SIKEN_ZEN_TYU

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI,"
      w_sSQL = w_sSQL & w_Table & "_DATAKBN_TYUKAN_Z as DataKbn ,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z AS KEKA,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_Z AS CHIKAI,"
      w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z as SOUJI,"
      w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z as JYUNJI, "

      if m_iKamokuKbn = C_JIK_JUGYO then
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI, "
      end if

    Case C_SIKEN_ZEN_KIM

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI,"
      w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_Z as DataKbn,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS KEKA_NASI,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_Z AS CHIKAI,"
      w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z as SOUJI, "
      w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z as JYUNJI, "

      if m_iKamokuKbn = C_JIK_JUGYO then
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
      end if

    Case C_SIKEN_KOU_TYU

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K AS KEKA,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS KEKA_NASI,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_K AS CHIKAI,"
      w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K as SOUJI, "
      w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K as JYUNJI, "
      w_sSQL = w_sSQL & w_Table & "_DATAKBN_TYUKAN_K as DataKbn,"

      if m_iKamokuKbn = C_JIK_JUGYO then
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
      end if

      '2002/12/25
      '八代高専用にselectフィールドの追加
      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA_ZK,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_Z AS CHIKAI_ZK,"

      '2003/01/06 UPD テーブル切り分け
      w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_Z AS KOUSINBI_ZK, "   '前期末
      w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_K AS KOUSINBI_TK, "   '後期中間

    Case C_SIKEN_KOU_KIM

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI_ZT,"
      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI_KT,"
      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI_KK,"

      w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z AS KEKA_ZT,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA_ZK,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K AS KEKA_KT,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_K AS KEKA,"

      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_K AS KEKA_NASI,"

      w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_Z AS CHIKAI_ZT,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_Z AS CHIKAI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_K AS CHIKAI_KT,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_K AS CHIKAI,"

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI,"

      w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K as SOUJI, "
      w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K as JYUNJI, "
		
      w_sSQL = w_sSQL & w_Table & "_SAITEI_JIKAN, "
      w_sSQL = w_sSQL & w_Table & "_KYUSAITEI_JIKAN, "

      w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_K as DataKbn,"
      w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_Z as DataKbn_ZK,"

      if m_iKamokuKbn = C_JIK_JUGYO then
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI_ZT, "
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI_ZK, "
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI_KT, "
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "

        w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_Z AS KOUSINBI_ZK, "
        w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_K AS KOUSINBI_KK, "

        '2002/12/25
        '八代高専用にselectフィールドの追加
        w_sSQL = w_sSQL & " T16_KOUSINBI_TYUKAN_K AS KOUSINBI_TK, "   '後期中間
      end if

  End Select

 '福島高専用　INS 2005/06/13 西村
'前期　中間
  w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z as SOUJI1, "
  w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z as JYUNJI1, "
'前期　期末
  w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z as SOUJI2, "
  w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z as JYUNJI2, "
'後期　中間
  w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K as SOUJI3, "
  w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K as JYUNJI3, "
'後期　期末
  w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K as SOUJI4, "
  w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K as JYUNJI4, "

  w_sSQL = w_sSQL & " T13_GAKUSEI_NO AS GAKUSEI_NO,"
  w_sSQL = w_sSQL & " T13_GAKUSEKI_NO AS GAKUSEKI_NO,"
  w_sSQL = w_sSQL & " T11_SIMEI AS SIMEI, "

  w_sSQL = w_sSQL & " T13_SYUSEKI_NO1 AS SYUSEKI_NO1, "
  w_sSQL = w_sSQL & " T13_SYUSEKI_NO2 AS SYUSEKI_NO2, "

  if m_iKamokuKbn = C_JIK_JUGYO then
    w_sSQL = w_sSQL & "   T16_SELECT_FLG, "
    w_sSQL = w_sSQL & "   T16_LEVEL_KYOUKAN, "
    w_sSQL = w_sSQL & "   T16_OKIKAE_FLG, "
    If m_sGakkoNO = cstr(C_NCT_KURUME) OR m_sGakkoNO = cstr(C_NCT_NUMAZU) then
	    w_sSQL = w_sSQL & "   T16_MENJYO_FLG, "
	End If
  Else
	    w_sSQL = w_sSQL & "   0 AS T16_MENJYO_FLG, "
  end if

	'2003.8.25 免除フラグ追加 免除フラグ＝1の場合、後期末に成績をコピーしない為。ITO （久留米対応）
	'2004.2.18 沼津追加　免除フラグ追加 免除フラグ＝1の場合、後期末に成績をコピーしない為。（沼津対応）
    'If m_sGakkoNO = cstr(C_NCT_KURUME) then
    'If m_sGakkoNO = cstr(C_NCT_KURUME) OR m_sGakkoNO = cstr(C_NCT_NUMAZU) then
	'    w_sSQL = w_sSQL & "   T16_MENJYO_FLG, "
	'End If

  w_sSQL = w_sSQL & w_Table & "_HYOKA_FUKA_KBN as HYOKA_FUKA "

  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL &     w_TableName & ","
  w_sSQL = w_sSQL & "   T11_GAKUSEKI,"
  w_sSQL = w_sSQL & "   T13_GAKU_NEN "

  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL &       w_Table & "_NENDO = " & Cint(m_iNendo)
  w_sSQL = w_sSQL & " AND " & w_Table & "_GAKUSEI_NO = T11_GAKUSEI_NO "
  w_sSQL = w_sSQL & " AND " & w_Table & "_GAKUSEI_NO = T13_GAKUSEI_NO "
  w_sSQL = w_sSQL & " AND T13_GAKUNEN = " & cint(m_sGakuNo)

  w_sSQL = w_sSQL & " AND " & w_KamokuName & " = '" & m_sKamokuCd & "' "
  w_sSQL = w_sSQL & " AND " & w_Table & "_NENDO = T13_NENDO "

  '//新居浜高専は免除科目は表示しない 2004/02/19
  if m_iKamokuKbn = C_JIK_JUGYO then
	  If m_sGakkoNO = cstr(C_NCT_NIIHAMA) then
		  w_sSQL = w_sSQL & " AND NVL(" & w_Table & "_MENJYO_FLG,0) <> 1 "
	  END IF
  end if

  if m_iKamokuKbn = C_JIK_JUGYO then
    '//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
    w_sSQL = w_sSQL & " AND T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
  end if

'****************************
  if m_iKongou <> C_CLASS_KONGO Then
    w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)

    'T13_GAKUSEKI_NOでソートするように修正。
    'T13とT16で学籍番号が違う場合が発生した為に発覚。
    '個人履修自動作成で学籍番号を更新していない可能性あり。2003.08.05 ITO
    w_sSQL = w_sSQL & " ORDER BY T13_GAKUSEKI_NO "

  else
    if m_sGakkoNO = C_NCT_KUMAMOTO Then
      w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
      w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "
    else
      if Cint(m_iKamoku_Kbn) <> C_KAMOKU_SENMON Then
        w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
        w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
      else
		if (m_sGakkoNO = C_NCT_MAIZURU) Then
			'INS 2007/06/11
			'--2019/06/17 DELETE FUJIBAYASHI(80001、80002の特別処理をやめる)
			'If (m_sKamokuCd = 80001) OR (m_sKamokuCd = 80002) Then
			'	w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
			'	w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
			'Else
			'--2019/06/17 DELETE END
			    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
				w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO1 "
			'End If		'--2019/06/17 DELETE FUJIBAYASHI(80001、80002の特別処理をやめる)
			'INS END 2007/06/11
			'DEL 2007/06/11  w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
			'DEL 2007/06/11  w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "
		else
			If m_sGakkoNO = C_NCT_OKINAWA Then
				If m_sKamokuCd >= 900000 Then
					w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
					w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
				Else
				    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
					w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO1 "
				End If
			Else
				w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO1 "
			End If
		end if
      end if
    end if
  end if
'****************************

  'w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "

'Response.Write w_sSQL
'response.end

  If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit function

'Response.Write w_sSQL

  m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
  m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))

  '福島高専用　INS 2005/06/13 西村	
  m_iSouJyugyou1 = gf_SetNull2String(m_Rs("SOUJI1"))
  m_iJunJyugyou1 = gf_SetNull2String(m_Rs("JYUNJI1"))
  m_iSouJyugyou2 = gf_SetNull2String(m_Rs("SOUJI2"))
  m_iJunJyugyou2 = gf_SetNull2String(m_Rs("JYUNJI2"))
  m_iSouJyugyou3 = gf_SetNull2String(m_Rs("SOUJI3"))
  m_iJunJyugyou3 = gf_SetNull2String(m_Rs("JYUNJI3"))
  m_iSouJyugyou4 = gf_SetNull2String(m_Rs("SOUJI4"))
  m_iJunJyugyou4 = gf_SetNull2String(m_Rs("JYUNJI4"))


  '//ﾚｺｰﾄﾞカウント取得
  m_rCnt = gf_GetRsCount(m_Rs)

  f_GetStudent = true

End Function


'********************************************************************************
'*  [機能]  授業時間数データの取得（福島高専の場合）
'*			作成 2005/12/16 西村
'*			MAX(授業時間数)で取得する
'********************************************************************************
Function f_GetJyugyoJIkan()

  Dim w_sSQL
  Dim w_FieldName
  Dim w_Table
  Dim w_TableName
  Dim w_KamokuName
  Dim m_RsMax

  On Error Resume Next
  Err.Clear

  f_GetJyugyoJIkan = false

  if m_iKamokuKbn = C_JIK_JUGYO then  '通常授業の場合
    w_Table = "T16"
    w_TableName = "T16_RISYU_KOJIN"
    w_KamokuName = "T16_KAMOKU_CD"
  else
    w_Table = "T34"
    w_TableName = "T34_RISYU_TOKU"
    w_KamokuName = "T34_TOKUKATU_CD"
  end if


  '//検索結果の値より一覧を表示
  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "

  Select Case m_sSikenKBN
    Case C_SIKEN_ZEN_TYU

      w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_TYUKAN_Z) as SOUJI,"
      w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_TYUKAN_Z) as JYUNJI, "


    Case C_SIKEN_ZEN_KIM

      w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_KIMATU_Z) as SOUJI, "
      w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_KIMATU_Z) as JYUNJI, "

    Case C_SIKEN_KOU_TYU

      w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_TYUKAN_K) as SOUJI, "
      w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_TYUKAN_K) as JYUNJI, "

    Case C_SIKEN_KOU_KIM


      w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_KIMATU_K) as SOUJI, "
      w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_KIMATU_K) as JYUNJI, "

  End Select

 '福島高専用　INS 2005/06/13 西村
'前期　中間
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_TYUKAN_Z) as SOUJI1, "
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_TYUKAN_Z) as JYUNJI1, "
'前期　期末
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_KIMATU_Z) as SOUJI2, "
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_KIMATU_Z) as JYUNJI2, "
'後期　中間
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_TYUKAN_K) as SOUJI3, "
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_TYUKAN_K) as JYUNJI3, "
'後期　期末
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_KIMATU_K) as SOUJI4, "
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_KIMATU_K) as JYUNJI4 "

  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL &     w_TableName & ","
  w_sSQL = w_sSQL & "   T11_GAKUSEKI,"
  w_sSQL = w_sSQL & "   T13_GAKU_NEN "

  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL &       w_Table & "_NENDO = " & Cint(m_iNendo)
  w_sSQL = w_sSQL & " AND " & w_Table & "_GAKUSEI_NO = T11_GAKUSEI_NO "
  w_sSQL = w_sSQL & " AND " & w_Table & "_GAKUSEI_NO = T13_GAKUSEI_NO "
  w_sSQL = w_sSQL & " AND T13_GAKUNEN = " & cint(m_sGakuNo)

  w_sSQL = w_sSQL & " AND " & w_KamokuName & " = '" & m_sKamokuCd & "' "
  w_sSQL = w_sSQL & " AND " & w_Table & "_NENDO = T13_NENDO "

  '//新居浜高専は免除科目は表示しない 2004/02/19
  if m_iKamokuKbn = C_JIK_JUGYO then
	  If m_sGakkoNO = cstr(C_NCT_NIIHAMA) then
		  w_sSQL = w_sSQL & " AND NVL(" & w_Table & "_MENJYO_FLG,0) <> 1 "
	  END IF
  end if

  if m_iKamokuKbn = C_JIK_JUGYO then
    '//置換元の生徒ははずす(C_TIKAN_KAMOKU_MOTO = 1    '置換元)
    w_sSQL = w_sSQL & " AND T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
  end if

'****************************
  if m_iKongou <> C_CLASS_KONGO Then
    w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)

  else
    if m_sGakkoNO = C_NCT_KUMAMOTO Then
      w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
    else
      if Cint(m_iKamoku_Kbn) <> C_KAMOKU_SENMON Then
        w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
      else
		if (m_sGakkoNO = C_NCT_MAIZURU) Then
			'INS 2007/06/11
			'--2019/06/17 DELETE FUJIBAYASHI(80001、80002の特別処理をやめる)
			'If (m_sKamokuCd = 80001) OR (m_sKamokuCd = 80002) Then
			'	w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
			'Else
			'--2019/06/17 DELETE END
			    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
			'End If		'--2019/06/17 DELETE FUJIBAYASHI(80001、80002の特別処理をやめる)
			'INS END 2007/06/11
			'DEL 2007/06/11 w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
		else
			If m_sGakkoNO = C_NCT_OKINAWA Then
				If m_sKamokuCd >= 900000 Then
					w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
				Else
				    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
				End If
			Else

			End If
		end if
      end if
    end if
  end if
'****************************


  If gf_GetRecordset(m_RsMax,w_sSQL) <> 0 Then Exit function


  m_iSouJyugyou = gf_SetNull2String(m_RsMax("SOUJI"))
  m_iJunJyugyou = gf_SetNull2String(m_RsMax("JYUNJI"))

  '福島高専用　INS 2005/06/13 西村	
  m_iSouJyugyou1 = gf_SetNull2String(m_RsMax("SOUJI1"))
  m_iJunJyugyou1 = gf_SetNull2String(m_RsMax("JYUNJI1"))
  m_iSouJyugyou2 = gf_SetNull2String(m_RsMax("SOUJI2"))
  m_iJunJyugyou2 = gf_SetNull2String(m_RsMax("JYUNJI2"))
  m_iSouJyugyou3 = gf_SetNull2String(m_RsMax("SOUJI3"))
  m_iJunJyugyou3 = gf_SetNull2String(m_RsMax("JYUNJI3"))
  m_iSouJyugyou4 = gf_SetNull2String(m_RsMax("SOUJI4"))
  m_iJunJyugyou4 = gf_SetNull2String(m_RsMax("JYUNJI4"))

   m_RsMax.close

f_GetJyugyoJIkan = true


End Function


'********************************************************************************
'*  [機能]  データの取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]
'********************************************************************************
Function f_Nyuryokudate()

  Dim w_sSysDate
  Dim w_Rs

  On Error Resume Next
  Err.Clear

  f_Nyuryokudate = false

  m_bKekkaNyuryokuFlg = false   '欠課入力ﾌﾗｸﾞ
  m_bSeiInpFlg = false

  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   T24_SEISEKI_KAISI, "
  w_sSQL = w_sSQL & "   T24_SEISEKI_SYURYO, "
  w_sSQL = w_sSQL & "   T24_KEKKA_KAISI, "
  w_sSQL = w_sSQL & "   T24_KEKKA_SYURYO, "
  w_sSQL = w_sSQL & "   T24_IDOU_SYURYO, "
  w_sSQL = w_sSQL & "   M01_SYOBUNRUIMEI, "
  w_sSQL = w_sSQL & "   SYSDATE "
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T24_SIKEN_NITTEI, "
  w_sSQL = w_sSQL & "   M01_KUBUN"
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   M01_SYOBUNRUI_CD = T24_SIKEN_KBN"
  w_sSQL = w_sSQL & " AND M01_NENDO = T24_NENDO"
  w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
  w_sSQL = w_sSQL & " AND T24_NENDO=" & Cint(m_iNendo)
  w_sSQL = w_sSQL & " AND T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
  w_sSQL = w_sSQL & " AND T24_SIKEN_CD='0'"
  w_sSQL = w_sSQL & " AND T24_GAKUNEN=" & Cint(m_sGakuNo)

  If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function

'   Response.Write "4" & w_sSQL

  If w_Rs.EOF Then
'   Response.Write "  EOF "
    exit function
  Else
    m_sSikenNm = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))    '試験名称
    m_iNKaishi = gf_SetNull2String(w_Rs("T24_SEISEKI_KAISI"))   '成績入力開始日
    m_iNSyuryo = gf_SetNull2String(w_Rs("T24_SEISEKI_SYURYO"))    '成績入力終了日
    m_iKekkaKaishi = gf_SetNull2String(w_Rs("T24_KEKKA_KAISI"))   '欠課入力開始
    m_iKekkaSyuryo = gf_SetNull2String(w_Rs("T24_KEKKA_SYURYO"))  '欠課入力終了

    m_iIdouEnd = gf_SetNull2String(w_Rs("T24_IDOU_SYURYO"))  '異動対象終了

    w_sSysDate = gf_SetNull2String(w_Rs("SYSDATE"))         'システム日付
  End If

  '入力期間内なら正常
  If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
    m_bSeiInpFlg = true
  End If

  '欠課入力可能ﾌﾗｸﾞ
  If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
    m_bKekkaNyuryokuFlg = True
  End If

  f_Nyuryokudate = true

End Function

'********************************************************************************
'*  [機能]  データの取得
'********************************************************************************
Function f_Syukketu2New(p_gaku,p_kbn)
  Dim w_GAKUSEI_NO
  Dim w_SYUKKETU_KBN

  f_Syukketu2New = ""
  w_GAKUSEI_NO = ""
  w_SYUKKETU_KBN = ""
  w_SKAISU = ""

  If m_SRs.EOF Then
    Exit Function
  Else
    Do Until m_SRs.EOF
      w_GAKUSEI_NO = m_SRs("T21_GAKUSEKI_NO")
      w_SYUKKETU_KBN = m_SRs("T21_SYUKKETU_KBN")
      w_SKAISU = gf_SetNull2String(m_SRs("KAISU"))

      If Cstr(w_GAKUSEI_NO) = Cstr(p_gaku) AND cstr(w_SYUKKETU_KBN) = cstr(p_kbn) Then
        f_Syukketu2New = w_SKAISU
        Exit Do
      End If

      m_SRs.MoveNext
    Loop

    m_SRs.MoveFirst
  End If

End Function

'********************************************************************************
'*  [機能]  確定欠課数、遅刻数を取得。
'*  [引数]  p_iNendo　 　：　処理年度
'*          p_iSikenKBN　：　試験区分
'*          p_sKamokuCD　：　科目コード
'*          p_sGakusei 　：　５年間番号
'*  [戻値]  p_iKekka   　：　欠課数
'*          p_ichikoku 　：　遅刻回数
'*          0：正常修了
'*  [説明]  試験区分に入っている、欠課数、遅刻数を取得する。
'*      2002.03.20
'*      NULLを0に変換しないために、関数をモジュール内で作成（CACommon.aspからコピー）
'********************************************************************************
Function f_GetKekaChi(p_iNendo,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku)
  Dim w_sSQL
  Dim w_Rs
  Dim w_sKek,w_sChi
  Dim w_Table,w_TableName
  Dim w_Kamoku

  On Error Resume Next
  Err.Clear

  f_GetKekaChi = false

  p_iKekka = ""
  p_iChikoku = ""

  '特別授業、その他(通常など)の切り分け
  if trim(m_sSyubetu) = "TOKU" then
    w_Table = "T34"
    w_TableName = "T34_RISYU_TOKU"
    w_Kamoku = "T34_TOKUKATU_CD"
  else
    w_Table = "T16"
    w_TableName = "T16_RISYU_KOJIN"
    w_Kamoku = "T16_KAMOKU_CD"
  end if

  '/試験区分によって取ってくる、フィールドを変える。
  Select Case p_iSikenKBN
    Case C_SIKEN_ZEN_TYU
      w_sKek   = w_Table & "_KEKA_TYUKAN_Z"
      w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_Z"
      w_sChi   = w_Table & "_CHIKAI_TYUKAN_Z"
    Case C_SIKEN_ZEN_KIM
      w_sKek   = w_Table & "_KEKA_KIMATU_Z"
      w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_Z"
      w_sChi   = w_Table & "_CHIKAI_KIMATU_Z"
    Case C_SIKEN_KOU_TYU
      w_sKek   = w_Table & "_KEKA_TYUKAN_K"
      w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_K"
      w_sChi   = w_Table & "_CHIKAI_TYUKAN_K"
    Case C_SIKEN_KOU_KIM
      w_sKek   = w_Table & "_KEKA_KIMATU_K"
      w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_K"
      w_sChi   = w_Table & "_CHIKAI_KIMATU_K"
  End Select

  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL &   w_sKek   & " as KEKA, "
  w_sSQL = w_sSQL &   w_sKekG  & " as KEKA_NASI, "
  w_sSQL = w_sSQL &   w_sChi   & " as CHIKAI "
  w_sSQL = w_sSQL & " FROM "   & w_TableName
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "      " & w_Table & "_NENDO =" & p_iNendo
  w_sSQL = w_sSQL & "  AND " & w_Table & "_GAKUSEI_NO= '" & p_sGakusei & "'"
  w_sSQL = w_sSQL & "  AND " & w_Kamoku & "= '" & p_sKamokuCD & "'"


  If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then exit function

  ' Response.Write "5"

  '//戻り値ｾｯﾄ
  If w_Rs.EOF = False Then
    p_iKekka = gf_SetNull2String(w_Rs("KEKA"))
    p_iChikoku = gf_SetNull2String(w_Rs("CHIKAI"))
  End If

  f_GetKekaChi = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能] 異動チェック
'********************************************************************************
Sub s_IdouCheck(p_GakusekiNo,p_IdouKbn,p_IdouName,p_bNoChange)
  Dim w_IdoutypeName  '異動状況名

  w_IdoutypeName = ""
  p_IdouName = ""


  m_Date = m_iIdouEnd
'debug
'response.write "m_Date = " & m_Date & "<br>"

  p_IdouKbn = gf_Get_IdouChk(p_GakusekiNo,m_Date,m_iNendo,w_IdoutypeName)

  if Cstr(p_IdouKbn) <> "" and Cstr(p_IdouKbn) <> CStr(C_IDO_FUKUGAKU) AND _
    Cstr(p_IdouKbn) <> Cstr(C_IDO_TEI_KAIJO) AND _
    Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKA) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_KOKUHI) Then

    p_IdouName = "[" & w_IdoutypeName & "]"
    p_bNoChange = True
  end if

end Sub

'********************************************************************************
'*  [機能] 成績のセット
'********************************************************************************
Sub s_SetGrades(p_sSeiseki,p_sHyoka,p_bNoChange)

  p_sSeiseki = gf_SetNull2String(m_Rs("SEI"))


  '2002/12/25
  '八代高専の場合、前期開設科目は試験によってコピー元の試験を変える
    If m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then

    '後期中間の時、前期期末からセット
    If m_sSikenKBN = C_SIKEN_KOU_TYU and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))  '前期期末
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_TK"))  '後期中間

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))
      End If
    End If

    '後期期末の時、後期中間からセット
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_TK"))  '後期中間
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))  '後期期末

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sSeiseki = gf_SetNull2String(m_Rs("SEI_KT"))
      End If
    End If

  '他高専（前期末->後期末にセット）
  Else
    '学年末試験の場合のみ
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
      'If gf_SetNull2String(m_Rs("SEI")) = "" Then
        p_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))
      End If
    End If

  End If



  '//通常授業のとき
  if m_iKamokuKbn = C_JIK_JUGYO then

    if m_HyokaDispFlg then
      p_sHyoka = gf_SetNull2String(m_Rs("HYOKAYOTEI"))
      if p_sHyoka = "" then p_sHyoka = "・"
    end if

    p_bNoChange = False

	'2004.02.20 ITO
	'久留米の場合、免除科目の点数は非表示
	If m_sGakkoNO = cstr(C_NCT_KURUME) then

		if cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG"))) = "1" Then

			p_bNoChange = True

		End If

	End If

    '//科目が選択科目の場合は、生徒が選択しているかどうかを判別する。選択しいない生徒は入力不可とする。
    if cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then

		if cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then p_bNoChange = True

    else
      if Cstr(m_iLevelFlg) = "1" then
        if isNull(m_Rs("T16_LEVEL_KYOUKAN")) = true then
          p_bNoChange = True
        else
          if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
            p_bNoChange = True
          End if
        End if
      End if
    end if

  end if

end Sub

'********************************************************************************
'*  [機能] 欠課、遅刻の日々計の取得
'********************************************************************************
Sub s_SetKekkaTotal(p_sKekkasu,p_sChikaisu)
  Dim w_sData
  Dim w_iKekka_rui,w_iChikoku_rui

  '//欠課
  p_sKekkasu = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_KEKKA))

  '//1欠課
  w_sData = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_KEKKA_1))

  if p_sKekkasu = "" and w_sData = "" then
    p_sKekkasu = ""
  else
    p_sKekkasu = cint(gf_SetNull2Zero(p_sKekkasu)) + cint(gf_SetNull2Zero(w_sData))
  end if

  '//遅刻数
  p_sChikaisu = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_TIKOKU))

  '//早退数
  w_sData = f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_SOTAI)

  if p_sChikaisu = "" and w_sData = "" then
    p_sChikaisu = ""
  else
    p_sChikaisu = cint(gf_SetNull2Zero(p_sChikaisu)) + cint(gf_SetNull2Zero(w_sData))
  end if

  '「出欠欠課が累積」で「前期中間でない」の場合

  if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and m_sSikenKBN <> C_SIKEN_ZEN_TYU then
    '以前の試験で登録されているデータを取得
    call f_GetKekaChi(m_iNendo,m_iShikenInsertType,m_sKamokuCd,m_Rs("GAKUSEI_NO"),w_iKekka_rui,w_iChikoku_rui)

    'どちらも""の時は""
    if p_sKekkasu = "" and w_iKekka_rui = "" then
      p_sKekkasu = ""
    else
      p_sKekkasu = cint(gf_SetNull2Zero(p_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))
    end if

    'どちらも""の時は""
    if p_sChikaisu = "" and w_iChikoku_rui = "" then
      p_sChikaisu = ""
    else
      p_sChikaisu = cint(gf_SetNull2Zero(p_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
    end if
  end if

End Sub

'********************************************************************************
'*  [機能]  欠課、遅刻数のセット
'********************************************************************************
Sub s_SetKekka(p_sKekka,p_sKekkaGai,p_sChikai)

  p_sKekka = gf_SetNull2String(m_Rs("KEKA"))
  p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI"))
  p_sChikai = gf_SetNull2String(m_Rs("CHIKAI"))


  '2002/12/25
  '八代高専の場合、前期開設科目は試験によってコピー元の試験を変える
If m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then

    '後期中間の時、前期期末からセット
    If m_sSikenKBN = C_SIKEN_KOU_TYU and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))  '前期期末
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_TK"))  '後期中間

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))     '欠課数
        p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK")) '欠課対象外
        p_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))    '遅刻回数
      End If
    End If

    '後期期末の時、後期中間からセット
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_TK"))  '後期中間
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))  '後期期末

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sKekka = gf_SetNull2String(m_Rs("KEKA_KT"))     '欠課数
        p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_KT")) '欠課対象外
        p_sChikai = gf_SetNull2String(m_Rs("CHIKAI_KT"))    '遅刻回数
      End If
    End If

 '他高専（前期末->後期末にセット）
 Else

    '//学年末試験の場合のみ
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

      'If gf_SetNull2String(m_Rs("KEKA")) = "" Then
      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))     '欠課数
        p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK")) '欠課対象外
        p_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))    '遅刻回数
      End If
    End If
 End If

End Sub

'********************************************************************************
'*  [機能]  評価不能処理(熊本電波のみ)
'********************************************************************************
Sub s_SetHyoka(p_IdouKbn,p_DataKbn,p_Checked,p_Disabled2,p_Disabled)

  '//評価不能データ設定
  p_DataKbn = 0
  p_Checked = ""
  p_Disabled2 = ""

  p_DataKbn = cint(gf_SetNull2Zero(m_Rs("DataKbn")))

  If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
    w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
    w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

    if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then

      p_DataKbn = cint(gf_SetNull2Zero(m_Rs("DataKbn_ZK")))

    end if
  end if

  if p_Disabled <> "" then p_Disabled2 = "disabled"

  if p_DataKbn = cint(C_HYOKA_FUNO) then
    p_Checked = "checked"
    p_Disabled2 = "disabled"

  elseif p_DataKbn = cint(C_MIHYOKA) then
    p_Disabled2 = "disabled"

  end if

  if not m_bSeiInpFlg Then p_Disabled2 = ""

  select case Cstr(p_IdouKbn)
    case Cstr(C_IDO_KYU_BYOKI),Cstr(C_IDO_KYU_HOKA)
      p_DataKbn = C_KYUGAKU

    case Cstr(C_IDO_TAI_2NEN),Cstr(C_IDO_TAI_HOKA),Cstr(C_IDO_TAI_SYURYO)
      p_DataKbn = C_TAIGAKU
  end select

End Sub

'********************************************************************************
'*  [機能]  テーブルサイズのセット
'********************************************************************************
Sub s_SetTableWidth(p_TableWidth)

  p_TableWidth = 610

  '//評価不能処理がある(熊本電波のみ)
  if m_SchoolFlg then
    p_TableWidth = 660
  end if

  '//評価予定表示フラグオン、または、通常授業のとき
  if m_HyokaDispFlg and Cstr(m_iKamokuKbn) = Cstr(C_TUKU_FLG_TUJO) then
    p_TableWidth = p_TableWidth + 50
  end if

  '//欠課外表示フラグオン
  if m_KekkaGaiDispFlg then
    p_TableWidth = p_TableWidth + 55
  end if

End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
Sub showPage()
  Dim w_sSeiseki
  Dim w_sHyoka

  Dim w_sChikai
  Dim w_sChikaisu

  Dim w_sKekka
  Dim w_sKekkaGai
  Dim w_sKekkasu

  Dim i

  Dim w_lSeiTotal '成績合計
  Dim w_lGakTotal '学生人数

  Dim w_IdouKbn '異動タイプ
  Dim w_IdouName

  Dim w_sInputClass
  Dim w_sInputClass1
  Dim w_sInputClass2

  Dim w_Padding
  Dim w_Padding2

  Dim w_Disabled
  Dim w_Disabled2
  Dim w_TableWidth

  '//遅刻、欠課(対象)、欠課(対象外)の合計
  Dim wChikokuSum,wKekkaSum,wKekkaGaiSum

  wChikokuSum = 0
  wKekkaSum = 0
  wKekkaGaiSum = 0

  w_Padding = "style='padding:2px 0px;'"
  w_Padding2 = "style='padding:2px 0px;font-size:10px;'"

  w_lSeiTotal = 0
  w_lGakTotal = 0
  i = 1

  '//NN対応
  If session("browser") = "IE" Then
    w_sInputClass  = "class='num'"
    w_sInputClass1 = "class='num'"
    w_sInputClass2 = "class='num'"
  Else
    w_sInputClass = ""
    w_sInputClass1 = ""
    w_sInputClass2 = ""
  End If

  '//テーブルサイズのセット
  Call s_SetTableWidth(w_TableWidth)

  if m_SchoolFlg then
    if m_MiHyokaFlg or (not m_bSeiInpFlg) then
      w_Disabled = "disabled"
    end if
  end if
%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--
  //************************************************************
  //  [機能]  ページロード時処理
  //************************************************************
  function window_onload() {
    //スクロール同期制御
    parent.init();
    //数値入力のときのみ
    <% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
      //成績合計値の取得
      f_GetTotalAvg();
    <% end if %>

    //総時間と純時間をhiddenにセット
    document.frm.hidSouJyugyou.value = "<%= m_iSouJyugyou %>";
    document.frm.hidJunJyugyou.value = "<%= m_iJunJyugyou %>";


  // INS 2005/06/13 西村　福島高専用 
  document.frm.hidSouJyugyou_TZ.value = "<%= m_iSouJyugyou1 %>";
  document.frm.hidJunJyugyou_TZ.value = "<%= m_iJunJyugyou1 %>";
  document.frm.hidSouJyugyou_KZ.value = "<%= m_iSouJyugyou2 %>";
  document.frm.hidJunJyugyou_KZ.value = "<%= m_iJunJyugyou2 %>";
  document.frm.hidSouJyugyou_TK.value = "<%= m_iSouJyugyou3 %>";
  document.frm.hidJunJyugyou_TK.value = "<%= m_iJunJyugyou3 %>";
  document.frm.hidSouJyugyou_KK.value = "<%= m_iSouJyugyou4 %>";
  document.frm.hidJunJyugyou_KK.value = "<%= m_iJunJyugyou4 %>";

    document.frm.target = "topFrame";
    document.frm.action = "sei0150_middle.asp";
    document.frm.submit();
  }

  //************************************************************
  //  [機能]  評価ボタンが押されたとき
  //************************************************************
  function f_change(p_iS){
    w_sButton = eval("document.frm.button"+p_iS);
    w_sHyouka = eval("document.frm.Hyoka"+p_iS);

    <%If m_sSikenKBN = C_SIKEN_ZEN_TYU Then%>
      if(w_sButton.value == "・") {
        w_sButton.value = "○";
        w_sHyouka.value = "○";
        return true;
      }
      if(w_sButton.value == "○") {
        w_sButton.value = "・";
        w_sHyouka.value = "";
        return true;
      }

    <%Else%>

      if(w_sButton.value == "・") {
        w_sButton.value = "○";
        w_sHyouka.value = "○";
        return true;
      }
      if(w_sButton.value == "○") {
        w_sButton.value = "◎";
        w_sHyouka.value = "◎";
        return true;
      }
      if(w_sButton.value == "◎") {
        w_sButton.value = "・";
        w_sHyouka.value = "";
        return true;
      }
    <%End If%>
  }

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //************************************************************
    function f_Touroku(){
    if(!f_InpCheck()){
      alert("入力値が不正です");
      return false;
    }

    //純時間のみ入力チェックを追加 2003.08.04 ITO
    if(parent.topFrame.document.frm.txtJunJyugyou.value == ""){
      parent.topFrame.document.frm.txtJunJyugyou.focus();
      alert("純授業時間数が入力されていません");
      return false;
    }

    // 久留米高専の場合
    <% If (m_sGakkoNO = C_NCT_KURUME) AND (m_bSeiInpFlg) AND (m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM) AND (m_bKekkaNyuryokuFlg) Then %>
      if(!jf_CheckInpVal()){
        alert("受講時数が入力されていません");
        return false;
      }

	  if(parent.topFrame.document.frm.txtSouJyugyou.value < (<%=m_lHaitoTani%> * 30)){
			    parent.topFrame.document.frm.txtSouJyugyou.focus();
		        alert("総授業時間数が足りません。単位数×30時間以上で入力してください。");
		        return false;
	  }

	  if(parent.topFrame.document.frm.txtJunJyugyou.value < (<%=m_lHaitoTani%> * 24)){
	    parent.topFrame.document.frm.txtJunJyugyou.focus();
        alert("純授業時間数が足りません。単位数×24時間以上で入力してください。");
        return false;
	  }
    <% End If %>

    if(!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return false;}
    document.frm.hidSouJyugyou.value = parent.topFrame.document.frm.txtSouJyugyou.value;
    document.frm.hidJunJyugyou.value = parent.topFrame.document.frm.txtJunJyugyou.value;

    //ヘッダ部空白表示
    parent.topFrame.document.location.href="white.asp";

    //登録処理
    <% if m_iKamokuKbn = C_JIK_JUGYO then %>
      document.frm.hidUpdMode.value = "TUJO";
      document.frm.action="sei0150_upd.asp";
    <% Else %>
      document.frm.hidUpdMode.value = "TOKU";
      document.frm.action="sei0150_upd_toku.asp";
    <% End if %>
    document.frm.target="main";
    document.frm.submit();
  }

  //************************************************************
  //  [機能]  キャンセルボタンが押されたとき
  //************************************************************
  function f_Cancel(){
    parent.document.location.href="default.asp";
  }

  //************************************************************
  //  [機能]  成績の合計と平均を求める
  //  [引数]  なし
  //  [戻値]  なし
  //  [説明]  成績入力期間外、期間内によって計算の仕方を変える
  //  [備考]
  //************************************************************
  function f_GetTotalAvg(){
    var i;
    var total;
    var avg;
    var cnt;

    total = 0;
    cnt = 0;
    avg = 0;

    <% If m_bSeiInpFlg Then %>
      //学生数でのループ
      for(i=0;i<<%=m_rCnt%>;i++) {
        //存在するかどうか
        textbox = eval("document.frm.Seiseki" + (i+1));
        if(textbox){
          //未入力チェック
          if (textbox.value != "") {
            //数字でないのは無視する
            if(!isNaN(textbox.value)){
              total = total + parseInt(textbox.value);
              cnt = cnt + 1;
            }
          }
        }
      }

    <% Else %>
      total = document.frm.hidTotal.value;
      cnt   = document.frm.hidGakTotal.value;
    <% End If%>

    document.frm.txtTotal.value = total;

    //四捨五入
    if (cnt!=0){
      avg = total/cnt;
      avg = avg * 10;
      avg = Math.round(avg);
      avg = avg / 10;
    }

    document.frm.txtAvg.value=avg;
  }

    //************************************************************
    //  [機能]  数値型チェック
    //************************************************************
  function f_CheckNum(pFromName){
    var wFromName,w_len;

    wFromName = eval(pFromName);

    if(isNaN(wFromName.value)){
      wFromName.focus();
      wFromName.select();
      return false;
    }else{
      //桁チェック
      if(wFromName.name.indexOf("Seiseki") != -1){
        if(wFromName.value > 100){
          wFromName.focus();
          wFromName.select();
          return false;
        }
      }

      //遅刻は、2桁まで
      if(wFromName.name.indexOf("Chikai") != -1){
        w_len = 2;
      }else{
        w_len = 3;
      }

      if(wFromName.value.length > w_len){
        wFromName.focus();
        wFromName.select();
        return false;
      }

      //マイナスをチェック
      var wStr = new String(wFromName.value)
      if (wStr.match("-")!=null){
        wFromName.focus();
        wFromName.select();
        return false;
      }

      if(wFromName.name.indexOf("txtAvg") == -1){
        //小数点チェック
        w_decimal = new Array();
        w_decimal = wStr.split(".")

        if(w_decimal.length>1){
          wFromName.focus();
          wFromName.select();
          return false;
        }
      }
    }

    return true;
  }

    //************************************************************
    //  [機能]  久留米高専時の成績と受講時数の入力チェック
  //  [詳細]　成績 & 受講時間共に入力可の時
  //  　　　　成績が入力されている場合は受講時間を必須入力とする
  //  [作成]  2003.05.13 hirota
    //************************************************************
  function jf_CheckInpVal(){
    //学生数でのループ
    for(i=0;i<<%=m_rCnt%>;i++) {
      //存在するかどうか
      textbox1 = eval("document.frm.Seiseki" + (i+1));  // 成績
      textbox2 = eval("document.frm.Kekka" + (i+1));    // 欠課
      if(textbox2){
        //未入力チェック
        if (textbox2.value == "") {
          if(textbox1){
            if(textbox1.value != ""){
              textbox2.focus();
              return false;
            }
          }
        }
      }
    }
    return true;
  }

    //************************************************************
    //  [機能]  大小チェック
    //************************************************************
  function f_CheckDaisyou(){
    wObj1 = eval("parent.topFrame.document.frm.txtSouJyugyou");
    wObj2 = eval("parent.topFrame.document.frm.txtJunJyugyou");

    if(wObj1.value != "" && wObj2.value != ""){
      if(wObj1.value < wObj2.value){
        wObj1.focus();
        return false;
      }
    }
    return true;
  }

  //************************************************
  //Enter キーで下の入力フォームに動くようになる
  //引数：p_inpNm 対象入力フォーム名
  //    ：p_frm 対象フォーム
  //　　：i   現在の番号
  //戻値：なし
  //入力フォーム名が、xxxx1,xxxx2,xxxx3,…,xxxxn
  //の名前のときに利用できます。
  //************************************************
  function f_MoveCur(p_inpNm,p_frm,i){
    if (event.keyCode == 13){   //押されたキーがEnter(13)の時に動く。
      i++;

      //入力可能のテキストボックスを探す。見つかったらフォーカスを移して処理を抜ける。
          for (w_li = 1; w_li <= 99; w_li++) {

        if (i > <%=m_rCnt%>) i = 1; //iが最大値を超えると、はじめに戻る。
        inpForm = eval("p_frm."+p_inpNm+i);

        //入力可能領域ならフォーカスを移す。
        if (typeof(inpForm) != "undefined") {
          inpForm.focus();      //フォーカスを移す。
          inpForm.select();     //移ったテキストボックス内を選択状態にする。
          break;
        //入力付加なら次の項目へ
        } else{
          i++
        }
          }
    }else{
      return false;
    }
    return true;
  }

  //************************************************
  //  文字入力時の成績処理
  //
  //************************************************
  function f_SetSeiseki(w_num){
    var ob = new Array();

    ob[0] = eval("parent.topFrame.document.frm.sltHyoka");
    ob[1] = eval("document.frm.Seiseki" + w_num);
    ob[2] = eval("document.frm.hidSeiseki" + w_num);
    ob[3] = eval("document.frm.hidHyokaFukaKbn" + w_num);

    if(ob[0].value.length == 0){
      ob[1].value = "";
      ob[2].value = "";
      ob[3].value = 0;
    }else{
      var vl = ob[0].value.split('#@#');

      ob[1].value = vl[0];
      ob[2].value = vl[0];
      ob[3].value = vl[1];
    }
  }

  //************************************************
  //  入力チェック
  //************************************************
  function f_InpCheck(){
    var w_length;
    var ob;

    //総時間・純時間入力チェック
    if(!f_CheckNum("parent.topFrame.document.frm.txtSouJyugyou")){ return false; }
    if(!f_CheckNum("parent.topFrame.document.frm.txtJunJyugyou")){ return false; }
    // 大小チェック不要 2003.02.20
    // if(!f_CheckDaisyou()){ return false; }

    w_length = document.frm.elements.length;

    for(i=0;i<w_length;i++){
      ob = eval("document.frm.elements[" + i + "]")

      //if(ob.type=="text" && ob.name != "txtAvg"  && ob.name != "txtTotal"){
      if(ob.type=="text" && ob.name != "txtAvg"  && ob.name != "txtTotal"  && ob.name != "ChikaiSum"  && ob.name != "KekkaSum"){
        ob = eval("document.frm." + ob.name);

        if(!f_CheckNum(ob)){return false;}
      }
    }
    return true;
  }

  //************************************************
  //評価不能がクリックされたときの処理
  //************************************************
  function f_InpDisabled(p_num){

    <% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
      var ob = new Array();

      ob[0] = eval("document.frm.chkHyokaFuno" + p_num);
      ob[1] = eval("document.frm.Seiseki" + p_num);

      if(ob[0].checked){
        ob[1].value = "";
        ob[1].disabled = true;

      }else{
        ob[1].disabled = false;
      }
    <% end if %>

    //数値入力のときのみ
    <% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
      f_GetTotalAvg();
    <% end if %>
  }

  //************************************************************
  //  [機能]  遅刻、欠課、欠課外の合計の計算
  //************************************************************
  function f_CalcSum(p_Name){
    var w_num;
    var total = 0;
    var cnt = 0;

    var ob_kei = eval("document.frm." + p_Name + "Sum")

    //学生数でのループ
    for(w_num=0;w_num<<%=m_rCnt%>;w_num++) {
      //存在するかどうか
      textbox = eval("document.frm." + p_Name + (w_num+1));
      if(textbox){
        //未入力チェック
        if (textbox.value != "") {
          //数字でないのは無視する
          if(!isNaN(textbox.value)){
            total = total + parseInt(textbox.value);
          }
        }
        cnt = cnt + 1;
      }
    }

    ob_kei.value = total;
  }

  //-->
  </SCRIPT>
  </head>
  <body LANGUAGE="javascript" onload="window_onload();">
  <center>
  <form name="frm" method="post">

  <table width="<%=w_TableWidth%>">
  <tr>
  <td>

  <table class="hyo" align="center" width="<%=w_TableWidth%>" border="1">
  <%

    m_Rs.MoveFirst

    Do Until m_Rs.EOF
      j = j + 1

      w_sSeiseki  = ""
      w_sHyoka    = ""
      w_sChikai   = ""
      w_sChikaisu = ""
      w_sKekka    = ""
      w_sKekkaGai = ""
      w_sKekkasu  = ""
      w_bNoChange = false

      Call gs_cellPtn(w_cell)

      'スタイルシート設定
      if not m_bSeiInpFlg Then
        w_sInputClass1 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
        w_Disabled = "disabled"
      End if

      if Not m_bKekkaNyuryokuFlg Then
        w_sInputClass2 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
      End if

      '//欠課、遅刻数のセット
      Call s_SetKekka(w_sKekka,w_sKekkaGai,w_sChikai)

      '//成績データセット
      Call s_SetGrades(w_sSeiseki,w_sHyoka,w_bNoChange)


      '//異動チェック

      Call s_IdouCheck(m_Rs("GAKUSEKI_NO"),w_IdouKbn,w_IdouName,w_bNoChange)

      '//欠課、遅刻の日々計の取得
      Call s_SetKekkaTotal(w_sKekkasu,w_sChikaisu)

    '// 2003/10/06 INSERT
	if (m_sGakkoNO <> cstr(C_NCT_NUMAZU)) and (m_sGakkoNO <> cstr(C_NCT_NIIHAMA)) Then
      '//欠入が0で,欠計が0より大きい場合
      if cint(gf_SetNull2Zero(w_sKekka)) = 0 and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then
        w_sKekka = cint(gf_SetNull2Zero(w_sKekkasu))
      end if
	end if

    '// 2003/10/06 INSERT
	if (m_sGakkoNO <> C_NCT_NUMAZU) and (m_sGakkoNO <> C_NCT_NIIHAMA) Then
      '//遅入が0で,遅計が0より大きい場合
      if cint(gf_SetNull2Zero(w_sChikai)) = 0 AND cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then
        w_sChikai = cint(gf_SetNull2Zero(w_sChikaisu))
      end if
	end if

      '//評価不能処理(熊本電波のみ)
      if m_SchoolFlg  then
        Call s_SetHyoka(w_IdouKbn,w_DataKbn,w_Checked,w_Disabled2,w_Disabled)
      end if

      %>

      <tr>
        <td class="<%=w_cell%>" align="center" width="65"  nowrap <%=w_Padding%>><%=m_Rs("GAKUSEKI_NO")%></td>
        <input type="hidden" name="txtGseiNo<%=i%>"   value="<%=m_Rs("GAKUSEI_NO")%>">
        <input type="hidden" name="hidNoChange<%=i%>" value="<%=w_bNoChange%>">

      	<% If m_sGakkoNO = cstr(C_NCT_KURUME) then %>

      		<!-- 免除フラグが１の場合  -->
      		<% If cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG"))) = "1" Then %>

		        <td class="<%=w_cell%>" align="left" width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%>[修得済]</td>

			<% Else %>

		        <td class="<%=w_cell%>" align="left" width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>

			<% End If %>

		<% Else %>

		    <td class="<%=w_cell%>" align="left" width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>

		<% End If %>

        <% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>

          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI1")))%></td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI2")))%></td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI3")))%></td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI4")))%></td>
        <% else %>

          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
        <% end if %>

        <!--選択科目の時に未選択の場合、入力不可。また、休学など-->
        <% If w_bNoChange = True Then %>

          <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>

          <% if m_HyokaDispFlg and m_iKamokuKbn = C_JIK_JUGYO then %>

            <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
          <% end if %>


          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>

          <% if m_KekkaGaiDispFlg then %>

            <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
          <% end if %>

          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>

          <% if m_SchoolFlg then %>

            <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
            <input type="hidden" name="chkHyokaFuno<%=i%>" value="<%=w_DataKbn%>">
          <% end if %>

        <% Else %>

          <!-- 成績 (数値入力、文字入力、成績なし入力により処理を分ける) -->
          <!--
          20040218 修正 shiki
          沼津の場合、免除フラグが１の場合は、成績のテキストBOXはbを付けて表示のみ
          未評価チェックボックスをクリックされても何もしないように
          hidMenjiFlg = 1 を立てる
          -->
		<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>

			<!-- 沼津  -->
          	<% If m_sGakkoNO = cstr(C_NCT_NUMAZU) then %>

          		<!-- 免除フラグが１の場合  -->
          		<% If cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG"))) = "1" Then %>
					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><font size="2">b</font>&nbsp;<input type="text" <%=w_sInputClass1%>  name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=3 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);" onChange="f_GetTotalAvg();" style="border-color:#FFFFFF #FFFFFF #FFFFFF #FFFFFF; border-style: solid; text-align:left; vertical-align:middle; " readonly></td>

					<!-- 未評価チェックボックスをクリックされた時用にフラグを立てる -->
					<input type="hidden" name="hidMenjiFlg<%=i%>" value="1">

				<% Else %>
					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass1%>  name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=3 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);" onChange="f_GetTotalAvg();"></td>

				<% End If %>

			<!-- 沼津以外  -->
			<% Else %>

					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass1%>  name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=3 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);" onChange="f_GetTotalAvg();"></td>

			<% End If %>

			<!-- END -->

		<% elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>

				<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>

				<% if not m_bSeiInpFlg Then %>
					<%=w_sSeiseki%>
				<% else %>
					<input type="button" class="<%=w_cell%>" style="text-align:center;" name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=2 onClick="f_SetSeiseki(<%=i%>);" <%=w_Disabled%>>
				<% end if %>

				</td>
				<input type="hidden" name="hidSeiseki<%=i%>" value="<%=w_sSeiseki%>">
				<input type="hidden" name="hidHyokaFukaKbn<%=i%>" value="<%=m_Rs("HYOKA_FUKA")%>">

		<% else %>

			<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>

		<% end if %>

          <!-- 評価予定 -->
          <% If m_HyokaDispFlg and m_iKamokuKbn = C_JIK_JUGYO then %>
            <% if m_bSeiInpFlg and (m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU) then %>
              <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>
                <input type="button" name="button<%=i%>" value="<%=w_sHyoka%>" size="2" onClick="f_change(<%=i%>);" class="<%=w_cell%>" style="text-align:center">
                <input type="hidden" name="Hyoka<%=i%>"  value="<%=trim(w_sHyoka)%>">
              </td>
            <% else %>
              <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sHyoka)%></td>
            <% end if %>
          <% end if %>

		<!-- 遅刻 -->
          <td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)" onChange="f_CalcSum('Chikai');" readonly = true></td>
		<!-- 欠課 -->
          <td class="<%=w_cell%>" align="right"  width="55" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sChikaisu)%></td>
          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)" onChange="f_CalcSum('Kekka');"></td>

          <% if m_KekkaGaiDispFlg then %>
            <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)" onChange="f_CalcSum('KekkaGai');"></td>
          <% end if %>

          <td class="<%=w_cell%>" align="right"  width="55" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sKekkasu)%></td>

		<!--公休 沼津のみ-->
		<%if m_sGakkoNO = cstr(C_NCT_NUMAZU) THEN %>
         <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>
			<input type="text" <%=w_sInputClass2%>  name=Kokyu<%=i%> value="<%=m_Rs("KEKA_NASI")%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kokyu',this.form,<%=i%>);">
         </td>
		<% end if %>

		<!-- 評価不能処理 -->
		<!--
          20040218 修正 shiki
          沼津の場合、免除フラグが１の場合は、評価不能のチェックBOXはDISABLED
		-->
		<% if m_SchoolFlg then %>
			<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>>
			<% if w_DataKbn = C_HYOKA_FUNO or w_DataKbn = C_MIHYOKA or w_DataKbn = 0 then %>

				<!-- 沼津  -->
				<% If m_sGakkoNO = cstr(C_NCT_NUMAZU) then %>
					<!-- 免除フラグが１の場合  -->
					<% If cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG"))) = "1" Then %>
						<input type="checkbox" name="chkHyokaFuno<%=i%>" <%=w_Disabled%> value="3"  <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);" disabled>
					<% Else %>
						<input type="checkbox" name="chkHyokaFuno<%=i%>" <%=w_Disabled%> value="3"  <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);">
					<% End If %>
				<% Else %>

					<input type="checkbox" name="chkHyokaFuno<%=i%>" <%=w_Disabled%> value="3"  <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);">

				<% End If %>

			<% else %>
				<input type="hidden" name="chkHyokaFuno<%=i%>" value="<%=w_DataKbn%>">
			<% end if %>

			</td>
 		<% end if %>
		<!-- END -->

          <%
            if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
              '表示のみの場合の合計・平均値を求める
              If IsNull(w_sSeiseki) = False and IsNumeric(CStr(w_sSeiseki)) = True Then
                w_lSeiTotal = w_lSeiTotal + CLng(w_sSeiseki)
                w_lGakTotal = w_lGakTotal + 1
              End If
            end if
          %>
        <%End If%>

		<%
		'2003.8.25 ITO
		'2004.02.18 沼津も免除科目が発生するようになった為、下記高専に沼津も含める
		'久留米の場合、免除科目は前期末成績を学年末にコピーしない為、免除フラグを設定
		If m_sGakkoNO = cstr(C_NCT_KURUME) OR m_sGakkoNO = cstr(C_NCT_NUMAZU) then
		'If m_sGakkoNO = cstr(C_NCT_KURUME) then
		%>
			<input type="hidden" name="hidMenjyo<%=i%>" value="<%=cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG")))%>">
		<%

		'その他の学校は、全て通常の科目で処理する為に0を設定
		Else
		%>
			<input type="hidden" name="hidMenjyo<%=i%>" value="0">
		<%
		End If
		%>

      </tr>

      <%
        wChikokuSum = wChikokuSum + cint(gf_SetNull2Zero(w_sChikai))
		If w_bNoChange = false Then
        	wKekkaSum = wKekkaSum + cint(gf_SetNull2Zero(w_sKekka))
			wKekkaGaiSum = wKekkaGaiSum + cint(gf_SetNull2Zero(w_sKekkaGai))
		'Else

		End If


        m_Rs.MoveNext
        i = i + 1
      Loop

      %>

      <% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
        <tr>
          <td class="header" align="right" colspan="7" nowrap>
            <FONT COLOR="#FFFFFF"><B>合計</B></FONT>
            <input type="text" name="txtTotal" size="5" <%=w_sInputClass%> readonly>
          </td>

          <td class="header" align="center" nowrap><input type="text" name="ChikaiSum" size="4" <%=w_sInputClass%> readonly value="<%=wChikokuSum%>"></td>
          <td class="header" align="center" nowrap>&nbsp;</td>
          <td class="header" align="center" nowrap><input type="text" name="KekkaSum" size="4" <%=w_sInputClass%> readonly value="<%=wKekkaSum%>"></td>

          <% if m_KekkaGaiDispFlg then %>
            <td class="header" align="center" nowrap><input type="text" name="KekkaGaiSum" size="4" <%=w_sInputClass%> readonly value="<%=wKekkaGaiSum%>"></td>
          <% end if %>

          <td class="header" align="center" colspan="2" nowrap>&nbsp;</td>
        </tr>

        <tr>
          <td class="header" align="right" colspan="7" nowrap>
            <FONT COLOR="#FFFFFF"><B>　平均点</B></FONT>
            <input type="text" name="txtAvg" size="5" <%=w_sInputClass%> readonly>
          </td>
          <td class="header" align="center" colspan="6" nowrap>&nbsp;</td>
        </tr>
      <% else %>
        <tr>
          <td class="header" align="center" colspan="7" nowrap><FONT COLOR="#FFFFFF"><B>合計</B></FONT></td>
          <td class="header" align="center" nowrap><input type="text" name="ChikaiSum" size="4" <%=w_sInputClass%> readonly value="<%=wChikokuSum%>"></td>
          <td class="header" align="center" nowrap>&nbsp;</td>
          <td class="header" align="center" nowrap><input type="text" name="KekkaSum" size="4" <%=w_sInputClass%> readonly value="<%=wKekkaSum%>"></td>

          <% if m_KekkaGaiDispFlg then %>
            <td class="header" align="center" nowrap><input type="text" name="KekkaGaiSum" size="4" <%=w_sInputClass%> readonly value="<%=wKekkaGaiSum%>"></td>
          <% end if %>

          <td class="header" align="center" colspan="2" nowrap>&nbsp;</td>
        </tr>
      <% end if %>

    </table>

    </td>
    </tr>

    <tr>
    <td align="center">
    <table>
      <tr>
        <td align="center" align="center" colspan="13">
          <%If m_bSeiInpFlg or m_bKekkaNyuryokuFlg Then%>
            <input type="button" class="button" value="　登　録　" onClick="f_Touroku();">
          <%End If%>
            <input type="button" class="button" value="キャンセル" onClick="f_Cancel();">
        </td>
      </tr>
    </table>
    </td>
    </tr>
  </table>

  <input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
  <input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
  <input type="hidden" name="KamokuCd"     value="<%=m_sKamokuCd%>">
  <input type="hidden" name="i_Max"        value="<%=i%>">
  <input type="hidden" name="sltShikenKbn" value="<%=m_sSikenKBN%>">
  <input type="hidden" name="txtGakuNo"    value="<%=m_sGakuNo%>">
  <input type="hidden" name="txtGakkaCd"   value="<%=m_sGakkaCd%>">
  <input type="hidden" name="txtClassNo"   value="<%=m_sClassNo%>">
  <input type="hidden" name="txtKamokuCd"  value="<%=m_sKamokuCd%>">
  <input type="hidden" name="PasteType"    value="">

  <input type="hidden" name="hidSouJyugyou">
  <input type="hidden" name="hidJunJyugyou">
  <input type="hidden" name="hidUpdMode">

 <!-- INS 2005/06/13 西村　福島高専用 -->
  <input type="hidden" name="hidSouJyugyou_TZ">
  <input type="hidden" name="hidJunJyugyou_TZ">
  <input type="hidden" name="hidSouJyugyou_KZ">
  <input type="hidden" name="hidJunJyugyou_KZ">
  <input type="hidden" name="hidSouJyugyou_TK">
  <input type="hidden" name="hidJunJyugyou_TK">
  <input type="hidden" name="hidSouJyugyou_KK">
  <input type="hidden" name="hidJunJyugyou_KK">



  <input type="hidden" name="hidKamokuKbn" value="<%=m_iKamokuKbn%>">
  <input type="hidden" name="hidKamokuBunrui" value="<%=m_sKamokuBunrui%>">
  <input type="hidden" name="hidSeisekiInpType" value="<%=m_iSeisekiInpType%>">

  <input type="hidden" name="hidKikan" value="<%=m_bSeiInpFlg%>">
  <input type="hidden" name="hidKekkaNyuryokuFlg" value="<%=m_bKekkaNyuryokuFlg%>">

  <input type="hidden" name="hidTotal" value="<%=w_lSeiTotal%>">
  <input type="hidden" name="hidGakTotal" value="<%=w_lGakTotal%>">
  <input type="hidden" name="txtUpdDate" value="<%=request("txtUpdDate")%>">

  <input type="hidden" name="hidZenkiOnly" value="<%=m_bZenkiOnly%>">

  <!--<input type="text" name="hidMihyoka" value ="<%=w_DataKbn%>">-->
  <input type="hidden" name="hidMihyoka" value ="<%=w_DataKbn%>">
  <input type="hidden" name="hidSchoolFlg" value ="<%=m_SchoolFlg%>">
  <input type="hidden" name="hidKekkaGaiDispFlg" value ="<%=m_KekkaGaiDispFlg%>">
  <input type="hidden" name="hidHyokaDispFlg" value ="<%=m_HyokaDispFlg%>">

  <input type="hidden" name="hidTableWidth" value ="<%=w_TableWidth%>">


  <input type="hidden" name="hidFromSei"   value ="<%=m_iNKaishi%>">
  <input type="hidden" name="hidToSei"     value ="<%=m_iNSyuryo%>">
  <input type="hidden" name="hidFromKekka" value ="<%=m_iKekkaKaishi%>">
  <input type="hidden" name="hidToKekka"   value ="<%=m_iKekkaSyuryo%>">

<%'2003.8.24 学校番号をsei0150_upd.aspに渡す為に追加。 ITO%>
  <input type="hidden" name="hidGakkoNo"   value ="<%=m_sGakkoNO%>">


  </form>
  </center>
  </body>
  </html>
<%
End sub
%>