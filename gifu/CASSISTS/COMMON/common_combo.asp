<%
'/******************************************************************************
' システム名：キャンパスアシスト
' 処　理　名：共通処理
' ﾌﾟﾛｸﾞﾗﾑID：-
' 機　　　能：コンボボックス共通化関数
'-------------------------------------------------------------------------------
' 作　　　成：2001.02.26 高丘　知央
' 変　　　更：2001.03.02 田部　雅幸     対応テーブル追加
' 　　　　　　2001.03.05 田部　雅幸     対応テーブル追加
' 　　　　　　2001.03.06 田部　雅幸     学年用関数追加
' 　　　　　　2001.03.26 岡田　誠　     教官室追加
' 　　　　　　2001.05.14 岡田　誠　     委員種別追加
'             2001.07.05 根本　直美　   複数名称表示追加
'             2001.07.12 岩下　幸一郎   対応テーブル追加
'******************************************************************************/

'ComboBoxにセットする時のレングスを固定長にする
Private Const mC_MST_NAMELEN = 50       ''マスタの名称データレングス
Private Const mC_MAX_IDNUMBER = 999999  ''最大のID番号（ログテーブル）

'ComboBox用null代用値
Public Const C_CBO_NULL = "@@@"       ''マスタの名称データレングス

'テーブルID用の固定値
Public Const C_CBO_M01_KUBUN = 1
Public Const C_CBO_M01_KUBUN_R = 101
Public Const C_CBO_M02_GAKKA = 2
Public Const C_CBO_M02_GAKKA_R = 102
Public Const C_CBO_M05_CLASS = 5
Public Const C_CBO_M05_CLASS_R = 105
Public Const C_CBO_M05_CLASS_N = 115
Public Const C_CBO_M05_CLASS_G = 125
Public Const C_CBO_M11_KYOKANSHITU = 11
Public Const C_CBO_M20_COURSE = 20
Public Const C_CBO_M12_SITYOSON = 12
Public Const C_CBO_M16_KEN = 16
Public Const C_CBO_M17_BUKATUDO = 17
Public Const C_CBO_M34_IIN = 34
Public Const C_CBO_M29_KYOKAN_YOTEI = 29
'Public Const C_CBO_M40_CALENDER = 40
'Public Const C_CBO_M40_CALENDER_H = 401
Public Const C_CBO_M27_SIKEN = 27
Public Const C_CBO_M04_KYOKAN = 04
Public Const C_CBO_T15_RISYU = 15
Public Const C_CBO_T15_RISYU_GRP = 150
Public Const C_CBO_T11_GAKUSEKI_N = 111
Public Const C_CBO_T16_R_KOJIN_S = 116
Public Const C_CBO_M06_KYOSITU = 6
Public Const C_CBO_T18_SEL_SYUBETU = 18
Public Const C_CBO_M42_JIKAN = 42
Public Const C_CBO_T26_SIKEN_JIK_KAMOKU = 126

Public Const C_CBO_T32_GYOJI_M = 32

Public Function gf_ComboSet(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' 機    能:ComboBoxセット
' 返    値:OK=True/NG=False
' 引    数:p_oCombo - ComboBox
'          p_sTableName - テーブル名
'          p_sWhere - Where条件(WHERE句は要らない)
'          p_sSelectOption - <SELECT>タグにつけるオプション( onchange = 'a_change()' )など
'          p_bWhite - 先頭に空白をつけるか
'          p_sSelectCD - 標準選択させたいコード(""なら選択なし)
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTしてComboBoxにセットする
' 備    考:(↓使用例）
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'上記の出力HTML（空白行のVALUEは’＠＠＠’となります。
'<select name='cboTest' ><Option Value='@@@'>　　　 <Option Value='1'>休学(病気・怪我) <Option Value='2'>休学(経済的他) 
'<Option Value='3'>復学 <Option Value='4'>停学 </select>
'*************************************************************************************
    Dim w_sId           'IDフィールド名
    Dim w_sName         '名称フィールド名
    Dim w_sTableName    '名称テーブル名
    Dim w_rst

    gf_ComboSet = False
    do 
    ''マスタ毎にSELECTするフィールド名を取得
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If
	
    ''マスタSELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If

'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '空白のOptionの代入
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">　　　　　 "& Chr(13)
    End If

    ''EOFでなければ、データをセット
    If Not w_rst.EOF Then
        Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If
    Response.write("</select>" & chr(13))

    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
   

    gf_ComboSet = True
    Exit Do
    Loop
End Function

Public Function gf_CalComboSet(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD,p_sDisp)
'*************************************************************************************
' 機    能:ComboBoxセット(日付Ver)
' 返    値:OK=True/NG=False
' 引    数:p_oCombo - ComboBox
'          p_sTableName - テーブル名
'          p_sWhere - Where条件(WHERE句は要らない)
'          p_sSelectOption - <SELECT>タグにつけるオプション( onchange = 'a_change()' )など
'          p_bWhite - 先頭に空白をつけるか
'          p_sSelectCD - 標準選択させたいコード(""なら選択なし)
'          p_sDISP - 表示方法(""ならyyyy/mm/ddで表記)
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTしてComboBoxにセットする
' 備    考:(↓使用例）
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'上記の出力HTML（空白行のVALUEは’＠＠＠’となります。
'<select name='cboTest' ><Option Value='@@@'>　　　 <Option Value='1'>休学(病気・怪我) <Option Value='2'>休学(経済的他) 
'<Option Value='3'>復学 <Option Value='4'>停学 </select>
'*************************************************************************************
    Dim w_sId           'IDフィールド名
    Dim w_sName         '名称フィールド名
    Dim w_sTableName    '名称テーブル名
    Dim w_rst

    gf_CalComboSet = False
    if p_sDisp = "" then p_sDisp = "yyyy/mm/dd"
    do 
    ''マスタ毎にSELECTするフィールド名を取得
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If
        
    ''マスタSELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '空白のOptionの代入
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">　　　 "& Chr(13)
    End If

    ''EOFでなければ、データをセット
    If Not w_rst.EOF Then
        Call s_CalMstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD,p_sDisp)
    End If
    Response.write("</select>" & chr(13))
    
    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
    
    
    gf_CalComboSet = True
    Exit Do
    Loop
End Function

Public Function gf_PluComboSet(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' 機    能:ComboBoxセット
' 返    値:OK=True/NG=False
' 引    数:p_oCombo - ComboBox
'          p_sTableName - テーブル名
'          p_sWhere - Where条件(WHERE句は要らない)
'          p_sSelectOption - <SELECT>タグにつけるオプション( onchange = 'a_change()' )など
'          p_bWhite - 先頭に空白をつけるか
'          p_sSelectCD - 標準選択させたいコード(""なら選択なし)
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTしてComboBoxにセットする
' 備    考:(↓使用例）
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'上記の出力HTML（空白行のVALUEは’＠＠＠’となります。
'<select name='cboTest' ><Option Value='@@@'>　　　 <Option Value='1'>休学(病気・怪我) <Option Value='2'>休学(経済的他) 
'<Option Value='3'>復学 <Option Value='4'>停学 </select>
'*************************************************************************************
    Dim w_sId           'IDフィールド名
    Dim w_sName         '名称フィールド名
    Dim w_sTableName    '名称テーブル名
    Dim w_rst

    gf_PluComboSet = False
    do 
    ''マスタ毎にSELECTするフィールド名を取得
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        response.write "ERR1"
        Exit Do
    End If

    ''マスタSELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        response.write "ERR2"
        Exit Do
    End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '空白のOptionの代入
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">　　　 "& Chr(13)
    End If

    ''EOFでなければ、データをセット
    If Not w_rst.EOF Then
        Call s_PluMstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If
    Response.write("</select>" & chr(13))

    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
   

    gf_PluComboSet = True
    Exit Do
    Loop
End Function


Public Function gf_ComboSetHead(p_sCombo, p_sSelectOption)
'*************************************************************************************
' 機    能:ComboBoxセット(ヘッダー)
' 返    値:OK=True/NG=False
' 引    数:p_oCombo - ComboBox
'          p_sSelectOption - <SELECT>タグにつけるオプション( onchange = 'a_change()' )など
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTしてComboBoxにセットする
' 備    考:
'*************************************************************************************
    Dim w_sId           'IDフィールド名
    Dim w_sName         '名称フィールド名
    Dim w_sTableName    '名称テーブル名
    Dim w_rst
    
    gf_ComboSetHead = False
    
    do 
        Response.write(chr(13)&"<select name='" & p_sCombo & "' " & p_sSelectOption & ">" & Chr(13))
        gf_ComboSetHead = True
    Exit Do
    Loop
End Function



Public Function gf_ComboSetData(p_iTableID, p_sWhere ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' 機    能:ComboBoxセット(データ)
' 返    値:OK=True/NG=False
' 引    数:p_oCombo - ComboBox
'          p_sTableName - テーブル名
'          p_sWhere - Where条件(WHERE句は要らない)
'          p_sSelectOption - <SELECT>タグにつけるオプション( onchange = 'a_change()' )など
'          p_bWhite - 先頭に空白をつけるか
'          p_sSelectCD - 標準選択させたいコード(""なら選択なし)
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTしてComboBoxにセットする
' 備    考:(↓使用例）
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'上記の出力HTML（空白行のVALUEは’＠＠＠’となります。
'<select name='cboTest' ><Option Value='@@@'>　　　 <Option Value='1'>休学(病気・怪我) <Option Value='2'>休学(経済的他) 
'<Option Value='3'>復学 <Option Value='4'>停学 </select>
'*************************************************************************************
    Dim w_sId           'IDフィールド名
    Dim w_sName         '名称フィールド名
    Dim w_sTableName    '名称テーブル名
    Dim w_rst
    
    gf_ComboSetData = False
    
    do 
    ''マスタ毎にSELECTするフィールド名を取得
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If
        
    ''マスタSELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------

    '空白のOptionの代入
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">　　　" & Chr(13)
    End If

    ''EOFでなければ、データをセット
    If Not w_rst.EOF Then
        Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If
    
    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
    
    
    gf_ComboSetData = True
    Exit Do
    Loop
End Function

Public Function gf_ComboSetFoot()
'*************************************************************************************
' 機    能:ComboBoxセット(フッダー)
' 返    値:OK=True/NG=False
' 引    数:
' 機能詳細:
' 備    考:
'*************************************************************************************
    
    gf_ComboSetFoot = False
    
    do 
        Response.write("</select>"&Chr(13))
        gf_ComboSetFoot = True
    Exit Do
    Loop
End Function










Private Function f_MstFieldName(p_iTableID, p_sID , p_sName , p_sTableName)
'*************************************************************************************
' 機    能: マスタフィールド名セット
' 返    値: OK=True/NG=False
' 引    数:  p_iTableID
'            p_sId - IDフィールド名(O)
'            p_sName - 名称フィールド名(O)
'            p_sTableName - テーブル名(I)
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTする（Comboセット用）
' 備    考:
'*************************************************************************************

    f_MstFieldName = True
    
    ''テーブル毎にIDと名称のフィールド名をセット
    Select Case p_iTableID
    Case C_CBO_M01_KUBUN            '区分マスタ(名称は普通の)
        p_sTableName = "M01_Kubun"
        p_sID = "M01_SYOBUNRUI_CD"
        p_sName = "M01_SYOBUNRUIMEI"
    Case C_CBO_M01_KUBUN_R          '区分マスタ(名称は略称)
        p_sTableName = "M01_Kubun"
        p_sID = "M01_SYOBUNRUI_CD"
        p_sName = "M01_SYOBUNRUIMEI_R"
    Case C_CBO_M02_GAKKA            '学科マスタ
        p_sTableName = "M02_GAKKA"
        p_sID = "M02_GAKKA_CD"
        p_sName = "M02_GAKKAMEI"
    Case C_CBO_M02_GAKKA_R          '学科マスタ略称
        p_sTableName = "M02_GAKKA"
        p_sID = "M02_GAKKA_CD"
        p_sName = "M02_GAKKARYAKSYO"
    Case C_CBO_M05_CLASS                'クラスマスタ
        p_sTableName = "M05_CLASS"
        p_sID = "M05_CLASSNO"
        p_sName = "M05_CLASSMEI"
    Case C_CBO_M05_CLASS_R          'クラスマスタ(略称)
        p_sTableName = "M05_CLASS"
        p_sID = "M05_CLASSNO"
        p_sName = "M05_CLASSRYAKU"
    Case C_CBO_M05_CLASS_N          'クラスマスタ(処理年度)
        p_sTableName = "M05_CLASS"
        p_sID = "M05_NENDO"
        p_sName = "M05_NENDO"
    Case C_CBO_M05_CLASS_G          'クラスマスタ(学年)
        p_sTableName = "M05_CLASS"
        p_sID = "M05_GAKUNEN"
        p_sName = "M05_GAKUNEN"
    Case C_CBO_M20_COURSE           'コースマスタ
        p_sTableName = "M20_COURSE"
        p_sID = "M20_COURSE_CD"
        p_sName = "M20_COURSEMEI"
    Case C_CBO_M11_KYOKANSHITU         '教官室マスタ
        p_sTableName = "M11_KYOKANSITU"
        p_sID = "M11_KYOKANSITU_CD"
        p_sName = "M11_KYOKANSITUMEI"
    Case C_CBO_M12_SITYOSON         '市町村マスタ
        p_sTableName = "M12_SITYOSON"
        p_sID = "M12_SITYOSON_CD"
        p_sName = "M12_SITYOSONMEI"
    Case C_CBO_M16_KEN              '県マスタ
        p_sTableName = "M16_KEN"
        p_sID = "M16_KEN_CD"
        p_sName = "M16_KENMEI"
    Case C_CBO_M17_BUKATUDO              '部活動マスタ
        p_sTableName = "M17_BUKATUDO"
        p_sID = "M17_BUKATUDO_CD"
        p_sName = "M17_BUKATUDOMEI"
    Case C_CBO_M34_IIN                  '委員マスタ
        p_sTableName = "M34_IIN"
        p_sID = "M34_DAIBUN_CD"
        p_sName = "M34_IIN_NAME"
    Case C_CBO_M29_KYOKAN_YOTEI                  '教官予定マスタ
        p_sTableName = "M29_KYOKAN_YOTEI"
        p_sID = "M29_YOTEI_CD"
        p_sName = "M29_YOTEIMEI"
'    Case C_CBO_M40_CALENDER                  'カレンダマスタ
'        p_sTableName = "M40_CALENDER"
'        p_sID = "M40_DATE"
'        p_sName = "M40_DATE"
'    Case C_CBO_M40_CALENDER_H                'カレンダマスタ(日を表示)
'        p_sTableName = "M40_CALENDER"
'        p_sID = "M40_DATE"
'        p_sName = "M40_HI"
    Case C_CBO_T32_GYOJI_M    '行事明細(カレンダー用、where句に"GROUP BY T32_HIDUKE"を入れる必要があります)
        p_sTableName = "T32_GYOJI_M"
        p_sID = "T32_HIDUKE"
        p_sName = "T32_HIDUKE"
    Case C_CBO_M27_SIKEN                  '試験マスタ
        p_sTableName = "M27_SIKEN"
        p_sID = "M27_SIKEN_CD"
        p_sName = "M27_SIKENMEI"
'    Case C_CBO_M04_KYOKAN                  '試験マスタ
'        p_sTableName = "M04_KYOKAN"
'        p_sID = "M04_KYOKAN_CD"
'        p_sName = "M04_KYOKANMEI_SEI"
    Case C_CBO_M04_KYOKAN                  '教官マスタ（複数名称表示）
        p_sTableName = "M04_KYOKAN"
        p_sID = "M04_KYOKAN_CD"
        p_sName = "M04_KYOKANMEI_SEI"
        p_sName = p_sName & " , M04_KYOKANMEI_MEI"
        'p_sName = p_sName & " , M04_KYOKANMEI_KANA_SEI"
        'p_sName = p_sName & " , M04_KYOKANMEI_KANA_MEI"
    Case C_CBO_T15_RISYU                  '履修テーブル
        p_sTableName = "T15_RISYU"
        p_sID = "T15_KAMOKU_CD"
        p_sName = "T15_KAMOKUMEI"
    Case C_CBO_T15_RISYU_GRP                  '履修テーブル
        p_sTableName = "T15_RISYU"
        p_sID = "T15_KAMOKU_CD"
        p_sName = "T15_KAMOKUMEI"
    Case C_CBO_T11_GAKUSEKI_N                  '学生名（複数名称表示）
        p_sTableName = "T11_GAKUSEKI T11"
        p_sTableName = p_sTableName & " , T13_GAKU_NEN T13"
        p_sID = "T11_GAKUSEI_NO"
        p_sName = "T13_GAKUSEKI_NO"
        p_sName = p_sName & " , T11_SIMEI"
    Case C_CBO_T16_R_KOJIN_S
        p_sTableName = "T16_RISYU_KOJIN"
        p_sID = "T16_GRP"
        p_sName = "T16_SYUBETU_MEI"
    Case C_CBO_M06_KYOSITU
        p_sTableName = "M06_KYOSITU"
        p_sID = "M06_KYOSITU_CD"
        p_sName = "M06_KYOSITUMEI"
    Case C_CBO_T18_SEL_SYUBETU
        p_sTableName = "T18_SELECTSYUBETU"
        p_sTableName = p_sTableName & " , T15_RISYU "
        p_sID = "T18_GRP"
        p_sName = "T18_SYUBETU_MEI"
    Case C_CBO_M42_JIKAN
        p_sTableName = "M42_SIKEN_PATTERN"
        p_sID = "M42_SIKEN_JIKAN"
        p_sName = "M42_SIKEN_JIKAN"
    Case C_CBO_T26_SIKEN_JIK_KAMOKU                  '試験時間割科目名（複数名称表示）
        p_sTableName = "T26_SIKEN_JIKANWARI"
        p_sTableName = p_sTableName & " , M03_KAMOKU "
        p_sID = "T26_KAMOKU"
        p_sName = "M03_KAMOKUMEI"
'以下足りないものはテーブル毎に追加する
'    Case C_CBO_M01_KUBUN            '区分マスタ(名称は普通の)
'        p_sTableName = "M01_Kubun"
'        p_sID = "M01_SYOBUNRUI_CD"
'        p_sName = "M01_SYOBUNRUIMEI"

    Case Else
        'TODO20010226T.TakaokaCall gs_criticalMsg(4101)        '"マスタ指定がありません"
        f_MstFieldName = False
    End Select
    
    

End Function


Private Function f_MstSelect(p_rst, p_sID , _
                             p_sName , p_sTableName , _
                             p_sWhere )
'*************************************************************************************
' 機    能:マスタSELECT
' 返    値:OK=True/NG=False
' 引    数:p_rst - Recordset
'            p_sId - IDフィールド名
'            p_sName - 名称フィールド名
'            p_sTableName - テーブル名
'            p_sWhere - WHERE条件
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTする（Comboセット用）
' 備    考:
'*************************************************************************************

    Dim w_sSQL
    
    f_MstSelect = False
    
    ''SQL作成
    w_sSQL = "SELECT "
    if p_sID <> p_sName Then
     w_sSQL = w_sSQL & vbCrLf & p_sID & ", "
    end if

    w_sSQL = w_sSQL & vbCrLf & p_sName & " "
    w_sSQL = w_sSQL & vbCrLf & "FROM " & p_sTableName
    If p_sWhere <> "" Then
        w_sSQL = w_sSQL & vbCrLf & " WHERE " & p_sWhere
    End If

'Response.write w_sSql

'------------20010810 tani
'強引ですが、「学籍番号　氏名」のコンボになるときは、
'学籍番号でソートします。
    If p_sName = "T13_GAKUSEKI_NO , T11_SIMEI" Then
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T13_GAKUSEKI_NO"
    Else
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY " & p_sID
    End If
'--------------
'戻すときは、上の分を消して、下のコメントをはずす。
'        w_sSQL = w_sSQL & vbCrLf & " ORDER BY " & p_sID



'----------------------------------------------------------
'response.write w_sSQL & "<BR>"	 '(デバッグ)
'response.write "<< デバッグ中・・・mochi >>"
'----------------------------------------------------------

    ''SELECT
    If gf_GetRecordset(p_rst, w_sSQL) <> 0 Then
        Exit Function

    End If

    f_MstSelect = True
    Exit Function

End Function


Private Sub s_MstDataSet(p_sCombo, p_rst, p_sID, p_sName,p_sSelectCD)
'*************************************************************************************
' 機    能:マスタデータセット
' 返    値:
' 引    数:p_oCombo - ComboBox
'            p_rst - RecordSet
'            p_sId - IDフィールド名
'            p_sName - 名称フィールド名
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTする（Comboセット用）
' 備    考:
'*************************************************************************************
    Dim w_sNameData
    Dim w_sData
    
    ''レコードセット=EOFまでループ
    Do Until p_rst.EOF
        '''Comboにセットするデータ編集
        'w_sNameData = gf_GetFieldValue(p_rst, p_sName)
        'w_sData = w_sNameData & Space(mC_MST_NAMELEN - len(w_sNameData))
        w_sData = gf_GetFieldValue(p_rst, p_sID)

        response.write(" <Option Value='" & w_sData & "'")
		if Not gf_IsNull(p_sSelectCD) then
	        If CStr(p_sSelectCD) = CStr(w_sData) Then
	            response.write " Selected "
	        End If
		End if
        response.Write(">" & gf_GetFieldValue(p_rst, p_sName) & Chr(13))
        p_rst.MoveNext      '''次レコードへ
        
    Loop

End Sub

Private Sub s_calMstDataSet(p_sCombo, p_rst, p_sID, p_sName,p_sSelectCD,p_sDisp)
'*************************************************************************************
' 機    能:カレンダマスタデータセット
' 返    値:
' 引    数:p_oCombo - ComboBox
'            p_rst - RecordSet
'            p_sId - IDフィールド名
'            p_sName - 名称フィールド名
'            p_sDisp - 表記方法
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTする（Comboセット用）
' 備    考:
'*************************************************************************************
    Dim w_sNameData
    Dim w_sData
    Dim w_sDisp
    Dim w_DE
    
    w_DE = array("","sun","mon","tue","wed","thu","fri","sat")
    ''レコードセット=EOFまでループ
    Do Until p_rst.EOF
        '''Comboにセットするデータ編集
        'w_sNameData = gf_GetFieldValue(p_rst, p_sName)
        'w_sData = w_sNameData & Space(mC_MST_NAMELEN - len(w_sNameData))
        w_sData = gf_GetFieldValue(p_rst, p_sID)
        
        response.write(" <Option Value='" & w_sData & "'")
        If CStr(p_sSelectCD) = CStr(w_sData) Then
            response.write " Selected "
        End If
    w_sDisp = p_sDisp
    w_sNameData = gf_GetFieldValue(p_rst, p_sName)
    w_sDisp = Replace(w_sDisp,"yyyy",year(w_sNameData))
    w_sDisp = Replace(w_sDisp,"yy",right(year(w_sNameData),2))
    w_sDisp = Replace(w_sDisp,"mm",Month(w_sNameData))
    w_sDisp = Replace(w_sDisp,"dd",day(w_sNameData))
    w_sDisp = Replace(w_sDisp,"DJ",WeekDayName(Weekday(w_sNameData),true))
    w_sDisp = Replace(w_sDisp,"DE",w_DE(Weekday(w_sNameData)))
        response.Write(">" & w_sDisp & Chr(13))
        p_rst.MoveNext      '''次レコードへ
        
    Loop

End Sub

Private Sub s_PluMstDataSet(p_sCombo, p_rst, p_sID, p_sName,p_sSelectCD)
'*************************************************************************************
' 機    能:マスタデータセット
' 返    値:
' 引    数:p_oCombo - ComboBox
'            p_rst - RecordSet
'            p_sId - IDフィールド名
'            p_sName - 名称フィールド名
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと複数名称をSELECTする（Comboセット用）
' 備    考:
'*************************************************************************************
    Dim w_sNameData
    Dim w_sData
    Dim w_sField
    
    ''レコードセット=EOFまでループ
    Do Until p_rst.EOF
        '''Comboにセットするデータ編集
        'w_sNameData = gf_GetFieldValue(p_rst, p_sName)
        'w_sData = w_sNameData & Space(mC_MST_NAMELEN - len(w_sNameData))
        w_sData = gf_GetFieldValue(p_rst, p_sID)
        
        w_sNameData = Split(p_sName,",")
        response.write(" <Option Value='" & w_sData & "'")
        If CStr(p_sSelectCD) = CStr(w_sData) Then
            response.write " Selected "
        End If
        'response.Write(">" & gf_GetFieldValue(p_rst, p_sName) & Chr(13))
        response.Write(">")
        
        For Each w_sField In w_sNameData
            response.Write(gf_GetFieldValue(p_rst, Trim(w_sField)))
            response.Write("　")
        Next
        
        response.Write(Chr(13))
        p_rst.MoveNext      '''次レコードへ
        
    Loop

End Sub




Public Function gf_ComboSet_99(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' 機    能:ComboBoxセット
' 返    値:OK=True/NG=False
' 引    数:p_oCombo - ComboBox
'          p_sTableName - テーブル名
'          p_sWhere - Where条件(WHERE句は要らない)
'          p_sSelectOption - <SELECT>タグにつけるオプション( onchange = 'a_change()' )など
'          p_bWhite - 先頭に空白をつけるか
'          p_sSelectCD - 標準選択させたいコード(""なら選択なし)
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTしてComboBoxにセットする
' 備    考:(↓使用例）
'Call gf_ComboSet_99("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'上記の出力HTML（空白行のVALUEは’＠＠＠’となります。
'<select name='cboTest' ><Option Value='@@@'>　　　 <Option Value='1'>休学(病気・怪我) <Option Value='2'>休学(経済的他) 
'<Option Value='3'>復学 <Option Value='4'>停学 </select>
'*************************************************************************************
    Dim w_sId           'IDフィールド名
    Dim w_sName         '名称フィールド名
    Dim w_sTableName    '名称テーブル名
    Dim w_rst

    gf_ComboSet_99 = False
    do 
    ''マスタ毎にSELECTするフィールド名を取得
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If

    ''マスタSELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If
    
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '空白のOptionの代入
    If p_bWhite Then
        response.Write " <Option Value="& 99 &">　　　 "& Chr(13)
    End If

    ''EOFでなければ、データをセット
    If Not w_rst.EOF Then
        Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If
    Response.write("</select>" & chr(13))

    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
   

    gf_ComboSet_99 = True
    Exit Do
    Loop
End Function



Public Function gf_cboNull(p_sStr)
'*************************************************************************************
' 機    能:ComboBox引数null代用値の変換
' 返    値:変換後の文字列
' 引    数:変換したい文字列
' 機能詳細:
' 備    考:
'*************************************************************************************
    
    gf_cboNull = p_sStr
    
    if p_sStr = C_CBO_NULL then gf_cboNull = ""

End Function

Public Function gf_ComboSet_Gakka(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' 機    能:ComboBoxセット
' 返    値:OK=True/NG=False
' 引    数:p_oCombo - ComboBox
'          p_sTableName - テーブル名
'          p_sWhere - Where条件(WHERE句は要らない)
'          p_sSelectOption - <SELECT>タグにつけるオプション( onchange = 'a_change()' )など
'          p_bWhite - 先頭に空白をつけるか
'          p_sSelectCD - 標準選択させたいコード(""なら選択なし)
' 機能詳細:指定されたテーブルから、ｺｰﾄﾞと名称をSELECTしてComboBoxにセットする
' 備    考:(↓使用例）
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'上記の出力HTML（空白行のVALUEは’＠＠＠’となります。
'<select name='cboTest' ><Option Value='@@@'>　　　 <Option Value='1'>休学(病気・怪我) <Option Value='2'>休学(経済的他) 
'<Option Value='3'>復学 <Option Value='4'>停学 </select>
'*************************************************************************************
    Dim w_sId           'IDフィールド名
    Dim w_sName         '名称フィールド名
    Dim w_sTableName    '名称テーブル名
    Dim w_rst

    gf_ComboSet_Gakka = False
    do 
    ''マスタ毎にSELECTするフィールド名を取得
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If

    ''マスタSELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '空白のOptionの代入
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">　　　　　 "& Chr(13)
    End If

    ''EOFでなければ、データをセット
    If Not w_rst.EOF Then
        Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If

    response.write(" <Option Value='" & C_CLASS_ALL & "'")
    If CStr(p_sSelectCD) = CStr(C_CLASS_ALL) Then
        response.write " Selected "
    End If
    response.Write(">" & "全学科" & Chr(13))

    Response.write("</select>" & chr(13))

    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
   
    gf_ComboSet_Gakka = True
    Exit Do
    Loop
End Function

%>