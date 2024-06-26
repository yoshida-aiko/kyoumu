<%
'/************************************************************************
' システム名: キャンパスアシストシステム
' 処  理  名: 共通処理−キャンパスアシスト
' ﾌﾟﾛｸﾞﾗﾑID : CACommon.asp
' 機      能: このファイルにはキャンパスアシスト固有の関数、定義をしてください。
'-------------------------------------------------------------------------
' 作      成: 2001.03.15 高丘 知央
' 変      更: 2001.07.12 谷脇 良也
' 　      　: 2001.07.18 根本 直美  '//最大時限数表示用関数追加
' 　      　: 2001.07.22 谷脇 良也  '//権限関係関数追加
' 　      　: 2001.12.01 田部 雅幸　'//教科書登録にFullとNormalの違いが無かったのを修正
' 　      　: 2002.04.26 shin	  　'//異動名称取得関数(gf_Set_Idou)の修正
'           : 2017/03/01 藤 美由紀  WEBアクセスログカスタマイズ
'*************************************************************************/


'//////////////////////////////////////////////////////////////////////////////////////////
'
'	関数一覧
'
'//////////////////////////////////////////////////////////////////////////////////////////
'交互にセルに色をつける				gs_cellPtn(p_sCell)
'nullを0に変換−数字ver.			gf_nInt(p_str)
'nullを""に変換−文字列ver.			gf_nStr(p_str)
'nullを指定文字に変換				gf_Null(p_str,p_henkan)
'四捨五入							gf_Round(p_num, p_keta)
'タイトルを出すサブルーチン			gs_title(p_title,p_subtitle)
'ページ関係の表示用サブルーチン		gs_pageBar(p_Rs,p_sPageCD,p_iDsp,p_pageBar)
'出欠データの取得					gf_GetSyukketuData(p_oRecordset, p_sSikenKbn, p_sGakunen, _p_sTantoKyokan, p_sClass, p_sKamokuCD, _p_sKaisibi, p_sSyuryobi, p_s1NenBango)
'出欠を取得する開始日と終了日の取得	gf_GetKaisiSyuryo(p_sSikenKbn, p_sGakunen, p_sKaisibi, p_sSyuryobi)
'引数を表示（デバッグ用）			gs_viewForm(p_form)
'最大時限数を取得					gf_GetJigenMax(p_iJMax)
'学籍項目レベル別表示				gf_empItem(p_ItemNo)
'メニュー項目レベル別表示			gf_empMenu(p_iMenuID)
'不正アクセスチェック				gf_userChk(p_PRJ_No)
'前期・後期情報を取得				gf_GetGakkiInfo(p_sGakki,p_sZenki_Start,p_sKouki_Start,p_sKouki_End)
'曜日ｺｰﾄﾞから曜日略称を返す			gf_GetYoubi(p_CD)
'LOGINした人が担任かどうかの判断	gf_Tannin(p_Nendo,p_Kyokan)				'8/11 前田 追加
'現在の日付に一番近い試験区分を取得	gf_Get_SikenKbn(p_iSiken_Kbn,p_kikan,p_gakunen)	'8/17 谷脇 追加
'教官名称を取得する					gf_GetKyokanNm(p_iNendo,p_sCD)
'学科名称を取得する					gf_GetGakkaNm(p_iNendo,p_sCD)
'クラス名称を取得する				gf_GetClassName(p_iNendo,p_iGakuNen,p_ClassNo)
'1年間番号,5年間番号の名称を取得    gf_GetGakuNomei(p_iNendo,p_iGakuKBN)
'USER名称を取得する					gf_GetUserNm(p_iNendo,p_sID)
'特別教室予約の権限を取得           gf_GetKengen_web0300(p_sKengen)
'使用教科書登録処理の権限を取得     gf_GetKengen_web0320(p_sKengen)
'個人履修選択科目決定の権限を取得   gf_GetKengen_web0340(p_sKengen)
'レベル別科目決定の権限を取得  		gf_GetKengen_web0390(p_sKengen)
'確定欠課数、遅刻数を取得           gf_GetKekaChi(p_iNendo,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku,p_iKekkaGai)
'管理マスタより出欠欠課種別を取得	gf_GetKanriInfo(p_iNendo,p_iSyubetu)
'出欠の入力ができなくなる日を取得	gf_Get_SyuketuEnd(p_iGakunen,p_sEndDay)
'郵便番号からの住所検索				gf_ComvZip(p_sKenMei,p_sSityosonCD,p_sSityosonMei,p_sZipCD,p_sTyoikiMei,p_iNendo)　Add 2001.12.5.岡田
'異動状況チェック関数				gf_Get_IdouChk(p_Gakusei_No,p_Date,p_iNendo) Add 2001.12.18 岡田
'異動名称取得関数					gf_Set_Idou(p_sGakusekiCd,p_iNendo,ByRef p_SSSS)

'出欠を取得する開始日と終了日、
'試験実績の登録をした試験区分の取得 gf_GetStartEnd(p_SyoriNen,p_sSikenKbn, p_sGakunen,p_GakusekiNo,p_Kamoku,p_sKaisibi, p_sSyuryobi,p_ShikenInsertKbn)
'↑Add 2002/06/13 shin
'試験実績登録の更新日取得			gf_GetUpdateDate(p_Nendo,p_KamokuCd,p_GakusekiNo,p_ShikenKbn,p_UpdateDate)
'↑Add 2002/06/13 shin
'試験実施終了日を取得する			gf_GetShikenDate(p_iNendo,p_sGakunen,p_ShikenKbn,p_UpdateDate)
'↑Add 2002/06/13 shin
'科目名を取得						gf_GetKamokuMei(p_SyoriNen,p_KamokuCd,p_KamokuKbn)
'操作ログを登録						gf_InsertOpeLog(p_nendo,p_syoriId,p_syoriName,p_taisyo,p_sosa,p_userId)		'add 2017/03/01 m.tou



'** 構造定義 **

'** 変数宣言 ** 

'** 外部ﾌﾟﾛｼｰｼﾞｬ定義 **

'////////////////////////////////////////////////////////////////////////
'// 交互にセルに色をつける
'//
'// 引　数：
'// 戻り値：セルのクラス名
'// 
'////////////////////////////////////////////////////////////////////////
sub gs_cellPtn(p_sCell) 

    if p_sCell = "" then p_sCell = C_CELL2

    if p_sCell = C_CELL1 then 
        p_sCell = C_CELL2
    else 
        p_sCell = C_CELL1
    end if

End sub

'////////////////////////////////////////////////////////////////////////
'// nullを0に変換−数字ver.
'//
'// 引　数：nullチェックするもの
'// 戻り値：nullを0に置換したもの
'// 
'////////////////////////////////////////////////////////////////////////
function gf_nInt(p_nstr)
    gf_nInt=gf_Null(p_nstr,"0")
end function

'////////////////////////////////////////////////////////////////////////
'// nullを""に変換−文字列ver.
'//
'// 引　数：nullチェックするもの
'// 戻り値：nullを""に置換したもの
'// 
'////////////////////////////////////////////////////////////////////////
function gf_nStr(p_str)
'    gf_nStr = gf_Null(p_str,"")

	If gf_Null(p_str,"") = "" Then
		gf_nStr =  ""
	Else
		gf_nStr = cstr(p_str)
	End If

end function

'////////////////////////////////////////////////////////////////////////
'// nullを指定文字に変換
'//
'// 引　数：nullチェックするもの，置換したいもの
'// 戻り値：nullを指定の物に置換したもの
'// 
'////////////////////////////////////////////////////////////////////////
Function  gf_Null(p_str,p_henkan) 
    if isnull(p_str) then 
        gf_isNull = p_henkan
    else
        gf_isNull=p_str
    end if
end function

'////////////////////////////////////////////////////////////////////////
'// 四捨五入
'//
'// 引　数：四捨五入したい値
'// 　　　：四捨五入対象桁
'// 戻り値：四捨五入した値
'// 
'////////////////////////////////////////////////////////////////////////
Function gf_Round(p_num, p_keta)
    Dim k
    Dim x
    If p_keta >= 0 Then
        k = CLng(10 ^ p_keta)
        x = Int(p_num * k + 0.5) / k
        gf_Round = x
    Else
        k = CLng(10 ^ (-p_keta))
        x = Int(p_num / k + 0.5) * k
        gf_Round = x
    End If
End Function

'////////////////////////////////////////////////////////////////////////
'// タイトルを出すサブルーチン
'//
'// 引　数：タイトルとサブタイトル
'// 戻り値：なし
'// 
'////////////////////////////////////////////////////////////////////////
sub gs_title(p_title,p_subtitle)

%>
    <table cellspacing="0" cellpadding="0" border="0" width="98%">
    <tr>
    <td height="27" width="100%" align="left"
    >

        <DIV class=title><%=p_title%></DIV>

    </td
    >
    </tr
    >

    <tr
    ><td height="4" width="5%" background="<%=C_IMAGE_DIR%>table_sita.gif"
    ><img src="<%=C_IMAGE_DIR%>sp.gif"
    ></td
    ></tr
    >

    <tr
    ><td height="10" class=title_Sub width="5%" align="right" valign="top"
    >

        <table class=title_Sub cellspacing="0" cellpadding="0" bgcolor=#393976 height="10" border="0"
        ><tr
        ><td align="center" valign="middle"
        ><DIV class=title_Sub
	><img src="<%=C_IMAGE_DIR%>sp.gif" width=8
        ><font color="#ffffff"
	><%=p_subtitle%></font
	><img src="<%=C_IMAGE_DIR%>sp.gif" width=8
        ></DIV
        ></td
        ></tr
        ></table
        >
    </td
    ></tr
    ></table>
<%

end sub


'********************************************************************************
'*  [機能]  ページ関係の表示用サブルーチン
'*  [引数]  p_Rs            ：一覧を表示するレコードセット
'*  　　　 p_sPageCD        ：ページ番号
'* 　　　  p_iDsp           ：1ページの最大表示すう。
'*  　　　 p_pageBar        ：
'*  [戻値]  p_pageBar       ：できたページバーHTML
'*  [説明]  
'********************************************************************************
sub gs_pageBar(p_Rs,p_sPageCD,p_iDsp,p_pageBar)
    Dim w_bNxt              '// NEXT表示有無
    Dim w_bBfr              '// BEFORE表示有無
    Dim w_iNxt              '// NEXT表示頁数
    Dim w_iBfr              '// BEFORE表示頁数
    Dim w_iCnt              '// ﾃﾞｰﾀ表示ｶｳﾝﾀ
    Dim w_iMax              '// ﾃﾞｰﾀ表示ｶｳﾝﾀ
    Dim i,w_iSt,w_iEd

    Dim w_iRecordCnt        '//レコードセットカウント

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

    '////////////////////////////////////////
    '      ページ関係の設定
    '////////////////////////////////////////
    'レコード数を取得
    w_iRecordCnt = gf_GetRsCount(p_Rs)
    w_iMax = gf_PageCount(p_Rs,p_iDsp)

    'EOFのときの設定
    If  p_sPageCD >= w_iMax Then
        p_sPageCD = w_iMax
    End If

    '前ページの設定
    If INT(p_sPageCD)=1 Then
        w_bBfr=False
        w_iBfr=0
    Else
        w_bBfr=True
        w_iBfr=p_sPageCD-1
    End If

    '後ページの設定
    If p_sPageCD=w_iMax Then
        w_bNxt=False
        w_iNxt=p_sPageCD
    Else
        w_bNxt=True
        w_iNxt=p_sPageCD+1
    End If
    
	'ページのリストの始め(w_iSt)と終わり(w_iEd)を代入
	'基本的に選択されているページ(p_sPageCD)が真中に来るようにする。
    w_iEd = p_sPageCD + 5
    w_iSt = p_sPageCD - 4
    
	'ページのリストが10個ない時、選択ページがリストの真中にこないとき。
    If p_sPageCD < 5 Then w_iEd = 10
    If w_iEd > w_iMax then w_iEd = w_iMax:w_iSt = w_iMax - 9
    If w_iSt < 1 or w_iMax < 10 then w_iSt = 1
    
    '絶対値ページの設定
    call gs_AbsolutePage(p_Rs,p_sPageCD,p_iDsp)
    
'////////////////////////////////////////
'      ページ関係の設定(ここまで)
'////////////////////////////////////////
    
        p_pageBar = ""
        p_pageBar = p_pageBar & vbCrLf & "<table border='0' width='100%'>"
        p_pageBar = p_pageBar & vbCrLf & "<tr>"
        p_pageBar = p_pageBar & vbCrLf & "<td align='left' width='10%'>"
    If w_bBfr = True Then
        p_pageBar = p_pageBar & vbCrLf & "<a href='javascript:f_PageClick("& w_iBfr &");' class='page'>前へ</a>"
    End If
        p_pageBar = p_pageBar & vbCrLf & " </td>"
        p_pageBar = p_pageBar & vbCrLf & "<td align=center width='80%'>"
        p_pageBar = p_pageBar & vbCrLf & " Page：[ "
    for i = w_iSt to w_iEd
'   for i = 1 to w_iMax
        If i = p_sPageCD then 
            p_pageBar = p_pageBar & vbCrLf & "<span class='page'>" & i & "</span>"
        Else
            p_pageBar = p_pageBar & vbCrLf & "<a href='javascript:f_PageClick("& i &");' class='page'>" & i & "</a>"
        End If
    next
        p_pageBar = p_pageBar & vbCrLf & "/" & w_iMax & "] "
        p_pageBar = p_pageBar & vbCrLf & " Results：" & w_iRecordCnt & "Hits"
        p_pageBar = p_pageBar & vbCrLf & "</td>"
        p_pageBar = p_pageBar & vbCrLf & "<td align='right' width='10%'> "
    If w_bNxt = True Then
        p_pageBar = p_pageBar & vbCrLf & "<a href='javascript:f_PageClick(" & w_iNxt & ")' class='page'>次へ</a>"
    End If
        p_pageBar = p_pageBar & vbCrLf & "</td>"
        p_pageBar = p_pageBar & vbCrLf & "</tr>"
        p_pageBar = p_pageBar & vbCrLf & "</table>"
end sub

'*******************************************************************************
' 機　　能：出欠データの取得
' 返　　値：取得結果
' 　　　　　(True)成功, (False)失敗
' 引　　数：p_oRecordset - レコードセット
' 　　　　　p_sSikenKbn - 試験区分
' 　　　　　p_sGakunen - 学年
' 　　　　　p_sTantoKyokan - 教官ＣＤ
' 　　　　　p_sClass - クラスNo
' 　　　　　p_sKamokuCD - 科目コード
' 　　　　　p_sKaisibi - 開始日
' 　　　　　p_sSyuryobi - 終了日
' 　　　　　p_s1NenBango - １年間番号
' 機能詳細：指定された条件の出欠のデータを取得する
' 備　　考：なし
'*******************************************************************************
Function gf_GetSyukketuData(p_oRecordset,p_sSikenKbn,p_sGakunen,p_sTantoKyokan,p_sClass,p_sKamokuCD,p_sKaisibi,p_sSyuryobi,p_s1NenBango)
	
	Dim w_sSql			'SQL
	
	On Error Resume Next
	
	'== 初期化 ==
	gf_GetSyukketuData = False
	
	'== 出欠を取得する開始日と終了日を取得する ==
	'//(試験間の期間)
	If gf_GetKaisiSyuryo(p_sSikenKbn, p_sGakunen, p_sKaisibi, p_sSyuryobi) <> True Then
		Exit Function
	End If
	
	'== 出欠を取得する ==
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & "SELECT "
	w_sSql = w_sSql & vbCrLf & "	Count(T21_GAKUSEKI_NO) as KAISU,"
	w_sSql = w_sSql & vbCrLf & "	T21_CLASS,"
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & "	T21_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "FROM "
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU "
	w_sSql = w_sSql & vbCrLf & "Where "
	w_sSql = w_sSql & vbCrLf & "	T21_NENDO = " & session("NENDO") & " "		'年度
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_GAKUNEN = " & p_sGakunen & " "			'学年
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_KAMOKU = '" & p_sKamokuCD & "' " 		'科目
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_KYOKAN = '" & p_sTantoKyokan & "' "		'教官
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_HIDUKE >= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sKaisibi & "' "						'開始日
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_HIDUKE <= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sSyuryobi & "' "						'終了日
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU_KBN IN ('" & C_KETU_KEKKA & "','" & C_KETU_TIKOKU & "','"& C_KETU_SOTAI &"','" & C_KETU_KEKKA_1 & "')"
	
	'== １年間番号が指定されている場合 ==
	If p_s1NenBango <> "" Then
		w_sSql = w_sSql & vbCrLf & "And "
		w_sSql = w_sSql & vbCrLf & "T21_GAKUSEKI_NO = " & p_s1NenBango & " "			'クラス
	End If
	
	w_sSql = w_sSql & vbCrLf & "Group By "
	w_sSql = w_sSql & vbCrLf & " T21_CLASS,"
	w_sSql = w_sSql & vbCrLf & " T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & " T21_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "Order By "
	w_sSql = w_sSql & vbCrLf & " T21_CLASS, "
	w_sSql = w_sSql & vbCrLf & " T21_GAKUSEKI_NO "
	
	'== データの取得 ==
	Set p_oRecordset = Server.CreateObject("ADODB.Recordset")
	
	'== 失敗したとき ==
	If gf_GetRecordset(p_oRecordset, w_sSql) <> 0 Then
		p_oRecordset.Close
		Set p_oRecordset = Nothing
		Exit Function
	End If
	
	gf_GetSyukketuData = True
	
End Function

Function gf_GetKaisiSyuryo(p_sSikenKbn, p_sGakunen, p_sKaisibi, p_sSyuryobi)
'*******************************************************************************
' 機　　能：出欠を取得する開始日と終了日の取得
' 返　　値：取得結果
' 　　　　　(True)成功, (False)失敗
' 引　　数：p_sSikenKbn - 試験区分
' 　　　　　p_sGakunen - 学年
' 　　　　　p_sKaisibi - 開始日
' 　　　　　p_sSyuryobi - 終了日
' 機能詳細：出欠を取得する開始日と終了日の取得
' 備　　考：なし
'*******************************************************************************
	Dim w_bRtn 						'戻り値
	Dim w_sSql
	Dim w_iNendo
	
	Dim w_oRecordset				'レコードセット
	
	w_iNendo = session("NENDO")

	On Error Resume Next
	
	'== 初期化 ==
	gf_GetKaisiSyuryo = False
	w_bRtn = False

	'== 試験によって取得するデータを変更する ==
	Select Case p_sSikenKbn
		Case C_SIKEN_ZEN_TYU		'前期中間
			'== SQL作成 ==
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "M00_KANRI "
			w_sSql = w_sSql & vbCrLf & "From M00_KANRI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "M00_NENDO = " & w_iNendo & " "	'年度
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "M00_NO = " & C_K_ZEN_KAISI & " " 				'前期開始日
			w_sSql = w_sSql & vbCrLf & "Union "
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_KAISI "
			w_sSql = w_sSql & vbCrLf & "FROM T24_SIKEN_NITTEI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "T24_NENDO = " & w_iNendo & " "	'年度
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN = " & C_SIKEN_ZEN_TYU & " "		'試験区分
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_GAKUNEN = " & p_sGakunen & " "				'学年
			w_sSql = w_sSql & vbCrLf & "Order By "
			w_sSql = w_sSql & vbCrLf & "1"
		Case C_SIKEN_ZEN_KIM		'前期期末

			'== SQL作成 ==
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_KAISI, "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_SYURYO "
			w_sSql = w_sSql & vbCrLf & "FROM T24_SIKEN_NITTEI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "T24_NENDO = " & w_iNendo & " "	'年度
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN IN (" & C_SIKEN_ZEN_TYU & " "		'試験区分
			w_sSql = w_sSql & vbCrLf & ", " & C_SIKEN_ZEN_KIM & " "		'試験区分
			w_sSql = w_sSql & vbCrLf & ") "
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_GAKUNEN = " & p_sGakunen & " "				'学年
			w_sSql = w_sSql & vbCrLf & "Order By "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN"

		Case C_SIKEN_KOU_TYU		'後期中間
			'== SQL作成 ==
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "M00_KANRI "
			w_sSql = w_sSql & vbCrLf & "From M00_KANRI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "M00_NENDO = " & w_iNendo & " "	'年度
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "M00_NO = " & C_K_KOU_KAISI & " " 				'後期開始日
			w_sSql = w_sSql & vbCrLf & "Union "
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_KAISI "
			w_sSql = w_sSql & vbCrLf & "FROM T24_SIKEN_NITTEI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "T24_NENDO = " & w_iNendo & " "	'年度
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN = " & C_SIKEN_KOU_TYU & " "		'試験区分
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_GAKUNEN = " & p_sGakunen & " "				'学年
			w_sSql = w_sSql & vbCrLf & "Order By "
			w_sSql = w_sSql & vbCrLf & "1"
			
		Case C_SIKEN_KOU_KIM		'後期期末
			'== SQL作成 ==
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_KAISI, "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_SYURYO "
			w_sSql = w_sSql & vbCrLf & "FROM T24_SIKEN_NITTEI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "T24_NENDO = " & w_iNendo & " "	'年度
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & " T24_SIKEN_KBN IN (" & C_SIKEN_KOU_TYU & " "		'試験区分
			w_sSql = w_sSql & vbCrLf & ", " & C_SIKEN_KOU_KIM & " "		'試験区分
			w_sSql = w_sSql & vbCrLf & ")"
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_GAKUNEN = " & p_sGakunen & " "				'学年
			w_sSql = w_sSql & vbCrLf & "Order By "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN"
	End Select

	'== データの取得 ==
	Set w_oRecordset = Server.CreateObject("ADODB.Recordset")
	
	'== 失敗したとき ==.
	    If gf_GetRecordset(w_oRecordset, w_sSql) <> 0 Then
		w_oRecordset.Close
		Set w_oRecordset = Nothing
		
		Exit Function
	End If


	'== ２件取れなかった場合 ==
	'If gf_GetRsCount(w_oRecordset) < 2 Then
	'	w_oRecordset.Close
	'	Set w_oRecordset = Nothing
	'	
	'	Exit Function
	'End If
	
	'== 開始日と終了日の設定 ==
	w_oRecordset.MoveFirst
		Select Case p_sSikenKbn
		Case C_SIKEN_ZEN_TYU, C_SIKEN_KOU_TYU		'前期中間、後期中間
			'== 開始日 ==
			p_sKaisibi = w_oRecordset("M00_KANRI")
			
			w_oRecordset.MoveNext
			
			'== 終了日 ==
			p_sSyuryobi = FormatDateTime(DateAdd("d", -1, w_oRecordset("M00_KANRI")))
		Case C_SIKEN_ZEN_KIM, C_SIKEN_KOU_KIM		'前期期末、後期期末
			'== 開始日 ==
			p_sKaisibi = FormatDateTime(DateAdd("d", 1, w_oRecordset("T24_JISSI_SYURYO")))
			w_oRecordset.MoveNext
			'== 終了日 ==
			p_sSyuryobi = FormatDateTime(DateAdd("d", -1, w_oRecordset("T24_JISSI_KAISI")))
			
	End Select
	
	'== 閉じる ==
	w_oRecordset.Close
	Set w_oRecordset = Nothing
	
	gf_GetKaisiSyuryo = True
	
End Function

'////////////////////////////////////////////////////////////////////////
'// ページに渡された引数を表示（デバッグ用）
'//
'// 引　数：request.form
'// 戻り値：なし
'// 詳細　：引数名＝引数値<br>の形で全ての引数を表示する。
'// 備考　：methodがpostの場合にのみ有効です。getの場合はプロパティを見てください。
'// 
'////////////////////////////////////////////////////////////////////////
sub gs_viewForm(p_form)
for each name In p_form
    response.write name&"="&p_form(name)&"<br>"
next

end sub

'/// 関数名変更のお知らせ。7/20 まで
sub s_viewForm(p_form)
    response.write "関数名が変わりました。<br>"
    response.write "call gs_viewForm(request.form)<br>"
    response.write "を使ってください。谷脇"
end sub

'********************************************************************************
'*  [機能]  最大時限数を取得
'*  [引数]  
'*  [戻値]  p_iJMax:最大時限数
'*  [説明]  
'********************************************************************************
Function gf_GetJigenMax(p_iJMax)

    Dim w_Rs                '// ﾚｺｰﾄﾞｾｯﾄｵﾌﾞｼﾞｪｸﾄ
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    
    On Error Resume Next
    Err.Clear
    
    p_iJMax = ""

    Do
        
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT MAX(""T20_JIGEN"") AS MAXJIGEN"
        w_sSQL = w_sSQL & vbCrLf & " FROM T20_JIKANWARI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & " T20_NENDO = " & SESSION("NENDO")

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            'm_bErrFlg = True
            Exit Do
        End If
        
        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            'm_bErrFlg = True
            Exit Do
        End If
        
        '// 取得した値を格納
        p_iJMax = CInt(w_Rs("MAXJIGEN"))
        '// 正常終了
        Exit Do

    Loop

    gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  学籍項目レベル別表示
'*  [引数]  p_ItemNo：項目のNO
'*  [戻値]  true/false
'*  [説明]  権限別の項目表示可否を出します。
'********************************************************************************
Function gf_empItem(p_ItemNo)
	Dim w_sLevel
	Dim w_iRet,w_Rs,w_sSql
	
 	gf_empItem = false

'===============================(デバッグ用)
' 	gf_empItem = True
'===============================

	w_sLevel = "T50_LEVEL" & Trim(Session("LEVEL"))
'	w_sLevel = "T50_LEVEL1"
'response.write Session("LEVEL")
    Do
	w_sSql = ""
	w_sSql = w_sSql & "Select " & w_sLevel & " "
	w_sSql = w_sSql & "From T50_KOMOKU_LEVEL "
	w_sSql = w_sSql & "Where T50_NO = " & p_ItemNo & " "

	w_iRet = gf_GetRecordset(w_Rs, w_sSql)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            'm_bErrFlg = True
            Exit Do
        End If

        If w_Rs.EOF Then
            '対象ﾚｺｰﾄﾞなし
            'm_bErrFlg = True
            Exit Do
        End If

        '// 表示権限がある場合はtrueを返す。
        If CInt(gf_SetNull2Zero(w_Rs(w_sLevel))) = 1 Then
			gf_empItem = true
			m_HyoujiFlg = 1			'<-- 表示ﾌﾗｸﾞ	08/01追加(ﾓﾁﾅｶﾞ)
		End if

        '// 正常終了
        Exit Do

    Loop
    gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  メニュー項目レベル別表示
'*  [引数]  p_iMenuID：項目のNO
'*  [戻値]  true/false
'*  [説明]  権限別の項目表示可否を出します。
'********************************************************************************
Function gf_empMenu(p_iMenuID)
	Dim w_sLevel
	Dim w_iRet,w_Rs,w_sSq
	Dim w_Where
	
	gf_empMenu = false

	'// Session("LEVEL")がNULLなら、ぬける
	if gf_IsNull(Trim(Session("LEVEL"))) then Exit Function

	'// Session("LEVEL")が"0"なら、ぬける
	if Cint(Session("LEVEL")) = Cint(0) then Exit Function

	w_sLevel = "T51_LEVEL" & Trim(Session("LEVEL"))

	'// WHERE文作成
	Select Case p_iMenuID
		Case "WEB0300" : w_Where = "T51_ID in ('WEB0300','WEB0301','WEB0302')"
		Case "WEB0320" : w_Where = "T51_ID in ('WEB0320','WEB0321')"
		Case "WEB0340" : w_Where = "T51_ID in ('WEB0340','WEB0341','WEB0342')"
		Case "WEB0390" : w_Where = "T51_ID in ('WEB0390','WEB0391','WEB0392')"
		Case "SEI0200" : w_Where = "T51_ID in ('SEI0200','SEI0210','SEI0221','SEI0222','SEI0223','SEI0224','SEI0230')"
		Case "SEI0300" : w_Where = "T51_ID in ('SEI0300','SEI0301','SEI0302')"
		Case Else :		 w_Where = "T51_ID =  '" & p_iMenuID & "'"
	End Select

    Do
		w_sSql = ""
		w_sSql = w_sSql & "Select " & w_sLevel & " "
		w_sSql = w_sSql & "From T51_SYORI_LEVEL "
		w_sSql = w_sSql & "Where "
		w_sSql = w_sSql & 		w_Where
		w_sSql = w_sSql & " ORDER BY  T51_ID "
		w_iRet = gf_GetRecordset(w_Rs, w_sSql)

		If w_iRet <> 0 Then
		    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		    'm_bErrFlg = True
		    Exit Do
		End If

		If w_Rs.EOF = true Then
		    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		    'm_bErrFlg = True
		    Exit Do
		End If

		w_flg = false
		w_Rs.movefirst
		Do Until w_Rs.EOF
			If trim(gf_SetNull2String(w_Rs(w_sLevel))) = "1" then 
				w_flg = true
				exit do
			end if
		w_Rs.movenext
		Loop

		If w_flg <> true Then
		    '対象ﾚｺｰﾄﾞなし
		    'm_bErrFlg = True
		    Exit Do
		End If

		'// 表示権限がある場合はtrueを返す。
	'	If Cint(gf_SetNull2Zero(w_Rs(w_sLevel))) = 1 Then gf_empMenu = true

		gf_empMenu = true

		'// 正常終了
		Exit Do

    Loop

End Function

'********************************************************************************
'*  [機能]  メニュー項目レベル別表示
'*  [引数]  p_iMenuID：項目のNO
'*  [戻値]  true/false
'*  [説明]  権限別の項目表示可否を出します。
'********************************************************************************
Function gf_empPasChg()
	Dim w_iRet,w_Rs,w_sSql,i	
	Dim w_Where
	
	gf_empPasChg = false

    Do
		w_sSql = ""
		w_sSql = w_sSql & "Select * "
		w_sSql = w_sSql & "From T51_SYORI_LEVEL "
		w_sSql = w_sSql & "Where "
		w_sSql = w_sSql & "T51_ID = 'WEB0400'"
		w_iRet = gf_GetRecordset(w_Rs, w_sSql)

		If w_iRet <> 0 Then
		    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		    'm_bErrFlg = True
		    Exit Do
		End If

		w_Rs.MoveFirst
		If w_Rs.EOF = true Then
		    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		    'm_bErrFlg = True
		    Exit Do
		End If

		w_flg = false
		For i = 3 to 13 
			If w_Rs(i) = 1 then 
				w_flg = true
				exit do
			end if
		Next

		If w_flg <> true Then
		    '対象ﾚｺｰﾄﾞなし
		    'm_bErrFlg = True
		    Exit Do
		End If
		'// 表示権限がある場合はtrueを返す。
	'	If Cint(gf_SetNull2Zero(w_Rs(w_sLevel))) = 1 Then gf_empMenu = true

		gf_empPasChg = true

		'// 正常終了
		Exit Do

    Loop

End Function

'********************************************************************************
'*  [機能]  不正アクセスチェック
'*  [引数]  p_PRJ_No = 権限ﾁｪｯｸのキー	C_LEVEL_NOCHKは、権限ﾁｪｯｸをしない
'*  [戻値]  なし
'*  [説明]  データベースに接続後に使用
'********************************************************************************
Function gf_userChk(p_PRJ_No)

	On Error Resume Next
	Err.Clear

	gf_userChk = False
	m_bErrFlg = False

	Do

		'// ログインチェック
		if gf_IsNull(Session("LOGIN_ID")) then
			m_bErrFlg = True
		    w_sWinTitle="キャンパスアシスト"
		    w_sMsgTitle="ログインエラー"
		    w_sRetURL = C_RetURL & "default.asp"
            m_sErrMsg = "セッションがタイムアウトされました\n再度ログインしなおしてください"
			w_sTarget = "_top"
			Exit do
		End if

		'// p_PRJ_NoがC_LEVEL_NOCHKは、権限ﾁｪｯｸをしない
		if p_PRJ_No = C_LEVEL_NOCHK then Exit Do

		'// 権限チェック
		if Not gf_empMenu(p_PRJ_No) then
			m_bErrFlg = True
		    w_sWinTitle="キャンパスアシスト"
		    w_sMsgTitle="権限エラー"
		    w_sRetURL = C_RetURL & "login/default.asp"
            m_sErrMsg = "権限がありません"
			w_sTarget = "_top"
			Exit do
		End if

		Exit do
	Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		'// 強制終了
'		response.end
    End If

	gf_userChk = True

End Function

'********************************************************************************
'*  [機能]  前期・後期情報を取得
'*  [引数]  なし
'*  [戻値]  p_sGakki		:学期CD
'*			p_sZenki_Start	:前期開始日
'*			p_sKouki_Start	:後期開始日
'*			p_sKouki_End	:後期終了日
'*  [説明]  
'********************************************************************************
Function gf_GetGakkiInfo(p_sGakki,p_sZenki_Start,p_sKouki_Start,p_sKouki_End)

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    gf_GetGakkiInfo = 1

	p_sZenki_Start = ""
	p_sKouki_Start = ""
	p_sKouki_End = ""
	p_sGakki = ""

    Do
        '管理マスタから学期情報を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NO, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_KANRI, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_BIKO"
        w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO=" & SESSION("NENDO") & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   (M00_KANRI.M00_NO=" & C_K_ZEN_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_SYURYO & ") "  '//[M00_NO]10:前期開始 11:後期開始

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrMsg = Err.description
            gf_GetGakkiInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            Do Until rs.EOF

                If cInt(rs("M00_NO")) = C_K_ZEN_KAISI Then
                    p_sZenki_Start = rs("M00_KANRI")
                ElseIf cInt(rs("M00_NO")) = C_K_KOU_KAISI Then
                    p_sKouki_Start = rs("M00_KANRI")
                ElseIf cInt(rs("M00_NO")) = C_K_KOU_SYURYO Then
                    p_sKouki_End = rs("M00_KANRI")
                End If
                rs.MoveNext
            Loop

            '//現在の前期後期判定
            If gf_YYYY_MM_DD(date(),"/") < p_sKouki_Start Then
                p_sGakki = C_GAKKI_ZENKI
            Else
                p_sGakki = C_GAKKI_KOUKI
            End If

        End If

        '//正常終了
        gf_GetGakkiInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  曜日を取得
'*  [引数]  p_CD(曜日ｺｰﾄﾞ)
'*  [戻値]  gf_GetYoubi
'*  [説明]  曜日の略称を返す
'********************************************************************************
Function gf_GetYoubi(p_CD)
Dim w_sYoubi

    On Error Resume Next
    Err.Clear

    w_sYoubi = ""
	w_sYoubi= WeekdayName(cInt(p_CD), True)

	'//戻り値をｾｯﾄ
    gf_GetYoubi = w_sYoubi

    Err.Clear

End Function

'********************************************************************************
'*  [機能]  担任の確認
'*  [引数]  p_iNendo	処理年度
'*		   p_iKyokan　教官コード
'*		   p_iBefore	有効年度（処理年度を含め、何年前までさかのぼって調べるのか）
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function gf_Tannin(p_iNendo,p_iKyokanCd,p_iBefore)

    Dim w_iRet
    Dim w_sSQL
    Dim rs
    Dim w_Cnt

    On Error Resume Next
    Err.Clear

    gf_Tannin = 1

    Do
        'クラスマスタから担任情報を取得
		w_sSQL = ""
		w_sSQL = w_sSQL & "	SELECT "
		w_sSQL = w_sSQL & "		M05_TANNIN "
		w_sSQL = w_sSQL & "	FROM "
		w_sSQL = w_sSQL & "		M05_CLASS "
		w_sSQL = w_sSQL & "	WHERE "
		w_sSQL = w_sSQL & "		M05_TANNIN = '"& p_iKyokanCd & "' "
		w_sSQL = w_sSQL & "	AND M05_NENDO <= " & p_iNendo & " "
		w_sSQL = w_sSQL & "	AND M05_NENDO > " & p_iNendo - p_iBefore& " "
		Set rs = Server.CreateObject("ADODB.Recordset")
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			m_bErrFlg = True
			Exit Do 
		End If
		w_Cnt=cint(gf_GetRsCount(rs))
'		If w_Cnt = 0 Then
'			Exit Do
'		End If
		If rs.EOF then Exit Do	'レコードセットが取れなかったとき。

        '//正常終了
        gf_Tannin = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  現在の日付に一番近い試験区分を取得
'*  [引数]  p_kikan			：対象となる期間（下の定数参照）
'* 		   p_gakunen		：対象となる学年（0の時は、全学年）
'*  [戻値]  なし
'*  [説明]  初期表示は現在の日付に一番近い試験を知る。
'* C_SIKEN_KIKAN：準備期間　C_JISSI_KIKAN：実施期間　C_SEISEKI_KIKAN：成績登録期間
'********************************************************************************
Function gf_Get_SikenKbn(p_iSiken_Kbn,p_kikan,p_gakunen)
    Dim w_iRet,w_kikanFld
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    gf_Get_SikenKbn = 1
    p_iSiken_Kbn = ""
    w_kikanFld = ""
    
    Select Case p_kikan
    	case C_SIKEN_KIKAN
		w_kikanFld = "T24_SIKEN_SYURYO"
    	case C_JISSI_KIKAN
		w_kikanFld = "T24_JISSI_SYURYO"
    	case C_SEISEKI_KIKAN
		w_kikanFld = "T24_SEISEKI_SYURYO"
    End Select
    
    Do
        '現在の日付に一番近い試験区分を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "    MIN(T24_SIKEN_KBN) as SIKEN_KBN"
        w_sSQL = w_sSQL & " FROM T24_SIKEN_NITTEI"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "       T24_NENDO = " & session("NENDO")
if p_gakunen > 0 then 
        w_sSQL = w_sSQL & "   AND T24_GAKUNEN = " & p_gakunen
end if
        w_sSQL = w_sSQL & "   AND " & w_kikanFld & " >= '" & gf_YYYY_MM_DD(date(),"/") & "'"
'        w_sSQL = w_sSQL & " ORDER BY " & w_kikanFld &" ASC"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            gf_Get_SikenKbn = 99
            Exit Do
        End If

        'If rs.EOF = False And ISNULL(rs("SIKEN_KBN")) = False Then
        If ISNULL(rs("SIKEN_KBN")) = False Then
            p_iSiken_Kbn = rs("SIKEN_KBN")
		Else
            p_iSiken_Kbn = C_SIKEN_ZEN_TYU
        End If

        '//正常終了
        gf_Get_SikenKbn = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  教官の氏名を取得(表示用)
'*  [引数]  なし
'*  [戻値]  f_GetKyokanNm:教官姓名
'*  [説明]  
'********************************************************************************
Function gf_GetKyokanNm(p_iNendo,p_sCD)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    gf_GetKyokanNm = ""
	w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M04_KYOKAN_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M04_NENDO = " & p_iNendo & " "

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M04_KYOKANMEI_SEI") & "　" & rs("M04_KYOKANMEI_MEI")
        End If

        Exit Do
    Loop

	'//戻り値ｾｯﾄ
    gf_GetKyokanNm = w_sName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

'********************************************************************************
'*  [機能]  学科名を取得(表示用)
'*  [引数]  なし
'*  [戻値]  gf_GetGakkaNm:学科名
'*  [説明]  
'********************************************************************************
Function gf_GetGakkaNm(p_iNendo,p_sCD)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    gf_GetGakkaNm = ""
	w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKAMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKA "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M02_GAKKA_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M02_NENDO = " & p_iNendo & " "

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M02_GAKKAMEI")
        End If

        Exit Do
    Loop

	'//戻り値ｾｯﾄ
    gf_GetGakkaNm = w_sName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

'********************************************************************************
'*  [機能]  クラス名を取得する
'*  [引数]  p_iNendo  ：処理年度
'*          p_iGakuNen：学年
'*          p_ClassNo ：クラスNO
'*  [戻値]  gf_GetClassName：クラス名
'*  [説明]  
'********************************************************************************
Function gf_GetClassName(p_iNendo,p_iGakuNen,p_ClassNo)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	gf_GetClassName = ""
	w_sClassName = ""

	Do

		'//クラス名称取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_CLASSMEI"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_NENDO=" & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN=" & p_iGakuNen
		w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_CLASSNO=" & p_ClassNo

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If rs.EOF = False Then
			'//クラス名
			w_sClassName = rs("M05_CLASSMEI")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	gf_GetClassName = w_sClassName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  １年間番号、５年間番号の名称を取得する
'*  [引数]  p_iNendo  ：処理年度
'*          p_iGakuKBN：１年間番号or５年間番号
'*  [戻値]  gf_GetGakuNomei：名称
'*  [説明]  
'********************************************************************************
Function gf_GetGakuNomei(p_iNendo,p_iGakuKBN)
	Dim w_iRet
	Dim w_sSQL
	Dim w_gaku_rs

	On Error Resume Next
	Err.Clear

	gf_GetGakuNomei = ""
	w_sGakuNomei = ""

	Do

		'//クラス名称取得
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M00_KANRI "
		w_sSql = w_sSql & vbCrLf & " FROM M00_KANRI "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M00_NENDO=" & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND M00_NO=" & p_iGakuKBN
		w_sSql = w_sSql & vbCrLf & "  AND M00_SYUBETU= 0 "

		'//ﾚｺｰﾄﾞｾｯﾄ取得
		w_iRet = gf_GetRecordset(w_gaku_rs, w_sSQL)
		If w_iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			Exit Do
		End If

		'//データが取得できたとき
		If w_gaku_rs.EOF = False Then
			'//クラス名
			w_sGakuNomei = w_gaku_rs("M00_KANRI")
		End If

		Exit Do
	Loop

	'//戻り値ｾｯﾄ
	gf_GetGakuNomei = w_sGakuNomei

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_gaku_rs)

End Function

'********************************************************************************
'*  [機能]  USERマスタよりUSER名を取得(表示用)
'*  [引数]  なし
'*  [戻値]  gf_GetUserNm:USER名
'*  [説明]  
'********************************************************************************
Function gf_GetUserNm(p_iNendo,p_sID)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    gf_GetUserNm = ""
	w_sName = ""

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M10_USER.M10_USER_NAME"
		w_sSQL = w_sSQL & vbCrLf & " FROM M10_USER"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M10_USER.M10_NENDO=" & p_iNendo
 		w_sSQL = w_sSQL & vbCrLf & " AND M10_USER.M10_USER_ID='" & p_sID & "'"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M10_USER_NAME")
        End If

        Exit Do
    Loop

	'//戻り値ｾｯﾄ
    gf_GetUserNm = w_sName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

'********************************************************************************
'*  [機能]  特別教室予約の権限を取得する
'*  [引数]  なし
'*  [戻値]  p_sKengen
'*  [説明]  
'********************************************************************************
Function gf_GetKengen_web0300(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0300 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('WEB0300','WEB0301','WEB0302') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            gf_GetKengen_web0300 = 99
            Exit Do
        End If

		if wLevRs.Eof then
            msMsg = "権限を取得できませんでした"
            gf_GetKengen_web0300 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "WEB0300" : p_sKengen = C_ACCESS_FULL			'//アクセス権限FULLアクセス可
			Case "WEB0301" : p_sKengen = C_ACCESS_NORMAL        '//アクセス権限一般
			Case "WEB0302" : p_sKengen = C_ACCESS_VIEW          '//アクセス権限参照のみ
		End Select

'		p_sKengen = C_ACCESS_FULL   'C_ACCESS_FULL   = "FULL"		
		'p_sKengen = C_ACCESS_NORMAL 'C_ACCESS_NORMAL = "NORMAL"	
		'p_sKengen = C_ACCESS_VIEW   'C_ACCESS_VIEW   = "VIEW"		

		'== 閉じる ==
	    Call gf_closeObject(wLevRs)

        gf_GetKengen_web0300 = 0
        Exit Do
    Loop

End Function


'********************************************************************************
'*  [機能]  使用教科書登録処理の権限を取得する
'*  [引数]  なし
'*  [戻値]  p_sKengen
'*  [説明]  
'********************************************************************************
Function gf_GetKengen_web0320(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0320 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('WEB0320','WEB0321') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            gf_GetKengen_web0320 = 99
            Exit Do
        End If

		if wLevRs.Eof then
            msMsg = "権限を取得できませんでした"
            gf_GetKengen_web0320 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "WEB0320" : p_sKengen = C_WEB0320_ACCESS_FULL  		'//アクセス権限FULLアクセス可
			Case "WEB0321" : p_sKengen = C_WEB0320_ACCESS_NORMAL       '//アクセス権限一般
		End Select

		'p_sKengen =  C_WEB0320_ACCESS_FULL     'C_ACCESS_FULL   = "FULL"		'//アクセス権限FULLアクセス可
		'p_sKengen =   C_WEB0320_ACCESS_NORMAL   'C_ACCESS_NORMAL = "NORMAL"		'//アクセス権限一般

		'== 閉じる ==
	    Call gf_closeObject(wLevRs)

        gf_GetKengen_web0320 = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  個人履修選択科目決定処理の権限を取得する
'*  [引数]  なし
'*  [戻値]  p_sKengen
'*  [説明]  
'********************************************************************************
Function gf_GetKengen_web0340(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0340 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('WEB0340','WEB0341','WEB0342') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            gf_GetKengen_web0340 = 99
            Exit Do
        End If

		if wLevRs.Eof then
            msMsg = "権限を取得できませんでした"
            gf_GetKengen_web0340 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "WEB0340" : p_sKengen = C_WEB0340_ACCESS_FULL  		'//アクセス権限FULLアクセス可
			Case "WEB0341" : p_sKengen = C_WEB0340_ACCESS_SENMON        '//アクセス権限担当教官のみ
			Case "WEB0342" : p_sKengen = C_WEB0340_ACCESS_TANNIN        '//アクセス権限担任のみ
		End Select

		'p_sKengen =  C_WEB0340_ACCESS_FULL     'C_ACCESS_FULL   = "FULL"		'//アクセス権限FULLアクセス可
'		p_sKengen =   C_WEB0340_ACCESS_SENMON   'C_ACCESS_SENMON = "SENMON"		'//アクセス権限担当教官のみ
		'p_sKengen =  C_WEB0340_ACCESS_TANNIN   'C_ACCESS_TANNIN = "TANNIN"		'//アクセス権限担任のみ

		'== 閉じる ==
	    Call gf_closeObject(wLevRs)

        gf_GetKengen_web0340 = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  レベル別科目決定処理の権限を取得する
'*  [引数]  なし
'*  [戻値]  p_sKengen
'*  [説明]  
'********************************************************************************
Function gf_GetKengen_web0390(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0390 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('WEB0390','WEB0391','WEB0392') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            gf_GetKengen_web0390 = 99
            Exit Do
        End If

		if wLevRs.Eof then
            msMsg = "権限を取得できませんでした"
            gf_GetKengen_web0390 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "WEB0390" : p_sKengen = C_WEB0340_ACCESS_FULL  		'//アクセス権限FULLアクセス可
			Case "WEB0391" : p_sKengen = C_WEB0340_ACCESS_SENMON        '//アクセス権限担当教官のみ
			Case "WEB0392" : p_sKengen = C_WEB0340_ACCESS_TANNIN        '//アクセス権限担任のみ
		End Select

		'p_sKengen =  C_WEB0340_ACCESS_FULL     'C_ACCESS_FULL   = "FULL"		'//アクセス権限FULLアクセス可
'		p_sKengen =   C_WEB0340_ACCESS_SENMON   'C_ACCESS_SENMON = "SENMON"		'//アクセス権限担当教官のみ
		'p_sKengen =  C_WEB0340_ACCESS_TANNIN   'C_ACCESS_TANNIN = "TANNIN"		'//アクセス権限担任のみ

		'== 閉じる ==
	    Call gf_closeObject(wLevRs)

        gf_GetKengen_web0390 = 0
        Exit Do
    Loop

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
'********************************************************************************
Function gf_GetKekaChi(p_iNendo,p_Syubetu,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku,p_iKekkaGai)

    Dim w_sSQL
    Dim w_KekaChiRs
    Dim w_sKek,p_sChi
    Dim w_Table,w_TableName
    Dim w_Kamoku
    
    On Error Resume Next
    Err.Clear
    
    gf_GetKekaChi = 1
    
    p_iKekka = 0
    p_iChikoku = 0
	
	'特別授業、その他(通常など)の切り分け
	if trim(p_Syubetu) = "TOKU" then
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_Kamoku = "T34_TOKUKATU_CD"
	else
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_Kamoku = "T16_KAMOKU_CD"
	end if
	
	'/ 試験区分によって取ってくる、フィールドを変える。
	Select Case p_iSikenKBN
		Case C_SIKEN_ZEN_TYU
			w_sKek = w_Table & "_KEKA_TYUKAN_Z"
			w_sKekG= w_Table & "_KEKA_NASI_TYUKAN_Z"
			p_sChi = w_Table & "_CHIKAI_TYUKAN_Z "
		Case C_SIKEN_ZEN_KIM
			w_sKek = w_Table & "_KEKA_KIMATU_Z"
			w_sKekG= w_Table & "_KEKA_NASI_KIMATU_Z"
			p_sChi = w_Table & "_CHIKAI_KIMATU_Z "
		Case C_SIKEN_KOU_TYU
			w_sKek = w_Table & "_KEKA_TYUKAN_K"
			w_sKekG= w_Table & "_KEKA_NASI_TYUKAN_K"
			p_sChi = w_Table & "_CHIKAI_TYUKAN_K"
		Case C_SIKEN_KOU_KIM
			w_sKek = w_Table & "_KEKA_KIMATU_K"
			w_sKekG= w_Table & "_KEKA_NASI_KIMATU_K"
			p_sChi = w_Table & "_CHIKAI_KIMATU_K"
	End Select
	
	w_sSQL = ""
	w_sSQL = w_sSQL &  " SELECT "
	w_sSQL = w_sSQL & " " & w_sKek & " as KEKA, "
	w_sSQL = w_sSQL & " " & w_sKekG & " as KEKA_NASI, "
	w_sSQL = w_sSQL & " " & p_sChi & " as CHIKAI "
	w_sSQL = w_sSQL & " FROM " & w_TableName
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "      " & w_Table & "_NENDO =" & p_iNendo
	w_sSQL = w_sSQL & "  AND " & w_Table & "_GAKUSEI_NO= '" & p_sGakusei & "'"
	w_sSQL = w_sSQL & "  AND " & w_Kamoku & "= '" & p_sKamokuCD & "'"
	
	'response.write "w_sSQL =" & w_sSQL & "<BR>"
	
    If gf_GetRecordset(w_KekaChiRs, w_sSQL) <> 0 Then
        'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        msMsg = Err.description
    End If
	
	'//戻り値ｾｯﾄ
	If w_KekaChiRs.EOF = False Then
		p_iKekka = w_KekaChiRs("KEKA")
		p_iKekkaGai = w_KekaChiRs("KEKA_NASI")
		p_iChikoku = w_KekaChiRs("CHIKAI")
	End If
	
    gf_GetKekaChi = 0
    
    Call gf_closeObject(w_KekaChiRs)

End Function

'********************************************************************************
'*  [機能]  管理マスタより出欠欠課の取り方を取得
'*  [引数]  なし
'*  [戻値]  p_sSyubetu = C_K_KEKKA_RUISEKI_SIKEN : 試験毎(=0)
'*  [戻値]  p_sSyubetu = C_K_KEKKA_RUISEKI_KEI   ：累積(=1)
'*  [説明]  
'********************************************************************************
Function gf_GetKanriInfo(p_iNendo,p_iSyubetu)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    gf_GetKanriInfo = 1

    Do 

		'//管理マスタより欠課累積情報区分を取得
		'//欠課累積情報区分(C_K_KEKKA_RUISEKI = 32)
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_SYUBETU"
		w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_NENDO=" & cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M00_KANRI.M00_NO=" & C_K_KEKKA_RUISEKI	'欠課累積情報区分(=32)

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            gf_GetKanriInfo = 99
            Exit Do
        End If

		'//戻り値ｾｯﾄ
		If w_Rs.EOF = False Then
			'//Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '試験毎
			'//Public Const C_K_KEKKA_RUISEKI_KEI = 1      '累積
			p_iSyubetu = w_Rs("M00_SYUBETU")

		End If

        gf_GetKanriInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [機能]  出欠の入力ができなくなる日を取得
'*  [引数]  p_gakunen		：対象となる学年
'*  [戻値]  p_sEndDay		：出欠の入力ができなくなる日
'*  [説明]  
'********************************************************************************
Function gf_Get_SyuketuEnd(p_iGakunen,p_sEndDay)
    Dim w_iRet,w_sSQL,rs
	Dim w_sGakki,w_sZenki_Start,w_sKouki_Start,w_sKouki_End
    Dim w_sDate

    On Error Resume Next
    Err.Clear

    gf_Get_SyuketuEnd = 1

	w_sDate = gf_YYYY_MM_DD(date(),"/")
	'学期情報の取得
	call gf_GetGakkiInfo(p_sGakki,p_sZenki_Start,p_sKouki_Start,p_sKouki_End)

	'初期値代入（前期開始日）
    p_sEndDay = p_sZenki_Start
    
    Do
        '現在の日付に一番近い試験区分を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "    T24_JISSI_SYURYO "
'        w_sSQL = w_sSQL & "    T24_SEISEKI_SYURYO "	'--2001/12/18 add 試験終了日を見る
        w_sSQL = w_sSQL & " FROM T24_SIKEN_NITTEI "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "       T24_NENDO = " & session("NENDO")
        w_sSQL = w_sSQL & "   AND T24_GAKUNEN = " & p_iGakunen
        w_sSQL = w_sSQL & " ORDER BY T24_JISSI_SYURYO DESC"
'        w_sSQL = w_sSQL & " ORDER BY T24_SEISEKI_SYURYO DESC" '--2001/12/18 add 試験終了日を見る

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            gf_Get_SikenKbn = 99
            Exit Do
        End If

		rs.MoveFirst
		Do Until rs.EOF

			'成績入力期間終了日を過ぎていたら、
			'出欠入力できなくなる日をその成績入力期間終了日に設定
			If rs("T24_JISSI_SYURYO") < w_sDate then 
'			If rs("T24_SEISEKI_SYURYO") < w_sDate then ' --2001/12/18 add 試験終了日を見る
				p_sEndDay = rs("T24_JISSI_SYURYO")
'				p_sEndDay = rs("T24_SEISEKI_SYURYO") ' --2001/12/18 add 試験終了日を見る
				Exit Do
			End If
			rs.MoveNext

		Loop
		
        '//正常終了
        gf_Get_SyuketuEnd = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  郵便番号から住所を取得する
'*  [引数]  
'*			p_sKenMei - 県名
'*			p_sSityosonCD - 市町村CD
'*			p_sSityosonMei - 市町村名
'*			p_sZipCD - 郵便番号
'*			p_sTyoikiMei - 町域名
'*  [戻値]  取得結果
'*  [戻値]  True(OK),False(Cancel)
'*  [説明]  
'********************************************************************************
Function gf_ComvZip(ByRef p_sZipCD,ByRef p_sKenMei,ByRef p_sSityosonCD,ByRef p_sSityosonMei,ByRef p_sTyoikiMei,ByRef p_iNendo)

    Dim w_bRtn
    Dim w_sSQL
    Dim w_oRecord
    Dim w_oclsSch
    
    On Error Resume Next
    Err.Clear
    
    '== 初期化 ==
    gf_ComvZip = 1
    
    p_sKenMei = ""
    p_sSityosonCD = ""
    p_sSityosonMei = ""
    p_sTyoikiMei = ""

Do
    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "M12_SITYOSON_CD,"
    w_sSQL = w_sSQL & "M12_SITYOSONMEI,"
    w_sSQL = w_sSQL & "M12_TYOIKIMEI "
    w_sSQL = w_sSQL & ", M16_KENMEI "           '2001/07/23 Mod
    w_sSQL = w_sSQL & "FROM M12_SITYOSON "
    w_sSQL = w_sSQL & ", M16_KEN "              '2001/07/23 Mod
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & " M12_YUBIN_BANGO = '" & p_sZipCD & "' "
    w_sSQL = w_sSQL & " And "
    w_sSQL = w_sSQL & " M16_NENDO = " & cint(p_iNendo)
    w_sSQL = w_sSQL & " And "
    w_sSQL = w_sSQL & " M16_KEN_CD = M12_KEN_CD "
    w_sSQL = w_sSQL & " Order By "
    w_sSQL = w_sSQL & " M12_YUBIN_BANGO, "
    w_sSQL = w_sSQL & " M12_RENBAN"

'response.write w_sSQL & "<br>"

    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==
    w_bRtn = gf_GetRecordset(w_oRecord, w_sSQL)

    If w_bRtn <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            gf_Get_SikenKbn = 99
            Exit Do
    End If

'    If w_oRecord.EOF = False Then

        p_sKenMei = w_oRecord("M16_KENMEI")
        p_sSityosonCD = w_oRecord("M12_SITYOSON_CD")
        p_sSityosonMei = w_oRecord("M12_SITYOSONMEI")
        p_sTyoikiMei = w_oRecord("M12_TYOIKIMEI")

'	End If

        '//正常終了
        gf_ComvZip = 0
    	Exit Do
    Loop

    Call gf_closeObject(w_oRecord)
    
End Function

'********************************************************************************
'*  [機能]  郵便番号、市町村コードから住所を取得する
'*  [引数]  
'*			p_sKenMei - 県名
'*			p_sSityosonCD - 市町村CD
'*			p_sSityosonMei - 市町村名
'*			p_sZipCD - 郵便番号
'*			p_sTyoikiMei - 町域名
'*  [戻値]  取得結果
'*  [戻値]  True(OK),False(Cancel)
'*  [説明]  
'********************************************************************************
Function gf_ComvZip2(ByRef p_sSityosonCD1,ByRef p_sZipCD,ByRef p_sKenMei,ByRef p_sSityosonCD,ByRef p_sSityosonMei,ByRef p_sTyoikiMei,ByRef p_iNendo)

    Dim w_bRtn
    Dim w_sSQL
    Dim w_oRecord
    Dim w_oclsSch
    
    On Error Resume Next
    Err.Clear
    
    '== 初期化 ==
    gf_ComvZip2 = 1
    
    p_sKenMei = ""
    p_sSityosonCD = ""
    p_sSityosonMei = ""
    p_sTyoikiMei = ""

Do
    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "M12_SITYOSON_CD,"
    w_sSQL = w_sSQL & "M12_SITYOSONMEI,"
    w_sSQL = w_sSQL & "M12_TYOIKIMEI "
    w_sSQL = w_sSQL & ", M16_KENMEI "
    w_sSQL = w_sSQL & "FROM M12_SITYOSON "
    w_sSQL = w_sSQL & ", M16_KEN "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & " M12_YUBIN_BANGO = '" & p_sZipCD & "' "
    w_sSQL = w_sSQL & " And "
    w_sSQL = w_sSQL & " M16_NENDO = " & cint(p_iNendo)
    w_sSQL = w_sSQL & " And "
    w_sSQL = w_sSQL & " M16_KEN_CD = M12_KEN_CD "
    w_sSQL = w_sSQL & " And "
    w_sSQL = w_sSQL & " M12_SITYOSON_CD = '" & p_sSityosonCD1 & "' "
    w_sSQL = w_sSQL & " Order By "
    w_sSQL = w_sSQL & " M12_YUBIN_BANGO, "
    w_sSQL = w_sSQL & " M12_RENBAN"

'response.write w_sSQL & "<br>"

    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==
    w_bRtn = gf_GetRecordset(w_oRecord, w_sSQL)

    If w_bRtn <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            gf_Get_SikenKbn = 99
            Exit Do
    End If

'    If w_oRecord.EOF = False Then

        p_sKenMei = w_oRecord("M16_KENMEI")
        p_sSityosonCD = w_oRecord("M12_SITYOSON_CD")
        p_sSityosonMei = w_oRecord("M12_SITYOSONMEI")
        p_sTyoikiMei = w_oRecord("M12_TYOIKIMEI")

'	End If

        '//正常終了
        gf_ComvZip2 = 0
    	Exit Do
    Loop

    Call gf_closeObject(w_oRecord)
    
End Function

'********************************************************************************
'*	[機能]	異動ありの場合移動状況の取得
'*	[引数]	p_Gakusei_No:学績NO
'*			p_Date		:授業実施日
'*	[戻値]	0:情報取得成功 99:失敗
'*	[説明]	2001.12.19 版：岡田
'********************************************************************************
Function gf_Get_IdouChk(p_Gakuseki_No,p_Date,p_iNendo,ByRef p_sKubunName)

	Dim w_sSQL
	Dim w_Rs
	Dim w_IdoFlg
	Dim w_sKubunName

	On Error Resume Next
	Err.Clear

	w_IdoFlg = False

	Do

		'// 明細データ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_8, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_8"
		w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cint(p_iNendo) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO='" & p_Gakuseki_No & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM>0"

'response.write w_sSQL & "<br>"


		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			msMsg = Err.description
			Exit Do
		End If

		If w_Rs.EOF = 0 Then
			i = 1
			'//8…最大移動回数
			Do Until Cint(i) > cint(8)    '//C_IDO_MAX_CNT = 8…最大移動回数
				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) = "" Then
					Exit Do
				End If
'Response.Write "[" & gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) & " > " & p_Date & "]"
				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) > p_Date  Then
					Exit Do
				End If
				i = i + 1
			Loop

'response.write "学生ＮＯ" & p_Gakuseki_No & " : i = " & i
			w_sKubunName = ""

			If i = 1 then
				'//最初の移動日が授業日より未来の場合、授業日に移動状態ではない
				'w_IdoFlg = False
				'w_sKubunName = ""

				w_sKubunName = gf_SetNull2String(w_Rs("T13_IDOU_KBN_" & i))

				w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i)),p_iNendo,p_sKubunName)

			Else

   				w_sKubunName = gf_SetNull2String(w_Rs("T13_IDOU_KBN_" & i-1))

				 w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i-1)),p_iNendo,p_sKubunName)

			End If
'response.write "結果:" & w_sKubunName & "異動事由：" & p_sKubunName
		End If

		Exit Do
	Loop

	gf_Get_IdouChk = w_sKubunName

	Call gf_closeObject(w_Rs)

	Err.Clear

End Function

'********************************************************************************
'*	[機能]	異動名称取得関数
'*	[引数]	p_iGakusei_No:学績NO
'*			p_iNendo		:処理年度
'*	[戻値]	0:情報取得成功 1:失敗  p_SSSS : 異動名称
'*	[説明]	2001.12.22 版：岡田
'*	[修正]	2002.04.26 shin 復学、停学解除の場合は、戻り値１に設定
'********************************************************************************
Function gf_Set_Idou(p_sGakusekiCd,p_iNendo,ByRef p_SSSS)

		gf_Set_Idou = 1

		Dim w_Date
		Dim w_SSSR
		
'response.write p_iNendo & "/" & month(date()) & "/" & day(date())


'		w_Date = gf_YYYY_MM_DD(p_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		w_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")
 		'//C_IDO_FUKUGAKU=3:復学、C_IDO_TEI_KAIJO=5:停学解除
		'p_SSSS = ""
		w_SSSR = ""

		p_SSSS = gf_Get_IdouChk(p_sGakusekiCd,w_Date,p_iNendo,w_SSSR)

'response.write w_Date
'response.write w_SSSR & "<br>"


		IF CStr(p_SSSS) <> "" Then

			IF Cstr(p_SSSS) <> CStr(C_IDO_FUKUGAKU) AND Cstr(p_SSSS) <> Cstr(C_IDO_TEI_KAIJO) Then

					p_SSSS = w_SSSR

					gf_Set_Idou =0
			Else

				w_SSSR = ""
				p_SSSS = ""
			
				gf_Set_Idou = 1

			End if

		End if

'response.write p_SSSS

End Function

'********************************************************************************
'*  [機能]  履修データから更新日を取得する。
'*  [引数]  
'*			p_iNendo - 処理年度
'*			p_iGakunen - 学年
'*			p_sGakkaCd - 学科コード
'*			p_sKamokuCd - 科目コード
'*			p_iCourseCd - コースコード
'*  [戻値]  更新日付
'*  [説明]  
'********************************************************************************
Function gf_GetT16UpdDate(p_iNendo,p_iGakunen,p_sGakkaCd,p_sKamokuCd,p_iCourseCd)

    Dim w_bRtn
    Dim w_sSQL
    Dim w_oRecord
    
    On Error Resume Next
    Err.Clear
    
    '== 初期化 ==
    gf_GetT16UpdDate = ""

Do
    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & " T16_UPD_DATE "
    w_sSQL = w_sSQL & " FROM T16_RISYU_KOJIN "
    w_sSQL = w_sSQL & "WHERE "
    w_sSQL = w_sSQL & "     T16_NENDO        =  " & p_iNendo
    w_sSQL = w_sSQL & " And T16_HAITOGAKUNEN =  " & p_iGakunen
    w_sSQL = w_sSQL & " And T16_GAKKA_CD     = '" & p_sGakkaCd & "'"
    w_sSQL = w_sSQL & " And T16_KAMOKU_CD    = '" & p_sKamokuCd & "'"

'response.write w_sSQL & "<br>"

    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==
    w_bRtn = gf_GetRecordset(w_oRecord, w_sSQL)

    If w_bRtn <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            Exit Do
    End If

    gf_GetT16UpdDate = gf_SetNull2String(w_oRecord("T16_UPD_DATE"))

    Exit Do
Loop

    Call gf_closeObject(w_oRecord)
    
End Function

'*******************************************************************************
' 機　　能：出欠を取得する開始日と終了日また、試験実績の登録をした試験区分の取得
' 返　　値：(True)成功, (False)失敗
' 引　　数：p_sSikenKbn  - 試験区分
'			p_sGakunen   - 学年
'			p_SyoriNen   - 処理年度使用
'			p_GakusekiNo - 学籍番号
'			p_Kamoku     - 科目コード
'			p_Mode       - 処理モード(kks→授業出欠、other→成績登録、欠席登録)
'			p_Syubetu    - 科目種別(TUJO：通常授業)
'			(戻り値)p_ShikenInsertKbn - 試験区分 [実績をとるときに使用]
'			(戻り値)p_sKaisibi   - 開始日
'			(戻り値)p_sSyuryobi  - 終了日
'			
' 機能詳細：出欠を取得する開始日と終了日の取得
' 備　　考：gf_GetKaisiSyuryoのカスタマイズ版
' 　　　　：似ているが日付のとり方が違うため別関数に
'
' 　　　　：2002/06/13　shin
'*******************************************************************************
'Function gf_GetStartEnd(p_Mode,p_SyoriNen,p_sSikenKbn,p_sGakunen,p_GakusekiNo,p_Kamoku,p_sKaisibi,p_sSyuryobi,p_ShikenInsertKbn)
Function gf_GetStartEnd(p_Mode,p_SyoriNen,p_Syubetu,p_sSikenKbn,p_sGakunen,p_GakusekiNo,p_Kamoku,p_sKaisibi,p_sSyuryobi,p_ShikenInsertKbn)
	
	Dim w_iNendo		'年度
	Dim w_UpdateDate	'更新日
	Dim w_sGakki		'学期
	Dim w_sZenki_Start	'前期開始日
	Dim w_sKouki_Start	'後期開始日
	Dim w_sKouki_End	'後期終了日
	Dim w_iSyubetu
	Dim w_num
	
	On Error Resume Next
	Err.clear
	
	gf_GetStartEnd = False
	
	w_iNendo = p_SyoriNen
	
	'//前期・後期情報を取得
	if gf_GetGakkiInfo(w_sGakki,w_sZenki_Start,w_sKouki_Start,w_sKouki_End) <> 0 then : Exit function
	
	'//終了日を取得
	if p_Mode = "kks" then
		if month(date()) <= 3 then
			p_sSyuryobi = gf_YYYY_MM_DD((w_iNendo + 1) & "/" & month(date()) & "/" & day(date()),"/")
		else
			p_sSyuryobi = gf_YYYY_MM_DD(w_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		end if
	else
		if not gf_GetShikenDate(w_iNendo,p_sGakunen,p_sSikenKbn,p_sSyuryobi) then : exit function
	end if
	
	'//試験区分が、前期中間なら、試験の実績登録していないため、前期開始日を累計取得開始日にする
	if cint(p_sSikenKbn) = 1 then
		p_sKaisibi  = gf_YYYY_MM_DD(w_sZenki_Start,"/")
		gf_GetStartEnd = True
		exit function
	end if
	
	'//授業出欠登録(kks0100より呼ばれたとき)
	if p_Mode = "kks" then
		'//累計開始日を取得するため、試験の実績を登録した試験区分を取得する
		for w_num = cint(p_sSikenKbn)-1 to 1 Step -1
			
			'//科目の実績登録してあるか調べるため、更新日を取得する ==
			if not gf_GetUpdateDate(w_iNendo,p_Syubetu,p_Kamoku,p_GakusekiNo,w_num,w_UpdateDate) then : exit function
			
			if gf_SetNull2String(w_UpdateDate) <> "" then
				
				if not gf_GetShikenDate(w_iNendo,p_sGakunen,w_num,p_sKaisibi) then : exit function
				
				p_ShikenInsertKbn = w_num
				
				'//試験実施終了日の次の日から累計を開始するため＋１する
				p_sKaisibi = gf_YYYY_MM_DD(DateAdd("d",1,p_sKaisibi),"/")
				
				gf_GetStartEnd = True
				exit function
			end if
		next
		
		'ここにくるのは、試験登録していないとき なので、前期開始日をセット
		p_sKaisibi  = gf_YYYY_MM_DD(w_sZenki_Start,"/")
		
	else
		
		'//出欠欠課の取り方を取得 科目区分(0:試験毎,1:累積)
        If gf_GetKanriInfo(p_SyoriNen,w_iSyubetu) <> 0 Then : exit function
		
		if cint(w_iSyubetu) = C_K_KEKKA_RUISEKI_SIKEN then
			'開始日
			if not gf_GetShikenDate(w_iNendo,p_sGakunen,p_sSikenKbn-1,p_sKaisibi) then : exit function
			
			p_sKaisibi = gf_YYYY_MM_DD(DateAdd("d",1,p_sKaisibi),"/")
			
		else
			'//累計開始日を取得するため、試験の実績を登録した試験区分を取得する
			for w_num = cint(p_sSikenKbn)-1 to 1 Step -1
				
				'== 科目の実績登録してあるか調べるため、更新日を取得する ==
				if not gf_GetUpdateDate(w_iNendo,p_Syubetu,p_Kamoku,p_GakusekiNo,w_num,w_UpdateDate) then : exit function
				
				if gf_SetNull2String(w_UpdateDate) <> "" then
					
					'//試験実施終了日を取得する
					if not gf_GetShikenDate(w_iNendo,p_sGakunen,w_num,p_sKaisibi) then : exit function
					
					p_ShikenInsertKbn = w_num
					
					'//試験実施終了日の次の日から累計を開始するため＋１する
					p_sKaisibi = gf_YYYY_MM_DD(DateAdd("d",1,p_sKaisibi),"/")
					
					gf_GetStartEnd = True
					exit function
				end if
			next
			
			'ここにくるのは、試験登録していないとき なので、前期開始日をセット
			p_sKaisibi  = gf_YYYY_MM_DD(w_sZenki_Start,"/")
			
		end if
	end if
	
	gf_GetStartEnd = True
	
End Function

'*******************************************************************************
' 機　　能：科目の実績登録してあるか調べるため、更新日を取得する
' 
' 返　　値：
' 　　　　　(True)成功, (False)失敗
' 引　　数：_sSikenKbn - 試験区分
' 　　　　　p_Nendo - 年度
' 　　　　　p_KamokuCd - 科目コード
'			p_GakusekiNo - 学籍NO
'			(戻り値)p_UpdateDate - 科目実績登録の更新日
' 機能詳細：
' 備　　考：gf_GetStartEndで使用
'			Add 2002/06/13 shin
'*******************************************************************************
function gf_GetUpdateDate(p_Nendo,p_Syubetu,p_KamokuCd,p_GakusekiNo,p_ShikenKbn,p_UpdateDate)
	Dim w_Sql,w_Rs
	Dim w_ShikenType
	Dim w_Table
	Dim w_TableName
	Dim w_KamokuName
	
	On Error Resume Next
	Err.Clear
	
	gf_GetUpdateDate = false
	
	if trim(p_Syubetu) = "TOKU" then
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_KamokuName = "T34_TOKUKATU_CD"
	else
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	end if
	
	select case cint(p_ShikenKbn)
		
		case C_SIKEN_ZEN_TYU '前期中間試験
			w_ShikenType = w_Table & "_KOUSINBI_TYUKAN_Z"
			
		case C_SIKEN_ZEN_KIM '前期期末試験
			w_ShikenType = w_Table & "_KOUSINBI_KIMATU_Z"
			
		case C_SIKEN_KOU_TYU '後期中間試験
			w_ShikenType = w_Table & "_KOUSINBI_TYUKAN_K"
			
		case C_SIKEN_KOU_KIM '後期期末試験
			w_ShikenType = w_Table & "_KOUSINBI_KIMATU_K"
			
		case else
			exit function
			
	end select
	
	w_Sql = ""
	w_Sql = w_Sql & " select "
	w_Sql = w_Sql & " 		" & w_ShikenType
	w_Sql = w_Sql & " from "
	w_Sql = w_Sql & " 		" & w_TableName
	w_Sql = w_Sql & " where "
	w_Sql = w_Sql & " 		" & w_Table & "_NENDO = " & p_Nendo
	w_Sql = w_Sql & " and	" & w_Table & "_GAKUSEKI_NO = '" & p_GakusekiNo & "' "
	w_Sql = w_Sql & " and	" & w_KamokuName & "= '"   & p_KamokuCd   & "' "
	
	If gf_GetRecordset(w_Rs,w_Sql) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit function
	End If
	
	p_UpdateDate = w_Rs(0)
	
	gf_GetUpdateDate = true
	
end function

'*******************************************************************************
' 機　　能：試験実施終了日を取得する
' 返　　値：
' 			(True)成功, (False)失敗
' 引　　数：p_sSikenKbn - 試験区分
' 			p_sGakunen - 学年
' 			p_iNendo - 年度
'			(戻り値)p_UpdateDate - 終了日
' 機能詳細：
' 備　　考：gf_GetStartEndで使用
'			Add 2002/06/13 shin
'*******************************************************************************
function gf_GetShikenDate(p_iNendo,p_sGakunen,p_ShikenKbn,p_UpdateDate)
	Dim w_sSql,w_Rs
	
	On Error Resume Next
	Err.Clear
	
	gf_GetShikenDate = false
	
	w_sSql = ""
	w_sSql = w_sSql & " select "
	'w_sSql = w_sSql & "		T24_JISSI_SYURYO "
	w_sSql = w_sSql & "		T24_IDOU_SYURYO "
	w_sSql = w_sSql & " from "
	w_sSql = w_sSql & "		T24_SIKEN_NITTEI "
	w_sSql = w_sSql & " Where "
	w_sSql = w_sSql & "		T24_NENDO = " & p_iNendo
	w_sSql = w_sSql & " And "
	w_sSql = w_sSql & "		T24_SIKEN_KBN = " & p_ShikenKbn
	w_sSql = w_sSql & " And "
	w_sSql = w_sSql & "		T24_GAKUNEN = " & p_sGakunen
	
	If gf_GetRecordset(w_Rs,w_sSql) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		Exit function
	End If
	
	if w_Rs.EOF then : exit function
	
	if gf_SetNull2String(w_Rs(0)) = "" then : exit function
	
	p_UpdateDate = w_Rs(0)
	
	gf_GetShikenDate = true
	
end function

'*******************************************************************************
' 機　　能：出欠データの取得
' 返　　値：取得結果
' 　　　　　(True)成功, (False)失敗
' 引　　数：p_oRecordset - レコードセット
' 　　　　　p_sSikenKbn - 試験区分
' 　　　　　p_sGakunen - 学年
' 　　　　　p_sTantoKyokan - 教官ＣＤ
' 　　　　　p_sClass - クラスNo
' 　　　　　p_sKamokuCD - 科目コード
' 　　　　　p_sKaisibi - 開始日
' 　　　　　p_sSyuryobi - 終了日
' 　　　　　p_s1NenBango - １年間番号
' 機能詳細：指定された条件の出欠のデータを取得する
' 備　　考：なし
'*******************************************************************************
Function gf_GetSyukketuData2(p_oRecordset,p_sSikenKbn,p_sGakunen,p_sTantoKyokan,p_sClass,p_sKamokuCD, _
								p_sKaisibi,p_sSyuryobi,p_s1NenBango,p_Nendo,p_ShikenInsertType,p_GakusekiNo,p_Syubetu)
	Dim w_sSql
	
	On Error Resume Next
	
	'== 初期化 ==
	gf_GetSyukketuData2 = False
	
	'== 出欠を取得する開始日と終了日を取得する ==
	'//(試験間の期間)
	if not gf_GetStartEnd("other",p_Nendo,p_Syubetu,p_sSikenKbn,p_sGakunen,p_GakusekiNo,p_sKamokuCD,p_sKaisibi,p_sSyuryobi,p_ShikenInsertType) then
	'if gf_GetKaisiSyuryo(p_sSikenKbn, p_sGakunen, p_sKaisibi, p_sSyuryobi) <> True Then
		Exit Function
	End If
	
	'== 出欠を取得する ==
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & "SELECT "
	w_sSql = w_sSql & vbCrLf & "	Count(T21_GAKUSEKI_NO) as KAISU,"
	w_sSql = w_sSql & vbCrLf & "	T21_CLASS,"
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & "	T21_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "FROM "
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU "
	w_sSql = w_sSql & vbCrLf & "Where "
	w_sSql = w_sSql & vbCrLf & "	T21_NENDO = " & p_Nendo & " "		'年度
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_GAKUNEN = " & p_sGakunen & " "			'学年
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_KAMOKU = '" & p_sKamokuCD & "' " 		'科目
'	w_sSql = w_sSql & vbCrLf & "	And "
'	w_sSql = w_sSql & vbCrLf & "	T21_KYOKAN = '" & p_sTantoKyokan & "' "		'教官
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_HIDUKE >= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sKaisibi & "' "						'開始日
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_HIDUKE <= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sSyuryobi & "' "						'終了日
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU_KBN IN (" & C_KETU_KEKKA & "," & C_KETU_TIKOKU & ","& C_KETU_SOTAI &"," & C_KETU_KEKKA_1 & ")"
	
	'== １年間番号が指定されている場合 ==
	If p_s1NenBango <> "" Then
		w_sSql = w_sSql & vbCrLf & "And "
		w_sSql = w_sSql & vbCrLf & "T21_GAKUSEKI_NO = " & p_s1NenBango & " "
	End If
	
	w_sSql = w_sSql & vbCrLf & "Group By "
	w_sSql = w_sSql & vbCrLf & " T21_CLASS,"
	w_sSql = w_sSql & vbCrLf & " T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & " T21_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "Order By "
	w_sSql = w_sSql & vbCrLf & " T21_CLASS, "
	w_sSql = w_sSql & vbCrLf & " T21_GAKUSEKI_NO "
	
	'response.write(w_sSql)
	
	If gf_GetRecordset(p_oRecordset,w_sSql) <> 0 Then : exit function
	
	gf_GetSyukketuData2 = True
	
End Function

'********************************************************************************
'*  [機能]  学校番号が登録されているかチェックする
'*  [引数]  p_ChkFlg(out),p_Type(in)→[C_KEKKAGAI_DISP,C_HYOKAYOTEI_DISP,C_DATAKBN_DISP]
'*  [戻値]  
'*          gf_ChkDisp(True→正常終了、False→エラー)
'*  [説明]  
'*  		学校ごとに処理が違う際に使用
'*  		p_ChkFlgがTrueなら処理をする
'*  		
'********************************************************************************
function gf_ChkDisp(p_Type,p_ChkFlg)
	Dim w_sSQL
	Dim w_Rs
	Const C_DISP = 1
	
	On Error Resume Next
	Err.Clear
	
	gf_ChkDisp = false
	p_ChkFlg = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "		M00_KANRI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M00_KANRI "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "      M00_NENDO = " & p_Type
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function
	
	if w_Rs.EOF then
		gf_ChkDisp = true
		exit function
	end if
	
	if cint(w_Rs(0)) = C_DISP then p_ChkFlg = true
	
	Call gf_closeObject(w_Rs)
	
	gf_ChkDisp = true
	
end function
'********************************************************************************
'*  [機能]  科目名を取得
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
function gf_GetKamokuMei(p_SyoriNen,p_KamokuCd,p_KamokuKbn)
	Dim w_sSQL,w_Rs
    
	gf_GetKamokuMei = ""
	
	On Error Resume Next
    Err.Clear
	
	'通常授業
	if p_KamokuKbn = C_JIK_JUGYO then
		w_sSQL = ""
		w_sSQL = w_sSQL & "select "
		w_sSQL = w_sSQL & "		M03_KAMOKUMEI "
		w_sSQL = w_sSQL & "from"
		w_sSQL = w_sSQL & "		M03_KAMOKU "
		w_sSQL = w_sSQL & "where "
		w_sSQL = w_sSQL & "		M03_NENDO =" & cint(p_SyoriNen)
		w_sSQL = w_sSQL & "	and	M03_KAMOKU_CD = " & p_KamokuCd
	'特別活動
	else
		w_sSQL = ""
		w_sSQL = w_sSQL & "select "
		w_sSQL = w_sSQL & "		M41_MEISYO "
		w_sSQL = w_sSQL & "from"
		w_sSQL = w_sSQL & "		M41_TOKUKATU "
		w_sSQL = w_sSQL & "where "
		w_sSQL = w_sSQL & "		M41_NENDO =" & cint(p_SyoriNen)
		w_sSQL = w_sSQL & "	and	M41_TOKUKATU_CD = " & p_KamokuCd
	end if
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function
	
	gf_GetKamokuMei = w_Rs(0)
	
end function

'********************************************************************************
'*  [機能]  学校番号取得
'*  [引数]  
'*  [戻値]  p_iGakkoNO:学校番号 レコードがない場合は""を返す
'*  [説明]  true:成功,false:失敗
'********************************************************************************
function gf_GetGakkoNO(p_iGakkoNO)
	Dim w_sSQL
	Dim w_Rs
	
	On Error Resume Next
	Err.Clear
	
	gf_GetGakkoNO = False
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "		M00_KANRI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M00_KANRI "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "     M00_NENDO = " & C_GAKKO_NO
	w_sSQL = w_sSQL & "     AND M00_NO = " & C_DISP_NO

	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

	if w_Rs.EOF then 
		p_iGakkoNO = ""
		gf_GetGakkoNO = True
		exit function
	end if

	p_iGakkoNO = w_Rs("M00_KANRI")
	gf_GetGakkoNO = True
	
end function

'add start 2017/03/01 m.tou
'********************************************************************************
'*  [機能]  操作ログデータINSERT
'*  [引数]  p_nendo：年度／p_syoriId：処理ID／p_syoriName：処理名／p_taisyo：対象／p_sosa：操作／p_userId：ユーザID
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function gf_InsertOpeLog(p_nendo,p_syoriId,p_syoriName,p_taisyo,p_sosa,p_userId)

    Dim w_iRet
    Dim w_sSQL

    On Error Resume Next
    Err.Clear

    gf_InsertOpeLog = 1

    Do

		'//ﾄﾗﾝｻﾞｸｼｮﾝ開始
		Call gs_BeginTrans()

		'//INSERT
		w_sSql = ""
		w_sSql = w_sSql   & " INSERT INTO T79_WEB_LOG"
		w_sSql = w_sSql   & " ("
		w_sSql = w_sSql   & " T79_ID"
		w_sSql = w_sSql   & " ,T79_NENDO"
		w_sSql = w_sSql   & " ,T79_SYORI_ID"
		w_sSql = w_sSql   & " ,T79_SYORI_NAME"
		w_sSql = w_sSql   & " ,T79_TAISYO"
		w_sSql = w_sSql   & " ,T79_SOSA"
		w_sSql = w_sSql   & " ,T79_USER_ID"
		w_sSql = w_sSql   & " ,T79_INS_DATE"
		w_sSql = w_sSql   & " ) VALUES ("
		w_sSql = w_sSql   & " SEQ_T79_ID.NEXTVAL"
		w_sSql = w_sSql   & " ," & p_nendo
		w_sSql = w_sSql   & " ,'" & p_syoriId & "'"
		w_sSql = w_sSql   & " ,'" & p_syoriName & "'"
		w_sSql = w_sSql   & " ,'" & p_taisyo & "'"
		w_sSql = w_sSql   & " ,'" & p_sosa & "'"
		w_sSql = w_sSql   & " ,'" & p_userId & "'"
		w_sSql = w_sSql   & " ,SYSDATE"
		w_sSql = w_sSql   & " )"

    		'response.write w_sSql &"<<br>"

		iRet = gf_ExecuteSQL(w_sSQL)
		If iRet <> 0 Then
			'//ﾛｰﾙﾊﾞｯｸ
			Call gs_RollbackTrans()
			'//登録失敗
			f_Insert = 99
			Exit Do
		End If

        '//ｺﾐｯﾄ
        Call gs_CommitTrans()

        '//正常終了
        gf_InsertOpeLog = 0
        Exit Do
    Loop


End Function
'add end   2017/03/01 m.tou

%>