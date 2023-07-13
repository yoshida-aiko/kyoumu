<%@ Language=VBScript %>
<% Response.Expires = 0%>
<% Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 異動状況一覧
' ﾌﾟﾛｸﾞﾗﾑID : web/web0380/default.asp
' 機      能: 異動状況一覧を出す。
'-------------------------------------------------------------------------
' 引      数:SESSION(""):教官コード     ＞      SESSIONより
' 変      数:なし
' 引      渡:SESSION(""):教官コード     ＞      SESSIONより
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/09/3 谷脇
' 変      更: 2002/02/20 高田
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public m_iNendo
    Public m_rs
    Public m_sZenki_Start		'前期ｽﾀｰﾄ

    Public m_iNowPage			'次のﾍﾟｰｼﾞﾅﾝﾊﾞｰ
	Public m_iPagesize			'表示件数
	m_iPagesize = C_PAGE_LINE


  '********** 表示用配列 **********
    Public m_sSimei()		'氏名
    Public m_sNendo()		'年度    
    Public m_sGakuNo()		'学生番号
    Public m_sGakuseiNo()	'学籍番号
    Public m_sGakunen()		'学年
    Public m_sGakka()		'学科
    Public m_sClass()		'クラス(組)
    Public m_sJiyu()		'異動事由
    Public m_sHiduke()		'日付(開始日）
    Public m_sEHiduke()		'日付(終了日）    
    Public m_sBiko()		'備考

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()
response.end
'///////////////////////////　ＥＮＤ　/////////////////////////////


'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	Dim w_lblMeisyo			'// １年間番号の名称取得用
	
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="異動状況一覧"
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

		'// 権限チェックに使用
'		session("PRJ_No") = "WEB0380"

		'// 不正アクセスチェック
'		Call gf_userChk(session("PRJ_No"))

		'// 変数初期化
		call f_paraSet()

		'// 異動有の学生取得
		If f_GetidoGaku() <> true then
			'データ取得失敗
			m_bErrFlg = True
			m_sErrMsg = "データの取得に失敗しました。"
			Exit Do
		End If
				
        If cint(gf_GetRsCount(m_rs)) = 0 Then
            '異動データがない
	        Call showNoPage()
            Exit Do
        End If

		'異動状況を配列に代入
		If f_InsAry() <> true then
			'データ取得失敗
			m_bErrFlg = True
			m_sErrMsg = "データの取得に失敗しました。"
			Exit Do
		End If
					
		'データのソート
		call s_sortBubble()
						
        '// ページを表示
        Call showPage()
     
        Exit Do
    Loop

    '// 終了処理
    Call gf_closeObject(m_rs)
    Call gs_CloseDatabase()

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
End Sub

'*******************************************************************************
' 機　　能：変数の初期化と代入
' 引　　数：なし
' 機能詳細：
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
Sub f_paraSet()

	m_iNendo = session("NENDO")

	'// 表示ﾍﾟｰｼﾞﾅﾝﾊﾞｰ
	m_iNowPage = Request("hidPageNo")
	if gf_IsNull(m_iNowPage) then
		m_iNowPage = 1
	End if

	m_iNowPage = Cint(m_iNowPage)

End Sub


'*******************************************************************************
' 機　　能：学科グループの取得
' 返　　値：TRUE:OK / FALSE:NG
' 引　　数：p_sGakkaGrp - 学科グループ
' 　　　　　p_sNendo - 年度
' 機能詳細：学科グループの取得
' 備　　考：なし
' 作　　成：2001/07/27　田部
' 変　　更：2001/08/28　谷脇
'*******************************************************************************
Function f_GetidoGaku()

    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt
    
    f_GetidoGaku = False
    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "T11_SIMEI,T13_GAKUSEI_NO,T13_GAKUNEN, "
    w_sSQL = w_sSQL & "T13_GAKUSEKI_NO,T13_NENDO, "    
    w_sSQL = w_sSQL & "M02_GAKKAMEI,M05_CLASSRYAKU as M05_CLASSMEI,T13_IDOU_NUM "
    w_sSQL = w_sSQL & "FROM T13_GAKU_NEN,T11_GAKUSEKI,M02_GAKKA,M05_CLASS "
    w_sSQL = w_sSQL & "WHERE "
    w_sSQL = w_sSQL & " T13_NENDO <= " & m_iNendo & " AND "
    w_sSQL = w_sSQL & " T13_GAKUSEI_NO = T11_GAKUSEI_NO AND "
    w_sSQL = w_sSQL & " T13_NENDO = M02_NENDO AND "
    w_sSQL = w_sSQL & " T13_GAKKA_CD = M02_GAKKA_CD AND "
    w_sSQL = w_sSQL & " T13_NENDO = M05_NENDO AND "
    w_sSQL = w_sSQL & " T13_GAKUNEN = M05_GAKUNEN AND "
    w_sSQL = w_sSQL & " T13_CLASS = M05_CLASSNO AND "
    w_sSQL = w_sSQL & " T13_IDOU_NUM > 0 "
    w_sSQL = w_sSQL & "ORDER BY T13_GAKUSEI_NO ,T13_NENDO"

    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==
    w_iRet = gf_GetRecordset_OpenStatic(m_rs, w_sSQL)
    If w_iRet <> 0 Then
        '== 取得されなかった場合 ==
        Exit function
    End If
    f_GetidoGaku = True
    Exit Function
    
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
Function f_GetZenki_Start(p_iNendo,p_sZenki_Start)

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	p_sZenki_Start = ""

	'管理マスタから学期情報を取得
	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO, "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NO, "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_KANRI, "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_BIKO"
	w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO=" & p_iNendo & " AND "
	w_sSQL = w_sSQL & vbCrLf & "   (M00_KANRI.M00_NO=" & C_K_ZEN_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_SYURYO & ") "  '//[M00_NO]10:前期開始 11:後期開始

	iRet = gf_GetRecordset(rs, w_sSQL)
	If iRet <> 0 Then
	    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
	    m_bErrMsg = Err.description
	    Exit Function
	End If

	If rs.EOF = False Then
	    Do Until rs.EOF
	        If cInt(rs("M00_NO")) = C_K_ZEN_KAISI Then
	            p_sZenki_Start = rs("M00_KANRI")
	        End If
	        rs.MoveNext
	    Loop
	End If

    Call gf_closeObject(rs)

End Function

'*******************************************************************************
' 機　　能：配列にデータを挿入。
' 返　　値：TRUE:OK / FALSE:NG
' 引　　数：
' 機能詳細：配列に表示用のデータを挿入
' 備　　考：なし
' 作　　成：2001/08/28　谷脇
'*******************************************************************************
Function f_InsAry()

    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt
    Dim w_rs
    
    f_InsAry = False

    w_iCnt = 1
    Do Until m_Rs.EOF

	for i = 1 to cint(m_Rs("T13_IDOU_NUM"))

	    w_sSQL = ""
	    w_sSQL = w_sSQL & "SELECT "	    
	    w_sSQL = w_sSQL & "M01_SYOBUNRUIMEI as IDO_KBN,"
	    w_sSQL = w_sSQL & "T13_IDOU_BI_" & i & " as IDO_BI,"
	    w_sSQL = w_sSQL & "T13_IDOU_ENDBI_" & i & " as IDO_BI_E,"
	    w_sSQL = w_sSQL & "T13_IDOU_BIK_" & i & " as IDO_BIK "
	    w_sSQL = w_sSQL & "FROM T13_GAKU_NEN, M01_KUBUN "
	    w_sSQL = w_sSQL & "WHERE "
'	    w_sSQL = w_sSQL & " T13_NENDO <= " & m_iNendo & " AND "
	    w_sSQL = w_sSQL & " T13_NENDO = " & m_rs("T13_NENDO") & " AND "
	    w_sSQL = w_sSQL & " M01_NENDO = " & m_iNendo & " AND "
	    w_sSQL = w_sSQL & " T13_GAKUSEI_NO = '"& m_rs("T13_GAKUSEI_NO") &"' AND "
	    w_sSQL = w_sSQL & " T13_GAKUSEKI_NO = '"& m_rs("T13_GAKUSEKI_NO") &"' AND "	    
	    w_sSQL = w_sSQL & " M01_DAIBUNRUI_CD = '"& C_IDO &"' AND "
	    w_sSQL = w_sSQL & " M01_SYOBUNRUI_CD = T13_IDOU_KBN_"& i &" "

		'== ﾚｺｰﾄﾞｾｯﾄ取得 ==
		w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
		If w_iRet <> 0 Then
			'== 取得されなかった場合 ==
			Exit function
		End If

		redim Preserve m_sSimei(w_iCnt)
		redim Preserve m_sNendo(w_iCnt)	    
		redim Preserve m_sGakuNo(w_iCnt)
		redim Preserve m_sGakuseiNo(w_iCnt)
		redim Preserve m_sGakunen(w_iCnt)
		redim Preserve m_sGakka(w_iCnt)
		redim Preserve m_sClass(w_iCnt)
		redim Preserve m_sJiyu(w_iCnt)
		redim Preserve m_sHiduke(w_iCnt)
		redim Preserve m_sEHiduke(w_iCnt)	    
		redim Preserve m_sBiko(w_iCnt)

		'// 今のﾚｺｰﾄﾞと前のﾚｺｰﾄﾞの「開始日付・異動事由・学生NO」が同じではない場合
		if Cstr(w_rs("IDO_BI") & w_rs("IDO_KBN") & m_rs("T13_GAKUSEI_NO")) <> Cstr(m_sHiduke(w_iCnt-1) & m_sJiyu(w_iCnt-1) & m_sGakuseiNo(w_iCnt-1)) then

			'// 前期開始日を取得
			Call f_GetZenki_Start(Cint(left(w_rs("IDO_BI"),4)),m_sZenki_Start)

			'// 処理年度の次年度の前期開始日データが存在しない場合、4/1をセットする。
            if m_sZenki_Start = "" AND Cint(left(w_rs("IDO_BI"),4)) > cint(m_inendo) then
                  m_sZenki_Start = m_inendo & "/04/01"
    		end if

			'// 前期開始日より開始日が前だったら、年度を-1する
			if right(gf_YYYY_MM_DD(m_sZenki_Start,"/"),5) > right(gf_YYYY_MM_DD(w_rs("IDO_BI"),"/"),5) then
			    m_sNendo(w_iCnt)	= Cint(left(w_rs("IDO_BI"),4)) - 1
			Else
			    m_sNendo(w_iCnt)	= left(w_rs("IDO_BI"),4)
			End if

			'// 配列にｾｯﾄする
		    m_sSimei(w_iCnt)		= m_rs("T11_SIMEI")
		    m_sGakunen(w_iCnt)		= m_rs("T13_GAKUNEN")
		    m_sGakka(w_iCnt)		= m_rs("M02_GAKKAMEI")
		    m_sClass(w_iCnt)		= m_rs("M05_CLASSMEI")
		    m_sGakuNo(w_iCnt)		= m_rs("T13_GAKUSEKI_NO")
			m_sGakuseiNo(w_iCnt)    = m_rs("T13_GAKUSEI_NO")

		    m_sJiyu(w_iCnt)			= w_rs("IDO_KBN")
		    m_sHiduke(w_iCnt)		= w_rs("IDO_BI")
		    m_sEHiduke(w_iCnt)		= w_rs("IDO_BI_E")
		    m_sBiko(w_iCnt)			= w_rs("IDO_BIK")

			w_iCnt = w_iCnt +1
		End if

	    Call gf_closeObject(w_rs)

	next
	m_rs.MoveNext
    loop
	
	'//ﾚｺｰﾄﾞｾｯﾄCLOSE

    f_InsAry = True
    Exit Function
    
End Function


'*******************************************************************************
' 機　　能：バブルソート
' 返　　値：TRUE:OK / FALSE:NG
' 引　　数：
' 機能詳細：日付でソートします。
' 備　　考：なし
' 作　　成：2001/08/28　谷脇
' 変　　更：2002/02/20　高田：日付・学籍番号順とする
'*******************************************************************************
sub s_sortBubble() 

    Dim i
    Dim j
	Dim loopc
	    
    For i = 0 To UBound(m_sHiduke) - 1
        For j = i + 1 To UBound(m_sHiduke)
            If m_sHiduke(j) > m_sHiduke(i) Then
                Call s_swap(m_sSimei(i),m_sSimei(j))
                Call s_swap(m_sNendo(i),m_sNendo(j))                
                Call s_swap(m_sGakuNo(i),m_sGakuNo(j))
                Call s_swap(m_sGakunen(i),m_sGakunen(j))
                Call s_swap(m_sGakka(i),m_sGakka(j))
                Call s_swap(m_sClass(i),m_sClass(j))
                Call s_swap(m_sJiyu(i),m_sJiyu(j))
                Call s_swap(m_sHiduke(i),m_sHiduke(j))
                Call s_swap(m_sEHiduke(i),m_sEHiduke(j))                
                Call s_swap(m_sBiko(i),m_sBiko(j))
            End If          
        Next
    Next


    
    For i = 0 To UBound(m_sHiduke) - 1
        For j = i + 1 To UBound(m_sHiduke)
            If m_sHiduke(j) = m_sHiduke(i) Then 
				For loopc = i  To UBound(m_sHiduke)    
					If m_sHiduke(i) = m_sHiduke(loopc) Then
						If m_sGakuNo(i) > m_sGakuNo(loopc) Then
						    Call s_swap(m_sSimei(i),m_sSimei(loopc))
						    Call s_swap(m_sNendo(i),m_sNendo(loopc))
						    Call s_swap(m_sGakuNo(i),m_sGakuNo(loopc))
						    Call s_swap(m_sGakunen(i),m_sGakunen(loopc))
						    Call s_swap(m_sGakka(i),m_sGakka(loopc))
						    Call s_swap(m_sClass(i),m_sClass(loopc))
						    Call s_swap(m_sJiyu(i),m_sJiyu(loopc))
						    Call s_swap(m_sHiduke(i),m_sHiduke(loopc))
						    Call s_swap(m_sEHiduke(i),m_sEHiduke(loopc))                
						    Call s_swap(m_sBiko(i),m_sBiko(loopc))
						End If	 						
					End If
				Next				
            End If            
        Next
    Next    
End Sub


'*******************************************************************************
' 機　　能：スワップ
' 返　　値：TRUE:OK / FALSE:NG
' 引　　数：Ａ，Ｂ
' 機能詳細：Ａ，Ｂの中身を入れ替える。
' 備　　考：なし
' 作　　成：2001/08/28　谷脇
'*******************************************************************************
Sub s_swap(a,b) 

dim tmp
	tmp = a
	a = b
	b = tmp
End sub


'********************************************************************************
'*  [機能]  ページ関係の表示用サブルーチン
'*  [引数]  p_iRsCnt        ：ﾚｺｰﾄﾞｶｳﾝﾄ
'*          p_iPageCd       ：ページ番号
'*          p_iDsp          ：1ページの最大表示すう。
'*  [戻値]  p_pageBar       ：できたページバーHTML
'*  [説明]  
'********************************************************************************
Sub s_pageBar(p_iRsCnt,p_iPageCd,p_iDsp)

	Dim w_bNxt					'// NEXT表示有無
	Dim w_bBfr					'// BEFORE表示有無
	Dim w_iNxt					'// NEXT表示頁数
	Dim w_iBfr					'// BEFORE表示頁数
	Dim w_iCnt					'// ﾃﾞｰﾀ表示ｶｳﾝﾀ
	Dim w_iMax					'// ﾃﾞｰﾀ表示ｶｳﾝﾀ
	Dim i,w_iSt,w_iEd

	Dim w_iRecordCnt			'//レコードセットカウント

	On Error Resume Next
	Err.Clear

	w_iCnt = 1
	w_bFlg = True

	'////////////////////////////////////////
	'ページ関係の設定
	'////////////////////////////////////////

	'レコード数を取得
	w_iRecordCnt = p_iRsCnt
	w_iMax = int((Cint(p_iRsCnt) / p_iDsp) + 0.9)

	'EOFのときの設定
	If Cint(p_iPageCd) >= w_iMax Then
		p_iPageCd = w_iMax
	End If

	'前ページの設定
	If Cint(p_iPageCd) = 1 Then
		w_bBfr = False
		w_iBfr = 0
	Else
		w_bBfr = True
		w_iBfr = Cint(p_iPageCd) - 1
	End If

	'後ページの設定
	If Cint(p_iPageCd) = w_iMax Then
		w_bNxt = False
		w_iNxt = Cint(p_iPageCd)
	Else
		w_bNxt = True
		w_iNxt = Cint(p_iPageCd) + 1
	End If

	'ページのリストの始め(w_iSt)と終わり(w_iEd)を代入
	'基本的に選択されているページ(p_iPageCd)が真中に来るようにする。
	w_iEd = Cint(p_iPageCd) + 5
	w_iSt = Cint(p_iPageCd) - 4

	'ページのリストが10個ない時、選択ページがリストの真中にこないとき。
	If Cint(p_iPageCd) < 5 Then w_iEd = 10
	If w_iEd > w_iMax then w_iEd = w_iMax : w_iSt = w_iMax - 9
	If w_iSt < 1 or w_iMax < 10 then w_iSt = 1

	'////////////////////////////////////////
	'ページ関係の設定(ここまで)
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
		If i = Cint(p_iPageCd) then 
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

	'// 書き出し
	response.write p_pageBar

End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  2003.08.07:高田:loopの条件が＝を含んでいた為に最終データが表示できなかった
'********************************************************************************
sub s_writecell()
	dim i,w_cell
	Dim w_sMaeSimei

	'// ﾙｰﾌﾟｶｳﾝﾀｰの初期値を取得
	i = ((m_iNowPage - 1) * m_iPagesize)

	'// ﾙｰﾌﾟｶｳﾝﾀｰの最大ﾙｰﾌﾟ数
	iMax = (m_iNowPage * m_iPagesize)

	Do Until i > (Ubound(m_sHiduke)-1) or i >= iMax
		call gs_cellPtn(w_cell)
		 %>
		<TR>
			<TD class="<%=w_cell%>"><%=m_sNendo(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sHiduke(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sEHiduke(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sJiyu(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sSimei(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sGakuNo(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sGakunen(i)%>-<%=m_sClass(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sGakka(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sBiko(i)%></TD>
		</TR>
		<%
		i = i + 1
	Loop

End sub

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showNoPage()
%>
<html>
<head>
    <title>異動状況一覧</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
</head>
<body>
<center>
<%call gs_title("異動状況一覧","一　覧")%>
<BR>
<span class="msg">現在、異動者はいません。</span>
</center>
</body>
</html>
<%
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
<title>異動状況一覧</title>
<link rel=stylesheet href=../../common/style.css type=text/css>
<SCRIPT LANGUAGE="javascript">
<!--

	//************************************************************
	//  [機能]  一覧表の次・前ページを表示する
	//  [引数]  p_iPage :表示頁数
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_PageClick(p_iPage){

		document.frm.action = "default.asp";
		document.frm.target = "fTopMain";
		document.frm.hidPageNo.value = p_iPage;
		document.frm.submit();

	}
//-->
</SCRIPT>
</head>
<body>
<center>
<form name="frm" method="post">
<%call gs_title("異動状況一覧","一　覧")%>

<BR>

<table class="hyo" border="1" width="">
	<tr>
		<th nowrap class="header">異動状況一覧</th>
		<td nowrap class="detail" align="center"><%=gf_fmtWareki(date())%>現在</td>
	</tr>
</table>

<BR>
<table border=0 cellpaddin="0" cellspacing="0" width="98%">
	<tr><td><% Call s_pageBar((Ubound(m_sHiduke)-1),m_iNowPage,m_iPagesize) %></td></tr>
	<tr>
		<td>
			<table border=1 class="hyo" width="100%">
				<tr>
					<th class="header" width="5%" nowrap>年度</th>
					<th class="header" width="10%" nowrap>開始日付</th>
					<th class="header" width="10%" nowrap>終了日付</th>
					<th class="header" nowrap>異動事由</th>
					<th class="header" nowrap>氏名</th>
					<th class="header" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
					<th class="header" nowrap>クラス</th>
					<th class="header" nowrap>学科</th>
					<th class="header" width="15%">備考</th>
				</tr>
				<% call s_writecell() %>
			</table>
		</td>
	</tr>
	<tr><td><% Call s_pageBar((Ubound(m_sHiduke)-1),m_iNowPage,m_iPagesize) %></td></tr>
</table>

<input type="hidden" name="hidPageNo" value="<%= m_iNowPage %>">
</from>

</center>
</body>
</head>
</html>
<%
End Sub
%>