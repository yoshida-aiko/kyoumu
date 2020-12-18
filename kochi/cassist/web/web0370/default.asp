<%@ Language=VBScript %>
<%Response.Expires = 0%>
<%Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生数一覧
' ﾌﾟﾛｸﾞﾗﾑID : web/web0370/default.asp
' 機      能: 学生数の一覧を出す。
'-------------------------------------------------------------------------
' 引      数:SESSION(""):教官コード     ＞      SESSIONより
' 変      数:なし
' 引      渡:SESSION(""):教官コード     ＞      SESSIONより
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2001/08/29 谷脇
' 変      更: 2015/08/27 藤林 変更内容(学科名取得時の年度、クラスを先頭1桁のみ使用する、混合クラス人数で、0人のクラスは表示しない)
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public m_iNendo
    Public m_iGrp_su
    Public m_ikongoCls()		'混合クラスフラグ
    Public m_sGakkaGrp()	'学科グループ
    Public m_gakka_cd()	'学科コード
    Public m_gakkamei()	'学科名
    Public m_Qgakkamei()	'旧学科名
    Public m_Fld_M()	'全体数
    Public m_Fld_F()	'全体数（女子）
    Public m_Fld_R()	'全体数（留学生）
    Public m_Fld_MK()	'休学者数
    Public m_Fld_FK()	'休学者数（女子）
    Public m_Fld_RK()	'休学者数（留学生）
    Public m_Cls_M()	'混合クラス
    Public m_Cls_F()		'混合クラス（女子）

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()
response.end
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
    w_sMsgTitle="学生数一覧"
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
'		session("PRJ_No") = "WEB0370"

		'// 不正アクセスチェック
'		Call gf_userChk(session("PRJ_No"))

		'// 変数初期化
		call f_paraSet()
		
		'// 学科グループ取得
		If f_GetGakkaGrp() <> true then
	            'データ取得失敗
	            m_bErrFlg = True
	            m_sErrMsg = "学科データがありません"
	            Exit Do
		End If
		
        If m_iGrp_su = 0 Then
            '学科データがない
            m_bErrFlg = True
            m_sErrMsg = "学科データがありません"
            Exit Do
        End If

		call f_arySet()
		
		'// 混合クラス取得
		If f_GetClass() <> true then
	            'データ取得失敗
	            m_bErrFlg = True
	            m_sErrMsg = "混合クラスデータがありません"
	            Exit Do
		End If
		
		'// データの集計
		for i = 1 to m_iGrp_su
			if f_GetGakusei(i) <> true then
				'データ取得失敗
				m_bErrFlg = True
				m_sErrMsg = "データがありません。"
				Exit for
			End If
		next
		
		'// 学科名取得
		If f_GetGakkaMei() <> true then
	            'データ取得失敗
	            m_bErrFlg = True
	            m_sErrMsg = "学科名がありません。"
	            Exit Do
		End If
		
        '// ページを表示
        Call showPage()
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
		response.write w_sMsg
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gs_CloseDatabase()
End Sub

Sub f_paraSet()
'*******************************************************************************
' 機　　能：変数の初期化と代入
' 引　　数：なし
' 機能詳細：
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
m_iNendo = session("NENDO")
'm_iNendo = 2001

End Sub

Sub f_arySet()
'*******************************************************************************
' 機　　能：変数の初期化と代入
' 引　　数：なし
' 機能詳細：
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
Dim i,j
Redim m_gakka_cd(m_iGrp_su)	'学科コード
Redim m_gakkamei(m_iGrp_su)	'学科名
Redim m_Qgakkamei(m_iGrp_su)	'旧学科名
Redim m_Fld_M(6,m_iGrp_su)	'全体数
Redim m_Fld_F(6,m_iGrp_su)	'全体数（女子）
Redim m_Fld_R(6,m_iGrp_su)	'全体数（留学生）
Redim m_Fld_MK(6,m_iGrp_su)	'休学者数
Redim m_Fld_FK(6,m_iGrp_su)	'休学者数（女子）
Redim m_Fld_RK(6,m_iGrp_su)	'休学者数（留学生）

'/*　配列の初期化　*/
for j = 0 to m_iGrp_su
	m_gakka_cd(j) = m_sGakkaGrp(j)
	m_gakkamei(j) = ""
	m_Qgakkamei(j) = ""
	for i = 0 to 6 
		m_Fld_M(i,j) = 0
		m_Fld_F(i,j) = 0
		m_Fld_R(i,j) = 0
		m_Fld_MK(i,j) = 0
		m_Fld_FK(i,j) = 0
		m_Fld_RK(i,j) = 0
	next 
next

End Sub

Function f_GetGakkaGrp()
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
    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt
    Dim w_rs
    
    f_GetGakkaGrp = False
    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M23_GROUP "
    w_sSQL = w_sSQL & "FROM M23_GAKKA_GRP "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M23_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M23_GAKKA_CD IS NOT NULL "
    w_sSQL = w_sSQL & "Group By M23_GROUP "
    w_sSQL = w_sSQL & "Order By M23_GROUP "
    w_sSQL = w_sSQL & ""

    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==
    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
    If w_iRet <> 0 Then
        '== 取得されなかった場合 ==
        Exit function
    End If
    
    If w_rs.eof = True Then
        m_iGrp_su = 0
    
    Else
        '== ﾚｺｰﾄﾞ件数の取得 ==
        m_iGrp_su = cint(gf_GetRsCount(w_rs))
'        m_iGrp_su = w_rs.RecordCount
    End If
    ReDim m_sGakkaGrp(m_iGrp_su)
   
    '== 学科グループのデータをセットする ==
    For w_iCnt = 1 to m_iGrp_su
        m_sGakkaGrp(w_iCnt) = w_rs("M23_GROUP")
        w_rs.MoveNext
    Next
    
	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_rs)

    f_GetGakkaGrp = True
    Exit Function
    
End Function

Function f_GetClass()
'*******************************************************************************
' 機　　能：混合クラスを探す。
' 返　　値：TRUE:OK / FALSE:NG
' 引　　数：
' 機能詳細：混合クラスを配列に入れる
' 備　　考：なし
' 作　　成：2001/08/28　谷脇
'*******************************************************************************
    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt,w_max_cls
    Dim w_rs
    
    f_GetClass = False
    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M05_GAKUNEN,Max(M05_CLASSNO) as MAX_CLS,M05_SYUBETU "
    w_sSQL = w_sSQL & "FROM M05_CLASS "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_GAKUNEN IS NOT NULL "
    w_sSQL = w_sSQL & "Group By M05_GAKUNEN,M05_SYUBETU "
    w_sSQL = w_sSQL & ""

    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==
    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
   If w_iRet <> 0 Then
        '== 取得されなかった場合 ==
        Exit function
    End If
    
    If w_rs.eof = True Then
        w_iCnt = 0
    
    Else
        '== ﾚｺｰﾄﾞ件数の取得 ==
        w_iCnt = cint(gf_GetRsCount(w_rs))
'        m_iGrp_su = w_rs.RecordCount
    End If

    '混合クラスフラグの配列の初期化
    ReDim m_ikongoCls(w_iCnt)
   for each i in m_ikongoCls
    i=0
     next
    '== 混合クラスのフラグをセットする ==
    w_max_cls = 0
    Do Until w_rs.EOF
	w_igak = cint(w_rs("M05_GAKUNEN"))
        m_ikongoCls(w_igak) = w_rs("M05_SYUBETU")

	'１学年でも混合クラスがあれば、フラグを立てる。
	If cint(w_rs("M05_SYUBETU")) = 1 then m_ikongoCls(0) = 1

	If w_max_cls < cint(w_rs("MAX_CLS")) then 
		w_max_cls = cint(w_rs("MAX_CLS"))
	End If

        w_rs.MoveNext
    Loop

	'集計用配列の初期化
	ReDim m_Cls_M(w_iCnt,w_max_cls)
	ReDim m_Cls_F(w_iCnt,w_max_cls)
    
	    For i = 0 to w_iCnt
		    For j = 0 to w_max_cls
			m_Cls_M(i,j) = 0
			m_Cls_F(i,j) = 0
		    Next
	    Next

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_rs)


    f_GetClass = True
    Exit Function
    
End Function

Function f_GetClassMei(p_gak,p_cls)
'*******************************************************************************
' 機　　能：混合クラス名取得。
' 返　　値：TRUE:OK / FALSE:NG
' 引　　数：クラス名
' 機能詳細：混合クラスのクラス名を出す。
' 備　　考：なし
' 作　　成：2001/08/28　谷脇
'*******************************************************************************
    Dim w_sSQL
    Dim w_iRet
    Dim w_rs
    
    f_GetClassMei = ""
    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M05_CLASSMEI "
    w_sSQL = w_sSQL & "FROM M05_CLASS "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M05_NENDO = " & m_iNendo & " AND "
    w_sSQL = w_sSQL & "M05_GAKUNEN = " & p_gak & " AND "
    w_sSQL = w_sSQL & "M05_CLASSNO = " & p_cls & "1 "
    w_sSQL = w_sSQL & ""

    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==
    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
   If w_iRet <> 0 Then
        '== 取得されなかった場合 ==
        Exit function
    End If
    If w_rs.eof = True Then
        f_GetClassMei = ""
    
    Else
	f_GetClassMei = w_rs("M05_CLASSMEI")
    End If
	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_rs)
End Function

Function f_GetGakkaMei()
'*******************************************************************************
' 機　　能：学科名のセット
' 返　　値：TRUE:OK / FALSE:NG
' 引　　数：p_sGakkaGrp - 学科グループ
' 　　　　　p_sNendo - 年度
' 機能詳細：学科グループの取得
' 備　　考：なし
' 作　　成：2001/07/27　田部
' 変　　更：2001/08/28　谷脇
'*******************************************************************************
    Dim w_sSQL
    Dim w_iRet
    Dim w_rs
    Dim w_grp,w_gakka,w_gakkamei,w_gakunen

    w_grp = 0:w_gakka=0:w_gakkamei="":w_gakunen=""

    f_GetGakkaMei = false
    
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M23_GROUP,"
    w_sSQL = w_sSQL & "M02_GAKKA_CD, "
    w_sSQL = w_sSQL & "M02_GAKKAMEI, "
    w_sSQL = w_sSQL & "M23_GAKUNEN "
    w_sSQL = w_sSQL & "FROM "
    w_sSQL = w_sSQL & "M02_GAKKA, M23_GAKKA_GRP "
    w_sSQL = w_sSQL & "Where "
'    w_sSQL = w_sSQL & "M23_NENDO = 2000 And "
    w_sSQL = w_sSQL & "M23_NENDO = " & m_iNendo & " And "
    w_sSQL = w_sSQL & "M02_NENDO = M23_NENDO And "
    w_sSQL = w_sSQL & "M02_GAKKA_CD = M23_GAKKA_CD "
    w_sSQL = w_sSQL & "order by M23_GROUP,M23_GAKUNEN,M02_GAKKA_CD"

    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==
    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
    If w_iRet <> 0 Then
        '== 取得されなかった場合 ==
        Exit function
    End If

    If w_rs.EOF = true then 
	m_bErrFlg = True
	m_sErrMsg = "学科データがありません"
	Exit Function
    End If 

   w_rs.MoveFirst
   Do Until w_rs.EOF
     w_grp = cint(w_rs("M23_GROUP"))
     
     '*** 学科グループが同じだら、学科コードが違うとき ***
     if w_group = cint(w_rs("M23_GROUP")) and w_gakka <> cint(w_rs("M02_GAKKA_CD")) then 

	     '*** 学年が大きい方が旧学科 ***
		If w_gakunen > cint(w_rs("M23_GAKUNEN")) then 
			m_gakkamei(w_grp) = w_rs("M02_GAKKAMEI")
			m_Qgakkamei(w_grp) = w_gakkamei
		else 
			m_gakkamei(w_grp) = w_gakkamei
			m_Qgakkamei(w_grp) = w_rs("M02_GAKKAMEI")
		End If 

     '*** 学科グループが違ったら、学科名の配列にレコードを入れる。 ***
     Elseif w_group <> cint(w_rs("M23_GROUP")) then
			m_gakkamei(w_grp) = w_rs("M02_GAKKAMEI")
			m_Qgakkamei(w_grp) = ""
     End If
     w_group = cint(w_rs("M23_GROUP"))
     w_gakkamei = w_rs("M02_GAKKAMEI")
     w_gakka = cint(w_rs("M02_GAKKA_CD"))
     w_gakunen = cint(w_rs("M23_GAKUNEN"))
	w_rs.MoveNext
   loop
   
     '*** 学年毎合計の名前 ***
	m_gakkamei(0) = "学年別　合計"
	m_Qgakkamei(0) = ""

    f_GetGakkaMei = true
End Function


Function f_GetGakusei(p_cnt)
'*******************************************************************************
' 機　　能：学科別の学生一覧を取得＆集計
' 返　　値：TRUE:OK / FALSE:NG
' 引　　数：p_Gakka - 学科コード
' 機能詳細：学科に所属する学生の一覧を取得
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt
    Dim w_Rs
    Dim w_iGak,w_iCls,w_iSeb,w_iNyu,w_iZai
	
    f_GetGakusei = False
    
    p_sGrp = m_sGakkaGrp(p_cnt)
	
    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "T13_GAKUNEN,SUBSTR(T13_CLASS,1,1) AS T13_CLASS, T13_ZAISEKI_KBN,T11_SEIBETU,T11_NYUGAKU_KBN,T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & "FROM T13_GAKU_NEN,T11_GAKUSEKI "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "T13_NENDO = "& m_iNENDO &" AND "
    w_sSQL = w_sSQL & "T13_GAKUSEI_NO = T11_GAkUSEI_NO AND "
'    w_sSQL = w_sSQL & "T11_NYUNENDO = ("&m_iNENDO&" - T13_GAKUNEN + 1) AND "
    w_sSQL = w_sSQL & "T13_ZAISEKI_KBN <= " & C_ZAI_TEIGAKU & " AND "
    w_sSQL = w_sSQL & "T13_GAKKA_CD IN "
    w_sSQL = w_sSQL & "    (select M23_GAKKA_CD from M23_GAKKA_GRP where M23_GROUP ='"& p_sGrp &"' and M23_NENDO = "& m_iNENDO &") "
    w_sSQL = w_sSQL & ""
    

    If gf_GetRecordset_OpenStatic(w_rs, w_sSQL) <> 0 Then
        '== 取得されなかった場合 ==
        Exit Function
    End If
	
    w_rs.MoveFirst

    
    Do Until w_rs.EOF
	    '// 変数に代入。nullの時は、０を代入する。
	    w_iGak = cint(gf_SetNull2Zero(w_rs("T13_GAKUNEN")))
	    w_iCls = cint(gf_SetNull2Zero(w_rs("T13_CLASS")))
	    w_iSeb = cint(gf_SetNull2Zero(w_rs("T11_SEIBETU")))
	    w_iNyu = cint(gf_SetNull2Zero(w_rs("T11_NYUGAKU_KBN")))
	    w_iZai = cint(gf_SetNull2Zero(w_rs("T13_ZAISEKI_KBN")))
		
		
'		if w_iGak >0 and w_iGak <= 6 and w_iCls > 0 and w_iCls <= m_iGrp_su then 
		if w_iGak >0 and w_iGak <= 6 and w_iCls > 0  then 
			'学年の全体数に加算
			m_Fld_M(w_iGak,p_cnt) = m_Fld_M(w_iGak,p_cnt) + 1
			
			'学年（女子）の全体数に加算
			If w_iSeb = C_SEIBETU_F then 
				m_Fld_F(w_iGak,p_cnt) = m_Fld_F(w_iGak,p_cnt) + 1
			End If 

			'留学生の全体数に加算　３年生以上
	'		If w_iNyu = C_NYU_RYUGAKU and w_iGak > 2 Then 
	'			m_Fld_R(w_iGak,p_cnt) = m_Fld_R(w_iGak,p_cnt) + 1
	'		End If

			'休学生の全体数に加算
			If w_iZai = C_ZAI_KYUGAKU Then 
				m_Fld_MK(w_iGak,p_cnt) = m_Fld_MK(w_iGak,p_cnt) + 1

				'休学生（女子）に加算
			    If w_iSeb = C_SEIBETU_F Then 
					m_Fld_FK(w_iGak,p_cnt) = m_Fld_FK(w_iGak,p_cnt) + 1
			    End If 
				
				'留学生（休学）に加算　３年生以上
	'		    If w_iNyu = C_NYU_RYUGAKU and w_iGak > 2 Then 
	'				m_Fld_RK(w_iGak,p_cnt) = m_Fld_RK(w_iGak,p_cnt) + 1
	'		    End If
			End If 
			
			'対象学年が混合クラスの場合、クラス別に集計を取る。
			If cint(m_ikongoCls(w_iGak)) = C_CLASS_KONGO then
					m_Cls_M(w_iGak,w_iCls) = m_Cls_M(w_iGak,w_iCls) + 1
					'女性の集計
					If w_iSeb = C_SEIBETU_F then 
						m_Cls_F(w_iGak,w_iCls) = m_Cls_F(w_iGak,w_iCls) + 1
					End If
			End If 
		end if

		w_rs.MoveNext

    Loop

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_rs)

'留学生の場合。

    '== SQL作成 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "T13_GAKUNEN,SUBSTR(T13_CLASS,1,1) AS T13_CLASS, T13_ZAISEKI_KBN,T11_SEIBETU,T11_NYUGAKU_KBN,T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & "FROM T13_GAKU_NEN,T11_GAKUSEKI "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "T13_NENDO = "& m_iNENDO &" AND "
    w_sSQL = w_sSQL & "T13_GAKUSEI_NO = T11_GAkUSEI_NO AND "
'    w_sSQL = w_sSQL & "T11_NYUNENDO = ("&m_iNENDO&" - T13_GAKUNEN + 1) AND "
    w_sSQL = w_sSQL & "T13_ZAISEKI_KBN <= " & C_ZAI_TEIGAKU & " AND "
    w_sSQL = w_sSQL & "T11_NYUGAKU_KBN = " & C_NYU_RYUGAKU & " AND "
    w_sSQL = w_sSQL & "T13_GAKKA_CD IN "
    w_sSQL = w_sSQL & "    (select M23_GAKKA_CD from M23_GAKKA_GRP where M23_GROUP ='"& p_sGrp &"' and M23_NENDO = "& m_iNENDO &") "
    w_sSQL = w_sSQL & ""
    '== ﾚｺｰﾄﾞｾｯﾄ取得 ==

'    response.write w_sSQL&"<BR>"
'    response.end


    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
    If w_iRet <> 0 Then
        '== 取得されなかった場合 ==
        Exit Function
    End If
'    w_iGak = cint(w_rs("T13_GAKUNEN"))
if w_rs.EOF = false then 
    w_rs.MoveFirst
    Do Until w_rs.EOF
    '// 変数に代入。nullの時は、０を代入する。
    w_iGak = cint(gf_SetNull2Zero(w_rs("T13_GAKUNEN")))
    w_iCls = cint(gf_SetNull2Zero(w_rs("T13_CLASS")))
    w_iSeb = cint(gf_SetNull2Zero(w_rs("T11_SEIBETU")))
    w_iNyu = cint(gf_SetNull2Zero(w_rs("T11_NYUGAKU_KBN")))
    w_iZai = cint(gf_SetNull2Zero(w_rs("T13_ZAISEKI_KBN")))

'	if w_iGak >2 and w_iGak <= 6 and w_iCls > 0 and w_iCls <= m_iGrp_su then 
	if w_iGak >2 and w_iGak <= 6 and w_iCls > 0  then 
		'留学生の全体数に加算　３年生以上
			m_Fld_R(w_iGak,p_cnt) = m_Fld_R(w_iGak,p_cnt) + 1

		'留学生（休学）に加算　３年生以上
		    If w_iZai = C_ZAI_KYUGAKU Then 
				m_Fld_RK(w_iGak,p_cnt) = m_Fld_RK(w_iGak,p_cnt) + 1
		    End If

	end if
	  w_rs.MoveNext
    Loop
end if
	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(w_rs)

    f_GetGakusei = True
    Exit Function
    
End Function

Sub s_writeSum(p_grp)
'*******************************************************************************
' 機　　能：集計値をテーブルに書き出す。
' 返　　値：
' 引　　数：
' 機能詳細：
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
dim i,w_sCell
dim w_mFld_M,w_mFld_F,w_mFld_R
w_mFld_M = ""
w_mFld_F = ""
w_mFld_R = ""
w_sCell = "CELL2"
%>
<tr>
<td class="<%=w_sCell%>"><%=m_gakkamei(p_grp)%>
<% if m_Qgakkamei(p_grp) <> "" then %>
<br>(<%=m_Qgakkamei(p_grp)%>)
<% End If %>
</td>

<% for i = 1 to 6  '学年毎にセルの書き込み
call gs_cellPtn(w_sCell)

'//*** セルの値を変数に代入。（休学者がいる場合は、カッコ書きも加える）
'// 学生数
if cint(m_Fld_MK(i,p_grp)) > 0 then w_mFld_M = "("&m_Fld_MK(i,p_grp)&")"
w_mFld_M = w_mFld_M & "<br>"&m_Fld_M(i,p_grp)

'// 学生数（女子）
if m_Fld_FK(i,p_grp) > 0 then w_mFld_F = "("&m_Fld_FK(i,p_grp)&")"
w_mFld_F = w_mFld_F & "<br>"&m_Fld_F(i,p_grp)

'// 留学生数
if m_Fld_RK(i,p_grp) > 0 then w_mFld_R = "("&m_Fld_RK(i,p_grp)&")"
w_mFld_R = w_mFld_R  & "<br>"&m_Fld_R(i,p_grp)

%>
<td class="<%=w_sCell%>" align="right" nowrap><%=w_mFld_M%></td>
<td class="<%=w_sCell%>" align="right" nowrap><%=w_mFld_F%></td>
<% If i >= 3 then %>
	<td class="<%=w_sCell%>" align="right" nowrap><%=w_mFld_R%></td>
<% End If %>

<%
w_mFld_M = ""
w_mFld_F = ""
w_mFld_R = ""


next

End Sub

Sub s_writeCell()
'*******************************************************************************
' 機　　能：学科毎、学年毎の集計を出す。
' 返　　値：
' 引　　数：
' 機能詳細：
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
dim w_igrp,w_igak
		for w_igrp = 1 to m_iGrp_su
			for w_igak = 1 to 6
				if w_igak < 6 then
					'学科毎合計
					m_Fld_M(6,w_igrp)   = m_Fld_M(6,w_igrp)   + m_Fld_M(w_igak,w_igrp)
					m_Fld_MK(6,w_igrp) = m_Fld_MK(6,w_igrp) + m_Fld_MK(w_igak,w_igrp)
					m_Fld_F(6,w_igrp)    = m_Fld_F(6,w_igrp)   + m_Fld_F(w_igak,w_igrp)
					m_Fld_FK(6,w_igrp)  = m_Fld_FK(6,w_igrp) + m_Fld_FK(w_igak,w_igrp)
					m_Fld_R(6,w_igrp)   = m_Fld_R(6,w_igrp)    + m_Fld_R(w_igak,w_igrp)
					m_Fld_RK(6,w_igrp) = m_Fld_RK(6,w_igrp)  + m_Fld_RK(w_igak,w_igrp) 
				end if
					'学年毎合計
					m_Fld_M(w_igak,0)  = m_Fld_M(w_igak,0)   + m_Fld_M(w_igak,w_igrp)
					m_Fld_MK(w_igak,0) = m_Fld_MK(w_igak,0) + m_Fld_MK(w_igak,w_igrp)
					m_Fld_F(w_igak,0)   = m_Fld_F(w_igak,0)    + m_Fld_F(w_igak,w_igrp)
					m_Fld_FK(w_igak,0) = m_Fld_FK(w_igak,0)  + m_Fld_FK(w_igak,w_igrp)
					m_Fld_R(w_igak,0)   = m_Fld_R(w_igak,0)    + m_Fld_R(w_igak,w_igrp)
					m_Fld_RK(w_igak,0) = m_Fld_RK(w_igak,0)  + m_Fld_RK(w_igak,w_igrp) 
			next
			call s_writeSum(w_igrp)
		next
			call s_writeSum(0)
End Sub

Sub s_kongoWrite()
'*******************************************************************************
' 機　　能：混合クラスの表を出す。
' 返　　値：
' 引　　数：
' 機能詳細：
' 備　　考：なし
' 作　　成：2001/08/29　谷脇
'*******************************************************************************
dim w_sCell
%>
	        <table class="hyo" border="1" width="">
	            <tr>

					
	                <th nowrap class="header">混合クラス</th>
	                <td nowrap class="detail" align="center"><%=gf_fmtWareki(date())%>現在</td>

	            </tr>
	        </table>
<BR>
<table border=1 class="hyo" >

<%
for j = 0 to UBound(m_Cls_M,2) 
	If j = 0 Then 'ヘッダの書き込み
%>
		<tr>
		<th class="header">クラス</th>
		<%
		for i = 1 to UBound(m_Cls_M)
			if cint(m_ikongoCls(i)) = C_CLASS_KONGO then 
			%>
			 <th class="header"><%=i%>年</th>
			 <th class="header">女子</th>
			<%

			end if 
		next
		%>
		</tr>
		<%
	Else
		iNinzu = 0
		for i = 1 to UBound(m_Cls_M) 		
			 iNinzu = iNinzu + m_Cls_M(i,j)
		next
		
		If iNinzu > 0 Then
			call gs_cellPtn(w_sCell)

%>
			<tr>
			<td class="<%=w_sCell%>"><%=f_GetClassMei(1,j)%></th>
			<% for i = 1 to UBound(m_Cls_M) 
			call gs_cellPtn(w_sCell)

				if cint(m_ikongoCls(i)) = C_CLASS_KONGO then 
				%>
				 <td class="<%=w_sCell%>" align="right"><%=m_Cls_M(i,j)%></td>
				 <td class="<%=w_sCell%>" align="right"><%=m_Cls_F(i,j)%></td>
				<%

				end if 
			next
			%>
			</tr>
		<%
		End If
	End If
next
%>
</table>
<%
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
    <title>学生数一覧</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
<body>
<center>
    <%call gs_title("学生数一覧","一　覧")%>

<BR>

	        <table class="hyo" border="1" width="">
	            <tr>

					
	                <th nowrap class="header">学年・学科別</th>
	                <td nowrap class="detail" align="center"><%=gf_fmtWareki(date())%>現在</td>

	            </tr>
	        </table>
<BR>
<table border=1 class="hyo">
<tr>
<th class="header" rowspan="2">学科</th>
<th class="header" colspan="2"><span style="font-size:12px;">１年</span></th>
<th class="header" colspan="2"><span style="font-size:12px;">２年</span></th>
<th class="header" colspan="3"><span style="font-size:12px;">３年</span></th>
<th class="header" colspan="3"><span style="font-size:12px;">４年</span></th>
<th class="header" colspan="3"><span style="font-size:12px;">５年</span></th>
<th class="header" colspan="3"><span style="font-size:12px;">学科別　合計</span></th>
</tr>
<tr>
<th class="header"><span style="font-size:12px;">計</span></th>
<th class="header"><span style="font-size:12px;">女子</span></th>
<th class="header"><span style="font-size:12px;">計</span></th>
<th class="header"><span style="font-size:12px;">女子</span></th>
<th class="header"><span style="font-size:12px;">計</span></th>
<th class="header"><span style="font-size:12px;">女子</span></th>
<th class="header"><span style="font-size:12px;">留学</span></th>
<th class="header"><span style="font-size:12px;">計</span></th>
<th class="header"><span style="font-size:12px;">女子</span></th>
<th class="header"><span style="font-size:12px;">留学</span></th>
<th class="header"><span style="font-size:12px;">計</span></th>
<th class="header"><span style="font-size:12px;">女子</span></th>
<th class="header"><span style="font-size:12px;">留学</span></th>
<th class="header"><span style="font-size:12px;">計</span></th>
<th class="header"><span style="font-size:12px;">女子</span></th>
<th class="header"><span style="font-size:12px;">留学</span></th>
</tr>
<% call s_writecell() %>
</table>
<table width="98%" border="0">
<TR><TD align="right">
<span class="CAUTION" style="text-align:right;">※（　）書きは、休学者数で内数です。</span><BR>
<span class="CAUTION" style="text-align:right;">※「留学」は、留学生を意味します。</span>
</td></tr>
</table>
<% '混合クラスが存在する場合は、混合クラスの表も出す。%>
<BR>
<% if m_ikongoCls(0) = 1 then %>
<% call s_kongoWrite()
 end If %>
</center>
</body>
</head>
</html>
<%
End Sub
%>