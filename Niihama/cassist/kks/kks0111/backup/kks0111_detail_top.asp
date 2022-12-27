<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠表示(詳細)
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0111/kks0111_detail_top.asp
' 機      能: 上ページ 授業出欠覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             
' 変      数:
' 引      渡: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
' 説      明:
'            
'-------------------------------------------------------------------------
' 作      成: 2002/05/07 shin
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	
	Public m_bErrFlg		'//エラーフラグ
	
	Public m_iSyoriNen		'//処理年度
	Public m_sGakunenCd		'//学年
	Public m_sClassCd		'//クラス
	Public m_sGakusekiNo	'//学籍NO
	Public m_sName			'//氏名
	Public m_sZaisekiName	'//在籍状況名
	
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
	Dim w_iRet			'// 戻り値
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="授業出欠入力"
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
            w_sMsg = "データベースとの接続に失敗しました。"
            'm_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//変数初期化
		Call s_ClearParam()
		
		'// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
		
		'//在籍区分のチェック処理
		if not f_ZaisekiChk() then
			m_bErrFlg = True
            Exit Do
		end if
		
		Call showPage()
		
        Exit Do
    Loop
	
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	m_sName = ""
	
    m_iSyoriNen = ""
    
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen = Session("NENDO")
	
	m_sGakunenCd = request("Nen")
	m_sClassCd = request("Class")
	m_sGakusekiNo = request("GakusekiNo")
	
	m_sName = f_GetName(m_iSyoriNen,m_sGakusekiNo)
    
End Sub
'********************************************************************************
'*  [機能]  学生名称取得
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
function f_GetName(p_SyoriNen,p_GakusekiNo)
	Dim w_iRet
    Dim w_sSQL,w_Rs
    
	f_GetName = ""
	
	On Error Resume Next
    Err.Clear
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T11_SIMEI "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T13_GAKU_NEN,"
	w_sSQL = w_sSQL & "  T11_GAKUSEKI "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  T13_GAKUSEI_NO = T11_GAKUSEI_NO "
	w_sSQL = w_sSQL & "  and T13_NENDO = " & p_SyoriNen
	w_sSQL = w_sSQL & "  and T13_GAKUSEKI_NO =" & p_GakusekiNo
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit function
	End If
	
	f_GetName = w_Rs(0)
	
end function

'********************************************************************************
'*  [機能]  在籍区分のチェック処理
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
function f_ZaisekiChk()
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_Rs_Zaiseki
	Dim w_ZaisekiKbn
	
	On Error Resume Next
	Err.Clear
	
	f_ZaisekiChk = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T13_ZAISEKI_KBN "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T13_GAKU_NEN "
	
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "      T13_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and T13_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T13_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & "  and T13_GAKUSEKI_NO ='" & m_sGakusekiNo & "'"
	
	w_iRet = gf_GetRecordset(w_Rs_Zaiseki,w_sSQL)
	
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	if not w_Rs_Zaiseki.EOF then
		w_ZaisekiKbn = cInt(w_Rs_Zaiseki("T13_ZAISEKI_KBN"))
		
		if w_ZaisekiKbn <> C_ZAI_ZAIGAKU then
			'在籍中でないとき、在籍区分名を取得
			if not f_Get_ZaisekiName(w_ZaisekiKbn,m_sZaisekiName) then exit function
		else
			m_sZaisekiName = ""
		end if
		
	end if
	
	f_ZaisekiChk = true
	
	
end function

'********************************************************************************
'*	[機能]	在籍区分名称の取得
'*	[引数]	
'*	[戻値]	
'*	[説明]	
'********************************************************************************
function f_Get_ZaisekiName(p_ZaisekiKbn,w_sZaisekiName)
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_Rs_Zaiseki
	
	On Error Resume Next
	Err.Clear
	
	f_Get_ZaisekiName = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  M01_SYOBUNRUIMEI "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  M01_KUBUN "
	
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "      M01_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and M01_DAIBUNRUI_CD = " & C_ZAISEKI
	w_sSQL = w_sSQL & "  and M01_SYOBUNRUI_CD = " & p_ZaisekiKbn
	
	w_iRet = gf_GetRecordset(w_Rs_Zaiseki,w_sSQL)
	
	If w_iRet <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		exit function
	End If
	
	w_sZaisekiName = w_Rs_Zaiseki("M01_SYOBUNRUIMEI")
	
	f_Get_ZaisekiName = true
	
end function 


'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	On Error Resume Next
    Err.Clear
%>
    <html>
    <head>
    <title>授業出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
	
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
	
    //************************************************************
    //  [機能]  ページロード時処理
    //************************************************************
    function window_onload() {
		
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    <%call gs_title("出欠状況","参照")%>
    <%Do %>
        
        <table>
        	<tr>
				<td class="search" nowrap>
					<table>
						<tr>
							<th class="header">学籍NO</th>
							<td><%=m_sGakusekiNo%></td>
							
							<th class="header">氏名</th>
							<td><%=m_sName%></td>
							
							<td><font color="#FF0000"><%=m_sZaisekiName%></font></td>
						</tr>
					</table>
				</td>
			</tr>
			
			<tr>
				<td align="center"><input type="button" value="閉じる" onClick="javascript:parent.close();"></td>
			</tr>
		</table>
		
		
		<table>
	        <tr>
	        	<td>
	        		<table width="540" class="hyo"  border="1">
	        			<tr>
							<th class="header" width="130" align="center" nowrap>日付</th>
				            <th class="header" width="60"  align="center" nowrap>時限</th>
				            <th class="header" width="150" align="center" nowrap>科目</th>
				            <th class="header" width="120" align="center" nowrap>入力教官</th>
				            <th class="header" width="80"  align="center" nowrap>状況</th>
	            		</tr>
	            	</table>
	            </td>
	        </tr>
        </table>
		
        <%Exit Do%>
    <%Loop%>
	
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>
