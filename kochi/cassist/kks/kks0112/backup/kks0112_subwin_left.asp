<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠表示(詳細)
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0112/kks0112_subwin_left.asp
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
	Public m_sClassCd		'//クラスCD
	Public m_sClassName		'//クラス名
	
	Public m_sKamokuCd		'//科目コード
    Public m_sKamokuName	'//科目名
    
    Public m_sSyubetu		'//種別
	Public m_iMonth			'//月
	
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
		
		'//ページ表示
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
	
	m_iSyoriNen = 0
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	m_sClassName = ""
	
    m_sKamokuCd = ""
    m_sKamokuName = ""
    
    m_sSyubetu = ""
	m_iMonth = ""
	
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen = Session("NENDO")
	
	m_sGakunenCd = request("hidGakunen")
	m_sClassCd = request("hidClassNo")
	
	m_sClassName = gf_GetClassName(m_iSyoriNen,m_sGakunenCd,m_sClassCd)
	
	m_sSyubetu = request("hidSyubetu")
	
    m_sKamokuCd = request("hidKamokuCd")
    m_sKamokuName = f_GetKamokuMei(m_iSyoriNen,m_sKamokuCd,m_sSyubetu)
    
	m_iMonth = request("sltMonth")
	
End Sub

'********************************************************************************
'*  [機能]  科目名を取得
'*  [引数]  
'*  [戻値]  
'*  [説明]  
'********************************************************************************
function f_GetKamokuMei(p_SyoriNen,p_KamokuCd,p_Syubetu)
	Dim w_iRet
    Dim w_sSQL,w_Rs
    
	f_GetKamokuMei = ""
	
	On Error Resume Next
    Err.Clear
	
	'通常授業
	if p_Syubetu = "TUJO" then
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
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit function
	End If
	
	f_GetKamokuMei = w_Rs(0)
	
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
	
	<!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    
	//************************************************************
    //  [機能]  ページクローズ
    //************************************************************
    function f_Close(){
		parent.close();
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript">
    <form name="frm" method="post">
    <center>
		<table>
			<tr><td><br><br></td></tr>
			
			<tr>
				<td align="center" colspan="2"><font size="+2"><%=m_iMonth%>月</font></td>
			</tr>
			
			<tr><td><br></td></tr>
			
			<tr>
				<td>
					<table class="hyo" border="1">
						<tr>
							<th class="header" width="45">学年</th>
							<td class="detail" width="120" align="left">&nbsp;<%=m_sGakunenCd%>年</td>
						</tr>
					</table>
				<td>
			</tr>
			
			<tr><td><br></td></tr>
			
			<tr>
				<td>
					<table class="hyo" border="1">
						<tr>	
							<th class="header" width="45">クラス</th>
							<td class="detail" width="120" align="left">&nbsp;<%=m_sClassName%>科</td>
						</tr>
					</table>
				<td>
			</tr>
			
			<tr><td><br></td></tr>
			
			<tr>
				<td>
					<table class="hyo" border="1">
						<tr>
							<th class="header" width="45">科目</th>
							<td class="detail" width="120" align="left">&nbsp;<%=m_sKamokuName%></td>
						</tr>
					</table>
				<td>
			</tr>
			
			<tr><td><br></td></tr>
			
			<tr>
				<td align="center" colspan="2"><input type="button" value="閉じる" onClick="f_Close();" name="btnClose"></td>
			</tr>
		</table>
		
	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>
