<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績参照（教官側）
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0800/default.asp
' 機      能: 
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード		＞		SESSIONより（保留）
' 説      明:
'           ■フレームページ
'-------------------------------------------------------------------------
' 作      成: 2003/05/13 廣田
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

	Public  m_iNendo   			'年度
	Public  m_sKyokanCd			'ログイン教官
	Public  m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ
	Dim     m_iGakunen
    Dim     m_iClass
	Dim     m_sGakName
    Dim     m_sGakNo

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

	Dim w_sWinTitle
	Dim w_sMsgTitle
	Dim w_sMsg
	Dim w_sRetURL
	Dim w_sTarget

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="成績参照"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_parent"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If

		'// 権限チェックに使用
'		Session("PRJ_No") = "SEI0800"

		'// 不正アクセスチェック
		Call gf_userChk(Session("PRJ_No"))

		'//ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()

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
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()

	m_iNendo    = Session("NENDO")
	m_sKyokanCd = Session("KYOKAN_CD")
	m_iGakunen  = Request("hidGakunen")
	m_iClass    = Request("hidClass")
	m_sGakName  = Request("hidGakuseiNM")
	m_sGakNo    = Request("hidGakuseiNo")

End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	On Error Resume Next
	Err.Clear

%>
<html>

<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>

<body LANGUAGE="javascript">
	<center>
	<form name="frm" METHOD="post">
	<% call gs_title(" 成績参照 "," 参　照 ") %>

	<table width="630" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="250" align="center" nowrap>　<%=m_iGakunen%>　年　　<%=gf_GetClassName(m_iNendo,m_iGakunen,m_iClass)%>　　<%=m_sGakName%></td>
			<td width="380" align="right"  nowrap>
				<table width="380" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="180" align="right" nowrap>
							<table border="1" class="hyo" cellspacing="0" cellpadding="0">
								<tr>
									<td width="30" class="CELL1" height="20" style="background : #33CCFF;" nowrap></td>
								</tr>
							</table>
						</td>
						<td align="left" nowrap>= 修得済</td>
						<td align="right" nowrap>
							<table border="1" class="hyo" cellspacing="0" cellpadding="0">
								<tr>
									<td width="30" class="CELL1" height="20" style="background : #FF9900;" nowrap></td>
								</tr>
							</table>
						</td>
						<td align="left" nowrap>= 未修得</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<br>

	<!-- TABLEヘッダ部 -->
	<table border="1" class="hyo" width="630">
		<tr>
			<th width="30"  class="header3" align="center" height="  " rowspan="2" nowrap>&nbsp;</th>
			<th width="30"  class="header3" align="center" height="  " rowspan="2" nowrap>&nbsp;</th>
			<th width="250" class="header3" align="center" height="  " rowspan="2" nowrap>科　目</th>
			<th width="70"  class="header3" align="center" height="  " rowspan="2" nowrap>修得単位</th>
			<th width="250" class="header3" align="center" height="20" colspan="5" nowrap>成　績</th>
		</tr>
		<tr>
			<th width="50" class="header2" align="center" height="20" nowrap>1年</th>
			<th width="50" class="header2" align="center" height="20" nowrap>2年</th>
			<th width="50" class="header2" align="center" height="20" nowrap>3年</th>
			<th width="50" class="header2" align="center" height="20" nowrap>4年</th>
			<th width="50" class="header2" align="center" height="20" nowrap>5年</th>
		</tr>
	</table>

	</form>
	</cinter>
</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
