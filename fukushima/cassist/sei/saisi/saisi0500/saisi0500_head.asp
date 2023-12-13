<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名:
' ﾌﾟﾛｸﾞﾗﾑID :
' 機      能:
'-------------------------------------------------------------------------
' 引      数:
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2003/02/24 hirota
'*************************************************************************/

%>
<!--#include file="../../Common/com_All.asp"-->
<%

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

	Public msURL
	Public m_bErrFlg

	Public m_iGakunen			'//学年
	Public m_iClassNo			'//クラスNO
	Public m_iSyoriNen			'//年度
	Public m_iKyokanCd			'//教官ｺｰﾄﾞ
	Public m_iGakka				'//学科
	Public m_sClass				'//クラス
	Public m_sClassNM			'//クラス名

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

	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	Dim w_iRet

	On Error Resume Next
	Err.Clear

	'Message用の変数の初期化
	w_sWinTitle="キャンパスアシスト"
	w_sMsgTitle="不合格学生一覧"
	w_sMsg=""
	w_sRetURL = C_RetURL & C_ERR_RETURL
	w_sTarget = "fTopMain"

	m_bErrFlg = False

	Do
		'//値の初期化
        Call s_ClearParam()

		'//パラメータ取得
		Call s_GetParameter()

		'//ページを表示
		Call showPage()

		m_bErrFlg = True
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If Not m_bErrFlg Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
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

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_iGakunen  = ""
    m_iClassNo  = ""

End Sub

'********************************************************************************
'*	[機能]	パラメータ取得
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub s_GetParameter()

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")
	m_sClass    = Request("hidClass")
	m_iGakunen  = Request("hidGakunen")
	m_sClassNM  = Request("hidClassNM")

End Sub

'********************************************************************************
'*	[機能]	HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub showPage()

	'---------- HTML START ----------
%>
<html>
<head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>不合格学生一覧</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--

	//-->
	</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<center>

<form name="frm" action="" target="main" Method="POST">

	<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
		<tr>
			<td valign="top" align="center">

			<table cellspacing="0" cellpadding="0" border="0" width="98%">
				<tr>
					<td height="27" width="100%" align="left"
					>
						<DIV class=title>不合格学生一覧</DIV>
					</td
					>
				</tr
				>
				<tr
					><td height="4" width="5%" background="/cassist/image/table_sita.gif"
						><img src="/cassist/image/sp.gif"
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
										><img src="/cassist/image/sp.gif" width=8
								        ><font color="#ffffff"
										>一覧</font
										><img src="/cassist/image/sp.gif" width=8
									></DIV
								></td
							></tr
						></table
						>
					</td
				></tr
			></table
			>
		</tr
		><tr>
			<td align="center">
				<table class="hyo" border="1" width="260" height="20">
				    <tr>
				        <th class="header" width="80"  align="center" nowrap>クラス</th>
				        <td class="detail" width="100" align="center" nowrap><%= m_iGakunen %> 年</td>
				        <td class="detail" width="180" align="center" nowrap><%= m_sClassNM %></td>
				    </tr>
				</table>
			</td>
		</tr>
		<tr>
			<td valign="bottom" align="center">
				<table class="hyo" border="1">
					<tr>
						<th width="60"  height="30" class="header3" nowrap>出席番号</th>
						<th width="150" height="30" class="header3" nowrap>氏名</th>
						<th width="150" height="30" class="header3" nowrap>科目</th>
						<th width="50"  height="30" class="header3" nowrap>年度</th>
						<th width="70"  height="30" class="header3" nowrap>欠課/授業</th>
						<th width="40"  height="30" class="header3" nowrap>旧評価</th>
						<th width="40"  height="30" class="header3" nowrap>新評価</th>
						<th width="100" height="30" class="header3" nowrap>担当教官</th>
					</tr>
				</table>
			</td>
		</tr>
	</table>

</form>
</center>

</body>
</html>
<%
'---------- HTML END   ----------
End Sub
%>