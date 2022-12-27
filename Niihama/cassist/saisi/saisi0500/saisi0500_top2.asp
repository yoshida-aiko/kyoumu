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
	Public m_sGakunenWhere		'//学年コンボセット条件
	Public m_sClassWhere		'//クラスコンボセット条件
	Public m_sClassOption       '//クラスコンボのオプション

	Public m_iGakunen			'//学年
	Public m_iClassNo			'//クラスNO
	Public m_iSyoriNen			'//年度
	Public m_iKyokanCd			'//教官ｺｰﾄﾞ
	Public m_iGakka				'//学科
	Public m_sClassNM			'//クラス名

	Const C_FIRST_DISP_GAKUNEN = 1

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
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		w_iRet = gf_OpenDatabase()
		If w_iRet <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If

		'//値の初期化
        Call s_ClearParam()

		'//パラメータ取得
		Call s_GetParameter()

		'//学年コンボセット時の条件
		Call s_MakeGakunenWhere()

		'//クラスコンボに関するWHEREを作成する
		Call s_MakeClassWhere()

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

End Sub

'********************************************************************************
'*  [機能]  学年コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeGakunenWhere()

    m_sGakunenWhere = ""
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyoriNen
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

End Sub

'********************************************************************************
'*  [機能]  クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeClassWhere()

    m_sClassWhere = ""
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen

    If gf_IsNull(Trim(m_iGakunen)) Then
        '//初期表示時は1年1組を表示する
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & C_FIRST_DISP_GAKUNEN
    Else
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_iGakunen)
    End If

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

    //************************************************************
    //  [機能]  フォームロード時
    //  [引数]  
    //  [戻値]  
    //  [説明]
    //************************************************************
	function jf_winload(){

	}

    //************************************************************
    //  [機能]  表示ボタン押下時
    //  [引数]  
    //  [戻値]  
    //  [説明]
    //************************************************************
	function jf_Search(){
		document.body.style.cursor = "wait";
		with(document.frm){
			f_SetVal();
			target = "_LOWER";
			action = "Wait.asp";
			submit();
		}
	}

    //************************************************************
    //  [機能]  送信データを変数にセット
    //  [引数]  
    //  [戻値]  
    //  [説明]
    //************************************************************
	function f_SetVal(){
		with(document.frm){
			hidGakunen.value = cboGakunenCd.value;
			hidClass.value   = cboClassCd.value;
			hidClassNM.value = cboClassCd.options[cboClassCd.selectedIndex].text;
		}
	}
	window.onload = jf_winload;
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
	></table>

	<br>

    <table border="0">
	    <tr>
	    	<td class="search">
				<table border="0" cellpadding="1" cellspacing="1">
					<tr>
						<td nowrap align="left">クラス</td>
						<td align="left">
<%
		Call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere," style='width:40px;' ",False,m_iGakunen)
%>
						</td>
						<td align="left" width="20">年</td>
						<td align="left" width="90">
<%
		Call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere," style='width:80px;' " & m_sClassOption,False,m_iClassNo)
%>
						</td>
						<td valign="bottom" align="right">
							<input class="button" type="button" onclick="javascript:jf_Search();" value="　表　示　" <%= w_sDisabled %>>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	</td>
	</tr>
	</table>

<input type="hidden" name="hidGakunen" value="<%= m_iGakunen %>">
<input type="hidden" name="hidClass"   value="<%= m_iClassNo %>">
<input type="hidden" name="hidClassNM" value="<%= m_sClassNM %>">
</form>
</center>

</body>
</html>
<%
'---------- HTML END   ----------
End Sub
%>