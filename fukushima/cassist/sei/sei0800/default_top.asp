<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 成績参照
'* ﾌﾟﾛｸﾞﾗﾑID : sei/sei0800/default_top.asp
'* 機      能: 
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*           :session("PRJ_No")      '権限ﾁｪｯｸのキー
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'* 説      明:
'*           ■初期表示
'*               コンボボックスは学年を表示
'*           ■表示ボタンクリック時
'*               下のフレームに指定した条件の留年該当者一覧を表示させる
'*-------------------------------------------------------------------------
'* 作      成: 2003/05/15 廣田
'* 変      更: 2015/03/19 清本 Win7対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
	Dim m_bErrFlg			'ｴﾗｰﾌﾗｸﾞ

    '選択用のWhere条件
	Dim m_sGakunenWhere		'学年の条件
	Dim m_sClassWhere		'試験の条件

    Dim m_sClassOption		'クラスコンボのオプション
	Dim m_iNendo			'処理年度
	Dim m_iGakunen
    Dim m_iClassNo
    
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

	Dim w_iRet              '// 戻り値
	Dim w_sSQL              '// SQL文
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message用の変数の初期化
	w_sWinTitle= "キャンパスアシスト"
	w_sMsgTitle= "留年該当者一覧"
	w_sMsg= ""
	w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget= ""

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// ﾃﾞｰﾀﾍﾞｰｽ接続
		if gf_OpenDatabase() <> 0 Then
			'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
			m_bErrFlg = True
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

		'// ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()

		'学年コンボ、クラスコンボに関するWHEREを作成する
		Call s_MakeGakunenWhere() 

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

Sub s_SetParam()
'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************

	m_iNendo    = Session("NENDO")			'//処理年度
    m_iGakunen  = Request("cboGakunenCD")	'//学年
    m_iClassNo  = Request("cboClassCD")		'//クラス

End Sub

Sub s_MakeGakunenWhere()
'********************************************************************************
'*  [機能]  学年コンボ、クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

	'学年
	m_sGakunenWhere = ""
	m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iNendo
	m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

	'クラス
	m_sClassWhere = ""
	m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iNendo

	If m_iGakunen = "" Then
		m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = 1"							'//初期表示時は1年1組を表示
	Else
		m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_iGakunen)		'//選択クラスを表示
	End If

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
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

	//************************************************************
	//  [機能]  表示ボタンクリック時の処理
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_Search(){
		with(document.frm){
			action="sei0800_listbottom.asp";
			target="main";
			submit();
		}
	}

	//************************************************************
	//  [機能]  学年が変更されたとき、本画面を再表示
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
    function f_ReLoadMyPage(){
		with(document.frm){
			action="default_top.asp";
			target="topFrame";
			submit();
		}
	}

//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" method="post">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
	<tr>
		<td valign="top" align="center">
		<%call gs_title("成績参照","参　照")%>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td class="search">
					<table border="0" cellpadding="1" cellspacing="1">
						<tr>
							<td align="left">
								<table border="0" cellpadding="1" cellspacing="1">
									<tr>
										<td align="left" nowrap>学年</td>
										<td align="left" nowrap><% call gf_ComboSet("cboGakunenCD",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' ",False,m_iGakunen) %>年</td>
										<td align="left" nowrap>クラス</td>
										<!-- 2015.03.19 Upd width:80→180 -->
										<td align="left" nowrap><% call gf_ComboSet("cboClassCD",C_CBO_M05_CLASS,m_sClassWhere,"style='width:180px;' " & m_sClassOption,False,m_iClassNo) %></td>
									    <td valign="bottom"><input type="button" value="　表　示　" onClick = "javascript:f_Search()" class="button"></td>
									</tr>
								</table>
							</td>
				        </tr>
			        </table>
			    </td>
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
