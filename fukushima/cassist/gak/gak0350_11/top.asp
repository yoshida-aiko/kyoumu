<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0350_11/top.asp
' 機      能: 上ページ 学籍データの検索を行う
'-------------------------------------------------------------------------
' 変      数:なし
' 引      渡:処理年度       ＞      SESSIONより（保留）
'			txtMode				   :動作モード
'           txtGakunen             :学年
'           txtGakkaCD             :学科
'           txtClass               :クラス
'           txtName                :名称
'           txtGakusekiNo          :学籍番号
'           txtGakuseiNo           :学生番号
' 説      明:
'           ■初期表示
'               コンボボックス      学年を表示
'                                  学科を表示
'                                  クラスを表示
'           ■表示ボタンクリック時
'               下のフレームに指定した検索条件にかなう学生情報を表示させる
'-------------------------------------------------------------------------
' 作      成: 2006/04/26 熊野
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg			   'ｴﾗｰﾌﾗｸﾞ
    
    '選択用のWhere条件
    Public s_sGakkaWhere		   '学科の抽出条件
    Public m_sClassWhere		   'クラスの抽出条件
       
    '取得したデータを持つ変数
    Public  m_iSyoriNen      	   ':処理年度
    Public  m_TxtMode      	       ':動作モード
    Public  m_sGakunen             ':学年
    Public  m_sGakkaCD             ':学科
    Public  m_sClass               ':クラス
    Public  m_sName                ':名称
    Public  m_sGakusekiNo          ':学籍番号
    Public  m_sGakuseiNo           ':学生番号
	
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

    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="学生情報検索"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_TxtMode=request("txtMode")

    '// ﾊﾟﾗﾒｰﾀSET
	if m_TxtMode = "" then
        	Call s_IntParam()
	else
        	Call s_SetParam()
	end if

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
        
        '学科コンボに関するWHEREを作成する
        Call s_MakeGakkaWhere() 
        
        'クラスコンボに関するWHEREを作成する
        Call s_MakeClassWhere() 
        
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
'*  [機能]  パラメータ初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_IntParam()

	m_iSyoriNen = cint(Session("Nendo"))	'処理年度
	m_sGakunen=""            				'学年
    m_sGakkaCD=""             				'学科
    m_sClass=""               				'クラス
    m_sName=""                				'名称
    m_sGakusekiNo=""          				'学籍番号
    m_sGakuseiNo=""           				'学生番号
 	
End Sub


'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen    = cint(Session("Nendo"))			 '処理年度
    m_sGakunen     = request("txtGakunen")           '学年
    m_sGakkaCD     = request("txtGakka")             '学科
    m_sClass       = request("txtClass")             'クラス
    m_sName        = request("txtName")              '名称
    m_sGakusekiNo  = request("txtGakusekiNo")        '学籍番号
    m_sGakuseiNo   = request("txtGakuseiNo")         '学生番号
 	
End Sub

'********************************************************************************
'*  [機能]  学科コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeGakkaWhere()

    s_sGakkaWhere = ""
    s_sGakkaWhere = s_sGakkaWhere & " M02_NENDO = " & m_iSyoriNen & " AND "
    s_sGakkaWhere = s_sGakkaWhere & " M02_GAKKA_CD <> '00'"

End Sub

'********************************************************************************
'*  [機能]  クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeClassWhere()
    
    m_sClassWhere = "" 
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyoriNen 		
    
    if m_sGakunen <> "@@@" then
        	m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_sGakunen)    
	end if

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
<link rel=stylesheet href=../../common/style.css type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {

    }

    //************************************************************
    //  [機能]  戻るボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_BackClick(){

        document.frm.action="../../menu/kensaku.asp";
        document.frm.target="_parent";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  検索実行ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

		document.frm.action="./main.asp";
        document.frm.target="fMain";
        document.frm.txtMode.value = "Search";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  学年, 年度が選択されたとき、再表示する
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="top.asp";
        document.frm.target="fTop";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //************************************************************
    //  [機能] クリアボタンを押されたとき
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function jf_Clear(){
        document.frm.txtGakunen.value = "";
        document.frm.txtGakka.value = "";
        document.frm.txtClass.value = "";
        document.frm.txtMode.value = "";
        
    }
    //-->
    </SCRIPT>

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="center">
<%call gs_title("学生情報検索","一　覧")%>
<form action="./main.asp" method="post" name="frm" target="fMain">

 	<input type="hidden" name="txtMode" width="100%" value="<%=m_TxtMode %>">

	<table cellspacing="0" cellpadding="0" border="0" width="100%" >
		<tr>
			<td valign="top" align="center">
					<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td class=search valign ="top">
								<table border="0" bgcolor="#E4E4ED" cellpadding="0" cellspacing="0">
									<tr>
										<td valign="top">
											<table border="0">
												<tr>
													<td nowrap height="16">学　年
													</td>
													<td>
														<select name="txtGakunen" style="width:110px;" onchange ="javascript:f_ReLoadMyPage()" >
															<option value="@@@" selected >  </option>
															<% For I = 1 To 5 
																if cstr(I) = cstr(m_sGakunen) then %>
																	<option value="<%=I%>" selected > <%=I%>年</option>
																<% Else %>
																	<option value="<%=I%>"> <%=I%>年</option>
																<% end if 
														Next %>
														</select>
													</td>
													<td nowrap height="16">学　科
													</td>
													<td>
														<!-- 2015.03.20 Upd width:110->180 -->
														<% call gf_ComboSet("txtGakka",C_CBO_M02_GAKKA,s_sGakkaWhere," style='width:180px;'",True,m_sGakkaCD) %>
													</td>
													<td nowrap height="16">ク ラ ス
													</td>
												
													<!-- '学年が選択されていない場合は、入力不可にする -->
													<td>
														<%IF m_sGakunen <> "@@@" and m_sGakunen <> "" then 
															<!-- 2015.03.20 Upd width:110->180 -->
								 							call gf_ComboSet("txtClass",C_CBO_M05_CLASS,m_sClassWhere," style='width:180px;'",True,m_sClass) 
							 							else %>
															<!-- 2015.03.20 Upd width:110->180 -->
															<select name="txtClass" DISABLED style="width:180px;" ID="Select1">
															<option value="@@@">　　　　　　　</option>
															</select>
														<% end if %>
													</td>
													<td>
														<input type="button" class=button value="ク　リ　ア" onclick="jf_Clear()" ID="Button1" NAME="Button1">
														<input type="button" class=button value=" 表　示 " onClick="javascript:f_Search()" ID="Button2" NAME="Button2">
													</td>
													<!-- '学年が選択されていない場合は、入力不可にする -->
												</tr>
											</table>
										</td>
										<td valign="top" nowrap>
										</td>
									</tr>
									<tr>
										<td align="right">
												<input type="hidden" name ="txtDisp" value="100" >
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
</div>

</body>
</html>

<%
End Sub
%>

