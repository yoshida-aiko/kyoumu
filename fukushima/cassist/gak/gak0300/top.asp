<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/top.asp
' 機      能: 上ページ 学籍データの検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtHyoujiNendo         :表示年度
'           txtGakunen             :学年
'           txtGakkaCD             :学科
'           txtClass               :クラス
'           txtName                :名称
'           txtGakusekiNo          :学籍番号
'           txtSeibetu             :性別
'           txtGakuseiNo           :学生番号
'           txtIdou                :異動
'           txtTyuClub             :中学校クラブ
'           txtClub                :現在クラブ
'           txtRyoseiKbn           :寮
'           txtMode                :動作モード
'                               BLANK   :初期表示
' 説      明:
'           ■初期表示
'               コンボボックス - 表示年度を表示
'                                学年を表示
'                                学科を表示
'                                クラスを表示
'                                中学校クラブを表示
'                                現在クラブを表示
'                                寮生区分を表示
'           ■表示ボタンクリック時
'               下のフレームに指定した検索条件にかなう学生情報を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/23 ﾓﾁﾅｶﾞ
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '試験選択用のWhere条件
    Public s_sGakkaWhere        '学科の抽出条件
    Public m_sBukatuWhere      	'クラブの抽出条件
    Public m_sClassWhere       	'クラスの抽出条件
    Public m_sRyoseiKbnWhere    '寮生区分の抽出条件
    Public m_sSeibetuWhere      '性別の抽出条件
    Public m_sIdouWhere         '異動の抽出条件
    
    '取得したデータを持つ変数
    Public  m_TxtMode      	       ':動作モード
    Public  m_iSyoriNen      	   ':処理年度
    Public  m_iHyoujiNendo         ':表示年度
    Public  m_sGakunen             ':学年
    Public  m_sGakkaCD             ':学科
    Public  m_sClass               ':クラス
    Public  m_sName                ':名称
    Public  m_sGakusekiNo          ':学籍番号
    Public  m_sSeibetu             ':性別
    Public  m_sGakuseiNo           ':学生番号
    Public  m_sIdou                ':異動
    Public  m_sTyuClub             ':中学校クラブ
    Public  m_sClub                ':現在クラブ
    Public  m_sRyoseiKbn           ':寮
	Public  m_sTyugaku			   ':出身中学校

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

    m_iSyoriNen = Session("NENDO")
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
        
        'クラブコンボに関するWHEREを作成する
        Call s_MakeBukatuWhere() 

        'クラスコンボに関するWHEREを作成する
        Call s_MakeClassWhere() 

        '性別コンボに関するWHEREを作成する
        Call s_MakeSeibetuWhere() 

        '寮コンボに関するWHEREを作成する
        Call s_MakeRyoseiKbnWhere() 

        '異動コンボに関するWHEREを作成する
        Call s_MakeIdouWhere() 

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
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_IntParam()
'response.write " s_IntParam <BR>" 

	m_iHyoujiNendo =m_iSyoriNen		'表示年度
    m_sGakunen=""            		'学年
    m_sGakkaCD=""             		'学科
    m_sClass=""               		'クラス
    m_sName=""                		'名称
    m_sGakusekiNo=""          		'学籍番号
    m_sSeibetu=""            		'性別
    m_sGakuseiNo=""           		'学生番号
    m_sIdou =""               		'異動
    m_sTyuClub =""            		'中学校クラブ
    m_sClub=""                		'現在クラブ
    m_sRyoseiKbn=""           		'寮
	m_sTyugaku = ""					'出身中学校

End Sub


'********************************************************************************
'*  [機能]  引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
'response.write " s_SetParam <BR>" 

    m_iHyoujiNendo = Session("NENDO")     	 '表示年度
    m_sGakunen     = request("txtGakunen")           '学年
    m_sGakkaCD     = request("txtGakka")             '学科
    m_sClass       = request("txtClass")             'クラス
    m_sName        = request("txtName")              '名称
    m_sGakusekiNo  = request("txtGakusekiNo")        '学籍番号
    m_sSeibetu     = request("txtSeibetu")           '性別
    m_sGakuseiNo   = request("txtGakuseiNo")         '学生番号
    m_sIdou        = request("txtIdou")              '異動
    m_sTyuClub     = request("txtTyuClub")           '中学校クラブ
    m_sClub        = request("txtClub")              '現在クラブ
    m_sRyoseiKbn   = request("txtRyoseiKbn")         '寮
	m_sTyugaku     = request("txtTyugaku")			 '出身中学校

'response.write " m_sGakunen = " & m_sGakunen & "<BR>"
End Sub


'********************************************************************************
'*  [機能]  学科コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeGakkaWhere()
    
    s_sGakkaWhere = ""
    s_sGakkaWhere = m_sGakkaWhere & " M02_NENDO = " & m_iHyoujiNendo  '//表示年度

End Sub


'********************************************************************************
'*  [機能]  クラブコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeBukatuWhere()
    
    m_sBukatuWhere = ""
    
    m_sBukatuWhere = m_sBukatuWhere & " M17_NENDO =" & m_iHyoujiNendo  '//表示年度
'response.write " m_sBukatuWhere=" & m_sBukatuWhere & "<BR>" 

End Sub


'********************************************************************************
'*  [機能]  クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeClassWhere()
    
    m_sClassWhere = "" 

        		
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iHyoujiNendo  			'//表示年度
    
    if m_sGakunen <> "@@@" then
    
    	m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_sGakunen)    '//学年

	end if
'response.write " m_sClassWhere=" & m_sClassWhere & "<BR>" 

End Sub


'********************************************************************************
'*  [機能]  寮コンボに関するWHEREを作成する（寮生区分）
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeRyoseiKbnWhere()

    m_sRyoseiKbnWhere = ""
    

    m_sRyoseiKbnWhere = m_sRyoseiKbnWhere & " M01_NENDO = " & m_iHyoujiNendo  	'//表示年度
    m_sRyoseiKbnWhere = m_sRyoseiKbnWhere & " AND M01_DAIBUNRUI_CD = 23 "  	' //入寮区分
    
'response.write " m_sRyoseiKbnWhere=" & m_sRyoseiKbnWhere & "<BR>" 

End Sub


'********************************************************************************
'*  [機能]  性別コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeSeibetuWhere()

    m_sSeibetuWhere = ""
    

    m_sSeibetuWhere  = m_sSeibetuWhere  & " M01_NENDO = " & m_iHyoujiNendo  	'//表示年度
    m_sSeibetuWhere  = m_sSeibetuWhere  & " AND M01_DAIBUNRUI_CD = 1 "			'//性別
    
'response.write " m_sSeibetuWhere =" & m_sSeibetuWhere  & "<BR>" 

End Sub

'********************************************************************************
'*  [機能]  異動コンボに関するWHEREを作成する（寮生区分）
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeIdouWhere()

    m_sIdouWhere = ""
    
    m_sIdouWhere = m_sIdouWhere & " M01_NENDO = " & m_iHyoujiNendo  	'//表示年度
    m_sIdouWhere = m_sIdouWhere & " AND M01_DAIBUNRUI_CD = 9 "			'//在籍異動区分
    
End Sub

'********************************************************************************
'*  [機能]  表示件数コンボの作成
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function s_MakeDispCmb()

	dim sel1,sel2,sel3
	
	sel1 = ""
	sel2 = ""
	sel3 = ""

	select case request("txtDisp")
		case "50"
			sel2 = " selected"
		case "100"
			sel3 = " selected"
		case else ' = "10"
			sel1 = " selected"
	End Select

    s_MakeDispCmb = ""
    
    s_MakeDispCmb = s_MakeDispCmb & "<Select name ='txtDisp'>"
    s_MakeDispCmb = s_MakeDispCmb & "<option value='10'"&sel1&"> 10</option>"
    s_MakeDispCmb = s_MakeDispCmb & "<option value='50'"&sel2&"> 50</option>"
    s_MakeDispCmb = s_MakeDispCmb & "<option value='100'"&sel3&">100</option>"
    s_MakeDispCmb = s_MakeDispCmb & "</Select>"
    
End Function

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
        document.frm.txtName.value = "";
        document.frm.txtGakusekiNo.value = "";
        document.frm.txtSeibetu.value = "";
        document.frm.txtGakuseiNo.value = "";
        document.frm.txtTyuClub.value = "";
        document.frm.txtClub.value = "";
        document.frm.txtRyoseiKbn.value = "";
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

	<table cellspacing="0" cellpadding="0" border="0" width="100%">
		<tr><td valign="top" align="center">

			<table border="0" cellpadding="0" cellspacing="0"><tr><td class=search>

				<table border="0" bgcolor="#E4E4ED" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top">
						<table border="0">
<!--							<tr><td nowrap height="16">表示年度</td>
								<td>
									<select name="txtHyoujiNendo" style="width:110px;" onchange ="javascript:f_ReLoadMyPage()" >
									<% For I = 0 To 4 
										w_iNen = m_iSyoriNen -I 
										if w_iNen = cint(m_iHyoujiNendo) then %>
										<option value="<%=w_iNen%>" selected > <%=w_iNen%>年度</option>
									<% Else %>
										<option value="<%=w_iNen%>"> <%=w_iNen%>年度</option>
									<% end if 
									   Next %>
									</select>
								</td>
							</tr>
//-->
							<tr><td nowrap height="16">学　年</td>
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
							</tr>
							<tr><td nowrap height="16">学　科</td>
							<td>
								<% call gf_ComboSet("txtGakka",C_CBO_M02_GAKKA,s_sGakkaWhere," style='width:110px;'",True,m_sGakkaCD) %>
							</td>
							</tr>				
							<tr><td nowrap height="16">ク ラ ス</td>
							
							<!-- '学年が選択されていない場合は、入力不可にする -->
							<td>
								<%IF m_sGakunen <> "@@@" and m_sGakunen <> "" then 
								 	call gf_ComboSet("txtClass",C_CBO_M05_CLASS,m_sClassWhere," style='width:110px;'",True,m_sClass) 
							 	else %>
									<select name="txtClass" DISABLED style="width:110px;">
									<option value="@@@">　　　　　　　</option>
									</select>
							    <% end if %>
							</td>
							</tr>
							<tr><td nowrap height="16">性　別</td>
							<td>
								<% call gf_ComboSet("txtSeibetu",C_CBO_M01_KUBUN,m_sSeibetuWhere," style='width:110px;'",True,m_sSeibetu) %>
							</td>
							</tr>
						</table>
					</td>
					<td valign="top">
						<table border="0">
							<tr><td nowrap height="16">名称(全角カナ)</td>
								<td>
									<input nowrap type="text" size="15" name="txtName" maxlength="60" value="<%=m_sName %>">
								</td>
								</tr>
								<tr><td nowrap height="16"><%=gf_GetGakuNomei(m_iHyoujiNendo,C_K_KOJIN_1NEN)%></td>
								<td>
									<input type="text" size="15" name="txtGakusekiNo" maxlength="5" value="<%=m_sGakusekiNo %>">
								</td>
							</tr>
							<tr><td nowrap height="16"><%=gf_GetGakuNomei(m_iHyoujiNendo,C_K_KOJIN_5NEN)%></td>
							<td>
								<input type="text" size="15" name="txtGakuseiNo" maxlength="10" value="<%=m_sGakuseiNo %>">
							</td>
							</tr>
						</table>
					</td>

					<td valign="top" nowrap>
						<table border="0">
							<tr><td nowrap height="16">出身中学校</td>
								<td>
									<input nowrap type="text" size="15" name="txtTyugaku" maxlength="60" value="<%=m_sTyugaku %>">
								</td>
							</tr>
							<tr><td nowrap height="16">中学校クラブ</td>
								<td>
									<% call gf_ComboSet("txtTyuClub",C_CBO_M17_BUKATUDO,m_sBukatuWhere," style='width:140px;'",True,m_sTyuClub) %>
								</td>
							</tr>
							<tr><td nowrap height="16">現在クラブ</td>
								<td>
									<% call gf_ComboSet("txtClub",C_CBO_M17_BUKATUDO,m_sBukatuWhere," style='width:140px;'",True,m_sClub) %>
								</td>
							</tr>
							<tr><td nowrap height="16">寮</td>
								<td>					
									<% call gf_ComboSet("txtRyoseiKbn",C_CBO_M01_KUBUN,m_sRyoseiKbnWhere," style='width:140px;'",True,m_sRyoseiKbn) %>
								</td>
							</tr>
<!--
							<tr>
								<td align="right">
									<input type="checkbox" name ="CheckImage" value="image" >画像検索する
								</td>
								<td>
									<input type="button" class=button value="ク　リ　ア" onclick="jf_Clear()" >
									<input type="button" class=button value="　表示　" onClick="javascript:f_Search()">
								</td>
							</tr>
//-->
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<input type="checkbox" name ="CheckImage" value="image" >画像検索する
					</td>
					<td>　表示件数　 　 　<%=s_MakeDispCmb()%>件
					</td>
					<td align="right">
					<input type="button" class=button value="ク　リ　ア" onclick="jf_Clear()" >
					<input type="button" class=button value=" 表　示 " onClick="javascript:f_Search()">
				</td></tr>
				</table>

			</td>
		</tr>
		</table>
	</td></tr>
<!--
	<tr><td align="right">
		<input class=button type="button" value="　表示　" onClick="javascript:f_Search()">
	</td></tr>
//-->
</table>
</form>
</div>

</body>
</html>

<%
End Sub
%>

