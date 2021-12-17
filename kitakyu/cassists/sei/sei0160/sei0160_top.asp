<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 放送大学成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0160/sei0160_top.asp
' 機      能: 上ページ 放送大学成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:
'           :
' 変      数:
' 引      渡:
'           :
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう放送大学成績登録画面を表示させる
'-------------------------------------------------------------------------
' 作      成: 2007/04/11 岩田
' 修      正: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Dim  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
	
	Dim m_iNendo             '年度
	Dim m_sKyokanCd          '教官コード

	Dim m_sGakunen		 '学年
	Dim m_sClass		 'クラス
	Dim m_sBunruiCD		 '分類コード
	Dim m_sBunruiNM		 '分類名称
	Dim m_sTani		 '単位

	Dim m_sClassWhere	'クラス取得Where文字列

	Dim m_bNoData		 ''入力科目がないときTrue
	Dim gRs
	
	
'///////////////////////////メイン処理/////////////////////////////
	
	Call Main()
	
'********************************************************************************
'*  [機能]  本ASPのﾒｲﾝﾙｰﾁﾝ
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub Main()
	Dim w_iRet              '// 戻り値
    Dim w_sSQL              	'// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="放送大学成績登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"
	
    On Error Resume Next
    Err.Clear
	
    m_bErrFlg = false
	
    	Do
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
			m_sErrMsg = "データベースとの接続に失敗しました。"
			Exit Do
		End If
		
		'//値を取得
		call s_SetParam()
		

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
	        '//クラスコンボに関するWHEREを作成する
        	Call s_MakeClassWhere() 

		'//ログイン教官の担当放送大学科目の取得
		if not f_GetNintei() then Exit Do
		
		If gRs.EOF Then
			m_bNoData = True
		Else
			m_bNoData = False
		End If

		'// ページを表示
		Call showPage()
		
		m_bErrFlg = true
		Exit Do
	Loop
	
	'// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
	If not m_bErrFlg Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// 終了処理
	Call gf_closeObject(gRs)
	Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
    m_iNendo    = session("NENDO")		'年度
    m_sKyokanCd = session("KYOKAN_CD")		'教官コード

    m_sGakunen  = request("txtGakunen")         '学年
    m_sClass    = request("txtClass")           'クラス
    m_sBunruiCD = request("txtBunruiCd")	'分類コード
    m_sBunruiNm = request("txtBunruiNm")	'分類名称
    m_sTani     = request("txtTani")		'単位

End Sub

'********************************************************************************
'*  [機能]  クラスコンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeClassWhere()
    
    m_sClassWhere = "" 

    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iNendo  		   '//処理年度

    if m_sGakunen <> "@@@" then
    	m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_sGakunen)    '//学年
    end if

'response.write " m_sClassWhere=" & m_sClassWhere & "<BR>" 

End Sub

'********************************************************************************
'*  [機能]  ログイン教官の認定科目を取得(年度、教官CDより)
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetNintei()
	Dim w_sSQL
    Dim w_sJiki
    Dim w_Rs
    Dim w_sMinNendo 
    
    On Error Resume Next
    Err.Clear
	
    f_GetNintei = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	 M110_BUNRUI_CD     AS BUNRUI_CD "
	w_sSQL = w_sSQL & " 	,M110_BUNRUI_MEISYO AS BUNRUI_NM "
	w_sSQL = w_sSQL & " 	,M110_TANI          AS TANI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & "     M110_NINTEI_H "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "		M110_NENDO =" & m_iNendo
	w_sSQL = w_sSQL & "	AND	M110_KYOKAN_CD ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & " ORDER BY M110_BUNRUI_CD "
'response.write w_sSQL
'response.end		
	If gf_GetRecordset(gRs,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit function
	End If
	
	f_GetNintei = true
    
End Function

'********************************************************************************
'*  HTMLを出力
'********************************************************************************
Sub showPage()
	Dim w_TukuName
	Dim w_SubjectDisp
	Dim w_SubjectValue
	
	On Error Resume Next
    Err.Clear
	
%>
	<html>
	<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//  [機能]  試験が変更されたとき、再表示する
	//************************************************************
	function f_ReLoadMyPage(){
		// 選択されたコンボの値をｾｯﾄ
		f_SetData();
		document.frm.txtClass.value="";

		document.frm.action="sei0160_top.asp";
		document.frm.target="topFrame";
		document.frm.submit();
	}
	
	//************************************************************
	//  [機能]  表示ボタンクリック時の処理
	//************************************************************
	function f_Search(){

		//入力チェック
		if(!f_InpCheck()){
			return false;
		}

		// 選択されたコンボの値をｾｯﾄ
		f_SetData();

		document.frm.action="sei0160_bottom.asp";
		document.frm.target="main";
		document.frm.submit();
	}
	//************************************************
	//	入力チェック
	//************************************************
	function f_InpCheck(){
		var w_length;
		var ob;

		//学年
		ob = eval("document.frm.txtGakunen");
		if(ob.value =="@@@"){
			alert("学年を選択してください");
			ob.focus();
			ob.select();
			return false;
		}

		//クラス
		ob = eval("document.frm.txtClass");
		if(ob.value =="@@@"){
			alert("クラスを選択してください");
			ob.focus();
			ob.select();
			return false;
		}

		//クラス
		ob = eval("document.frm.sltSubject");
		if(ob.value =="@@@"){
			alert("科目を選択してください");
			ob.focus();
			ob.select();
			return false;
		}

		return true;
	}
	
	//************************************************************
	//  [機能]  表示ボタンクリック時に選択されたデータをｾｯﾄ
	//************************************************************
	function f_SetData(){
		//データ取得
		var vl = document.frm.sltSubject.value.split('#@#');
		
		//選択されたデータをｾｯﾄ(分類CD、分類名称、単位を取得)
		document.frm.txtBunruiCd.value=vl[0];
		document.frm.txtBunruiNm.value=vl[1];
		document.frm.txtTani.value=vl[2];
	}
	
	
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>
	
    	<body LANGUAGE="javascript">
	
	<center>
	<form name="frm" METHOD="post">
	
	<% call gs_title(" 放送大学成績登録 "," 登　録 ") %>
	<br>
	
	<table border="0">
		<tr><td valign="bottom">
			
			<table border="0" width="100%">
				<tr><td class="search">
					
					<table border="0">
						<tr valign="middle">
							<td align="left" nowrap>学　年</td>
							<td>
								<select name="txtGakunen" style="width:110px;" onchange ="javascript:f_ReLoadMyPage()" >
									<option value="@@@" selected >  </option>
									<% For I = 1 To 2 
										if cstr(I) = cstr(m_sGakunen) then %>
									<option value="<%=I%>" selected > <%=I%>年</option>
										<% Else %>
									<option value="<%=I%>"> <%=I%>年</option>
										<% end if 
								           Next %>
								</select>
							</td>

							<td>&nbsp;</td>
							
							<td align="left" nowrap>ク ラ ス</td>
							
							<!-- '学年が選択されていない場合は、入力不可にする -->
							<td>
								<%IF m_sGakunen <> "@@@" and m_sGakunen <> "" then 
								 	call gf_ComboSet("txtClass",C_CBO_M05_CLASS,m_sClassWhere," style='width:200px;'",True,m_sClass) 
							 	else %>
								<select name="txtClass" DISABLED style="width:200px;">
									<option value="@@@">　　　　　　　</option>
								</select>
							    <% end if %>
							</td>
	                    			</tr>
						
						<tr>
							<td align="left" nowrap>科目</td>
							<td align="left">
								<% if not gRs.EOF then %>
								<select name="sltSubject" style="width:250px;">
									<% 
									do until gRs.EOF
										
										'科目コンボ表示部分生成
										w_SubjectDisp = gf_SetNull2String(gRs("BUNRUI_NM")) 

										'科目コンボVALUE部分生成
										w_SubjectValue = ""
										w_SubjectValue = w_SubjectValue & gRs("BUNRUI_CD")   & "#@#"
										w_SubjectValue = w_SubjectValue & gRs("BUNRUI_NM")   & "#@#"
										w_SubjectValue = w_SubjectValue & gRs("TANI")  

								
										if cstr(m_sBunruiCD) = gf_SetNull2String(gRs("BUNRUI_CD")) then %>
									<option value="<%=w_SubjectValue%>" selected > <%=w_SubjectDisp%></option>
										<% Else %>
									<option value="<%=w_SubjectValue%>"> <%=w_SubjectDisp%></option>
										<% end if 
										gRs.movenext
									loop 
									%>
								</select>
								<% 	'' eof のときのメッセージを追加
								   else %>
									：担当教官の入力科目はありません
								<!--select name="sltSubject">
									<option value="@@@">　　　　　　　</option>
								</select-->
								<% end if %>
							</td>

							
							<td colspan="7" align="right">
							<% 'EOF のときの処理を追加
								If Not m_bNoData Then %>
								<input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
							<% 	Else %>
							<% 	end if %>
							</td>
						</tr>
					</table>
					
				</td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
	
	<input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtGakuNo"    value="<%=w_sGakunen%>">
	<input type="hidden" name="txtClassNo"   value="<%=m_sClass%>">
	<input type="hidden" name="txtBunruiCd"  value="<%=m_sBunruiCD%>">
	<input type="hidden" name="txtBunruiNm"  value="<%=m_sBunruiNM%>">
	<input type="hidden" name="txtTani"      value="<%=m_sTaniD%>">
	</form>
	</center>
	</body>
	</html>
<%
End Sub

'********************************************************************************
'*	空白HTMLを出力
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
	<html>
	<head>
	<title>放送大学成績登録</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>
	
	<body LANGUAGE="javascript">
	<form name="frm" mothod="post">
	
	<center>
	<br><br><br>
		<span class="msg"><%=Server.HTMLEncode(p_Msg)%></span>
	</center>
	
	<input type="hidden" name="txtMsg" value="<%=Server.HTMLEncode(p_Msg)%>">
	</form>
	</body>
	</html>
<%
End Sub
%>