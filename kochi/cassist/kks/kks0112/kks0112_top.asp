<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0112/kks0112_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
'            
'            
' 説      明:
'           ■初期表示
'               科目コンボボックス：ログイン者の担当科目
'               [欠課入力]
'					授業日：システム日付
'					時限  ：時限マスタより取得
'               [欠課一覧参照]
'					指定月：
'           ■選択ボタンクリック時
'               カレンダーを出す
'           ■入力ボタンクリック時
'               下のフレームに指定した条件にかなう授業の出欠入力画面を表示
'           ■表示ボタンクリック時
'               サブウィンドウで指定した条件の出欠状況を表示
'-------------------------------------------------------------------------
' 作      成: 2002/05/16 shin
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    
    Public m_iSyoriNen          '//年度
    Public m_iKyokanCd          '//教官ｺｰﾄﾞ
    
    Public m_sGakki             '//学期
    Public m_sZenki_Start		'//前期開始日
    Public m_sKouki_Start		'//後期開始日
    Public m_sKouki_End			'//後期終了日
    
	Public m_Rs_Jigen			'//時限
	Public m_Rs_Subject			'//科目
	
	Public m_JigenCount			'//時限数
    
    Public m_Month				'//現在の月
    'エラー系
    Public m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
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
        If gf_OpenDatabase() <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If
		
        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))
		
        '//値の初期化
		Call s_ClearParam()
		
        '//変数セット
        Call s_SetParam()
		
		'//前期・後期情報を取得
		if gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End) <> 0 then
			m_bErrFlg = True
        	Exit Do
		end if
		
		'//ログイン教官の担当科目の取得
		if not f_GetSubject() then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//授業データが取得できないとき
		if m_Rs_Subject.EOF then
			Call showWhitePage("授業データがありません")
			Exit Do
		end if
		
		'//時限数の取得
		if not f_Get_JigenData() then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//時限数取得できないとき
		if m_Rs_Jigen.EOF then
			Call showWhitePage("時限数が取得できません")
			Exit Do
		end if
		
        '// ページを表示
        Call showPage()
        Exit Do
    Loop
	
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
    End If
    
    '// 終了処理
    Call gf_closeObject(m_Rs_Jigen)
	Call gf_closeObject(m_Rs_Subject)
	
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
    m_iKyokanCd = ""
    m_Month = 0
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")
	m_Month = month(date())
End Sub

'********************************************************************************
'*  [機能]  ログイン教官の受持教科を取得(年度、教官CD、学期より)
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSubject()
	Dim w_sSQL
    
    On Error Resume Next
    Err.Clear
	
    f_GetSubject = false
	
	'通常、留学生代替科目取得
	w_sSQL = ""
	w_sSQL = w_sSQL & "select "
	w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & "from"
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & "		,M03_KAMOKU "
	w_sSQL = w_sSQL & "		,M100_KAMOKU_ZOKUSEI "
	w_sSQL = w_sSQL & "where "
	w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_iKyokanCd & "'"
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M03_KAMOKU_CD "
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI = " & C_JIK_JUGYO
	
	w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	M03_ZOKUSEI_CD = M100_ZOKUSEI_CD "
	
	w_sSQL = w_sSQL & "	and	M100_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	M100_SYUKKETSU_FLG = 0 "
	
	w_sSQL = w_sSQL & "union "
	
	'特別活動取得
	w_sSQL = w_sSQL & "select "
	w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M41_MEISYO as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & "from "
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & "		,M41_TOKUKATU "
	w_sSQL = w_sSQL & "		,M100_KAMOKU_ZOKUSEI "
	w_sSQL = w_sSQL & "where "
	w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_iKyokanCd & "'"
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M41_TOKUKATU_CD "
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI = " & C_JIK_TOKUBETU
	
	w_sSQL = w_sSQL & "	and	M41_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	M41_ZOKUSEI_CD = M100_ZOKUSEI_CD "
	
	w_sSQL = w_sSQL & "	and	M100_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	M100_SYUKKETSU_FLG = 0 "
	
	w_sSQL = w_sSQL & "order by GAKUNEN,CLASS,KAMOKU_KBN "
	
	If gf_GetRecordset(m_Rs_Subject,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		msMsg = Err.description
		Exit function
	End If
	
	f_GetSubject = true
    
End Function

'********************************************************************************
'*	[機能]	処理年度の時限数の取得
'*	[引数]	
'*	[戻値]	true:成功 false:失敗
'*	[説明]	
'********************************************************************************
function f_Get_JigenData()
	Dim w_sSQL
	
	On Error Resume Next
	Err.Clear
	
	f_Get_JigenData = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  MAX(M07_JIKAN) "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  M07_JIGEN "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  M07_NENDO = " & m_iSyoriNen
	
	If gf_GetRecordset(m_Rs_Jigen,w_sSQL) <> 0 Then
		'ﾚｺｰﾄﾞｾｯﾄの取得失敗
		exit function
	End If
	
	m_JigenCount = cInt(m_Rs_Jigen(0))
	
	f_Get_JigenData = true
	
end function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	Dim w_num
	Dim w_ClassName
%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>授業出欠入力</title>
	
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		if(location.href.indexOf('#')==-1){
			//ヘッダ部を表示submit
			document.frm.target = "main";
			document.frm.action = "white.asp?data_flg=OK"
			document.frm.submit();
		}
    }
    
    //************************************************************
    //  [機能]  入力ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Insert(){
		if(!f_InpChk()){ return false; }
		
		f_SetHidden();
		
    	document.frm.action="WaitAction.asp";
        document.frm.target="main";
        document.frm.submit();
    }
    //************************************************************
    //  [機能]  入力チェック
    //  [引数]  
    //  [戻値]  
    //  [説明]
    //
    //************************************************************
    function f_InpChk(){
		var obj = eval("document.frm.txtDate");
		
		//■開始日
        //NULLチェック
        if(f_Trim(obj.value) == ""){
            f_InpChkErr("授業日が入力されていません",obj);
            return false;
        }
        
        //型チェック
        if(IsDate(obj.value) != 0){
        	f_InpChkErr("授業日の日付が不正です",obj);
        	return false;
        }
        
        //前期開始日<=授業日<=後期終了日のチェック
        if(DateParse("<%=m_sZenki_Start%>",obj.value) < 0 || DateParse(obj.value,"<%=m_sKouki_End%>") < 0){
			f_InpChkErr("授業日には、前期開始日以後、後期終了日以前の日付を入力してください",obj);
			return false;
		}
		
        return true;
		
	}
	
	//************************************************************
    //  [機能]  入力チェックエラー時のalert,focus,select処理
    //************************************************************
    function f_InpChkErr(p_AlertMsg,p_Object){
		alert(p_AlertMsg);
		p_Object.focus();
		p_Object.select();
	}
	
    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){
		var PositionX,PositionY,w_position;
		
		var vl = document.frm.sltKamoku.value.split('#@#');
		
		url = "kks0112_subwin.asp";
		url = url + "?hidSyubetu=" + vl[0];
		url = url + "&hidKamokuCd=" + vl[1];
		url = url + "&hidGakunen=" + vl[2];
		url = url + "&hidClassNo=" + vl[3];
		url = url + "&sltMonth=" + document.frm.sltMonth.value;
		
		w   = window.screen.availWidth;
		h   = window.screen.availHeight-30;
		
		PositionX = window.screen.availWidth  / 2 - w / 2;
		PositionY = 0; //window.screen.availHeight / 2 - h / 2;
		
		w_position = ",left=" + PositionX + ",top=" + PositionY;
		
		opt = "directoris=0,location=0,menubar=0,status=0,toolbar=0,resizable=no";
		opt = opt + ",width=" + w + ",height=" + h;
		opt = opt + w_position;
		
		nWin = window.open(url,"kks0112_subwin",opt);
	}
	
    //************************************************************
    //  [機能]  コンボの年、クラス、科目をばらしてセットする
    //************************************************************
    function f_SetHidden(){
		var vl = document.frm.sltKamoku.value.split('#@#');
		
		//通常・特別授業(種別、課目ｺｰﾄﾞ、学年、ｸﾗｽNOを取得)
		document.frm.hidSyubetu.value = vl[0];
        document.frm.hidKamokuCd.value = vl[1];
        document.frm.hidGakunen.value = vl[2];
        document.frm.hidClassNo.value = vl[3];
		
		document.frm.hidClassName.value = vl[4];
        document.frm.hidKamokuName.value = vl[5];
	}
	//************************************************************
    //  [機能]  エンターキー処理
    //************************************************************
	function f_EnterClick(p_Type){
		if(event.keyCode==13){
			if(p_Type == "INSERT"){
				f_Insert();
			}else{
				f_Search();
			}
		}
	}
	
	//-->
    </SCRIPT>
	
	</head>
	<body LANGUAGE="javascript" onload="return window_onload();">
	<% call gs_title("授業出欠入力","検　索") %>
	<form name="frm" method="post">
	<center>
	<table border="0">
		<tr>
			<td align="right" class="search" nowrap>
				<table border="0">
					<tr>
						<td nowrap>科目</td>
						<td colspan="7" nowrap>
							<select name="sltKamoku" style="width:200px;">
								<% 
									do until m_Rs_Subject.EOF
										w_ClassName = ""
										w_ClassName = gf_GetClassName(m_iSyoriNen,m_Rs_Subject("GAKUNEN"),m_Rs_Subject("CLASS"))
										
								%>
										<option value="<%=CStr(cint(m_Rs_Subject("KAMOKU_KBN")) & "#@#" & m_Rs_Subject("KAMOKU_CD") & "#@#" & m_Rs_Subject("GAKUNEN") & "#@#" & m_Rs_Subject("CLASS") & "#@#" & w_ClassName & "#@#" & m_Rs_Subject("KAMOKU_NAME"))%>"><%=m_Rs_Subject("GAKUNEN") & "年&nbsp;&nbsp;" & w_ClassName & "&nbsp;&nbsp;&nbsp;" & m_Rs_Subject("KAMOKU_NAME") %>
								<%
										m_Rs_Subject.movenext
									loop
								%>
								
							</select>
						</td>
					</tr>
					
				    <tr><td colspan="7" height="10"><img src="../../image/sp_black.gif" width="100%" height="1"></td></tr>
				    
				    <tr>
				    	<th class="header" colspan="4" align="center">欠課入力</td>
				    	
				    	<td rowspan="4"><img src="../../image/sp_black.gif" width="1" height="80"></td>
				    	
				    	<th class="header" colspan="4" align="center">欠課一覧参照</td>
				    </tr>
				    
				    <tr>
				    	<td>授業日</td>
				    	<td>
				    		<input type="text" name="txtDate" value="<%=gf_YYYY_MM_DD(date(),"/")%>" onKeyDown="f_EnterClick('INSERT');">
				    		<input type="button" class="button" onClick="fcalender('txtDate')" value="選択">
				    	</td>
				    	
				    	<td>時限</td>
				    	<td>
				    		<select name="sltJigen">
				    		
				    		<% for w_num=1 to m_JigenCount %>
				    			<option value="<%=w_num%>"><%=w_num%>
				    		<% next %>
				    		
				    	</td>
				    	
				    	<td>指定月</td>
				    	<td><select name="sltMonth"  onKeyDown="f_EnterClick('DISP');">
				    			<option value="4"  <%=gf_iif(m_Month = 4,"selected","")%>  >4
				    			<option value="5"  <%=gf_iif(m_Month = 5,"selected","")%>  >5
				    			<option value="6"  <%=gf_iif(m_Month = 6,"selected","")%>  >6
				    			<option value="7"  <%=gf_iif(m_Month = 7,"selected","")%>  >7
				    			<option value="8"  <%=gf_iif(m_Month = 8,"selected","")%>  >8
				    			<option value="9"  <%=gf_iif(m_Month = 9,"selected","")%>  >9
				    			<option value="10" <%=gf_iif(m_Month = 10,"selected","")%> >10
				    			<option value="11" <%=gf_iif(m_Month = 11,"selected","")%> >11
				    			<option value="12" <%=gf_iif(m_Month = 12,"selected","")%> >12
				    			<option value="1"  <%=gf_iif(m_Month = 1,"selected","")%>  >1
				    			<option value="2"  <%=gf_iif(m_Month = 2,"selected","")%>  >2
				    			<option value="3"  <%=gf_iif(m_Month = 3,"selected","")%>  >3
				    		</select>
				    	</td>
				    </tr>
				    
				    <tr>
				    	
				    </tr>
				    
				    <tr>
				    	<td colspan="4" align="center" nowrap>
							<input class="button" type="button" onclick="javascript:f_Insert();" value="　入　力　">
						</td>
				    	
						<td colspan="2" align="center" nowrap>
							<input class="button" type="button" onclick="javascript:f_Search();" value="　表　示　">
						</td>
					</tr>
			    </table>
				
		    </td>
	    </tr>
    </table>
	
    <!--値渡し用-->
    <input type="hidden" name="Tuki_Zenki_Start" value="<%=m_sZenki_Start%>">
    <input type="hidden" name="Tuki_Kouki_Start" value="<%=m_sKouki_Start%>">
    <input type="hidden" name="Tuki_Kouki_End"   value="<%=m_sKouki_End%>">
    
    <INPUT TYPE="hidden" name="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE="hidden" name="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
    
    <INPUT TYPE="hidden" name="hidGakunen"   value = "">
    <INPUT TYPE="hidden" name="hidClassNo"   value = "">
    <INPUT TYPE="hidden" name="hidKamokuCd" value = "">
    <INPUT TYPE="hidden" name="hidSyubetu"   value = "">
	
    <INPUT TYPE="hidden" name="hidClassName" value = "">
	<INPUT TYPE="hidden" name="hidKamokuName"   value = "">
	
    <input TYPE="hidden" name="txtURL" VALUE="kks0112_bottom.asp">
    <input TYPE="hidden" name="txtMsg" VALUE="しばらくお待ちください">
	
	</form>
	</center>
	</body>
	</html>
<%
End Sub

'********************************************************************************
'*	[機能]	空白HTMLを出力
'*	[引数]	なし
'*	[戻値]	なし
'*	[説明]	
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
	<html>
	<head>
	<title>授業出欠入力</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	
	//************************************************************
	//	[機能]	ページロード時処理
	//	[引数]
	//	[戻値]
	//	[説明]
	//************************************************************
	function window_onload() {
	}
	//-->
	</SCRIPT>
	
	</head>
	<body LANGUAGE="javascript" onload="return window_onload()">
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