<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠表示
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0111/kks0111_top.asp
' 機      能: 授業出欠の検索ページ
'-------------------------------------------------------------------------
' 引      数:年度           ＞      SESSION("NENDO")
'            
' 変      数:
' 引      渡:
'            
'            
' 説      明:
'           ■初期表示
'               学年のコンボボックスは1年を表示
'               クラスのコンボボックスはCLASSNOが1を表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう授業出欠一覧を表示させる
'-------------------------------------------------------------------------
' 作      成: 2002/05/07 shin
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const C_FIRST_DISP_GAKUNEN = 1   '//初期表示の時の学年(1年)
	
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	Public m_iSyoriNen		'//処理年度
	Public m_sGakki			'//学期
	Public m_sZenki_Start	'//前期開始日
	Public m_sKouki_Start	'//後期開始日
	Public m_sKouki_End		'//後期終了日
	
	Public m_iGakunenCd		'//学年CD
	Public m_Date			'//システム日付
	
	Public m_sGakunenWhere	'//学年コンボのWHERE文
	Public m_sClassWhere	'//クラスコンボのWHERE文
	
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
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
	Dim w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget
    
    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="授業出欠一覧"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"
	
    On Error Resume Next
    Err.Clear
	
    m_bErrFlg = False
	
    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        If gf_OpenDatabase() <> 0 Then
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
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()
	m_Date = ""
	
	m_iGakunenCd	= 0
    m_iSyoriNen		= 0
    
    m_sGakki		= ""
	m_sZenki_Start	= ""
	m_sKouki_Start	= ""
	m_sKouki_Start	= ""
	
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	m_Date = gf_YYYY_MM_DD(date(),"/")				'//システム日付をセット
	
	m_iGakunenCd = cInt(request("cboGakunenCd"))	'リロード時にセット(新規時は、"")
	m_iSyoriNen = Session("NENDO")
	
End Sub

'********************************************************************************
'*  [機能]  学年コンボのWHERE文の作成
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_GakunenWhere
	
	m_sGakunenWhere = ""
	m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyorinen
	m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"
	
End Sub

'********************************************************************************
'*  [機能]  クラスコンボのWHERE文の作成
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClassWhere
	
	m_sClassWhere = ""
	m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen
	
	If m_iGakunenCd = 0 Then
		'//初期表示時は1年1組を表示する
		m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & C_FIRST_DISP_GAKUNEN
	Else
		m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & m_iGakunenCd
	End If
	
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
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>授業出欠一覧</title>
	
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
	//************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){
		if(!f_InpChk()){ return false; }
		
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
		var ob = new Array();
		ob[0] = eval("document.frm.txtFromDate");
		ob[1] = eval("document.frm.txtToDate");
		
		//■開始日
        //NULLチェック
        if(f_Trim(ob[0].value) == ""){
            f_InpChkErr("開始日が入力されていません",ob[0]);
            return false;
        }
        
        //型チェック
        if(IsDate(ob[0].value) != 0){
        	f_InpChkErr("開始日の日付が不正です",ob[0]);
        	return false;
        }
        
        //前期開始日<=開始日<=後期終了日のチェック
        if(DateParse("<%=m_sZenki_Start%>",ob[0].value) < 0 || DateParse(ob[0].value,"<%=m_sKouki_End%>") < 0){
			f_InpChkErr("開始日には、前期開始日以後、後期終了日以前の日付を入力してください",ob[0]);
			return false;
		}
        
        //■終了日
        //NULLチェック
        if(f_Trim(ob[1].value) == ""){
			f_InpChkErr("終了日が入力されていません",ob[1]);
			return false;
        }
        
        //型チェック
        if(IsDate(ob[1].value) != 0){
			f_InpChkErr("終了日の日付が不正です",ob[1]);
        	return false;
        }
        
        //前期開始日<=終了日<=後期終了日のチェック
        if(DateParse("<%=m_sZenki_Start%>",ob[1].value) < 0 || DateParse(ob[1].value,"<%=m_sKouki_End%>") < 0){
			f_InpChkErr("終了日には、前期開始日以後、後期終了日以前の日付を入力してください",ob[1]);
			return false;
		}
        
        //■期間の取得のﾁｪｯｸ■
        if(DateParse(ob[0].value,ob[1].value) < 0){
        	f_InpChkErr("開始日と終了日を正しく入力してください",ob[0]);
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
    //  [機能]  学年を変更した時(クラス情報をセットしなおすため)
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){
		document.frm.action = "kks0111_top.asp";
        document.frm.target = "topFrame";
        document.frm.submit();
		
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript">
    <%call gs_title("出欠状況","参照")%>
    <form name="frm" method="post">
	
	<center>
    <table border="0">
	    <tr>
		    <td align="right" class="search" nowrap>
				
			    <table border="0">
					<tr>
						<td align="left" nowrap>学年</td>
						<td align="left" nowrap>
							<%  
								call s_GakunenWhere()	'学年コンボのWHERE文
								
								call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange='javascript:f_ReLoadMyPage();' style='width:40px;' ",False,m_iGakunenCd)
							%>
							
							<font>年</font>
							
							<font>　　　　　　　　クラス</font>
							<%
								call s_ClassWhere()		'クラスコンボのWHERE文
								
								call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"style='width:80px;' ",False,"")
							%>
						</td>
						<td align="left" nowrap><br></td>
					</tr>
					
					<tr>
						<td align="left" nowrap>日付</td>
						<td nowrap>
							<input type="text" name="txtFromDate" value="<%=m_Date%>">
							<input type="button" class="button" onclick="fcalender('txtFromDate')" value="選択">　～　
							
							<input type="text" name="txtToDate" value="<%=m_Date%>">
							<input type="button" class="button" onclick="fcalender('txtToDate')" value="選択">
							
						</td>
						
						<td valign="bottom" align="right" nowrap>
							<input class="button" type="button" onclick="javascript:f_Search();" value="　表　示　">
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	
    <!--値渡し用-->
    <input type="hidden" name="txtURL" VALUE="kks0111_bottom.asp">
    <input type="hidden" name="txtMsg" VALUE="しばらくお待ちください">
	
	</center>
    </form>
    
    </body>
    </html>
<%
End Sub
%>