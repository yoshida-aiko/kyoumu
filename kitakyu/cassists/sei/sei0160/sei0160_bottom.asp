<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 放送大学成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0160_bottom.asp
' 機      能: 下ページ 放送大学成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード		＞		SESSIONより
'           :年度		＞		SESSIONより
' 変      数:なし
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2007/04/11 岩田
' 変      更: 2008/09/30 西村　北九州高専の場合　欠課の評価は出力しない
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
	'エラー系
    Dim m_bErrFlg				'//ｴﾗｰﾌﾗｸﾞ
    
    Const C_ERR_GETDATA = "データの取得に失敗しました"
    
    Dim m_iNendo				'//年度
    Dim m_sKyokanCd				'//教官コード
    Dim m_sGakunen				'//学年
    Dim m_sClass				'//クラス

    Dim m_sBunruiCD		 		'//分類コード
    Dim m_sBunruiNM		 		'//分類名称
    Dim m_sTani		 			'//単位

    Dim m_lDataCount,m_uData()			'//評価データ
    Dim m_rCnt					'//レコードカウント
    Dim m_Rs
	
    Dim m_iSeisekiInpType
	Public m_sGakkoNO       '学校番号  INS 2008/09/30
	Public m_iMaxTem        '欠課最大点数  INS 2008/09/30

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
	Dim w_iRet
	Dim w_sSQL
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
	'Message用の変数の初期化
	w_sWinTitle = "キャンパスアシスト"
	w_sMsgTitle = "放送大学成績登録"
	w_sMsg = ""
	w_sRetURL = C_RetURL & C_ERR_RETURL
	w_sTarget = ""
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = false
	
	Do
		'//ﾃﾞｰﾀﾍﾞｰｽ接続
		If gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//ﾊﾟﾗﾒｰﾀSET
		Call s_SetParam()

		'//成績入力方法の取得(0:点数[C_SEISEKI_INP_TYPE_NUM]、1:文字[C_SEISEKI_INP_TYPE_STRING]、2:欠課、遅刻[C_SEISEKI_INP_TYPE_KEKKA])
		if not gf_GetKamokuSeisekiInp(m_iNendo,m_sBunruiCd,C_KAMOKUBUNRUI_NINTEI,m_iSeisekiInpType) then 
			m_bErrFlg = True
			Exit Do
		end if

		'//学校番号の取得
		if Not gf_GetGakkoNO(m_sGakkoNO) then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//成績、学生データ取得
		If not f_GetStudent() Then m_bErrFlg = True : Exit Do
		
		If m_Rs.EOF Then
			Call gs_showWhitePage("個人履修データが存在しません。","放送大学成績登録")
			Exit Do
		End If

		'//評価データの取得
		if not f_GetKamokuHyokaData(m_iNendo,m_sBunruiCd,C_KAMOKUBUNRUI_NINTEI,m_lDataCount,m_uData) then
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
		
		if w_sMsg = "" then w_sMsg = C_ERR_GETDATA
		
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// 終了処理
	Call gf_closeObject(m_Rs)
	
	Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*	[機能]	全項目に引き渡されてきた値を設定
'********************************************************************************
Sub s_SetParam()
	
    	m_iNendo    	= session("NENDO")		'年度
    	m_sKyokanCd 	= session("KYOKAN_CD")		'教官コード

	m_sGakunen  	= request("txtGakunen")         '学年
	m_sClass    	= request("txtClass")           'クラス
	m_sBunruiCD 	= request("txtBunruiCd")	'分類コード
	m_sBunruiNm 	= request("txtBunruiNm")	'分類名称
	m_sTani     	= request("txtTani")		'単位

End Sub

'********************************************************************************
'*	[機能]	データの取得
'********************************************************************************
Function f_GetStudent()
	
	Dim w_sSQL
	
	On Error Resume Next
	Err.Clear
	
	f_GetStudent = false

	'//検索結果の値より一覧を表示
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " T13_GAKUSEI_NO  AS GAKUSEI_NO,"
	w_sSQL = w_sSQL & " T13_GAKUSEKI_NO AS GAKUSEKI_NO,"
	w_sSQL = w_sSQL & " T13_GAKKA_CD    AS GAKKA_CD,"
	w_sSQL = w_sSQL & " T13_CLASS    　 AS CLASS,"
	w_sSQL = w_sSQL & " T13_COURCE_CD   AS COURCE_CD,"

	w_sSQL = w_sSQL & " T11_SIMEI       AS SIMEI, "

    	w_sSQL = w_sSQL & " T100_HAITOTANI  AS HAITOTANI,"
    	w_sSQL = w_sSQL & " T100_HYOTEI     AS HYOTEI,"
    	w_sSQL = w_sSQL & " T100_HYOKA      AS HYOKA,"
    	w_sSQL = w_sSQL & " T100_NINTEIBI   AS NINTEIBI,"
    	w_sSQL = w_sSQL & " T100_SYUTOKU_NENDO AS SYUTOKU_NENDO,"
    	w_sSQL = w_sSQL & " T100_SEISEKI    AS SEISEKI,"
    	w_sSQL = w_sSQL & " T100_HYOKA_FUKA_KBN AS FUKA_KBN"

	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI,"
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN, "
	w_sSQL = w_sSQL & " 	T100_RISYU_NINTEI "
	
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "     T13_NENDO    = " & Cint(m_iNendo)
	w_sSQL = w_sSQL & " AND	T13_GAKUNEN  = " & Cint(m_sGakunen)
	w_sSQL = w_sSQL & " AND	T13_CLASS    = " & Cint(m_sClass)
    	w_sSQL = w_sSQL & " AND T13_GAKUSEI_NO = T11_GAKUSEI_NO"
    	w_sSQL = w_sSQL & " AND T13_GAKUSEI_NO = T100_GAKUSEI_NO(+)"
    	w_sSQL = w_sSQL & " AND T100_BUNRUI_CD(+) = '" & m_sBunruiCd & "'"

    
    	w_sSQL = w_sSQL & " ORDER BY"
    	w_sSQL = w_sSQL & " T13_GAKUSEKI_NO"

	
'response.write w_sSQL
'response.end


	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit function
	
	'//ﾚｺｰﾄﾞカウント取得
	m_rCnt = gf_GetRsCount(m_Rs)
	
	f_GetStudent = true
	
End Function

'********************************************************************************
'*  [機能]  テーブルサイズのセット
'********************************************************************************
Sub s_SetTableWidth(p_TableWidth)
	
	p_TableWidth = 610
	
End Sub

'*******************************************************************************
' 機　　能：科目評価リスト取得(放送大学専用)
' 返　　値：True/False
' 引　　数：p_iNendo - 年度(IN)
'        　 p_sKamokuCD - 科目コード(IN)
'               （認定科目の場合は分類コードを指定する）
'           p_sKamokuBunrui - 科目分類コード(IN)
'               C_KAMOKUBUNRUI_TUJYO = 通常科目
'               C_KAMOKUBUNRUI_NINTEI = 認定科目
'               C_KAMOKUBUNRUI_TOKUBETU = 特別科目
'           p_lDataCount -  評価データ件数(OUT)
'           p_uData - 評価データ(OUT)
'
' 機能詳細：指定の科目コードと点数からp_uData()に評価、評定、欠点科目を設定する
'           p_uData()は動的配列でcall元で宣言すること。（宣言は関数内で行う）
'           p_uData()の件数はp_lDataCountにセットされる。また、配列インデックスは
'           1 ～ p_lDataCountまでが有効。
' 備　　考：点数評価の科目対象
'           call例
'           ret = gf_GetKamokuHyokaData(m_iNendo, w_KamokuCD, C_KAMOKUBUNRUI_TUJYO, w_lConut, w_udata())
'
'           gf_GetKamokuHyokaData　より作成、f_GetHyokaDataをCallするように変更
'*******************************************************************************
Function f_GetKamokuHyokaData(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_lDataCount,p_uData)
    Dim w_iZokuseiCD         '科目属性
    Dim w_iHyokaNo
    
    On Error Resume Next
    
    f_GetKamokuHyokaData = False
    
    '科目属性取得
    If Not gf_GetKamokuZokusei(p_iNendo, p_sKamokuCD, p_sKamokuBunrui, w_iZokuseiCD) Then
        Exit Function
    End If
    '科目属性から評価NO取得
    w_iHyokaNo = gf_iGetHyokaNo(w_iZokuseiCD, p_iNendo)
    
    '評価NOから評価データ取得
    If Not f_GetHyokaData(p_iNendo, w_iHyokaNo, p_lDataCount, p_uData) Then
        Exit Function
    End If
	
    '評価NOから評価データ取得　INS 2008/09/30西村
	IF m_sGakkoNO = cstr(C_NCT_KITAKYU) then
		If Not f_GetKekkaMaxTen(p_iNendo, w_iHyokaNo) Then
	        Exit Function
	    End If
	end if
	

    f_GetKamokuHyokaData = True
             
End Function


'*******************************************************************************
' 機　　能：科目評価取得(放送大学専用)
' 返　　値：True/False
' 引　　数：p_iNendo - 年度(IN)
'        　 p_iHyokaNo - 評価NO(IN)
'           p_lDataCount - 件数(OUT)
'           p_uData - 評価データ(OUT)
'
' 機能詳細：点数から評価NOのp_uDataに評価、評定、欠点科目を設定する
' 備　　考：評価NOがすでに分かっている場合には直接callも可
'           評価NOが分からないときは、gf_GetKamokuTensuHyokaをcall
'           f_GetHyokaData　より作成、M08_HYOKA_TAISYO_KBN　をC_HYOKA_TAISHO_HOKAに変更
'*******************************************************************************
Function f_GetHyokaData(p_iNendo,p_iHyokaNo,p_lDataCount,p_uData)
    Dim w_oRecord
    Dim w_sSql
    Dim w_lIdx
    
    On Error Resume Next
    
    f_GetHyokaData = False
    
    p_lDataCount = 0
    
    w_sSql = ""
    w_sSql = w_sSql & " SELECT"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_MEI,"
    w_sSql = w_sSql & " 	M08_HYOTEI,"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_RYAKU"
    
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & " 	M08_HYOKAKEISIKI"
    
    w_sSql = w_sSql & " WHERE"
    w_sSql = w_sSql & " 	M08_HYOUKA_NO = " & p_iHyokaNo      		'評価NO
    w_sSql = w_sSql & " AND M08_NENDO = " & p_iNendo        			'年度
    w_sSql = w_sSql & " AND M08_HYOKA_TAISYO_KBN = " & C_HYOKA_TAISHO_HOKA     	'他大学

	' INS 2008/09/30 西村 欠点記号はコンボにセットしない
	IF m_sGakkoNO = cstr(C_NCT_KITAKYU) then
		w_sSql = w_sSql & " AND M08_HYOKA_SYOBUNRUI_RYAKU = 0 "
	END IF
    
    w_sSql = w_sSql & " ORDER BY"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_CD"
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '科目Mない時エラー
    If w_oRecord.EOF Then
        Exit Function
    End If
    
    p_lDataCount = gf_GetRsCount(w_oRecord)
    
    '配列データ宣言
    ReDim p_uData(p_lDataCount,3)
    w_lIdx = 0
    
    Do Until w_oRecord.EOF
        
        'データセット
        p_uData(w_lIdx,0) = w_oRecord("M08_HYOKA_SYOBUNRUI_MEI")	'評価
        p_uData(w_lIdx,1) = w_oRecord("M08_HYOTEI")			'評定
        p_uData(w_lIdx,2) = w_oRecord("M08_HYOKA_SYOBUNRUI_RYAKU")	
        
        w_lIdx = w_lIdx + 1
        w_oRecord.MoveNext
    Loop
    
    Call gf_closeObject(w_oRecord)
    
    f_GetHyokaData = True

End Function

'*******************************************************************************
' 機　　能：欠課科目のMax点数(放送大学専用)
' 返　　値：True/False
' 引　　数：p_iNendo - 年度(IN)
'        　 p_iHyokaNo - 評価NO(IN)
'
' 機能詳細：欠点科目の最大点数を取得
' 備　　考：
'*******************************************************************************
Function f_GetKekkaMaxTen(p_iNendo,p_iHyokaNo)

    Dim w_oRecord
    Dim w_sSql
    Dim w_lIdx
    
    On Error Resume Next
    
    f_GetKekkaMaxTen = False
    
	m_iMaxTem = 0   
    w_sSql = ""
    w_sSql = w_sSql & " SELECT"
    w_sSql = w_sSql & " 	MAX(M08_MAX) MAX_TEN"
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & " 	M08_HYOKAKEISIKI"
    w_sSql = w_sSql & " WHERE"
    w_sSql = w_sSql & " 	M08_HYOUKA_NO = " & p_iHyokaNo      		'評価NO
    w_sSql = w_sSql & " AND M08_NENDO = " & p_iNendo        			'年度
    w_sSql = w_sSql & " AND M08_HYOKA_TAISYO_KBN = " & C_HYOKA_TAISHO_HOKA     	'他大学
    w_sSql = w_sSql & " AND M08_HYOKA_SYOBUNRUI_RYAKU = 1 "
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '科目Mない時エラー
    If w_oRecord.EOF Then
        Exit Function
    End If
    
    
    'データセット
    m_iMaxTem = w_oRecord("MAX_TEN")
    
    Call gf_closeObject(w_oRecord)

f_GetKekkaMaxTen = true

End Function


'********************************************************************************
'*  [機能]  HTMLを出力
'********************************************************************************
Sub showPage()
	DIm w_cell
	DIm w_Padding
	DIm w_Padding2

	Dim i
	Dim ii

	DIm w_sValue	''コンボVALUE部分

	Dim w_sInputClass
	
	Dim w_Disabled
	Dim w_Disabled2
	Dim w_TableWidth

	'初期設定
	w_Padding = "style='padding:2px 0px;'"
	w_Padding2 = "style='padding:2px 0px;font-size:10px;'"
	
	i = 1

	'//NN対応
	If session("browser") = "IE" Then
		w_sInputClass  = "class='num'"
	Else
		w_sInputClass = ""
	End If
	
	'//テーブルサイズのセット
	Call s_SetTableWidth(w_TableWidth)
	
%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--
	//************************************************************
	//  [機能]  ページロード時処理
	//************************************************************
	function window_onload() {
//		//スクロール同期制御
//		parent.init();
//		
//		document.frm.target = "topFrame";
//		document.frm.action = "sei0160_top.asp";
//		document.frm.submit();
	}
	//************************************************************
	//  [機能]  登録ボタンが押されたとき
	//************************************************************
	function f_Touroku(){
		if(!f_InpCheck()){
		//	alert("入力値が不正です");
			return false;
		}

		if(!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return false;}
		
		//ヘッダ部空白表示
		parent.topFrame.document.location.href="white.asp";
		
		//登録処理
		document.frm.action="sei0160_upd.asp";
		document.frm.target="main";
		document.frm.submit();
	}
	
	//************************************************************
	//	[機能]	キャンセルボタンが押されたとき
	//************************************************************
	function f_Cancel(){
		parent.document.location.href="default.asp";
	}

	//************************************************
	//	入力チェック
	//************************************************
	function f_InpCheck(){
		var w_length;
		var ob;
		
		w_length = document.frm.elements.length;
		
		for(i=0;i<w_length;i++){
			ob = eval("document.frm.elements[" + i + "]")
			
			if(ob.type=="text"){
				ob = eval("document.frm." + ob.name);
				
				if(!f_CheckNum(ob)){
						alert("入力値が不正です");
						return false;}
			}


		}


    	<% IF m_sGakkoNO = cstr(C_NCT_KITAKYU) then %>
		for(i=1;i < <%=m_rCnt%> + 1 ;i++){

				ob = eval("document.frm.Seiseki" + i);
				if(ob.value.length == 0){
				
				}
				else
				{
				if(!f_CheckKekka(ob)){
					
					alert(<%=m_iMaxTem%> + "点以下は入力できません");
					return false;	
				}
				}
		}
		<% END IF%>


		return true;
	}

	//************************************************************
	//  [機能]  数値型チェック
	//************************************************************
	function f_CheckNum(pFromName){
		var wFromName,w_len;
		
		wFromName = eval(pFromName);
		
		if(isNaN(wFromName.value)){
			wFromName.focus();
			wFromName.select();
			return false;
		}else{
			//桁チェック
			if(wFromName.name.indexOf("Seiseki") != -1){
				if(wFromName.value > 100){
					wFromName.focus();
					wFromName.select();
					return false;
				}
			}

			//修得年度は、4桁まで
			if(wFromName.name.indexOf("SyuNendo") != -1){
				w_len = 4;
			}else{
				w_len = 3;
			}
			
			if(wFromName.value.length > w_len){
				wFromName.focus();
				wFromName.select();
				return false;
			}

			//マイナスをチェック
			var wStr = new String(wFromName.value)
			if (wStr.match("-")!=null){
				wFromName.focus();
				wFromName.select();
				return false;
			}
		}

		return true;
	}

	//************************************************************
	//  [機能]  欠課点チェック INS2008/09/30西村
	//************************************************************
	function f_CheckKekka(pFromName){
		var wFromName,w_len;

		var wFromName,w_len;
		
		wFromName = eval(pFromName);
		
		if(isNaN(wFromName.value)){
			wFromName.focus();
			wFromName.select();
			return false;
		}else{

			if(wFromName.value <= <%=m_iMaxTem%>){
				wFromName.focus();
				wFromName.select();
				return false;
			}
		}
		return true;
	}

	
	//************************************************
	//Enter キーで下の入力フォームに動くようになる
	//引数：p_inpNm	対象入力フォーム名
	//    ：p_frm	対象フォーム
	//　　：i		現在の番号
	//戻値：なし
	//入力フォーム名が、xxxx1,xxxx2,xxxx3,…,xxxxn 
	//の名前のときに利用できます。
	//************************************************
	function f_MoveCur(p_inpNm,p_frm,i){
		if (event.keyCode == 13){		//押されたキーがEnter(13)の時に動く。
			i++;
			
			//入力可能のテキストボックスを探す。見つかったらフォーカスを移して処理を抜ける。
	        for (w_li = 1; w_li <= 99; w_li++) {
				
				if (i > <%=m_rCnt%>) i = 1; //iが最大値を超えると、はじめに戻る。
				inpForm = eval("p_frm."+p_inpNm+i);
				
				//入力可能領域ならフォーカスを移す。
				if (typeof(inpForm) != "undefined") {
					inpForm.focus();			//フォーカスを移す。
					inpForm.select();			//移ったテキストボックス内を選択状態にする。
					break;
				//入力付加なら次の項目へ
				} else{
					i++
				}
	        }
		}else{
			return false;
		}
		return true;
	}

	//************************************************
	//	評価コンボ入力時の処理
	//	
	//************************************************
	function f_ChgHyoka(w_num){
		var ob = new Array();
		ob[0] = eval("document.frm.sltHyoka" + w_num);

		ob[1] = eval("document.frm.hidHyoka" + w_num);
		ob[2] = eval("document.frm.hidHyotei" + w_num);
		ob[3] = eval("document.frm.hidHyokaFukaKbn" + w_num);
		ob[4] = eval("document.frm.SyuNendo" + w_num);

		if(ob[0].value.length == 0||ob[0].value =="@@@"){
			ob[1].value = "";
			ob[2].value = "";
			ob[3].value = "";
			ob[4].value = "";
		}else{
			var vl = ob[0].value.split('#@#');
			
			ob[1].value = vl[0];
			ob[2].value = vl[1];
			ob[3].value = vl[2];
			
			//合格で修得年度が未入力のとき処理年度を表示する
			if(ob[3].value=="0"){
				if(ob[4].value==""){
					ob[4].value ="<%=m_iNendo%>";
				}
			}else{
				ob[4].value ="";
			}
		}
	}

	//************************************************
	//	処理年度入力時の処理
	//	
	//************************************************
	function f_ChkSyoriNendo(w_num){
		var ob = new Array();
		ob[0] = eval("document.frm.SyuNendo" + w_num);
		ob[1] = eval("document.frm.hidHyokaFukaKbn" + w_num);

		if(ob[1].value=="0"){
			//評価不可区分が"可" のとき処理年度はを表示する
			if(ob[0].value==""){
				ob[0].value ="<%=m_iNendo%>";
			}
		}else{
			//評価不可区分が"可" 以外のとき処理年度は入力できない
			ob[0].value ="";
		}

	}
	
	//-->
	</SCRIPT>
	</head>
	<body LANGUAGE="javascript" onload="window_onload();">
	<center>
	<form name="frm" method="post">
	
	<table width="<%=w_TableWidth%>">
	<tr>
	<td>
	
	<table class="hyo" align="center" width="<%=w_TableWidth%>" border="1">
		<tr>
			<th class="header" width="80" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
			<th class="header" width="300"氏　名</th>
			<th class="header" width="50" nowrap>成績</th>
			<th class="header" width="50" nowrap>評価</th>
			<th class="header" width="100" nowrap>修得年度</th>
		</tr>
	<%
	m_Rs.MoveFirst

	i = 0
	
	Do Until m_Rs.EOF

		i = i + 1
		
		Call gs_cellPtn(w_cell)
	%>
			
		<tr>
			<td class="<%=w_cell%>" align="center" width="80"  nowrap <%=w_Padding%>><%=m_Rs("GAKUSEKI_NO")%></td>
			<input type="hidden" name="txtGsekiNo<%=i%>"   value="<%=m_Rs("GAKUSEKI_NO")%>">
			<input type="hidden" name="txtGseiNo<%=i%>"    value="<%=m_Rs("GAKUSEI_NO")%>">
			<input type="hidden" name="txtGakkaCD<%=i%>"   value="<%=m_Rs("GAKKA_CD")%>">
			<input type="hidden" name="txtClass<%=i%>"     value="<%=m_Rs("CLASS")%>">
			<input type="hidden" name="txtCorceCD<%=i%>"   value="<%=m_Rs("COURCE_CD")%>">
			<td class="<%=w_cell%>" align="left" width="300" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%></td>
				

			<!-- 成績  -->
			<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>
                        	<input type="text" <%=w_sInputClass%>  name="Seiseki<%=i%>" value="<%=trim(m_Rs("SEISEKI"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);">
                        </td>

				
			<!-- 評価 -->
			<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>
				<select name="sltHyoka<%=i%>" onchange ="f_ChgHyoka(<%=i%>);">
					<Option Value="@@@">

		<% 	For ii = 0 to  m_lDataCount-1 

				'コンボVALUE部分生成
				w_sValue = ""
				w_sValue = w_sValue & m_uData(ii,0) & "#@#"
				w_sValue = w_sValue & m_uData(ii,1) & "#@#"
				w_sValue = w_sValue & m_uData(ii,2) 
		%>
				<% if trim(m_Rs("HYOKA")) = gf_SetNull2String(m_uData(ii,0)) then %>
					<option value="<%=w_sValue%>" selected> <%=m_uData(ii,0)%></option>
				<% Else %>
					<option value="<%=w_sValue%>"> <%=m_uData(ii,0)%></option>
				<% end if 
		    	NEXT 
		%>
				</select>
			</td>

			<!-- 取得年度 -->
			<td class="<%=w_cell%>" align="center" width="100" nowrap <%=w_Padding%>>
	                	<input type="text" <%=w_sInputClass%> name="SyuNendo<%=i%>" value="<%=trim(m_Rs("SYUTOKU_NENDO"))%>" size=4 maxlength=4 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);"onBlur ="f_ChkSyoriNendo(<%=i%>);">
                        </td>

			<input type="hidden" name="hidHyoka<%=i%>"           value="<%=m_Rs("HYOKA")%>">
			<input type="hidden" name="hidHyotei<%=i%>"          value="<%=m_Rs("HYOTEI")%>">
			<input type="hidden" name="hidHyokaFukaKbn<%=i%>"    value="<%=m_Rs("FUKA_KBN")%>">
		</tr>
	<%
		m_Rs.MoveNext
	Loop
	%>
			
			
	</table>
	
	</td>
	</tr>
	
	<tr>
	<td align="center">
	<table>
		<tr>
			<td align="center" align="center" colspan="13">
					<input type="button" class="button" value="　登　録　" onClick="f_Touroku();">　
					<input type="button" class="button" value="キャンセル" onClick="f_Cancel();">
			</td>
		</tr>
	</table>
	</td>
	</tr>

	</table>
	
	
	<input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">

	<input type="hidden" name="txtGakunen"   value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass"     value="<%=m_sClass%>">
	<input type="hidden" name="txtBunruiCd"  value="<%=m_sBunruiCD%>">
	<input type="hidden" name="txtBunruiNm"  value="<%=m_sBunruiNM%>">
	<input type="hidden" name="txtTani"      value="<%=m_sTani%>">


	<input type="hidden" name="i_Max"        value="<%=i%>">
	
	</form>
	</center>
	</body>
	</html>
<%
End sub
%>