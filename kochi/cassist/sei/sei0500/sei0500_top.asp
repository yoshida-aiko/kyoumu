<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 実力試験成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0500/sei0500_top.asp
' 機      能: 上ページ 成績登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:

'-------------------------------------------------------------------------
' 作      成: 2001/09/06 モチナガ
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    Public m_iNendo             '年度
    Public m_sKyokanCd          '教官コード
    Public m_iSikenCD			'試験CD

    Public m_Rs_Siken			'試験情報を取得
    Public m_Rs					'学年、クラス、科目取得RS

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
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="実力試験成績登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do
        '// ﾃﾞｰﾀﾍﾞｰｽ接続
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If

		'//値を取得
		call s_SetParam()

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

		'//試験コンボを取得
        w_iRet = f_GetSiken()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//試験コードがNULLだったら、コンボのはじめの試験コードを入れる
		if gf_IsNull(m_iSikenCd) then m_iSikenCd = m_Rs_Siken("M28_SIKEN_CD")

		if Not gf_IsNull(m_iSikenCd) then

			'//学年・クラス・科目コンボを取得
			w_iRet = f_GetKamoku()
			If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		End if

       '// ページを表示
       Call showPage()
       Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    Call gf_closeObject(m_Rs_Siken)
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iSikenCd  = Request("txtSikenCD")    '//コンボ試験区分

End Sub

'********************************************************************************
'*  [機能]  試験コンボを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSiken()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKENMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD "
		w_sSQL = w_sSQL & vbCrLf & "  FROM  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28_SIKEN_KAMOKU M28,  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD         = M27.M27_SIKEN_CD AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KAMOKU     = M27.M27_SIKEN_KAMOKU AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KBN        = M27.M27_SIKEN_KBN AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_NENDO            = M27.M27_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KBN        =  " & C_SIKEN_JITURYOKU & " AND "	'試験区分(実力試験のみ)
		w_sSQL = w_sSQL & vbCrLf & "  	(M28.M28_SEISEKI_KYOKAN1 = '" & m_sKyokanCd & "' OR "		'入力教官1
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN2 = '" & m_sKyokanCd & "' OR "		'入力教官2
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN3 = '" & m_sKyokanCd & "' OR "		'入力教官3
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN4 = '" & m_sKyokanCd & "' OR "		'入力教官4
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN5 = '" & m_sKyokanCd & "' ) AND "	'入力教官5
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_NENDO            =  " & m_iNendo					'処理年度
		w_sSQL = w_sSQL & vbCrLf & "  GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKENMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD "
'Response.Write w_ssql & "<br>"
        iRet = gf_GetRecordset(m_Rs_Siken, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If

        f_GetSiken = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  学年・クラス・科目コンボを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetKamoku()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetKamoku = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD,"
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KAMOKU,"
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_GAKUNEN,  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_CLASS,  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_KAMOKUMEI "
		w_sSQL = w_sSQL & vbCrLf & "  FROM  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27, "
		w_sSQL = w_sSQL & vbCrLf & "  	M28_SIKEN_KAMOKU M28 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KAMOKU     = M28.M28_SIKEN_KAMOKU AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_CD         = M28.M28_SIKEN_CD AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KBN        = M28.M28_SIKEN_KBN AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO            = M28.M28_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO            =  " & m_iNendo & " AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KBN        =  " & C_SIKEN_JITURYOKU & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD         =  " & m_iSikenCD  & " AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	(M28.M28_SEISEKI_KYOKAN1 = '" & m_sKyokanCd & "' OR "
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN2 = '" & m_sKyokanCd & "' OR "
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN3 = '" & m_sKyokanCd & "' OR "
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN4 = '" & m_sKyokanCd & "' OR "
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN5 = '" & m_sKyokanCd & "' ) "

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKamoku = 99
            Exit Do
        End If

        f_GetKamoku = 0
        Exit Do
    Loop

End Function

'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'[引数] pData1 : データ１
'       pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'****************************************************
Function f_Selected(pData1,pData2)

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else 
            f_Selected = "" 
        End If
    End If

End Function

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
	//************************************************************
	//  [機能]  試験が変更されたとき、再表示する
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_ReLoadMyPage(){

	    document.frm.action="sei0500_top.asp";
	    document.frm.target="topFrame";
	    document.frm.submit();

	}

	//************************************************************
	//  [機能]  表示ボタンクリック時の処理
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_Search(){

	    // ■■■NULLﾁｪｯｸ■■■
	    // ■学年
	    if( f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
	        window.alert("学年の選択を行ってください");
	        document.frm.txtGakuNo.focus();
	        return ;
	    }
	    // ■クラス
	    if( f_Trim(document.frm.txtClassNo.value) == "<%=C_CBO_NULL%>" ){
	        window.alert("クラスの選択を行ってください");
	        document.frm.txtClassNo.focus();
	        return ;
	    }

	    // ■科目名
	    if( f_Trim(document.frm.txtKamokuCd.value) == "<%=C_CBO_NULL%>" ){

			if (document.frm.txtKamokuCd.length ==1){
		        window.alert("試験科目がありません");
		        return ;
			}else{
		        window.alert("科目の選択を行ってください");
		        document.frm.txtKamokuCd.focus();
		        return ;
			}
	    }

		// 選択されたコンボの値をｾｯﾄ
		iRet = f_SetData();
		if( iRet != 0 ){
			return;
		}

	    document.frm.action="sei0500_bottom.asp";
	    document.frm.target="main";
	    document.frm.submit();

	}

	//************************************************************
	//  [機能]  表示ボタンクリック時に選択されたデータをｾｯﾄ
	//  [引数]  なし
	//  [戻値]  なし
	//  [説明]
	//
	//************************************************************
	function f_SetData(){

        if (document.frm.cboKamoku.value==""){
            alert("科目データがありません。")
            return 1;
        };

		//データ取得
        var vl = document.frm.cboKamoku.value.split('$$$');

        //選択されたデータをｾｯﾄ(学年、クラス、科目CDを取得)
        document.frm.txtGakuNo.value=vl[0];
        document.frm.txtClassNo.value=vl[1];
        document.frm.txtKamokuCd.value=vl[2];

        return 0;
	}

	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>

	<body>
	<center>
	<form name="frm" METHOD="post">

	<% call gs_title(" 実力試験成績登録 "," 登　録 ") %>
	<br>
	<table border="0">
	    <tr><td valign="bottom">

	        <table border="0" width="100%">
	            <tr><td class="search">

	                <table border="0">
	                    <tr valign="middle">
	                        <td align="left" nowrap>試験区分</td>
	                        <td align="left" colspan="3">
								<%If m_Rs_Siken.EOF Then%>
									<select name="txtSikenCD" style='width:150px;' DISABLED>
										<option value="">データがありません
								<%Else%>
									<select name="txtSikenCD" style='width:150px;' onchange = 'javascript:f_ReLoadMyPage()'>
									<%Do Until m_Rs_Siken.EOF%>
										<option value='<%=m_Rs_Siken("M28_SIKEN_CD")%>'  <%=f_Selected(cstr(m_Rs_Siken("M28_SIKEN_CD")),cstr(m_iSikenCD))%>><%=m_Rs_Siken("M27_SIKENMEI")%>
										<%m_Rs_Siken.MoveNext%>
									<%Loop%>
								<%End If%>
								</select>
							</td>
	                        <td>&nbsp;</td>

	                        <td align="left" nowrap>科目</td>
	                        <td align="left">
								<%If m_iSikenCd = "" Then%>
									<select name="cboKamoku" style='width:200px;' DISABLED>
										<option value="">データがありません
								<%Else%>
									<%If m_Rs.EOF Then%>
										<select name="cboKamoku" style='width:200px;' DISABLED>
											<option value="">科目データがありません
									<%Else%>
										<select name="cboKamoku" style='width:200px;'>
										<%Do Until m_Rs.EOF%>
											<%
											wSikenCd   = m_Rs("M28_SIKEN_CD") 

											'//表示内容を作成
											w_Str=""
											w_Str= w_Str & m_Rs("M28_GAKUNEN") & "年　"
											w_Str= w_Str & gf_GetClassName(m_iNendo,m_Rs("M28_GAKUNEN"),m_Rs("M28_CLASS")) & "　"
											w_Str= w_Str & m_Rs("M27_KAMOKUMEI") & "　"
											%>
											<option value="<%=m_Rs("M28_GAKUNEN")%>$$$<%=m_Rs("M28_CLASS")%>$$$<%=m_Rs("M28_SIKEN_KAMOKU")%>"><%=w_Str%>
											<%m_Rs.MoveNext%>
										<%Loop%>
									<%End If%>
								<%End If%>
								</select>
							</td>
	                    </tr>
						<tr>
					        <td colspan="7" align="right">
					        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
					        </td>
						</tr>
	                </table>
	            </td>
				</tr>
	        </table>
	        </td>
	    </tr>
	</table>

	<input type="hidden" name="txtNendo"     value="<%= m_iNendo %>">
	<input type="hidden" name="txtKyokanCd"  value="<%= m_sKyokanCd %>">
	<input type="hidden" name="txtShikenCd"  value="<%= wSikenCd   %>">

	<input type="hidden" name="txtGakuNo"    value="">
	<input type="hidden" name="txtClassNo"   value="">
	<input type="hidden" name="txtKamokuCd"  value="">

	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>