<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 欠席日数登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/sei0600/sei0600_topDisp.asp
' 機      能: 
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:
'           ■初期表示
'               上部画面表示のみ
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/09/26 谷脇 良也
' 修      正: 2002/06/07 金澤香織
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系

    '市町村選択用のWhere条件
    Public m_iNendo           '年度
    Public m_sKyokanCd        '教官コード
    Public m_sGakunen         '学年
    Public m_sClass           'クラス
    Public m_sClassNm         'クラス名
    Public m_sGakusei()       '学生の配列
    Public m_sGakka       	  '学生の所属学科
    Public m_sShiken

    Public  m_GRs
    Public  m_Rs
    Public  m_iMax       	   '最大ページ
    Public  m_iDsp       	   '一覧表示行数
    Public  m_iTokuketu_Su     '特別欠席のレコード数 2002/6/7 金澤
    Public  m_sTokuketu_Na()   '特別欠席名称 2002/6/7 金澤
    
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
    w_sMsgTitle="欠席日数登録"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

	  '//試験名を取得
            If f_GetSiken(m_sShiken) <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If

		Call f_Gakusei()

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
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm    = request("txtClassNm")
	m_sShiken    = request("txtSikenKBN")

End Sub

'********************************************************************************
'*  [機能]  試験コンボを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSiken(p_sShiken)
    Dim w_sSQL,w_Rs

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & " M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_SYOBUNRUI_CD = " & cint(p_sShiken)

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If
	p_sShiken = w_Rs("M01_SYOBUNRUIMEI")

        f_GetSiken = 0
        Exit Do
    Loop
	Call gf_closeObject(w_Rs)
    f_GetTokuketuName
End Function

'********************************************************************************
'*  [機能]  特別欠席名称を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  2002/06/07	金澤追加
'********************************************************************************
Function f_GetTokuketuName()
    Dim w_sSQL,w_Rs2,i,w_sTENMP
    
    On Error Resume Next
    Err.Clear
    
    Do 
	'------------------------ kanazawa 2002/6/7--------------------------------------
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & "  M01_SYOBUNRUIMEI "
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "   MM01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_DAIBUNRUI_CD = " & C_M01_DAIBUNRUI150 '特別欠席
		w_sSQL = w_sSQL & vbCrLf & "  ORDER BY M01_SYOBUNRUI_CD "

'		response.write " =" & w_sSQL & "<BR>"     ' 2002/6/7 KANAZAWA 
'		response.end
	
        iRet = gf_GetRecordset_OpenStatic(w_Rs2, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            Exit Do
        End If
        
		m_iTokuketu_Su = cint(gf_GetRsCount(w_Rs2))
		
	'----------------------------------------------------------------------------------
        Exit Do
    Loop
       
    i = 1
    Redim m_sTokuketu_Na(m_iTokuketu_Su)
    
 '   w_Rs2.movelast
    
    do until w_Rs2.eof

        m_sTokuketu_Na(i) = w_Rs2("M01_SYOBUNRUIMEI")
		response.write err.description
        i = i + 1
		w_Rs2.movenext
    loop
    
	Call gf_closeObject(w_Rs2)

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
<link rel="stylesheet" href="../../common/style.css" type="text/css">

<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="JavaScript">
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
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){
        parent.main.f_Touroku();
    }

	//************************************************************
	//	[機能]	キャンセルボタンが押されたとき
	//	[引数]	なし
	//	[戻値]	なし
	//	[説明]
	//************************************************************
	function f_Cansel(){

        //初期ページを表示
        parent.document.location.href="default.asp"
	
	}


//-->
</SCRIPT>

</head>
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">
<center>
<%call gs_title("欠席日数登録","登　録")%>
<table border="0" width="300" class=hyo align="center">
	<tr>
		<th width="300" class="header2" colspan="2"><%=m_sShiken%></th>
	</tr>
	<tr>
		<th width="50" class="header">クラス</th>
		<td width="250" align="center" class="detail"><%=m_sGakunen%>年　<%=m_sClassNm%></td>
	</tr>
</table>
<br>
<div align="center"><span class=CAUTION>※「累計」は、日毎出欠入力メニューより日々入力された上記試験までの欠席状況です。<br>
</span></div>
	<table width="50%">
	<tr>
		<td align="center"><input type="button" class="button" value="　登　録　" onclick="javascript:f_Touroku()">　
		<input type="button" class="button" value="キャンセル" onclick="javascript:f_Cansel()"></td>
	</tr>
	</table>
<table border="1" cellpadding="1" cellspacing="1" class="hyo">
			<tr>
				<!--<th nowrap class="header" rowspan="2" width="50"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>   2002/6/10 kana -->
				<th nowrap class="header" rowspan="2" width="80"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
				<th nowrap class="header" rowspan="2" width="150">氏 名</th>
				<th nowrap class="header" colspan="2"><font size="2">欠 席</font></th>
				
				<!--'colspan も可変で対応-->
				<% if m_iTokuketu_Su <> 0 then %>
				<th nowrap class="header" colspan= <%=m_iTokuketu_Su%>><font size="2">特別欠席</font></th>
				<% end if %>
			</tr>
			<tr>
				<th nowrap class="header2" width="43"><font size="1">入力</font></th>
				<th nowrap class="header2" width="43"><font size="1">累計</font></th>
				<!--'以下ループで対応　金澤6/7-->
				<% Dim i %>
				<% FOR i = 1 TO m_iTokuketu_Su %>
					<th nowrap class="header2" width="43"><font size="1"><%= m_sTokuketu_Na(i) %></font></th>
				<%NEXT%>
				
			</tr>
</table>
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
