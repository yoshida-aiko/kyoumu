<%@ Language=VBScript %>
<%
'*************************************************************************
'* システム名: 教務事務システム
'* 処  理  名: 留年該当者一覧
'* ﾌﾟﾛｸﾞﾗﾑID : han/han0121/ichiran.asp
'* 機      能: 下ページ 留年該当者一覧リスト表示を行う
'*-------------------------------------------------------------------------
'* 引      数:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'*           cboGakunenCd      :学年コード
'*           :session("PRJ_No")      '権限ﾁｪｯｸのキー
'* 変      数:なし
'* 引      渡:教官コード     ＞      SESSIONより（保留）
'*           :処理年度       ＞      SESSIONより（保留）
'* 説      明:
'*           選択されたクラスの留年該当者一覧を表示
'*-------------------------------------------------------------------------
'* 作      成: 2001/08/08 前田　智史
'* 変      更: 2014/07/29 藤林　聡子 VB版に合わせて区分を表示し、ソート順を指定する。ページあたりの表示件数を20件に。
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_sMsg              'ﾒｯｾｰｼﾞ
    
    '取得したデータを持つ変数
    Public  m_iNendo         ':処理年度
    Public  m_iKyokanCd         ':教官コード
    Public  m_iGakunen          ':学年コード
    
    Public  m_Rs                'recordset
    Public  m_sMode             'モード
    
    'ページ関係
    Public  m_iMax              ':最大ページ
    Public  m_iDsp              '// 一覧表示行数
    Public  m_iPageCD


	Public m_sSINKYU
	Public m_sTAIGAKU
	Public m_sRYUUNEN
	Public m_sSYUTAI
	Public m_sSYUTAIGAKU
	
	Public Const m_C_PAGE_LINE20 = 20		'1ページあたりの表示件数 --2014/07/29 INSERT FUJIBAYASHI

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
    Dim w_sWHERE            '// WHERE文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//レコードカウント用

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="留年該当者一覧"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = m_C_PAGE_LINE20

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

		'//リストの詳細データ取得
		w_iRet = f_getdate()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

		If m_Rs.EOF Then
	        '// ページを表示
	        Call NO_showPage()
		Else
	        '// ページを表示
	        Call showPage()
		End If

        Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()    '//2001/07/30変更

    m_iKyokanCd = Request("txtKyokanCd")          ':教官コード
    m_iNendo = Request("txtNendo")              ':処理年度
    m_iGakunen = Request("cboGakunenCd")   ':学年コード
    m_sMode = Request("txtMode")

    If m_sMode = "Hyouji" Then
        m_iPageCD = 1
    Else
        m_iPageCD = INT(Request("txtPageCd")) ':表示済表示頁数（自分自身から受け取る引数）
    End If
    
End Sub

'********************************************************************************
'*  [機能]  リストの詳細取得
'*  [引数]  
'*  [戻値]  0:情報取得成功、1:レコードなし、99:失敗
'*  [説明]  
'********************************************************************************
Function f_getdate()
    
    On Error Resume Next
    Err.Clear
    
    f_getdate = 1

    Do

        '// クラスマスタを取得
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & "A.T48_GAKUSEKI_NO,A.T48_SIMEI,B.M02_GAKKAMEI "
        '--2014/07/29 INSERT FUJIBAYASHI
        w_sSQL = w_sSQL & vbCrLf & ",A.T48_SINKYU,A.T48_TAIGAKU,A.T48_RYUUNEN,A.T48_SYUTAI,A.T48_SYUTAIGAKU "
        '--2014/07/29 INSERT END
        w_sSQL = w_sSQL & vbCrLf & "FROM "
        w_sSQL = w_sSQL & vbCrLf & "T48_RYUNEN A,M02_GAKKA B "
        w_sSQL = w_sSQL & vbCrLf & "WHERE "
        w_sSQL = w_sSQL & vbCrLf & "A.T48_NENDO = " & m_iNendo & " "
        w_sSQL = w_sSQL & vbCrLf & "AND A.T48_GAKUNEN = " & m_iGakunen & " "
        w_sSQL = w_sSQL & vbCrLf & "AND A.T48_NENDO = B.M02_NENDO(+) "
        w_sSQL = w_sSQL & vbCrLf & "AND A.T48_GAKKA_CD = B.M02_GAKKA_CD(+) "
        '--2014/07/29 INSERT FUJIBAYASHI
        w_sSQL = w_sSQL & vbCrLf & "ORDER BY "
        w_sSQL = w_sSQL & vbCrLf & "         A.T48_GAKUNEN"
        w_sSQL = w_sSQL & vbCrLf & "        ,A.T48_GAKKA_CD"
        w_sSQL = w_sSQL & vbCrLf & "        ,A.T48_GAKUSEKI_NO"
		'--2014/07/29 INSERT END
		
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            f_getdate = 99
            Exit Do 'GOTO LABEL_f_GetClassMei_END
        End If

		f_getdate = 0

        Exit Do
    
    Loop
    

'// LABEL_f_GetClassMei_END
End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_iCnt

	w_iCnt = 0
    
    '--2014/07/29 INSERT FUJIBAYASHI(ページあたりの表示件数を指定値に合わせるため追加)
    w_iCnt  = 1
    w_bFlg  = True
	'--2014/07/29 INSERT FUJIBAYASHI
	    
    On Error Resume Next
    Err.Clear

    'ページBAR表示
    Call gs_pageBar(m_Rs,m_iPageCD,m_iDsp,w_pageBar)

%>

<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    //************************************************************
    //  [機能]  一覧表の次・前ページを表示する
    //  [引数]  p_iPage :表示頁数
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageCd.value = p_iPage;
        document.frm.submit();
    
    }
-->
</SCRIPT>
</head>

<body>

<center>
<form name ="frm" method="post">

<!--2014/07/29 UPDATE FUJIBAYASHI 『width="500" →width="600"』-->
<table border=0 width="600">
<tr>
<td align="center">
<%=w_pageBar %>

	<table border="1" class=hyo width="<%=C_TABLE_WIDTH%>">
		<tr>
			<!--2014/07/29 INSERT FUJIBAYASHI -->
			<th class=header width="15" height="40">進級</th>
			<th class=header width="15">留年</th>
			<th class=header width="15">退学</th>
			<th class=header width="15">修退</th>
			<th class=header width="35">修了退学</th>
			<!--2014/07/29 INSERT END -->

			<th class=header width="140">学 科 名</th>
			
			<!--2014/07/29 UPDATE FUJIBAYASHI 『width="100" →width="65"』-->
			<th class=header width="65"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
			<th class=header width="120">氏　　名</th>
		</tr>

<%
	
    Do While(w_bFlg)
        '//テーブルセル背景色
        call gs_cellPtn(w_cell)
        
        '--2014/07/29 INSERT FUJIBAYASHI
        '//各フラグが1の場合、記号に『○』を設定する
        If m_Rs("T48_SINKYU") = "1" Then
        	m_sSINKYU = "○"
        Else
        	m_sSINKYU = ""
        End if
        If m_Rs("T48_RYUUNEN") = "1" Then
        	m_sRYUUNEN = "○"
        Else
        	m_sRYUUNEN = ""
        End if
        If m_Rs("T48_TAIGAKU") = "1" Then
        	m_sTAIGAKU = "○"
        Else
        	m_sTAIGAKU = ""
        End if
        If m_Rs("T48_SYUTAI") = "1" Then
        	m_sSYUTAI = "○"
        Else
        	m_sSYUTAI = ""
        End if
        If m_Rs("T48_SYUTAIGAKU") = "1" Then
        	m_sSYUTAIGAKU = "*"
        Else
        	m_sSYUTAIGAKU = ""
        End if
        '--2014/07/29 INSERT END
%>
		<tr>
			
			<!--2014/07/29 INSERT FUJIBAYASHI -->
			<td align="center" class=<%=w_cell%>><%=m_sSINKYU %><BR></td>
			<td align="center" class=<%=w_cell%>><%=m_sRYUUNEN %><BR></td>
			<td align="center" class=<%=w_cell%>><%=m_sTAIGAKU %><BR></td>
			<td align="center" class=<%=w_cell%>><%=m_sSYUTAI %><BR></td>
			<td align="center" class=<%=w_cell%>><%=m_sSYUTAIGAKU %><BR></td>
			<!--2014/07/29 INSERT END -->
			
			<td class=<%=w_cell%>><%=m_Rs("M02_GAKKAMEI")%><BR></td>
			<td class=<%=w_cell%>><%=m_Rs("T48_GAKUSEKI_NO")%><BR></td>
			<td class=<%=w_cell%>><%=m_Rs("T48_SIMEI")%><BR></td>
		</tr>
<%
		m_Rs.MoveNext
		
        '--2014/07/29 UPDATE FUJIBAYASHI(ページあたりの表示件数を指定値に合わせるため)
		'If w_iCnt >= C_PAGE_LINE Then
		'	Exit Do
		'Else
		'	w_iCnt = w_iCnt + 1
        'End If
		If m_Rs.EOF Then
			w_bFlg = False
		ElseIf w_iCnt >= m_iDsp Then
			w_bFlg = False
		Else
			w_iCnt = w_iCnt + 1
        End If
        '--2014/07/29 UPDATE END
        
        
    Loop
%>
    </table>
<%=w_pageBar %>

</td>
</tr>
</table>
	<input type="hidden" name="txtMode" value="<%=m_sMode%>">
	<input type="hidden" name="txtPageCd" value="<%=m_iPageCD %>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_iKyokanCd%>">
	<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">
	<input type="hidden" name="cboGakunenCd" value="<%=m_iGakunen%>">

</form>
</center>

</body>

</html>
<%
    '---------- HTML END   ----------
End Sub

Sub NO_showPage()
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
   </head>

    <body>

    <center>
		<br><br><br>
		<span class="msg">対象データは存在しません。条件を入力しなおして検索してください。</span>
    </center>
    </body>

    </html>
<%
    '---------- HTML END   ----------
End Sub
%>
