<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0112/kks0112_middle.asp
' 機      能: 上ページ 前ページの検索条件を表示
'-------------------------------------------------------------------------
' 引      数: 
'             
'             
'             
'             
' 変      数:
' 引      渡: 
'             
'             
'             
'             
' 説      明:
'           ■初期表示
'               日付：前ページの検索条件を表示
'               時限：前ページの検索条件を表示
'               科目：前ページの検索条件を表示
'               クラス：前ページの検索条件を表示
'               入力区分：欠課、遅刻、早退、クリアのラジオボタン
'           ■登録ボタンクリック時
'               入力情報を登録する
'           ■戻るボタンクリック時
'               前ページに戻る
'-------------------------------------------------------------------------
' 作      成: 2002/05/16 shin
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    '//エラー系
    Dim  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '//変数
    Dim m_iSyoriNen		'//処理年度
    Dim m_sDate			'//日付
	Dim m_iJigen		'//時限数
	Dim m_iGakunen		'//学年
	Dim m_sClassName	'//クラス名
	Dim m_sKamokuName	'//科目名
	Dim m_sClassNo		'//クラスNO
	Dim m_sKamokuCd		'//科目CD
	Dim m_iKamokuKbn	'//科目区分(0:通常授業、1:特別活動)
	
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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
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
            w_sMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
        '//変数初期化
        Call s_ClearParam()
		
        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()
		
		Call showPage_middle()
		
        Exit Do
    Loop
	
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
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
	
	m_sDate = ""
	m_iJigen = 0
	m_iGakunen = 0
	m_iSyoriNen = 0
	
    m_sClassNo = 0
    m_sClassName = ""
    
    m_sKamokuCd = ""
    m_iKamokuKbn = 0
    m_sKamokuName = ""
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
	m_sDate = gf_YYYY_MM_DD(trim(Request("txtDate")),"/")
	m_iJigen = trim(Request("sltJigen"))
	m_iSyoriNen = Session("NENDO")
	m_iGakunen = trim(Request("hidGakunen"))
	
	m_sClassNo = cint(Request("hidClassNo"))
	m_sClassName = gf_GetClassName(m_iSyoriNen,m_iGakunen,m_sClassNo)
	
	m_sKamokuCd = Request("hidKamokuCd")
	m_iKamokuKbn = cint(Request("hidSyubetu"))
	m_sKamokuName = gf_GetKamokuMei(m_iSyoriNen,m_sKamokuCd,m_iKamokuKbn)
	
End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage_middle()
	dim w_str	'表示メッセージ

    On Error Resume Next
    Err.Clear
	
	w_str = "<span class='CAUTION'>※ 入力したい「入力区分」を選択後、該当する学生の出欠状況覧をクリックして下さい。<BR></span>" & vbCrLf
	
%>
    <html>
    <head>
    <title>授業出欠入力</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
		//スクロール同期制御
		parent.init();
	}
	
    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Insert(){
		parent.frames["main"].f_Insert();
		return;
    }
	
    //************************************************************
    //  [機能]  戻るボタンが押されたとき
    //  [引数]  
    //  [戻値]  
    //  [説明]
    //************************************************************
    function f_Back(){
        //空白ページを表示
        parent.document.location.href="default.asp"
    }
	
    //-->
    </SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    <%call gs_title("授業出欠入力","登　録")%>
    	<table height="160">
    		<tr>
				<td>
	        		<table class="hyo" border="1" width="550">
	            		<tr>
							<th nowrap class="header" width="65" align="center">日付</th>
			                <td nowrap class="detail" width="120" align="center"><%=m_sDate%></td>
			                <th nowrap class="header" width="70" align="center">時限</th>
			                <td nowrap class="detail" width="30" align="center"><%=m_iJigen%></td>
			                <th nowrap class="header" width="65" align="center">科目</th>
			                <td nowrap class="detail" width="150" align="center"><%=m_sKamokuName%></td>
						</tr>

						
						<tr>
							<th nowrap class="header" width="65"  align="center">クラス</th>
			                <td nowrap class="detail" width="120"  align="center"><%=m_iGakunen & "年 " & m_sClassName & "科 " %></td>
			                <th nowrap class="header" width="70"  align="center">入力区分</th>
			                
			                <td nowrap class="detail" width=""  align="center" colspan="3">
			                	<input type="radio" name="rdoType" value="1" checked>
			                	<input type="text" name="txtKekka" size="2" maxlength="2" value="1">欠課
			                	
			                	<input type="radio" name="rdoType" value="2" >遅刻
			                	<input type="radio" name="rdoType" value="3" >早退
			                	<input type="radio" name="rdoType" value="4" >クリア
			                </td>
			            </tr>
	        		</table>
				</td>
			</tr>
			
			<tr>
				<td align="center">
					<table>
						<tr>
							<td><input type="button" name="btnInsert" value="　登　録　" onClick="f_Insert();"></td>
							<td><input type="button" name="btnBack" value="　戻　る　" onClick="f_Back();"></td>
						</tr>
	      			</table>
				</td>
			</tr>
			
			
			<tr>
				<td align="center">
					<table>
						<tr>
							<td><%=w_str%></td>
						</tr>
	      			</table>
				</td>
			</tr>
			
			<tr>
				<td align="center" valign="bottom">
					<table class="hyo" border="1" width="300">
						<tr>
							<th nowrap class="header" width="80"  align="center">学籍番号</th>
							<th nowrap class="header" width="150"  align="center">氏　名</th>
							<th nowrap class="header" width="70"  align="center">状況</th>
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
End Sub
%>
