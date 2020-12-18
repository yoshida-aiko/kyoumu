<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 授業出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0111/kks0111_main.asp
' 機      能: 下ページ 授業出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             TUKI           '//月
' 変      数:
' 引      渡: NENDO          '//処理年
'             KYOKAN_CD      '//教官CD
'             GAKUNEN        '//学年
'             CLASSNO        '//ｸﾗｽNo
'             TUKI           '//月
' 説      明:
'           ■初期表示
'               検索条件にかなう行事出欠入力を表示
'           ■登録ボタンクリック時
'               入力情報を登録する
'-------------------------------------------------------------------------
' 作      成: 
' 変      更: 2015.03.19 kiyomoto Win7対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '取得したデータを持つ変数
    Public m_iSyoriNen      '//処理年度
    
	Public m_JigenCount		'//時限数
	
	Public m_sGakunenCd		'//学年
	Public m_sClassCd		'//クラスCD
	Public m_sFromDate		'//kks0111_top.aspで入力した期間の始まり
	Public m_sToDate		'//kks0111_top.aspで入力した期間の終わり
	Public m_sClassName		'//クラス名
	
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
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		'//変数初期化
		Call s_ClearParam()
		
		'//パラメータSET
        Call s_SetParam()
		
		'//ページ表示
		Call showPage()
		
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
	
	m_JigenCount = 0
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	m_sFromDate = ""
	m_sToDate = ""
	
	m_iSyoriNen = ""
    
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen = Session("NENDO")
	
	m_JigenCount = request("JigenCount")
	
	m_sGakunenCd = request("cboGakunenCd")
	m_sClassCd = request("cboClassCd")
	m_sFromDate = gf_YYYY_MM_DD(request("txtFromDate"),"/")
	m_sToDate = gf_YYYY_MM_DD(request("txtToDate"),"/")
	
	m_sClassName = gf_GetClassName(m_iSyoriNen,m_sGakunenCd,m_sClassCd)
		
End Sub

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
	On Error Resume Next
    Err.Clear
	
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
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Back(){
        //空白ページを表示
        parent.document.location.href="default.asp"
    }
	//-->
    </SCRIPT>
	</head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    
    <%call gs_title("出欠状況","参照")%>
    	<table>
			<tr>
				<td nowrap>
			        <table class="hyo" border="1" width="300">
			            <tr>
							<th class="header" width="50"  align="center" nowrap>クラス</th>
							<td class="detail" width="150" align="left" nowrap>　<%=m_sGakunenCd%>年　<%=m_sClassName%>科　</td>
						</tr>
						
						<tr>
							<th class="header" width="50"  align="center" nowrap>日付</th>
							<td class="detail" align="left" nowrap>　<%=m_sFromDate%>　～　<%=m_sToDate%>　</td>
						</tr>
					</table>
				</td>
			</tr>
			
			<tr>
				<td align="center" nowrap>
					<table>
						<tr>
							<td valign="bottom"align="center" nowrap>
				        	    <input class="button" type="button" onclick="javascript:f_Back();" value=" 戻　る ">
				        	</td>
						</tr>
			        </table>
				</td>
			</tr>
        </table>
		
		<!-- 2015.03.19 Upd Start kiyomoto-->
        <!--<table width=800>-->
        <table>
		<!-- 2015.03.19 Upd End kiyomoto-->
        <tr>
            <td align="center" nowrap>
	            <table class="hyo"  border="1">
				     <tr>
		                <th class="header"  rowspan="2" width="100" align="center" nowrap><font color="#ffffff">
		                    <%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></font>
		                </th>
						<th class="header" width="100" align="center" rowspan="2" nowrap><font color="#ffffff">氏名</font></th>
						<th class="header" width="50" align="center" rowspan="2" nowrap><font color="#ffffff">詳細</font></th>
						
						<%Dim w_num%>
						<%for w_num = 1 to m_JigenCount%>
							<th class="header" width="50" align="center" colspan="2" nowrap><font color="#ffffff"><%=w_num%></font></th>
						<%next%>
					</tr>
					
					<tr>
						<%for w_num = 1 to m_JigenCount%>
							<th class="header" width="20" align="center" nowrap><font color="#ffffff">欠</font></th>
							<th class="header" width="20" align="center" nowrap><font color="#ffffff">遅</font></th>
						<%next%>
					</tr>
				</table>
			</td>
        	
        	<td width="10" nowrap><br></td>
        	
        	<td align="center" width="120" nowrap>
				
	            <table width="120" class="hyo" border="1">
		            <tr>
		                <th colspan="2" class="header" align="center" width="60" nowrap><font color="#ffffff">前期</font></th>
		                <th colspan="2" class="header" align="center" width="60" nowrap><font color="#ffffff">後期</font></th>
		            </tr>
		            <tr>
		                <th class="header" width="30" align="center" nowrap><font color="#ffffff">欠</font></th>
		                <th class="header" width="30" align="center" nowrap><font color="#ffffff">遅</font></th>
		                <th class="header" width="30" align="center" nowrap><font color="#ffffff">欠</font></th>
		                <th class="header" width="30" align="center" nowrap><font color="#ffffff">遅</font></th>
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
