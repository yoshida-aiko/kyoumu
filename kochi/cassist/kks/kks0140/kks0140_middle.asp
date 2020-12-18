<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 行事出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0140/kks0140_middle.asp
' 機      能: 下ページ 行事出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数: NENDO     '//年度
'             KYOKAN_CD '//教官CD
'             GAKUNEN   '//学年
'             CLASSNO   '//クラスNO
'             GYOJI_CD  '//行事CD
'             GYOJI_MEI '//行事名
'             KAISI_BI  '//開始日
'             SYURYO_BI '//終了日
'             SOJIKANSU '//総時間数
' 変      数:
' 引      渡: NENDO     '//年度
'             KYOKAN_CD '//教官CD
'             GAKUNEN   '//学年
'             CLASSNO   '//クラスNO
'             GYOJI_CD  '//行事CD
'             GYOJI_MEI '//行事名
'             KAISI_BI  '//開始日
'             SYURYO_BI '//終了日
'             SOJIKANSU '//総時間数
' 説      明:
'           ■初期表示
'               検索条件にかなう行事出欠入力を表示
'           ■表示ボタンクリック時
'               指定した条件にかなう中学校を表示させる
'           ■登録ボタンクリック時
'               入力情報を登録する
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
    Const DebugPrint = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public m_iSyoriNen      '//処理年度
    Public m_iKyokanCd      '//教官CD
    Public m_sGakunen       '//学年
    Public m_sClassNo       '//ｸﾗｽNO
    Public m_sTuki          '//月
    Public m_sGyoji_Cd      '//行事CD
    Public m_sGyoji_Mei     '//行事名
    Public m_sKaisi_Bi      '//開始日
    Public m_sSyuryo_Bi     '//終了日
    Public m_sSoJikan       '//総時間数
	Public m_sEndDay		'//入力できなくなる日

    '//ﾚｺｰﾄﾞセット
    Public m_Rs_M           '//recordset明細情報
    Public m_Rs_G           '//recordset行事出欠情報
    Public m_iRsCnt         '//ヘッダﾚｺｰﾄﾞ数

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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="行事出欠入力"
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

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))


        '//変数初期化
        Call s_ClearParam()

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

'//デバッグ
'Call s_DebugPrint()

        '// 生徒リスト情報取得
        w_iRet = f_Get_DetailData()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

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
    Call gf_closeObject(m_Rs_M)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen   = ""
    m_iKyokanCd  = ""
    m_sGakunen = ""
    m_sClassNo = ""
    m_sTuki = ""

    m_sGyoji_Cd  = ""
    m_sGyoji_Mei = ""
    m_sKaisi_Bi  = ""
    m_sSyuryo_Bi = ""
    m_sSoJikan   = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = trim(Request("NENDO"))
    m_iKyokanCd = trim(Request("KYOKAN_CD"))
    m_sGakunen  = trim(Request("GAKUNEN"))
    m_sClassNo  = trim(Request("CLASSNO"))
    m_sTuki     = trim(Request("TUKI"))

    m_sGyoji_Cd  = trim(Request("GYOJI_CD"))
    m_sGyoji_Mei = trim(Request("GYOJI_MEI"))
    m_sKaisi_Bi  = trim(Request("KAISI_BI"))
    m_sSyuryo_Bi = trim(Request("SYURYO_BI"))
    m_sSoJikan   = trim(Request("SOJIKANSU"))
    m_sEndDay   = trim(Request("ENDDAY"))

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()

    response.write "<font color=#000000>m_iSyoriNen = " & m_iSyoriNen  & "</font><br>"
    response.write "<font color=#000000>m_iKyokanCd = " & m_iKyokanCd  & "</font><br>"
    response.write "<font color=#000000>m_sGakunen  = " & m_sGakunen   & "</font><br>"
    response.write "<font color=#000000>m_sClassNo  = " & m_sClassNo   & "</font><br>"
    response.write "<font color=#000000>m_sTuki     = " & m_sTuki      & "</font><br>"

    response.write "<font color=#000000>m_sGyoji_Cd = " & m_sGyoji_Cd  & "</font><br>"
    response.write "<font color=#000000>m_sGyoji_Mei= " & m_sGyoji_Mei & "</font><br>"
    response.write "<font color=#000000>m_sKaisi_Bi = " & m_sKaisi_Bi  & "</font><br>"
    response.write "<font color=#000000>m_sSyuryo_Bi= " & m_sSyuryo_Bi & "</font><br>"
    response.write "<font color=#000000>m_sSoJikan  = " & m_sSoJikan   & "</font><br>"

End Sub

'********************************************************************************
'*  [機能]  クラス一覧を取得
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_DetailData()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_Get_DetailData = 1

    Do 

        '// 明細データ
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN," 
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "   T11_GAKUSEKI.T11_SIMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN,T11_GAKUSEKI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN=" & cInt(m_sGakunen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS=" & cInt(m_sClassNo)
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write "<font color=#000000>" & w_sSQL & "</font><BR>"
        iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_DetailData = 99
            Exit Do
        End If

        '//件数を取得
        m_iRsCnt = 0
        If m_Rs_M.EOF = False Then
            m_iRsCnt = gf_GetRsCount(m_Rs_M)
        End If

        f_Get_DetailData = 0
        Exit Do
    Loop

End Function

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
    <title>行事用出欠入力</title>
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

		//スクロール同期制御
		parent.init();

        <%If m_Rs_M.EOF = True Then%>
			document.location.href="white.htm"
			return;
		<%End If%>
    }

    //************************************************************
    //  [機能]  キャンセルボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Cancel(){
        //空白ページを表示
        parent.document.location.href="default.asp"
    }

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){

		parent.frames["main"].f_Touroku();
		return;
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>

    <form name="frm" method="post" >
    <%call gs_title("行事出欠入力","一　覧")%>
	<br>
    <%Do%>
        <%If m_Rs_M.EOF = True Then
			Exit Do
		End If%>

        <table>
		<tr><td>
            <table class=hyo width="590" border="1" >
                <tr>
                    <th nowrap class="header" width="80"  align="center">行事名</th>
                    <td nowrap class="detail" width="200" align="left">　<%=m_sGyoji_Mei%></td>
                    <th nowrap class="header" width="80"  align="center">時限数</th>
                    <td nowrap class="detail" width="50"  align="center"><%=m_sSoJikan%></td>
                    <th nowrap class="header" width="80"  align="center">実施日</th>
                    <td nowrap class="detail" width="100" align="center"><%=month(m_sKaisi_Bi) & "/" & day(m_sKaisi_Bi)%>～<%=month(m_sSyuryo_Bi) & "/" & day(m_sSyuryo_Bi)%></td>
                </tr>
            </table>
		</td></tr><tr>
		<td align="center">
	<% 'If m_sEndDay < m_sSyuryo_Bi then %>
            <table>
                <td ><input class=button type="button" onclick="javascript:f_Touroku();" value="　登　録　"></td>
                <td ><input class=button type="button" onclick="javascript:f_Cancel();" value="キャンセル"></td>
            </table>
	<% 'Else %>
            <!--table>
                <td ><input class=button type="button" onclick="javascript:f_Cancel();" value=" 戻　る "></td>
            </table-->
	<% 'End If %>
		</td></tr>
        </table>

        <!--明細ヘッダ部-->

        <table >
            <tr>
				<td valign="top">
		            <table class="hyo"  border="1" >
		               <tr>
		                   <th class="header" width="80"  height="23" align="center"  nowrap><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
		                   <th class="header" width="150" height="23" align="center"  nowrap>氏　名</th>
		                   <th class="header" width="80"  height="23" align="center"  nowrap>欠課時間</th>
		               </tr>
		            </table>
	            </td>
			<%If m_iRsCnt <> 1 Then%>
				<td width="10"><br></td>
				<td valign="top" >
	                <table class="hyo"  border="1" >
                        <tr>
                            <th class="header" width="80"  height="23" align="center"  nowrap><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
                            <th class="header" width="150" height="23" align="center"  nowrap>氏　名</th>
                            <th class="header" width="80"  height="23" align="center"  nowrap>欠課時間</th>
		               </tr>
                  </table>
                </td>
			<%End If%>
			</tr>
        </table>

        <%
        Exit Do

    Loop%>

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>