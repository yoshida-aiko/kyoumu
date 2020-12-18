<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 連絡掲示板
' ﾌﾟﾛｸﾞﾗﾑID : web/web0330/web0330_main.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/10 前田
' 変      更: 2001/08/27 伊藤公子 コメントをリストの上部に表示するように変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_iMax          ':最大ページ
    Public  m_iDsp          '// 一覧表示行数
    Public  m_sPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_rs
    Dim     m_sNendo
    Dim     m_sKyokanCd
    Dim     m_rCnt          '//レコード件数

    'エラー系
    Public  m_bErrFlg       'ｴﾗｰﾌﾗｸﾞ
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

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="連絡掲示板"
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
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '//リストの一覧データの詳細取得
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

        '// ページを表示
        Call showPage()
        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_rs)
    '// 終了処理
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_sNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp = C_PAGE_LINE

    If Request("txtPageCD") <> "" Then
        m_sPageCD = INT(Request("txtPageCD"))   ':表示済表示頁数（自分自身から受け取る引数）
    Else
        m_sPageCD = 1   ':表示済表示頁数（自分自身から受け取る引数）
    End If
    If m_sPageCD = 0 Then m_sPageCD = 1

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_sNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp = C_PAGE_LINE

    If Request("txtPageCD") <> "" Then
        m_sPageCD = INT(Request("txtPageCD"))   ':表示済表示頁数（自分自身から受け取る引数）
    Else
        m_sPageCD = 1   ':表示済表示頁数（自分自身から受け取る引数）
    End If
    If m_sPageCD = 0 Then m_sPageCD = 1

End Sub

Function f_GetData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
        '//リストの表示
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT DISTINCT"
        m_sSQL = m_sSQL & "     T46_NO,T46_KENMEI,T46_KAISI,T46_SYURYO "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T46_RENRAK "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T46_INS_USER = '" & Session("LOGIN_ID") & "' "

		'応急処置。期限が過ぎたら表示しない。	2001/12/17
		'本当は表示しないだけじゃなくて期限が過ぎているデータは削除する。
        m_sSQL = m_sSQL & " AND T46_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "'"
        m_sSQL = m_sSQL & " AND T46_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
    m_rCnt=gf_GetRsCount(m_rs)

    f_GetData = 0

    Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub S_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    Dim w_pageBar           'ページBAR表示用

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True
%>
<div align="center"><span class=CAUTION>※ 新規登録の場合は｢新規登録はこちら｣をクリックしてください。<br>
										※ 修正の場合は｢>>｣をクリックしてください。<br>
										※ 件名をクリックすると送信内容を確認できます。
</span></div>
<br>
<%
    'ページBAR表示
    Call gs_pageBar(m_rs,m_sPageCD,m_iDsp,w_pageBar)

%>
<%=w_pageBar %>
</td></tr>
<tr><td>
<table width="100%" border="1" CLASS="hyo">
    <TR>
<!--         <TH CLASS="header" width="40" nowrap>処理<br>番号</TH> -->
        <TH CLASS="header" width="55%" nowrap>件　名</TH>
        <TH CLASS="header" nowrap>期　間</TH>
        <TH CLASS="header" width="16" nowrap>修正</TH>
        <!--TH CLASS="header" width="16" nowrap>削除</TH-->
    </TR>

<%	Do While (w_bFlg)
    call gs_cellPtn(w_cell)
	call gs_ColorPtnNN(w_color)
%>
    <TR>
<!--        <TD CLASS="<%=w_cell%>" ALIGN="right"><%=m_rs("T46_NO")%></TD> -->
        <TD CLASS="<%=w_cell%>"><a href="javascript:f_Kakunin(<%=m_rs("T46_NO")%>)"><%=m_rs("T46_KENMEI")%></a></TD>
        <TD CLASS="<%=w_cell%>" ALIGN="center"><%=m_rs("T46_KAISI")%>～<%=m_rs("T46_SYURYO")%></TD>
        <TD CLASS="<%=w_cell%>" ALIGN="center"><input type="button" value=">>" class=button onclick="javascript:f_Syusei(<%=m_rs("T46_NO")%>)"></TD>
        <!--TD CLASS="<%=w_cell%>" ALIGN="center"><input type="checkbox" name=Delchk value="<%=m_rs("T46_NO")%>"></TD-->
    </TR>
<% m_rs.MoveNext

		If m_rs.EOF Then
		    w_bFlg = False
		ElseIf w_iCnt >= C_PAGE_LINE Then
		    w_bFlg = False
		Else
		    w_iCnt = w_iCnt + 1
		End If

    Loop %>
    <tr>
    <!--td colspan=5 align="right" bgcolor=#9999BD><input class=button type=button value="×削除" onclick="javascript:f_delete()"></td-->
    </tr>
 </table>
 </td></tr>
 <tr><td>
<%=w_pageBar %>
<BR>
<!--
<div align="center"><span class=CAUTION>※ 新規登録の場合は｢新規登録はこちら｣をクリックしてください。<br>
										※ 修正の場合は｢>>｣をクリックしてください。<br>
										※ 件名をクリックすると送信内容を確認できます。
</span></div>
-->
<div align="center"><span class=CAUTION>※ メッセージは、表示期間を過ぎると自動的に削除されます。<br>
</span></div>
<%End sub

Sub showPage_NoData()
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
        <span class="msg">連絡データは存在しません。<br>｢新規登録はこちら｣をクリックしてください。</span>
    </center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
    Dim i
    i=1

%>
<HTML>
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>連絡掲示板</title>

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
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }
    //************************************************************
    //  [機能]  削除ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_delete(){

        if (f_chk()==1){
        alert( "削除の対象となる件名が選択されていません" );
        return;
        }

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "web0330_DEL.asp";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  リスト一覧のチェックボックスの確認
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_chk(){
        var i;
        i = 0;

        //0件のとき
        if (document.frm.txtRcnt.value<=0){
            return 1;
            }

        //1件のとき
        if (document.frm.txtRcnt.value==1){
            if (document.frm.Delchk.checked == false){
                return 1;
            }else{
                return 0;
                }
        }else{
        //それ以外の時
        var checkFlg
            checkFlg=false

        do { 
            
            if(document.frm.Delchk[i].checked == true){
                checkFlg=true
                break;
             }

        i++; }  while(i<document.frm.txtRcnt.value);
            if (checkFlg == false){
                return 1;
                }
        }
        return 0;
    }

    //************************************************************
    //  [機能]  件名ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Kakunin(p_Int){

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "view.asp";
        document.frm.txtNo.value = p_Int;
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  修正ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Syusei(p_Int){

        //リスト情報をsubmit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.txtNo.value = p_Int;
        document.frm.txtMode.value = "UPD";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" ACTION="post">
<table width="90%" border="0"><tr><td>
<%
    If m_rs.EOF Then
        Call showPage_NoData()
    Else
        Call S_syousai()
    End If
%>
    <INPUT TYPE=HIDDEN  NAME=txtNo          value="">
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_sNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtPageCD      value="<%= m_sPageCD %>">
    <INPUT TYPE=HIDDEN  NAME=txtRcnt        value="<%=m_rCnt%>">
</td></tr></table>

</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>