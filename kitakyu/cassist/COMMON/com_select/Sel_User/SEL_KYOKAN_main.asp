<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 教官参照選択画面
' ﾌﾟﾛｸﾞﾗﾑID : web/web0330/sousin_main.asp
' 機      能: 下ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:
' 説      明:
'-------------------------------------------------------------------------
' 作      成: 2001/07/10 前田
' 変      更: 2001/08/08 根本 直美     NN対応に伴うソース変更
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public  m_iMax          ':最大ページ
    Public  m_iDsp          '// 一覧表示行数
    Public  m_iNendo        '年度
    Public  m_sKyokanCd     '教官ｺｰﾄﾞ
    Public  m_sJoukin       '常勤区分
    Public  m_sGakka        '学科区分
    Public  m_sKkanKBN      '教官区分
    Public  m_sKkeiKBN      '教科系列区分
    Public  m_rs
    Public  m_sPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_iI
    Public  m_sKNm

	Public m_sUserKbn		'//USER区分
	Public m_sSimei			'//氏名

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bErrMsg           'ｴﾗｰﾒｯｾｰｼﾞ
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
    w_sMsgTitle="利用者選択画面"
    w_sMsg=""
    w_sRetURL="../../../../default.asp"
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

        '// ﾊﾟﾗﾒｰﾀSET
        Call s_SetParam()

        '//データの表示
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            Exit Do
        End If

        If m_Rs.EOF Then
            '// ページを表示
            Call showPage_NoData()
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

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_Rs)
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

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sJoukin   = request("Joukin")

    m_sGakka   = Trim(Replace(request("Gakka"),"@@@",""))
    m_sKkanKBN = Trim(Replace(request("KkanKBN"),"@@@",""))
    m_sKkeiKBN = Trim(Replace(request("KkeiKBN"),"@@@",""))
	m_sUserKbn = Trim(Replace(request("UserKbn"),"@@@",""))
	m_sSimei   = request("txtSimei")

    m_iI        = request("txtI")
    m_sKNm      = request("txtKNm")
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
        '//絞り込まれた条件で一覧の表示
        m_sSQL = ""
		m_sSQL = m_sSQL & vbCrLf & " SELECT "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_USER_ID "
		m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_KBN "
		m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_NAME "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_CD "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_GAKKA_CD "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAKEIRETU_KBN "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_KBN"
		m_sSQL = m_sSQL & vbCrLf & " FROM "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN "
		m_sSQL = m_sSQL & vbCrLf & " WHERE "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_KYOKAN_CD = M04_KYOKAN.M04_KYOKAN_CD(+) "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO = M04_KYOKAN.M04_NENDO(+)"
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_iNendo

        If m_sGakka <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_GAKKA_CD= '" & m_sGakka & "' "
        End If

        If m_sKkanKBN <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_KYOKAN_KBN=" & Cint(m_sKkanKBN)
        End If

        If m_sKkeiKBN <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_KYOKAKEIRETU_KBN=" & Cint(m_sKkeiKBN)
        End If

        If m_sUserKbn <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_USER_KBN= " & m_sUserKbn
        End If

        If m_sSimei <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_USER_NAME LIKE '%" & m_sSimei & "%'"
        End If

		m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_GAKKA_CD,M04_KYOKAKEIRETU_KBN,M10_USER_NAME"

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

'********************************************************************************
'*  [機能]  学科記号を取得
'*  [引数]  なし
'*  [戻値]  gf_GetUserNm:
'*  [説明]  
'********************************************************************************
Function f_GetGakkaKigoName(p_sGakkaCd)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetGakkaKigoName = ""
	w_sName = ""

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKA_KIGO"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_GAKKA_CD='" & p_sGakkaCd & "'"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M02_GAKKA_KIGO")
        End If

        Exit Do
    Loop

	'//戻り値ｾｯﾄ
    f_GetGakkaKigoName = w_sName

	'//ﾚｺｰﾄﾞｾｯﾄCLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

Sub S_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Dim w_iCnt
Dim w_pageBar           'ページBAR表示用

    w_iCnt  = 1
    w_bFlg  = True

    Call gs_pageBar(m_Rs,m_sPageCD,m_iDsp,w_pageBar)


%>
<%=w_pageBar %>
<table width="90%">
<tr><td>

<table border="1" width="100%" class="hyo">
<tr>
    <th width="30%" class="header">利用者区分</th>
    <th width="10%" class="header">学科</th>
    <th width="15%" class="header">教科系</th>
    <th width="45%" class="header">氏名</th>
</tr>
<%Do While (w_bFlg)
    Call gs_cellPtn(w_cell)
    %>
    <tr><form name="aaa" method="post">

		<%
		'========================================================
		'//区分名称等取得

		w_sKyokanKbnName = ""
		w_sKeiretuKbnName = ""
		w_sGakkaKigo = ""

		'//教官CDをセット
		w_sKyokanCd = m_rs("M04_KYOKAN_CD")

		'//教官の時(教官CDありの場合)
		If LenB(w_sKyokanCd) <> 0 Then
			'//教官区分名称を取得
			Call gf_GetKubunName(C_KYOKAN,m_rs("M04_KYOKAN_KBN"),m_iNendo,w_sKyokanKbnName)

			'//教科系列区分名称を取得
			Call gf_GetKubunName(C_KYOKA_KEIRETU,m_rs("M04_KYOKAKEIRETU_KBN"),m_iNendo,w_sKeiretuKbnName)

			w_sGakkaKigo = f_GetGakkaKigoName(m_rs("M04_GAKKA_CD"))
		Else

			'//教官以外の場合USER区分名称を表示
			Call gf_GetKubunName(C_USER,m_rs("M10_USER_KBN"),m_iNendo,w_sKyokanKbnName)
			w_sKeiretuKbnName = "―"
			w_sGakkaKigo = "―"

		End If
		'========================================================
		%>

        <td align="center" class="<%=w_cell%>"><%=w_sKyokanKbnName%><BR></td>
        <td align="center" class="<%=w_cell%>"><%=w_sGakkaKigo%><BR></td>
        <td align="center" class="<%=w_cell%>"><%=w_sKeiretuKbnName%><BR></td>
        <td align="center" class="<%=w_cell%>">
        <input type="button" class="<%=w_cell%>" name="KNm" value='<%=m_rs("M10_USER_NAME")%>' onclick="iinSelect(this.form)">
        <input type="hidden" name="KCd" value='<%=gf_HTMLTableSTR(m_Rs("M10_USER_ID")) %>'>

        </td>
    </form></tr>
<%
    m_rs.MoveNext

        If m_rs.EOF Then
            w_bFlg = False
        ElseIf w_iCnt >= C_PAGE_LINE Then
            w_bFlg = False
        Else
            w_iCnt = w_iCnt + 1
        End If

Loop %>

</table>
</td></tr></table>
<%=w_pageBar %>
<br>

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
<link rel="stylesheet" href="../../style.css" type="text/css">
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

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************


%>
<HTML>
<BODY>

<link rel="stylesheet" href="../../style.css" type="text/css">
    <title>利用者選択画面</title>

    <!--#include file="../../jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
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
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  申請内容表示用ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function iinSelect(p_sct) {

        //挿入元のフォームを取得
            w_sctNm = p_sct.KNm;
            w_sctNo = p_sct.KCd;

        //挿入処理
            parent.opener.document.frm.SKyokanNm<%=m_iI%>.value = w_sctNm.value;
            parent.opener.document.frm.SKyokanCd<%=m_iI%>.value = w_sctNo.value;

            document.frm.SKyokanNm.value = w_sctNm.value;
            document.frm.SKyokanCd.value = w_sctNo.value;

        return true;    
        //window.close()

    }

    //************************************************************
    //  [機能]  クリアボタンをクリックした場合
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function f_Clear() {

        document.frm.SKyokanNm.value = "";
        document.frm.SKyokanCd.value = "";

        //挿入させたいフォームを取得
            w_NmStr = parent.opener.document.frm.SKyokanNm<%=m_iI%>;
            w_NoStr = parent.opener.document.frm.SKyokanCd<%=m_iI%>;

        //挿入処理

            w_NmStr.value = document.frm.SKyokanNm.value;
            w_NoStr.value = document.frm.SKyokanCd.value;
        return true;    
    }
    
    //-->
    </SCRIPT>

	<center>

	<FORM NAME="frm" method="post">
	    <INPUT TYPE="HIDDEN" NAME="txtNendo"    value="<%=m_iNendo%>">
	    <INPUT TYPE="HIDDEN" NAME="txtKyokanCd" value="<%=m_sKyokanCd%>">
	    <INPUT TYPE="HIDDEN" NAME="txtPageCD"   value="<%=m_sPageCD%>">
	    <input type="hidden" name="txtI"        value="<%=m_iI%>">
	    <input type="hidden" name="txtKNm"      value="<%=m_sKNm%>">
	<table width="50%" class="hyo">
	    <tr>
	        <td align="center" width="40%"><font color="white">利用者名</font></td>
	        <td align="center" class="detail"><input type="text" class="noBorder" name="SKyokanNm" value="<%=m_sKNm%>" readonly>
	        <input type="hidden" name="SKyokanCd" value=""></td>
	    </tr>
	</table>

<!--	<span class="CAUTION">※ 選択をするには教官名をクリックしてください。</span>-->
	<span class="CAUTION">※ 氏名をクリックして、利用者を選択してください。</span>
	        <input type="button" value=" クリア " class="button" onclick="javascript:f_Clear();">
	        <input type="button" value="閉じる" class="button" onclick="javascript:parent.window.close();">

	<%
	        Call S_syousai()
	%>
	<table>
	    <tr>
	        <td colspan="4" align="center" nowrap>
	        <form>
	        <input type="button" value=" クリア " class="button" onclick="javascript:f_Clear();">
	        <input type="button" value="閉じる" class="button" onclick="javascript:parent.window.close();">
	        </form>
	        </td>
	    </tr>
	</table>
	</FORM>
	</center>
	</BODY>
	</HTML>
<%
End Sub
%>
