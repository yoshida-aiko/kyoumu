<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 就職先マスタ
' ﾌﾟﾛｸﾞﾗﾑID : mst/mst0133/main.asp
' 機      能: 下ページ 就職先マスタの一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           txtSinroKBN     :進路先コード
'           txtSingakuCd        :進学コード
'           txtSyusyokuName     :就職先名称（一部）
'           txtPageCD       :表示済表示頁数（自分自身から受け取る引数）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           txtRenrakusakiCD    :選択された連絡先コード
'           txtPageCD       :表示済表示頁数（自分自身に引き渡す引数）
' 説      明:
'           ■初期表示
'               検索条件にかなう就職・進学先を表示
'           ■次へ、戻るボタンクリック時
'               指定した条件にかなう就職・進学を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/06/18 岩下　幸一郎
' 変      更: 2001/07/13 谷脇　良也
' 変      更: 2001/08/22 伊藤　公子 業種区分追加対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public  m_iNendo            '処理年度
    Public  m_sSinroCD      ':進路先コード
    Public  m_sSingakuCd        ':進学コード
    Public  m_sSyusyokuName     ':就職先名称（一部）
    Public  m_sPageCD       ':表示済表示頁数（自分自身から受け取る引数）
    Public  m_skubun
    Public  m_Rs            'recordset
    Public  m_iDisp         ':表示件数の最大値をとる
    Public  m_sMode
    'ページ関係
    Public  m_iMax          ':最大ページ
    Public  m_iDsp          '// 一覧表示行数

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
    w_sMsgTitle="就職マスタ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    '// ﾊﾟﾗﾒｰﾀSET
    Call s_SetParam()

        If m_sMode = "" Then
        '// ページを表示
        Call NoPage()
    Else
        
        On Error Resume Next
        Err.Clear

        m_bErrFlg = False
        m_iDsp = C_PAGE_LINE

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

            '就職マスタを取得
            w_sWHERE = ""

            w_sSQL = w_sSQL & vbCrLf & " SELECT "
            w_sSQL = w_sSQL & vbCrLf & " M01.M01_SYOBUNRUIMEI "
            w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_NENDO "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_CD "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_KBN "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINGAKU_KBN "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_GYOSYU_KBN "
            w_sSQL = w_sSQL & vbCrLf & " FROM "
            w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
            w_sSQL = w_sSQL & vbCrLf & "    ,M01_KUBUN M01 "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "    M01_NENDO = " & m_iNendo & " AND "
            w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo & " AND "

'---2001/08/22 ito 業種区分追加対応
	         w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD = "&C_SINRO&""
	         w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN = M01.M01_SYOBUNRUI_CD "

            '抽出条件の作成
            If m_sSinroCD<>"" Then
                w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN =" & m_sSinroCD & " "
            End If

'---2001/08/22 ito 業種区分追加対応
	        If m_sSingakuCd <> "" Then
				if cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU then
		            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN =" & m_sSingakuCd & " "
				ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU then
		            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_GYOSYU_KBN =" & m_sSingakuCd & " "
				End if
	        End If

            If m_sSyusyokuName<>"" Then
                w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINROMEI Like '%" & m_sSyusyokuName & "%' "
            End If

            w_sSQL = w_sSQL & vbCrLf & " ORDER BY M32.M32_SINRO_CD "

'   Response.Write w_sSQL & "<br>"

            Set m_Rs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)

            If w_iRet <> 0 Then
                'ﾚｺｰﾄﾞｾｯﾄの取得失敗
                m_bErrFlg = True
                Exit Do 'GOTO LABEL_MAIN_END
            Else
                'ページ数の取得
                m_iMax = gf_PageCount(m_Rs,m_iDsp)
'   Response.Write "m_iMax:" & m_iMax & "<br>"
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
        
        '// 終了処理
        Call gf_closeObject(m_Rs)
        Call gs_CloseDatabase()
    End If

End Sub


'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()
    m_sSinroCD = Request("txtSinroCD")              ':進路コード
    If m_sSinroCD="@@@" Then m_sSinroCD=""      'コンボ未選択時

    m_sSingakuCd = Request("txtSingakuCd")          ':進学コード
    If m_sSingakuCd="@@@" Then m_sSingakuCd=""  'コンボ未選択時

    m_sMode = Request("txtMode")            ':モード

    m_iNendo = Session("NENDO")     ':年度
    m_sSyusyokuName = Request("txtSyusyokuName")    ':就職先名称（一部）

    If Request("txtPageCD") <> "" Then
        m_sPageCD = INT(Request("txtPageCD"))   ':表示済表示頁数（自分自身から受け取る引数）
    Else
        m_sPageCD = 1   ':表示済表示頁数（自分自身から受け取る引数）
    End If
    If m_sPageCD = 0 Then m_sPageCD = 1

    If m_sSinroCD = "1" Then            ':ヘッダーの区分名称変更
        m_skubun = "進学区分"
    else
        m_skubun = "進路区分"
    End If
    
    m_iDisp = C_PAGE_LINE       '１ページ最大表示数

End Sub


Sub S_syousai()
'********************************************************************************
'*  [機能]  詳細を表示
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

Dim w_slink
Dim w_iCnt
Dim w_i
Dim w_cell
w_iCnt  = 1
w_i     = 0
w_cell = ""

Do While not m_Rs.EOF
	w_i = w_i + 1

	w_slink = "　"

	if m_Rs("M32_SINRO_URL") <> "" Then 
	    w_sLink= "<a href='" & gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "' target='_blank'>" 
	    w_sLink= w_sLink &  gf_HTMLTableSTR(trim(m_Rs("M32_SINRO_URL"))) & "</a>"
	End if
	call gs_cellPtn(w_cell)
        %>

        <tr>
        <td align="center" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M01_SYOBUNRUIMEI")) %></td>

		<%
		'//初期化
		w_sKbn = ""
		w_sKbnName = ""

		'//進路区分OR業種区分名称を取得
		Select case cint(gf_SetNull2Zero(m_Rs("M32_SINRO_KBN")))
			Case C_SINRO_SINGAKU	'//進路区分が進学の場合

				'//進学区分名称を取得
				w_sKbn = trim(m_Rs("M32_SINGAKU_KBN"))
				If w_sKbn <> "" Then
					Call gf_GetKubunName(C_SINGAKU,m_Rs("M32_SINGAKU_KBN"),m_iNendo,w_sKbnName)
				End If

			Case C_SINRO_SYUSYOKU	'//進路区分が就職の場合

				'//業種区分名称を取得
				w_sKbn = trim(m_Rs("M32_GYOSYU_KBN"))
				If w_sKbn <> "" Then
					Call gf_GetKubunName(C_GYOSYU_KBN,m_Rs("M32_GYOSYU_KBN"),m_iNendo,w_sKbnName)
				End If

			Case C_SINRO_SONOTA	'//進路区分がその他の場合

		End Select

		%>

        <td align="center" class=<%=w_cell%>><%=gf_HTMLTableSTR(w_sKbnName) %></td>

        <td align="left" class=<%=w_cell%>><%=gf_HTMLTableSTR(trim(m_Rs("M32_SINROMEI"))) %></a></td>
        <td align="left" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M32_DENWABANGO")) %></td>
        <td align="left" class=<%=w_cell%>><%=w_slink%></td>
        <td align="center" class=<%=w_cell%>><input class=button type="button" value=">>" onclick="javascript:f_Henko('<%=cstr(m_Rs("M32_SINRO_CD")) %>')"></td>
        <td align="center" class=<%=w_cell%>><input type="checkbox" name="deleteNO<%= w_i %>" value="<%=gf_HTMLTableSTR(m_Rs("M32_SINRO_CD")) %>"></td>
        </tr>

        <%
            m_Rs.MoveNext
            If w_iCnt >= C_PAGE_LINE Then
                Exit Sub
            Else
                w_iCnt = w_iCnt + 1
            End If
        Loop

        m_iDisp = w_i

End sub



Sub NoPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

%>

    <html>

    <head>
	<link rel=stylesheet href=../../common/style.css type=text/css>

    </head>

    <body>

	<center>
	<br><br><br>
	<span class="msg"><%=C_BRANK_VIEW_MSG%></span>
	</center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub


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
    Dim w_pageBar           'ページBAR表示用
    
    On Error Resume Next
    Err.Clear

%>

<html>
    <head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
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

        document.frm.action="main.asp";
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }
   

    //************************************************************
    //  [機能]  修正画面を表示する
    //  [引数]  p_sSinroCD
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_GoSyosai(p_sSinroCD){

        document.frm.action="syousai.asp";
        document.frm.target="";
        document.frm.txtPageCD.value = p_sSinroCD;
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    
    }

    //************************************************************
    //  [機能]  syosai_frmへのパラメータの受け渡し
    //  [引数]  p_sSyuseiCD
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Henko(p_sSyuseiCD){

        document.frm.action="syusei.asp";
        document.frm.target="fTopMain";
        document.frm.txtRenrakusakiCD.value = p_sSyuseiCD;
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    }


    //************************************************************
    //  [機能]  削除ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Delete(){

    MainFrm = parent.window.frames["main"]
    var i;
    i = 1;

    var checkFlg
        checkFlg=false

    do { 
        obj = eval("document.frm.deleteNO" + i);
        if(obj.checked == true){
        
            checkFlg=true
            break;
         }

    i++; }  while(i<=document.frm.txtDisp.value);
    if (checkFlg == false){
        alert( "削除の対象となる進路が選択されていません" );
    }else{

        document.frm.action="./del_kakunin.asp";
        document.frm.target="fTopMain";
        document.frm.txtMode.value = "Delete";
        document.frm.submit();
        }
    }


    //-->
    </SCRIPT>

    </head>

<body>
<center>

<form name="frm" action="" target="" method="post">
<br>
<table><tr><td align="center" width="800">
	<%
		'ページBAR表示
		Call gs_pageBar(m_Rs,m_sPageCD,m_iDsp,w_pageBar)
	%>
	<%=w_pageBar %>

    <table border=1 class=hyo width="100%">
	    <tr>
		    <th class=header width="80">進路区分</th>
		    <th class=header width="80">種別区分</th>
		    <th class=header>進　路　名</th>
		    <th class=header width="96">Ｔ Ｅ Ｌ</th>
		    <th class=header width="30%">Ｕ Ｒ Ｌ</th>
		    <th class=header width="32">修正</th>
		    <th class=header width="32">削除</th>
	    </tr>
	    <% S_syousai() %>
	    <tr>
		    <td colspan=7 align=right bgcolor=#9999BD><input class=button type=button value="×削除" Onclick="f_Delete()"></td>
	    </tr>
	</table>

	<%=w_pageBar %>
</td></tr></tabel>

<br>
</center>

<input type="hidden" name="txtMode" value="">
<input type="hidden" name="txtRenrakusakiCD" value="">
<input type="hidden" name="txtSinroCD2" value="<%= m_sSinroCD %>">
<input type="hidden" name="txtSingakuCD2" value="<%= m_sSingakuCd %>">
<input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD %>">
<input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCd %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
<input type="hidden" name="txtNendo" value="<%= Session("SYORI_NENDO") %>">
<input type="hidden" name="txtDisp" value="<%= m_iDisp %>">
</form>

</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>