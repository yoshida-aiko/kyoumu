<%@Language=VBScript %>
<%
'******************************************************************
'システム名     ：教務事務システム
'処　理　名     ：各種委員登録
'プログラムID   ：gak/gak0470/default.asp
'機　　　能     ：フレームページ 学籍委員情報入力の表示を行う
'------------------------------------------------------------------
'引　　　数     ：
'変　　　数     ：
'引　　　渡     ：
'説　　　明     ：
'------------------------------------------------------------------
'作　　　成     ：2001.07.02    前田　智史
'変      更     : 2001/08/09 根本 直美     NN対応に伴うソース変更
'
'******************************************************************
'*******************　ASP共通モジュール宣言　**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******　モ ジ ュ ー ル 変 数　********
    'ページ関係
Public  m_iMax          ':最大ページ
Public  m_iDsp                      '// 一覧表示行数
Public  m_bErrFlg       '//エラーフラグ（DB接続エラー等の場合にエラーページを表示するためのフラグ）
Public  m_sDebugStr     '//以下デバック用
Public  m_iNendo
Public  m_sKyokanCd
Public  m_rs            '//レコードセット
Public  m_Irs           '//レコードセット（委員用）
Public  m_Grs           '//レコードセット（学籍番号用）
Public  m_sDaiNm()
Public  m_iDai()
Public  m_iSyo()
Public  m_iIinNm()

Public  m_IrCnt           '//レコードカウント
Public  m_GrCnt           '//レコードカウント
Public  m_iGAKKIKBN '学期区分

'******　メイン処理　********

    'ﾒｲﾝﾙｰﾁﾝ実行
    Call Main()

'******　Ｅ　Ｎ　Ｄ　********

Sub Main()
'******************************************************************
'機　　能：本ASPのﾒｲﾝﾙｰﾁﾝ
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

    '******共通関数******
    Dim w_iRet              '// 戻り値
    Dim w_sSQL              '// SQL文
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="各種委員登録"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")

   if Request("cboGakki") = "" then 'cboGakki

	m_iGAKKIKBN = session("GAKKI")
   
   else

	m_iGAKKIKBN = Request("cboGakki")

   End if

Response.Write m_iGAKKIKBN

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

        '// 権限チェックに使用
        session("PRJ_No") = "GAK0470"

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '// 担任チェック
	  If gf_Tannin(m_iNendo,m_sKyokanCd,1) <> 0 Then
	            m_bErrFlg = True
	            m_sErrMsg = "担任以外の入力はできません。"
	            Exit Do
	  End If

        w_iRet = f_getData()
        If w_iRet <> 0 Then
            'エラー処理
            m_bErrFlg = True
            Exit Do
        End If
		If m_IrCnt = 0 then
	        '// ページを表示
	        Call showPage_NO()
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

    '// 終了処理
    Call gs_CloseDatabase()
End Sub

Function f_getData()
'******************************************************************
'機　　能：データの取得
'返　　値：なし
'引　　数：なし
'機能詳細：
'備　　考：特になし
'******************************************************************

Dim i
i = 1

    On Error Resume Next
    Err.Clear
    f_getData = 1

    Do
        '//学年･クラスのデータ
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT"
        m_sSQL = m_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M05_CLASS "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     M05_NENDO = " & Cint(m_iNendo) & " "
        m_sSQL = m_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_rs, m_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If

        '//リスト(委員種別，委員名称)のデータ
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT"
        m_sSQL = m_sSQL & "     M34_DAIBUN_CD,M34_SYOBUN_CD,M34_IIN_NAME "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M34_IIN "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     M34_NENDO = " & Cint(m_iNendo) & "  "
        m_sSQL = m_sSQL & " AND M34_IIN_KBN <> " & C_IIN_GAKKO & " "
        m_sSQL = m_sSQL & " UNION"
        m_sSQL = m_sSQL & " SELECT"
        m_sSQL = m_sSQL & "     M34_DAIBUN_CD,M34_SYOBUN_CD,M34_IIN_NAME "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M34_IIN "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     M34_NENDO = " & Cint(m_iNendo) & "  "
        m_sSQL = m_sSQL & " AND M34_SYOBUN_CD = " & C_M34_SYOBUN_CD & " "

        Set m_Irs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Irs, m_sSQL)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
    m_IrCnt=gf_GetRsCount(m_Irs)
       '//リスト(氏名)のデータ
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT "
        m_sSQL = m_sSQL & "     A.T06_GAKUSEI_NO,A.T06_DAIBUN_CD,A.T06_SYOBUN_CD,B.T11_SIMEI "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T06_GAKU_IIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     A.T06_NENDO = " & Cint(m_iNendo) & " "
        m_sSQL = m_sSQL & " AND C.T13_GAKUNEN  = " & Cint(m_rs("M05_GAKUNEN")) & " "
        m_sSQL = m_sSQL & " AND C.T13_CLASS = " & Cint(m_rs("M05_CLASSNO")) & " "
        m_sSQL = m_sSQL & " AND A.T06_NENDO = C.T13_NENDO "
        m_sSQL = m_sSQL & " AND A.T06_GAKUSEI_NO = B.T11_GAKUSEI_NO "
        m_sSQL = m_sSQL & " AND B.T11_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		m_sSQL = m_sSQL & " AND A.T06_GAKKI_KBN = " & m_iGAKKIKBN

        Set m_Grs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Grs, m_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
    m_GrCnt=gf_GetRsCount(m_Grs)

        m_Irs.Movefirst

        Do Until m_Irs.EOF
            If Cint(m_Irs("M34_SYOBUN_CD")) = 0 Then
                ReDim Preserve m_sDaiNm(m_Irs("M34_DAIBUN_CD"))
                m_sDaiNm(m_Irs("M34_DAIBUN_CD")) = m_Irs("M34_IIN_NAME")
            Else
                ReDim Preserve m_iDai(i)
                ReDim Preserve m_iSyo(i)
                ReDim Preserve m_iIinNm(i)
                m_iDai(i) = m_Irs("M34_DAIBUN_CD")
                m_iSyo(i) = m_Irs("M34_SYOBUN_CD")
                m_iIinNm(i) = m_Irs("M34_IIN_NAME")
                i = i + 1
            End If
            
            m_Irs.MoveNext
            
        Loop

    f_getData = 0

    Exit Do

    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
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

    Dim i
    i = 0 

    '---------- HTML START ----------
    %>
    <html>
    <head>
    <title>各種委員登録</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <script language="javascript">
    <!--
    //************************************************************
    //  [機能]  申請内容表示用ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function NewWin(p_Int,p_Str,p_GInt,p_CInt,p_sNm) {

		var obj=eval("document.frm."+p_sNm)
        URL = "select.asp?i="+p_Int+ "&IINNM="+escape(p_Str)+"&GAKUNEN="+p_GInt+"&CLASS="+p_CInt+"&GName="+escape(obj.value)+"";
        //URL = "select.asp?i="+p_Int+ "&IINNM="+escape(p_Str)+"&GAKUNEN="+p_GInt+"&CLASS="+p_CInt+"&GName="+escape(p_sNm)+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,toolbar=no,width=700,height=600,top=0,left=0");
        return true;    
    }

    //************************************************************
    //  [機能] クリアボタンを押されたとき
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function jf_Clear(p_Name,p_Cd){
        eval("document.frm."+p_Name).value = "";
        eval("document.frm."+p_Cd).value = "";
    }

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Touroku(){

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        //リスト情報をsubmit
        document.frm.action="gak0470_edt.asp";
        document.frm.submit();

    }
	//************************************************************
    //  [機能]  学期が変更されたとき、本画面を再表示
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./default.asp";
        //document.frm.target="fTopMain";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //-->
    </script>

    </head>
    <body>
    <center>
    <form action="" name="frm" method="post">
<table border="0" width="90%">
<tr>
<td align="center">
<%call gs_title("各種委員登録","一　覧")%>

    <table border="0" cellpadding="1" cellspacing="1" width="100%">
    <tr>
    <td align="center">
        <table border="0" width="500" class=hyo align="center">
        <tr>
        <th width="65" class="header">学年</th>
        <td width="50" align="center" class="detail"><%=m_rs("M05_GAKUNEN")%>年</td>
        <th width="65" class="header">クラス</th>
        <td width="220" align="center" class="detail"><%=m_rs("M05_CLASSMEI")%></td>
		<th width="50" class="header">学期</th>
		<td width="50" class="detail">
		<select name="cboGakki" onchange = 'javascript:f_ReLoadMyPage()' ><Option Value="1"　Selected>前期
		<Option Value="2" Selected>後期</select></td>
        </tr>
        </table></td>
    </tr>
    <tr>
    <td align="center">
 <!--    <img src="../../image/sp.gif" height="30"> -->
<span class="msg">各委員の>>ボタンを押すと、学生選択画面が出てきます。<BR>学生を選択し、登録ボタンを押して下さい。</span>
    </td>
    </tr>
    <tr>
    <td align="center">

        <table width="100%" border="1" class="hyo">
        <tr>
        <th width="25%" class="header">委員種別</th>
        <th width="25%" class="header">委員名称</th>
        <th width="40%" class="header">氏　名</th>
        <th width="5%" class="header">選択</th>
        <th width="5%" class="header">　</th>
        </tr>

        <tr>

        <%
        For i = 1 to UBound(m_iDai)
            call gs_cellPtn(w_cell)%>

            <td  class="<%=w_cell%>" align="center"><font color="#000000"><%= m_sDaiNm(m_iDai(i)) %></font></td>
            <td  class="<%=w_cell%>" align="center"><font color="#000000"><%= m_iIinNm(i) %><br></font></td>

                <%
                If m_Grs.EOF = False Then
                    w_Name = ""
                    w_Gakusei_No = ""
                    m_Grs.MoveFirst
                    Do Until m_Grs.EOF
                        If Cint(m_iDai(i)) = Cint(m_Grs("T06_DAIBUN_CD")) and Cint(m_iSyo(i)) = Cint(m_Grs("T06_SYOBUN_CD")) Then
                            w_Name = m_Grs("T11_SIMEI")
                            w_Gakusei_No = m_Grs("T06_GAKUSEI_NO")
                            Exit Do
                        End If
                        m_Grs.MoveNext
                    Loop
                    m_Grs.MoveFirst
                End If %>
            <td  class="<%=w_cell%>" align="center">
                <input type="text" class="<%=w_cell%>" name="gakuNm<%=i%>" value="<%= w_Name %>" readonly><br>
                <input type="hidden" name="gakuNo<%=i%>" value="<%= w_Gakusei_No %>">
                <input type="hidden" name="iinDai<%=i%>" value="<%= Cint(m_iDai(i)) %>">
                <input type="hidden" name="iinSyo<%=i%>" value="<%= Cint(m_iSyo(i)) %>">
                <input type="hidden" name="Before<%=i%>" value="<%= w_Gakusei_No %>"></td>
            <!--<td  class="<%=w_cell%>" align="center"><input type="button" class="button" value=">>" onclick="NewWin(<%=i%>,'<%= m_iIinNm(i) %>',<%=m_rs("M05_GAKUNEN") %>,<%=m_rs("M05_CLASSNO") %>,'<%= w_Name %>')"></td>-->
            <td  class="<%=w_cell%>" align="center"><input type="button" class="button" value=">>" onclick="NewWin(<%=i%>,'<%= m_iIinNm(i) %>',<%=m_rs("M05_GAKUNEN") %>,<%=m_rs("M05_CLASSNO") %>,'gakuNm<%=i%>')"></td>
            <td  class="<%=w_cell%>"><input type="button" class="button" value="クリア" onclick="jf_Clear('gakuNm<%=i%>','gakuNo<%=i%>')"></td>
        </tr>
        <%Next%>
        </table>

    </td>
    </tr>
    </table><br><br>
        <input type="button" value="登　録" class="button" onclick="javascript:f_Touroku()">

    <INPUT TYPE="HIDDEN" NAME="HIDMAX" VALUE="<%= i-1 %>">
    <INPUT TYPE="HIDDEN" NAME="CLASS" VALUE="<%= m_rs("M05_CLASSNO") %>">
    <INPUT TYPE="HIDDEN" NAME="GAKUNEN" VALUE="<%= m_rs("M05_GAKUNEN")%>">
	<INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
	<INPUT TYPE="HIDDEN" NAME="GAKKI"	  value = "<% m_iGAKKI %>">

</td>
</tr>
</table>
    </form>
    </center>
    </body>
    </html>

<%
    '---------- HTML END   ----------
End Sub

Sub showPage_NO()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
    <html>
    <head>
    <title>各種委員登録</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <script language="javascript">
    </script>
    </head>
    <body>
    <center>
    <form action="" name="frm" method="post">
<table border="0" width="90%">
<tr>
<td align="center">
<%call gs_title("各種委員登録","一　覧")%>
<br><br><br><br><br>
        <span class="msg">学籍委員情報のデータがありません。</span>


    </center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub
%>