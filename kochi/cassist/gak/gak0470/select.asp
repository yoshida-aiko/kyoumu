<%@Language=VBScript %>
<%
'******************************************************************
'システム名     ：教務事務システム
'処　理　名     ：各種委員登録
'プログラムID   ：gak/gak0470/select.asp
'機　　　能     ：クラス情報表示、選択を行う
'------------------------------------------------------------------
'引　　　数     ：
'変　　　数     ：
'引　　　渡     ：
'説　　　明     ：
'------------------------------------------------------------------
'作　　　成     ：2001.07.02    前田　智史
'変      更     : 2001/08/08 根本 直美     NN対応に伴うソース変更
'変      更     : 2002/04/23 宮井　　項目名「学籍番号」を管理マスタから取得するように変更
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
Dim     m_iNendo        '//処理年度
Dim     m_sKyokanCd     '//教官コード
Dim     m_sGakunen      '//学年
Dim     m_sClass        '//クラス名
Dim     m_sIinNm        '//委員名称
Dim     m_iI            '//defaultのリストの位置
Dim     m_rs            '//レコードセット
Dim     m_Irs           '//レコードセット（委員用）
Dim     m_Grs           '//レコードセット（学籍番号用）
Dim     m_rCnt          '//レコード件数
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
    m_sGakunen  = request("GAKUNEN")
    m_sClass    = request("CLASS")
    m_sIinNm    = request("IINNM")
    m_iI        = request("i")
    m_iDsp      = C_PAGE_LINE

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

        '//リスト一覧の表示
        w_iRet = f_getData()
        If w_iRet <> 0 Then
            'エラー処理
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

    On Error Resume Next
    Err.Clear
    f_getData = 1

    Do
        '//学年･クラスのデータ
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT"
        m_sSQL = m_sSQL & "     A.T13_GAKUSEI_NO,A.T13_GAKUSEKI_NO,B.T11_SIMEI "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T13_GAKU_NEN A,T11_GAKUSEKI B "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     A.T13_NENDO = '" & m_iNendo & "' "
        m_sSQL = m_sSQL & " AND A.T13_GAKUNEN = '" & m_sGakunen & "' "
        m_sSQL = m_sSQL & " AND A.T13_CLASS = '" & m_sClass & "' "
        m_sSQL = m_sSQL & " AND A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO(+) "
        m_sSQL = m_sSQL & " ORDER BY A.T13_GAKUSEKI_NO "

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do 
        End If
    m_rCnt=gf_GetRsCount(m_rs)
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

Dim w_Half
Dim w_sKNM
Dim j
j = 0

w_sKNM = request("GName")

    On Error Resume Next
    Err.Clear

	'// ﾌﾞﾗｳｻﾞｰによってﾎﾞﾀﾝのｻｲｽﾞを変える 
	w_btnWidth = ""
	if session("browser") = "NN" then
		w_btnWidth = "style='width:200'"
	End if

    '---------- HTML START ----------
    %>
<html>
<head>
<title>各種委員登録</title>
<link rel=stylesheet href="../../common/style.css" type="text/css">
<script language=javascript>
<!--
        //************************************************************
        //  [機能]  申請内容表示用ウィンドウオープン
        //  [引数]
        //  [戻値]
        //  [説明]
        //************************************************************
        function iinSelect(p_sct,p_No) {

            //挿入させたいフォームを取得
				w_NmStr = eval("opener.document.frm.gakuNm" + p_No);
				w_NoStr = eval("opener.document.frm.gakuNo" + p_No);

            //挿入元のフォームを取得
                w_sctNm = p_sct.gakuNm;
                w_sctNo = p_sct.gakuNo;

            //挿入処理
                w_NmStr.value = w_sctNm.value;
                w_NoStr.value = w_sctNo.value;

                document.frm.SearchNm.value = w_sctNm.value;
                document.frm.SearchNo.value = w_sctNo.value;

            return true;    
            //window.close()

        }

        //************************************************************
        //  [機能]  氏名,学生コードの申請内容表示
        //  [引数]
        //  [戻値]
        //  [説明]
        //************************************************************
        //function f_SearchSelect(p_sct) {
        //  //挿入元のフォームを取得
        //      w_sctNm = p_sct.gakuNm;
        //      w_sctNo = p_sct.gakuNo;
        //
        //  //挿入処理
        //      document.frm.SearchNm.value = w_sctNm.value;
        //      document.frm.SearchNo.value = w_sctNo.value;
        //  return true;    
        //}

        //************************************************************
        //  [機能]  クリアボタンをクリックした場合
        //  [引数]
        //  [戻値]
        //  [説明]
        //************************************************************
        function f_Clear(p_No) {

            document.frm.SearchNm.value = "";
            document.frm.SearchNo.value = "";

            //挿入させたいフォームを取得
				w_NmStr = eval("opener.document.frm.gakuNm" + p_No);
				w_NoStr = eval("opener.document.frm.gakuNo" + p_No);
                
            //挿入処理
                w_NmStr.value = "";
                w_NoStr.value = "";
            return true;    
        }

    //-->
    </script>


</head>

<body onload="focus();">
<form name="frm">
<center>
<%call gs_title("各種委員登録","参　照")%>

<br>

<table border="0" cellpadding="1" cellspacing="1" width="500">
	<tr>
		<td align="center" colspan="2">

		    <table class="hyo">
			    <tr>
				    <td align="center" width="150"><font color="white"><%=m_sIinNm%></font></td>
				    <td align="center" width="250" class="detail"><input type="text" class="CELL2" name="SearchNm" value="<%=w_sKNM%>" readonly><input type="hidden" name="SearchNo" value="<%=gf_fmtZero(m_rs("T13_GAKUSEI_NO"),10)%>"></td>
			    </tr>
		    </table>
			<br>
			<form>
			    <input type="button" value=" クリア " class="button" onclick="javascript:f_Clear('<%=m_iI%>');">　
			    <input type="button" value="閉じる" class="button" onclick="javascript:window.close();">
			</form>
			<span class="msg">※選択をするには名前をクリックし、閉じるボタンをクリックしてください</span><br>

		</td>
	</tr>
	<tr>
		<td align="center" width="250" valign="top">

		    <table border="1" class="hyo">
		    <tr>
				<!-- 2002/04/23 miyai -->
				<th width="80" class="header" nowrap><%=gf_GetGakuNomei(Session("NENDO"),C_K_KOJIN_1NEN)%></th>
		        <th width="170" class="header" nowrap>氏名</th>
		    </tr>
		    <%
		        m_rs.MoveFirst
		        w_Half = gf_Round(m_rCnt / 2,0)
		        Do Until m_rs.EOF
		            Call gs_cellPtn(w_cell)
		            j = j + 1 
		            If w_Half + 1 = j then
		            w_cell = ""
		            Call gs_cellPtn(w_cell)
		    %>
		    </table>

		</td>
		<td align="center" width="250" valign="top">

		    <table border="1" class="hyo">
			    <tr>
					<!-- 2002/04/23 miyai -->
					<th width="80" class="header" nowrap><%=gf_GetGakuNomei(Session("NENDO"),C_K_KOJIN_1NEN)%></th>
				    <th width="170" class="header" nowrap>氏名</th>
			    </tr>
		    <% End If %>
			    <tr><form>
				    <td class="<%=w_cell%>" align="center"><%=m_rs("T13_GAKUSEKI_NO")%><input type="hidden" name="gakuNo" value="<%=m_rs("T13_GAKUSEI_NO")%>"></td>
				    <td class="<%=w_cell%>" align="left"><input type="button" class="<%=w_cell%>" <%=w_btnWidth%> name="gakuNm" value="<%=gf_SetNull2String(m_rs("T11_SIMEI"))%>" onclick="iinSelect(this.form,'<%=m_iI%>')"></td>
			    </form>
			    </tr>
		    <%
		        m_rs.MoveNext
		        Loop%>

		    </table>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2">

			<br>
			<form>
			    <input type="button" value=" クリア " class="button" onclick="javascript:f_Clear('<%=m_iI%>');">　
			    <input type="button" value="閉じる" class="button" onclick="javascript:window.close();">
			</form>

		</td>
	</tr>
</table>

<INPUT TYPE="HIDDEN" NAME="GAKUNEN" VALUE="<%=request("m_sGakunen") %>">
<INPUT TYPE="HIDDEN" NAME="CLASS"   VALUE="<%=request("m_sClass") %>">
<INPUT TYPE="HIDDEN" NAME="IINNM"   VALUE="<%=request("m_sIinNm") %>">
<INPUT TYPE="HIDDEN" NAME="i"       VALUE="<%=request("m_iI") %>">

</center>
</form>
</center>
</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>