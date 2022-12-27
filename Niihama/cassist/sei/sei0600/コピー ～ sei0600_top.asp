<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 欠席日数登録
' ﾌﾟﾛｸﾞﾗﾑID : gak/sei0600/sei0600_top.asp
' 機      能: 上ページ 試験毎所見登録の検索を行う
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :年度           ＞      SESSIONより（保留）
' 説      明:
'           ■初期表示
'               コンボボックスは空白で表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう調査書の内容を表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/09/26 谷脇
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ

    '市町村選択用のWhere条件
    Public m_iNendo         '年度
    Public m_sKyokanCd      '教官コード
    Public m_iSikenKBN   '試験区分
    Public m_sGakuNo        '氏名コンボボックスに入る値
    Public m_sGakuNoWhere   '氏名コンボボックスのwhere条件

    Public  m_Rs
    Public  m_Rs_Siken
    Public  m_iMax          '最大ページ
    Public  m_iDsp          '一覧表示行数

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
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakuNo   = request("txtGakuNo")
    m_iSikenKBN   = request("txtSikenKBN")
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

        '//学年の対象のデータ取得
        w_iRet = f_getData()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

        Call f_GakuNoWhere()
        
		'=====================
		'//試験コンボを取得
		'=====================
        w_iRet = f_GetSiken()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'==============================================
		'//試験区分をｾｯﾄする(科目取得時に使用)
		'==============================================
		If Request("txtSikenKBN")  = "" Then

			'//最近の試験区分を取得
            w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_SEISEKI_KIKAN,m_rs("M05_GAKUNEN"))
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If

		Else
		    m_iSikenKbn = Request("txtSikenKBN")    '//コンボ試験区分
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

'********************************************************************************
'*  [機能]  試験コンボを取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetSiken()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & "  M01_SYOBUNRUI_CD"
		w_sSQL = w_sSQL & vbCrLf & " ,M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_SYOBUNRUI_CD < " & cint(C_SIKEN_JITURYOKU)
		w_sSQL = w_sSQL & vbCrLf & "  ORDER BY M01_SYOBUNRUI_CD"

        iRet = gf_GetRecordset(m_Rs_Siken, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If


        f_GetSiken = 0
        Exit Do
    Loop

End Function

Function f_Selected(pData1,pData2)
'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'[引数] pData1 : データ１
'       pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'****************************************************

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else 
            f_Selected = "" 
        End If
    End If

End Function

Function f_getData()
'********************************************************************************
'*  [機能]  学年の対象のデータ取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_getData = 1

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     M05_CLASS "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     M05_NENDO = '" & m_iNendo & "' "
        w_sSQL = w_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            f_getData = 99
            m_bErrFlg = True
            Exit Do 
        End If

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

%>
<html>

<head>

<title>欠席日数登録</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        document.frm.action="sei0600_main.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  クリアボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Clear(){

        document.frm.txtGakuNo.value = "";
    
    }

    //-->
    </SCRIPT>

    <link rel="stylesheet" href="../../common/style.css" type="text/css">

</head>

<body>

<center>

<form name="frm" METHOD="post">

<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
<td valign="top" align="center">
<%call gs_title("欠席日数登録","登　録")%>
<br>
    <table border="0">
    <tr>
    <td class="search">
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left">
            <table border="0" cellpadding="1" cellspacing="1">
	                    <tr valign="middle">
	                        <td align="left">試験区分</td>
	                        <td align="left">
								<%If m_Rs_Siken.EOF Then%>
									<select name="txtSikenKBN" style='width:150px;' DISABLED>
										<option value="">データがありません
								<%Else%>
									<select name="txtSikenKBN" style='width:150px;' onchange = 'javascript:f_ReLoadMyPage()'>
									<%Do Until m_Rs_Siken.EOF%>
										<option value='<%=m_Rs_Siken("M01_SYOBUNRUI_CD")%>'  <%=f_Selected(cstr(m_Rs_Siken("M01_SYOBUNRUI_CD")),cstr(m_iSikenKbn))%>><%=m_Rs_Siken("M01_SYOBUNRUIMEI")%>
										<%m_Rs_Siken.MoveNext%>
									<%Loop%>
								<%End If%>
								</select>
							</td>
            <td Nowrap align="center">　クラス　</td>
            <td Nowrap><%=m_Rs("M05_GAKUNEN")%>年　<%=m_Rs("M05_CLASSMEI")%></td>
            </tr>
			<tr>
		        <td colspan="4" align="right">
		        <input type="button" class="button" value=" ク　リ　ア " onclick="javasript:f_Clear();">
		        <input type="button" class="button" value="　表　示　" onclick="javasript:f_Search();">
		        </td>
			</tr>
            </table>
        </td>
        </tr>
        </table>
    </td>
    </tr>
    </table>
</td>
</tr>
</table>
	<input type="hidden" name="txtGakunen" value="<%=m_Rs("M05_GAKUNEN")%>">
	<input type="hidden" name="txtClass" value="<%=m_Rs("M05_CLASSNO")%>">
	<input type="hidden" name="txtClassNm" value="<%=m_Rs("M05_CLASSMEI")%>">
</form>

</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
