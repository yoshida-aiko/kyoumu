<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 特別教室予約
' ﾌﾟﾛｸﾞﾗﾑID : web/web0300/web0300_top.asp
' 機      能: 上ページ 表示情報を表示
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
' 変      数:
' 引      渡:   NENDO           '//:年
'               KYOKAN_CD       '//教官CD
'               cboGakunenCd    '//学年
'               cboClassCd      '//クラス
'               cboSikenKbn     '//試験区分
'               cboSikenCd      '//試験CD
' 説      明:
'           ■初期表示
'               アクセス権限がFULLの場合は、利用者を変更できる
'-------------------------------------------------------------------------
' 作      成: 2001/08/06 伊藤公子
' 変      更: 2001/08/07 根本 直美     NN対応に伴うソース変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////

'	Const C_ACCESS_FULL   = "FULL"		'//アクセス権限FULLアクセス可
'	Const C_ACCESS_NORMAL = "NORMAL"	'//アクセス権限一般
'	Const C_ACCESS_VIEW   = "VIEW"		'//アクセス権限参照のみ

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iSyoriNen          '//教官ｺｰﾄﾞ
    Public m_iKyokanCd          '//年度
    Public m_iTuki
    Public m_sLoginId

    '//コンボ用Where条件等
    Public m_sKyosituWhere      '//教室取得条件
    Public m_sKyokanName        '//教官名称
    Public m_sKengen			'//アクセス権限

    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
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
    w_sMsgTitle="特別教室予約"
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

        '//値の初期化
        Call s_ClearParam()

        '//変数セット
        Call s_SetParam()

        '//教室コンボに関するWHEREを作成する
        Call s_MakeKyosituWhere() 

		'//権限を取得
		w_iRet = gf_GetKengen_web0300(m_sKengen)
		If w_iRet <> 0 Then
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
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [機能]  変数初期化
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
	m_sLoginId  = ""
    m_iTuki     = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")
	m_sLoginId  = trim(Session("LOGIN_ID"))
    m_iTuki     = month(date())

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sLoginId  = " & m_sLoginId  & "<br>"
    response.write "m_iTuki     = " & m_iTuki     & "<br>"

End Sub

'********************************************************************************
'*  [機能]  教室コンボに関するWHEREを作成する
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_MakeKyosituWhere()

    m_sKyosituWhere = ""
    m_sKyosituWhere = m_sKyosituWhere & "     M06_NENDO = " & m_iSyorinen
	'使用フラグが使用可のものだけ表示 2001/12/11 
    m_sKyosituWhere = m_sKyosituWhere & " AND M06_SIYO_FLG = '1'"

End Sub

'********************************************************************************
'*  [機能]  教官の氏名を取得
'*  [引数]  なし
'*  [戻値]  p_sName
'*  [説明]  
'********************************************************************************
Function f_GetKyokanNm(p_sCD,p_iNENDO,p_sName)
Dim rs
Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetKyokanNm = 1
    w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M04_KYOKAN_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M04_NENDO = " & p_iNENDO & " "

'response.write w_sSQL

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetKyokanNm = 99
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M04_KYOKANMEI_SEI") & "　" & rs("M04_KYOKANMEI_MEI")
        End If

        f_GetKyokanNm = 0
        Exit Do
    Loop

    p_sName = w_sName

End Function

'****************************************************
'[機能] データ1とデータ2が同じ時は "SELECTED" を返す
'       (リストダウンボックス選択表示用)
'[引数] pData1 : データ１
'       pData2 : データ２
'[戻値] f_Selected : "SELECTED" OR ""
'                   
'****************************************************
Function f_Selected(pData1,pData2)

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else 
            f_Selected = "" 
        End If
    End If

End Function

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>特別教室予約</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [機能]  ページロード時処理
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function window_onload() {
    }

    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Show(){

        if(document.frm.SKyokanNm1.value==""){
			alert("利用者を選択してください")
			return;
		}

        document.frm.action="./web0300_main.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  クリアボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function fj_Clear(){

		//下画面を空白表示
		parent.main.location.href="default2.asp"

		//利用者欄を空白にする
		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}

    //************************************************************
    //  [機能]  教官参照選択画面ウィンドウオープン
    //  [引数]
    //  [戻値]
    //  [説明]
    //************************************************************
    function KyokanWin(p_iInt,p_sKNm) {

		//下画面を空白表示
		parent.main.location.href="default2.asp"

		var obj=eval("document.frm."+p_sKNm)

		//利用者選択画面を表示する
        //URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
        URL = "../../Common/com_select/Sel_User/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
//2015.8.19 UPDATE URAKAWA 表示サイズを変更
//        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=600,top=0,left=0");
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=570,height=640,top=0,left=0");
        nWin.focus();
        return true;    
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <%call gs_title("特別教室予約","一　覧")%>

    <form name="frm" method="post" onClick="return false;">
<%
'//デバッグ
'Call s_DebugPrint()
%>

    <center>
        <table border="0">
        <tr>
        <td class="search">
            <table border="0" cellpadding="1" cellspacing="1">
            <tr>
            <td align="left">

                <table border="0" cellpadding="1" cellspacing="1">
                <tr>
                <td Nowrap align="left">教室</td>
                <td Nowrap align="left">
                    <% call gf_ComboSet("cboKyositu",C_CBO_M06_KYOSITU,m_sKyosituWhere,"style='width:220px;' ",False,m_iSikenKbn) %>
                </td>
                <td Nowrap align="left">
                    <select name="TUKI">
                        <option value="4"  <%=f_Selected("4" ,cstr(m_iTuki))%> >4
                        <option value="5"  <%=f_Selected("5" ,cstr(m_iTuki))%> >5
                        <option value="6"  <%=f_Selected("6" ,cstr(m_iTuki))%> >6
                        <option value="7"  <%=f_Selected("7" ,cstr(m_iTuki))%> >7
                        <option value="8"  <%=f_Selected("8" ,cstr(m_iTuki))%> >8
                        <option value="9"  <%=f_Selected("9" ,cstr(m_iTuki))%> >9
                        <option value="10" <%=f_Selected("10",cstr(m_iTuki))%> >10
                        <option value="11" <%=f_Selected("11",cstr(m_iTuki))%> >11
                        <option value="12" <%=f_Selected("12",cstr(m_iTuki))%> >12
                        <option value="1"  <%=f_Selected("1" ,cstr(m_iTuki))%> >1
                        <option value="2"  <%=f_Selected("2" ,cstr(m_iTuki))%> >2
                        <option value="3"  <%=f_Selected("3" ,cstr(m_iTuki))%> >3
                    </select>月</td>
                </tr>

                <tr>
                <td Nowrap align="left">利用者</td>
                <td align="left" nowrap colspan="1">
                    <input type="text" class="text" name="SKyokanNm1" VALUE='<%=gf_GetUserNm(m_iSyoriNen,m_sLoginId)%>' readonly>
                    <input type="hidden" name="SKyokanCd1" VALUE='<%=m_sLoginId%>'>
					<%
					'//最高権限者のみ利用者の変更を可とする
					If m_sKengen = C_ACCESS_FULL Then%>
	                    <input type="button" class="button" value="選択" onclick="KyokanWin(1,'SKyokanNm1')">
						<input type="button" class="button" value="クリア" onClick="fj_Clear()">
					<%End If%>
                </td>
                <td align="right" nowrap colspan="1">
		        <input class="button" type="button" value="　表　示　" onclick="javascript:f_Show()">
                </td>
                </tr>
                </table>
            </td>
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