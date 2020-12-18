<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 行事出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0140/kks0140_top.asp
' 機      能: 上ページ 月ごとの行事情報を表示する
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSION("KYOKAN_CD")
'            年度           ＞      SESSION("NENDO")
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
'               月のコンボボックスは当月を表示
'           ■表示ボタンクリック時
'               下のフレームに指定した条件にかなう行事一覧を表示させる
'           ■登録ボタンクリック時
'               入力された情報を登録する
'-------------------------------------------------------------------------
' 作      成: 2001/07/03 伊藤公子
' 変      更: 2001/12/07 佐野大悟 行事区分の追加に対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙコンスト /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    Public m_iSyoriNen          '教官ｺｰﾄﾞ
    Public m_iKyokanCd          '年度

    Public m_sGakunen           '//学年
    Public m_sClassNo           '//クラスNO
    Public m_sClassMei          '//クラス名
    Public m_sTuki              '//月
    Public m_Rs
    Public m_sNoTanMsg

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
            Call gs_SetErrMsg("データベースとの接続に失敗しました。")
            Exit Do
        End If

        '// 不正アクセスチェック
        Call gf_userChk(session("PRJ_No"))

        '//値の初期化
        Call s_ClearParam()

        '//変数セット
        Call s_SetParam()

        '//月情報を取得
        If request("TUKI") <> "" Then
            m_sTuki = request("TUKI")
        Else
            m_sTuki = month(date())
        End If

        '//年度、教官CDより担任クラス情報を取得
        w_iRet = f_GetClassInfo()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//担任クラスがあるかどうか
        If trim(m_sGakunen) = "" AND trim(m_sClassNo) = "" Then
            '//受持クラスがない時
            m_sNoTanMsg = "担任クラスがありません"
        Else
            '//担任クラスがあるときのみ表示
            '// ヘッダリスト情報取得
            w_iRet = f_Get_HeadData()
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If
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

    '//ﾚｺｰﾄﾞｾｯﾄCLOSE
    Call gf_closeObject(m_Rs)
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
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sClassMei = ""

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

End Sub

'********************************************************************************
'*  [機能]  デバッグ用
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_DebugPrint()
Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"
    response.write "m_sClassMei = " & m_sClassMei & "<br>"

End Sub

'********************************************************************************
'*  [機能]  教官CDより、担任クラス情報を取得
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Function f_GetClassInfo()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetClassInfo = 1

    Do
        'クラスマスタからクラス情報を取得
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "    M05_NENDO,"
        w_sSQL = w_sSQL & "    M05_GAKUNEN,"
        w_sSQL = w_sSQL & "    M05_CLASSNO,"
        w_sSQL = w_sSQL & "    M05_CLASSMEI"
        w_sSQL = w_sSQL & " FROM M05_CLASS"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "       M05_TANNIN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & "   AND M05_NENDO = " & cInt(m_iSyoriNen)

'response.write w_sSQL & "<br>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetClassInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            m_sGakunen  = rs("M05_GAKUNEN")
            m_sClassNo  = rs("M05_CLASSNO")
            m_sClassMei = rs("M05_CLASSMEI")
        End If

        '//正常終了
        f_GetClassInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  行事情報取得処理を行う
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_HeadData()

    Dim w_sSQL
    Dim w_Rs
    Dim w_sSDate,w_sEDate

    On Error Resume Next
    Err.Clear
    
    f_Get_HeadData = 1

    Do 

        '// 行事ヘッダデータ
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_MEI, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_KAISI_BI, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_SYURYO_BI, "
        w_sSQL = w_sSQL & vbCrLf & "  max(T32_GYOJI_M.T32_SOJIKANSU) AS T32_SOJIKANSU "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H, "
        w_sSQL = w_sSQL & vbCrLf & "  T32_GYOJI_M "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_CD = T32_GYOJI_M.T32_GYOJI_CD(+) AND "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_NENDO = T32_GYOJI_M.T32_NENDO(+) AND"
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_KYUKA_FLG='0' AND "   '//長期休暇FLG (0:通常 1:長期休暇 2:祝日)
'2001/12/07 「両方」を追加
'        w_sSQL = w_sSQL & vbCrLf & "  T32_GYOJI_M.T32_COUNT_KBN='0' AND "   '//カウント区分(0:行事 1:授業 2:その他)
        w_sSQL = w_sSQL & vbCrLf & "  (T32_GYOJI_M.T32_COUNT_KBN='0' OR "   '//カウント区分(0:行事 1:授業 2:その他 3:両方)
        w_sSQL = w_sSQL & vbCrLf & "  T32_GYOJI_M.T32_COUNT_KBN='3') AND "

        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  (T32_GYOJI_M.T32_TAISYO_GAKUNEN=" & cInt(m_sGakunen) & " Or T32_GYOJI_M.T32_TAISYO_GAKUNEN=" & C_GAKUNEN_ALL & ") AND "   '//対象学年(0:全学年 1-5:1-5年 )
        w_sSQL = w_sSQL & vbCrLf & "  (T32_GYOJI_M.T32_TAISYO_CLASS="   & cInt(m_sClassNo) & " Or T32_GYOJI_M.T32_TAISYO_CLASS=" & C_CLASS_ALL & ")"        '//対象クラス(99:全クラス)
        w_sSQL = w_sSQL & vbCrLf & " AND SUBSTR((T32_GYOJI_M.T32_HIDUKE),6,2)='" & gf_fmtZero(m_sTuki,2) & "'"
        w_sSQL = w_sSQL & vbCrLf & "  GROUP BY"
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_MEI, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_KAISI_BI, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_SYURYO_BI "

'response.write "<font color=#000000>" & w_sSQL & "<BR>"
        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then

            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_HeadData = 99
            Exit Do
        End If

        '//正常終了
        f_Get_HeadData = 0
        Exit Do
    Loop

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
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>行事出欠入力</title>

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

        <%
        '//担任クラスがない場合
        If m_sNoTanMsg <> "" Then%>
            parent.main.location.href="default2.asp?NoTanMsg=<%=m_sNoTanMsg%>"
        <%End If%>

    }

    //************************************************************
    //  [機能]  表示ボタンクリック時の処理
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_Search(){

        //行事がない場合
        if (document.frm.GYOJI.value==""){
            //parent.main.location.href = "default2.asp"
            alert("行事がありません");
            return;
        };

        var vl = document.frm.GYOJI.value.split('_')

        document.frm.GYOJI_CD.value  = vl[0];
        document.frm.GYOJI_MEI.value = vl[1];
        document.frm.KAISI_BI.value  = vl[2];
        document.frm.SYURYO_BI.value = vl[3];
        document.frm.SOJIKANSU.value = vl[4];

		//リスト画面表示
        document.frm.action="./kks0140_bottom.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [機能]  月を変更した時
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_ChangeTuki(){

        //本画面をsubmit
        document.frm.target = "topFrame";
        document.frm.action = "./kks0140_top.asp"
        document.frm.submit();
        return;
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

<%
'//デバッグ
'Call s_DebugPrint()
%>

    <center>
    <%call gs_title("行事出欠入力","一　覧")%>
    <%Do %>

        <%
        '//担任クラス情報がない時
        If m_sGakunen = "" Or m_sClassMei = "" Then
            response.write "<span class=msg>" & m_sNoTanMsg & "<span>"
            Exit Do
        End If
        %>
		<br>
        <table >
        <tr><td class="search">
                    <table cellpadding="1" cellspacing="1">
                        <tr>
                        <td nowrap >　<%=m_sGakunen%>年&nbsp;&nbsp;<%=m_sClassMei%></td>
                        <td nowrap >　<select name="TUKI" onchange="javascript:f_ChangeTuki();" style="width:50px;">
                                <option value="4"  <%=f_Selected("4" ,cstr(m_sTuki))%> >4
                                <option value="5"  <%=f_Selected("5" ,cstr(m_sTuki))%> >5
                                <option value="6"  <%=f_Selected("6" ,cstr(m_sTuki))%> >6
                                <option value="7"  <%=f_Selected("7" ,cstr(m_sTuki))%> >7
                                <option value="8"  <%=f_Selected("8" ,cstr(m_sTuki))%> >8
                                <option value="9"  <%=f_Selected("9" ,cstr(m_sTuki))%> >9
                                <option value="10" <%=f_Selected("10",cstr(m_sTuki))%> >10
                                <option value="11" <%=f_Selected("11",cstr(m_sTuki))%> >11
                                <option value="12" <%=f_Selected("12",cstr(m_sTuki))%> >12
                                <option value="1"  <%=f_Selected("1" ,cstr(m_sTuki))%> >1
                                <option value="2"  <%=f_Selected("2" ,cstr(m_sTuki))%> >2
                                <option value="3"  <%=f_Selected("3" ,cstr(m_sTuki))%> >3
                            </select></td>
						<td>月</td>
						<td>&nbsp;&nbsp;行事</td>
                        <td nowrap  valign="middle" >

                        <%If m_Rs.EOF Then%>
                            <select name="GYOJI" style='width:200px;' DISABLED>
                                <option value="">行事がありません
                        <%Else%>
                            <select name="GYOJI" style='width:200px;'>
                            <%Do Until m_Rs.EOF%>
                                <option value=<%=m_Rs("T31_GYOJI_CD") & "_" & m_Rs("T31_GYOJI_MEI") & "_" & m_Rs("T31_KAISI_BI") & "_" & m_Rs("T31_SYURYO_BI") & "_" & m_Rs("T32_SOJIKANSU")%>>&nbsp;<%=m_Rs("T31_GYOJI_MEI")%>&nbsp;&nbsp;

                                <%m_Rs.MoveNext%>
                            <%Loop%>
                        <%End If%>

                            </select>
                        </td>
						<td valign="bottom" align="right">
			            <input class="button" type="button" onclick="javascript:f_Search();" value="　表　示　">
						</tr>
                    </table>
		        </td>
	        </tr>
        </table>

        <%Exit Do%>
    <%Loop%>

    </center>

    <!--値渡し用-->
    <INPUT TYPE=HIDDEN NAME="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE=HIDDEN NAME="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
    <INPUT TYPE=HIDDEN NAME="GAKUNEN"   value = "<%=m_sGakunen%>">
    <INPUT TYPE=HIDDEN NAME="CLASSNO"   value = "<%=m_sClassNo%>">
    <INPUT TYPE=HIDDEN NAME="GYOJI_CD"  value = "">
    <INPUT TYPE=HIDDEN NAME="GYOJI_MEI" value = "">
    <INPUT TYPE=HIDDEN NAME="KAISI_BI"  value = "">
    <INPUT TYPE=HIDDEN NAME="SYURYO_BI" value = "">
    <INPUT TYPE=HIDDEN NAME="SOJIKANSU" value = "">

    </form>
    </body>
    </html>
<%
    '---------- HTML END   ----------
End Sub
%>
