<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 日毎出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0170/kks0170_bottom.asp
' 機      能: 下ページ 授業出欠入力の一覧リスト表示を行う
'-------------------------------------------------------------------------
' 引      数: SESSION("NENDO")           '//処理年
'             SESSION("KYOKAN_CD")       '//教官CD
'             TUKI           '//月
'             cboDate        '//日付
' 変      数:
' 引      渡: NENDO"        '//処理年
'             KYOKAN_CD     '//教官CD
'             GAKUNEN"      '//学年
'             CLASSNO"      '//ｸﾗｽNo
'             cboDate"      '//日付
' 説      明:
'           ■初期表示
'               検索条件にかなう担任ｸﾗｽ生徒情報を表示
'           ■登録ボタンクリック時
'               入力情報を登録する
'-------------------------------------------------------------------------
' 作      成: 2001/07/24 伊藤公子
' 変      更: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙCONST /////////////////////////////
    Const C_SYOBUNRUICD_IPPAN = 4   '//欠席区分(0:出席,1:欠席,2:遅刻,3:早退,4:忌引,…)

'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    Public  m_bTannin       '//担任ﾌﾗｸﾞ

    '取得したデータを持つ変数
    Public m_iSyoriNen      '//処理年度
    Public m_iKyokanCd      '//教官CD
    Public m_sDate          '//日付
    Public m_iGakunen       '//学年
    Public m_iClassNo       '//クラスNo
    Public m_sClassNm       '//クラス名称
    Public m_iRsCnt         '//クラスﾚｺｰﾄﾞ数
    Public m_sDispMsg       '//エラー時メッセージ
	Public m_sEndDay		'//入力できなくなる日

    'ﾚｺｰﾄﾞセット
    Public m_Rs

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
    w_sMsgTitle="日毎出欠入力"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_bTannin = False

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

        '// 担任クラス情報取得
        w_iRet = f_GetClassInfo(m_bTannin)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'入力不可になる日を取得
		call gf_Get_SyuketuEnd(m_iGakunen,m_sEndDay)

        '//ログイン教官が担任クラスを持っているとき生徒リストを取得
        If m_bTannin = True Then
            '// 生徒リスト情報取得
            w_iRet = f_GetClassList()
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If

			If m_Rs.EOF Then
		        '// データなしの場合、空白ページを表示
		        Call showWhitePage("クラス情報がありません")
			Else
		        '// 詳細ページを表示
		        Call showPage()
			End If
		Else
	        '// 担任クラスをもっていない場合、空白ページを表示
	        Call showWhitePage("受持クラスがありません。")
        End If

        Exit Do
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// 終了処理
    Call gf_closeObject(m_Rs)
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
    m_sDate     = ""
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_sClassNm = ""

End Sub

'********************************************************************************
'*  [機能]  全項目に引き渡されてきた値を設定
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = SESSION("NENDO")
    m_iKyokanCd = SESSION("KYOKAN_CD")

    m_sDate     = trim(Request("cboDate"))
    If m_sDate = "" Then
        m_sDate = gf_YYYY_MM_DD(date(),"/")
    End If

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
    response.write "m_sDate     = " & m_sDate     & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_sClassNm  = " & m_sClassNm  & "<br>"

End Sub

'********************************************************************************
'*  [機能]  担任クラス情報取得
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetClassInfo(p_bTannin)

    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear
    
    f_GetClassInfo = 1

    Do 

        '// 担任クラス情報
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_TANNIN"
        w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & m_iKyokanCd & "'"

'response.write w_sSQL & "<BR>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetClassInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            p_bTannin = True 
            m_iGakunen = rs("M05_GAKUNEN")
            m_iClassNo = rs("M05_CLASSNO")
            m_sClassNm = rs("M05_CLASSMEI")
        End If

        f_GetClassInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [機能]  担任クラス一覧取得
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_GetClassList()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetClassList = 1

    Do 

        '// 担任クラス情報取得
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_IDOU_NUM, "
        w_sSQL = w_sSQL & vbCrLf & "  B.T11_SIMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  B.T11_GAKUSEI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  C.T30_HIDUKE, "
        w_sSQL = w_sSQL & vbCrLf & "  C.T30_SYUKKETU_KBN,"
        '//"出席"は表示しない
        w_sSQL = w_sSQL & vbCrLf & "  DECODE(D.M01_SYOBUNRUIMEI_R,'出','　',D.M01_SYOBUNRUIMEI_R) AS SYUKKETU_MEI"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN A"
        w_sSQL = w_sSQL & vbCrLf & "  ,T11_GAKUSEKI B"
        w_sSQL = w_sSQL & vbCrLf & "  ,(SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     T30_HIDUKE,"
        w_sSQL = w_sSQL & vbCrLf & "     T30_SYUKKETU_KBN,"
        w_sSQL = w_sSQL & vbCrLf & "     T30_GAKUSEKI_NO"
        w_sSQL = w_sSQL & vbCrLf & "    FROM T30_KESSEKI"
        w_sSQL = w_sSQL & vbCrLf & "    WHERE T30_HIDUKE='" & m_sDate & "'"
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_GAKUNEN=" & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_CLASS=" & m_iClassNo & ") C"
        w_sSQL = w_sSQL & vbCrLf & "  ,(SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     M01_SYOBUNRUI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "     M01_SYOBUNRUIMEI_R"
        w_sSQL = w_sSQL & vbCrLf & "    FROM M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & "    WHERE "
        w_sSQL = w_sSQL & vbCrLf & "          M01_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "      AND M01_DAIBUNRUI_CD=" & C_KESSEKI & ") D"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        'w_sSQL = w_sSQL & vbCrLf & "      A.T13_NENDO - A.T13_GAKUNEN + 1 = B.T11_NYUNENDO(+) "
        w_sSQL = w_sSQL & vbCrLf & "      A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO "
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_GAKUSEKI_NO = C.T30_GAKUSEKI_NO(+)"
        w_sSQL = w_sSQL & vbCrLf & "  AND C.T30_SYUKKETU_KBN = D.M01_SYOBUNRUI_CD(+)"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_GAKUNEN=" & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_CLASS=" & m_iClassNo
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T13_GAKUSEKI_NO"

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_GetClassList = 99
            Exit Do
        End If

        '//ﾚｺｰﾄﾞカウントを取得
        m_iRsCnt = gf_GetRsCount(m_Rs)

        f_GetClassList = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [機能]  移動ありの場合移動状況の取得
'*  [引数]  p_Gakusei_No:学績NO
'*          p_Date      :授業実施日
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_IdouInfo(p_Gakusei_No)

    Dim w_sSQL
    Dim w_Rs
    Dim w_sKubunName

    On Error Resume Next
    Err.Clear

    w_IdoFlg = False

    Do

        '// 移動情報
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_1, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_1, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_2, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_2, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_3, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_3, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_4, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_4, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_5, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_5, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_6, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_6, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_7, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_7, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_8, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_8"
        w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO='" & p_Gakusei_No & "' AND"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM>0"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            Exit Do
        End If

        If w_Rs.EOF = false Then

            i = 1
            Do Until i>8    '//8…最大移動回数

                If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) = "" Then
                    Exit Do
                End If

                If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) > m_sDate  Then
                    Exit Do
                End If
                i = i + 1
            Loop

            If i = 1 then
                '//最初の移動日が授業日より未来の場合、授業日に移動状態ではない
                w_sKubunName = ""
            Else
                '//移動中の場合、移動区分・移動理由を取得
                Select Case Trim(w_Rs("T13_IDOU_KBN_" & i-1))
                 Case cstr(C_IDO_FUKUGAKU),cstr(C_IDO_TEI_KAIJO)  '//C_IDO_FUKUGAKU=3:復学、C_IDO_TEI_KAIJO=5:停学解除
                    w_sKubunName = ""
                 Case Else
                    '//移動理由を取得(区分マスタ、大分類=C_IDO)
                    w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i-1)),m_iSyoriNen,w_sKubunName)
                    If w_bRet<> True Then
                        Exit Do
                    End If
                End Select

            End If

        End If

        Exit Do
    Loop

    f_Get_IdouInfo = w_sKubunName

    Call gf_closeObject(w_Rs)

    Err.Clear

End Function

'********************************************************************************
'*  [機能]  出欠区分と名称を取得(javascript生成用)
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  出欠入力のJAVASCRIPT作成
'********************************************************************************
Function f_Get_SYUKETU_KBN(p_MaxNo)

    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear
    
    f_Get_SYUKETU_KBN = 1

    Do 

        '// 明細データ
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI_R"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_KESSEKI) & " AND "
        '//C_SYOBUNRUICD_IPPAN = 4  '//欠席区分(0:出席,1:欠席,2:遅刻,3:早退,4:忌引,…)
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD<" & C_SYOBUNRUICD_IPPAN
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M01_KUBUN.M01_SYOBUNRUI_CD"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_SYUKETU_KBN = 99
            Exit Do
        End If

        i=0
        If rs.EOF = True Then
            response.write ("var ary = new Array(0);")
            response.write ("ary[0] = '　';")
        Else

            '//ﾚｺｰﾄﾞカウント取得
            w_iCnt = gf_GetRsCount(rs) - 1
            response.write ("var ary = new Array(" & w_iCnt & ");") & vbCrLf

            Do Until rs.EOF
                If i = 0 Then
                    response.write ("ary[0] = '　';") & vbCrLf
                Else
                    response.write ("ary[" & rs("M01_SYOBUNRUI_CD") &  "] = '" & rs("M01_SYOBUNRUIMEI_R") & "';") & vbCrLf
                End If

                i=i+1
                rs.MoveNext
            Loop

        End If

        p_MaxNo = w_iCnt

        f_Get_SYUKETU_KBN = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)
    Err.Clear

End Function

'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showPage()
Dim i 
Dim w_sIduoRiyu

    On Error Resume Next
    Err.Clear

%>
    <html>
    <head>
    <title>日毎出欠入力</title>
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

        //ヘッダ部を表示submit
        document.frm.txtMsg.value = "<%=m_sDispMsg%>";
        document.frm.target = "topFrame";
        document.frm.action = "kks0170_middle.asp"
        document.frm.submit();
        return;

    }

    //************************************************************
    //  [機能]  出欠入力
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function chg(chgInp) {

        no = 0;
        <%
        '//出欠区分を取得
        Call f_Get_SYUKETU_KBN(w_MaxNo)
        %>

        str = chgInp.value;
        for(i=0; i<<%=w_MaxNo+1%>; i++){
            if (ary[i]==str){
                break;
            }
        };

        no = i + 1;
        if (no > <%=w_MaxNo%>) no = 0;
        chgInp.value = ary[no];

        //隠しフィールドにデータをセット
        var obj=eval("document.frm.hid"+chgInp.name);
        obj.value=no;
        return;
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

		//ヘッダ部空白ページ表示
		parent.topFrame.location.href="./white.htm"

        //リスト情報をsubmit
        document.frm.target = "main";
        //document.frm.target = "<%=C_MAIN_FRAME%>";
        document.frm.action = "./kks0170_edt.asp"
        document.frm.submit();
        return;
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
        //parent.document.location.href="default2.asp"

        document.frm.target = "<%=C_MAIN_FRAME%>";
        document.frm.action = "./default.asp"
        document.frm.submit();
        return;

    }

    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">

<%
'//デバッグ
'Call s_DebugPrint()
%>
    <center>
    <form name="frm" method="post" onClick="return false;">

    <%Do
        '//担任クラスがない場合
        If m_bTannin = False Then
            Exit Do
        End If

        If m_Rs.EOF = False Then
            '//リスト改行カウントを取得
            w_iCnt = INT(m_iRsCnt/2 + 0.9)
        Else
            '//データなしの場合
            Exit Do
        End If%>

        <!--リスト部-->
        <table >
            <tr><td valign="top">

                <!--ヘッダ-->
                <table class=hyo border="1" bgcolor="#FFFFFF">

        <%
		'入力期間が過ぎていれば、入力はできない。
		if m_sEndDay < m_sDate then 
				w_tmp =  " onclick='return chg(this)'"
		else
				w_tmp = " DISABLED"
		End If

        i=1
        Do Until m_Rs.EOF

                    '//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
                    Call gs_cellPtn(w_Class) 

                    '//入力許可・非許可の判定
                    '//移動状況の考慮(T13_IDOU_NUMが1以上の場合は移動状況を判別する)移動中の場合は、出欠入力不可
                    If gf_SetNull2Zero(m_Rs_D("T13_IDOU_NUM")) > 0 Then 
                        w_sIduoRiyu = ""
                        w_sIduoRiyu = f_Get_IdouInfo(m_Rs("T11_GAKUSEI_NO"))
                    End If

                    %>

                    <!--詳細-->
                    <tr>
                        <td nowrap class="<%=w_Class%>" width="80"  align="center"><%=m_Rs("T13_GAKUSEKI_NO")%><input type="hidden" name="GAKUSEKI_NO" value="<%=m_Rs("T13_GAKUSEKI_NO")%>"><br></td>
                        <td nowrap class="<%=w_Class%>" width="150" align="left"><%=m_Rs("T11_SIMEI")%><br></td>
                        <%
                        If w_sIduoRiyu <> "" Then
                            %>
                            <td class="NOCHANGE" width="80" align="center" ><%=gf_SetNull2String(w_sIduoRiyu)%><br>
                            <input type="hidden" name="hidKBN<%=m_Rs("T13_GAKUSEKI_NO")%>" size="2" value="---"></td>
                            <%
                        Else
                            %>
                            <td class="<%=w_Class%>" width="80" align="center">
                            <input type="button" name="KBN<%=m_Rs("T13_GAKUSEKI_NO")%>" value="<%=gf_HTMLTableSTR(m_Rs("SYUKKETU_MEI"))%>" size="20" maxlength="2" class=<%=w_Class%> style="border-style:none" style="text-align:center" <%=w_tmp%>>
                            <input type="hidden" name="hidKBN<%=m_Rs("T13_GAKUSEKI_NO")%>" size="2" value="<%=gf_SetNull2Zero(m_Rs("T30_SYUKKETU_KBN"))%>"></td>
                            <%
                        End If
                        %>

                    </tr>

            <%If i = w_iCnt Then
                '//リストを改行する

                '//ｽﾀｲﾙｼｰﾄのｸﾗｽを初期化
				w_Class = ""
                %>
                </table>
                </td>
				<td width="10"></td>
				<td valign="top">
                <!--ヘッダ-->
                <table class="hyo" border="1" >

            <%End If

            i = i + 1
            m_Rs.MoveNext%>
        <%Loop%>

                </table>
                </td></tr>
            </table>
            <br>
<%		if m_sEndDay < m_sDate then %>
            <table>
				<tr>
                <td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="　登　録　"></td>
                <td ><input class="button" type="button" onclick="javascript:f_Cancel();" value="キャンセル"></td>
				</tr>
            </table>
<% Else %>
            <table>
				<tr>
                <td ><input class="button" type="button" onclick="javascript:f_Cancel();" value=" 戻　る "></td>
				</tr>
            </table>
<% End If %>
        <%Exit Do%>
    <%Loop%>

    <!--値渡し用-->
    <input type="hidden" name="NENDO"     value="<%=m_iSyoriNen%>">
    <input type="hidden" name="KYOKAN_CD" value="<%=m_iKyokanCd%>">
    <input type="hidden" name="GAKUNEN"   value="<%=m_iGakunen%>">
    <input type="hidden" name="CLASSNO"   value="<%=m_iClassNo%>">
    <input type="hidden" name="cboDate"   value="<%=m_sDate%>">

    <input type="hidden" name="txtMsg"    value="">

    </form>
    </center>
    </body>
    </html>
<%
End Sub

'********************************************************************************
'*  [機能]  空白HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
    <html>
    <head>
    <title>日毎出欠入力</title>
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
		//parent.location.href="white.htm"
    }
    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">

	<center>
	<br><br><br>
		<span class="msg"><%=p_Msg%></span>
	</center>

    </body>
    </html>
<%
End Sub
%>
