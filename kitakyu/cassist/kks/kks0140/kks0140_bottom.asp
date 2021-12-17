<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 行事出欠入力
' ﾌﾟﾛｸﾞﾗﾑID : kks/kks0140/kks0140_bottom.asp
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

        '// 行事出欠明細情報取得
        w_iRet = f_Get_AbsInfo()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		If m_Rs_M.EOF = True Then
	        '// ページを表示
	        Call showWhitePage("生徒情報がありません")
		Else
	        '// ページを表示
	        Call showPage()
		End If

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

	call gf_Get_SyuketuEnd(cint(m_sGakunen),m_sEndDay)

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
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO,"
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
'*  [機能]  行事出欠データを取得
'*  [引数]  なし
'*  [戻値]  0:情報取得成功 99:失敗
'*  [説明]  
'********************************************************************************
Function f_Get_AbsInfo()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_Get_AbsInfo = 1

    Do 

        '// 出欠データ
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN," 
        w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_KEKKA "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_SYUKKETU "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "   T22_NENDO=" & cInt(m_iSyoriNen)   & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN=" & cInt(m_sGakunen)  & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS=" & cInt(m_sClassNo)    & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD='" & m_sGyoji_Cd & "'"

'response.write "<font color=#000000><br>" & w_sSQL & "<br>"
        iRet = gf_GetRecordset(m_Rs_G, w_sSQL)
        If iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            msMsg = Err.description
            f_Get_AbsInfo = 99
            Exit Do
        End If

        f_Get_AbsInfo = 0
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

        //ヘッダ部を表示submit
        document.frm.target = "topFrame";
        document.frm.action = "kks0140_middle.asp"
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
        parent.document.location.href="default.asp"

    }

    //************************************************************
    //  [機能]  入力チェック(onBlur時)
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_CheckData(p_ObjName,p_Total){

        var objName="document.frm."+p_ObjName
        var Kekka = eval(objName);

        if (f_Trim(Kekka.value)!=""){

            //if (isNaN(f_Trim(Kekka.value))){
            if (f_chkNumber(f_Trim(Kekka.value))==1){
                alert("入力値が不正です")
                Kekka.focus();
                return;
            }else{
                var vKekka = new Number(Kekka.value)
                var vTotal = new Number(p_Total)
                if(vKekka > vTotal){
                    alert("総時間を超えた時間が入力されています。")
                    Kekka.focus();
                    return;
                };
            };
        };

    };

    //************************************************************
    //  [機能]  数字チェック
    //  [引数]  p_num
    //  [戻値]  成功：0   失敗：1
    //  [説明]	数字かどうかをチェック(マイナス値、小数点有の場合はエラーを返す)
    //************************************************************
	function f_chkNumber(p_num){

		//数値チェック
		if (isNaN(p_num)){
			return 1;
		}else{

			//マイナスをチェック
			var wStr = new String(p_num)
			if (wStr.match("-")!=null){
				return 1;
			};

			//小数点チェック
			w_decimal = new Array();
			w_decimal = wStr.split(".")
			if(w_decimal.length>1){
				return 1;
			}

		};
		return 0;
	}

    //************************************************************
    //  [機能]  登録ボタンが押されたとき
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_Touroku(){

        //生徒数
        if (document.frm.iMax.value <= 0){
            //alert("データがありません。")
            return;
        };

		//入力チェック(NN対応)
		iRet = f_CheckData_All();
		if( iRet != 0 ){
			return;
		}

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		//ヘッダ部空白表示
		parent.topFrame.document.location.href="white.htm"

        //リスト情報をsubmit
        document.frm.target = "main";
        document.frm.action = "./kks0140_edt.asp"
        document.frm.submit();
        return;
    }

    //************************************************************
    //  [機能]  入力値のﾁｪｯｸ(登録ボタン押下時)
    //  [引数]  なし
    //  [戻値]  0:ﾁｪｯｸOK、1:ﾁｪｯｸｴﾗｰ
    //  [説明]  入力値のNULLﾁｪｯｸ、英数字ﾁｪｯｸ、桁数ﾁｪｯｸを行う
    //          引渡ﾃﾞｰﾀ用にﾃﾞｰﾀを加工する必要がある場合には加工を行う
    //************************************************************
    function f_CheckData_All() {

		if(document.frm.iMax.value==1){
			var wKekka = new String("SU_" & document.frm.GAKUSEKI_NO.value)

			iRet = f_CheckData_NN(wKekka,<%=m_sSoJikan%>);
			if( iRet != 0 ){
				return 1;
			}

		}else{

			var i
			var w_bCheck = 0
			for (i = 0; i < document.frm.iMax.value; i++) {
				var wKekka = new String("SU_" + document.frm.GAKUSEKI_NO[i].value)
				iRet = f_CheckData_NN(wKekka,<%=m_sSoJikan%>);
				if( iRet != 0 ){
					w_bCheck = 1;
					break;
				};
			};

			if (w_bCheck == 1){
				return 1;
			};
		};
		return 0;
	};

    //************************************************************
    //  [機能]  入力チェック(NN用)
    //  [引数]  なし
    //  [戻値]  なし
    //  [説明]
    //************************************************************
    function f_CheckData_NN(p_ObjName,p_Total){

        var objName="document.frm."+p_ObjName
        var Kekka = eval(objName);
        
		if (typeof(Kekka) != "undefined"){

			if (f_Trim(Kekka.value)!=""){

			    //if (isNaN(f_Trim(Kekka.value))){
			    if (f_chkNumber(f_Trim(Kekka.value))==1){
			        alert("入力値が不正です")
			        Kekka.focus();
			        return 1;
			    }else{
			        var vKekka = new Number(Kekka.value)
			        var vTotal = new Number(p_Total)
			        if(vKekka > vTotal){
			            alert("総時間を超えた時間が入力されています。")
			            Kekka.focus();
			        	return 1;
			        };
			    };
			};
		};
        return 0;
    };


    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>

    <form name="frm" method="post" onClick="return false;">
    <%Do%>

        <%If m_Rs_M.EOF = True Then%>
            <br><br>
            <span class="msg">生徒情報がありません</span>
			<%Exit Do%>
		<%End If%>

        <%If Trim(Request("GYOJI_CD")) = "" Then%>
        <%Else%>

            <!--明細リスト部-->

            <table >
                <tr><td valign="top" >
            <table class="hyo"  border="1" >

            <%i = 0%>
            <%If m_Rs_M.EOF = True Then%>
            <%Else%>

                <%
				'//改行カウント
	            w_iCnt = INT(m_iRsCnt/2 + 0.9)

				Dim w_IdouCnt
				Dim w_sIdouMei
				w_IdouCnt = 1

                Do Until m_Rs_M.EOF
                    i = i + 1

                    '//学籍NOを取得
                    w_iGakusekiNo = m_Rs_M("T13_GAKUSEKI_NO")

					'//異動がある場合取得する
					w_IdouCnt = gf_Set_Idou(Cstr(w_iGakusekiNo),m_iSyoriNen,w_sIdouMei)

                    '//ｽﾀｲﾙｼｰﾄのｸﾗｽをセット
                    Call gs_cellPtn(w_Class) 
                %>
                    <tr>
                        <td nowrap class=<%=w_Class%> width="80"  align="center"><%=w_iGakusekiNo%><input type="hidden" name="GAKUSEKI_NO" value="<%=w_iGakusekiNo%>"><br></td>
                        
						<td nowrap class=<%=w_Class%> width="150" align="left"><%=m_Rs_M("T11_SIMEI")%><br></td>
                        <%
                        If m_Rs_G.EOF = False Then
                            m_Rs_G.MoveFirst
                            Do Until m_Rs_G.EOF
                                w_iKekka = ""

                                If cStr(trim(m_Rs_G("T22_GAKUSEKI_NO"))) = cStr(trim(w_iGakusekiNo)) Then
     
	                               w_iKekka = gf_SetNull2Zero(m_Rs_G("T22_GYOJI_KEKKA"))

                                    Exit Do
                                End If
                                m_Rs_G.MoveNext
                            Loop
                            m_Rs_G.MoveFirst
                            If cInt(w_iKekka) = 0 Then
                                w_iKekka = ""
                            End If

                        End If
                        %>
                       	<% IF w_IdouCnt = 1 Then %>
						 	<td class="<%=w_Class%>" width="80" align="center">
					   	<% Else %>
							<td class="NOCHANGE" width="80" align="center" >
						<% End IF %>
						<%
						'//NN対応
						If session("browser") = "IE" Then
							w_sInputClass = "class='num'"
						Else
							w_sInputClass = ""
						End If

						%>
<% IF w_IdouCnt = 1 Then %>
					<% 'If m_sEndDay < m_sSyuryo_Bi then %>

							<input <%=w_sInputClass%> type="text" name="SU_<%=w_iGakusekiNo%>" value="<%=w_iKekka%>"  size="5" maxlength="2"><br></td>

					<% 'Else %>

							<%'=w_iKekka%><!-- <br></td> -->

					<% 'End If %>
<% Else %>

<%=w_sIdouMei%><br></td>

<%End if%>
                    </tr>
                    <%
					'//2列目表示
					If i = w_iCnt Then
					%>
                        </table>
                    </td>
					<td width="10"><br></td>
					<td valign="top" >
                        <table class="hyo"  border="1" >

                        <%'//ｽﾀｲﾙｼｰﾄのｸﾗｽをクリア
						w_Class = ""
						%>

                    <%End If%>

                    <%m_Rs_M.MoveNext%>
                <%Loop%>
            <%End If%>

                    </table>
                </td></tr>
            </table>
            <br>
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

        <%
        End If

        Exit Do

    Loop%>

    <!--値渡し用-->
    <input type="hidden" name="NENDO"     value="<%=m_iSyoriNen%>">
    <input type="hidden" name="KYOKAN_CD" value="<%=m_iKyokanCd%>">
    <input type="hidden" name="GAKUNEN"   value="<%=m_sGakunen%>">
    <input type="hidden" name="CLASSNO"   value="<%=m_sClassNo%>">
    <input type="hidden" name="TUKI"      value="<%=m_sTuki%>">
    <input type="hidden" name="iMax"      value="<%=i%>">

    <INPUT TYPE=HIDDEN NAME="GYOJI_CD"  value = "<%=m_sGyoji_Cd%>">
    <INPUT TYPE=HIDDEN NAME="GYOJI_MEI" value = "<%=m_sGyoji_Mei%>">
    <INPUT TYPE=HIDDEN NAME="KAISI_BI"  value = "<%=m_sKaisi_Bi%>">
    <INPUT TYPE=HIDDEN NAME="SYURYO_BI" value = "<%=m_sSyuryo_Bi%>">
    <INPUT TYPE=HIDDEN NAME="SOJIKANSU" value = "<%=m_sSoJikan%>">
    <INPUT TYPE=HIDDEN NAME="ENDDAY" value = "<%=m_sEndDay%>">

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
    <title>行事用出欠入力</title>
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