<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 成績登録
' ﾌﾟﾛｸﾞﾗﾑID : sei/sei0100/sei0150_paste.asp
' 機      能: 成績貼り付け用
'-------------------------------------------------------------------------
' 引      数:学生総数・貼り付け対象(成績・遅刻・欠課)
' 変      数:なし
' 説      明:クリップボードから成績データを取得する
'-------------------------------------------------------------------------
' 作      成: 2002/02/04 佐野 大悟
' 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
Dim m_iSeisekiInpType
Dim m_KekkaGaiDispFlg	'//欠課外表示フラグ
Dim m_bKekkaNyuryokuFlg	'//欠課入力期間フラグ

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
	w_sMsgTitle="成績登録"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_parent"


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
		
		'// 権限チェックに使用
		session("PRJ_No") = "SEI0100"
		
		'// 不正アクセスチェック
		Call gf_userChk(session("PRJ_No"))
		
		m_iSeisekiInpType = cint(request("hidSeisekiInpType"))
		m_KekkaGaiDispFlg = request("hidKekkaGaiDispFlg")
		m_bKekkaNyuryokuFlg	= request("hidKekkaNyuryokuFlg")
		
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

Sub showPage()
'********************************************************************************
'*  [機能]  HTMLを出力
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  
'********************************************************************************

    '---------- HTML START ----------
    %>
<html>
<head>
<title>成績貼り付け転送</title>
<link rel=stylesheet href="../../common/style.css" type="text/css">
<script language=javascript>
<!--
        //************************************************************
        //  [機能]  クリアボタンをクリックした場合
        //  [引数]
        //  [戻値]
        //  [説明]
        //************************************************************
        function f_Clear(p_No) {

            document.frm.paste.value = "";

            return true;    
        }

        //************************************************************
        //  [機能]  貼り付けボタンをクリックした場合
        //  [引数]
        //  [戻値]
        //  [説明]
        //************************************************************
        function f_Paste() {
			var str;
			var i;
			var textbox;
			var strLen;
			var strMaxLen = "3";
			var w_sObj = "<%=request("PasteType")%>";

			if(w_sObj == "txtKekka"){
				strMaxLen = "3"
			}
			if(w_sObj == "txtKibi"){
				strMaxLen = "2"
			}
			if(w_sObj == "txtTeisi"){
				strMaxLen = "3"
			}
			if(w_sObj == "txtHaken"){
				strMaxLen = "2"
			}

			//未入力チェック
			if (document.frm.paste.value=="") {
				alert("転送対象データがありません。");
				return false;
			}

			//貼り付け文字列の取得
			str = (document.frm.paste.value).split("\n");
			strLen = str.length;
			
			//学生数でのループ
			for(i=1;i<=<%=request("i_Max")-1%>;i++) {
				//親ウィンドウに存在するかどうか
				textbox = eval("opener.parent.main.document.frm.<%=request("PasteType")%>" + i);
				
				//(取得できたデータ数に関係なく全データを一旦クリアする)
				//if (textbox && i<=strLen + 1) {
				if (textbox){
					
					//ﾃｷｽﾄﾎﾞｯｸｽにﾛｯｸがかかってなかったら
					if(textbox.readOnly == false && textbox.disabled == false){
						
						//初期化
						textbox.value = "";
						
						if (str[i-1] != "") {
							//数字でないのは無視する
							if (!isNaN(str[i-1])) {
								textbox.value = jf_Left(str[i-1],strMaxLen);
							}
						}
					}
				}
			}
			
			<% if m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_NUM) then %>
				//合計・平均の計算
				eval("opener.parent.main").f_GetTotalAvg();
			<% end if %>
			
			<% if m_bKekkaNyuryokuFlg then %>
				//遅刻の合計
//				eval("opener.parent.main").f_CalcSum("Chikai");
				
				//欠課の合計
//				eval("opener.parent.main").f_CalcSum("Kekka");
				
				<% if m_KekkaGaiDispFlg then %>
					//欠課外の合計
//					eval("opener.parent.main").f_CalcSum("KekkaGai");
				<% end if %>
			<% end if %>
			
			window.close();
        }

	//**************************************************************************************
	//////////////////////////   左から指定数の文字を抜き取る   ////////////////////////////
	//--------------------------------------------------------------------------------------
	// Arguments: String length
	// Return:  String
	//**************************************************************************************
	function jf_Left(str,len){
	 if(str==null) return "";
	 if(len==null) return "";
	 if(isNaN(len)) return "";
	 str = str.substr(0,len);
	 return str;
	}

    //-->
    </script>

</head>

<body>
<form name="frm">
<center>
<%call gs_title("成績貼り付け転送","登　録")%>

<br>

<table border="0" cellpadding="1" cellspacing="1">
	<tr>
		<td align="center" colspan="2">

			<span class="msg">※Excelファイルからコピーしたデータを<BR>貼り付けてください。</span><br>

		</td>
	</tr>
	<tr>
		<td align="center" width="250" valign="top">
			<textarea name="paste" COLS="20" ROWS="27"></textarea>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2">
			<br>
		    <input type="button" value=" 転　送 " class="button" onclick="javascript:f_Paste();">　
		    <input type="button" value=" クリア " class="button" onclick="javascript:f_Clear();">　
		    <input type="button" value="閉じる" class="button" onclick="javascript:window.close();">
		</td>
	</tr>
</table>

</center>
</form>
</center>
</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>