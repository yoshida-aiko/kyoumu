<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 操作マニュアル
' ﾌﾟﾛｸﾞﾗﾑID : mst/manual/default.asp
' 機      能: リンク先のページの変更を行う
'-------------------------------------------------------------------------
' 引      数:なし
' 変      数:なし
' 引      渡:m_sLinkNo	:選択されたリンク先ナンバー
' 説      明:
'           ■初期表示
'			任意のページをメインフレームに表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/26 岩下　幸一郎
' 変      更: 
'*************************************************************************/

	Public	m_shtmlName

%>
<!--#include file="../Common/com_All.asp"-->
<html>

<head>
<title>教務事務システムマニュアル：Campus Assist manual</title>
<link rel=stylesheet href="../common/style.css" type=text/css>
<script language="javascript">
<!--
    //************************************************************
    //  [機能]  任意のページを表示
    //  [引数]  txtpage :表示頁
    //  [戻値]  なし
    //  [説明]
    //
    //************************************************************
    function f_PageClick(p_No){

        document.frm.action = "./main.asp";
        document.frm.target = "_top";
	document.frm.txtLinkNo.value = p_No;
        document.frm.submit();

    }
//-->
</script>
</head>

<body marginheight=0 marginwidth=0 bgcolor="#ffffff" topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0">
<div align="center">

<form name="frm" action="./main.asp" target="" Method="POST">

<table class=manual cellspacing="0" cellpadding="0" width=504 height=100% border="0">
	<tr>
		<td class=manual colspan="2" width="504" height="35%" align="center">
		<img src="../image/title.gif" width="504" height="214">
		</td>
	</tr>
	<tr>
		<td class=manual align="right" valign="top" height="65%" width="504">

			<img src="../image/sp.gif" height="15"><br>

			<table border="0" width=100% cellspacing="0" cellpadding="0">
			<tr>
			<td class=manual colspan="3">
				<table width=100% bgcolor="#3A449E" cellspacing="0" cellpadding="0" border="0">
					<tr>
						<td align="center"><font size="3" color="#ffffff"><b>操作マニュアル</b></font></td>
					</tr>
				</table>
			<img src="../image/sp.gif" height="5"><br>

			</td>
			</tr>

			<tr>
			<td valign="top">
				<table width=156 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>1・システム概要</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(1)">システム概要</a>
						</td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif"></td>
					</tr>
					<tr>
						<td class=manual><b>2・システム基本操作</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(2)">キーボード操作</a>
						</td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(3)">マウス操作</a>
						</td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(4)">画面操作</a>
						</td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif"></td>
					</tr>
					<tr>
						<td class=manual><b>3・ログイン</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(5)">ログイン</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif"></td>
					</tr>
					<tr>
						<td class=manual><b>4・メインメニュー</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(6)">メインメニュー</a></td>
					</tr>
				</table>
			</td>
			<td valign="top">

				<table width=174 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>5・出欠入力</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(7)">授業出欠入力</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(8)">行事出欠入力</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(9)">日毎出欠入力</a></td>
					</tr>
					<tr>
						<td class=manual><b>6・試験・成績</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(10)">試験実施科目登録</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(11)">試験監督免除申請登録</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(12)">成績登録</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(13)">試験時間割（クラス別）</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(14)">試験期間教官予定一覧</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(15)">成績一覧</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(16)">個人別成績一覧</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(17)">留年該当者一覧</a></td>
					</tr>
					<tr>
						<td class=manual><b>7・スケジュール</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(18)">行事日程一覧</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(19)">クラス別授業時間一覧</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(20)">教官別授業時間一覧</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(21)">時間割交換連絡</a></td>
					</tr>
				</table>
			</td>
			<td valign="top">
				<table width=174 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>8・その他入力</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(22)">進路先情報登録</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(23)">使用教科書登録</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(24)">指導要録所見等登録</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(25)">各種委員登録</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(26)">個人履修選択科目登録</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(27)">部活動部員登録</a></td>
					</tr>
					<tr>
						<td class=manual><b>9・情報検索</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(28)">学生情報検索</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(29)">中学校情報検索</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(30)">高等学校情報検索</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(31)">進路先情報検索</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(32)">空き時間</a></td>
					</tr>
					<tr>
						<td class=manual><b>10・支援機能</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(33)">特別教室予約</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(34)">連絡掲示板</a></td>
					</tr>
				</table>
			</td>
			</tr>
			</table>
		</td>
		<td>
			<img src="../image/sp.gif" width="10">
		</td>
	</tr>
</table>

<input type="hidden" name="txtLinkNo" value="">
</form>

</body>

</html>