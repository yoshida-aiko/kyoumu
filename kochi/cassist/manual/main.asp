<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 操作マニュアル
' ﾌﾟﾛｸﾞﾗﾑID : mst/manual/main.asp
' 機      能: リンク先のページの変更を行う
'-------------------------------------------------------------------------
' 引      数:m_sLinkNo	:選択された連絡先コード
' 変      数:なし
' 引      渡:m_spage	:表示させるhtml名
' 説      明:
'           ■初期表示
'			任意のページをメインフレームに表示させる
'-------------------------------------------------------------------------
' 作      成: 2001/07/26 岩下　幸一郎
' 変      更: 2001/09/25 岩下　幸一郎
'*************************************************************************/

	Public	m_spage			':ページNAME
	Public	m_sLinkNo		':順番につけられたリンクの番号

	m_sLinkNo = Request("txtLinkNo")


Select case m_sLinkNo
	Case 1
		m_spage = "login.htm"
	Case 2
		m_spage = "0-1.htm"
	Case 3
		m_spage = "0-2.htm"
	Case 4
		m_spage = "1-1.htm"
	Case 5
		m_spage = "1-2.htm"
	Case 6
		m_spage = "1-3.htm"
	Case 7
		m_spage = "2-1.htm"
	Case 8
		m_spage = "2-2.htm"
	Case 9
		m_spage = "2-3.htm"
	Case 10
		m_spage = "2-tuika1.htm"
	Case 11
		m_spage = "2-tuika2.htm"
	Case 12
		m_spage = "2-4.htm"
	Case 13
		m_spage = "2-5.htm"
	Case 14
		m_spage = "2-6.htm"
	Case 15
		m_spage = "2-7.htm"
	Case 16
		m_spage = "2-8.htm"
	Case 17
		m_spage = "3-1.htm"
	Case 18
		m_spage = "3-2.htm"
	Case 19
		m_spage = "3-3.htm"
	Case 20
		m_spage = "3-4.htm"
	Case 21
		m_spage = "4-1.htm"
	Case 22
		m_spage = "4-2.htm"
	Case 23
		m_spage = "4-3.htm"
	Case 24
		m_spage = "4-4.htm"
	Case 25
		m_spage = "4-5.htm"
	Case 26
		m_spage = "4-6.htm"
	Case 27
		m_spage = "4-7.htm"
	Case 28
		m_spage = "5-1.htm"
	Case 29
		m_spage = "5-2.htm"
	Case 30
		m_spage = "5-3.htm"
	Case 31
		m_spage = "5-4.htm"
	Case 32
		m_spage = "5-5.htm"
	Case 33
		m_spage = "6-1.htm"
	Case 34
		m_spage = "6-2.htm"
	Case 35
		m_spage = "2-tuika3.htm"
End select

%>

<HTML>
<HEAD>
<TITLE>教務事務システムマニュアル：Campus Assist manual</TITLE>
</HEAD>

<FRAMESET cols="190,*" BORDER=0 FRAMESPACING=0 FRAMEBORDER="NO">
<FRAME SRC="./menu.htm" NAME="menu" NORESIZE SCROLLING="auto">
<FRAME SRC="./<%= m_spage %>" NAME="main" NORESIZE SCROLLING="auto">
<NOFRAMES>
<BODY>
</BODY>
</NOFRAMES>
</FRAMESET>


</HTML>
