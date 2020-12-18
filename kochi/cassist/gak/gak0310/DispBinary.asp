<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索結果(画像表示)
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0310/DispBinary.asp
' 機      能: 下ページ 学籍データの学生写真Imageデータを表示する
'-------------------------------------------------------------------------
' 変      数:なし
' 引      渡:txtGakuseiNo           :学生番号
'            txtMode                :動作モード
' 説      明:
'           ■初期表示
'               なし
'           ■結果表示
'               学生番号より学生写真Imageデータを画像Binaryデータとして送信
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
' 変      更: 2005/05/06 大前　BLOB型対応
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

    '取得したデータを持つ変数
    Public  m_sGakuseiNo           ':学生番号

'///////////////////////////メイン処理/////////////////////////////

    'ﾒｲﾝﾙｰﾁﾝ実行

    Call Main()

'///////////////////////////　ＥＮＤ　/////////////////////////////

Sub Main()
'********************************************************************************
'*  [機能]  画像を取得してBINARYとしてResponceする
'*  [引数]  なし
'*  [戻値]  なし
'*  [説明]  Global.asaで宣言しているクエリSession("qurs")を使用する
'********************************************************************************

    'BLOB型対応の為追加 DB接続もoo4oで行うがgf_AutoOpen内で行っている
    Dim wOraDyn
    Dim Chunksize, BytesRead, CurChunkEx

    'Message用の変数の初期化
    w_sWinTitle="キャンパスアシスト"
    w_sMsgTitle="学生情報検索結果"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do
'Response.Write request("gakuNo")
        '// ﾊﾟﾗﾒｰﾀSET(	'学生番号)
		m_sGakuseiNo = request("gakuNo")
        if Trim(m_sGakuseiNo) = "" then exit do
        Session("OraDatabase").Parameters("IMG_KEY").value = m_sGakuseiNo
        Session("qurs").Refresh
        If Err.number <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
           Exit Do
        End If

        '// 画像の出力
       if Not Session("qurs").EOF then
			 '//// BLOB型対応の為
			 BytesRead = 0
			 'Reading in 32K chunks
			 ChunkSize= 32768
			 i = 0
			 Do
               Response.Expires=0
               Response.ContentType="image/jpeg"
			   BytesRead = Session("qurs").Fields("T09_IMAGE").GetChunkByteEx(CurChunkEx, i * ChunkSize, ChunkSize)
			   if BytesRead > 0 then
			      Response.BinaryWrite CurChunkEx
			    end if
			    i = i + 1
			 Loop Until BytesRead < ChunkSize
       End If

       Exit Do

    Loop

End Sub

%>
