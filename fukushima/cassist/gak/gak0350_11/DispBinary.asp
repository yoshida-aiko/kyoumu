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
' 変      更: 2023/11/24 清本　oo4o廃止により画像データ読み込み方法を変更
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////

    '取得したデータを持つ変数
    Public  m_sGakuseiNo           ':学生番号
    Public  m_Rs		   'recordset	'2023.11.24 ADD kiyomoto

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
	Dim w_iRet              '// 戻り値	'2023.11.24 ADD kiyomoto
    Dim w_sSQL              '// SQL文	'2023.11.24 ADD kiyomoto
    
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
' 20231124 kiyomoto ADD ---------------------------------------------ST
        Response.Expires = 0
        Response.Buffer = TRUE
        Response.Clear

        '// ﾃﾞｰﾀﾍﾞｰｽ接続
		w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            'ﾃﾞｰﾀﾍﾞｰｽとの接続に失敗
            m_bErrFlg = True
            m_sErrMsg = "データベースとの接続に失敗しました。"
            Exit Do
        End If
' 20231124 kiyomoto ADD ---------------------------------------------ED

'Response.Write request("gakuNo")
        '// ﾊﾟﾗﾒｰﾀSET(	'学生番号)
		m_sGakuseiNo = request("gakuNo")
        if Trim(m_sGakuseiNo) = "" then exit do
        
 ' 20231124 kiyomoto DEL ---------------------------------------------ST       
        'Session("OraDatabase").Parameters("IMG_KEY").value = m_sGakuseiNo
        'Session("qurs").Refresh
        'If Err.number <> 0 Then
        '    'ﾚｺｰﾄﾞｾｯﾄの取得失敗
        '    m_bErrFlg = True
        '   Exit Do
        'End If
' 20231124 kiyomoto DEL ---------------------------------------------ED

' 20231124 kiyomoto ADD ---------------------------------------------ST
        'Response.ContentType="image/jpeg"
        'データ抽出SQLを作成する
        Call s_MakeSQL(w_sSQL)

        'レコードセットの取得
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

        If w_iRet <> 0 Then
            'ﾚｺｰﾄﾞｾｯﾄの取得失敗
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// ページを表示
        If Not m_Rs.EOF Then
			Response.BinaryWrite m_Rs("T09_IMAGE")
		Else
			Response.Write "Img0000000000.gif"
        End If

        Exit Do
' 20231124 kiyomoto ADD ---------------------------------------------ED

' 20231124 kiyomoto DEL ---------------------------------------------ST
       ' '// 画像の出力
       'if Not Session("qurs").EOF then
		'	 '//// BLOB型対応の為
		'	 BytesRead = 0
		'	 'Reading in 32K chunks
		'	 ChunkSize= 32768
		'	 i = 0
		'	 Do
        '       Response.Expires=0
        '       Response.ContentType="image/jpeg"
		'	   BytesRead = Session("qurs").Fields("T09_IMAGE").GetChunkByteEx(CurChunkEx, i * ChunkSize, ChunkSize)
		'	   if BytesRead > 0 then
		'	      Response.BinaryWrite CurChunkEx
		'	    end if
		'	    i = i + 1
		'	 Loop Until BytesRead < ChunkSize
       'End If
' 20231124 kiyomoto DEL ---------------------------------------------ED

       Exit Do

    Loop

' 20231124 kiyomoto ADD ---------------------------------------------ST
    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
' 20231124 kiyomoto ADD ---------------------------------------------ED

End Sub
' 20231124 kiyomoto ADD ---------------------------------------------ST
Sub s_MakeSQL(p_sSql)
'********************************************************************************
'*  [機能]  学籍データ抽出SQL文字列の作成
'*  [引数]  p_sSql - SQL文字列
'*  [戻値]  なし 
'*  [説明]  
'********************************************************************************

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " T09_IMAGE "
    p_sSql = p_sSql & " FROM T09_GAKU_IMG "
    p_sSql = p_sSql & " WHERE T09_GAKUSEI_NO = '" & cstr(m_sGakuseiNo) & "'"

End Sub
' 20231124 kiyomoto ADD ---------------------------------------------ED

%>
