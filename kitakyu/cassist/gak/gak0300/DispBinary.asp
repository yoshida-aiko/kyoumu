<%@ Language=VBScript %>
<%
'/************************************************************************
' システム名: 教務事務システム
' 処  理  名: 学生情報検索結果(画像表示
' ﾌﾟﾛｸﾞﾗﾑID : gak/gak0300/DispBinary.asp
' 機      能: 下ページ 学籍データの学生写真Imageデータを表示する
'-------------------------------------------------------------------------
' 引      数:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
' 変      数:なし
' 引      渡:教官コード     ＞      SESSIONより（保留）
'           :処理年度       ＞      SESSIONより（保留）
'           txtGakuseiNo           :学生番号
'           txtMode                :動作モード
'                               BLANK   :初期表示
'                               SEARCH  :結果表示
' 説      明:
'           ■初期表示
'               タイトルのみ表示
'           ■結果表示
'               学生番号より学生写真Imageデータを画像表示する
'-------------------------------------------------------------------------
' 作      成: 2001/07/02 岩田
' 変      更: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// ﾓｼﾞｭｰﾙ変数 /////////////////////////////
    'エラー系
    Public  m_bErrFlg           'ｴﾗｰﾌﾗｸﾞ
    
    '取得したデータを持つ変数
    Public  m_TxtMode      	       ':動作モード
    Public  m_sGakuseiNo           ':学生番号
    
    Public	m_Rs					'recordset
    Public	m_iDsp					'// 一覧表示行数

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
    w_sMsgTitle="学生情報検索結果"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

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

        '// ﾊﾟﾗﾒｰﾀSET
		m_sGakuseiNo = request("gakuNo")           	'学生番号
		if Trim(m_sGakuseiNo) = "" then exit do

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
    Loop

    '// ｴﾗｰの場合はｴﾗｰﾍﾟｰｼﾞを表示（ﾏｽﾀﾒﾝﾃﾒﾆｭｰに戻る）
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// 終了処理
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub


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

%>


