<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�����(�摜�\��)
' ��۸���ID : gak/gak0310/DispBinary.asp
' �@      �\: ���y�[�W �w�Ѓf�[�^�̊w���ʐ^Image�f�[�^��\������
'-------------------------------------------------------------------------
' ��      ��:�Ȃ�
' ��      �n:txtGakuseiNo           :�w���ԍ�
'            txtMode                :���샂�[�h
' ��      ��:
'           �������\��
'               �Ȃ�
'           �����ʕ\��
'               �w���ԍ����w���ʐ^Image�f�[�^���摜Binary�f�[�^�Ƃ��đ��M
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/02
' ��      �X: 2005/05/06 ��O�@BLOB�^�Ή�
' ��      �X: 2023/11/24 ���{�@oo4o�p�~�ɂ��摜�f�[�^�ǂݍ��ݕ��@��ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

    '�擾�����f�[�^�����ϐ�
    Public  m_sGakuseiNo           ':�w���ԍ�
    Public  m_Rs		   'recordset	'2023.11.24 ADD kiyomoto

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s

    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

Sub Main()
'********************************************************************************
'*  [�@�\]  �摜���擾����BINARY�Ƃ���Responce����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  Global.asa�Ő錾���Ă���N�G��Session("qurs")���g�p����
'********************************************************************************

    'BLOB�^�Ή��̈גǉ� DB�ڑ���oo4o�ōs����gf_AutoOpen���ōs���Ă���
    Dim wOraDyn
    Dim Chunksize, BytesRead, CurChunkEx
	Dim w_iRet              '// �߂�l	'2023.11.24 ADD kiyomoto
    Dim w_sSQL              '// SQL��	'2023.11.24 ADD kiyomoto
    
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�w����񌟍�����"
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

        '// �ް��ް��ڑ�
		w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If
' 20231124 kiyomoto ADD ---------------------------------------------ED

'Response.Write request("gakuNo")
        '// ���Ұ�SET(	'�w���ԍ�)
		m_sGakuseiNo = request("gakuNo")
        if Trim(m_sGakuseiNo) = "" then exit do
        
 ' 20231124 kiyomoto DEL ---------------------------------------------ST       
        'Session("OraDatabase").Parameters("IMG_KEY").value = m_sGakuseiNo
        'Session("qurs").Refresh
        'If Err.number <> 0 Then
        '    'ں��޾�Ă̎擾���s
        '    m_bErrFlg = True
        '   Exit Do
        'End If
' 20231124 kiyomoto DEL ---------------------------------------------ED

' 20231124 kiyomoto ADD ---------------------------------------------ST
        'Response.ContentType="image/jpeg"
        '�f�[�^���oSQL���쐬����
        Call s_MakeSQL(w_sSQL)

        '���R�[�h�Z�b�g�̎擾
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// �y�[�W��\��
        If Not m_Rs.EOF Then
			Response.BinaryWrite m_Rs("T09_IMAGE")
		Else
			Response.Write "Img0000000000.gif"
        End If

        Exit Do
' 20231124 kiyomoto ADD ---------------------------------------------ED

' 20231124 kiyomoto DEL ---------------------------------------------ST
       ' '// �摜�̏o��
       'if Not Session("qurs").EOF then
		'	 '//// BLOB�^�Ή��̈�
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
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// �I������
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
' 20231124 kiyomoto ADD ---------------------------------------------ED

End Sub
' 20231124 kiyomoto ADD ---------------------------------------------ST
Sub s_MakeSQL(p_sSql)
'********************************************************************************
'*  [�@�\]  �w�Ѓf�[�^���oSQL������̍쐬
'*  [����]  p_sSql - SQL������
'*  [�ߒl]  �Ȃ� 
'*  [����]  
'********************************************************************************

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " T09_IMAGE "
    p_sSql = p_sSql & " FROM T09_GAKU_IMG "
    p_sSql = p_sSql & " WHERE T09_GAKUSEI_NO = '" & cstr(m_sGakuseiNo) & "'"

End Sub
' 20231124 kiyomoto ADD ---------------------------------------------ED

%>
