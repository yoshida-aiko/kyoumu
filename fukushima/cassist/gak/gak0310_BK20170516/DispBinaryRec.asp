<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�����(�摜�\��)
' ��۸���ID : gak/gak0310/DispBinaryRec.asp
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
' ��      ��: 2011/04/05 ��c DispBinary ���쐬(DB���摜�f�[�^���擾����)
' ��      �X: 2001/07/02
' ��      �X: 2005/05/06 ��O�@BLOB�^�Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

    '�擾�����f�[�^�����ϐ�
    Public  m_sGakuseiNo           ':�w���ԍ�
    Public  m_ImgRs                ':�w���ʐ^Image�f�[�^

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
	Dim w_sSQL

    'BLOB�^�Ή��̈גǉ� DB�ڑ���oo4o�ōs����gf_AutoOpen���ōs���Ă��� Dim OraDynaset As OraDynaset
    Dim wOraDyn
    Dim Chunksize, BytesRead, CurChunkEx

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
        '// ���Ұ�SET(	'�w���ԍ�)
				m_sGakuseiNo = request("gakuNo")
        if Trim(m_sGakuseiNo) = "" then exit do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & " T09_IMAGE "
        w_sSQL = w_sSQL & " FROM T09_GAKU_IMG "
        w_sSQL = w_sSQL & " WHERE T09_GAKUSEI_NO = '" & cstr(m_sGakuseiNo) & "'"


        Set wOraDyn = Session("OraDatabasePh").CreateDynaset(w_sSQL, 0)

        '// �摜�̏o��

        '//// BLOB�^�Ή��̈�
        BytesRead = 0
        'Reading in 32K chunks
        ChunkSize= 32768
        i = 0

        Do
          Response.Expires=0
          Response.ContentType="image/jpeg"
          BytesRead = wOraDyn.Fields("T09_IMAGE").GetChunkByteEx(CurChunkEx, i * ChunkSize, ChunkSize)
          if BytesRead > 0 then
            Response.BinaryWrite CurChunkEx
          end if
          i = i + 1
        Loop Until BytesRead < ChunkSize

        Exit Do

    Loop

    '// �I������
    Set wOraDyn = Nothing

End Sub

%>
