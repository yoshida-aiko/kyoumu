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
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

    '�擾�����f�[�^�����ϐ�
    Public  m_sGakuseiNo           ':�w���ԍ�

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
'Response.Write request("gakuNo")
        '// ���Ұ�SET(	'�w���ԍ�)
		m_sGakuseiNo = request("gakuNo")
        if Trim(m_sGakuseiNo) = "" then exit do
        Session("OraDatabase").Parameters("IMG_KEY").value = m_sGakuseiNo
        Session("qurs").Refresh
        If Err.number <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
           Exit Do
        End If

        '// �摜�̏o��
       if Not Session("qurs").EOF then
			 '//// BLOB�^�Ή��̈�
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
