<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�����(�摜�\��)
' ��۸���ID : gak/gak0310/DispBinary.asp
' �@      �\: ���y�[�W �w�Ѓf�[�^�̊w���ʐ^Image�f�[�^��\������
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtGakuseiNo           :�w���ԍ�
'           txtMode                :���샂�[�h
'                               BLANK   :�����\��
'                               SEARCH  :���ʕ\��
' ��      ��:
'           �������\��
'               �^�C�g���̂ݕ\��
'           �����ʕ\��
'               �w���ԍ����w���ʐ^Image�f�[�^���摜�\������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/02
' ��      �X: 2005/05/06 ��O�@BLOB�^�Ή�
' ��      �X: 2022/12/27 �g�c�@�ʐ^���\������悤�ɏC��(�{��gak0300/DispBinary.asp�Ɠ��l�̃\�[�X�ɏC�����R�����gADD�̕�����ǉ�)
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�擾�����f�[�^�����ϐ�
    Public  m_TxtMode      	       ':���샂�[�h
    Public  m_sGakuseiNo           ':�w���ԍ�
    
    Public	m_Rs					'recordset
    Public	m_iDsp					'// �ꗗ�\���s��

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s

    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

Sub Main()
'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

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
        '2022.12.27�@ADD Yoshida�@-->
        Response.Expires = 0
        Response.Buffer = TRUE
        Response.Clear
        '2022.12.27�@ADD Yoshida�@<--

        '// �ް��ް��ڑ�
		w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

        '// ���Ұ�SET
		m_sGakuseiNo = request("gakuNo")           	'�w���ԍ�
		if Trim(m_sGakuseiNo) = "" then exit do

        '2022.12.27�@ADD Yoshida�@-->
        Response.ContentType="image/jpeg"
        '2022.12.27�@ADD Yoshida�@<--

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
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// �I������
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub


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

%>