<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����\��}�X�^
' ��۸���ID : skn/skn0120/syossai.asp
' �@      �\: ���y�[�W �����\��}�X�^�̏ڍו\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j

'           txtSikenKbn      :�����敪
'           txtSikenCd      :�����R�[�h
'           txtMode         :���샂�[�h
'                           BLANK   :�o�^�̂��ߑS���ڋ󔒂ŕ\��
'                           INSERT  :�o�^���������s
'                           UPDATE  :�X�V���������s
'                           DISP    :�w�肳�ꂽ�敪�̃f�[�^��\��
'
'           cmbJissiDate    :���t
'           cmbRiyu         :���R
'           txtKaisiH       :�J�n�����i���j
'           txtKaisiN       :�J�n�����i���j
'           txtSyuRyoH      :�I�������i���j
'           txtSyuryoN      :�I�������i���j
'
'           txtRenban       :�A��
'           txtPage         :�\���Ő�
'
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtSikenKbn      :�����敪�i�߂�Ƃ��j
'           txtSikenCd      :�����R�[�h
'           txtMode         :���샂�[�h�i�߂�Ƃ��j
'                           BLANK   :�S���ڋ󔒂ŕ\��
'
'           cmbJissiDate    :���t
'           cmbRiyu         :���R
'           txtKaisiH       :�J�n�����i���j
'           txtKaisiN       :�J�n�����i���j
'           txtSyuryoH      :�I�������i���j
'           txtSyuryoN      :�I�������i���j
'
'           txtRenban       :�A��
'           txtPage         :�\���Ő�
'
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��������\���\��
'           ���X�V�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ������\����X�V������
'           ���o�^�{�^���N���b�N��
'               �����\����͂�\�������āA�o�^������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/16 ���u �m��
' ��      �X: 2001/06/26 ���{
' ��      �X: 2001/07/27 �ɓ����q M40_CALENDER�폜�̈וύX
' ��      �X: 2001/08/03 �ɓ����q �������Ԃ�\������悤�C��
' ��      �X: 2001/08/03 �ɓ����q ���t�̓��͂�FromTo�Ŕ͈͓��͂ł���p�ɕύX
'                                 �S���`�F�b�N�{�b�N�X��ǉ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٺݽ� /////////////////////////////
	CONST C_MIN_TIME = "00:00"		'//�ŏ�����
	CONST C_MAX_TIME = "23:55"		'//�ő厞��(����)
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�擾�����f�[�^�����ϐ�
    Public  m_iSikenKbn     ':�����敪
    Public  m_iSikenCd      ':�����R�[�h
    Public  m_sMode         ':���샂�[�h
    Public  m_iKyokanCd     ':�����R�[�h
    Public  m_iSyoriNen     ':�����N�x
    Public  m_iRenban       ':�A��
    Public  m_iJissiDate    ':���t
    Public  m_iJissiDateE    ':�I�����t
    Public  m_iJissiKaisi   ':�������{�J�n��
    Public  m_iJissiSyuryo  ':�������{�I����
    Public  m_sBiko         ':���l
    Public  m_iKaisiH       ':�J�n�����i���j
    Public  m_iKaisiN       ':�J�n�����i���j
    Public  m_iSyuryoH      ':�I�������i���j
    Public  m_iSyuryoN      ':�I�������i���j
    Public  m_iKaisi       ':�J�n����
    Public  m_iSyuryo      ':�I������
    Public  m_Rs            'recordset
    Public  m_iRenbanCount  '�A�Ԑ�
    Public  m_iKikanWhere
    Public  m_iPage     ':�\���ϕ\���Ő��i�������g����󂯎������j

    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bBack             '// ����I�����׸�
    Public  m_bMsgFlg           '// ү�����׸ށi���ᔽ�Ȃǂ̴װ�̏ꍇ���޲�۸ނ�\�����邽�߂��׸ށj
    Public  m_sDebugStr         '// �ȉ����ޯ�ޗp
    Public  m_sMsg              '// ү���ޗp

    Public  m_sMinDate,m_sMaxDate	'//�������Ԃ̍ŏ����t�A�ő���t

    '�y�[�W�֌W
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp                      '// �ꗗ�\���s��

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
    Dim w_sWHERE            '// WHERE��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�����ēƏ��o�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bBack = False
    m_bMsgFlg = False
    m_sMode = Request("txtMode")

    w_iRet = 0

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    Do
    
        '// �l�̏�����
        Call s_SetBlank()

        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        Call s_setParam()
                
			'===============================
			'//���ԃf�[�^�̎擾
			'===============================
	        w_iRet = f_Nyuryokudate()
			If w_iRet = 1 Then
				    '// �I������
				    Call gs_CloseDatabase()
					response.Redirect "default.asp?txtMode=no&txtSikenKbn="&m_iSikenKbn&""
					response.end
				Exit Do
			End If
			If w_iRet <> 0 Then 
				m_bErrFlg = True
				Exit Do
			End If

        '���R�R���{�Ɋւ���WHERE���쐬����i�X�V���ɕK�v�j
        Call c_MakeRiyuWhere() 

        '// �eӰ�ނɂ�菈���𕪂���
        Select Case m_sMode
            '// �������f�[�^�\���i�o�^�p�t�H�[���j
            Case "BLANK"
                m_sDebugStr = "BLANK"

                w_iRet = f_GetJissiDate()
                If w_iRet <> 0 Then
                    '�G���[����
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If
                
                '���ԃR���{�Ɋւ���WHERE���쐬����i�X�V���ɕK�v�j
                Call s_MakeKikanWhere()
                
                ' �y�[�W��\��
                Call showPage()

            '// �������f�[�^�\���i�X�V�p�t�H�[���j
            Case "DISP"
                '�w��f�[�^�擾
                m_sDebugStr = "DISP"

'                Call s_setParam()
                
                w_iRet = f_GetData()
                If w_iRet <> 0 Then
                    '�G���[����
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If

                w_iRet = f_GetJissiDate()
                If w_iRet <> 0 Then
                    '�G���[����
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If
                
                Call s_MakeRiyuWhere()
                
                '���ԃR���{�Ɋւ���WHERE���쐬����i�X�V���ɕK�v�j
                Call s_MakeKikanWhere()
                
                '�y�[�W��\��
                Call showPage()

            '// �������f�[�^�ǉ�
            Case "INSERT"
                m_sDebugStr = "INSERT"
                
'                Call s_SetParam()
                
                Call s_SetParamDate()

                w_iRet = f_GetJissiDate()
                If w_iRet <> 0 Then
                    '�G���[����
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If

				'//�o�^����
				w_iRet = f_Insert()
                If w_iRet = 1 Then
                    '�������d�����Ă���ꍇ
                    m_bMsgFlg = True
                    m_bBack = False
                    Call showPage()
                    Exit Do
				Else
                    If w_iRet <> 0 Then
                        '�G���[����
                        m_bErrFlg = True
                        m_sErrMsg = m_sMsg
                        Exit Do
                    Else
                        '����I���ňꗗ�\��ʂɖ߂�
                        m_bBack = True
                        Call showPage()

                    End If
                End If

            '// �������f�[�^�X�V
            Case "UPDATE"
                m_sDebugStr = "UPDATE"
                
'                Call s_SetParam()
                
                Call s_SetParamDate()

                w_iRet = f_GetJissiDate()
                If w_iRet <> 0 Then
                    '�G���[����
                    m_bErrFlg = True
                    m_sErrMsg = m_sMsg
                    Exit Do
                End If


				'//���͂��ꂽ��񂪂��łɓo�^����ĂȂ����`�F�b�N����
				w_bRet = f_ChkDate(m_iJissiDate,m_iRenban)
                if w_bRet = True Then
                    w_iRet = f_UpdateData()
                    If w_iRet <> 0 Then
                        m_bErrFlg = True
                        m_sErrMsg = m_sMsg
                        Exit Do
                    Else
                        '����I���ňꗗ�\��ʂɖ߂�
                        m_bBack = True
                        
                        '���ԃR���{�Ɋւ���WHERE���쐬����i�X�V���ɕK�v�j
                        Call s_MakeKikanWhere()
                        
                        Call showPage()
                    End If
                else
                        '�G���[����
                        '�������d�����Ă���ꍇ
                        m_bMsgFlg = True
                        m_bBack = False
                        Call showPage()
                        Exit Do
                
                end if
            '// ���������̑��i�G���[�j
            Case Else
                m_sDebugStr = "ETC"
                m_bErrFlg = True
                Call gs_SetErrMsg("�������[�h���ݒ肳��Ă��܂���(���Ѵװ)")
                Exit Do
        End Select
            
        '// ����I��
        Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        if m_sErrMsg <> "" Then
            w_sMsg = m_sErrMsg
        else
            w_sMsg = gf_GetErrMsg()
        end if
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()

    On Error Goto 0
    Err.Clear
            
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂ��󔒂ɏ�����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetBlank()

    m_iSikenKbn = ""
    m_iSikenCd = ""
    m_iRenban = ""
    m_iJissiDate = ""
    m_sBiko = ""
    m_iKaisiH = ""
    m_iKaisiN = ""
    m_iSyuryoH = ""
    m_iSyuryoN = ""
    m_iKaisi = ""
    m_iSyuryo = ""
    m_iRenbanCount = ""
    m_iSyoriNen = ""
    m_iPage = ""

End Sub

Function f_Nyuryokudate()
'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
	dim w_date

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1


	w_date = gf_YYYY_MM_DD(date(),"/")
'	w_date = "2000/06/18"

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MIN(T24_SIKEN_NITTEI.T24_SIKEN_KAISI) as KAISI"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SIKEN_SYURYO) as SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iSyoriNen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_iSikenKbn)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KAISI <= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_SYURYO >= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  Group By M01_SYOBUNRUIMEI"

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
'//���ѓ��͊��ԃe�X�g�p

'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '2002/04/30'"
'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '2000/03/01'"

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_

'response.write w_sSQL & "<<<BR>"

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If m_DRs.EOF Then
			Exit Do
		Else
			m_sSikenNm = m_DRs("M01_SYOBUNRUIMEI")
		End If
		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function


'********************************************************************************
'*  [�@�\]  �������{�J�n���E�I�������Z�b�g
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾�����A1:ں��ނȂ��A99:���s
'*  [����]  
'********************************************************************************
Function f_GetJissiDate()

    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    f_GetJissiDate = 0
    m_iJissiKaisi = ""
    m_iJissiSyuryo = ""

    Do 
        '// ��������ں��޾�Ă��擾
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & " T24_JISSI_KAISI"
        w_sSQL = w_sSQL & vbCrLf & " ,T24_JISSI_SYURYO"
        w_sSQL = w_sSQL & vbCrLf & " FROM T24_SIKEN_NITTEI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE T24_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND T24_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & vbCrLf & " AND T24_SIKEN_CD = '" & m_iSikenCd & "'"

'response.write w_sSQL & "<br>"

        Set w_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            'm_sMsg = "���������̎擾�Ɏ��s���܂���"
            m_sMsg = Err.description
            f_GetJissiDate = 99
            Exit Do
        End If
    
        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            m_sMsg = "�����������o�^����Ă��܂���"
            f_GetJissiDate = 1
            Exit Do
        End If

        m_iJissiKaisi = w_Rs("T24_JISSI_KAISI")      ':�������{�J�n��
        m_iJissiSyuryo = w_Rs("T24_JISSI_SYURYO")    ':�������{�I����
        
        Exit Do

    Loop

    gf_closeObject(w_Rs)
End Function

'********************************************************************************
'*  [�@�\]  ���擾�������s���i�f�[�^�X�V���\���Ɏg�p�j
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾�����A1:ں��ނȂ��A99:���s
'*  [����]  
'********************************************************************************
Function f_GetData()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim i                   '// ����
    
    On Error Resume Next
    Err.Clear
    f_GetData = 0

    Do 
        '// �����\��ں��޾�Ă��擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & " T25_NENDO"
        w_sSQL = w_sSQL & " ,T25_SIKEN_KBN"
        w_sSQL = w_sSQL & " ,T25_SIKEN_CD"
        w_sSQL = w_sSQL & " ,T25_KYOKAN"
        w_sSQL = w_sSQL & " ,T25_YOTEIBI"
        w_sSQL = w_sSQL & " ,T25_RENBAN"
        w_sSQL = w_sSQL & " ,T25_YOTEI_KAISI"
        w_sSQL = w_sSQL & " ,T25_YOTEI_SYURYO"
        w_sSQL = w_sSQL & " ,T25_BIKO"
        w_sSQL = w_sSQL & " FROM T25_KYOKAN_YOTEI "
        w_sSQL = w_sSQL & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        'w_sSQL = w_sSQL & " AND T25_YOTEIBI = '" & m_iJissiDate & "'"
        w_sSQL = w_sSQL & " AND T25_YOTEIBI = '" & gf_YYYY_MM_DD(Request("txtKeyYoteibi"),"/") & "'"
        w_sSQL = w_sSQL & " AND T25_RENBAN = " & m_iRenban

'response.write "w_sSQL = " &w_sSQL & "<BR>"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
       If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_sMsg = Err.description
            f_GetData = 99
            Exit Do 'GOTO LABEL_f_GetData_END
        End If

        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            f_GetData = 1
            m_sMsg = "�����\��̑Ώۃ��R�[�h������܂���"
            Exit Do
        End If

        '// �擾�����l���۰��ٕϐ��Ɋi�[
        m_iSikenKbn = w_Rs("T25_SIKEN_KBN")            ':�����敪
        m_iSikenCd = w_Rs("T25_SIKEN_CD")              ':�����R�[�h
        m_iRenban = w_Rs("T25_RENBAN")                 ':�A��
        m_iJissiDate = w_Rs("T25_YOTEIBI")             ':���t
        m_sBiko = w_Rs("T25_BIKO")                     ':���R
        m_iKaisiH = Left(w_Rs("T25_YOTEI_KAISI"),2)    ':�J�n�����i���j
        m_iKaisiN = Mid(w_Rs("T25_YOTEI_KAISI"),4,2)   ':�J�n�����i���j
        m_iSyuryoH = Left(w_Rs("T25_YOTEI_SYURYO"),2)  ':�I�������i���j
        m_iSyuryoN = Mid(w_Rs("T25_YOTEI_SYURYO"),4,2) ':�I�������i���j
        m_iKaisi = w_Rs("T25_YOTEI_KAISI")             ':�J�n����
        m_iSyuryo = w_Rs("T25_YOTEI_SYURYO")           ':�I������

        Exit Do

    Loop

    gf_closeObject(w_Rs)
    
'// LABEL_f_GetData_END
End Function

'********************************************************************************
'*  [�@�\]  �o�^����
'*  [����]  �Ȃ�
'*  [�ߒl]  0:����,1:�d������A99:���s
'*  [����]  
'********************************************************************************
Function f_Insert()

	f_Insert = 0

	Do

		'//���{���t���擾
		m_iJissiDate = Request("cmbJissiDate")  ':���t
		m_iJissiDateE = Request("cmbJissiDateE")  ':�I�����t

		If m_iJissiDateE <> "" Then
			iMax = DateDiff("d",m_iJissiDate,m_iJissiDateE)+1
		Else
			iMax = 1
		End If 

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

		'//�J�n������I�����܂ŁA�����ɂ�����INSERT����
		For i = 1 To imax
			w_sJissiBi = FormatDateTime(DateAdd("d",i-1,m_iJissiDate))

			'//�A�Ԏ擾
			w_iRet = f_GetCountRenban(w_sJissiBi,w_RenBan)
			If w_iRet <> 0 Then
			    '�G���[����
				f_Insert = 99
			    Exit Do
			End If

			'//�d���`�F�b�N
			w_bRet = f_ChkDate(w_sJissiBi,w_RenBan)
			if w_bRet = True Then

				'//�d�����Ă��Ȃ����o�^����
			    w_iRet = f_InsertData(w_sJissiBi,w_RenBan)
			    If w_iRet <> 0 Then
					'//۰��ޯ�
					Call gs_RollbackTrans()
			        '�G���[����
					f_Insert = 99
			        Exit Do
			    Else

			    End If
			else
		        '�������d�����Ă���ꍇ
				f_Insert = 1
		        Exit Do
			end if

		Next

	   '//����I�����A�Я�
	   Call gs_CommitTrans()

		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �A�Ԃ��J�E���g����i�V�K�o�^���Ɏg�p�j
'*  [����]  p_JissiBi
'*  [�ߒl]  0:���擾�����A99:���s
'*  [����]  
'********************************************************************************
Function f_GetCountRenban(p_JissiBi,p_RenBan)
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear

    f_GetCountRenban = 0
	p_RenBan = 0

    Do 
        '// �����\��ں��޾�Ă��擾�i�A��Max�l�j
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & " MAX("
        w_sSQL = w_sSQL & vbCrLf & " T25_RENBAN "
        w_sSQL = w_sSQL & vbCrLf & ")"
        w_sSQL = w_sSQL & vbCrLf & " AS T25_MAXRENBAN"
        w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_YOTEIBI = '" & p_JissiBi & "'"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_sMsg = Err.description
            f_GetCountRenban = 99
            Exit Do
        End If

		If ISNULL(w_Rs("T25_MAXRENBAN")) = False Then
			p_RenBan = cint(w_Rs("T25_MAXRENBAN")) + 1
		Else
			p_RenBan = 0
		End If

    	Exit Do

    Loop

	'//ں��޾��CLOSE
    gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  ���{�����A�w�莎�����Ԃ̊��ԓ��ɐݒ肳�ꂽ��
'*  [����]  �Ȃ�
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
'Function f_ChkDate()
Function f_ChkDate(p_sJissiBi,p_RenBan)


    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    f_ChkDate = True

    Do

        '// �����\��ں��޾�Ă��擾
        w_sSQL2 = ""
        w_sSQL2 = w_sSQL2 & "SELECT"
        w_sSQL2 = w_sSQL2 & " T25_YOTEI_KAISI"
        w_sSQL2 = w_sSQL2 & " ,T25_YOTEI_SYURYO"
        w_sSQL2 = w_sSQL2 & " ,T25_RENBAN"
        w_sSQL2 = w_sSQL2 & " FROM T25_KYOKAN_YOTEI "
        w_sSQL2 = w_sSQL2 & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL2 = w_sSQL2 & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL2 = w_sSQL2 & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL2 = w_sSQL2 & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL2 = w_sSQL2 & " AND T25_YOTEIBI = '" & p_sJissiBi & "'"
        w_sSQL2 = w_sSQL2 & " AND T25_RENBAN <> " & p_RenBan

        w_iRet2 = gf_GetRecordset(w_Rs2, w_sSQL2)
        If w_iRet2 <> 0 Then
            'ں��޾�Ă̎擾���s
            m_sMsg = Err.description
            f_ChkDate = False
            Exit Do
        End If

        If w_Rs2.EOF Then
            '�Ώ�ں��ނȂ�
            f_ChkDate = True
            Exit Do
        End If

        Do Until w_Rs2.EOF
            If CDate(m_iKaisi) < CDate(w_Rs2("T25_YOTEI_SYURYO")) Then
                if CDate(m_iSyuryo) <= CDate(w_Rs2("T25_YOTEI_KAISI")) Then
                else
                    m_sMsg = "���łɓo�^����Ă���\�莞�ԂƏd�����Ă��܂�\n" & m_iKaisi & "-" & m_iSyuryo & "<" & w_Rs2("T25_YOTEI_KAISI") & "-" & w_Rs2("T25_YOTEI_SYURYO")
                    f_ChkDate = False
                    exit Do
                end if
            else
                If CDate(m_iSyuryo) > CDate(w_Rs2("T25_YOTEI_KAISI")) Then
                    if CDate(m_iKaisi) >= CDate(w_Rs2("T25_YOTEI_SYURYO")) Then
                    else
                        m_sMsg = "���łɓo�^����Ă���\�莞�ԂƗ\�莞�Ԃ��d�����Ă��܂�\n" & m_iKaisi & "-" & m_iSyuryo & "<" & w_Rs2("T25_YOTEI_KAISI") & "-" & w_Rs2("T25_YOTEI_SYURYO")
                        f_ChkDate = False
                        exit Do
                    end if
                end if
            end if
	        w_Rs2.MoveNext
        Loop
    

        Exit Do
    Loop

    gf_closeObject(w_Rs2)

End Function

'********************************************************************************
'*  [�@�\]  �o�^�������s��(�����\��)
'*  [����]  �Ȃ�
'*  [�ߒl]  0:�X�V�����A1:�L�[�ᔽ�A99:���s
'*  [����]�@
'********************************************************************************
'Function f_InsertData()
Function f_InsertData(p_sJissiBi,p_RenBan)

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��

    On Error Resume Next
    Err.Clear
    f_InsertData = 1
    
    Do

        '// �w�肳�ꂽ�ް��}��
        w_sSQL = ""
        w_sSQL = w_sSQL & "INSERT INTO T25_KYOKAN_YOTEI"
        w_sSQL = w_sSQL & " ("
        w_sSQL = w_sSQL & "  T25_NENDO"
        w_sSQL = w_sSQL & ", T25_SIKEN_KBN"
        w_sSQL = w_sSQL & ", T25_SIKEN_CD"
        w_sSQL = w_sSQL & ", T25_KYOKAN"
        w_sSQL = w_sSQL & ", T25_YOTEIBI"
        w_sSQL = w_sSQL & ", T25_RENBAN"
        w_sSQL = w_sSQL & ", T25_YOTEI_KAISI"
        w_sSQL = w_sSQL & ", T25_YOTEI_SYURYO"
        w_sSQL = w_sSQL & ", T25_BIKO"
        w_sSQL = w_sSQL & ", T25_INS_DATE"
        w_sSQL = w_sSQL & ", T25_INS_USER"
        w_sSQL = w_sSQL & " ) VALUES ("
        w_sSQL = w_sSQL & m_iSyoriNen
        w_sSQL = w_sSQL & "," & m_iSikenKbn
        w_sSQL = w_sSQL & ", '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & ", '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & ", '" & gf_YYYY_MM_DD(p_sJissiBi,"/") & "'"
        w_sSQL = w_sSQL & "," & p_RenBan
        w_sSQL = w_sSQL & ", '" & m_iKaisi & "'"
        w_sSQL = w_sSQL & ", '" & m_iSyuryo & "'"
        w_sSQL = w_sSQL & ", '" & m_sBiko & "'"
        w_sSQL = w_sSQL & ", '" & gf_YYYY_MM_DD(date(),"/") & "'"
        w_sSQL = w_sSQL & ", '" & Session("LOGIN_ID") & "'"
        w_sSQL = w_sSQL & ")"

        w_iRet = gf_ExecuteSQL(w_sSQL)
        If w_iRet <> 0 Then
            '�}���������s
            If w_iRet = C_ERR_DATA_EXIST or w_iRet = C_ERR_DATA_EXIST2 Then
                m_sMsg = "�o�^�Ɏ��s���܂���"
                Exit Do
            Else
                m_sMsg = "�o�^�G���[�ł�"
                f_InsertData = 99
                Exit Do
            End If
        End If

        '//����I��
        f_InsertData = 0
        Exit Do
    Loop

End function

'********************************************************************************
'*  [�@�\]  �X�V�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:�X�V�����A1:�L�[�ᔽ�A99:���s
'*  [����]  
'********************************************************************************
Function f_UpdateData()

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    
    On Error Resume Next
    Err.Clear

    f_UpdateData = 1
    
    Do 

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

        '// �w�肳�ꂽ�ް��̑��݊m�F
        '// UPDATE���s�����ް������݂��Ȃ��ꍇ�A�G���[���������Ȃ����ߎ��O�Ɋm�F����
        '// �����\��e�[�u��ں��޾�Ă��擾
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & "SELECT T25_NENDO "
        w_sSQL = w_sSQL & vbCrLf & ", T25_SIKEN_KBN "
        w_sSQL = w_sSQL & vbCrLf & ", T25_SIKEN_CD "
        w_sSQL = w_sSQL & vbCrLf & ", T25_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & ", T25_RENBAN "
        w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_YOTEIBI = '" & gf_YYYY_MM_DD(request("txtKeyYoteibi"),"/") & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T25_RENBAN = " & m_iRenban

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            '//۰��ޯ�
            Call gs_RollbackTrans()
            'ں��޾�Ă̎擾���s
            f_UpdateData = 99
            Exit Do
        End If

        If w_Rs.EOF Then
            '//۰��ޯ�
            Call gs_RollbackTrans()
            '�Ώ�ں��ނȂ�
            m_sMsg = "�Ώۃ��R�[�h������܂���"
            'f_UpdateData = 1
            Exit Do
        End If

        '// �w�肳�ꂽ�ް��X�V
        w_sSQL = ""
        w_sSQL = w_sSQL & "UPDATE T25_KYOKAN_YOTEI"
        w_sSQL = w_sSQL & " SET "
        w_sSQL = w_sSQL & " T25_YOTEIBI = '" & gf_YYYY_MM_DD(m_iJissiDate,"/") & "'"    '//�ۗ�
        w_sSQL = w_sSQL & ", T25_YOTEI_KAISI = '" & m_iKaisi & "'"  '//�ۗ�
        w_sSQL = w_sSQL & ", T25_YOTEI_SYURYO = '" & m_iSyuryo & "'" '//�ۗ�
        w_sSQL = w_sSQL & ", T25_BIKO = '" & m_sBiko & "'"
        w_sSQL = w_sSQL & ", T25_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "'"
        w_sSQL = w_sSQL & ", T25_UPD_USER = '" & Session("LOGIN_ID") & "'"
        w_sSQL = w_sSQL & " WHERE T25_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & " AND T25_SIKEN_KBN = " & m_iSikenKbn
        w_sSQL = w_sSQL & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
        w_sSQL = w_sSQL & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & " AND T25_YOTEIBI = '" & gf_YYYY_MM_DD(request("txtKeyYoteibi"),"/") & "'"
        w_sSQL = w_sSQL & " AND T25_RENBAN = " & m_iRenban

        w_iRet = gf_ExecuteSQL(w_sSQL)
        If w_iRet <> 0 Then
            '//۰��ޯ�
            Call gs_RollbackTrans()
            '�X�V�������s
            If w_iRet = C_ERR_DATA_EXIST Or w_iRet = C_ERR_DATA_EXIST2 Then
                m_sMsg = "�X�V�����Ɏ��s���܂���"
                'f_UpdateData = 1
                Exit Do
            Else
                m_sMsg = "�X�V�G���[�ł�"
                f_UpdateData = 99
                Exit Do
            End If
        End If
        
        '//ں��޾��CLOSE
        Call gf_closeObject(w_Rs)
        
        '//�Я�
        Call gs_CommitTrans()
        
        '//����I��
        f_UpdateData = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSikenKbn = Request("txtSikenKbn")            ':�����敪
    
    if Request("txtSikenCd") <> "" Then
        m_iSikenCd = Request("txtSikenCd")      ':�����R�[�h
    else
        m_iSikenCd = 0
    end if

    m_sMode = Request("txtMode")                    ':���샂�[�h
    m_iSyoriNen = Session("NENDO")                  ':�����N�x
    m_iKyokanCd = Session("KYOKAN_CD")              ':�����R�[�h
    m_iJissiDate = Request("cmbJissiDate")			':���t
    m_iRenban = Request("txtRenban")                ':�A��

    if m_sMode = "UPDATE" or m_sMode = "INSERT" Then
'        m_iJissiDate = Request("cmbJissiDate")  ':���t
        m_iJissiDateE = Request("cmbJissiDateE")  ':�I�����t
        m_sBiko = Request("txtBiko")            ':���l
        m_iKaisiH = Request("txtKaisiH")        ':�J�n�����i���j
        m_iKaisiN = Request("txtKaisiN")        ':�J�n�����i���j
        m_iSyuryoH = Request("txtSyuRyoH")      ':�I�������i���j
        m_iSyuryoN = Request("txtSyuryoN")      ':�I�������i���j
    end if

    '// BLANK�̏ꍇ�͍s���ر
    If Request("txtMode") = "BLANK" Then
        m_iPage = 1
    Else
        m_iPage = INT(Request("txtPage"))   ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �����n����Ă��������̒l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParamDate()
    
    if m_sMode = "UPDATE" or m_sMode = "INSERT" Then
    
    Dim w_sKaisiH
    Dim w_sKaisiN
    Dim w_sSyuryoH
    Dim w_sSyuryoN
    
    w_sKaisiH = ""
    w_sKaisiN = ""
    w_sSyuryoH = ""
    w_sSyuryoN = ""
    
    w_sKaisiH = Trim(m_iKaisiH)
    w_sKaisiN = Trim(m_iKaisiN)
    w_sSyuryoH = Trim(m_iSyuryoH)
    w_sSyuryoN = Trim(m_iSyuryoN)
    
    if Len(w_sKaisiH) > 0 Then
         w_sKaisiH = Right("0" & w_sKaisiH,2)
    else
    end if
    
    if Len(w_sKaisiN) > 0 Then
        w_sKaisiN = Right("0" & w_sKaisiN,2)
    else
    end if

    if Len(w_sSyuryoH) > 0 Then
         w_sSyuryoH = Right("0" & w_sSyuryoH,2)
    else
    end if

    if Len(w_sSyuryoN) > 0 Then
        w_sSyuryoN = Right("0" & w_sSyuryoN,2)
    else
    end if
    
    m_iKaisi = w_sKaisiH & ":" & w_sKaisiN
    m_iSyuryo = w_sSyuryoH & ":" & w_sSyuryoN
    
    end if

End Sub

Sub s_MakeKikanWhere()
'********************************************************************************
'*  [�@�\]  �\��R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

'-----2001/07/27 ito M40_CALENDER�폜�̈וύX

'    m_iKikanWhere= ""
'    m_iKikanWhere = m_iKikanWhere & " M40_DATE >= '" & m_iJissiKaisi & "' "
'    m_iKikanWhere = m_iKikanWhere & " AND M40_DATE <= '" & m_iJissiSyuryo & "' "

    m_iKikanWhere= ""
    m_iKikanWhere = m_iKikanWhere & " T32_HIDUKE >= '" & m_iJissiKaisi & "' "
    m_iKikanWhere = m_iKikanWhere & " AND T32_HIDUKE <= '" & m_iJissiSyuryo & "' "
    m_iKikanWhere = m_iKikanWhere & " GROUP BY T32_HIDUKE"

'response.write m_iKikanWhere & "<BR>"

End Sub

'********************************************************************************
'*  [�@�\]  �w�N���Ƃ̎������Ԃ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
Function f_GetSikenKikan()

    Dim rs                '// ں��޾�ĵ�޼ު��
    Dim iRet              '// �߂�l
    Dim w_sSQL              '// SQL��

    On Error Resume Next
    Err.Clear
    f_GetSikenKikan = false

    Do

        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  T24.T24_SIKEN_KBN"
        w_sSql = w_sSql & vbCrLf & "  ,T24.T24_SIKEN_CD"
        w_sSql = w_sSql & vbCrLf & "  ,T24.T24_GAKUNEN"
        w_sSql = w_sSql & vbCrLf & "  ,T24.T24_JISSI_KAISI"
        w_sSql = w_sSql & vbCrLf & "  ,T24.T24_JISSI_SYURYO"
        w_sSql = w_sSql & vbCrLf & " FROM T24_SIKEN_NITTEI T24"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      T24.T24_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND T24.T24_SIKEN_KBN= " & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND T24.T24_SIKEN_CD='" & m_iSikenCd & "'"
        w_sSql = w_sSql & vbCrLf & " ORDER BY T24.T24_GAKUNEN"

        iRet = gf_GetRecordset(rs,w_sSql)
        If iRet <> 0  Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSikenKikan = False
            Exit Do
        End If

		'// �������̎擾
		iRet = f_GetDisp_Data_Siken(w_sSikenName)
        If iRet <> 0  Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSikenKikan = false
            Exit Do
        End If

		'===========================================
		'HTML�����o��
		'===========================================
		%>
		<table class=hyo border=1 width=420>
		<tr><th class="header" colspan="6"><%=w_sSikenName%>����</th></tr>
		<tr>

		<%
		i = 1
		For i = 1 To 5
			If rs.EOF = False Then
				If i=cint(rs("T24_GAKUNEN")) Then%>
					<th class="header" width="33"  align="center"><font size=2><%=i%>�N</font></th>
					<td class="detail" width="100" align="center"><font size=2><%=right(rs("T24_JISSI_KAISI"),5) & "�`" & right(rs("T24_JISSI_SYURYO"),5) %></font></td>
					<%
					rs.MoveNext
				Else%>
					<th class="header" width="33" align="center"><font size=2><%=i%>�N</font></th>
					<td class="detail" width="100" align="center" ><font size=2>�\</font></td>
				<%
				End If

			Else%>
				<th class="header" width="33" align="center"><font size=2><%=i%>�N</font></th>
				<td class="detail" width="100" align="center" ><font size=2>�\</font></td>
				<%
			End If

			If i=3 Then%>
				</tr><tr>
			<%
			End If
		Next
		%>

		<td class="detail" width="100" align="center" colspan="2"></td>
		</tr>
		</table>
		<br>
		<%
		'===========================================

		f_GetSikenKikan = True
        Exit Do
    Loop

    gf_closeObject(rs)

'// LABEL_f_ChkDate_END
End Function

'********************************************************************************
'*  [�@�\]  �w�N���Ƃ̎������Ԃ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sMinDate:�������Ԃ̍ŏ����t
'*          p_sMaxDate:�������Ԃ̍ő���t
'*  [����]  
'********************************************************************************
Function f_GetKikanLimit(p_sMinDate,p_sMaxDate)

    Dim rs                '// ں��޾�ĵ�޼ު��
    Dim iRet              '// �߂�l
    Dim w_sSQL              '// SQL��

    On Error Resume Next
    Err.Clear

    f_GetKikanLimit = True
	p_sMinDate = ""
	p_sMaxDate = ""

    Do

        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  MIN(T24.T24_JISSI_KAISI) AS MIN_JISSI_KAISI"
        w_sSql = w_sSql & vbCrLf & "  ,MAX(T24.T24_JISSI_SYURYO) AS MAX_JISSI_SYURYO"
        w_sSql = w_sSql & vbCrLf & " FROM T24_SIKEN_NITTEI T24"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      T24.T24_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND T24.T24_SIKEN_KBN= " & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND T24.T24_SIKEN_CD='" & m_iSikenCd & "'"
        w_sSql = w_sSql & vbCrLf & " ORDER BY T24.T24_GAKUNEN"

        iRet = gf_GetRecordset(rs,w_sSql)
        If iRet <> 0  Then
            'ں��޾�Ă̎擾���s
            f_GetKikanLimit = False
            Exit Do
        End If

		If rs.EOF = false Then
			p_sMinDate = rs("MIN_JISSI_KAISI")
			p_sMaxDate = rs("MAX_JISSI_SYURYO")
		End If

        Exit Do
    Loop

	'//�I������
    gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �\������(����)���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDisp_Data_Siken(p_sSikenName)
    Dim iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Siken = 1

    Do
        '�����}�X�^���f�[�^���擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_iSikenKbn

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetDisp_Data_Siken = 99
            Exit Do
        End If

        p_sSikenName = ""
        If rs.EOF = False Then
            p_sSikenName = rs("M01_SYOBUNRUIMEI")

            '//���͎����܂��́A�ǎ����I�����ꂽ�ꍇ�����ڍז����ǉ��\��
            If cint(m_sSikenCd) <> 0  Then
                p_sSikenName = p_sSikenName & " (" 
                p_sSikenName = p_sSikenName & rs("M27_SIKENMEI")
                p_sSikenName = p_sSikenName & " )" 
            End If

        End If

        '//����I��
        f_GetDisp_Data_Siken = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************


%>
<html>

<head>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

        <%If m_bBack = True Then%>
           // Ӱ�ނ�BLANK��ݒ肵�A�ꗗ�\�߰�ނɖ߂�
            document.frm.action="default.asp";
            document.frm.target="<%=C_MAIN_FRAME%>";
            document.frm.txtMode.value = "Search";
            document.frm.submit();

        <%Else%>
            // �װ�̏ꍇ�Aү���ނ�\��
            <%If m_bMsgFlg = True Then%>
                window.alert("<%=m_sMsg%>");
            <%End If%>
            // ����̫���
        <%End If%>
    }

    //************************************************************
    //  [�@�\]  �o�^�E�X�V����
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_SaveClick() {
        var w_iRet;
        // ���͒l������
        w_iRet = f_CheckData();

        if( w_iRet == 0 ){
            if( confirm("<%=C_TOUROKU_KAKUNIN%>") == true ){

				//�S���̏ꍇ�A������ϊ�
				if(document.frm.txtDayAll.checked==true){
					document.frm.txtKaisiH.readOnly = true;
					document.frm.txtKaisiH.value = "00";

					document.frm.txtKaisiN.readOnly = true;
					document.frm.txtKaisiN.value = "00";

					document.frm.txtSyuryoH.readOnly = true;
					document.frm.txtSyuryoH.value = "23";

					document.frm.txtSyuryoN.readOnly = true;
					document.frm.txtSyuryoN.value = "55";
				};

            <%If m_sMode = "BLANK" Or m_sMode = "INSERT" Then%>
                // Ӱ�ނ�INSERT��ݒ肵�A�{�߰�ނ��Ă�
                document.frm.action="syousai.asp";
                document.frm.txtMode.value = "INSERT";
                document.frm.submit();
            <%ElseIf m_sMode = "DISP" Or m_sMode = "UPDATE" Then%>
                // Ӱ�ނ�UPDATE��ݒ肵�A�{�߰�ނ��Ă�
                document.frm.action="syousai.asp";
                document.frm.txtMode.value = "UPDATE";
                document.frm.submit();
            <%End If%>
            }
        }
    }

    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //  [�쐬��] 
    //************************************************************
    function f_BackClick(){

        document.frm.action="default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "BLANK";
        document.frm.submit();

    }
    
    //************************************************************
    //  [�@�\]  ���͒l������
    //  [����]  �Ȃ�
    //  [�ߒl]  0:����OK�A1:�����װ
    //  [����]  ���͒l��NULL�����A���������A�����������s��
    //          ���n�ް��p���ް������H����K�v������ꍇ�ɂ͉��H���s��
    //************************************************************
    function f_CheckData() {

        // ���������t����������
        // �� ���{��
        if(f_Trim(document.frm.cmbJissiDate.value) == "" ){
                window.alert("���t�����͂���Ă��܂���");
                document.frm.cmbJissiDate.focus();
                return 1;
		}else{
            if( IsDate(document.frm.cmbJissiDate.value) != 0 ){
                window.alert("���t�̓��͂��s���ł�");
                document.frm.cmbJissiDate.focus();
                return 1;
            }
        }

		<%'//�������������ԓ����ǂ����`�F�b�N������%>
		<%
		'//�������Ԃ̍ŏ����t�ƁA�ő���t���擾����
		w_bRet = f_GetKikanLimit(m_sMinDate,m_sMaxDate)
		If w_bRet = False Then
			m_bErrFlg = True
		End If
		%>

        var MinDate = new Date("<%=m_sMinDate%>");
        var MaxDate = new Date("<%=m_sMaxDate%>");
        var vl = new Date(document.frm.cmbJissiDate.value);
		if( vl< MinDate ){
            window.alert("�������ԊO�̓��t�͓��͂ł��܂���");
            document.frm.cmbJissiDate.focus();
            return 1;
		}else{
			if( vl > MaxDate ){
	            window.alert("�������ԊO�̓��t�͓��͂ł��܂���");
	            document.frm.cmbJissiDate.focus();
	            return 1;
			}
		}

		<%
		'//�V�K�o�^���́A�\��I������ǉ��\������
		If m_sMode = "BLANK" Or m_sMode = "INSERT" Then%>
	        if(f_Trim(document.frm.cmbJissiDateE.value) != "" ){

		        <%'// �������\��I�����̓��t�`�F�b�N������%>
	            if( IsDate(document.frm.cmbJissiDateE.value) != 0 ){
	                window.alert("���t�̓��͂��s���ł�");
	                document.frm.cmbJissiDateE.focus();
	                return 1;
	            }

		        <%'// ���������Ԃ̑召������������%>
		        if( DateParse(document.frm.cmbJissiDate.value,document.frm.cmbJissiDateE.value) < 0){
		            window.alert("�J�n���ƏI�����𐳂������͂��Ă�������");
		            document.frm.cmbJissiDate.focus();
		            return 1;
		        }

				<%'//�������������ԓ����ǂ����`�F�b�N������%>
		        var MinDate = new Date("<%=m_sMinDate%>");
		        var MaxDate = new Date("<%=m_sMaxDate%>");
		        var vl = new Date(document.frm.cmbJissiDateE.value);
				if( vl< MinDate ){
		            window.alert("�������ԊO�̓��t�͓��͂ł��܂���");
		            document.frm.cmbJissiDateE.focus();
		            return 1;
				}else{
					if( vl > MaxDate ){
			            window.alert("�������ԊO�̓��t�͓��͂ł��܂���");
			            document.frm.cmbJissiDateE.focus();
			            return 1;
					}
				}

	        }

		<%End If%>

        // ���������l�̌�����������
       if(f_Trim(document.frm.txtBiko.value) == "" ){
            window.alert("���R�����͂���Ă��܂���");
            document.frm.txtBiko.focus();
            return 1;
		}else{
	        if( getLengthB(document.frm.txtBiko.value) > "200" ){
	            window.alert("���R�̗��͑S�p100�����ȓ��œ��͂��Ă�������");
	            document.frm.txtBiko.focus();
	            return 1;
	        }
		}


<%'==================================================%>
	//�\�肪�I���̎��͓��̓`�F�b�N�͂��Ȃ�
	if (document.frm.txtDayAll.checked==true){
		return 0;
	}
<%'==================================================%>

<%' �������ȉ������`�F�b�N ������%>

        // ������NULL����������
        // ���J�n����
        if( f_Trim(document.frm.txtKaisiH.value) == "" ){
            window.alert("�J�n���������͂���Ă��܂���");
            document.frm.txtKaisiH.focus();
            return 1;
        }
        if( f_Trim(document.frm.txtKaisiN.value) == "" ){
            window.alert("�J�n���������͂���Ă��܂���");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        // ���I������
        if( f_Trim(document.frm.txtSyuryoH.value) == "" ){
            window.alert("�I�����������͂���Ă��܂���");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
        if( f_Trim(document.frm.txtSyuryoN.value) == "" ){
            window.alert("�I�����������͂���Ă��܂���");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        // ���������p������������
        // ���J�n����
        //var str = new String(document.frm.txtKaisiH.value);
        var str = document.frm.txtKaisiH.value;
        if( isNaN(str) ){
            window.alert("�J�n���������p�����ł͂���܂���");
            document.frm.txtKaisiH.focus();
            return 1;
        }
        if( str < 0 ){
            window.alert("�J�n�������s���ł�");
            document.frm.txtKaisiH.focus();
            return 1;
        }

		n = str.match(".");
		if (n == ".") {
			alert("�J�n�������s���ł�"); 
		    document.frm.txtKaisiH.focus();
		    return 1;
		}

        //var str = new String(document.frm.txtKaisiN.value);
        var str = document.frm.txtKaisiN.value;
        if( isNaN(str) ){
            window.alert("�J�n���������p�����ł͂���܂���");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        if( str < 0 ){
            window.alert("�J�n�������s���ł�");
            document.frm.txtKaisiN.focus();
            return 1;
        }

		n = str.match(".");
		if (n == ".") {
			alert("�J�n�������s���ł�"); 
		    document.frm.txtKaisiN.focus();
		    return 1;
		}

        // ���I������
        //var str = new String(document.frm.txtSyuryoH.value);
        var str = document.frm.txtSyuryoH.value;
        if( isNaN(str) ){
            window.alert("�I�����������p�����ł͂���܂���");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
        if( str < 0 ){
            window.alert("�I���������s���ł�");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
		n = str.match(".");
		if (n == ".") {
			alert("�I���������s���ł�"); 
		    document.frm.txtSyuryoH.focus();
		    return 1;
		}

        //var str = new String(document.frm.txtSyuryoN.value);
        var str = document.frm.txtSyuryoN.value;
        if( isNaN(str) ){
            window.alert("�I�����������p�����ł͂���܂���");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        if( str < 0 ){
            window.alert("�I���������s���ł�");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
		n = str.match(".");
		if (n == ".") {
			alert("�I���������s���ł�"); 
		    document.frm.txtSyuryoN.focus();
		    return 1;
		}

        // ������������������
        // ���J�n����
        var str = new String(document.frm.txtKaisiH.value);
        if( str.length > 2 ){
            window.alert("�J�n������2���ȓ��ł͂���܂���");
            document.frm.txtKaisiH.focus();
            return 1;
        }
        var str = new String(document.frm.txtKaisiN.value);
        if( str.length > 2 ){
            window.alert("�J�n������2���ȓ��ł͂���܂���");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        // ���I������
        var str = new String(document.frm.txtSyuryoH.value);
        if( str.length > 2 ){
            window.alert("�I��������2���ȓ��ł͂���܂���");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
        var str = new String(document.frm.txtSyuryoN.value);
        if( str.length > 2 ){
            window.alert("�I��������2���ȓ��ł͂���܂���");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        // ��������������������
        // ���J�n����
<%
'//        if( f_Trim(document.frm.txtKaisiH.value) < 9 ){
'//            window.alert("�\�莞�Ԃ̓��͂́A9�F00����23:55�܂łł�");
'//            document.frm.txtKaisiH.focus();
'//            return 1;
'//        }
%>
        if( f_Trim(document.frm.txtKaisiH.value) >= 24 ){
            window.alert("�\�莞�Ԃ̓��͂́A23:55�܂łł�");
            document.frm.txtKaisiH.focus();
            return 1;
        }

        if( f_Trim(document.frm.txtKaisiN.value) >= 60 ){
            window.alert("�\�莞�Ԃ𐳊m�ɓ��͂��Ă�������");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        if( f_Trim(document.frm.txtKaisiN.value) < 0 ){
            window.alert("�\�莞�Ԃ𐳊m�ɓ��͂��Ă�������");
            document.frm.txtKaisiN.focus();
            return 1;
        }
        // ���I������
<%
'//        if( f_Trim(document.frm.txtSyuryoH.value) < 9 ){
'//            window.alert("�\�莞�Ԃ̓��͂́A9�F00����23:55�܂łł�");
'//            document.frm.txtSyuryoH.focus();
'//            return 1;
'//        }
%>
        if( f_Trim(document.frm.txtSyuryoH.value) >= 24 ){
            window.alert("�\�莞�Ԃ̓��͂́A23:55�܂łł�");
            document.frm.txtSyuryoH.focus();
            return 1;
        }


        if( f_Trim(document.frm.txtSyuryoN.value) >= 60 ){
            window.alert("�\�莞�Ԃ𐳊m�ɓ��͂��Ă�������");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        if( f_Trim(document.frm.txtSyuryoN.value) < 0 ){
            window.alert("�\�莞�Ԃ𐳊m�ɓ��͂��Ă�������");
            document.frm.txtSyuryoN.focus();
            return 1;
        }
        // ���I�������i���j���J�n�����i���j�����ĂȂ���
        if( Number(f_Trim(document.frm.txtKaisiH.value)) > Number(f_Trim(document.frm.txtSyuryoH.value)) ){
            window.alert("�I�������͊J�n�����ȍ~�ɂ��Ă�������");
            document.frm.txtSyuryoH.focus();
            return 1;
        }
        // ���I�������i���j���J�n�����i���j�����ĂȂ���
        if( Number(f_Trim(document.frm.txtKaisiH.value)) == Number(f_Trim(document.frm.txtSyuryoH.value)) ){
            if( Number(f_Trim(document.frm.txtKaisiN.value)) > Number(f_Trim(document.frm.txtSyuryoN.value)) ){
                window.alert("�I�������͊J�n�����ȍ~�ɂ��Ă�������");
                document.frm.txtSyuryoN.focus();
                return 1;
            }
        }
        // ���J�n����
        var str = new String(document.frm.txtKaisiN.value);
        if( str.length < 2 ){
            str = 0 + str;
        }
        if( f_Trim(str).substr(1,1) != 0 ){
            if( f_Trim(str).substr(1,1) != 5 ){
                window.alert("�\�莞�Ԃ�5���P�ʂœ��͂��Ă�������");
                document.frm.txtKaisiN.focus();
                return 1;
            }
        }
        
        // ���I������
        var str = new String(document.frm.txtSyuryoN.value);
        if( str.length < 2 ){
            str = 0 + str;
        }
        if( f_Trim(str).substr(1,1) != 0 ){
            if( f_Trim(str).substr(1,1) != 5 ){
                window.alert("�\�莞�Ԃ�5���P�ʂœ��͂��Ă�������");
                document.frm.txtSyuryoN.focus();
                return 1;
            }
        }
        // ���J�n�����ƏI������������łȂ���
        if( f_Trim(document.frm.txtKaisiH.value) == f_Trim(document.frm.txtSyuryoH.value) ){
            if( f_Trim(document.frm.txtKaisiN.value) == f_Trim(document.frm.txtSyuryoN.value) ){
                window.alert("�J�n�����ƏI������������ł�");
                document.frm.txtSyuryoN.focus();
                return 1;
            }
        }
        
        return 0;
    }
<%
'    //************************************************************
'    //  [�@�\]  �I���`�F�b�N��
'    //  [����]  �Ȃ�
'    //  [�ߒl]  
'    //  [����]  
'    //************************************************************%>
	function f_ZenCheck(obj){
		if(obj.checked==true){
			document.frm.txtKaisiH.value = " -";
			document.frm.txtKaisiH.readOnly = true;

			document.frm.txtKaisiN.value = " -";
			document.frm.txtKaisiN.readOnly = true;

			document.frm.txtSyuryoH.value = " -";
			document.frm.txtSyuryoH.readOnly = true;

			document.frm.txtSyuryoN.value = " -";
			document.frm.txtSyuryoN.readOnly = true;

		}else{
			document.frm.txtKaisiH.value = "";
			document.frm.txtKaisiH.readOnly = false;

			document.frm.txtKaisiN.value = "";
			document.frm.txtKaisiN.readOnly = false;

			document.frm.txtSyuryoH.value = "";
			document.frm.txtSyuryoH.readOnly = false;

			document.frm.txtSyuryoN.value = "";
			document.frm.txtSyuryoN.readOnly = false;

		}
	}

    //-->

</SCRIPT>
<link rel=stylesheet href="../../common/style.css" type=text/css>
</head>

<body LANGUAGE="javascript" onload="return window_onload()">
<form name="frm" Method="POST">
<div align="center">
<%
if m_sMode = "DISP" or m_sMode = "UPDATE" Then
    call gs_title("�����ēƏ��\���o�^","�C�@��")
Else
    call gs_title("�����ēƏ��\���o�^","�V�K�o�^")
End If
%>
<br>
<br>


<%
'//�������ԕ\��
Call f_GetSikenKikan()
%>

<table border="0">
	<tr>
	    <td>

			<table border="0" cellpadding="1" cellspacing="1">
		    <COLGROUP  ALIGN=center>
			    <tr>
			        <td align="center">

			            <table border=1 CLASS="hyo">
					        <tr>
					            <TH nowrap CLASS="header" width="120" align="center">���@�@�t</TH>
					            <TD CLASS="detail"  width="370">
								<input type="text" name="cmbJissiDate" value="<%=m_iJissiDate%>" maxlength="10" size="15">
								<input type="button" class="button" onclick="fcalender('cmbJissiDate')" value="�I��">

							<%If m_sMode = "BLANK" Or m_sMode = "INSERT" Then%>
								�`
								<input type="text" name="cmbJissiDateE" value="<%=m_iJissiDateE%>" maxlength="10" size="15">
								<input type="button" class="button" onclick="fcalender('cmbJissiDateE')" value="�I��">
								<br>
							<%End If%>

								<font size=2>�i���͗�:<%=date()%>�j</font>
								</td>
					        </tr>
					        <tr>
					            <TH nowrap CLASS="header" width="120" align="center">���@�@�R</TH>
					            <TD CLASS="detail" width="280">
							    <textarea rows=4 cols=50 class=text name="txtBiko"><%=m_sBiko%></textarea><BR>
							    <font size=2>�i�S�p100�����ȓ��j</font>
								</TD>
					        </tr>

							<%

							'//���͂���Ă��鎞�Ԃ��I�����ǂ����𔻕�
							'//C_MIN_TIME = "00:00"(�ŏ�����),C_MAX_TIME = "23:55"(�ő厞��)
							If m_iKaisi = C_MIN_TIME And m_iSyuryo = C_MAX_TIME Then
								w_sCheck="checked"
								w_iKaisiH = " -"
								w_iKaisiN = " -"
								w_iSyuryoH = " -"
								w_iSyuryoN = " -"
							Else
								w_sCheck=""
								w_iKaisiH  = m_iKaisiH 
								w_iKaisiN  = m_iKaisiN 
								w_iSyuryoH = m_iSyuryoH
								w_iSyuryoN = m_iSyuryoN

							End If
							%>

					        <tr>
					            <TH nowrap CLASS="header" width="120" align="center">�\��</TH>
					            <TD CLASS="detail" width="280">
								<input type="checkbox" name="txtDayAll" onclick="javascript:f_ZenCheck(this);"  <%=w_sCheck%> >�S��
								</TD>
					        </tr>


					        <tr>
					            <TH nowrap CLASS="header" width="120" align="center">�\�莞��</TH>

								
					            <TD CLASS="detail" width="280">
								<input type="text" name="txtKaisiH" size="2" maxlength="2" value="<%=w_iKaisiH%>">���@
								<input type="text" name="txtKaisiN" size="2" maxlength="2" value="<%=w_iKaisiN%>">���@
								�`
								<input type="text" name="txtSyuryoH" size="2" maxlength="2" value="<%=w_iSyuryoH%>">���@
								<input type="text" name="txtSyuryoN" size="2" maxlength="2" value="<%=w_iSyuryoN%>">���@
								</TD>
					        </tr>
			            </TABLE>

			        </td>
			    </TR>
			</table>

		</td>
	</tr>
    <tr>
		<td align="center">

		    <table>
		        <tr>
		            <td align="center">
		                <%if m_sMode = "DISP" or m_sMode = "UPDATE" Then%>
		                <input type="button" value="�@�X�@�V�@" onClick="javascript:f_SaveClick();return false;" class=button>
		                <%else%>
		                <input type="button" value="�@�o�@�^�@" onClick="javascript:f_SaveClick();return false;" class=button>
		                <%end if%>
		                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value=" �N�@���@�A " class=button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		                <input type="button" value="�L�����Z��" onClick="javascript:f_BackClick();return false;" class=button>
		            </td>
		        </tr>
		    </table>

	    </td>
	</tr>
</table>

</div>


<input type="hidden" name="txtMode" value="<%=m_sMode%>">
<input type="hidden" name="txtSikenKbn" value="<%=m_iSikenKbn%>">
<input type="hidden" name="txtSikenCd" value="<%=m_iSikenCd%>">
<input type="hidden" name="txtRenban" value="<%=m_iRenban%>">
<input type="hidden" name="txtPage" value="<%=m_iPage%>">


<input type="hidden" name="txtKeyYoteibi" value="<%=Request("txtKeyYoteibi")%>">

</form>
</body>

</html>


<%
    '---------- HTML END   ----------
End Sub
%>

