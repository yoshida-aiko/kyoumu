<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: �i�����p�j�����\��o�^
'* ��۸���ID : skn/skn0120/main.asp
'* �@      �\: ���y�[�W �����\��}�X�^�̈ꗗ���X�g�\�����s��
'*-------------------------------------------------------------------------
'* ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*          txtSikenKbn         :�����敪
'*          txtSikenCd          :�����R�[�h
'*          txtMode             :���샂�[�h
'*                              BLANK   :�����\��
'*                              DISP    :�w�肳�ꂽ�敪�̃f�[�^��\��
'*                              CHK     :�w�肳�ꂽ�폜���̃f�[�^��\��
'*                              DEL     :�폜���������s
'*          chkDelRenbanX   :�폜�A�ԁi�������g����󂯎������j
'*          txtPage         :�\���Ő�
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*          txtSikenKbn      :�I�����ꂽ�����敪
'*          chkDelRenbanX   :�폜�A�ԁi�������g�ɓn�������j
'*          txtPage         :�\���Ő�
'* ��      ��:
'*           �������\��
'*               ���������ɂ��Ȃ��������\���\��
'*           ���C���{�^���N���b�N��
'*               �w�肵�������ɂ��Ȃ������\���\�������āA�C��������
'*           ���o�^�{�^���N���b�N��
'*               �����\����͂�\�������āA�o�^������
'*           ���폜�{�^���N���b�N��
'*               �w�肵�������ɂ��Ȃ��������폜����
'*              �{�y�[�W�ɂāA�폜�̏������s��
'*-------------------------------------------------------------------------
'* ��      ��: 2001/06/18 ���u �m��
'* ��      �X: 2001/06/26 ���{
'* ��      �X: 2001/08/03 �ɓ����q �������Ԃ�\������悤�C��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٺݽ� /////////////////////////////
	CONST C_MIN_TIME = "00:00"		'//�ŏ�����
	CONST C_MAX_TIME = "23:55"		'//�ő厞��(����)
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_sMsg              'ү����

    '�擾�����f�[�^�����ϐ�
    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iSikenKbn         ':�����敪
    Public  m_iSikenCd          ':�����R�[�h
    Public  m_sMode             ':���샂�[�h
    Public  m_iRenban           ':�A��
    Public  m_sYoteiMei         ':�\�薼��
    'Public  m_iRiyu            ':���R�i�\��R�[�h�j
    Public  m_dtYoteiKaisi      ':�\��J�n����
    Public  m_dtYoteiSyuryo     ':�\��I������
    Public  m_iMonth            '�\����i���j
    Public  m_iDay              '�\����i���j
    Public  m_sYobi             '�\����i�j���j
    Public  m_iRMax             '�ő�A�Ԓl
    Public  m_iCnt              '�J�E���g����
    Public  m_iPage             '�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_iYoteiBi


    Public  m_Rs                'recordset

    '�y�[�W�֌W
    Public  m_iMax              ':�ő�y�[�W
    Public  m_iDsp              '// �ꗗ�\���s��

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
    w_sWinTitle = "�L�����p�X�A�V�X�g"
    w_sMsgTitle = "�����ēƏ��o�^"
    w_sMsg = ""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    Do
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

        '// ���Ұ�SET
        Call s_SetParam()

'response.write "���[�h = " & Request("txtMode") & "<BR>"

		'===============================
		'//���ԃf�[�^�̎擾
		'===============================
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			'// �y�[�W��\��
			Call showPage_NoData("�����������Ԃł͂���܂���B")
			Exit Do
		End If
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If

        '// �폜����
        If trim(Request("txtMode")) = "DEL" Then
			'//�폜�������s
			w_iRet = s_DelData()
			If w_iRet <> 0 Then
				m_bErrFlg = True
                Exit Do
			End If
        End If

		'�폜�m�F��ʂ�\��
		If Request("txtMode") = "CHK" Then

			'//�폜�I�����ꂽ�f�[�^���擾
			w_iRet = f_GetDeleteData()
			If w_iRet <> 0 Then
				m_bErrFlg = True
                Exit Do
			End If

		Else
	        '�����\��}�X�^���擾
	        w_sWHERE = ""

	        w_sSQL = ""
	        w_sSQL = w_sSQL & "SELECT "
	        w_sSQL = w_sSQL & vbCrLf & " T25_NENDO "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_SIKEN_KBN "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_SIKEN_CD "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_KYOKAN "        
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEIBI "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_RENBAN "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEI_KAISI "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEI_SYURYO "
	        w_sSQL = w_sSQL & vbCrLf & " ,T25_BIKO "
	        w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
	        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
	        w_sSQL = w_sSQL & vbCrLf & " T25_NENDO = " & m_iSyoriNen
	        w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"

	        '���o�����̍쐬
	        If m_iSikenKbn <> "" Then
	           w_sSQL = w_sSQL & " AND T25_SIKEN_KBN = " & m_iSikenKbn
	           w_sSQL = w_sSQL & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
	        End If

	        'w_sSQL = w_sSQL & " ORDER BY T25_YOTEIBI ASC"
	        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T25_YOTEIBI ,T25_YOTEI_KAISI"

	        Set m_Rs = Server.CreateObject("ADODB.Recordset")
	        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
	        If w_iRet <> 0 Then
	            'ں��޾�Ă̎擾���s
	            m_bErrFlg = True
	            m_sErrMsg = "���R�[�h�Z�b�g�̎擾�Ɏ��s���܂���"
	            Exit Do
	        Else
	            '�y�[�W���̎擾
	            m_iMax = gf_PageCount(m_Rs,m_iDsp)
	        End If
		End If

		If m_Rs.EOF Then
            '// �y�[�W��\��
            Call showPage_NoData("�Ώۃf�[�^�͑��݂��܂���B��������͂��Ȃ����Č������Ă��������B")
	        Exit Do
		End If
            '// �y�[�W��\��
            Call showPage()
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// �I������
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  DB����l���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_GetParam()

    Dim w_iDay
    Dim w_iMonth
    Dim w_sYobi
    
    '//���t�Ɨj���\��
    w_iDay = ""
    w_iMonth = ""
    w_iYobi = ""
    w_iDay = f_GetDay(m_Rs("T25_YOTEIBI"))    
    w_iMonth = f_GetMonth(m_Rs("T25_YOTEIBI"))    
    w_sYobi = left(WeekdayName(Weekday(CDate(m_Rs("T25_YOTEIBI")))) ,1)
    m_iMonth = gf_fmtZero(w_iMonth,2)
    m_iDay = gf_fmtZero(w_iDay,2)
    m_sYobi = w_sYobi
    m_sYoteiMei = m_Rs("T25_BIKO")

    '//�����\��
    m_dtYoteiKaisi = m_Rs("T25_YOTEI_KAISI")
    m_dtYoteiSyuryo = m_Rs("T25_YOTEI_SYURYO")

    '//�A�ԕ\��
    m_iRenban = m_Rs("T25_RENBAN")

    m_iYoteiBi = m_Rs("T25_YOTEIBI")

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")         ':�����R�[�h
    m_iSyoriNen = Session("NENDO")             ':�����N�x
    m_iSikenKbn = Request("txtSikenKbn")       ':�����敪

    if Request("txtSikenCd") <> "" Then
        m_iSikenCd = Request("txtSikenCd")      ':�����R�[�h
    else
        m_iSikenCd = 0
    end if

    m_sMode = Request("txtMode")                ':���샂�[�h
    
    m_iRenban = Request("txtRenban")            ':�A��    '//�ۗ�

    '// BLANK�̏ꍇ�͍s���ر
    If Request("txtMode") = "Search" Then
        m_iPage = 1
    Else
        m_iPage = INT(Request("txtPage"))   ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �����n����Ă����l��\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ShowRequest()

    Response.write "txtMode=BLANK" & "&txtSikenKbn=" & m_iSikenKbn

End Sub

'********************************************************************************
'*  [�@�\]  �\������(����)���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDeleteData()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDeleteData = 1

    Do

		'//�I�����ꂽ�����擾
		w_sDelData = replace(Request("chkDel")," ","")
		w_sDelData = split(w_sDelData,",")
		w_iCnt = UBound(w_sDelData)

        '�}�X�^���f�[�^���擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "SELECT "
		w_sSQL = w_sSQL & vbCrLf & " T25_NENDO "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_SIKEN_KBN "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_SIKEN_CD "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_KYOKAN "        
		w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEIBI "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_RENBAN "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_YOTEI_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & " ,T25_BIKO "
		w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
		w_sSQL = w_sSQL & vbCrLf & " WHERE " 
		w_sSQL = w_sSQL & vbCrLf & " T25_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
		w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " AND ("

		For i = 0 To w_iCnt
			If i <> 0 Then
	            w_sSQL = w_sSQL & vbCrLf & " Or "
			End If

			w_Ary = split(w_sDelData(i),"_")
            w_sSQL = w_sSQL & vbCrLf & "  ( T25_YOTEIBI = '" & w_Ary(0) & "'"
            w_sSQL = w_sSQL & vbCrLf & "      AND T25_RENBAN = '" & w_Ary(1) & "'"
            w_sSQL = w_sSQL & vbCrLf & "   )"
		Next

            w_sSQL = w_sSQL & vbCrLf & " )"
	        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T25_YOTEIBI ,T25_YOTEI_KAISI"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetDeleteData = 99
            Exit Do
        End If

        '//����I��
        f_GetDeleteData = 0

        Exit Do
    Loop

End Function

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
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KAISI Is Not Null "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_SYURYO Is Not Null "
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

'
''********************************************************************************
''*  [�@�\]  �w��S���҂��폜����
''*  [����]  �Ȃ�
''*  [�ߒl]  �Ȃ�
''*  [����]  
''********************************************************************************
'Sub s_DelData()
'
'    Dim w_sCD               '// �A��
'    Dim w_iRet              '// �߂�l
'    Dim w_sSQL              '// SQL��
'    
'    m_iCnt = CInt(Request("txtCnt"))
'    
'    For i = 1 to m_iCnt
'        '//��ݻ޸��݊J�n
'        Call gs_BeginTrans()
'
'        If Request("chkDelRenban" & i) <> "" Then
'
'            w_sCD = Request("chkDelRenban" & i)
'            '// �����\��}�X�^ں��޾�Ă��擾
'            w_sSQL = ""
'            w_sSQL = w_sSQL & "DELETE "
'            w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
'            w_sSQL = w_sSQL & vbCrLf & " WHERE " 
'            w_sSQL = w_sSQL & vbCrLf & " T25_NENDO = " & m_iSyoriNen
'            w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
'            w_sSQL = w_sSQL & vbCrLf & " AND T25_RENBAN = " & w_sCD
'            w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
'
'            w_iRet = gf_ExecuteSQL(w_sSQL)
'            If w_iRet <> 0 Then
'                '//۰��ޯ�
'                Call gs_RollbackTrans()
'                
'                'ں��޾�Ă̎擾���s
'                m_bErrFlg = True
'                m_sErrMsg = "�폜�Ɏ��s���܂����B"
'                Exit Sub 'GOTO MAIN
'            End If
'
'
'        End If
'
'    Next
'    
'    '//�Я�
'    Call gs_CommitTrans()
'
'End Sub

'********************************************************************************
'*  [�@�\]  �\����폜����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function  s_DelData()

    Dim w_sCD               '// �A��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��

    On Error Resume Next
    Err.Clear

	s_DelData = 1

	Do

		'//�폜�����擾
		w_sDelData = replace(Request("chkDel")," ","")
		w_sDelData = split(w_sDelData,",")
		w_iCnt = UBound(w_sDelData)

		'//�I�����ꂽ�����\����폜����
		w_sSQL = ""
		w_sSQL = w_sSQL & "DELETE "
		w_sSQL = w_sSQL & vbCrLf & " FROM T25_KYOKAN_YOTEI "
		w_sSQL = w_sSQL & vbCrLf & " WHERE " 
		w_sSQL = w_sSQL & vbCrLf & " T25_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND T25_KYOKAN = '" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_KBN = " & m_iSikenKbn
		w_sSQL = w_sSQL & vbCrLf & " AND T25_SIKEN_CD = '" & m_iSikenCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " AND ("

		For i = 0 To w_iCnt
			If i <> 0 Then
	            w_sSQL = w_sSQL & vbCrLf & " Or "
			End If

			w_Ary = split(w_sDelData(i),"_")
            w_sSQL = w_sSQL & vbCrLf & "  ( T25_YOTEIBI = '" & w_Ary(0) & "'"
            w_sSQL = w_sSQL & vbCrLf & "      AND T25_RENBAN = '" & w_Ary(1) & "'"
            w_sSQL = w_sSQL & vbCrLf & "   )"
		Next
        w_sSQL = w_sSQL & vbCrLf & " )"

		w_iRet = gf_ExecuteSQL(w_sSQL)
		If w_iRet <> 0 Then
		    '�폜���s
			s_DelData = 99
		    m_bErrFlg = True
		    m_sErrMsg = "�폜�Ɏ��s���܂����B"
		    Exit Do
		End If

		'//����I����
		s_DelData = 0
		Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  YYYYMMDD�`���̓��t���猎�𒊏o
'*  [����]  YYYYMMDD�`���̓��t
'*  [�ߒl]  MM�`���̌�
'*  [����]  
'********************************************************************************
Function f_GetMonth(p_sDate)

    f_GetMonth = ""

    If Trim(gf_SetNull2String(p_sDate)) = "" Then
        f_GetMonth = ""
        Exit Function
    End If
    
    f_GetMonth = Month(gf_FormatDate(p_sDate,"/"))

End Function

'********************************************************************************
'*  [�@�\]  YYYYMMDD�`���̓��t������𒊏o
'*  [����]  YYYYMMDD�`���̓��t
'*  [�ߒl]  DD�`���̓�
'*  [����]  
'********************************************************************************
Function f_GetDay(p_sDate)

    f_GetDay = ""

    If Trim(gf_SetNull2String(p_sDate)) = "" Then
        f_GetDay = ""
        Exit Function
    End If
    
    f_GetDay = Day(gf_FormatDate(p_sDate,"/"))

End Function

'********************************************************************************
'*  [�@�\]  �������`
'*  [����]  �����̂ݎ���(hhnn�`���̂��̂̂�)
'*  [�ߒl]  ��؂蕶���t������(�G���[���A���������̂܂�)
'*  [����]  �����݂̂̎�������؂蕶���ŕ�����B
'*  [�ύX]  DB4����5��
'*          hhnn��hh:nn
'********************************************************************************
Function gf_FormatTime(p_Time,p_Delimiter)
    Dim w_sTime 
    Dim w_sHour
    Dim w_sMinute

    '�󔒂Ȃ�G���[
    If IsNull(p_Time)  Then
        gf_FormatTime = p_Time
        Exit Function
    End If
    If p_Time = "" Then 
        gf_FormatTime = p_Time
        Exit Function
    End If

    '�����łȂ��Ȃ�G���[
    If Not IsNumeric(p_Time) Then 
        gf_FormatTime = p_time
        Exit Function
    End If

    '4���łȂ��Ȃ�G���[
    If Len(p_Time) <> 4 Then
        gf_FormatTime = p_Time
        Exit Function
    End If

    w_sHour = Mid(p_Time,1,2)
    w_sMinute  = Mid(p_Time,3,2)

    w_sTime = w_sHour & p_Delimiter 
    w_sTime = w_sTime & w_sMinute

    '�ŏI�I�ɓ��t�łȂ��Ȃ�G���[
    'If Not IsDate(w_sDate) Then    
    '   gf_FormatDate = p_Date
    '   Exit Function
    'End If

    gf_FormatTime = w_sTime

End Function

'********************************************************************************
'*  [�@�\]  �w�N���Ƃ̎������Ԃ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
Function f_GetSikenKikan()

    Dim w_Rs2                '// ں��޾�ĵ�޼ު��
    Dim w_iRet2              '// �߂�l
    Dim w_sSQL2              '// SQL��

    On Error Resume Next
    Err.Clear
    f_GetSikenKikan = True

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

        iRet = gf_GetRecordset(rs,w_sSQL)
        If iRet <> 0  Then
            'ں��޾�Ă̎擾���s
            f_GetSikenKikan = False
            Exit Do
        End If

		'// �������̎擾
		iRet = f_GetDisp_Data_Siken(w_sSikenName)
        If iRet <> 0  Then
            'ں��޾�Ă̎擾���s
            f_GetSikenKikan = False
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

        Exit Do
    Loop

    gf_closeObject(rs)

'// LABEL_f_ChkDate_END
End Function

'********************************************************************************
'*  [�@�\]  �\������(����)���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDisp_Data_Siken(p_sSikenName)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Siken = 1

    Do
        '�����}�X�^���f�[�^���擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI, "
        w_sSql = w_sSql & vbCrLf & "  M27_SIKEN.M27_SIKENMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN ,M27_SIKEN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      M01_KUBUN.M01_SYOBUNRUI_CD = M27_SIKEN.M27_SIKEN_KBN(+)"
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_NENDO = M27_SIKEN.M27_NENDO(+)"
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND M27_SIKEN.M27_SIKEN_CD='" & m_iSikenCd & "'"

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

    Dim w_pageBar           '�y�[�WBAR�\���p
    Dim w_iRecordCnt        '//���R�[�h�Z�b�g�J�E���g
    Dim w_iCnt

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    m_iCnt  = 1

    '�y�[�WBAR�\��
    Call gs_pageBar(m_Rs,m_iPage,m_iDsp,w_pageBar)

%>

<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    //************************************************************
    //  [�@�\]  �o�^��ʂ�\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_NewClick(){
    
        document.frm.action="syousai.asp";
        document.frm.target = "main";
        document.frm.txtMode.value = "BLANK";
        document.frm.submit();
        
    }
    
    //************************************************************
    //  [�@�\]  �I�����ꂽ�������̍X�V��ʂ�\������
    //  [����]  p_sCode     :�I�����ꂽ�A�ԁi�����\��j
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ListClick(p_day,p_sCode){

        document.frm.action="syousai.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtRenban.value = p_sCode;
        //document.frm.cmbJissiDate.value = p_day;

        document.frm.txtKeyYoteibi.value = p_day;

        document.frm.txtMode.value = "DISP";
        document.frm.submit();
        
    }
    
    //************************************************************
    //  [�@�\]  �폜�{�^���������ꂽ�Ƃ��i�m�F�p�j
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ChkDelClick(){
        if( confirm("�\����폜���܂��B") == true ){
            document.frm.action="main.asp";
            document.frm.target = "main";
            document.frm.txtMode.value = "CHK";
            document.frm.submit();
        }
    }
    
    //************************************************************
    //  [�@�\]  �폜�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_DelClick(){
        if( confirm("�\����폜���܂��B") == true ){
            document.frm.action="main.asp";
            document.frm.target = "main";
            document.frm.txtMode.value = "DEL";
            document.frm.submit();
        }
    }

    
    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ��i�m�F�p�j
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //  [�쐬��] 
    //************************************************************
    function f_ChkBackClick(){
    
        document.frm.action="main.asp?<%Call s_ShowRequest()%>";
        document.frm.target="main";
        document.frm.submit();
        
        
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
        document.frm.target="_parent"
        document.frm.submit();
        
        
    }
    
    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="main.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPage.value = p_iPage;
        document.frm.submit();
    
    }
    

	//-->
	</SCRIPT>
	</head>

	<body>
	<center>
	<form name="frm" Method="POST">
	    <input type="hidden" name="txtBiko" value="<%=m_sYoteiMei%>">
	    <input type="hidden" name="txtMode" value="<%=m_sMode%>">
	    <input type="hidden" name="txtRenban">
	    <input type="hidden" name="txtSikenKbn" value="<%=m_iSikenKbn%>">
	    <input type="hidden" name="txtSikenCd" value="<%=m_iSikenCd%>">
	    <input type="hidden" name="txtPage" value="<%= m_iPage %>">
	<table border=0 width="<%=C_TABLE_WIDTH%>">
		<tr>
		<td align="center">
		<%
		if m_sMode = "CHK" Then
		    response.write "�ȉ��̓��e���폜���܂��B<br>" & chr(13)
		else

			'//�������Ԃ�\��
			Call f_GetSikenKikan()

		    response.write w_pageBar
		End if
		%>


		<%Call showTableHead()%>
		<%
		    Do Until m_Rs.EOF
		%>
		<%Call s_GetParam()%>
		<%
		if m_sMode = "CHK" Then
		    Call ShowTableChk()
		else
		    Call ShowTable()
		End if
		%>
		<%m_iCnt = m_iCnt + 1%>
		<%
		            m_Rs.MoveNext

		            If w_iCnt >= C_PAGE_LINE Then
		                Exit Do
		            Else
		                w_iCnt = w_iCnt + 1
		            End If

		    Loop

		'//�폜�m�F��ʂŁA�Ȃ��ꍇ
		if m_sMode <> "CHK" Then
		%>
		    <tr>
		    <td colspan=6 align=right bgcolor=#9999BD><input class=button type=button value="�~�폜" Onclick="f_ChkDelClick()"></td>
		    </tr>
		<%
		Else%>
			<!--�폜�m�F��ʕ\�����l�ێ��p-->
			<input type="hidden" name="chkDel" value="<%=Request("chkDel")%>">
		<%
		End If
		%>
	</table>

	<br>
	<%
	if m_sMode = "CHK" Then
	    response.write "���s���܂����H<br>" & chr(13)
	else
	    response.write w_pageBar
	End if
	%>
	<%Call showButton()%>
	</td>
	</tr>
	</table>

	<input type="hidden" name="txtKeyYoteibi" value="">

	</form>

	</center>
	</body>

	</html>
<%
    '---------- HTML END   ----------
End Sub

Sub showTableHead()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

if m_sMode = "CHK" Then
%>
<table border="1" width="<%=C_TABLE_WIDTH%>" CLASS="hyo">
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="60%" ALIGN=center>
<tr>
    <th CLASS="Header">���t</th>
    <th CLASS="Header">�\�莞��</th>
    <th CLASS="Header">���R</th>
</tr>
<%
else
%>
<table border="1" width="<%=C_TABLE_WIDTH%>" CLASS="hyo">
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="46%" ALIGN=center>
    <COLGROUP WIDTH="6%" ALIGN=center>
    <COLGROUP WIDTH="8%" ALIGN=center>
<tr>
    <th CLASS="header">���t</th>
    <th CLASS="header">�\�莞��</th>
    <th CLASS="header">���R</th>
    <th CLASS="header">�C��</th>
    <th CLASS="header">�폜</th>
</tr>
<%
end if
    '---------- HTML END   ----------
End Sub


Sub showTable()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
<tr>

	<%

	'//���͂���Ă��鎞�Ԃ��I�����ǂ����𔻕�
	'//C_MIN_TIME = "00:00"(�ŏ�����),C_MAX_TIME = "23:55"(�ő厞��)
	If m_dtYoteiKaisi = C_MIN_TIME And m_dtYoteiSyuryo = C_MAX_TIME Then
		w_sStr = "�I��"
	Else
		w_sStr = m_dtYoteiKaisi & "-" & m_dtYoteiSyuryo
	End If
	%>

    <td class=detail>
    <%=m_iMonth%>/<%=m_iDay%>(<%=m_sYobi%>)
    </td>
    <td class=detail>
    <%=w_sStr%>
    </td>
    <td class=detail align="left"><%=m_sYoteiMei%></td>
    <td class=detail align="center"><input type="button" value=">>" onClick="f_ListClick('<%=m_iYoteiBi%>','<%=m_iRenban%>');return false;" class=button></td>
    <td class=detail align="center"><input type="checkbox" name="chkDel" value="<%=m_iYoteiBi & "_" & m_iRenban%>"></td>

</tr>
<%
End Sub

'********************************************************************************
'*  [�@�\]  �m�F��ʂ�\������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub ShowTableChk()

	%>
	<tr>
	    <td class=detail>
	    <%=m_iMonth%>/<%=m_iDay%>(<%=m_sYobi%>)
	    </td>
	    <td class=detail>
	    <%=m_dtYoteiKaisi%>-<%=m_dtYoteiSyuryo%>
	    </td>
	    <td class=detail align="left"><%=m_sYoteiMei%></td>
	    <input type="hidden" name="chkDelRenban<%=m_iCnt%>" value="<%=m_iRenban%>">
	    
	</tr>
	<%

End Sub


Sub showButton()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
if m_sMode = "CHK" Then
%>
<table border="0" width="40%">
<COLGROUP span=2 WIDTH="50%" ALIGN=center>
<tr>
    <td><input type="button" value="�@��@���@" onClick="javascript:f_DelClick();return false;" class=button></td>
    <td><input type="button" value="�L�����Z��" onClick="javascript:f_ChkBackClick();return false;" class=button></td>
</tr>
<input type="hidden" name="txtCnt" value="<%=m_iCnt%>">

</table>
<%
else
end if

End Sub

Sub showPage_NoData(p_msg)
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
    <html>
    <head>
	<link rel=stylesheet href="../../common/style.css" type=text/css>
    </head>
    <body>
    <center>
		<br><br><br>
		<span class="msg"><%=p_msg%></span>
    </center>
    </body>
    </html>

<%
    '---------- HTML END   ----------
End Sub
%>