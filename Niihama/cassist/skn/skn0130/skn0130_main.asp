<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: �������{�Ȗړo�^
'* ��۸���ID : skn/skn0130/main.asp
'* �@      �\: ���y�[�W �������{�Ȗڂ̈ꗗ���X�g�\�����s��
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
'*          txtPageCD         :�\���Ő�
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*          txtSikenKbn      :�I�����ꂽ�����敪
'*          chkDelRenbanX   :�폜�A�ԁi�������g�ɓn�������j
'*          txtPageCD         :�\���Ő�
'* ��      ��:
'*           �������\��
'*               ���������ɂ��Ȃ��������\���\��
'*           ���C���{�^���N���b�N��
'*               �w�肵�������ɂ��Ȃ������\���\�������āA�C��������
'*-------------------------------------------------------------------------
'* ��      ��: 2001/06/18 ���u �m��
'* ��      �X: 2001/06/26 ���{
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_sMsg              'ү����

    '�擾�����f�[�^�����ϐ�
    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iSikenKbn         ':�����敪
    Public  m_iSikenCode        ':��������
    Public  m_sMode             ':���샂�[�h
    Public  m_sGetTable         ':���샂�[�h�AT26�Ƀf�[�^������ꍇ��"26"�AT27�Ƀf�[�^������ꍇ��"27"
    Public  m_iCnt              '�J�E���g����
    Public  m_sPageCD           ':�\���ϕ\���Ő��i�������g����󂯎������j
	Public  m_seisekiF    
    Public  m_Rs                'recordset
	Public  m_sSikenDate		'�������ʃf�[�^
	Public	m_iGakunen			'�w�N

    '�y�[�W�֌W
    Public  m_iMax              ':�ő�y�[�W
    Public  m_iDsp              '// �ꗗ�\���s��
	
	private const C_MAIN_FLG_YES = 1				'���C�������i���C�������t���O�j
	'private const C_SEISEKI_INP_FLG_YES = 1			'���ѓ��͋����i���ѓ��͋����t���O�j
	
	Dim m_AryGakunen()
	
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
    w_sMsgTitle = "�������{�Ȗړo�^"
    w_sMsg = ""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget = ""


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

        if m_sMode <> "" then
			if m_sMode = "no" then
				'// �y�[�W��\��
				Call No_showPage("�����������Ԃł͂���܂���B")
				Exit Do
			Else
				'===============================
				'//���ԃf�[�^�̎擾
				'===============================
		        w_iRet = f_Nyuryokudate()
				
				If w_iRet = 1 Then
					    '// �I������
					    Call gf_closeObject(m_Rs)
					    Call gs_CloseDatabase()
						response.Redirect "default.asp?txtMode=no&txtSikenKbn="&m_iSikenKbn&""
						response.end
					Exit Do
				End If
				
				If w_iRet <> 0 Then 
					m_bErrFlg = True
					Exit Do
				End If
				
				'//�ꗗ�ɕ\������Ȗڂ̊w�N���擾
				If not f_GetGakunen()  then Exit Do
				
				'// �\���p�ް����擾����
				If f_GetKamoku() = False then Exit Do
				
				'// �y�[�W��\��
				Call showPage()
			'Exit Do
			End if
		Else
            '// �󔒃y�[�W��\��
            Call showBrankPage()
        end if
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")          ':�����R�[�h
    m_iSyoriNen = Session("NENDO")              ':�����N�x
    m_iSikenKbn = Request("txtSikenKbn")            ':�����敪
    m_iSikenCode = Request("txtSikenCd")            ':��������
    
    if m_iSikenCode = "" then
        m_iSikenCode = 0 
    end if
    
    m_sMode = Request("txtMode")                ':���샂�[�h
    
    '// BLANK�̏ꍇ�͍s���ر
    If Request("txtMode") = "Search" Then
        m_sPageCD = 1
    Else
        m_sPageCD = INT(Request("txtPageCD"))   ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
	
End Sub

'********************************************************************************
'*  [�@�\]  �N���X�����擾����
'*  [����]  p_iNendo  �F�����N�x
'*          p_iGakuNen�F�w�N
'*  [�ߒl]  f_GetClassMax�F�N���X��
'*  [����]  
'********************************************************************************
Function f_GetClassMax(p_iNendo,p_iGakuNen)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassMax = 0

	Do

		'//�N���X���̎擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  COUNT(M05_CLASSNO) as ClassMax"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M05_NENDO=" & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND M05_GAKUNEN=" & p_iGakuNen

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If
		
		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//�N���X��
			f_GetClassMax = cint(rs("ClassMax"))
		End If

		Exit Do
	Loop

	'//�߂�l���
'	f_GetClassMax = rs("ClassMax")

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

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
	w_Syuryo = "T24_SIKEN_SYURYO"
	w_kyokan = Session("KYOKAN_CD")
	
	if w_kyokan = NULL or w_kyokan = "" then w_kyokan = "@@@"

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MIN(T24_SIKEN_NITTEI.T24_SIKEN_KAISI) as KAISI"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SIKEN_SYURYO) as SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO) as SEI_SYURYO"
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
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KAISI <= '" & w_date & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI." & w_Syuryo & " >= '" & w_date & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_SYURYO >= '" & w_date & "' "
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

			m_sSikenDate = m_DRs("SYURYO") '//okada 2001.12.25

			if m_DRs("SYURYO") > w_date then m_seisekiF = 0 '�������ԓ��̏ꍇ�́A���ѓ��͋����̂݃��[�h����
				
				m_sSikenNm = m_DRs("M01_SYOBUNRUIMEI")
		End If
		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �w�N�E�N���X�E�ȖڃR���{���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKamoku()
	Dim w_sSQL
    Dim w_num
	
    On Error Resume Next
    Err.Clear
    
    f_GetKamoku = False
	w_iCnt = 0
	
	Do 
		m_sGetTable = "T27"
		
		Select Case cint(m_iSikenKbn) '�I�񂾎����ɂ���āA�擾�Ȗڂ̊J�݊��Ԃ�ς���
			Case C_SIKEN_ZEN_TYU
				w_sJiki = C_KAI_ZENKI
			Case C_SIKEN_ZEN_KIM
				w_sJiki = C_KAI_ZENKI
			Case C_SIKEN_KOU_TYU
				w_sJiki = C_KAI_KOUKI
			Case C_SIKEN_KOU_KIM
				w_sJiki = C_KAI_KOUKI
		End Select
		
		w_sGakunenWhere = ""		'//�l���C������Ƃ���Where�Ɏg�p
		w_sSQL = ""
		
		for w_num = 1 to ubound(m_AryGakunen)
			w_sSql = w_sSql & "  SELECT "
			w_sSql = w_sSql & "  T27_GAKUNEN AS GAKUNEN ,"
			w_sSql = w_sSql & "  T27_KAMOKU_CD AS KAMOKU ,"
			w_sSql = w_sSql & "  T15_KAMOKUMEI AS KAMOKUMEI�@"
			w_sSql = w_sSql & "  FROM "
			w_sSql = w_sSql & "  T27_TANTO_KYOKAN ,"
			w_sSql = w_sSql & "  M05_CLASS, "
			w_sSql = w_sSql & "  T15_RISYU "
			w_sSql = w_sSql & "  WHERE "
			w_sSql = w_sSql & "      M05_NENDO = T27_NENDO "
			w_sSql = w_sSql & "  AND M05_GAKUNEN =T27_GAKUNEN "
			w_sSql = w_sSql & "  AND M05_CLASSNO = T27_CLASS "
			w_sSql = w_sSql & "  AND M05_GAKKA_CD = T15_GAKKA_CD "
			w_sSql = w_sSql & "  AND T27_KAMOKU_CD = T15_KAMOKU_CD(+) "
			w_sSql = w_sSql & "  AND T15_NYUNENDO(+) = T27_NENDO - T27_GAKUNEN + 1"
			w_sSql = w_sSql & "  AND T27_NENDO = " & m_iSyoriNen
			w_sSql = w_sSql & "  AND T27_KYOKAN_CD ='" & m_iKyokanCd & "' "
			w_sSql = w_sSql & "  AND T27_MAIN_FLG = " & C_MAIN_FLG_YES
			w_sSql = w_sSql & "  and T27_GAKUNEN = " & m_AryGakunen(w_num-1)
			w_sSql = w_sSql & "  AND (T15_KAISETU" & m_AryGakunen(w_num-1) & " =" & w_sJiki & " OR T15_KAISETU" & m_AryGakunen(w_num-1) & " =" & C_KAI_TUNEN & " )"
			w_sSql = w_sSql & "  AND (T27_KAMOKU_CD Not IN (" & f_SubQuery(m_AryGakunen(w_num-1)) & "))"
			
			w_sSql = w_sSql & "  Union "
			
			w_sGakunenWhere = w_sGakunenWhere & m_AryGakunen(w_num-1)
			if w_num <> ubound(m_AryGakunen) then w_sGakunenWhere = w_sGakunenWhere & ","
			
		next
		
		w_sSQL = w_sSQL & " SELECT DISTINCT "
		w_sSQL = w_sSQL & " 	T27_GAKUNEN AS GAKUNEN,"
		w_sSQL = w_sSQL & " 	T27_KAMOKU_CD AS KAMOKU,"
		w_sSQL = w_sSQL & " 	T16_KAMOKUMEI AS KAMOKUMEI"
		w_sSQL = w_sSQL & " FROM"
		w_sSQL = w_sSQL & " 	T27_TANTO_KYOKAN,"
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 	T27_GAKUNEN in (" & w_sGakunenWhere & ") and "
		w_sSQL = w_sSQL & " 	T27_KAMOKU_CD = T16_KAMOKU_CD(+) and "
		w_sSQL = w_sSQL & " 	T27_NENDO = T16_NENDO(+) and "
		w_sSQL = w_sSQL & " 	T27_NENDO = " & m_iSyoriNen & " and "
		w_sSQL = w_sSQL & " 	T27_KYOKAN_CD ='" & m_iKyokanCd & "' and "
		w_sSql = w_sSql & " 	T27_MAIN_FLG = " & C_MAIN_FLG_YES & " and "
		w_sSQL = w_sSQL & " 	T16_OKIKAE_FLG >= " & C_TIKAN_KAMOKU_SAKI 
		w_sSQL = w_sSQL & " GROUP BY "
		w_sSQL = w_sSQL & " 	T27_NENDO"
		w_sSQL = w_sSQL & " 	,T27_GAKUNEN"
		w_sSQL = w_sSQL & " 	,T27_CLASS"
		w_sSQL = w_sSQL & " 	,T27_KAMOKU_CD"
		w_sSQL = w_sSQL & " 	,T16_KAMOKUMEI"
		
		If gf_GetRecordset(m_Rs, w_sSQL) <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKamoku = 99
            Exit Do
        End If

        f_GetKamoku = True
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �u����ȖڃR�[�h������Ă����޸�ذ
'*  [����]  pGakunen = �w�N�b�c
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_SubQuery(pGakunen)
Dim w_sSubSql

	On Error Resume Next
	Err.Clear

	w_sSubSql = ""
	w_sSubSql = w_sSubSql & " SELECT "
	w_sSubSql = w_sSubSql & " 	  T65_KAMOKU_CD_SAKI "
	w_sSubSql = w_sSubSql & " FROM "
	w_sSubSql = w_sSubSql & " 	  T65_RISYU_SENOKIKAE "
	w_sSubSql = w_sSubSql & " WHERE "
	w_sSubSql = w_sSubSql & " 	  T65_NENDO    = " & m_iSyoriNen
'	w_sSubSql = w_sSubSql & " AND T65_GAKKA_CD = '06' "
	w_sSubSql = w_sSubSql & " AND T65_GAKUNEN  = " & pGakunen

	f_SubQuery = w_sSubSql

End Function


'********************************************************************************
'*  [�@�\]  �������Ԋ��̒��Ƀ��O�C�����[�U�̒S��(����)����Ȗڐ����o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_SikenJWariCnt(p_iCnt)

    Dim w_sSQL,iRet,w_Rs

    On Error Resume Next
    Err.Clear
    
    f_SikenJWariCnt = 1

		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & " T26_GAKUNEN AS GAKUNEN"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_CLASS AS CLASS"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_KAMOKU AS KAMOKU"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE T26_NENDO = " & m_iSyoriNen

If m_iSikenKbn < C_SIKEN_KOU_KIM then '�N�x�������̏ꍇ�́A���ׂĂ��Ώ�
		w_sSQL = w_sSQL & vbCrLf & "    AND T26_SIKEN_KBN =" & m_iSikenKbn
End If

		w_sSQL = w_sSQL & vbCrLf & "    AND T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "'"
		w_sSQL = w_sSQL & vbCrLf & "    AND T26_JISSI_KYOKAN ='" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "  T26_NENDO"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_GAKUNEN"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " ,T26_KAMOKU"
		
'response.write w_sSQL  & "<BR>"
'response.end
        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_SikenJWariCnt = 99
			Exit Function
        End If

		p_iCnt = gf_GetRsCount(w_Rs)

        f_SikenJWariCnt = 0

End Function

'********************************************************************************
'*  [�@�\]  �N���X�����擾����
'*  [����]  p_iNendo  �F�����N�x
'*          p_iGakuNen�F�w�N
'*          p_Kamoku  �F�Ȗ�
'*  [�ߒl]  p_Class	      �F�N���XNO
'* �@�@�@�@ f_GetClassName�F�N���X��
'*  [����]  
'********************************************************************************
Function f_GetClassName(p_iNendo,p_iGakuNen,p_Kamoku,p_iKyositu,p_ClassNo)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassName = ""
	w_sClassName = ""
    p_ClassNo = ""
	w_clsMax = f_GetClassMax(m_iSyoriNen,m_Rs("GAKUNEN"))

	Do

		'//�N���X���̎擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT DISTINCT"
'		w_sSql = w_sSql & vbCrLf & "  M05.M05_CLASSMEI AS CLASSMEI,"
		w_sSql = w_sSql & vbCrLf & "   M05.M05_CLASSRYAKU AS CLASSMEI"
		w_sSql = w_sSql & vbCrLf & "  ,M05.M05_CLASSNO "
		w_sSql = w_sSql & vbCrLf & "  ,T27.T27_CLASS"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS M05 , T27_TANTO_KYOKAN T27"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T27.T27_NENDO=" & p_iNendo
'		w_sSQL = w_sSQL & vbCrLf & "  AND T27.T27_KYOSITU_CD = " & p_iKyositu
		w_sSql = w_sSql & vbCrLf & "  AND T27.T27_GAKUNEN=" & p_iGakuNen
		w_sSql = w_sSql & vbCrLf & "  AND T27.T27_KYOKAN_CD ='" & m_iKyokanCd & "' "
		w_sSql = w_sSql & vbCrLf & "  AND T27.T27_KAMOKU_CD ='" & p_Kamoku & "' "
	    w_sSql = w_sSql & vbCrLf & "  AND T27_MAIN_FLG  = " & C_MAIN_FLG_YES 
		w_sSql = w_sSql & vbCrLf & "  AND M05.M05_NENDO = T27.T27_NENDO"
		w_sSql = w_sSql & vbCrLf & "  AND M05.M05_GAKUNEN = T27.T27_GAKUNEN"
		w_sSql = w_sSql & vbCrLf & "  AND M05.M05_CLASSNO = T27.T27_CLASS "
		w_sSql = w_sSql & vbCrLf & " ORDER BY  "
		'w_sSql = w_sSql & vbCrLf & " M05.M05_CLASSNO "
		w_sSql = w_sSql & vbCrLf & " T27.T27_CLASS "

'response.write w_sSql
		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		i = 0
		Do Until rs.EOF
			w_sClassName = w_sClassName & "," & rs("CLASSMEI")
			p_ClassNo = p_ClassNo & "#" & rs("T27_CLASS")
			
			i = i + 1
			rs.MoveNext
		Loop

		If p_ClassNo <> "" then p_ClassNo = Mid(p_ClassNo,2)
		If w_sClassName <> "" then w_sClassName = Mid(w_sClassName,2)

		If i >= w_clsMax then w_sClassName = "�S"

		Exit Do
	Loop

	'//�߂�l���
	f_GetClassName = w_sClassName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �������Ԋ��f�[�^�����{���Ԃ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSikenJikan(p_iNendo,p_iGakuNen,p_Kamoku,p_sSikenJikan,p_sSJissi,p_sKyositu)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetSikenJikan = False
    w_sSikenJikan = 0

    Do
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  T26_NENDO,"					'�N�x
        w_sSql = w_sSql & vbCrLf & "  T26_SIKEN_KBN,"               '�����敪
        w_sSql = w_sSql & vbCrLf & "  T26_SIKEN_CD,"                '�����R�[�h
        w_sSql = w_sSql & vbCrLf & "  T26_GAKUNEN,"                 '�w�N
        w_sSql = w_sSql & vbCrLf & "  T26_CLASS,"                   '�N���X�m�n
        w_sSql = w_sSql & vbCrLf & "  T26_KAMOKU,"                  '�ȖڃR�[�h
        w_sSql = w_sSql & vbCrLf & "  T26_JISSI_KYOKAN,"            '���{�����R�[�h/���ѓ��͋���
        w_sSql = w_sSql & vbCrLf & "  T26_JISSI_FLG,"               '���{�t���O
        w_sSql = w_sSql & vbCrLf & "  T26_SIKENBI,"                 '���{���t
        w_sSql = w_sSql & vbCrLf & "  T26_MAIN_FLG,"                '���C�������t���O 
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_INP_FLG,"         '���ѓ��͋����t���O 
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN1 ,"        '���ѓ��͋����R�[�h1  
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN2,"         '���ѓ��͋����R�[�h2  
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN3,"         '���ѓ��͋����R�[�h3  
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN4,"         '���ѓ��͋����R�[�h4  
        w_sSql = w_sSql & vbCrLf & "  T26_SEISEKI_KYOKAN5,"         '���ѓ��͋����R�[�h5  
        w_sSql = w_sSql & vbCrLf & "  T26_KANTOKU_KYOKAN,"          '�ē����R�[�h
        w_sSql = w_sSql & vbCrLf & "  T26_KYOSITU,"                 '���{�����R�[�h
        w_sSql = w_sSql & vbCrLf & "  T26_SIKEN_JIKAN,"             '��������
        w_sSql = w_sSql & vbCrLf & "  T26_KAISI_JIKOKU,"            '�J�n����
        w_sSql = w_sSql & vbCrLf & "  T26_SYURYO_JIKOKU,"           '�I������
        w_sSql = w_sSql & vbCrLf & "  T26_KYOKAN_RENMEI "           '�����A��
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  T26_SIKEN_JIKANWARI "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      T26_NENDO = " & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND T26_JISSI_KYOKAN='" & m_iKyokanCd & "'"
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_KBN = " & m_iSikenKbn
        w_sSql = w_sSql & vbCrLf & "  AND T26_SIKEN_CD = '" & m_iSikenCode & "' "
        w_sSql = w_sSql & vbCrLf & "  AND T26_GAKUNEN = " & p_iGakuNen
        w_sSql = w_sSql & vbCrLf & "  AND T26_KAMOKU ='" & p_Kamoku & "' "
'response.write w_ssql & "<br>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		p_sSikenJikan = 0

        If rs.EOF = False Then
            p_sSikenJikan = rs("T26_SIKEN_JIKAN")
            p_sSJissi = rs("T26_JISSI_FLG")
            p_sKyositu = rs("T26_KYOSITU")
'response.write "aa(" & p_sKyositu & ")<br>"

        End If
        Exit Do
    Loop

	f_GetSikenJikan = True
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �\������(����)���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSikenName()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetSikenName = ""
    w_sSikenName = ""

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

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

        If rs.EOF = False Then
            w_sSikenName = rs("M01_SYOBUNRUIMEI")
        End If

        Exit Do
    Loop

	f_GetSikenName = w_sSikenName

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �ꗗ�ɕ\������Ȗڂ̊w�N���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  True�������AFalse�����s
'*  [����]  �N�x�A�����敪�A�܂����{�J�n�������{�I������NULL�łȂ����̂Ō���
'*          �f�[�^�����Ȃ��w�N�́A�\�����Ȃ�
'********************************************************************************
Function f_GetGakunen()
	Dim w_sSQL
	Dim wRs
	Dim wCnt,w_num
	
	On Error Resume Next
	Err.Clear
	
	f_GetGakunen = false
	
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	* "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & " 	T24_SIKEN_NITTEI "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	T24_NENDO = " & m_iSyoriNen & " and "
	w_sSql = w_sSql & " 	T24_SIKEN_KBN = " & m_iSikenKbn & " and "
	w_sSql = w_sSql & " 	T24_JISSI_KAISI is not NULL and "
	w_sSql = w_sSql & " 	T24_JISSI_SYURYO is not NULL "
	
	w_sSql = w_sSql & " order by "
	w_sSql = w_sSql & " 	T24_GAKUNEN "
	
	If gf_GetRecordset(wRs, w_sSQL) <> 0 Then exit function
	
	If wRs.EOF Then
		ReDim m_AryGakunen(1)
		m_AryGakunen(0) = 9
		exit function
	end if
	
	wCnt = gf_GetRsCount(wRs)
	
	ReDim m_AryGakunen(wCnt)
	
	for w_num = 1 to wCnt
		m_AryGakunen(w_num-1) = wRs("T24_GAKUNEN")
		wRs.movenext
	next
	
	f_GetGakunen = true
	
	Call gf_closeObject(wRs)
	
End Function

'********************************************************************************
'*  [�@�\]  �\������(����)���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKyositu(p_lKyositu)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetKyositu = ""
    w_sKyositu = ""

    Do
        '�����}�X�^���f�[�^���擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "   M06_KYOSITUMEI "
        w_sSql = w_sSql & vbCrLf & "  ,M06_RYAKUSYO "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      M06_NENDO = " & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M06_KYOSITU_CD = " & p_lKyositu

'response.write w_sSql

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

        If rs.EOF = False Then
            w_sKyositu = rs("M06_KYOSITUMEI")
        End If

        Exit Do
    Loop

	f_GetKyositu = w_sKyositu

    Call gf_closeObject(rs)

End Function

    '---------- �֐� end ----------

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear

    '---------- HTML START ----------
    Dim w_lJikan	'�������{����
	Dim w_className
	Dim w_class
	Dim w_sSikenJikan
	Dim w_sSJissi
	Dim w_sKyosituCd
	Dim w_sKyositu
%>

<html>

<head>
    <title>�������{�Ȗړo�^</title>

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

    }

    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �C���{�^���������̏���
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_regist(p_Gakunen,p_Class,p_Kamoku) {

        document.frm.txtGakunen.value=p_Gakunen;
        document.frm.txtClass.value=p_Class;
        document.frm.txtKamoku.value=p_Kamoku;
        
        document.frm.action="skn0130_regist.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();
    }

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������̏���
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_Back() {
		location.href = "default.asp"
    }
    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>
	<body>
	    <center>
	    <form name=frm>
		<%call gs_title("�������{�Ȗړo�^","��@��")%>
		<br>
		<table class="hyo" border="1" width="260" height="20">
		    <tr>
		        <th class="header" width="80"  align="center" nowrap>���{����</th>
		        <td class="detail" width="180" align="center" nowrap><%=f_GetSikenName()%></td>
		    </tr>
		</table>
	<br>
	<hr size="1" color="#000000" >
	<input class="button" type="button" onclick="javascript:f_Back();" value="�L�����Z��">
	<br><br>
		<% if m_Rs.eof then %>
			<br><br><br>
			<span class="msg">�Ώۃf�[�^�͑��݂��܂���B��������͂��Ȃ����Č������Ă��������B</span>
		<% Else %>
	<span class=CAUTION>
		�� �C���̏ꍇ�͢>>����N���b�N���Ă��������B<br>
	</span>

	    <table width="80%">
	        <tr>
	            <td align="center">
	                <%
	                    '�y�[�WBAR�\��
	                    Call gs_pageBar(m_Rs,m_sPageCD,m_iDsp,w_pageBar)
	                %>
	                <%=w_pageBar %>

	                <table border="1" width="100%" class="hyo">
	                    <tr>
	                        <th class=header align="center" nowrap><font color="#ffffff">�N���X</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">�Ȗږ���</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">���@�{</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">���@��</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">���{����</font></th>
	                        <th class=header align="center" nowrap><font color="#ffffff">�C�@��</font></th>
	                    </tr>
			<%
				w_i = 1
				
				do until m_Rs.eof or w_i > C_PAGE_LINE
					m_JISSI_FLG = ""
					w_sSikenJikan = 0
					w_sSJissi = 0
					w_sKyosituCd = 0
					w_lJikan = f_GetSikenJikan(m_iSyoriNen,m_Rs("GAKUNEN"),m_Rs("KAMOKU"),w_sSikenJikan,w_sSJissi,w_sKyosituCd)				'�������Ԏ擾
					m_JISSI_FLG = w_sSJissi
					m_sjissi_cls = "detail"
					
					Call gf_GetKubunName(C_SIKEN_KBN,m_JISSI_FLG,Session("NENDO"),m_JISSI_FLG)
					
					if m_JISSI_FLG = "0" Or m_JISSI_FLG = "" then 
						m_JISSI_FLG = "������" '���{�敪�����Ȃ��Ƃ��͖����͂Ƃ݂Ȃ��B
					End IF
					
					if m_JISSI_FLG = "������" then 
						m_sjissi_cls = "JISSHIMI"
					End If
					
					if isnull(w_sKyosituCd) = true Then
						w_sKyositu = "" 'w_sKyosituCd = 'm_Rs("KYOSITU") '//Add 2002.1.23
					Else
						 w_sKyositu = f_GetKyositu(w_sKyosituCd)	
					end if
					
					w_className = f_GetClassName(m_iSyoriNen,m_Rs("GAKUNEN"),m_Rs("KAMOKU"),w_sKyosituCd,w_class)'m_Rs("KYOSITU"),w_class)	'�N���X���擾
										'�������擾
					m_iGakunen = m_Rs("GAKUNEN") '// Add 2001.12.26
					
					%>
					<tr>
						<td class="detail" align="left" nowrap>�@<%=gf_HTMLTableSTR(m_Rs("GAKUNEN"))%> - <%=w_className%></td>
						<td class="detail" align="center" nowrap><%=gf_HTMLTableSTR(m_Rs("KAMOKUMEI"))%></td>
						<td class="<%=m_sjissi_cls%>" align="center" nowrap><%=gf_HTMLTableSTR(m_JISSI_FLG)%></td>
						
						<td class=detail align="center" nowrap>
							<%If isnull(w_sSikenJikan) = True Or Trim(w_sSikenJikan) = "" Then w_sSikenJikan = "0"%>
							<%=gf_HTMLTableSTR(w_sSikenJikan)%>��
						</td>
						
						<td class=detail align="center" nowrap><%=gf_HTMLTableSTR(w_sKyositu)%></td>
						<td class=detail align="center" nowrap>
					<%
						'//2001/12/07 ���C�������łȂ��Ȃ�ڍ׉�ʂɑJ�ڂ����Ȃ�
						If CStr(m_Rs("T27_MAIN_FLG")) = CStr(C_MAIN_KYOKAN_YES) Then
					%>
						<input type="button" class="button" name="Change" value=">>" onclick="f_regist(<%=gf_HTMLTableSTR(m_Rs("GAKUNEN"))%>,'<%=w_class%>','<%=gf_HTMLTableSTR(m_Rs("KAMOKU"))%>')">
					<%
						End If
					%>
						</td>
					</tr>
			<%
					w_i = w_i + 1
					m_Rs.movenext
				loop
			%>
				<%=w_pageBar %>

	                </td>
	            </tr>
	        </table>

		<% End if %>

	        <input type="hidden" name="txtGetTable" value = "<%=m_sGetTable%>">
	        <input type="hidden" name="txtMode" value = "<%=m_sMode%>">
	        <input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
	        <input type="hidden" name="txtSikenKbn" value="<%= m_iSikenKbn %>">
	        <input type="hidden" name="txtSikenCd" value="<%= m_iSikenCode %>">
	        <input type="hidden" name="txtSeisekiFlg" value="<%= m_seisekiF %>">

	        <input type="hidden" name="txtGakunen" value="<%= m_iGakunen %>">
	        <input type="hidden" name="txtClass" value="">
	        <input type="hidden" name="txtKamoku" value="">

	        <input type="hidden" name="txtKikan" value="<%=m_sSikenDate%>">
<!--	        <input type="hidden" name="txtkikanEnd" value=""> -->
			

	    </form>

	    </table>

	    </center>

	</body>

	</html>

<%
    '---------- HTML END   ----------
End Sub

'********************************************************************************
'*  [�@�\]  ��HTML���o��
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
Sub showBrankPage()
%>
<html>
<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
</head>

<body>
<center>
<br><br><br>
<span class="msg"><%=C_BRANK_VIEW_MSG%></span>
</center>
</body>

</html>
<% End Sub 

Sub No_showPage(p_msg)
'********************************************************************************
'*  [�@�\]  ��HTML���o��
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
%>
<html>
<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
</head>

<body>
<center>
<br><br><br>
<span class="msg"><%=p_msg%></span>
</center>
</body>

</html>
<% End Sub %>
