<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0100_bottom.asp
' �@      �\: ���y�[�W ���ѓo�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'				�R���{�{�b�N�X�͋󔒂ŕ\��
'			���\���{�^���N���b�N��
'				���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/26 �O�c �q�j
' ��      �X: 2001/08/21 �ɓ� ���q
' ��      �X: 2001/08/21 �ɓ� ���q �w�b�_���؂藣��
' ��      �X: 2002/03/20 ����		NULL��""�ŕ\������
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bErrMsg           '�װү����

	'�����I��p��Where����
    Public m_iNendo			'�N�x
    Public m_sKyokanCd		'�����R�[�h
    Public m_sSikenKBN		'�����敪
    Public m_sGakuNo		'�w�N
    Public m_sClassNo		'�w��
    Public m_sKamokuCd		'�ȖڃR�[�h
    Public m_sSikenNm		'������
    Public m_sSikenbi		'������
    Public m_sKaisiT		'�������{�J�n����
    Public m_sSyuryoT		'�������{�I������
    Public m_sKamokuNo		'�Ȗږ�
    Public m_sTKyokanCd		'�S���Ȗڂ̋���
	Dim		m_rCnt			'���R�[�h�J�E���g
    Public m_sGakkaCd
    Public m_iSyubetu		'�o���l�W�v���@
    Public m_TUKU_FLG
	Public m_iJigenTani		'//�P�����̒P�ʐ�
    
    Public m_iKaisetu_Kbn
    Public m_iKamoku_Kbn
    Public m_iHissen_Kbn
	Public m_ilevelFlg
	Public	m_Rs
	Public	m_TRs
	Public	m_DRs
	Public	m_SRs
	Public	m_iMax			'�ő�y�[�W

	Dim m_iSouJyugyou		'�����Ǝ���
	DIm m_iJunJyugyou		'�����Ǝ���

	Public	m_iKikan		'���͊��ԃt���O
	Public	m_bKekkaNyuryokuFlg		'���ۓ��͉\�׸�(True:���͉� / False:���͕s��)

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
	w_sMsgTitle="���ѓo�^"
	w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""


'    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

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

'//�f�o�b�O
'Call s_DebugPrint

		'//���ԃf�[�^�̎擾
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			m_iKikan = "NO"
		End If

		'=================
		'//�o�����ۂ̎������擾
		'=================
		'//�Ȗڋ敪(0:������,1:�ݐ�)
        w_iRet = gf_GetKanriInfo(m_iNendo,m_iSyubetu)
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		'=================
		'//���ʊ������擾
		'=================
		'//���ʊ���(0:�ʏ����,1:���ʊ���)
        w_iRet = f_getTUKU(m_iNendo,m_sKamokuCd,m_sGakuNo,m_sClassNo,m_TUKU_FLG)
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		
	    '**********************************************************
	    '�ʏ���ƂƓ��ʊ����ŁA�Ƃ��ė���ꏊ���ς��B
	    '**********************************************************
		If m_TUKU_FLG = C_TUKU_FLG_TUJO then  '�ʏ���Ƃ̏ꍇ
			'=================
			'//�Ȗڏ����擾
			'=================
			'//�Ȗڋ敪(0:��ʉȖ�,1:���Ȗ�)�A�y�сA�K�C�I���敪(1:�K�C,2:�I��)�𒲂ׂ�
			'//���x���ʋ敪(0:��ʉȖ�,1:���x���ʉȖ�)�𒲂ׂ�
	        w_iRet = f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg)
			If w_iRet <> 0 Then 
				m_bErrFlg = True
				Exit Do
			End If

			'===============================
			'//���сA�w���f�[�^�擾
			'===============================
			'//�Ȗڋ敪��C_KAMOKU_SENMON(0:��ʉȖ�)�̏ꍇ�̓N���X�ʂɐ��k��\��
			'//�Ȗڋ敪��C_KAMOKU_SENMON(1:���Ȗ�)�̏ꍇ�͊w�ȕʂɐ��k��\��
	        'w_iRet = f_getdate()
	        w_iRet = f_getdate(m_iKamoku_Kbn)
			If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
			If m_rs.EOF Then
				Call ShowPage_No()
				Exit Do
			End If

			'===============================
			'//���ې��̎擾
			'===============================
			w_iRet = f_GetSyukketu()
			If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		Else 
			'===============================
			'//���сA�w���f�[�^�擾
			'===============================
	        w_iRet = f_getTUKUclass(m_iNendo,m_sKamokuCd,m_sGakuNo,m_sClassNo)
			If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
			If m_rs.EOF Then
				Call ShowPage_No()
				Exit Do
			End If


	    End If
		'===============================
		'//�������ԓ��̎擾
		'===============================
'		w_iRet = f_GetSikenJikan()
'		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

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
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_SetParam()

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSikenKBN	= Cint(request("txtSikenKBN"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sGakkaCd	= request("txtGakkaCd")
	m_iJigenTani = Session("JIKAN_TANI") '�P�����̒P�ʐ�

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub
    response.write "m_iNendo	=" & m_iNendo	 & "<br>"
    response.write "m_sKyokanCd	=" & m_sKyokanCd & "<br>"
    response.write "m_sSikenKBN	=" & m_sSikenKBN & "<br>"
    response.write "m_sGakuNo	=" & m_sGakuNo	 & "<br>"
    response.write "m_sClassNo	=" & m_sClassNo	 & "<br>"
    response.write "m_sKamokuCd	=" & m_sKamokuCd & "<br>"
    response.write "m_sGakkaCd	=" & m_sGakkaCd  & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  ���ې��A�x�������擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSyukketu()

    Dim w_iRet
	Dim w_iSyubetu
	Dim w_bZenkiOnly
	Dim w_sSikenKBN
	Dim w_sTKyokanCd

    On Error Resume Next
    Err.Clear

    f_GetSyukketu = 1

'�f�o�b�O***************
'	Call s_DebugPrint
'***********************
	Do
		'==========================================
		'//�ȖڒS�������̋���CD�̎擾
		'==========================================
        'w_iRet = f_GetTantoKyokan()
        'w_iRet = f_GetTantoKyokan(w_sTKyokanCd)
		w_iRet = f_GetTantoKyokan2(w_sTKyokanCd)
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'==========================================
		'//�Ǘ��}�X�^���A�o�����ۂ̎������擾
		'==========================================
		w_iRet = f_GetKanriInfo(w_iSyubetu)
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'==========================================
		'//�����Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
		'==========================================
		'//�O���݂̂̏ꍇ��T21���O�L���������܂ł̌��ې����擾����
		w_iRet = f_SikenInfo(w_bZenkiOnly)
		If w_iRet<> 0 Then
			Exit Do
		End If 

		If w_bZenkiOnly = True Then
			w_sSikenKBN = C_SIKEN_ZEN_KIM
		Else
			w_sSikenKBN = m_sSikenKBN
		End If

		'==========================================
		'//�Ȗڂɑ΂��錋��,�x���̒l�擾
		'==========================================
		'Call gf_GetSyukketuData(m_SRs,m_sSikenKBN,m_sGakuNo,m_sTKyokanCd,m_sClassNo,m_sKamokuCd,w_skaisibi,w_sSyuryobi,"",w_iSyubetu)
		Call gf_GetSyukketuData(m_SRs,w_sSikenKBN,m_sGakuNo,w_sTKyokanCd,m_sClassNo,m_sKamokuCd,w_skaisibi,w_sSyuryobi,"")
		'if m_SRs.EOF = false then m_SRs.MoveFirst

		'//����I��
	    f_GetSyukketu = 0
		Exit Do
	Loop

End Function 

'********************************************************************************
'*  [�@�\]  �Ǘ��}�X�^���o�����ۂ̎������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sSyubetu = C_K_KEKKA_RUISEKI_SIKEN : ������(=0)
'*  [�ߒl]  p_sSyubetu = C_K_KEKKA_RUISEKI_KEI   �F�ݐ�(=1)
'*  [����]  
'********************************************************************************
Function f_GetKanriInfo(p_iSyubetu)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKanriInfo = 1

    Do 

		'//�Ǘ��}�X�^��茇�ۗݐϏ��敪���擾
		'//���ۗݐϏ��敪(C_K_KEKKA_RUISEKI = 32)
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_SYUBETU"
		w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M00_KANRI.M00_NO=" & C_K_KEKKA_RUISEKI	'���ۗݐϏ��敪(=32)

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKanriInfo = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			'//Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '������
			'//Public Const C_K_KEKKA_RUISEKI_KEI = 1      '�ݐ�
			p_iSyubetu = w_Rs("M00_SYUBETU")

		End If

        f_GetKanriInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_SikenInfo(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_SikenInfo = 1
	p_bZenkiOnly = false

    Do 

		'//�����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSQL = w_sSQL & vbCrLf & " T15_RISYU.T15_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & Cint(m_iNendo)-cint(m_sGakuNo)+1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & Trim(m_sKamokuCd) & "'" 
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAISETU" & m_sGakuNo & "=" & C_KAI_ZENKI

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_SikenInfo = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_bZenkiOnly = True
		End If

        f_SikenInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �R���{�őI�����ꂽ�Ȗڂ̉Ȗڋ敪�y�сA�K�C�I���敪�𒲂ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKamokuInfo(p_iKamoku_Kbn,p_iHissen_Kbn,p_ilevelFlg)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKamokuInfo = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKU_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  ,T15_RISYU.T15_HISSEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  ,T15_RISYU.T15_LEVEL_FLG"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(m_sGakuNo) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & m_sKamokuCd & "' "

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKamokuInfo = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_iKamoku_Kbn = w_Rs("T15_KAMOKU_KBN")
			p_iHissen_Kbn = w_Rs("T15_HISSEN_KBN")
			p_ilevelFlg = w_Rs("T15_LEVEL_FLG")
		End If

        f_GetKamokuInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function


'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_getdate(p_iKamoku_Kbn)
Dim w_iNyuNendo


	On Error Resume Next
	Err.Clear
	f_getdate = 1

	Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakuNo) + 1

		'//�������ʂ̒l���ꗗ��\��
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " A.T16_SEI_TYUKAN_Z AS SEI1,A.T16_SEI_KIMATU_Z AS SEI2,A.T16_SEI_TYUKAN_K AS SEI3,A.T16_SEI_KIMATU_K AS SEI4, "

		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_Z AS SEI,A.T16_KEKA_TYUKAN_Z AS KEKA,A.T16_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_Z AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_TYUKAN_Z as SOUJI, A.T16_JUNJIKAN_TYUKAN_Z as JYUNJI, "
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI,A.T16_KEKA_KIMATU_Z AS KEKA,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T16_CHIKAI_KIMATU_Z AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_KIMATU_Z as SOUJI, A.T16_JUNJIKAN_KIMATU_Z as JYUNJI, "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI,A.T16_KEKA_TYUKAN_K AS KEKA,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_K AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_TYUKAN_K as SOUJI, A.T16_JUNJIKAN_TYUKAN_K as JYUNJI, "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_Z AS SEI_ZT,A.T16_KEKA_TYUKAN_Z AS KEKA_ZT,A.T16_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,A.T16_CHIKAI_TYUKAN_Z AS CHIKAI_ZT,A.T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI_ZT, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI_ZK,A.T16_KEKA_KIMATU_Z AS KEKA_ZK,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,A.T16_CHIKAI_KIMATU_Z AS CHIKAI_ZK,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI_ZK, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI_KT,A.T16_KEKA_TYUKAN_K AS KEKA_KT,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,A.T16_CHIKAI_TYUKAN_K AS CHIKAI_KT,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI_KT, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI_KK,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
				w_sSQL = w_sSQL & "		A.T16_SOJIKAN_KIMATU_K as SOUJI, A.T16_JUNJIKAN_KIMATU_K as JYUNJI, A.T16_SAITEI_JIKAN, A.T16_KYUSAITEI_JIKAN, "
		End Select

		w_sSQL = w_sSQL & " 	A.T16_GAKUSEI_NO AS GAKUSEI_NO,A.T16_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_SELECT_FLG"
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_LEVEL_KYOUKAN"
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_OKIKAE_FLG"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T16_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(m_sGakuNo) & " "

		'//�Ȗڋ敪��C_KAMOKU_SENMON(1:���Ȗ�)�̏ꍇ�͊w�ȕʂɐ��k��\��
		'If cint(m_iLevelFlg) = 1 then
		'		w_sSQL = w_sSQL & vbCrLf & " AND	A.T16_LEVEL_KYOUKAN = '" & m_sKyokanCd & "' "
		'		w_sSQL = w_sSQL & vbCrLf & " AND	C.T13_GAKKA_CD = '" & m_sGakkaCd & "' "
		'Else
			If cint(p_iKamoku_Kbn) = cint(C_KAMOKU_SENMON) Then
				'w_sSQL = w_sSQL & vbCrLf & " AND	C.T13_GAKKA_CD = '" & m_sGakkaCd & "' "
				w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
			Else
				w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
			End If
		'End If
		w_sSQL = w_sSQL & " AND	A.T16_NENDO = C.T13_NENDO "

		'//�u�����̐��k�͂͂���(C_TIKAN_KAMOKU_MOTO = 1    '�u����)
		w_sSQL = w_sSQL & " AND	A.T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
'		w_sSQL = w_sSQL & " AND	B.T11_NYUNENDO = " & w_iNyuNendo & " "
		w_sSQL = w_sSQL & " ORDER BY A.T16_GAKUSEKI_NO "

'Response.Write w_ssql
'Response.end

		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_getdate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
		m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))

		'//ں��ރJ�E���g�擾
		m_rCnt=gf_GetRsCount(m_Rs)

		f_getdate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	���ʊ�����u�w���擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_getTUKUclass(p_iNendo,p_sKamokuCd,p_iGakunen,p_iClass)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet
    Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear
    
    f_getTUKUclass = 1
	p_sTKyokanCd = ""

	Do

        w_iNyuNendo = Cint(p_iNendo) - Cint(p_iGakunen) + 1

		'//�������ʂ̒l���ꗗ��\��
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "

		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_Z AS KEKA,A.T34_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,A.T34_CHIKAI_TYUKAN_Z AS CHIKAI, "
				w_sSQL = w_sSQL & "		A.T34_SOJIKAN_TYUKAN_Z as SOUJI, A.T34_JUNJIKAN_TYUKAN_Z as JYUNJI, "
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_Z AS KEKA,A.T34_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T34_CHIKAI_KIMATU_Z AS CHIKAI, "
				w_sSQL = w_sSQL & "		A.T34_SOJIKAN_KIMATU_Z as SOUJI, A.T34_JUNJIKAN_KIMATU_Z as JYUNJI, "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_K AS KEKA,A.T34_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T34_CHIKAI_TYUKAN_K AS CHIKAI, "
				w_sSQL = w_sSQL & "		A.T34_SOJIKAN_TYUKAN_K as SOUJI, A.T34_JUNJIKAN_TYUKAN_K as JYUNJI, "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_Z AS KEKA_ZT,A.T34_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,A.T34_CHIKAI_TYUKAN_Z AS CHIKAI_ZT, "
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_Z AS KEKA_ZK,A.T34_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,A.T34_CHIKAI_KIMATU_Z AS CHIKAI_ZK, "
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_K AS KEKA_KT,A.T34_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,A.T34_CHIKAI_TYUKAN_K AS CHIKAI_KT, "
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_K AS KEKA,A.T34_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T34_CHIKAI_KIMATU_K AS CHIKAI, "
				w_sSQL = w_sSQL & "		A.T34_SOJIKAN_KIMATU_K as SOUJI, A.T34_JUNJIKAN_KIMATU_K as JYUNJI, A.T34_SAITEI_JIKAN, A.T34_KYUSAITEI_JIKAN, "
		End Select

		w_sSQL = w_sSQL & " 	A.T34_GAKUSEI_NO AS GAKUSEI_NO,A.T34_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T34_RISYU_TOKU A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T34_NENDO = " & Cint(p_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T34_TOKUKATU_CD = '" & p_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T34_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T34_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(p_iGakunen) & " "
		w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(p_iClass) & " "

		w_sSQL = w_sSQL & " AND	A.T34_NENDO = C.T13_NENDO "

		w_sSQL = w_sSQL & " ORDER BY A.T34_GAKUSEKI_NO "

'response.write w_sSQL &"<<br>"
'Response.End

		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_getTUKUclass = 99
			m_bErrFlg = True
			Exit Do 
		End If

		m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
		m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))

		'//ں��ރJ�E���g�擾
		m_rCnt=gf_GetRsCount(m_Rs)

		f_getTUKUclass = 0
		Exit Do
	Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[�@�\]	�ȖڒS�������̋���CD�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetTantoKyokan(p_sTKyokanCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetTantoKyokan = 1
	p_sTKyokanCd = ""

    Do 
		'//�ȖڒS�������̋���CD�̎擾
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & "  T20_KYOKAN "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & "  T20_JIKANWARI "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & "  T20_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND T20_KAMOKU = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND T20_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND T20_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " GROUP BY T20_KYOKAN "

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetTantoKyokan = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			'm_sTKyokanCd = w_Rs("T20_KYOKAN")
			p_sTKyokanCd = w_Rs("T20_KYOKAN")
		End If

        f_GetTantoKyokan = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[�@�\]	�ȖڒS�������̋���CD�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetTantoKyokan2(p_sTKyokanCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetTantoKyokan = 1
	p_sTKyokanCd = ""

    Do 
		'//�ȖڒS�������̋���CD�̎擾
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & "  T27_KYOKAN_CD "
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & "  T27_TANTO_KYOKAN "
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & "  T27_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND T27_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND T27_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND T27_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " GROUP BY T27_KYOKAN_CD "

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetTantoKyokan = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then

			p_sTKyokanCd = w_Rs("T27_KYOKAN_CD")
		End If

        f_GetTantoKyokan = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_Nyuryokudate()

	Dim w_sSysDate

	On Error Resume Next
	Err.Clear

	f_Nyuryokudate = 1
	m_bKekkaNyuryokuFlg = False		'���ۓ����׸�

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI "
		w_sSQL = w_sSQL & vbCrLf & "  ,SYSDATE "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakuNo)
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
'//���ѓ��͊��ԃe�X�g�p

'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '2002/04/30'"
'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '2000/03/01'"

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_

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
			m_sSikenNm = gf_SetNull2String(m_DRs("M01_SYOBUNRUIMEI"))		'��������
			m_iNKaishi = gf_SetNull2String(m_DRs("T24_SEISEKI_KAISI"))		'���ѓ��͊J�n��
			m_iNSyuryo = gf_SetNull2String(m_DRs("T24_SEISEKI_SYURYO"))		'���ѓ��͏I����
			m_iKekkaKaishi = gf_SetNull2String(m_DRs("T24_KEKKA_KAISI"))	'���ۓ��͊J�n
			m_iKekkaSyuryo = gf_SetNull2String(m_DRs("T24_KEKKA_SYURYO"))	'���ۓ��͏I��
			w_sSysDate = Left(gf_SetNull2String(m_DRs("SYSDATE")),10)		'�V�X�e�����t
		End If

		'���͊��ԓ��Ȃ琳��
		If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			f_Nyuryokudate = 0
		End If

		'���ۓ��͉\�׸�
		If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			m_bKekkaNyuryokuFlg = True
		End If

		Exit Do
	Loop

End Function


'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_getTUKU(p_iNendo,p_sKamoku,p_iGakunen,p_iClass,p_TUKU_FLG)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

	On Error Resume Next
	Err.Clear
	f_getTUKU = 0
	p_TUKU_FLG = C_TUKU_FLG_TUJO

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T20_TUKU_FLG "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T20_JIKANWARI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T20_NENDO=" & Cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_KAMOKU ='" & p_sKamoku & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_GAKUNEN =" & Cint(p_iGakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_CLASS =" & Cint(p_iClass)

'response.write w_sSQL & "<<<BR>"

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_getTUKU = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If w_Rs.EOF = false Then
			p_TUKU_FLG = cStr(gf_SetNull2Zero(w_Rs("T20_TUKU_FLG")))
		End If

		Exit Do
	Loop
    Call gf_closeObject(w_Rs)

End Function


'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'********************************************************************************
Function f_Syukketu2(p_gaku,p_kbn)

	Dim w_GAKUSEI_NO
	Dim w_SYUKKETU_KBN

	f_Syukketu2 = 0
	w_GAKUSEI_NO = ""
	w_SYUKKETU_KBN = ""
	w_SKAISU = ""

'	On Error Resume Next
'	Err.Clear

	If m_SRs.EOF Then
		Exit Function
	Else
'		m_SRs.MoveFirst
		Do Until m_SRs.EOF

		w_GAKUSEI_NO = m_SRs("T21_GAKUSEKI_NO")
		w_SYUKKETU_KBN = m_SRs("T21_SYUKKETU_KBN")
		w_SKAISU = m_SRs("KAISU")

			If Cstr(w_GAKUSEI_NO) = Cstr(p_gaku) AND cstr(w_SYUKKETU_KBN) = cstr(p_kbn) Then
				f_Syukketu2 = w_SKAISU

				Exit Do
			End If
			m_SRs.MoveNext
		Loop
'set m_SRs=nothing
		m_SRs.MoveFirst
	End If

End Function


'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'********************************************************************************
Function f_Syukketu2New(p_gaku,p_kbn)

	Dim w_GAKUSEI_NO
	Dim w_SYUKKETU_KBN

	f_Syukketu2New = ""
	w_GAKUSEI_NO = ""
	w_SYUKKETU_KBN = ""
	w_SKAISU = ""

	If m_SRs.EOF Then
		Exit Function
	Else
		Do Until m_SRs.EOF

			w_GAKUSEI_NO = m_SRs("T21_GAKUSEKI_NO")
			w_SYUKKETU_KBN = m_SRs("T21_SYUKKETU_KBN")
			w_SKAISU = gf_SetNull2String(m_SRs("KAISU"))

			If Cstr(w_GAKUSEI_NO) = Cstr(p_gaku) AND cstr(w_SYUKKETU_KBN) = cstr(p_kbn) Then
				f_Syukketu2New = w_SKAISU
				Exit Do
			End If

			m_SRs.MoveNext
		Loop

		m_SRs.MoveFirst
	End If

End Function


'********************************************************************************
'*  [�@�\]  �m�茇�ې��A�x�������擾�B
'*  [����]  p_iNendo�@ �@�F�@�����N�x
'*          p_iSikenKBN�@�F�@�����敪
'*          p_sKamokuCD�@�F�@�ȖڃR�[�h
'*          p_sGakusei �@�F�@�T�N�Ԕԍ�
'*  [�ߒl]  p_iKekka   �@�F�@���ې�
'*          p_ichikoku �@�F�@�x����
'*          0�F����C��
'*  [����]  �����敪�ɓ����Ă���A���ې��A�x�������擾����B
'*			2002.03.20
'*			NULL��0�ɕϊ����Ȃ����߂ɁA�֐������W���[�����ō쐬�iCACommon.asp����R�s�[�j
'********************************************************************************
Function f_GetKekaChi(p_iNendo,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku,p_iKekkaGai)

    Dim w_sSQL
    Dim w_KekaChiRs
    Dim w_iRet
    Dim w_sKek,p_sChi
	Dim w_sSouG,w_sJyunG

    On Error Resume Next
    Err.Clear
    
    'p_iKekka = 0
    'p_iChikoku = 0
    p_iKekka = ""
    p_iChikoku = ""


    f_GetKekaChi = 1

	'/ �����敪�ɂ���Ď���Ă���A�t�B�[���h��ς���B
	Select Case p_iSikenKBN
		Case C_SIKEN_ZEN_TYU
			w_sKek = "T16_KEKA_TYUKAN_Z"
			w_sKekG = "T16_KEKA_NASI_TYUKAN_Z"
			p_sChi = "T16_CHIKAI_TYUKAN_Z"
			w_sSouG  = "T16_SOJIKAN_TYUKAN_Z"
			w_sJyunG = "T16_JUNJIKAN_TYUKAN_Z"
		Case C_SIKEN_ZEN_KIM
			w_sKek = "T16_KEKA_KIMATU_Z"
			w_sKekG = "T16_KEKA_NASI_KIMATU_Z"
			p_sChi = "T16_CHIKAI_KIMATU_Z"
			w_sSouG  = "T16_SOJIKAN_KIMATU_Z"
			w_sJyunG = "T16_JUNJIKAN_KIMATU_Z"
		Case C_SIKEN_KOU_TYU
			w_sKek = "T16_KEKA_TYUKAN_K"
			w_sKekG = "T16_KEKA_NASI_TYUKAN_K"
			p_sChi = "T16_CHIKAI_TYUKAN_K"
			w_sSouG  = "T16_SOJIKAN_TYUKAN_K"
			w_sJyunG = "T16_JUNJIKAN_TYUKAN_K"
		Case C_SIKEN_KOU_KIM
			w_sKek = "T16_KEKA_KIMATU_K"
			w_sKekG = "T16_KEKA_NASI_KIMATU_K"
			p_sChi = "T16_CHIKAI_KIMATU_K"
			w_sSouG  = "T16_SOJIKAN_KIMATU_K"
			w_sJyunG = "T16_JUNJIKAN_KIMATU_K"
	End Select

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " " & w_sKek & " as KEKA, "
		w_sSQL = w_sSQL & vbCrLf & " " & w_sKekG & " as KEKA_NASI, "
		w_sSQL = w_sSQL & vbCrLf & " " & p_sChi & " as CHIKAI, "
		w_sSQL = w_sSQL & vbCrLf & " " & w_sSouG & " as SOUJI, "
		w_sSQL = w_sSQL & vbCrLf & " " & w_sJyunG & " as JYUNJI "
		w_sSQL = w_sSQL & vbCrLf & " FROM T16_RISYU_KOJIN "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T16_NENDO =" & p_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND T16_GAKUSEI_NO= '" & p_sGakusei & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T16_KAMOKU_CD= '" & p_sKamokuCD & "'"

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_KekaChiRs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKekaChi = 99
            Exit Do
        End If

		'//�߂�l���
		If w_KekaChiRs.EOF = False Then
			'p_iKekka = w_KekaChiRs("KEKA")
			'p_iKekkaGai = w_KekaChiRs("KEKA_NASI")
			'p_iChikoku = w_KekaChiRs("CHIKAI")
			p_iKekka = gf_SetNull2String(w_KekaChiRs("KEKA"))
			p_iKekkaGai = gf_SetNull2String(w_KekaChiRs("KEKA_NASI"))
			p_iChikoku = gf_SetNull2String(w_KekaChiRs("CHIKAI"))

			m_iSouJyugyou = gf_SetNull2String(w_KekaChiRs("SOUJI"))
			m_iJunJyugyou = gf_SetNull2String(w_KekaChiRs("JYUNJI"))
		End If

        f_GetKekaChi = 0
        Exit Do
    Loop

    Call gf_closeObject(w_KekaChiRs)

End Function


Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_sGakusekiCd
Dim w_sSeiseki
Dim w_sHyoka
Dim w_sKekka,w_sKekkaGai
Dim w_sChikai
Dim w_sKekkasu
Dim w_sChikaisu
Dim w_sShikenKBN_RUI
Dim w_iKekka_rui,w_iChikoku_rui

Dim w_ihalf
Dim i

Dim w_lSeiTotal	'���э��v
Dim w_lGakTotal	'�w���l��

'�f�[�^��NULL�̏ꍇ��0�ɕϊ����Ȃ����߂ɁA��U�f�[�^��ۑ����郏�[�N�Ŏg�p
'2002.03.20
Dim w_sData 	
dim w_sData2


w_lSeiTotal = 0
w_lGakTotal = 0

i = 1

%>
<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

		//�X�N���[����������
		parent.init();

		//���э��v�l�̎擾
		f_GetTotalAvg();

		// �����ԂƏ����Ԃ�hidden�ɃZ�b�g
		document.frm.hidSouJyugyou.value = "<%= m_iSouJyugyou %>";
		document.frm.hidJunJyugyou.value = "<%= m_iJunJyugyou %>";

        //submit
        document.frm.target = "topFrame";
        document.frm.action = "sei0100_middle.asp"
        document.frm.submit();
        return;

    }

   //************************************************************
    //  [�@�\]  �]���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_change(p_iS){

        w_sButton = eval("document.frm.button"+p_iS);
        w_sHyouka = eval("document.frm.Hyoka"+p_iS);

<%If m_sSikenKBN = C_SIKEN_ZEN_TYU Then%>

        if(w_sButton.value == "�E") {
			w_sButton.value = "��";
			w_sHyouka.value = "��";
	        return;
		}
        if(w_sButton.value == "��") {
			w_sButton.value = "�E";
			w_sHyouka.value = "";
	        return;
		}

<%Else%>

        if(w_sButton.value == "�E") {
			w_sButton.value = "��";
			w_sHyouka.value = "��";
	        return;
		}
        if(w_sButton.value == "��") {
			w_sButton.value = "��";
			w_sHyouka.value = "��";
	        return;
		}
        if(w_sButton.value == "��") {
			w_sButton.value = "�E";
			w_sHyouka.value = "";
	        return;
		}
<%End If%>

    }
    
   //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){

		if(f_CheckData_All() == 1){
	            alert("���͒l���s���ł�");
	            return 1;
		}else{
			if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return;}

			document.frm.hidSouJyugyou.value = parent.topFrame.document.frm.txtSouJyugyou.value;
			document.frm.hidJunJyugyou.value = parent.topFrame.document.frm.txtJunJyugyou.value;

			//�w�b�_���󔒕\��
			parent.topFrame.document.location.href="white.asp"

			//�o�^����
		<% if m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
			document.frm.hidUpdMode.value = "TUJO";
			document.frm.action="sei0100_upd.asp";
		<% Else %>
			document.frm.hidUpdMode.value = "TOKU";
			document.frm.action="sei0100_upd_toku.asp";
		<% End if %>
	        document.frm.target="main";
	        document.frm.submit();
		}
	}

	//************************************************************
	//	[�@�\]	�L�����Z���{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//************************************************************
	function f_Cansel(){

        //�����y�[�W��\��
        parent.document.location.href="default.asp"

	}

    //************************************************************
    //  [�@�\]  ���͒l������(�o�^�{�^��������)
    //  [����]  �Ȃ�
    //  [�ߒl]  0:����OK�A1:�����װ
    //  [����]  ���͒l��NULL�����A�p���������A�����������s��
    //          ���n�ް��p���ް������H����K�v������ꍇ�ɂ͉��H���s��
    //************************************************************
    function f_CheckData_All() {
		var i
		var w_Seiseki
		var w_bFLG

		// �����ԁE�����ԓ��̓`�F�b�N
		if(!f_CheckNum("parent.topFrame.document.frm.txtSouJyugyou")){ return 1; }
		if(!f_CheckNum("parent.topFrame.document.frm.txtJunJyugyou")){ return 1; }
		if(!f_CheckDaisyou()){ return 1; }

	<% if m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
		for (i = 1; i < document.frm.i_Max.value; i++) {
			
			w_Seiseki = eval("document.frm.Seiseki"+i);
			w_bFLG = true

			if (w_Seiseki){		//2001/12/17 Add
				if (isNaN(w_Seiseki.value)){
					w_bFLG = false;
					w_Seiseki.focus();
					return 1;
					break;
				}else{
					//����l���`�F�b�N 2001/12/09 �ǉ� �ɓ�
					//var wStr = new String(w_Seiseki.value)
					if (w_Seiseki.value > 100){
						w_bFLG = false;
						w_Seiseki.focus();
						return 1;
						break;
					};

					//�}�C�i�X���`�F�b�N
					var wStr = new String(w_Seiseki.value)
					if (wStr.match("-")!=null){
						w_bFLG = false;
						w_Seiseki.focus();
						return 1;
						break;
					};

					//�����_�`�F�b�N
					w_decimal = new Array();
					w_decimal = wStr.split(".")
					if(w_decimal.length>1){
						w_bFLG = false;
						w_Seiseki.focus();
						return 1;
						break;
					};
				};
			};
		};
		if (w_bFLG == false){
			return 1;
		};
	<% End if %>

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_Chikai = eval("document.frm.Chikai"+i);
			w_bFLG = true
if (w_Chikai){		//2001/12/17 Add
			if (isNaN(w_Chikai.value)){
				w_bFLG = false;
				w_Chikai.focus();
				return 1;
				break;
			}else{

				//�}�C�i�X���`�F�b�N
				var wStr = new String(w_Chikai.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					w_Chikai.focus();
					return 1;
					break;
				};

				//�����_�`�F�b�N
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					w_Chikai.focus();
					return 1;
					break;
				}

			};
};
		};

			if (w_bFLG == false){
				return 1;
			};

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {

			w_Kekka = eval("document.frm.Kekka"+i);
			w_bFLG = true
if (w_Kekka){		//2001/12/17 Add
			if (isNaN(w_Kekka.value)){
				w_bFLG = false;
				w_Kekka.focus();
				return 1;
				break;
			}else{

				//�}�C�i�X���`�F�b�N
				var wStr = new String(w_Kekka.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					w_Kekka.focus();
					return 1;
					break;
				};

				//�����_�`�F�b�N
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					w_Kekka.focus();
					return 1;
					break;
				}

			};
};
		};
			if (w_bFLG == false){
				return 1;
			};

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_KekkaGai = eval("document.frm.KekkaGai"+i);
			w_bFLG = true
if (w_KekkaGai){		//2001/12/17 Add
			if (isNaN(w_KekkaGai.value)){
				w_bFLG = false;
				w_KekkaGai.focus();
				return 1;
				break;
			}else{

				//�}�C�i�X���`�F�b�N
				var wStr = new String(w_KekkaGai.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					w_KekkaGai.focus();
					return 1;
					break;
				};

				//�����_�`�F�b�N
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					w_KekkaGai.focus();
					return 1;
					break;
				}

			};
};
		};
			if (w_bFLG == false){
				return 1;
			};

		return 0;
	};

    //************************************************************
    //  [�@�\]  �ȈՐ��l�^�`�F�b�N
    //************************************************************
	function f_CheckNum(pFromName){

		wFromName = eval(pFromName);
		if (isNaN(wFromName.value)){
			wFromName.focus();
			return false;
		}else{

			//�}�C�i�X���`�F�b�N
			var wStr = new String(wFromName.value)
			if (wStr.match("-")!=null){
				wFromName.focus();
				return false;
			};

			//�����_�`�F�b�N
			w_decimal = new Array();
			w_decimal = wStr.split(".")
			if(w_decimal.length>1){
				wFromName.focus();
				return false;
			};
		}
		return true;
	}

    //************************************************************
    //  [�@�\]  �召�`�F�b�N
    //************************************************************
	function f_CheckDaisyou(){

		wObj1 = eval("parent.topFrame.document.frm.txtSouJyugyou");
		wObj2 = eval("parent.topFrame.document.frm.txtJunJyugyou");

		if(wObj1.value != "" && wObj2.value != ""){
			if(wObj1.value < wObj2.value){
				wObj1.focus();
				return false;
			}
		}
		return true;
	}

	//************************************************
	//Enter �L�[�ŉ��̓��̓t�H�[���ɓ����悤�ɂȂ�
	//�����Fp_inpNm	�Ώۓ��̓t�H�[����
	//    �Fp_frm	�Ώۃt�H�[��
	//�@�@�Fi		���݂̔ԍ�
	//�ߒl�F�Ȃ�
	//���̓t�H�[�������Axxxx1,xxxx2,xxxx3,�c,xxxxn 
	//�̖��O�̂Ƃ��ɗ��p�ł��܂��B
	//************************************************
	function f_MoveCur(p_inpNm,p_frm,i){
		if (event.keyCode == 13){		//�����ꂽ�L�[��Enter(13)�̎��ɓ����B
			i++;

			//���͉\�̃e�L�X�g�{�b�N�X��T���B����������t�H�[�J�X���ڂ��ď����𔲂���B
	        for (w_li = 1; w_li <= 99; w_li++) {
				
				if (i > <%=m_rCnt%>) i = 1; //i���ő�l�𒴂���ƁA�͂��߂ɖ߂�B
				inpForm = eval("p_frm."+p_inpNm+i);

				//���͉\�̈�Ȃ�t�H�[�J�X���ڂ��B
				if (typeof(inpForm) != "undefined") {
					inpForm.focus();			//�t�H�[�J�X���ڂ��B
					inpForm.select();			//�ڂ����e�L�X�g�{�b�N�X����I����Ԃɂ���B
					break;

				//���͕t���Ȃ玟�̍��ڂ�
				} else{
					i++
				}
	        }
		}else{
	//		alert(event.keyCode);
			return false;
		}
		return true;
	}

	//-->
	</SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post" onClick="return false;">
	
	<%	m_Rs.MoveFirst
		Do Until m_Rs.EOF
			w_ihalf = gf_Round(m_rCnt / 2,0)
			'i = i + 1 
			j = j + 1 
			w_sSeiseki = ""
			w_sHyoka = ""
			w_sKekka = ""
			w_sChikai = ""
			w_sGakusekiCd = ""
			w_sKekkasu = ""
			w_sChikaisu = ""
			Call gs_cellPtn(w_cell)
'If w_ihalf + 1 = i then
'				If w_ihalf + 1 = j then
'				w_cell = ""
'				Call gs_cellPtn(w_cell)
%><center><table class="hyo" border="1" align="center" width="555" nowrap>
<!--		</table>
	</td>
	<td valign="top" width="50%">
		<table class="hyo" border="1" align="center" width="98%">
--><%
	'	 		End If 

		'//�e�f�[�^���擾����
		'** ��O�̎����敪
		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU								'//�O������
				w_sShikenKBN_RUI = 99
				
			Case C_SIKEN_ZEN_KIM								'//�O������
				w_sShikenKBN_RUI = C_SIKEN_ZEN_TYU
				
			Case C_SIKEN_KOU_TYU								'//�������
				w_sShikenKBN_RUI = C_SIKEN_ZEN_KIM
				
			Case C_SIKEN_KOU_KIM								'//�������
				w_sShikenKBN_RUI = C_SIKEN_KOU_TYU
		End Select

'/**** �ȉ��\�������Ő��сA���ۂ̕\����NULL->0�ϊ����Ȃ���NULL��""�ŕ\������ 2002.03.20 matsuo ****/

		w_sGakusekiCd = m_Rs("GAKUSEKI_NO")

		'2002.03.20
		'w_sKekka = gf_SetNull2Zero(m_Rs("KEKA"))
		'w_sKekkaGai = gf_SetNull2Zero(m_Rs("KEKA_NASI"))
		'w_sChikai = gf_SetNull2Zero(m_Rs("CHIKAI"))

		w_sKekka = gf_SetNull2String(m_Rs("KEKA"))
		w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI"))
		w_sChikai = gf_SetNull2String(m_Rs("CHIKAI"))

'�O���ŏI����Ă���Ȗڂ̌��ۂ��擾���Ċw�������тɃZ�b�g����B2002/02/21 ITO

		'//�O���݂̂̏ꍇ��T21���O�L���������܂ł̌��ې����擾����
		
		w_iRet = f_SikenInfo(w_bZenkiOnly)

		'�w�N�������̏ꍇ�̂�
		If m_sSikenKBN = C_SIKEN_KOU_KIM Then
		
			'�O���J�݂�������O�������̌��ۂ��w�N���̐��тɃZ�b�g����
			If w_bZenkiOnly = True Then

				'�w�������т�0
				'2002.03.20
				'If gf_SetNull2Zero(m_Rs("KEKA")) = "0" Then 
				If gf_SetNull2String(m_Rs("KEKA")) = "" Then 

					'w_sKekka = gf_SetNull2Zero(m_Rs("KEKA_ZK"))			'���ې�
					'w_sKekkaGai = gf_SetNull2Zero(m_Rs("KEKA_NASI_ZK"))	'���ۑΏۊO
					'w_sChikai = gf_SetNull2Zero(m_Rs("CHIKAI_ZK"))		'�x����

					w_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))			'���ې�
					w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK"))	'���ۑΏۊO
					w_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))		'�x����
				
				End If
				
			End If
		End If

'�O���ŏI����Ă���Ȗڂ̌��ۂ��擾���Ċw�������тɃZ�b�g����B2002/02/21 ITO


	'�l�̏������B
	w_bNoChange = False
	'2002.03.20
	'w_sKekkasu = 0
	'w_sChikaisu = 0
	w_sKekkasu = ""
	w_sChikaisu = ""

	'---------------------------------------------------------------------------------------------
	'�ʏ���ƂƂ��̏���
	if m_TUKU_FLG = C_TUKU_FLG_TUJO then 
		'2002.03.20
		'w_sSeiseki = m_Rs("SEI")
		w_sSeiseki = gf_SetNull2String(m_Rs("SEI"))
		w_sHyoka = gf_HTMLTableSTR(m_Rs("HYOKAYOTEI"))

'�O���ŏI����Ă���Ȗڂ̌��ۂ��擾���Ċw�������тɃZ�b�g����B2002/02/21 ITO

		'�w�N�������̏ꍇ�̂�
		If m_sSikenKBN = C_SIKEN_KOU_KIM Then
		
			'�O���J�݂�������O�������̌��ۂ��w�N���̐��тɃZ�b�g����
			If w_bZenkiOnly = True Then

				'�w�������т�0
				'2002.03.20
				'If gf_SetNull2Zero(m_Rs("SEI")) = "0" Then 
				If gf_SetNull2String(m_Rs("SEI")) = "" Then 

					'w_sSeiseki = gf_SetNull2Zero(m_Rs("SEI_ZK"))			'�O����������
					w_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))			'�O����������
				
				End If
				
			End If
		End If

'�O���ŏI����Ă���Ȗڂ̌��ۂ��擾���Ċw�������тɃZ�b�g����B2002/02/21 ITO

		if w_sHyoka = "�@" then w_sHyoka = "�E"

		'//�Ȗڂ��I���Ȗڂ̏ꍇ�́A���k���I�����Ă��邩�ǂ����𔻕ʂ���B�I�������Ȃ����k�͓��͕s�Ƃ���B
		w_bNoChange = False

		If cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then 
			If cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then
				w_bNoChange = True
			End If 
		Else
			if Cstr(m_iLevelFlg) = "1" then
				if isNull(m_Rs("T16_LEVEL_KYOUKAN")) = true then
'					response.write " A "
					w_bNoChange = True
				else
'					response.write " B " & m_Rs("T16_LEVEL_KYOUKAN")
					if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
'					response.write " C "
						w_bNoChange = True
					End if
				End if
			End if
		End If

	End if

	'==�ٓ��b�g�j�i2001/12/19���o�[�W����:okada�j================================
		Dim w_Date
		Dim w_SSSS
		Dim w_SSSR
		
		'w_Date = gf_YYYY_MM_DD(m_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		w_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")
 '//C_IDO_FUKUGAKU=3:���w�AC_IDO_TEI_KAIJO=5:��w����
		w_SSSS = ""
		w_SSSR = ""

		w_SSSS = gf_Get_IdouChk(w_sGakusekiCd,w_Date,m_iNendo,w_SSSR)

		IF CStr(w_SSSS) <> "" Then
			IF Cstr(w_SSSS) <> CStr(C_IDO_FUKUGAKU) AND Cstr(w_SSSS) <> Cstr(C_IDO_TEI_KAIJO) Then
				w_SSSS = "[" & w_SSSR & "]"
				w_bNoChange = True
			Else
				w_SSSR = ""
				w_SSSS = ""
			End if
		End if

	if Cstr(m_TUKU_FLG) = Cstr(C_TUKU_FLG_TUJO) then 

	'=========================================================

		'���ےx�����̎擾�@���ݒʏ���Ƃ̂�
		
		'2002.03.20
		'//���ې��~�P�ʐ��̎擾
		w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_KEKKA)		'�߂�l��NULL�̎���""

'		w_sKekkasu = gf_IIF(trim(w_sData) = "", "", cint(gf_SetNull2Zero(w_sData) * m_iJigenTani))

		'gf_IIF�ɓn���Ƃ��Ƀp�����[�^���v�Z����̂ŁA�p�����[�^��0�ɕϊ�
		w_sKekkasu = gf_IIF(w_sData = "", "", cint(gf_SetNull2Zero(w_sData)) * cint(m_iJigenTani))
		'w_sKekkasu = cint(f_Syukketu2(w_sGakusekiCd,C_KETU_KEKKA)) * m_iJigenTani			'//���ې��~�P�ʐ��̎擾

		'//�P���ۂ̏ꍇ�̌��ې��̎擾
		w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_KEKKA_1)

		if w_sKekkasu = "" and w_sData = "" then
			'w_sKekkasu��w_sData���ǂ����""�̎���""�̂܂�
			w_sKekkasu = ""
		else
			'�ǂ��炩���""�łȂ���Όv�Z
			w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_sData))			'//�P���ۂ̏ꍇ�̌��ې��̎擾
		end if
		'w_sKekkasu = w_sKekkasu + cint(f_Syukketu2(w_sGakusekiCd,C_KETU_KEKKA_1))			'//�P���ۂ̏ꍇ�̌��ې��̎擾

		'//�x�����̎擾
		w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_TIKOKU)
		w_sChikaisu = gf_IIF(w_sData = "", "", cint(gf_SetNull2Zero(w_sData)))
		'w_sChikaisu = cint(f_Syukketu2(w_sGakusekiCd,C_KETU_TIKOKU))		'//�x�����̎擾
		
		'//���ސ��̎擾
		w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_SOTAI)
		if w_sChikaisu = "" and w_sData = "" then
			'w_sKekkasu��w_sData���ǂ����""�̎���""�̂܂�
			w_sChikaisu = ""
		else
			'�ǂ��炩���""�łȂ���Όv�Z
			w_sChikaisu = cint(gf_SetNull2Zero(w_sChikaisu)) + cint(gf_SetNull2Zero(w_sData))			'//�P���ۂ̏ꍇ�̌��ې��̎擾
		end if

		'w_sChikaisu = w_sChikaisu + cint(f_Syukketu2(w_sGakusekiCd,C_KETU_SOTAI))		'//���ސ��̎擾

	end if
	'---------------------------------------------------------------------------------------------

		'�u�o�����ۂ��ݐρv�Łu�O�����ԂłȂ��v�̏ꍇ
		'���ہE���Ȃ�Null�������ꍇ�A�����邽�ߊ֐��ǉ� Add 2001.12.16 okada
		if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and w_sShikenKBN_RUI <> 99 then 
	 		'2002.03.20
			'call gf_GetKekaChi(m_iNendo,w_sShikenKBN_RUI,m_sKamokuCd,cstr(m_Rs("GAKUSEI_NO")),w_iKekka_rui,w_iChikoku_rui,w_iKekkaGai_rui) '��O�̎����̍��v�l�𑫂��B
	 		call f_GetKekaChi(m_iNendo,w_sShikenKBN_RUI,m_sKamokuCd,cstr(m_Rs("GAKUSEI_NO")),w_iKekka_rui,w_iChikoku_rui,w_iKekkaGai_rui) '��O�̎����̍��v�l�𑫂��B
			
			'�ǂ����""�̎���""
			if w_sKekkasu = "" and w_iKekka_rui = "" then
				w_sKekkasu = ""
			else
				w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))
			end if
			'w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))

			'�ǂ����""�̎���""
			if w_sChikaisu = "" and w_iChikoku_rui = "" then
				w_sChikaisu = ""
			else
				w_sChikaisu = cint(gf_SetNull2Zero(w_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
			end if
			'w_sChikaisu = cint(gf_SetNull2Zero(w_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
		end if
		
		'2002.03.20
		'If cint(w_sKekka) = 0 and cint(w_sKekkasu) > 0 Then 		'//������0��,���v��0���傫���ꍇ
		If cint(gf_SetNull2Zero(w_sKekka)) = 0 and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then 		'//������0��,���v��0���傫���ꍇ

			'w_sKekka = cint(w_sKekkasu)		'//���������v
			w_sKekka = cint(gf_SetNull2Zero(w_sKekkasu))								'//���������v
		End If

		'If cint(w_sChikai) = 0 AND cint(w_sChikaisu) > 0 Then		'//�x����0��,�x�v��0���傫���ꍇ
		If cint(gf_SetNull2Zero(w_sChikai)) = 0 AND cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then		'//�x����0��,�x�v��0���傫���ꍇ
			'w_sChikai = cint(w_sChikaisu)							'//�x�����x�v
			w_sChikai = cint(gf_SetNull2Zero(w_sChikaisu))							'//�x�����x�v
		End If
	%><!-- <tr> --><%
			'========================================================================================
			'//�Ȗڂ��I���Ȗڂ̎��ɉȖڂ�I�����Ă��Ȃ��ꍇ(���͕s��)
			'========================================================================================
			If w_bNoChange = True Then%>
			<%
				if Cstr(m_TUKU_FLG) = Cstr(C_TUKU_FLG_TUJO) Then%>
					<input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>">
					<input type="hidden" name="hidUpdFlg<%=i%>" value="False">
					<td class="<%=w_cell%>" width="45" nowrap><%=w_sGakusekiCd%></td>
					<td class="<%=w_cell%>" align="left" width="250" nowrap><%=m_Rs("SIMEI")%><%=w_SSSS%></td>
					<td class="<%=w_cell%>" align="center" width="25" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="25" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="25" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="25" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="35" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="35" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="35" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="45" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="35" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="40" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="45" nowrap>-</td>
				<%Else%>
					<input type="hidden" name=txtGseiNo<%=i%> value="<%=m_Rs("GAKUSEI_NO")%>">
					<input type="hidden" name="hidUpdFlg<%=i%>" value="False">
					<td class="<%=w_cell%>" width="45" ><%=w_sGakusekiCd%></td>
					<td class="<%=w_cell%>" align="left" width="250" nowrap><%=m_Rs("SIMEI")%><%=w_SSSS%></td>
					<td class="<%=w_cell%>" align="center" width="25" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="25" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="25" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="25" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="35" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="35" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="80" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="60" nowrap>-</td>
					<td class="<%=w_cell%>" align="center" width="60" nowrap>-</td>
				<%End if%>
			<%
			'=========================================================================
			'//�Ȗڂ��K�C���A�܂��͑I���Ȗڂ̎��ɐ��k���Ȗڂ�I�����Ă���ꍇ(���͉�)
			'=========================================================================
			Else
				%>
						<td class="<%=w_cell%>"  width="40" nowrap><%=w_sGakusekiCd%><input type="hidden" name=txtGseiNo<%=i%> value="<%=m_Rs("GAKUSEI_NO")%>"></td>
						<input type="hidden" name="hidUpdFlg<%=i%>" value="True">
						<td class="<%=w_cell%>" align="left"  width="250" nowrap><%=m_Rs("SIMEI")%><%=w_SSSS%></td>
					<!-- 2002.03.20 -->
					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
						<td class="<%=w_cell%>" align="center" width="25" style="font-size:10px" nowrap><%=m_Rs("SEI1")%></td>
						<td class="<%=w_cell%>" align="center" width="25" style="font-size:10px" nowrap><%=m_Rs("SEI2")%></td>
						<td class="<%=w_cell%>" align="center" width="25" style="font-size:10px" nowrap><%=m_Rs("SEI3")%></td>
						<td class="<%=w_cell%>" align="center" width="25" style="font-size:10px" nowrap><%=m_Rs("SEI4")%></td>
					<%Else%>
						<td class="<%=w_cell%>" width="25" nowrap>&nbsp;</td>
						<td class="<%=w_cell%>" width="25" nowrap>&nbsp;</td>
						<td class="<%=w_cell%>" width="25" nowrap>&nbsp;</td>
						<td class="<%=w_cell%>" width="25" nowrap>&nbsp;</td>
					<%End If%>
						<%
						'//NN�Ή�
						If session("browser") = "IE" Then
							w_sInputClass = "class='num'"
							w_sInputClass1 = "class='num'"
							w_sInputClass2 = "class='num'"
						Else
							w_sInputClass = ""
							w_sInputClass1 = ""
							w_sInputClass2 = ""
						End If

						if m_iKikan = "NO" Then
							w_sInputClass1 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
						End if

						'// ���ۓ��͉\�׸�
						if Not m_bKekkaNyuryokuFlg then
							w_sInputClass2 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
						End if

				'=========================================================================
				'//�ʏ���Ƃ̏ꍇ 
				'=========================================================================
						%>
				<%If m_iKikan <> "NO" Then%>
					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
							
						<td class="<%=w_cell%>" width="35"align="center" nowrap><input type="text" <%= w_sInputClass1 %>  name=Seiseki<%=i%> value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)" onChange="f_GetTotalAvg()"></td>
						<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>
								<td class="<%=w_cell%>"  width="35" align="center" nowrap><input type="button" size="2" name="button<%=i%>" value="<%=w_sHyoka%>" onClick="return f_change(<%=i%>)" style="text-align:center" class="<%=w_cell%>"><!-- class="<%=w_cell%>"-->
								<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>"></td>
						<%Else%>
								<td class="<%=w_cell%>"  width="35" align="center" nowrap><%=w_sHyoka%><input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>"></td>
						<%End If%>
							<td class="<%=w_cell%>" width="35" align="center" nowrap><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="45" align="right" nowrap><%=w_sChikaisu%></td>
							<td class="<%=w_cell%>" width="35" align="center" nowrap><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="40" align="center" nowrap><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="45" align="right" nowrap><%=w_sKekkasu%></td>
					<%Else%>
							<td class="<%=w_cell%>" width="35" nowrap align="center">-</td>
							<td class="<%=w_cell%>" width="35" nowrap align="center">-</td>
							<td class="<%=w_cell%>" width="80" nowrap align="center"><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="60" nowrap align="center"><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="60" nowrap align="center"><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
					<%End If%>

				<%Else%>

					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
							
						<td class="<%=w_cell%>" width="35" align="right" nowrap><input type="text" <%= w_sInputClass1 %>  name=Seiseki<%=i%> value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)" onChange="f_GetTotalAvg()"></td>
						<%	'�\���݂̂̏ꍇ�̍��v�E���ϒl�����߂�
							If IsNull(w_sSeiseki) = False Then
								If IsNumeric(CStr(w_sSeiseki)) = True Then
									w_lSeiTotal = w_lSeiTotal + CLng(w_sSeiseki)
									w_lGakTotal = w_lGakTotal + 1
								End If
							End If
						%>
						<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>
								<td class="<%=w_cell%>"  width="35" align="center" nowrap><%=trim(w_sHyoka)%></td>
						<%Else%>
								<td class="<%=w_cell%>"  width="35" align="center" nowrap><%=trim(w_sHyoka)%></td>
						<%End If%>
						<td class="<%=w_cell%>" width="35" align="right" nowrap><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="45" align="right" nowrap><%=w_sChikaisu%></td>
						<td class="<%=w_cell%>" width="35" align="right" nowrap><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="40" align="right" nowrap><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="45" align="right" nowrap><%=w_sKekkasu%></td>
					<%Else%>
						<td class="<%=w_cell%>" width="35" align="center" nowrap>-</td>
						<td class="<%=w_cell%>" width="35" align="center" nowrap>-</td>
						<td class="<%=w_cell%>" width="80" align="center" nowrap><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="60" align="center" nowrap><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="60" align="center" nowrap><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
					<%End If%>
				<%End If%>
			<%End If%>
					</tr>
			<%
				m_Rs.MoveNext
				i = i + 1
			Loop%>
			<tr>
				<td class="header" nowrap align="right" colspan="4">
					<FONT COLOR="#FFFFFF"><B>���э��v</B></FONT>
					<input type="text" name="txtTotal" size="5" <%=w_sInputClass%> readonly>
				</td>
				<td class="header" nowrap colspan="6"></td>
			</tr>
			<tr>
				<td class="header" nowrap align="right" colspan="4">
					<FONT COLOR="#FFFFFF"><B>���ϓ_</B></FONT>
					<input type="text" name="txtAvg" size="5" <%=w_sInputClass%> readonly>
				</td>
				<td class="header" nowrap colspan="6"></td>
			</tr>
		</table>
<!--
	</td>
	</tr>
	</table>
-->
	<table width="50%">
	<tr>
		<td align="center">
		<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%>
			<input type="button" class="button" value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
		<%End If%>
		<input type="button" class="button" value="�L�����Z��" onclick="javascript:f_Cansel()"></td>
	</tr>
	</table>

	<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="KamokuCd"    value="<%=m_sKamokuCd%>">
	<input type="hidden" name="i_Max"       value="<%=i%>">
	<input type="hidden" name="txtSikenKBN" value="<%=m_sSikenKBN%>">
	<input type="hidden" name="txtGakuNo"   value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakkaCd"  value="<%=m_sGakkaCd%>">
	<input type="hidden" name="txtClassNo"  value="<%=m_sClassNo%>">
	<input type="hidden" name="txtKamokuCd" value="<%=m_sKamokuCd%>">
	<input type="hidden" name="txtTUKU_FLG" value="<%=m_TUKU_FLG%>">
	<input type="hidden" name="PasteType"   value="">
	<!-- 02/03/27 �ǉ� -->
	<input type="hidden" name="hidSouJyugyou">
	<input type="hidden" name="hidJunJyugyou">
	<input type="hidden" name="hidUpdMode">

	</FORM>
	</center>
	</body>
	<SCRIPT>
		//2002/02/05 ���� �ǉ�
		//************************************************************
		//	[�@�\]	���т��ύX���ꂽ�Ƃ�
		//	[����]	�Ȃ�
		//	[�ߒl]	�Ȃ�
		//	[����]	���т̍��v�ƕ��ς����߂�
		//	[���l]	�w���̑�����������͍̂Ō�ł��邽�߁A���̈ʒu�ɏ����B
		//************************************************************
		function f_GetTotalAvg(){
			var i;
			var total;
			var avg;
			var cnt;

			total = 0;
			cnt = 0;
			avg = 0;

	<%If m_iKikan <> "NO" Then	'���͊��Ԓ�%>

			//�w�����ł̃��[�v
			for(i=0;i<<%=i%>;i++) {

				//���݂��邩�ǂ���
				textbox = eval("document.frm.Seiseki" + (i+1));
				if (textbox) {
					//�����̓`�F�b�N
					if (textbox.value != "") {
						//�����łȂ��͖̂�������
						if (!isNaN(textbox.value)) {
							total = total + parseInt(textbox.value);
						}
					}
					cnt = cnt + 1;
				}
			}

	<% Else	'���͊��Ԓ��ł͂Ȃ�%>
		total = <%=w_lSeiTotal%>;
		cnt   = <%=w_lGakTotal%>;
	<% End If%>

			document.frm.txtTotal.value=total;

			//�l�̌ܓ�
			if (cnt!=0){
				avg = total/cnt;
				avg = avg * 10;
				avg = Math.round(avg);
				avg = avg / 10;
			}
			
			document.frm.txtAvg.value=avg;
		}
	</SCRIPT>

	</html>
<%
End sub

Sub No_showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT language="javascript">
	<!--

	    //************************************************************
	    //  [�@�\]  �y�[�W���[�h������
	    //  [����]
	    //  [�ߒl]
	    //  [����]
	    //************************************************************
	    function window_onload() {

	    }

	//-->
	</SCRIPT>
	</head>

    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	<center>
	<br><br><br>
		<span class="msg">���ѓ��͊��ԊO�ł��B</span>
	</center>
	<input type="hidden" name="txtMsg" value="���ѓ��͊��ԊO�ł��B">

	</form>
	</body>
	</html>

<%
End Sub
Sub showPage_No()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT language="javascript">
	<!--

	    //************************************************************
	    //  [�@�\]  �y�[�W���[�h������
	    //  [����]
	    //  [�ߒl]
	    //  [����]
	    //************************************************************
	    function window_onload() {

	    }

	//-->
	</SCRIPT>
	</head>

    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	</head>

	<body>
	<br><br><br>
	<center>
		<span class="msg">�l���C�f�[�^�����݂��܂���B</span>
	</center>

	<input type="hidden" name="txtMsg" value="�l���C�f�[�^�����݂��܂���B">

	</form>
	</body>
	</html>

<%
End Sub
%>