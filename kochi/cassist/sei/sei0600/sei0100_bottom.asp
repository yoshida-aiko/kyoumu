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
    
    Public m_iKamoku_Kbn
    Public m_iHissen_Kbn

	Public	m_Rs
	Public	m_TRs
	Public	m_DRs
	Public	m_SRs
	Public	m_iMax			'�ő�y�[�W

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
			'// �y�[�W��\��
			Call No_showPage()
			Exit Do
		End If
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
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
		'//�Ȗڋ敪(0:��ʉۖ�,1:���Ȗ�)�A�y�сA�K�C�I���敪(1:�K�C,2:�I��)�𒲂ׂ�
        w_iRet = f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn)
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

Sub s_SetParam()
'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSikenKBN	= Cint(request("txtSikenKBN"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sGakkaCd	= request("txtGakkaCd")

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

	Do
		'==========================================
		'//�ȖڒS�������̋���CD�̎擾
		'==========================================
        'w_iRet = f_GetTantoKyokan()
        w_iRet = f_GetTantoKyokan(w_sTKyokanCd)
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
		if m_SRs.EOF = false then m_SRs.MoveFirst
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
Function f_GetKamokuInfo(p_iKamoku_Kbn,p_iHissen_Kbn)

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
		End If

        f_GetKamokuInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'Function f_getdate()
Function f_getdate(p_iKamoku_Kbn)
'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Dim w_iNyuNendo


	On Error Resume Next
	Err.Clear
	f_getdate = 1

	Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakuNo) + 1

		'//�������ʂ̒l���ꗗ��\��
		w_sSQL = ""
		w_sSQL = w_sSQL & " SELECT "

		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_Z AS SEI,A.T16_KEKA_TYUKAN_Z AS KEKA,A.T16_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_Z AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI, "
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI,A.T16_KEKA_KIMATU_Z AS KEKA,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T16_CHIKAI_KIMATU_Z AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI,A.T16_KEKA_TYUKAN_K AS KEKA,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T16_CHIKAI_TYUKAN_K AS CHIKAI,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
		End Select

		w_sSQL = w_sSQL & " 	A.T16_GAKUSEI_NO AS GAKUSEI_NO,A.T16_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_SELECT_FLG"
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
		If cint(p_iKamoku_Kbn) = cint(C_KAMOKU_SENMON) Then
			w_sSQL = w_sSQL & vbCrLf & " AND	C.T13_GAKKA_CD = '" & m_sGakkaCd & "' "
		Else
			w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
		End If

		w_sSQL = w_sSQL & " AND	A.T16_NENDO = C.T13_NENDO "

		'//�u�����̐��k�͂͂���(C_TIKAN_KAMOKU_MOTO = 1    '�u����)
		w_sSQL = w_sSQL & " AND	A.T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
'		w_sSQL = w_sSQL & " AND	B.T11_NYUNENDO = " & w_iNyuNendo & " "
		w_sSQL = w_sSQL & " ORDER BY A.T16_GAKUSEKI_NO "

'response.write w_sSQL &"<<br>"

		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_getdate = 99
			m_bErrFlg = True
			Exit Do 
		End If

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
			Case C_SIKEN_ZEN_KIM
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_Z AS KEKA,A.T34_KEKA_NASI_KIMATU_Z AS KEKA_NASI,A.T34_CHIKAI_KIMATU_Z AS CHIKAI, "
			Case C_SIKEN_KOU_TYU
				w_sSQL = w_sSQL & " 	A.T34_KEKA_TYUKAN_K AS KEKA,A.T34_KEKA_NASI_TYUKAN_K AS KEKA_NASI,A.T34_CHIKAI_TYUKAN_K AS CHIKAI, "
			Case C_SIKEN_KOU_KIM
				w_sSQL = w_sSQL & " 	A.T34_KEKA_KIMATU_K AS KEKA,A.T34_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T34_CHIKAI_KIMATU_K AS CHIKAI, "
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

		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_getTUKUclass = 99
			m_bErrFlg = True
			Exit Do 
		End If

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
'Function f_GetTantoKyokan()
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

Function f_Nyuryokudate()
'********************************************************************************
'*	[�@�\]	���ѓ��͊��ԃf�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI"
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
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
'//���ѓ��͊��ԃe�X�g�p

'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '2003/04/30'"
'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '1999/03/01'"

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

Function f_getTUKU(p_iNendo,p_sKamoku,p_iGakunen,p_iClass,p_TUKU_FLG)
'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

	On Error Resume Next
	Err.Clear
	f_getTUKU = 0
	p_TUKU_FLG = "0"

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
			p_TUKU_FLG = w_Rs("T20_TUKU_FLG")
		End If

		Exit Do
	Loop
	
    Call gf_closeObject(w_Rs)

End Function

Function f_Syukketu(p_gaku,p_kbn)
'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	f_Syukketu = 0

	If m_SRs.EOF Then
		Exit Function
	Else
		m_SRs.MoveFirst
		Do Until m_SRs.EOF
			If m_SRs("T21_GAKUSEKI_NO") = p_gaku AND cstr(m_SRs("T21_SYUKKETU_KBN")) = cstr(p_kbn) Then
				f_Syukketu = m_SRs("KAISU")
				Exit Do
			End If
		m_SRs.MoveNext
		Loop
	End If

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

i = 0

%>
<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
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

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		//�w�b�_���󔒕\��
		parent.topFrame.document.location.href="white.asp"

		//�o�^����
<% if m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
        document.frm.action="sei0100_upd.asp";
<% Else %>
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
<% if m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_Seiseki = eval("document.frm.Seiseki"+i);
			w_bFLG = true
			if (isNaN(w_Seiseki.value)){
				w_bFLG = false;
				return 1;
				break;
			}else{

				//�}�C�i�X���`�F�b�N
				var wStr = new String(w_Seiseki.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					return 1;
					break;
				};

				//�����_�`�F�b�N
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					return 1;
					break;
				}

			};
		};
			if (w_bFLG == false){
				return 1;
			};
<% End if %>
		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_Kekka = eval("document.frm.Kekka"+i);
			w_bFLG = true
			if (isNaN(w_Kekka.value)){
				w_bFLG = false;
				return 1;
				break;
			}else{

				//�}�C�i�X���`�F�b�N
				var wStr = new String(w_Kekka.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					return 1;
					break;
				};

				//�����_�`�F�b�N
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					return 1;
					break;
				}

			};
		};
			if (w_bFLG == false){
				return 1;
			};

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_KekkaGai = eval("document.frm.KekkaGai"+i);
			w_bFLG = true
			if (isNaN(w_KekkaGai.value)){
				w_bFLG = false;
				return 1;
				break;
			}else{

				//�}�C�i�X���`�F�b�N
				var wStr = new String(w_KekkaGai.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					return 1;
					break;
				};

				//�����_�`�F�b�N
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					return 1;
					break;
				}

			};
		};
			if (w_bFLG == false){
				return 1;
			};

		var i
		for (i = 1; i < document.frm.i_Max.value; i++) {
			w_Chikai = eval("document.frm.Chikai"+i);
			w_bFLG = true
			if (isNaN(w_Chikai.value)){
				w_bFLG = false;
				return 1;
				break;
			}else{

				//�}�C�i�X���`�F�b�N
				var wStr = new String(w_Chikai.value)
				if (wStr.match("-")!=null){
					w_bFLG = false;
					return 1;
					break;
				};

				//�����_�`�F�b�N
				w_decimal = new Array();
				w_decimal = wStr.split(".")
				if(w_decimal.length>1){
					w_bFLG = false;
					return 1;
					break;
				}

			};
		};

			if (w_bFLG == false){
				return 1;
			};
		return 0;
	};

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
		if (i > <%=m_rCnt%>) i = 1; //i���ő�l�𒴂���ƁA�͂��߂ɖ߂�B
		inpForm = eval("p_frm."+p_inpNm+i);
		inpForm.focus();			//�t�H�[�J�X���ڂ��B
		inpForm.select();			//�ڂ����e�L�X�g�{�b�N�X����I����Ԃɂ���B
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
	<center>

<!--
	<table border=1>
	<tr>
	<td valign="top">
-->
		<table class="hyo" border="1" align="center" width="550">
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
'				Call gs_cellPtn(w_cell)%>
<!--		</table>
	</td>
	<td valign="top" width="50%">
		<table class="hyo" border="1" align="center" width="98%">
-->
	<%
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

		w_sGakusekiCd = m_Rs("GAKUSEKI_NO")
		w_sKekka = gf_SetNull2Zero(m_Rs("KEKA"))
		w_sKekkaGai = gf_SetNull2Zero(m_Rs("KEKA_NASI"))
		w_sChikai = gf_SetNull2Zero(m_Rs("CHIKAI"))
	'�l�̏������B
	w_bNoChange = False
	w_sKekkasu = 0
	w_sChikaisu = 0
	'---------------------------------------------------------------------------------------------
	'�ʏ���ƂƂ��̏���
	if m_TUKU_FLG = C_TUKU_FLG_TUJO then 
		w_sSeiseki = m_Rs("SEI")
		w_sHyoka = gf_HTMLTableSTR(m_Rs("HYOKAYOTEI"))
if w_sHyoka = "�@" then w_sHyoka = "�E"
			'//�Ȗڂ��I���Ȗڂ̏ꍇ�́A���k���I�����Ă��邩�ǂ����𔻕ʂ���B�I�������Ȃ����k�͓��͕s�Ƃ���B
			w_bNoChange = False

			If cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then 
				If cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then
					w_bNoChange = True
				End If 
			End If
		'���ےx�����̎擾�@���ݒʏ���Ƃ̂�
		w_sKekkasu = cint(f_Syukketu(w_sGakusekiCd,C_KETU_KEKKA))			'//���ې��̎擾

		w_sChikaisu = cint(f_Syukketu(w_sGakusekiCd,C_KETU_TIKOKU))		'//�x�����̎擾
		w_sChikaisu = w_sChikaisu + cint(f_Syukketu(w_sGakusekiCd,C_KETU_SOTAI))		'//���ސ��̎擾

	end if
	'---------------------------------------------------------------------------------------------

		'�u�o�����ۂ��ݐρv�Łu�O�����ԂłȂ��v�̏ꍇ
		if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and w_sShikenKBN_RUI <> 99 then 
	 		call gf_GetKekaChi(m_iNendo,w_sShikenKBN_RUI,m_sKamokuCd,cstr(m_Rs("GAKUSEI_NO")),w_iKekka_rui,w_iChikoku_rui,w_iKekkaGai_rui) '��O�̎����̍��v�l�𑫂��B
			w_sKekkasu = cint(w_sKekkasu) + cint(w_iKekka_rui)
			w_sChikaisu = cint(w_sChikaisu) + cint(w_iChikoku_rui)
		end if
		
		If cint(w_sKekka) = 0 and cint(w_sKekkasu) > 0 Then 		'//������0��,���v��0���傫���ꍇ
			w_sKekka = cint(w_sKekkasu)								'//���������v
		End If
		If cint(w_sChikai) = 0 AND cint(w_sChikaisu) > 0 Then		'//�x����0��,�x�v��0���傫���ꍇ
			w_sChikai = cint(w_sChikaisu)							'//�x�����x�v
		End If
	%>
		<tr>
	<%

			'========================================================================================
			'//�Ȗڂ��I���Ȗڂ̎��ɉȖڂ�I�����Ă��Ȃ��ꍇ(���͕s��)
			'========================================================================================
			If w_bNoChange = True Then%>

						<td class="<%=w_cell%>" width="40" ><%=w_sGakusekiCd%></td>
						<td class="<%=w_cell%>" align="left"   width="260"><%=m_Rs("SIMEI")%></td>
						<td class="<%=w_cell%>" align="center" width="35" >-</td>
						<td class="<%=w_cell%>" align="center" width="35" >-</td>
						<td class="<%=w_cell%>" align="center" width="40" >-</td>
						<td class="<%=w_cell%>" align="center" width="40" >-</td>
						<td class="<%=w_cell%>" align="center" width="40" >-</td>
						<td class="<%=w_cell%>" align="center" width="35" >-</td>
						<td class="<%=w_cell%>" align="center" width="35" >-</td>
			<%
			'=========================================================================
			'//�Ȗڂ��K�C���A�܂��͑I���Ȗڂ̎��ɐ��k���Ȗڂ�I�����Ă���ꍇ(���͉�)
			'=========================================================================
			Else
				i = i+1
				%>
						<td class="<%=w_cell%>"  width="40"><%=w_sGakusekiCd%>
						<input type="hidden" name=txtGseiNo<%=i%> value="<%=m_Rs("GAKUSEI_NO")%>"></td>
						<td class="<%=w_cell%>" align="left"  width="210"><%=m_Rs("SIMEI")%></td>

						<%
						'//NN�Ή�
						If session("browser") = "IE" Then
							w_sInputClass = "class='num'"
						Else
							w_sInputClass = ""
						End If
				'=========================================================================
				'//�ʏ���Ƃ̏ꍇ
				'=========================================================================
						%>
				<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
						
							<td class="<%=w_cell%>" width="30"><input type="text" <%=w_sInputClass%>  name=Seiseki<%=i%> value="<%=w_sSeiseki%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)"></td>
					<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>
							<td class="<%=w_cell%>"  width="30"><input type="button" size="2" name="button<%=i%>" value="<%=w_sHyoka%>" onClick="return f_change(<%=i%>)" style="text-align:center" class="<%=w_cell%>"><!-- class="<%=w_cell%>"-->
							<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>"></td>
					<%Else%>
							<td class="<%=w_cell%>"  width="30"><%=w_sHyoka%><input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>"></td>
					<%End If%>
						<td class="<%=w_cell%>" width="20"><input type="text" <%=w_sInputClass%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="20"><input type="text" <%=w_sInputClass%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="30"align="right"  ><%=w_sKekkasu%></td>
						<td class="<%=w_cell%>" width="20"><input type="text" <%=w_sInputClass%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=1 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="25"align="right"  ><%=w_sChikaisu%></td>
				<%Else%>
						<td class="<%=w_cell%>" align="center" width="30" >-</td>
						<td class="<%=w_cell%>" align="center" width="30" >-</td>
						<td class="<%=w_cell%>" width="45" align="center"><input type="text" <%=w_sInputClass%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="45" align="center"><input type="text" <%=w_sInputClass%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
						<td class="<%=w_cell%>" width="60" align="center"><input type="text" <%=w_sInputClass%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
			
				<%End If%>
			<%End If%>
					</tr>
			<%
			m_Rs.MoveNext
			Loop%>
		</table>
<!--
	</td>
	</tr>
	</table>
-->
	<table width="50%">
	<tr>
		<td align="center"><input type="button" class="button" value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
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

	</FORM>
	</center>
	</body>
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
		<span class="msg">�f�[�^�����݂��܂���B</span>
	</center>

	<input type="hidden" name="txtMsg" value="�f�[�^�����݂��܂���B">

	</form>
	</body>
	</html>

<%
End Sub
%>