<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���i���Ґ��ѓo�^
' ��۸���ID : sei/sei0900/sei0900_bottom.asp
' �@      �\: ���y�[�W ���i���Ґ��ѓo�^�̌������s��
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
' ��      ��: 2022/2/1 �g�c�@�Ď������ѓo�^��ʂ𗬗p���쐬
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�����I��p��Where����
    Public m_iNendo			'�N�x
    Public m_sKyokanCd		'�����R�[�h
    Public m_sSikenKBN		'�����敪
    Public m_sGakuNo		'�w�N
    Public m_sClassNo		'�w��
    Public m_sKamokuCd		'�ȖڃR�[�h
    Public m_sKamokuNM		'�Ȗږ� Ins 2017/12/26 Nishimura
    Public m_sSikenNm		'������
    Public m_rCnt			'���R�[�h�J�E���g
    Public m_sGakkaCd
    Public m_iSyubetu		'�o���l�W�v���@
    Public m_TUKU_FLG
	Public m_iJigenTani		'//�P�����̒P�ʐ�
    
    Public m_iKamoku_Kbn
    Public m_iHissen_Kbn
	Public m_ilevelFlg
	Public m_Rs
	Public m_DRs
	Public m_SRs
	
	Dim m_iSouJyugyou		'�����Ǝ���
	DIm m_iJunJyugyou		'�����Ǝ���

	Dim m_SeisekiIndex 'ADD 2022/1/12
	
	Public	m_iKikan		'���͊��ԃt���O
	Public	m_bKekkaNyuryokuFlg		'���ۓ��͉\�׸�(True:���͉� / False:���͕s��)
	
	Public	m_iShikenInsertType
	Public m_FirstGakusekiNo
	
	m_iShikenInsertType = 0
	
	Public m_sSyubetu
	Dim m_SchoolFlg

	Public Const C_GOUKAKUTEN = 60  '���i�_
	
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
'response.write "bottom START" & "<BR>"
	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���i���Ґ��ѓo�^"
	w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""
	
    On Error Resume Next
    Err.Clear
	
	m_bErrFlg = False
	
	Do
		'// �ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'//�s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'// ���Ұ�SET
		Call s_SetParam()
			
		'//�]���s�\�`�F�b�N�̏����́A�F�{�����̂��ߊw�Z�ԍ��̃`�F�b�N
		'//True�Ȃ�F�{�d�g
		if not gf_ChkDisp(C_DATAKBN_DISP,m_SchoolFlg) then
			m_bErrFlg = True
			Exit Do
		End If

		'//���ԃf�[�^�̎擾
		If f_Nyuryokudate() = 1 Then
			m_iKikan = "NO"
		else
			m_iKikan = ""
		End If

		'=================
		'//�o�����ۂ̎������擾
		'=================
		'//�Ȗڋ敪(0:������,1:�ݐ�)
		If gf_GetKanriInfo(m_iNendo,m_iSyubetu) <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If

		'=================
		'//���ʊ������擾
		'=================
		'//���ʊ���(0:�ʏ����,1:���ʊ���)
        If f_getTUKU(m_iNendo,m_sKamokuCd,m_sGakuNo,m_sClassNo,m_TUKU_FLG) <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
	
		'**********************************************************
		'�ʏ���Ƃ̂ݕ\���@�i���ʊ����͕\�����Ȃ��j
		'**********************************************************
		If m_TUKU_FLG = C_TUKU_FLG_TUJO then  '�ʏ���Ƃ̏ꍇ
			'=================
			'//�Ȗڏ����擾
			'=================
			'//�Ȗڋ敪(0:��ʉȖ�,1:���Ȗ�)�A�y�сA�K�C�I���敪(1:�K�C,2:�I��)�𒲂ׂ�
			'//���x���ʋ敪(0:��ʉȖ�,1:���x���ʉȖ�)�𒲂ׂ�
			If f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) <> 0 Then 
				m_bErrFlg = True
				Exit Do
			End If			

			'===============================
			'//���сA�w���f�[�^�擾
			'===============================
			'//�Ȗڋ敪��C_KAMOKU_SENMON(0:��ʉȖ�)�̏ꍇ�̓N���X�ʂɐ��k��\��
			'//�Ȗڋ敪��C_KAMOKU_SENMON(1:���Ȗ�)�̏ꍇ�͊w�ȕʂɐ��k��\��
			If f_getdate(m_iKamoku_Kbn) <> 0 Then m_bErrFlg = True : Exit Do
			If m_rs.EOF Then
				Call ShowPage_No()
				Exit Do
			End If
			
			'===============================
			'//���ې��̎擾
			'===============================
			' If f_GetSyukketu() <> 0 Then m_bErrFlg = True : Exit Do		
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
	m_sSikenKBN	= C_SIKEN_KOU_KIM
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sKamokuNM	= request("txtKamokuNM")	'Ins 2017/12/26 Nishimura
	m_sGakkaCd	= request("txtGakkaCd")
	m_iJigenTani = Session("JIKAN_TANI") '�P�����̒P�ʐ�
	m_sSyubetu	= trim(Request("SYUBETU"))
	
End Sub

'********************************************************************************
'*  [�@�\]  ���ې��A�x�������擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSyukketu()
	Dim w_iRet
	'Dim w_bZenkiOnly
	'Dim w_sSikenKBN
	Dim w_sTKyokanCd
	
    On Error Resume Next
    Err.Clear
	
    f_GetSyukketu = 1

	Do
		'==========================================
		'//�ȖڒS�������̋���CD�̎擾
		'==========================================
        w_iRet = f_GetTantoKyokan2(w_sTKyokanCd)
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
	
		'==========================================
		'//�����Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
		'==========================================
		'//�O���݂̂̏ꍇ��T21���O�L���������܂ł̌��ې����擾����
		'w_iRet = f_SikenInfo(w_bZenkiOnly)
		'If w_iRet<> 0 Then
		'	Exit Do
		'End If 
		
		'If w_bZenkiOnly = True Then
		'	w_sSikenKBN = C_SIKEN_ZEN_KIM
		'Else
		'	w_sSikenKBN = m_sSikenKBN
		'End If
		
		'//�ŏ��̐��k�̊w�Дԍ����擾
		if not m_Rs.EOF then
			m_FirstGakusekiNo = m_Rs("GAKUSEKI_NO")
			m_Rs.movefirst
		end if

		'==========================================
		'//�Ȗڂɑ΂��錋��,�x���̒l�擾
		'==========================================
		'if not gf_GetSyukketuData(m_SRs,w_sSikenKBN,m_sGakuNo,w_sTKyokanCd,m_sClassNo,m_sKamokuCd,w_skaisibi,w_sSyuryobi,"") then
		if not gf_GetSyukketuData2(m_SRs,m_sSikenKBN,m_sGakuNo,w_sTKyokanCd,m_sClassNo,m_sKamokuCd,w_skaisibi,w_sSyuryobi,"",m_iNendo,m_iShikenInsertType,m_FirstGakusekiNo,m_sSyubetu) then
			Exit Do
		end if
		
		'//����I��
	    f_GetSyukketu = 0
		Exit Do
	Loop

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

'    Do 		'2020.03.09 Del �s�v��Loop���폜

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

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_SikenInfo = 99
'            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_bZenkiOnly = True
		End If

        f_SikenInfo = 0
'        Exit Do
'    Loop

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

		w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_Z AS SEI_ZT,A.T16_KEKA_TYUKAN_Z AS KEKA_ZT,A.T16_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,A.T16_CHIKAI_TYUKAN_Z AS CHIKAI_ZT,A.T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI_ZT, "
		w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_Z AS SEI_ZK,A.T16_KEKA_KIMATU_Z AS KEKA_ZK,A.T16_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,A.T16_CHIKAI_KIMATU_Z AS CHIKAI_ZK,A.T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI_ZK, "
		w_sSQL = w_sSQL & " 	A.T16_SEI_TYUKAN_K AS SEI_KT,A.T16_KEKA_TYUKAN_K AS KEKA_KT,A.T16_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,A.T16_CHIKAI_TYUKAN_K AS CHIKAI_KT,A.T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI_KT, "
		w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI_KK,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
		w_sSQL = w_sSQL & " 	A.T16_SEI_KIMATU_K AS SEI,A.T16_KEKA_KIMATU_K AS KEKA,A.T16_KEKA_NASI_KIMATU_K AS KEKA_NASI,A.T16_CHIKAI_KIMATU_K AS CHIKAI,A.T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "
		w_sSQL = w_sSQL & "		A.T16_SOJIKAN_KIMATU_K as SOUJI, A.T16_JUNJIKAN_KIMATU_K as JYUNJI, A.T16_SAITEI_JIKAN, A.T16_KYUSAITEI_JIKAN, "
		
		w_sSQL = w_sSQL & "		A.T16_KOUSINBI_TYUKAN_Z,"
		w_sSQL = w_sSQL & "		A.T16_KOUSINBI_KIMATU_Z,"
		w_sSQL = w_sSQL & "		A.T16_KOUSINBI_TYUKAN_K,"
		w_sSQL = w_sSQL & "		A.T16_KOUSINBI_KIMATU_K,"
		w_sSQL = w_sSQL & "		A.T16_DATAKBN_KIMATU_K as DataKbn,"
		
		w_sSQL = w_sSQL & " 	A.T16_GAKUSEI_NO AS GAKUSEI_NO,A.T16_GAKUSEKI_NO AS GAKUSEKI_NO,B.T11_SIMEI AS SIMEI "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_SELECT_FLG "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_LEVEL_KYOUKAN "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_OKIKAE_FLG "
		w_sSQL = w_sSQL & vbCrLf & " ,A.T16_KAISETU  AS KAISETU "	'//�J�ݎ����ǉ� Ins 2018/03/22 Nishimura
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T16_NENDO = " & Cint(m_iNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T16_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	C.T13_GAKUNEN = " & Cint(m_sGakuNo) & " "
		w_sSQL = w_sSQL & " AND	C.T13_CLASS = " & Cint(m_sClassNo) & " "
		w_sSQL = w_sSQL & " AND	A.T16_NENDO = C.T13_NENDO "

		'//�u�����̐��k�͂͂���(C_TIKAN_KAMOKU_MOTO = 1    '�u����)
		w_sSQL = w_sSQL & " AND	A.T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO

		'//�K�C���I���Ȗڂ̂����I�����Ă���w���݂̂��擾����		'INS 2019/03/06 ����
		w_sSQL = w_sSQL & " AND	( T16_HISSEN_KBN = " & C_HISSEN_HIS
		w_sSQL = w_sSQL & "       OR (T16_HISSEN_KBN = " & C_HISSEN_SEN & " AND T16_SELECT_FLG = 1) "
		w_sSQL = w_sSQL & " 	) "

		w_sSQL = w_sSQL & " AND A.T16_SEI_KIMATU_K < " & C_GOUKAKUTEN & " "
		w_sSQL = w_sSQL & " AND T16_HYOKA_FUKA_KBN NOT IN(2 , 3) "
		w_sSQL = w_sSQL & " AND T16_HYOKA_FUKA_KBN NOT IN(" & C_HYOKA_FUKA_KEKKA &  "," & C_HYOKA_FUKA_BOTH & ") "
		w_sSQL = w_sSQL & " ORDER BY A.T16_GAKUSEKI_NO "

		'  response.write w_sSQL & "<BR>"
		'  response.end
		If gf_GetRecordset(m_Rs, w_sSQL) <> 0 Then
			'ں��޾�Ă̎擾���s
			f_getdate = 99
			m_bErrFlg = True
			Exit Do 
		End If
		
		m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
		m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))
		
		'//ں��ރJ�E���g�擾
		m_rCnt=gf_GetRsCount(m_Rs)
		' response.write "m_rCnt:" & m_rCnt & "<BR>"
		f_getdate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	�ȖڒS�������̋���CD�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetTantoKyokan(p_sTKyokanCd)
response.write "f_GetTantoKyokan START" & "<BR>"
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
response.write "w_sSQL:" & w_sSQL & "<BR>"
response.end
        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetTantoKyokan = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
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
' response.write "w_sSQL:" & w_sSQL & "<BR>"

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
	'm_bKekkaNyuryokuFlg = False		'���ۓ����׸�

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
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & C_SIKEN_SAISI
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakuNo)
' response.write w_sSQL  & "<BR>"
'rensponse.end		
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
			w_sSysDate = gf_SetNull2String(m_DRs("SYSDATE"))				'�V�X�e�����t
		End If
		
		'���͊��ԓ��Ȃ琳��
		If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			f_Nyuryokudate = 0
		End If
' response.write "m_iNKaishi:" & m_iNKaishi  & "<BR>"
' response.write "w_sSysDate:" & w_sSysDate  & "<BR>"
' response.write "m_iNSyuryo:" & m_iNSyuryo  & "<BR>"
' response.write "w_sSysDate:" & w_sSysDate  & "<BR>"		
' response.write "f_Nyuryokudate:" & f_Nyuryokudate  & "<BR>"
		
		Exit Do
	Loop

End Function


'********************************************************************************
'*	[�@�\]	�Ȗڂ����ʊ����ł��邩������
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

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_getTUKU = 99
			m_bErrFlg = True
			Exit Do 
		End If
'response.write w_sSQL  & "<BR>"
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

	On Error Resume Next
	Err.Clear

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
    Dim w_sKek,p_sChi
	Dim w_sSouG,w_sJyunG
	Dim w_Table,w_TableName
    Dim w_Kamoku
    
    On Error Resume Next
    Err.Clear
    
    p_iKekka = ""
    p_iChikoku = ""
	
	'���ʎ��ƁA���̑�(�ʏ�Ȃ�)�̐؂蕪��
	if trim(m_sSyubetu) = "TOKU" then
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_Kamoku = "T34_TOKUKATU_CD"
	else
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_Kamoku = "T16_KAMOKU_CD"
	end if
	
    f_GetKekaChi = 1
	
	'/�����敪�ɂ���Ď���Ă���A�t�B�[���h��ς���B
	Select Case p_iSikenKBN
		Case C_SIKEN_ZEN_TYU
			w_sKek   = w_Table & "_KEKA_TYUKAN_Z"
			w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_Z"
			p_sChi   = w_Table & "_CHIKAI_TYUKAN_Z"
			w_sSouG  = w_Table & "_SOJIKAN_TYUKAN_Z"
			w_sJyunG = w_Table & "_JUNJIKAN_TYUKAN_Z"
		Case C_SIKEN_ZEN_KIM
			w_sKek   = w_Table & "_KEKA_KIMATU_Z"
			w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_Z"
			p_sChi   = w_Table & "_CHIKAI_KIMATU_Z"
			w_sSouG  = w_Table & "_SOJIKAN_KIMATU_Z"
			w_sJyunG = w_Table & "_JUNJIKAN_KIMATU_Z"
		Case C_SIKEN_KOU_TYU
			w_sKek   = w_Table & "_KEKA_TYUKAN_K"
			w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_K"
			p_sChi   = w_Table & "_CHIKAI_TYUKAN_K"
			w_sSouG  = w_Table & "_SOJIKAN_TYUKAN_K"
			w_sJyunG = w_Table & "_JUNJIKAN_TYUKAN_K"
		Case C_SIKEN_KOU_KIM
			w_sKek   = w_Table & "_KEKA_KIMATU_K"
			w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_K"
			p_sChi   = w_Table & "_CHIKAI_KIMATU_K"
			w_sSouG  = w_Table & "_SOJIKAN_KIMATU_K"
			w_sJyunG = w_Table & "_JUNJIKAN_KIMATU_K"
	End Select
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & 	w_sKek   & " as KEKA, "
	w_sSQL = w_sSQL & 	w_sKekG  & " as KEKA_NASI, "
	w_sSQL = w_sSQL & 	p_sChi   & " as CHIKAI, "
	w_sSQL = w_sSQL & 	w_sSouG  & " as SOUJI, "
	w_sSQL = w_sSQL & 	w_sJyunG & " as JYUNJI "
	w_sSQL = w_sSQL & " FROM "   & w_TableName
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "      " & w_Table & "_NENDO =" & p_iNendo
	w_sSQL = w_sSQL & "  AND " & w_Table & "_GAKUSEI_NO= '" & p_sGakusei & "'"
	w_sSQL = w_sSQL & "  AND " & w_Kamoku & "= '" & p_sKamokuCD & "'"
	
	If gf_GetRecordset(w_KekaChiRs, w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		f_GetKekaChi = 99
	End If
	
	'//�߂�l���
	If w_KekaChiRs.EOF = False Then
		p_iKekka = gf_SetNull2String(w_KekaChiRs("KEKA"))
		p_iKekkaGai = gf_SetNull2String(w_KekaChiRs("KEKA_NASI"))
		p_iChikoku = gf_SetNull2String(w_KekaChiRs("CHIKAI"))
		
		m_iSouJyugyou = gf_SetNull2String(w_KekaChiRs("SOUJI"))
		m_iJunJyugyou = gf_SetNull2String(w_KekaChiRs("JYUNJI"))
	End If
	
	f_GetKekaChi = 0
	
	Call gf_closeObject(w_KekaChiRs)
	
End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
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
	Dim w_DataKbn
	Dim w_Checked

	Dim w_Padding
	Dim w_Padding2
	Dim w_Disabled
	Dim w_Disabled2
	Dim w_TableWidth
	
	on error resume next
	
	w_Padding = "style='padding:2px 0px;'"
	w_Padding2 = "style='padding:2px 0px;font-size:10px;'"

	w_lSeiTotal = 0
	w_lGakTotal = 0

	i = 1
	
	if m_SchoolFlg and m_TUKU_FLG = C_TUKU_FLG_TUJO then
		w_TableWidth = 760
		
		if cint(gf_SetNull2Zero(m_Rs("DataKbn"))) = cint(C_MIHYOKA) or m_iKikan = "NO" then
			w_Disabled = "disabled"
		end if
	else
		w_TableWidth = 710
	end if
	
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
			
		// �����ԂƏ����Ԃ�hidden�ɃZ�b�g
		document.frm.hidSouJyugyou.value = "<%= m_iSouJyugyou %>";
		document.frm.hidJunJyugyou.value = "<%= m_iJunJyugyou %>";
		
        //submit
        document.frm.target = "topFrame";
        document.frm.action = "sei0900_middle.asp"
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
		var ob,w_num;
		var i;
		var w_Seiseki;
		var w_hidSeiseki;
		var indx;
		var w_bFLG;
		var SeisekiArray =[];
			
		if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return;}
		//�w�b�_���󔒕\��
		parent.topFrame.document.location.href="white.asp";
		
		//�o�^����
	
		for (i = 0; i < document.frm.i_Max.value; i++) {
			w_Seiseki = eval("document.frm.Seiseki"+(i+1));
			indx = w_Seiseki.selectedIndex;
			m_SeisekiIndex =  w_Seiseki.options[indx].value;
			SeisekiArray[i] =  w_Seiseki.options[indx].value;
		}
		document.frm.hidSeiseki.value = SeisekiArray;
		document.frm.hidUpdMode.value = "TUJO";
		document.frm.action="sei0900_upd.asp";
		document.frm.target="main";
		document.frm.submit();
	
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


	//-->
	</SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<center>
	<form name="frm" method="post">
	
	<table width="<%=w_TableWidth%>">
	<tr>
	<td>
	
	<table class="hyo" align="center" width="<%=w_TableWidth%>" border="1">
	
	<%	m_Rs.MoveFirst
		Do Until m_Rs.EOF
			w_ihalf = gf_Round(m_rCnt / 2,0)
			j = j + 1 
			w_sSeiseki = ""
			w_sHyoka = ""
			w_sKekka = ""
			w_sChikai = ""
			w_sGakusekiCd = ""
			w_sKekkasu = ""
			w_sChikaisu = ""
			Call gs_cellPtn(w_cell)
%>
	<tr>
<%
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
			
			w_sKekka = gf_SetNull2String(m_Rs("KEKA"))
			w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI"))
			w_sChikai = gf_SetNull2String(m_Rs("CHIKAI"))

			'//�O���ŏI����Ă���Ȗڂ̌��ۂ��擾���Ċw�������тɃZ�b�g����B2002/02/21 ITO
			'//�O���݂̂̏ꍇ��T21���O�L���������܂ł̌��ې����擾����
			
			'w_iRet = f_SikenInfo(w_bZenkiOnly)
			If cint(m_Rs("KAISETU")) = C_KAI_ZENKI Then	'2020.03.09 Upd Kiyomoto �J�ݎ�����T16���R�[�h����擾����
				w_bZenkiOnly = True
			End if 
			
			'�w�N�������̏ꍇ�̂�
			If m_sSikenKBN = C_SIKEN_KOU_KIM Then
				
				'�O���J�݂�������O�������̌��ۂ��w�N���̐��тɃZ�b�g����
				If w_bZenkiOnly = True Then

					'�w�������т��X�V����Ă��Ȃ��ꍇ�A�O���̌��ہA�x�����w�N���ɃZ�b�g����B
					If gf_SetNull2String(m_Rs("T16_KOUSINBI_KIMATU_K")) = "" Then 
						w_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))			'���ې�
						w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK"))	'���ۑΏۊO
						w_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))		'�x����
					End If

					'�w�������т�0
'					If gf_SetNull2String(m_Rs("KEKA")) = "" Then 
'						w_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))			'���ې�
'						w_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK"))	'���ۑΏۊO
'						w_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))		'�x����
'					End If
					
				End If
			End If
			
			'�O���ŏI����Ă���Ȗڂ̌��ۂ��擾���Ċw�������тɃZ�b�g����B2002/02/21 ITO
			'�l�̏������B
			w_bNoChange = False
			w_sKekkasu = ""
			w_sChikaisu = ""
			
			'---------------------------------------------------------------------------------------------
			'�ʏ���ƂƂ��̏���
			if m_TUKU_FLG = C_TUKU_FLG_TUJO then 
				w_sSeiseki = gf_SetNull2String(m_Rs("SEI"))
				w_sHyoka = gf_HTMLTableSTR(m_Rs("HYOKAYOTEI"))
				
				'�O���ŏI����Ă���Ȗڂ̌��ۂ��擾���Ċw�������тɃZ�b�g����B2002/02/21 ITO
				
				'�w�N�������̏ꍇ�̂�
				If m_sSikenKBN = C_SIKEN_KOU_KIM Then
					
					'�O���J�݂�������O�������̌��ۂ��w�N���̐��тɃZ�b�g����
					If w_bZenkiOnly = True Then
						'�w�������т��X�V����Ă��Ȃ��ꍇ�A�O���̐��т��w�N���ɃZ�b�g����B
						If gf_SetNull2String(m_Rs("T16_KOUSINBI_KIMATU_K")) = "" Then 
							w_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))			'�O����������
						End If
						'�w�������т�0
'						If gf_SetNull2String(m_Rs("SEI")) = "" Then 
'							w_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))			'�O����������
'						End If
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
							w_bNoChange = True
						else
							if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
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
			
			w_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")
			'//C_IDO_FUKUGAKU=3:���w�AC_IDO_TEI_KAIJO=5:��w����
			w_SSSS = ""
			w_SSSR = ""
			
			w_SSSS = gf_Get_IdouChk(w_sGakusekiCd,w_Date,m_iNendo,w_SSSR)
			
			IF CStr(w_SSSS) <> "" Then
				IF Cstr(w_SSSS) <> CStr(C_IDO_FUKUGAKU) AND Cstr(w_SSSS) <> Cstr(C_IDO_TEI_KAIJO) AND Cstr(w_SSSS) <> Cstr(C_IDO_TENKO) _
					AND Cstr(w_SSSS) <> Cstr(C_IDO_TENKA) AND Cstr(w_SSSS) <> Cstr(C_IDO_KOKUHI) AND Cstr(w_SSSS) <> Cstr(C_IDO_NYUGAKU) _
					AND  Cstr(w_SSSS) <> Cstr(C_IDO_TENNYU) Then
					w_SSSS = "[" & w_SSSR & "]"
					w_bNoChange = True
				Else
					w_SSSR = ""
					w_SSSS = ""
				End if
			End if
			
			'�ʏ����
			if Cstr(m_TUKU_FLG) = Cstr(C_TUKU_FLG_TUJO) then 
				
				'���ےx�����̎擾
				
				'//���ې��~�P�ʐ��̎擾
				w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_KEKKA)		'�߂�l��NULL�̎���""
				
				'gf_IIF�ɓn���Ƃ��Ƀp�����[�^���v�Z����̂ŁA�p�����[�^��0�ɕϊ�
				w_sKekkasu = gf_IIF(w_sData = "", "", cint(gf_SetNull2Zero(w_sData)) * cint(m_iJigenTani))

'response.write(w_sGakusekiCd & ":" & " T21 ���� w_sKekkasu = " & w_sKekkasu & "<BR>")
				
				'//�P���ۂ̏ꍇ�̌��ې��̎擾
				w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_KEKKA_1)
				if w_sKekkasu = "" and w_sData = "" then
					w_sKekkasu = ""
				else
					'�ǂ��炩���""�łȂ���Όv�Z
					w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_sData))			'//�P���ۂ̏ꍇ�̌��ې��̎擾
				end if

'response.write(w_sGakusekiCd & ":" & " T21 1���� w_sKekkasu = " & w_sKekkasu & "<BR>")
				
				'//�x�����̎擾
				w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_TIKOKU)
				w_sChikaisu = gf_IIF(w_sData = "", "", cint(gf_SetNull2Zero(w_sData)))
				
				'//���ސ��̎擾
				w_sData=f_Syukketu2New(w_sGakusekiCd,C_KETU_SOTAI)
				if w_sChikaisu = "" and w_sData = "" then
					'w_sKekkasu��w_sData���ǂ����""�̎���""�̂܂�
					w_sChikaisu = ""
				else
					'�ǂ��炩���""�łȂ���Όv�Z
					w_sChikaisu = cint(gf_SetNull2Zero(w_sChikaisu)) + cint(gf_SetNull2Zero(w_sData))			'//�P���ۂ̏ꍇ�̌��ې��̎擾
				end if
			end if
			
			'---------------------------------------------------------------------------------------------
			
			'�u�o�����ۂ��ݐρv�Łu�O�����ԂłȂ��v�̏ꍇ
			'���ہE���Ȃ�Null�������ꍇ�A�����邽�ߊ֐��ǉ� Add 2001.12.16 okada
			if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and w_sShikenKBN_RUI <> 99 then 
				'��O�̎����̍��v�l�𑫂��B
				call f_GetKekaChi(m_iNendo,m_iShikenInsertType,m_sKamokuCd,cstr(m_Rs("GAKUSEI_NO")),w_iKekka_rui,w_iChikoku_rui,w_iKekkaGai_rui)
				
				'�ǂ����""�̎���""
				if w_sKekkasu = "" and w_iKekka_rui = "" then
					w_sKekkasu = ""
				else
					w_sKekkasu = cint(gf_SetNull2Zero(w_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))
				end if

'response.write(w_sGakusekiCd & ":" & " T16 w_sKekkasu = " & w_sKekkasu & "<BR><BR>")
				
				'�ǂ����""�̎���""
				if w_sChikaisu = "" and w_iChikoku_rui = "" then
					w_sChikaisu = ""
				else
					w_sChikaisu = cint(gf_SetNull2Zero(w_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
				end if
			end if
			
			'If cint(gf_SetNull2Zero(w_sKekka)) = 0 and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then 		'//������0��,���v��0���傫���ꍇ
			If gf_SetNull2String(w_sKekka) = "" and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then 		'//������0��,���v��0���傫���ꍇ
				w_sKekka = cint(gf_SetNull2Zero(w_sKekkasu))								'//���������v
			End If
			
			'If cint(gf_SetNull2Zero(w_sChikai)) = 0 AND cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then		'//�x����0��,�x�v��0���傫���ꍇ
			If gf_SetNull2String(w_sChikai) = "" and cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then		'//�x����0��,�x�v��0���傫���ꍇ
				w_sChikai = cint(gf_SetNull2Zero(w_sChikaisu))							'//�x�����x�v
			End If
			
			
			if m_SchoolFlg and m_TUKU_FLG = C_TUKU_FLG_TUJO then
				'//�]���s�\�f�[�^�ݒ�
				w_DataKbn = 0
				w_Checked = ""
				w_Disabled2 = ""
				
				w_DataKbn = cint(gf_SetNull2Zero(m_Rs("DataKbn")))
				
				if w_DataKbn = cint(C_HYOKA_FUNO) then
					w_Checked = "checked"
					w_Disabled2 = "disabled"
					
				elseif w_DataKbn = cint(C_MIHYOKA) then
					w_Disabled2 = "disabled"
					
				end if
				
				select case Cstr(w_SSSS)
					case Cstr(C_IDO_KYU_BYOKI),Cstr(C_IDO_KYU_HOKA)
						w_DataKbn = C_KYUGAKU
						
					case Cstr(C_IDO_TAI_2NEN),Cstr(C_IDO_TAI_HOKA),Cstr(C_IDO_TAI_SYURYO)
						w_DataKbn = C_TAIGAKU
				end select
			end if
			
			'========================================================================================
			'//�Ȗڂ��I���Ȗڂ̎��ɉȖڂ�I�����Ă��Ȃ��ꍇ(���͕s��)
			'========================================================================================
			If w_bNoChange = True Then
				if Cstr(m_TUKU_FLG) = Cstr(C_TUKU_FLG_TUJO) Then%>

					<input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>">
					<input type="hidden" name="txtKaisetu<%=i%>" value="<%=m_Rs("KAISETU")%>">
					<input type="hidden" name="hidUpdFlg<%=i%>" value="False">
					<td class="<%=w_cell%>" align="center" width="65"  nowrap <%=w_Padding%>><%=w_sGakusekiCd%></td>
					<td class="<%=w_cell%>" align="left"   width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
				<%Else%>

					<input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>">
					<input type="hidden" name="txtKaisetu<%=i%>" value="<%=m_Rs("KAISETU")%>">
					<input type="hidden" name="hidUpdFlg<%=i%>" value="False">
					<td class="<%=w_cell%>" align="center" width="65"  nowrap <%=w_Padding%>><%=w_sGakusekiCd%></td>
					<td class="<%=w_cell%>" align="left"   width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="50"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="100" nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="80"  nowrap <%=w_Padding%>>-</td>
					<td class="<%=w_cell%>" align="center" width="85"  nowrap <%=w_Padding%>>-</td>
				<%End if%>
			<%
			'=========================================================================
			'//�Ȗڂ��K�C���A�܂��͑I���Ȗڂ̎��ɐ��k���Ȗڂ�I�����Ă���ꍇ(���͉�)
			'=========================================================================
			Else
				%>
					<td class="<%=w_cell%>" align="center" width="65" nowrap <%=w_Padding%>>
						<%=w_sGakusekiCd%>
						<input type="hidden" name="txtGseiNo<%=i%>" value="<%=m_Rs("GAKUSEI_NO")%>">
						<input type="hidden" name="txtKaisetu<%=i%>" value="<%=m_Rs("KAISETU")%>">
					</td>
						
					<input type="hidden" name="hidUpdFlg<%=i%>" value="True">
					<td class="<%=w_cell%>" align="left"  width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_SSSS%></td>
					
					<!-- 2002.03.20 -->
					<!-- �R�����g
					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI1"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI2"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI3"))%></td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding2%>><%=gf_HTMLTableSTR(m_Rs("SEI4"))%></td>
					<%Else%>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>&nbsp;&nbsp;</td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>&nbsp;&nbsp;</td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>&nbsp;&nbsp;</td>
						<td class="<%=w_cell%>" align="center" width="30" nowrap <%=w_Padding%>>&nbsp;&nbsp;</td>
					<%End If%>
					 -->
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
					
					'=========================================================================
					'//�ʏ���Ƃ̏ꍇ 
					'=========================================================================
					%>
				<%If m_iKikan <> "NO" Then%>
					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
							
						<td class="<%=w_cell%>" width="50"align="center" nowrap <%=w_Padding%>>
							<select name="Seiseki<%=i%>" style='width:50px;' onchange="">
								<option value="1">��</option>
								<option value="2" selected="">��</option>
							</select>
						<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>
								<td class="<%=w_cell%>"  width="50" align="center" nowrap <%=w_Padding%>>
									<input type="button" size="2" name="button<%=i%>" value="<%=w_sHyoka%>" onClick="return f_change(<%=i%>)" class="<%=w_cell%>" style="text-align:center">
								</td>
								<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>">
						<%Else%>
								<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>><%=w_sHyoka%></td>
								<input type="hidden" name="Hyoka<%=i%>" value="<%=trim(w_sHyoka)%>">
						<%End If%>
					<%Else%>
							<td class="<%=w_cell%>" width="50"  nowrap align="center" <%=w_Padding%>>-</td>
							<td class="<%=w_cell%>" width="50"  nowrap align="center" <%=w_Padding%>>-</td>
							<td class="<%=w_cell%>" width="100" nowrap align="center" <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="80"  nowrap align="center" <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)"></td>
							<td class="<%=w_cell%>" width="85"  nowrap align="center" <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)"></td>
					<%End If%>
					
				<%Else%>
					
					<%If m_TUKU_FLG = C_TUKU_FLG_TUJO Then%>
							
						<td class="<%=w_cell%>" width="50" align="right" nowrap <%=w_Padding%>>
							<select name="Seiseki<%=i%>" style='width:50px;' onchange="">
									<option value="1">��</option>
									<option value="2" selected="">��</option>
							</select>
						</td>
						
						<%	'�\���݂̂̏ꍇ�̍��v�E���ϒl�����߂�
							If IsNull(w_sSeiseki) = False Then
								If IsNumeric(CStr(w_sSeiseki)) = True Then
									w_lSeiTotal = w_lSeiTotal + CLng(w_sSeiseki)
									w_lGakTotal = w_lGakTotal + 1
								End If
							End If
						%>
						<%If m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU Then%>
								<td class="<%=w_cell%>"  width="50" align="center" nowrap <%=w_Padding%>><%=trim(w_sHyoka)%></td>
						<%Else%>
								<td class="<%=w_cell%>"  width="50" align="center" nowrap <%=w_Padding%>><%=trim(w_sHyoka)%></td>
						<%End If%>
					<%Else%>
					<%End If%>
				<%End If%>
			<%End If%>
				
				<% if m_SchoolFlg and m_TUKU_FLG = C_TUKU_FLG_TUJO then %>
					
					<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>>
						<% if w_DataKbn = C_HYOKA_FUNO or w_DataKbn = C_MIHYOKA or w_DataKbn = 0 then %>
							<input type="checkbox" name="chkHyokaFuno<%=i%>" value="3" <%=w_Disabled%> <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);">
						<% else %>
							&nbsp;
							<input type="hidden" name="chkHyokaFuno<%=i%>" value="<%=w_DataKbn%>">
						<% end if %>
					</td>
					
				<% end if %>
			</tr>
			
			<%
				m_Rs.MoveNext
				i = i + 1
			Loop
			%>
		</table>
		
		</td>
		</tr>
		
		<tr>
		<td align="center">
		<table>
			<tr>
				<td align="center" align="center" colspan="13">
					<%If m_iKikan <> "NO" Then%>
						<input type="button" class="button" value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
					<%End If%>
						<input type="button" class="button" value="�L�����Z��" onclick="javascript:f_Cansel()">
				</td>
			</tr>
		</table>
		</td>
		</tr>
	</table>
	
	<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="KamokuCd"    value="<%=m_sKamokuCd%>">
	<input type="hidden" name="i_Max"       value="<%=i-1%>">
	<input type="hidden" name="txtSikenKBN" value="<%=m_sSikenKBN%>">
	<input type="hidden" name="txtGakuNo"   value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakkaCd"  value="<%=m_sGakkaCd%>">
	<input type="hidden" name="txtClassNo"  value="<%=m_sClassNo%>">
	<input type="hidden" name="txtKamokuCd" value="<%=m_sKamokuCd%>">
	<input type="hidden" name="txtKamokuNM" value="<%=m_sKamokuNM%>">
	<input type="hidden" name="txtTUKU_FLG" value="<%=m_TUKU_FLG%>">
	<input type="hidden" name="PasteType"   value="">
	<!-- 02/03/27 �ǉ� -->
	<input type="hidden" name="hidSouJyugyou">
	<input type="hidden" name="hidJunJyugyou">
	<input type="hidden" name="hidUpdMode">
	
	<input type="hidden" name="hidFirstGakusekiNo" value="<%=m_FirstGakusekiNo%>">
	<input type="hidden" name="hidMihyoka" value ="<%=w_DataKbn%>">
	<input type="hidden" name="hidSchoolFlg" value ="<%=m_SchoolFlg%>">
	<input type="hidden" name="hidSeiseki" value="">
	
	</FORM>
	</center>
	</body>
	<SCRIPT>
	<!--
	//-->
	</SCRIPT>

	</html>
<%
End sub

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
	//-->
	</SCRIPT>
	</head>
	
    <body LANGUAGE="javascript">
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