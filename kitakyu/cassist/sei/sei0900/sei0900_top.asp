<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���i���Ґ��ѓo�^
' ��۸���ID : sei/sei0900/sei0900_top.asp
' �@      �\: ��y�[�W ���i���Ґ��ѓo�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
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
    Public  m_bErrMsg           '�װү����

    Public m_iNendo             '�N�x
	Public m_iRisyuKakoNendo    '���C�ߋ��N�x
    Public m_iGakki             '�w��
    Public m_sKyokanCd          '�����R�[�h
    Public m_iSikenKbn			'�����敪
	Public m_sTxtMode           '//���샂�[�h

    Public m_iDispFlg			'�X�V���\���t���O 0:�\���A1:��\��

	Public m_sGetTable			'�ȖڃR���{���쐬�����e�[�u��
    
    Public m_Rs_Nendo			'�N�x�����擾
    Public m_Rs					'�w�N�A�N���X�A�Ȗڎ擾RS
	Public m_Rs_NendoCount			'�N�x���̌���
	Public m_RsCnt			'���R�[�h�J�E���g �Ȗڎ擾RS

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
    w_sMsgTitle="���i���Ґ��ѓo�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"

    On Error Resume Next
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
		'//�l���擾
		call s_SetParam()

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

		' �w�N�������̎����敪���擾
		m_iSikenKbn =  C_SIKEN_KOU_KIM
		'//�N�x�R���{���擾
        w_iRet = f_GetRisyuKakoNendo()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//�N�x��NULL��������A�R���{�̍Ō�̔N�x������
		If  gf_IsNull(m_iRisyuKakoNendo) Then
	        m_Rs_Nendo.MoveLast
			m_iRisyuKakoNendo  = m_Rs_Nendo("T17_NENDO")
			m_Rs_Nendo.MoveFirst
    	End If

		if Not gf_IsNull(m_iRisyuKakoNendo) then

			'//�ȖڃR���{���擾
			w_iRet = f_GetKamoku_Nenmatu()
			If w_iRet <> 0 Then m_bErrFlg = True : Exit Do	

		End if

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
    Call gf_closeObject(m_Rs_Nendo)
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

    m_iNendo    = session("NENDO")
    m_iGakki    = Session("GAKKI")
    m_sKyokanCd = session("KYOKAN_CD")
	m_sTxtMode  = Request("txtMode")
	m_iRisyuKakoNendo  = Request("txtRisyuKakoNendo")
	

End Sub

'********************************************************************************
'*  [�@�\]  �N�x�R���{���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetRisyuKakoNendo()

    Dim w_sSQL
	Dim w_bRtn
	Dim w_oRecord
	Dim w_Nendo
	Dim w_Count

    On Error Resume Next
    Err.Clear
    
    f_GetRisyuKakoNendo = 1



		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & "  DISTINCT T17_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  TT17_RISYUKAKO_KOJIN"
		w_sSQL = w_sSQL & vbCrLf & "  ,TT27_TANTO_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE "
 		w_sSQL = w_sSQL & vbCrLf & "  T17_NENDO = T27_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_KAMOKU_CD = T17_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_GAKUNEN = T17_HAITOGAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO 
		w_sSQL = w_sSQL & vbCrLf & "    AND (T17_TANI_SUMI =NULL OR T17_TANI_SUMI = 0) "
		w_sSQL = w_sSQL & vbCrLf & "  ORDER BY T17_NENDO"
' response.write "w_sSQL:" & w_sSQL & "<BR>"
' response.end
        ' w_bRtn = gf_GetRecordset(m_Rs_Nendo, w_sSQL)
		w_bRtn = gf_GetRecordset_OpenStatic(m_Rs_Nendo, w_sSQL)

        If w_bRtn <> 0 Then
             Exit Function
        End If
			
        f_GetRisyuKakoNendo = 0


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
		w_sSQL = w_sSQL & vbCrLf & "  WHERE T26_NENDO = " & m_iNendo

If m_iSikenKbn < C_SIKEN_KOU_KIM then '�N�x�������̏ꍇ�́A���ׂĂ��Ώ�
		w_sSQL = w_sSQL & vbCrLf & "    AND T26_SIKEN_KBN =" & m_iSikenKbn
End If

		w_sSQL = w_sSQL & vbCrLf & "    AND T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "'"
		w_sSQL = w_sSQL & vbCrLf & "    AND ("
		w_sSQL = w_sSQL & vbCrLf & "       T26_JISSI_KYOKAN    ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN1 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN2 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN3 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN4 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    OR T26_SEISEKI_KYOKAN5 ='" & m_sKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    )"
		'w_sSQL = w_sSQL & vbCrLf & "    AND T26_JISSI_FLG = " & Cint(C_SIKEN_KBN_JISSI)
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
'*  [�@�\]  ��������̎��A�ȖڃR���{���擾
'*          ���̔N�x�Ɏ��{���ꂽ������S�ĕ\������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKamoku_Nenmatu()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetKamoku_Nenmatu = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
		w_sSQL = w_sSQL & vbCrLf & " 	T27_KAMOKU_CD AS KAMOKU"
		w_sSQL = w_sSQL & vbCrLf & " 	,MAX(T17_KAMOKUMEI) AS KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM"
		w_sSQL = w_sSQL & vbCrLf & " 	T27_TANTO_KYOKAN "
		w_sSQL = w_sSQL & vbCrLf & " 	,T17_RISYUKAKO_KOJIN "
		w_sSQL = w_sSQL & vbCrLf & " 	,M05_CLASS "
		w_sSQL = w_sSQL & vbCrLf & "	,("
		w_sSQL = w_sSQL & vbCrLf & " 		SELECT * FROM TT13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " 		WHERE  T13_NENDO = " & cInt(m_iRisyuKakoNendo) - 1
		w_sSQL = w_sSQL & vbCrLf & " 		 AND T13_KARI_SINKYU = 1) TT13_GAKU_NEN "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " 		T27_NENDO = M05_NENDO "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_GAKUNEN = M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_CLASS = M05_CLASSNO	"
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_KAMOKU_CD = T17_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_GAKUNEN = T17_HAITOGAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_NENDO = T27_NENDO "
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_GAKUSEI_NO = T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_NENDO = " & cInt(m_iRisyuKakoNendo)
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
		w_sSQL = w_sSQL & vbCrLf & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "
		w_sSQL = w_sSQL & vbCrLf & "    AND (T17_TANI_SUMI =NULL OR T17_TANI_SUMI = 0) " & " "
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO 
		w_sSQL = w_sSQL & vbCrLf & "    AND T17_COURSE_CD IN ( '0' , CASE WHEN M05_GAKKA_CD = T17_GAKKA_CD THEN (CASE WHEN M05_COURSE_CD IS NOT NULL THEN M05_COURSE_CD ELSE T17_COURSE_CD END ) ELSE T17_COURSE_CD END ) " '2019.02.12 Upd Kiyomoto
		w_sSQL = w_sSQL & vbCrLf & "  GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & " 	T27_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY "
		w_sSQL = w_sSQL & vbCrLf & "  KAMOKU"

' response.write w_sSQL  & "<BR>"
' rensponse.end

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
		' iRet =gf_GetRecordset_OpenStatic(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKamoku_Nenmatu = 99
            Exit Do
        End If	
		'//ں��ރJ�E���g�擾
		m_RsCnt=gf_GetRsCount(m_Rs)

        f_GetKamoku_Nenmatu = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  ���C�e�[�u�����Ȗږ��̂��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_KamokuName
'*  [����]  
'********************************************************************************
Function f_GetKamokuName(p_Gakunen,p_Class,p_KamokuCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_GakkaCd
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_GetKamokuName = ""
	p_KamokuName = ""

    Do 

		'//�����s���̂Ƃ�
		If trim(p_Gakunen)="" Or trim(p_Class) = "" Or  trim(p_KamokuCd) = "" Then
            Exit Do
		End If

		'//�w��CD���擾
		w_iRet = f_GetGakkaCd(p_Gakunen,p_Class,w_GakkaCd)
		If w_iRet<> 0 Then
            Exit Do
		End If

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,T15_RISYU.T15_LEVEL_FLG"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(p_Gakunen) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & w_GakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD=" & p_KamokuCd

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		If w_Rs.EOF = False Then
			p_KamokuName = w_Rs("T15_KAMOKUMEI")
		End If

        Exit Do
    Loop

    f_GetKamokuName = p_KamokuName

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  ���x���ʂ��ǂ����𒲂ׂ�B
'*  [����]  �Ȃ�
'*  [�ߒl]  ���x���ʁFtrue
'*  [����]  
'********************************************************************************
Function f_LevelChk(p_Gakunen,p_KamokuCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_GakkaCd
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_LevelChk = false
	p_KamokuName = ""
    Do 

		'//�����s���̂Ƃ�
		If trim(p_Gakunen)="" Or  trim(p_KamokuCd) = "" Then
            Exit Do
		End If

		'//�w��CD���擾
'		w_iRet = f_GetGakkaCd(p_Gakunen,p_Class,w_GakkaCd)
		If w_iRet<> 0 Then
            Exit Do
		End If

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(T15_LEVEL_FLG) "
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_NYUNENDO = " & cint(m_iNendo) - cint(p_Gakunen) + 1
'		w_sSQL = w_sSQL & vbCrLf & "  AND T15_GAKKA_CD='" & w_GakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_KAMOKU_CD = '" & p_KamokuCd & "'"


        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		If w_Rs.EOF = False and cint(gf_SetNull2Zero(w_Rs("MAX(T15_LEVEL_FLG)"))) = 1 Then
			f_LevelChk = true
		End If

        Exit Do
    Loop
    Call gf_closeObject(w_Rs)
End Function

'********************************************************************************
'*  [�@�\]  �w��CD���擾
'*  [����]  p_Gakunen:�w�N,p_Class:�N���X
'*  [�ߒl]  p_GakkaCd:�w��CD
'*  [����]  
'********************************************************************************
Function f_GetGakkaCd(p_Gakunen,p_Class,p_GakkaCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetGakkaCd = 1
	p_GakkaCd = ""

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKKA_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO= " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN=" & cint(p_Gakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_CLASSNO=" & cint(p_Class)

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetGakkaCd = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_GakkaCd = w_Rs("M05_GAKKA_CD")
		End If

        f_GetGakkaCd = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'****************************************************
'[�@�\] �u��������ꂽ�I���Ȗڂ͕\�����Ȃ����߂̊֐��B
'[����] 
'       
'[�ߒl] 
'****************************************************
Function f_chkOkikae(p_KamokuCd)
	Dim w_sSql
	Dim w_Rs
	Dim i_Ret

	On Error Resume Next
    Err.Clear

	f_chkOkikae = 1

Do

	w_sSql = "Select "
	w_sSql = w_sSql & "T65_KAMOKU_CD_SAKI "
	w_sSql = w_sSql & "From "
	w_sSql = w_sSql & "T65_RISYU_SENOKIKAE "
	w_sSql = w_sSql & "Where "
	w_sSql = w_sSql & "T65_KAMOKU_CD_SAKI = '" & p_KamokuCd & "'"

	iRet = gf_GetRecordset(w_Rs, w_sSql)
	If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_chkOkikae = 99
            Exit Do
    End If

	If w_Rs.EOF = False Then
		f_chkOkikae = 0
	End If
	
    Exit Do
Loop
	Call gf_closeObject(w_Rs)

End Function

'****************************************************
'[�@�\] �u��������ꂽ��։Ȗڂ�\�����邽�߂̊֐��B�i���w���p�j
'[����] 
'       
'[�ߒl] 
'****************************************************
Function f_chkRyuOkikae(p_KamokuCd)
	Dim w_sSql
	Dim w_Rs
	Dim i_Ret

	On Error Resume Next
    Err.Clear

	f_chkRyuOkikae = 1

Do

	w_sSql = ""
	w_sSql = "Select "
	w_sSql = w_sSql & "T75_KAMOKU_CD_SAKI "
	w_sSql = w_sSql & "From "
	w_sSql = w_sSql & "T75_RYU_OKIKAE "
	w_sSql = w_sSql & "Where "
	w_sSql = w_sSql & "T75_KAMOKU_CD_SAKI = '" & p_KamokuCd & "'"
	w_sSql = w_sSql & "And "
	w_sSql = w_sSql & "T75_NENDO = " & m_iNendo

	iRet = gf_GetRecordset(w_Rs, w_sSql)
	If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_chkRyuOkikae = 99
            Exit Do
    End If

	If w_Rs.EOF = False Then
		f_chkRyuOkikae = 0
	End If
	
    Exit Do
Loop
	Call gf_closeObject(w_Rs)

End Function


'****************************************************
'[�@�\] �f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'[����] pData1 : �f�[�^�P
'       pData2 : �f�[�^�Q
'[�ߒl] f_Selected : "SELECTED" OR ""
'****************************************************
Function f_Selected(pData1,pData2)

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else 
            f_Selected = "" 
        End If
    End If

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
	
	If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		f_getTUKU = 99
		m_bErrFlg = True
	End If
	
	If w_Rs.EOF = false Then
		p_TUKU_FLG = cStr(gf_SetNull2Zero(w_Rs("T20_TUKU_FLG")))
	End If
	
    Call gf_closeObject(w_Rs)
	
End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	Dim i
	On Error Resume Next
    Err.Clear
	i = 1
%>
	<html>
	<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//  [�@�\]  �������ύX���ꂽ�Ƃ��A�ĕ\������
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_ReLoadMyPage(){

	    document.frm.action="sei0900_top.asp";
	    document.frm.target="topFrame";
	    document.frm.submit();

	}

	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���̏���
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_Search(){

	    // ������NULL����������
	    // ���N�x
	    if( f_Trim(document.frm.txtRisyuKakoNendo.value) == "<%=C_CBO_NULL%>" ){
	        window.alert("�N�x�̑I�����s���Ă�������");
	        document.frm.txtRisyuKakoNendo.focus();
	        return ;
	    }

	    // ���Ȗږ�
	    if( f_Trim(document.frm.txtKamokuCd.value) == "<%=C_CBO_NULL%>" ){

			if (document.frm.txtKamokuCd.length ==1){
		        window.alert("�����Ȗڂ�����܂���");
		        return ;
			}else{
		        window.alert("�Ȗڂ̑I�����s���Ă�������");
		        document.frm.txtKamokuCd.focus();
		        return ;
			}
	    }

		// �I�����ꂽ�R���{�̒l���
		iRet = f_SetData();
		if( iRet != 0 ){
	        window.alert("�Ȗڂ̃f�[�^������܂���");
			return;
		}

	    document.frm.action="sei0900_bottom.asp";
	    document.frm.target="main";
	    document.frm.submit();

	}

	//************************************************************
	//  [�@�\]  �N�x�R���{Change�C�x���g
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
			
		function f_changeKamoku(){
	    
		// �I�����ꂽ�R���{�̒l���
		iRet = f_GetKamokuCombo();
		if( iRet != 0 ){
	        window.alert("�N�x�̑I�����s���Ă�������");
			return;
		}
	
		document.frm.action="sei0900_top.asp";
	    document.frm.target="topFrame";
		 document.frm.txtMode.value = "Reload";
	    document.frm.submit();

	}

	//************************************************************
	//  [�@�\]  �N�x�R���{�̃f�[�^����Ȗڂ��擾���鏈��
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_GetKamokuCombo(){

		if (document.frm.cboRisyuKakoNendo.value==""){
			return 1;
        };
		
		//�f�[�^�擾
        m_iRisyuKakoNendo = document.frm.cboRisyuKakoNendo.value;
		document.frm.txtRisyuKakoNendo.value=m_iRisyuKakoNendo;
		// document.frm.cboRisyuKakoNendo.selected= true;

        return 0;
	}
	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���ɑI�����ꂽ�f�[�^���
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_SetData(){
		if (document.frm.cboKamoku.value==""){
			return 1;
        };
		if (document.frm.cboRisyuKakoNendo.value==""){
			return 1;
        };

		//�f�[�^�擾
        var vl = document.frm.cboKamoku.value.split('#@#');
		
        //�I�����ꂽ�f�[�^���(�Ȗ�CD���擾)
        document.frm.txtKamokuCd.value=vl[0];
        document.frm.txtKamokuNM.value=vl[1];
		
		m_iRisyuKakoNendo = document.frm.cboRisyuKakoNendo.value;
		document.frm.txtRisyuKakoNendo.value=m_iRisyuKakoNendo;

        return 0;
	}

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

		// �I�����ꂽ�R���{�̒l���
		iRet = f_SetData();
		if( iRet != 0 ){
			return;
		}
		
    }

	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>

    <body LANGUAGE=javascript onload="return window_onload()">
	
	<center>
	<form name="frm" METHOD="post">

	<% 
		Dim w_iGakunen_s
		Dim w_sGakkaCd_s
		Dim w_sKamokuCd_s
		Dim w_sKamokuNM_s

		call gs_title(" ���i���Ґ��ѓo�^ "," �o�@�^ ") %>
<br>
	<table border="0">
	    <tr><td valign="bottom">

	        <table border="0" width="100%">
	            <tr><td class="search">

	                <table border="0">
	                    <tr valign="middle">
							<td align="left" nowrap>�N�x</td>
	                        <td align="left" colspan="3">
								<%If m_Rs_Nendo.EOF Then%>
									<select name="cboRisyuKakoNendo" style='width:150px;' DISABLED>
										<option value="">�f�[�^������܂���
								<%Else%>
									<select name="cboRisyuKakoNendo" style='width:150px;' onchange = 'javascript:f_changeKamoku()'>
									<%Do Until m_Rs_Nendo.EOF%>
										<option value='<%=m_Rs_Nendo("T17_NENDO")%>'  <%=f_Selected(m_Rs_Nendo("T17_NENDO"),m_iRisyuKakoNendo)%>><%=m_Rs_Nendo("T17_NENDO")%>
										<%m_Rs_Nendo.MoveNext%>
									<%Loop%>
								<%End If%>
								</select>
							</td>
	                        <td>&nbsp;</td>
	                        <td align="left" nowrap>�Ȗ�</td>
	                        <td align="left">
								<%If m_iSikenKbn = "" Then%>
									<select name="cboKamoku" style='width:230px;' DISABLED>
										<option value="">�f�[�^������܂���
								<%Else%>
									<%If m_Rs.EOF Then%>
										<select name="cboKamoku" style='width:230px;' DISABLED>
											<option value="">�Ȗڃf�[�^������܂���
									<%Else%>
										<select name="cboKamoku" style='width:230px;' onclick="javasript:f_SetData();">
										<%Do Until m_Rs.EOF%>
											<%
											
											'//�I���Ȗڂ��u���������Ă��ꍇ�̕\�� Add 2001.12.17 ���c
											If f_chkOkikae(m_Rs("KAMOKU")) = 0 then
												m_Rs.MoveNext
											
											Else
												w_sKamokuCd_s = m_Rs("KAMOKU")
												w_sKamokuNM_s = m_Rs("KAMOKUMEI")
													'//�\�����e���쐬
													If f_LevelChk(m_Rs("GAKUNEN"),m_Rs("KAMOKU")) = true then 
														w_Str=""
														w_Str= w_Str & m_Rs("KAMOKUMEI") & "�@"
														
													Else
															w_Str=""
															w_Str= w_Str & m_Rs("KAMOKUMEI") & "�@"	
													End If
													
											%>

											<option value=<%=w_sKamokuCd_s  & "#@#" & w_sKamokuNM_s%> ><%=w_Str%>
											<%
											'2002/02/21 �ǉ� ITO ���уf�[�^�̍X�V�����擾���邽�߂�KEY��ޔ�
											w_sKamokuCd_s = m_Rs("KAMOKU")
											w_sKamokuNM_s = m_Rs("KAMOKUMEI")
											%>

											<%m_Rs.MoveNext%>
											<% End IF %>
										<%Loop%>
									<%End If
								End If
								%>
								</select>
							</td>
	                    </tr>
						<tr>
					        <td colspan="7" align="right">
							<%If m_RsCnt = 0 Then%>
								<input type="button" class="button" value="�@�\�@���@" DISABLED>
							<%Else%>
								<input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
							<% End IF %>
					        </td>
						</tr>
	                </table>
	            </td>
				</tr>
	        </table>
	        </td>
	    </tr>
	</table>

	<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtKamokuCd" value="<%=w_sKamokuCd_s%>">
	<input type="hidden" name="txtKamokuNM" value="<%=w_sKamokuNM_s%>">
	<input type="hidden" name="txtRisyuKakoNendo" value="<%=m_iRisyuKakoNendo%>">
	<input type="hidden" name="txtTable" value="<%=m_sGetTable%>">
	 <input type="hidden" name="txtMode"  value = "">
	<!--ADD ST-->  
	<input type="hidden" name="txtUpdDate" value="<%=gf_GetT16UpdDate(m_iNendo,w_iGakunen_s,w_sGakkaCd_s,w_sKamokuCd_s,"")%>">
	<!--ADD ED --> 
	<input type="hidden" name="SYUBETU" value="">
	
	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>