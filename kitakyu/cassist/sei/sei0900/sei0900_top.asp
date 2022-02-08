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
    Public m_iGakki              '�w��
    Public m_sKyokanCd          '�����R�[�h
    Public m_iSikenKbn			'�����敪

    Public m_iDispFlg			'�X�V���\���t���O 0:�\���A1:��\��

	Public m_sGetTable			'�ȖڃR���{���쐬�����e�[�u��
    
    Public m_Rs_Siken			'���������擾
    Public m_Rs					'�w�N�A�N���X�A�Ȗڎ擾RS
    Public m_Rs_Ryu				'�w�N�A�N���X�A�Ȗڎ擾RS�@���w���̑�։Ȗ�

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
		w_iRet = f_GetKamoku_Nenmatu()
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do	

		'//���w���̑�։Ȗڎ擾�@2001/12/20 add
	        w_iRet = f_GetRyuDaigae(m_iSikenKbn)
	        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

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
    Call gf_closeObject(m_Rs_Siken)
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

End Sub

'********************************************************************************
'*  [�@�\]  �����R���{���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSiken()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & "  M01_SYOBUNRUI_CD"
		w_sSQL = w_sSQL & vbCrLf & " ,M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_SYOBUNRUI_CD < " & cint(C_SIKEN_JITURYOKU)
		w_sSQL = w_sSQL & vbCrLf & "  ORDER BY M01_SYOBUNRUI_CD"

response.write "w_sSQL:" & w_sSQL & "<BR>"
        iRet = gf_GetRecordset(m_Rs_Siken, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If

        f_GetSiken = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  ���w����֎��Ԋ������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetRyuDaigae(p_sikenKBN)

    Dim w_iRet
    Dim w_sSQL
    Dim w_sGakkiKbn '//�w���敪

    On Error Resume Next
    Err.Clear

    f_GetRyuDaigae = 1

    Do

        '//�O����敪���擾
        If m_sGakki = "ZENKI" Then
            w_sGakkiKbn = cstr(C_GAKKI_ZENKI)   '//1:�O��
        Else
            w_sGakkiKbn = cstr(C_GAKKI_KOUKI)   '//2:���
        End If

        '//�󎝎��Ƃ��擾 
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT DISTINCT "
        w_sSQL = w_sSQL & "     T27.T27_GAKUNEN, "
        w_sSQL = w_sSQL & "     T27.T27_CLASS, "
        w_sSQL = w_sSQL & "     M05.M05_CLASSMEI, "
        w_sSQL = w_sSQL & "     M05.M05_GAKKA_CD, "
        w_sSQL = w_sSQL & "     T27.T27_KAMOKU_CD, "
        w_sSQL = w_sSQL & "     T27.T27_KYOKAN_CD, "
        w_sSQL = w_sSQL & "     T16.T16_KAMOKUMEI, "
        w_sSQL = w_sSQL & "     T16.T16_OKIKAE_FLG,"
        w_sSQL = w_sSQL & "     T13_CLASS "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T16_RISYU_KOJIN T16,"
        w_sSQL = w_sSQL & "     T27_TANTO_KYOKAN T27 ,"
        w_sSQL = w_sSQL & "     M05_CLASS M05,"
        w_sSQL = w_sSQL & "     T13_GAKU_NEN T13 "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "     T27.T27_KYOKAN_CD 	= '"	&	m_sKyokanCd			&	"' AND "
        w_sSQL = w_sSQL & "     T16.T16_OKIKAE_FLG 	= "		&	C_TIKAN_KAMOKU_SAKI	&	"  AND "
        w_sSQL = w_sSQL & "     T27.T27_NENDO		= "		& 	cInt(m_iNendo) 	& 	"  AND "
        w_sSQL = w_sSQL & "     T16.T16_KAMOKU_CD  = T27.T27_KAMOKU_CD AND "
        w_sSQL = w_sSQL & "     T16.T16_NENDO      = T27.T27_NENDO AND "
        w_sSQL = w_sSQL & "     T16.T16_HAITOGAKUNEN     = T27.T27_GAKUNEN AND "
        w_sSQL = w_sSQL & "     T27.T27_NENDO      = T13.T13_NENDO AND "
        w_sSQL = w_sSQL & "     T27.T27_CLASS      = T13.T13_CLASS AND "
        w_sSQL = w_sSQL & "     T16.T16_GAKUSEI_NO = T13.T13_GAKUSEI_NO AND "
        w_sSQL = w_sSQL & "     T27.T27_NENDO      = M05.M05_NENDO AND "
        w_sSQL = w_sSQL & "     T27.T27_GAKUNEN    = M05.M05_GAKUNEN AND "
        w_sSQL = w_sSQL & "     T27.T27_CLASS      = M05.M05_CLASSNO "

'response.write w_ssql
'response.end
        iRet = gf_GetRecordset(m_Rs_Ryu, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetRyuDaigae = 99
            Exit Do
        End If

        '//����I��
        f_GetRyuDaigae = 0
        Exit Do
    Loop

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
'*  [�@�\]  ��������̎��A�w�N�E�N���X�E�ȖڃR���{���擾
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
			w_sSQL = w_sSQL & vbCrLf & "  SELECT distinct "
			w_sSQL = w_sSQL & vbCrLf & "  GAKUNEN"
			w_sSQL = w_sSQL & vbCrLf & " ,CLASS"
			w_sSQL = w_sSQL & vbCrLf & " ,KAMOKU"
			w_sSQL = w_sSQL & vbCrLf & " ,CLASSMEI"
			w_sSQL = w_sSQL & vbCrLf & " ,GAKKA_CD"
			w_sSQL = w_sSQL & vbCrLf & " ,KAMOKUMEI"
			w_sSQL = w_sSQL & vbCrLf & "  FROM"

			w_sSQL = w_sSQL & vbCrLf & "("
			
			w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
			w_sSQL = w_sSQL & vbCrLf & " 	T27_GAKUNEN AS GAKUNEN"
			w_sSQL = w_sSQL & vbCrLf & " 	,T27_CLASS AS CLASS"
			w_sSQL = w_sSQL & vbCrLf & " 	,T27_KAMOKU_CD AS KAMOKU"
			w_sSQL = w_sSQL & vbCrLf & " 	,M05_CLASSMEI AS CLASSMEI"
			w_sSQL = w_sSQL & vbCrLf & " 	,M05_GAKKA_CD AS GAKKA_CD"
			w_sSQL = w_sSQL & vbCrLf & " 	,T17_KAMOKUMEI AS KAMOKUMEI"
			w_sSQL = w_sSQL & vbCrLf & " FROM"
			w_sSQL = w_sSQL & vbCrLf & " 	T27_TANTO_KYOKAN "
			w_sSQL = w_sSQL & vbCrLf & " 	,T17_RISYUKAKO_KOJIN "
			w_sSQL = w_sSQL & vbCrLf & " 	,M05_CLASS "
			w_sSQL = w_sSQL & vbCrLf & " WHERE "
			w_sSQL = w_sSQL & vbCrLf & " 		T27_NENDO = M05_NENDO "
			w_sSQL = w_sSQL & vbCrLf & "    AND T27_GAKUNEN = M05_GAKUNEN "
			w_sSQL = w_sSQL & vbCrLf & "    AND T27_CLASS = M05_CLASSNO	"

'			w_sSQL = w_sSQL & vbCrLf & "    AND T27_KAMOKU_CD = T17_KAMOKU_CD(+)"

			w_sSQL = w_sSQL & vbCrLf & "    AND T27_KAMOKU_CD = T17_KAMOKU_CD"
			w_sSQL = w_sSQL & vbCrLf & "    AND T27_GAKUNEN = T17_HAITOGAKUNEN "

'			w_sSQL = w_sSQL & vbCrLf & "    AND M05_GAKKA_CD(+) = T17_GAKKA_CD "�@DEL 2017.12.22 Nishimura 1�N���X�ɕ����w�Ȃ����݂���̂ŃR�[�XCD�Ń����N����悤�ɏC��

'			w_sSQL = w_sSQL & vbCrLf & "    AND T17_NENDO(+) = T27_NENDO "
			w_sSQL = w_sSQL & vbCrLf & "    AND T17_NENDO = T27_NENDO "

			w_sSQL = w_sSQL & vbCrLf & "    AND T27_NENDO = " & m_iNendo
			w_sSQL = w_sSQL & vbCrLf & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
			w_sSQL = w_sSQL & vbCrLf & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "

'			w_sSQL = w_sSQL & vbCrLf & "    AND T17_OKIKAE_FLG >= " & C_TIKAN_KAMOKU_SAKI 
			w_sSQL = w_sSQL & vbCrLf & "    AND (T17_OKIKAE_FLG > " & C_TIKAN_KAMOKU_SAKI 
			w_sSQL = w_sSQL & vbCrLf & "    OR  T17_OKIKAE_FLG = 0) "
'			w_sSQL = w_sSQL & vbCrLf & "    AND T17_COURSE_CD IN ( '0' , CASE WHEN M05_COURSE_CD IS NOT NULL THEN M05_COURSE_CD ELSE T17_COURSE_CD END) " '2017.12.22 Ins Nishimura
			w_sSQL = w_sSQL & vbCrLf & "    AND T17_COURSE_CD IN ( '0' , CASE WHEN M05_GAKKA_CD = T17_GAKKA_CD THEN (CASE WHEN M05_COURSE_CD IS NOT NULL THEN M05_COURSE_CD ELSE T17_COURSE_CD END ) ELSE T17_COURSE_CD END ) " '2019.02.12 Upd Kiyomoto

			w_sSQL = w_sSQL & vbCrLf & "  GROUP BY "
			w_sSQL = w_sSQL & vbCrLf & " 	T27_NENDO"
			w_sSQL = w_sSQL & vbCrLf & " 	,T27_GAKUNEN"
			w_sSQL = w_sSQL & vbCrLf & " 	,T27_CLASS"
			w_sSQL = w_sSQL & vbCrLf & " 	,T27_KAMOKU_CD"
			w_sSQL = w_sSQL & vbCrLf & " 	,M05_GAKKA_CD"
			w_sSQL = w_sSQL & vbCrLf & " 	,M05_CLASSMEI"
			w_sSQL = w_sSQL & vbCrLf & " 	,T17_KAMOKUMEI"
			w_sSQL = w_sSQL & vbCrLf & " 	,M05_GAKKA_CD"
			
			
			
			w_sSQL = w_sSQL & vbCrLf & " union all "
'		Else 
'			w_sSQL = ""

			w_sSQL = w_sSQL & vbCrLf & " SELECT "
			w_sSQL = w_sSQL & vbCrLf & "  T26.T26_GAKUNEN AS GAKUNEN "
			w_sSQL = w_sSQL & vbCrLf & "  ,T26.T26_CLASS AS CLASS "
			w_sSQL = w_sSQL & vbCrLf & "  ,T26.T26_KAMOKU AS KAMOKU "
			w_sSQL = w_sSQL & vbCrLf & "  ,M05.M05_CLASSMEI AS CLASSMEI "
			w_sSQL = w_sSQL & vbCrLf & "  ,M05.M05_GAKKA_CD AS GAKKA_CD "
			w_sSQL = w_sSQL & vbCrLf & "  ,T15.T15_KAMOKUMEI AS KAMOKUMEI "
			w_sSQL = w_sSQL & vbCrLf & " FROM "
			w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI T26"
			w_sSQL = w_sSQL & vbCrLf & "  ,T15_RISYU T15"
			w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS M05"
			w_sSQL = w_sSQL & vbCrLf & " WHERE "
			w_sSQL = w_sSQL & vbCrLf & "   T26.T26_CLASS = M05.M05_CLASSNO "
			w_sSQL = w_sSQL & vbCrLf & "  AND T26.T26_GAKUNEN = M05.M05_GAKUNEN "
			w_sSQL = w_sSQL & vbCrLf & "  AND T26.T26_NENDO = M05.M05_NENDO "
			w_sSQL = w_sSQL & vbCrLf & "  AND T26.T26_KAMOKU = T15.T15_KAMOKU_CD(+)"
			w_sSQL = w_sSQL & vbCrLf & "  AND M05.M05_GAKKA_CD = T15.T15_GAKKA_CD"
			w_sSQL = w_sSQL & vbCrLf & "  AND T15.T15_NYUNENDO(+) = T26.T26_NENDO - T26.T26_GAKUNEN + 1"
			w_sSQL = w_sSQL & vbCrLf & "  AND T26.T26_NENDO= " & m_iNendo
			w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "'"
			w_sSQL = w_sSQL & vbCrLf & "  AND ("
			w_sSQL = w_sSQL & vbCrLf & "     T26_JISSI_KYOKAN    ='" & m_sKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & "  OR T26_SEISEKI_KYOKAN1 ='" & m_sKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & "  OR T26_SEISEKI_KYOKAN2 ='" & m_sKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & "  OR T26_SEISEKI_KYOKAN3 ='" & m_sKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & "  OR T26_SEISEKI_KYOKAN4 ='" & m_sKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & "  OR T26_SEISEKI_KYOKAN5 ='" & m_sKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & "  )"
			w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
			w_sSQL = w_sSQL & vbCrLf & "  T26.T26_GAKUNEN "
			w_sSQL = w_sSQL & vbCrLf & "  ,T26.T26_CLASS "
			w_sSQL = w_sSQL & vbCrLf & "  ,T26.T26_KAMOKU "
			w_sSQL = w_sSQL & vbCrLf & "  ,M05.M05_CLASSMEI "
			w_sSQL = w_sSQL & vbCrLf & "  ,T15.T15_KAMOKUMEI"
			w_sSQL = w_sSQL & vbCrLf & "  ,M05.M05_GAKKA_CD"
'
'		End If

		w_sSQL = w_sSQL & vbCrLf & " UNION ALL "

		w_sSQL = w_sSQL & vbCrLf & " SELECT  DISTINCT "
		w_sSQL = w_sSQL & vbCrLf & "  T20_JIKANWARI.T20_GAKUNEN AS GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T20_JIKANWARI.T20_CLASS AS CLASS, "
		w_sSQL = w_sSQL & vbCrLf & "  T20_JIKANWARI.T20_KAMOKU AS KAMOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSMEI AS CLASSMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKKA_CD AS GAKKA_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  M41_TOKUKATU.M41_MEISYO AS KAMOKUMEI "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T20_JIKANWARI ,M05_CLASS,M41_TOKUKATU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO "
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN"
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_NENDO = M05_CLASS.M05_NENDO "
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_KAMOKU = M41_TOKUKATU.M41_TOKUKATU_CD "
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_NENDO = M41_TOKUKATU.M41_NENDO "
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_NENDO=" & m_iNendo & " "

'2015/10/08 �w�����V�X�e���N����擾�������t�Ŏ擾�����w�����g�p���Ă��邽�߁A����ɂȂ������ɑO���̎��Ԋ����Q�Ƃł��Ȃ��Ȃ��Ă���B
'		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_GAKKI_KBN='" & m_iGakki & "' " '//2001.12.28.okada
if m_iSikenKbn = 1 or m_iSikenKbn = 2 then
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_GAKKI_KBN='1' "
elseif m_iSikenKbn = 3 or m_iSikenKbn = 4 then
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_GAKKI_KBN='2' "
end if
		w_sSQL = w_sSQL & vbCrLf & "  AND T20_JIKANWARI.T20_KYOKAN='" & m_sKyokanCd & "' "
		'//���Ƌ敪(C_JUGYO_KBN_JUHYO = 0�F���ƂƂ݂Ȃ�, C_JUGYO_KBN_NOT_JUGYO = 1:���ƂƂ݂Ȃ��Ȃ�)
		w_sSQL = w_sSQL & vbCrLf & "  AND M41_TOKUKATU.M41_JUGYO_KBN=" & C_JUGYO_KBN_JUHYO
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY "
		w_sSQL = w_sSQL & vbCrLf & "  GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  ,CLASS "
		w_sSQL = w_sSQL & vbCrLf & "  ,KAMOKU)"
'2017/12/27 Add Kiyomoto -->
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY "
		w_sSQL = w_sSQL & vbCrLf & "  GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  ,CLASS "
		w_sSQL = w_sSQL & vbCrLf & "  ,KAMOKU"
'2017/12/27 Add Kiyomoto <--
' response.write w_sSQL  & "<BR>"
' rensponse.end

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKamoku_Nenmatu = 99
            Exit Do
        End If

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
	On Error Resume Next
    Err.Clear

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
	    // ���w�N
	    if( f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
	        window.alert("�w�N�̑I�����s���Ă�������");
	        document.frm.txtGakuNo.focus();
	        return ;
	    }
	    // ���N���X
	    if( f_Trim(document.frm.txtClassNo.value) == "<%=C_CBO_NULL%>" ){
	        window.alert("�N���X�̑I�����s���Ă�������");
	        document.frm.txtClassNo.focus();
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
	        window.alert("�Ȗڂ̑I�����s���Ă�������");
			return;
		}

	    document.frm.action="sei0900_bottom.asp";
	    document.frm.target="main";
	    document.frm.submit();

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
		
		//�f�[�^�擾
        var vl = document.frm.cboKamoku.value.split('#@#');
		
        //�I�����ꂽ�f�[�^���(�w�N�A�N���X�A�Ȗ�CD���擾)
        document.frm.txtGakuNo.value=vl[0];
        document.frm.txtClassNo.value=vl[1];
        document.frm.txtKamokuCd.value=vl[2];
        document.frm.txtGakkaCd.value=vl[3];
        document.frm.txtUpdDate.value=vl[4];
        document.frm.SYUBETU.value=vl[5];
        document.frm.txtKamokuNM.value=vl[6];

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
	                        <td align="left" nowrap>�Ȗ�</td>
	                        <td align="left">
								<%If m_iSikenKbn = "" Then%>
									<select name="cboKamoku" style='width:230px;' DISABLED>
										<option value="">�f�[�^������܂���
								<%Else%>
									<%If m_Rs.EOF AND m_Rs_Ryu.EOF Then%>
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
											
													'//�\�����e���쐬
													If f_LevelChk(m_Rs("GAKUNEN"),m_Rs("KAMOKU")) = true then 
														w_Str=""
														'w_Str= w_Str & m_Rs("GAKUNEN") & "�N�@"
														'w_Str= w_Str & m_Rs("CLASSMEI") & "�@"
														w_Str= w_Str & m_Rs("KAMOKUMEI") & "�@"
													Else

															w_Str=""
															'w_Str= w_Str & m_Rs("GAKUNEN") & "�N�@"
															' w_Str= w_Str & m_Rs("CLASSMEI") & "�@"
															w_Str= w_Str & m_Rs("KAMOKUMEI") & "�@"	

													End If
													'w_Str= w_Str & f_GetKamokuName(m_Rs("T26_GAKUNEN"),m_Rs("T26_CLASS"),m_Rs("T26_KAMOKU"))
													Dim w_TUKU_FLG,w_TukuName
													
													w_TUKU_FLG = 0
													
													Call f_getTUKU(m_iNendo,m_Rs("KAMOKU"),m_Rs("GAKUNEN"),m_Rs("CLASS"),w_TUKU_FLG)
													
													if cint(gf_SetNull2Zero(w_TUKU_FLG)) = 1 then
														w_TukuName = "TOKU"
													else
														w_TukuName = "TUJO"
													end if
													
											%>
											<!-- �R�����g
											<option value="<%=m_Rs("GAKUNEN") & "#@#" & m_Rs("CLASS") & "#@#" & m_Rs("KAMOKU") & "#@#" & m_Rs("GAKKA_CD") & "#@#" & gf_GetT16UpdDate(m_iNendo,m_Rs("GAKUNEN"),m_Rs("GAKKA_CD"),m_Rs("KAMOKU"),"") & "#@#" & w_TukuName & "#@#" & m_Rs("KAMOKUMEI")%>"  ><%=w_Str%>
											-->
											<option value="<%m_Rs("KAMOKUMEI")%>"  ><%=w_Str%>

											<%
											'2002/02/21 �ǉ� ITO ���уf�[�^�̍X�V�����擾���邽�߂�KEY��ޔ�
											w_iGakunen_s = m_Rs("GAKUNEN")
											w_sGakkaCd_s = m_Rs("GAKKA_CD")
											w_sKamokuCd_s = m_Rs("KAMOKU")
											w_sKamokuNM_s = m_Rs("KAMOKUMEI")
											%>

											<%m_Rs.MoveNext%>
											<% End IF %>
										<%Loop%>

										<% '���w���̑�։Ȗ� 2001/12/20 add%>
										<%Do Until m_Rs_Ryu.EOF%>
											<%
      											IF cint(m_Rs_Ryu("T16_OKIKAE_FLG")) = cint(C_TIKAN_KAMOKU_SAKI) THEN
													w_Str=""
													w_Str= w_Str & "���w�����"
													w_Str= w_Str & ""
													w_Str= w_Str & "�@�@�@"
													w_Str= w_Str & m_Rs_Ryu("T16_KAMOKUMEI") & "�@"	
												%>
												<option value="<%=m_Rs_Ryu("T27_GAKUNEN") & "#@#" & m_Rs_Ryu("T27_CLASS") & "#@#" & m_Rs_Ryu("T27_KAMOKU_CD") & "#@#" & m_Rs_Ryu("M05_GAKKA_CD") & "#@#" & gf_GetT16UpdDate(m_iNendo,m_Rs_Ryu("T27_GAKUNEN"),m_Rs_Ryu("M05_GAKKA_CD"),m_Rs_Ryu("T27_KAMOKU_CD"),"") & "#@#" & m_Rs_Ryu("T16_KAMOKUMEI")%>"><%=w_Str%>


											<%
											'2002/02/21 �ǉ� ITO ���уf�[�^�̍X�V�����擾���邽�߂�KEY��ޔ�
											w_iGakunen_s = m_Rs_Ryu("T27_GAKUNEN")
											w_sGakkaCd_s = m_Rs_Ryu("M05_GAKKA_CD")
											w_sKamokuCd_s = m_Rs_Ryu("T27_KAMOKU_CD")
											w_sKamokuNM_s = m_Rs_Ryu("T16_KAMOKUMEI")
											%>

											<%	Else
													Exit Do
												END IF
											m_Rs_Ryu.MoveNext
										Loop
									End If
								End If
								%>
								</select>
							</td>
	                    </tr>
						<tr>
					        <td colspan="7" align="right">
					        <input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
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
	<input type="hidden" name="txtGakuNo"   value="<%=w_iGakunen_s%>">
	<input type="hidden" name="txtClassNo"  value="">
	<input type="hidden" name="txtKamokuCd" value="<%=w_sKamokuCd_s%>">
	<input type="hidden" name="txtKamokuNM" value="<%=w_sKamokuNM_s%>">
	<input type="hidden" name="txtGakkaCd" value="<%=w_sGakkaCd_s%>">
	<input type="hidden" name="txtTable" value="<%=m_sGetTable%>">
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