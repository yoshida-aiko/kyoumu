<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ȓ����o�^
' ��۸���ID : gak/sei0600/sei0600_main.asp
' �@      �\: ���y�[�W ���ȓ����̓o�^���
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:
'               �I�����ꂽ�����敪�̌��ȓ�����o�^���邽�߂̉�ʕ\��
'-------------------------------------------------------------------------
' ��      ��: 2001/09/26 �J�e �ǖ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n

    '�s�����I��p��Where����
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '�����R�[�h
    Public m_sGakuNo        '�����R���{�{�b�N�X�ɓ���l
    Public m_sClass         '�N���X
    Public m_sClassNm       '�N���X��
    Public m_sGakka     '�w���̏����w��
    Public m_iSikenKBN
    Public m_iSyubetu    
    Public m_sGakunen
	'//�z��
    Public m_iKesseki()     '���Ȑ�
    Public m_iKibiki()		'��������
    Public m_iKessekiRui()  '���ȏW�v�l
    Public m_iKibikiRui()	'�������W�v�l
    Public m_sGakuseiNo()	'�w���ԍ��i5�N�Ԕԍ��j
    Public m_sGakusekiNo()	'�w�Дԍ��i1�N�Ԕԍ��j
    Public m_sGakuSimei()	'�w������

    Public  m_GRs,m_DRs
    Public  m_Rs,m_KskRs
    Public  m_iMax          '�ő�y�[�W
    Public  m_iDsp          '�ꗗ�\���s��
	Public  m_rCnt
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
    w_sMsgTitle="���ȓ����o�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '// ���Ұ�SET
        Call s_SetParam()
		'===============================
		'//���ԃf�[�^�̎擾
		'===============================
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			'// �y�[�W��\��
			Call No_showPage("���ѓ��͊��ԊO�ł��B")
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

        '//�w���f�[�^�擾
        w_iRet = f_Gakusei()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//�����o���W�v�l�f�[�^�擾
        w_iRet = f_GetKessekiData(m_KskRs, m_iSikenKBN, m_sGakunen, m_sClass, w_sKaisibi, w_sSyuryobi, "")
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
		'//�W�v�l�f�[�^�̉��H�擾
        w_iRet = f_Kesseki(m_KskRs)
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

      '// �y�[�W��\��
       Call showPage()
       Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()
End Sub

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakuNo   = request("txtGakuNo")
    m_iDsp      = C_PAGE_LINE
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm    = request("txtClassNm")
	m_iSikenKBN    = request("txtSikenKBN")
	'//�O��OR���փ{�^���������ꂽ��
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

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
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_iSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & w_date & "' "

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

Function f_Gakusei()
'********************************************************************************
'*  [�@�\]  �w���f�[�^���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	dim w_Rs,w_sSQL,w_iRet
	dim w_sSikenKBN,i

'	On Error Resume Next
	Err.Clear
	f_Gakusei = 1
	
  Do
	select case cint(m_iSikenKBN)
		case C_SIKEN_ZEN_TYU '�O������
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_Z AS KIBIKI"
		case C_SIKEN_ZEN_KIM '�O������
			w_sSikenKBN = "T13_KESSEKI_KIMATU_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_KIMATU_Z AS KIBIKI"
		case C_SIKEN_KOU_TYU '�������
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_K AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_K AS KIBIKI"
		case C_SIKEN_KOU_KIM '��������i�w�N���j
			w_sSikenKBN = "T13_SUMKESSEKI AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_SUMKIBTEI AS KIBIKI"
	End select

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI,T11_GAKUSEI_NO,T13_GAKUSEKI_NO,"
    w_sSQL = w_sSQL & 		w_sSikenKBN
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI,T13_GAKU_NEN"
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND T11_GAKUSEI_NO = T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & " ORDER BY T13_GAKUSEKI_NO "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
		Exit do
    End If
	m_rCnt = gf_GetRsCount(w_Rs)

	'//�z��̍쐬
    Redim m_iKesseki(m_rCnt)        '���Ȑ�
    Redim m_iKibiki(m_rCnt)			'��������
    Redim m_iKessekiRui(m_rCnt)  			'���ȏW�v�l
    Redim m_iKibikiRui(m_rCnt)			'�������W�v�l
    Redim m_sGakuseiNo(m_rCnt)		'�w���ԍ��i5�N�Ԕԍ��j
    Redim m_sGakusekiNo(m_rCnt)		'�w�Дԍ��i1�N�Ԕԍ��j
    Redim m_sGakuSimei(m_rCnt)		'�w������

	w_Rs.MoveFirst
	i = 1

	Do Until w_Rs.EOF
		
		m_iKesseki(i) = cint(gf_SetNull2Zero(w_Rs("KESSEKI")))
		m_iKibiki(i)	= cint(gf_SetNull2Zero(w_Rs("KIBIKI")))
		m_sGakuseiNo(i) = w_Rs("T11_GAKUSEI_NO")
		m_sGakusekiNo(i) = w_Rs("T13_GAKUSEKI_NO")
		m_sGakuSimei(i) = w_Rs("T11_SIMEI")
		m_iKessekiRui(i) = 0
		m_iKibikiRui(i) = 0
		i = i + 1
		w_Rs.MoveNext
	Loop

	f_Gakusei = 0 '����I��
	exit do
  Loop

    Call gf_closeObject(w_Rs)

End Function

Function f_GetKesskiSu(p_iSikenKBN,p_sGakuseiNo,p_iKessekiSu,p_iKibikiSu)
'********************************************************************************
'*  [�@�\]  �w���f�[�^���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	dim w_Rs,w_sSQL,w_iRet
	dim w_sSikenKBN,i

'	On Error Resume Next
	Err.Clear
	f_GetKesskiSu = 1
	
  Do
	select case cint(p_iSikenKBN)
		case C_SIKEN_ZEN_TYU '�O������
			f_GetKesskiSu = 0
			Exit Do
		case C_SIKEN_ZEN_KIM '�O������
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_Z AS KIBIKI"
		case C_SIKEN_KOU_TYU '�������
			w_sSikenKBN = "T13_KESSEKI_KIMATU_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_KIMATU_Z AS KIBIKI"
		case C_SIKEN_KOU_KIM '��������i�w�N���j
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_K AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_K AS KIBIKI"
	End select

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & 		w_sSikenKBN
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T13_GAKU_NEN"
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND T13_GAKUSEI_NO = '" & p_sGakuseiNo & "' "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
	    Call gf_closeObject(w_Rs)
		Exit do
    End If
		
		p_iKessekiSu = cint(gf_SetNull2Zero(w_Rs("KESSEKI")))
		p_iKibikiSu	= cint(gf_SetNull2Zero(w_Rs("KIBIKI")))

	f_GetKesskiSu = 0 '����I��

    Call gf_closeObject(w_Rs)

	exit do
  Loop

End Function

Function F_GetKekkaKubun(p_KksKBN)
'*******************************************************************************
' �@�@�@�\�F�o���敪�����Ȉ����ɂȂ�̂��Ȃ�Ȃ��̂��𔻒�
' �ԁ@�@�l�F�擾����
' �@�@�@�@�@(1)���Ȉ�������, (0)���Ȉ������Ȃ�
' ���@�@���Fp_sKksKBN - ���ׂ����敪
' �@�\�ڍׁF�w�肳�ꂽ�����̏o���̃f�[�^���擾����
' ���@�@�l�F�Ȃ�
'*******************************************************************************
	dim w_sSQL,w_sRs,s_iRet
	F_GetKekkaKubun = 0
	
    On Error Resume Next
    Err.Clear
		
		w_sSQL = ""
		w_sSql = w_sSql & vbCrLf & "Select "
		w_sSql = w_sSql & vbCrLf & " M01_KEKKA_KBN "
		w_sSql = w_sSql & vbCrLf & "From "
		w_sSql = w_sSql & vbCrLf & " M01_KUBUN "
		w_sSql = w_sSql & vbCrLf & "Where "
		w_sSql = w_sSql & vbCrLf & "     M01_DAIBUNRUI_CD =" & C_KESSEKI 'No19 ���ۋ敪
		w_sSql = w_sSql & vbCrLf & " AND M01_SYOBUNRUI_CD =" & cint(p_KksKBN)
		w_sSql = w_sSql & vbCrLf & " AND M01_NENDO =" & m_iNendo 

		w_iRet = gf_GetRecordset(w_sRs, w_sSQL)
		If w_iRet <> 0 then m_bErrFlg = True : Exit Function
		if w_sRs.EOF = true then m_bErrFlg = True : Exit Function
		
	F_GetKekkaKubun = cint(w_sRs("M01_KEKKA_KBN"))

    Call gf_closeObject(w_sRs)

End Function

Function f_Kesseki(m_KskRs)
'********************************************************************************
'*  [�@�\]  �擾�f�[�^�𐮗�����B
'*  [����]  �Ȃ�
'*  [�ߒl]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
    Dim w_iKaisu,w_iSktKBN,w_sGakuNo

'    On Error Resume Next
    Err.Clear
   
	f_Kesseki = 1
	'// �W�v���ʂŃ��[�v
    Do Until m_KskRs.EOF
    	
    	w_iKaisu = cint(m_KskRs("KAISU"))
    	w_iSktKBN = cint(m_KskRs("T30_SYUKKETU_KBN"))
    	w_sGakuNo = m_KskRs("T30_GAKUSEKI_NO")
    	
		'//�w���������[�v
		For i = 1 to m_rCnt 

			If w_sGakuNo = m_sGakusekiNo(i) then 
			
				If w_iSktKBN = 1 or w_iSktKBN > 3 then '�x���Ƒ��ނ͏���
					'//���ȋ敪�ɂ��W�v�l�ւ̊���U��
			    	If F_GetKekkaKubun(w_iSktKBN) = 1 then 
							m_iKessekiRui(i) = m_iKessekiRui(i) + w_iKaisu
					Else
							m_iKibikiRui(i) = m_iKibikiRui(i) + w_iKaisu
					End If
				End If
			End If
			
    	Next
	
		m_KskRs.MoveNext
    Loop

	f_Kesseki = 0

End Function

Function f_GetKessekiData(p_oRecordset, p_sSikenKbn, p_sGakunen, p_sClass, p_sKaisibi, p_sSyuryobi, p_s1NenBango)
'*******************************************************************************
' �@�@�@�\�F�o���f�[�^�̎擾
' �ԁ@�@�l�F�擾����
' �@�@�@�@�@(True)����, (False)���s
' ���@�@���Fp_oRecordset - ���R�[�h�Z�b�g
' �@�@�@�@�@p_sSikenKbn - �����敪
' �@�@�@�@�@p_sGakunen - �w�N
' �@�@�@�@�@p_sTantoKyokan - �����b�c
' �@�@�@�@�@p_sClass - �N���XNo
' �@�@�@�@�@p_sKaisibi - �J�n��
' �@�@�@�@�@p_sSyuryobi - �I����
' �@�@�@�@�@p_s1NenBango - �P�N�Ԕԍ�
' �@�\�ڍׁF�w�肳�ꂽ�����̏o���̃f�[�^���擾����
' ���@�@�l�F�Ȃ�
'*******************************************************************************
	Dim w_bRtn			'�߂�l
	Dim w_sSql			'SQL
	
'	On Error Resume Next
	'== ������ ==
	gf_GetKessekiData = 1
	w_bRtn=False
	w_sSql=""
	'== �o�����擾����J�n���ƏI�������擾���� ==
	'//(�����Ԃ̊���)
	w_bRtn = gf_GetKaisiSyuryo(cint(p_sSikenKbn), p_sGakunen, p_sKaisibi, p_sSyuryobi)

	If w_bRtn <> True Then
		Exit Function
	End If

	'== �o�����擾���� ==
	'SQL�쐬
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & "SELECT "
	w_sSql = w_sSql & vbCrLf & "	Count(T30_GAKUSEKI_NO) as KAISU,"
	w_sSql = w_sSql & vbCrLf & "	T30_CLASS,"
	w_sSql = w_sSql & vbCrLf & "	T30_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & "	T30_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "FROM "
	w_sSql = w_sSql & vbCrLf & "	T30_KESSEKI "
	w_sSql = w_sSql & vbCrLf & "Where "
	w_sSql = w_sSql & vbCrLf & "	T30_NENDO = " & session("NENDO") & " "		'�N�x
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_GAKUNEN = " & p_sGakunen & " "					'�w�N
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_CLASS = " & p_sClass & " "					'�N���X
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_HIDUKE >= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sKaisibi & "' "								'�J�n��
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_HIDUKE <= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sSyuryobi & "' "								'�I����
'	w_sSql = w_sSql & vbCrLf & "	And "
'	w_sSql = w_sSql & vbCrLf & "	T30_SYUKKETU_KBN IN ('" & C_KETU_KEKKA & "','" & C_KETU_TIKOKU & "','"& C_KETU_SOTAI &"')"
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T30_SYUKKETU_KBN >= " & C_KETU_KEKKA & " "

	'== �P�N�Ԕԍ����w�肳��Ă���ꍇ ==
	If p_s1NenBango <>"" Then
		w_sSql = w_sSql & vbCrLf & "And "
		w_sSql = w_sSql & vbCrLf & "T30_GAKUSEKI_NO = " & p_s1NenBango & " "			'�N���X
	End If
	
	w_sSql = w_sSql & vbCrLf & "Group By "
	w_sSql = w_sSql & vbCrLf & "T30_CLASS,"
	w_sSql = w_sSql & vbCrLf & "T30_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & "T30_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "Order By "
	w_sSql = w_sSql & vbCrLf & "T30_CLASS, "
	w_sSql = w_sSql & vbCrLf & "T30_GAKUSEKI_NO "

	'== �f�[�^�̎擾 ==
	Set p_oRecordset = Server.CreateObject("ADODB.Recordset")

	'== ���s�����Ƃ� ==
	    If gf_GetRecordset(p_oRecordset, w_sSql) <> 0 Then
		p_oRecordset.Close
		Set p_oRecordset = Nothing
		
		Exit Function
	End If
	gf_GetKessekiData = 0
	
End Function

Sub No_showPage(p_msg)
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
	</head>

	<body>
	<center>
	<br><br><br>
			<span class="msg"><%=p_msg%></span>
	</center>
	</body>

	</html>

<%
End Sub

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
<link rel="stylesheet" href="../../common/style.css" type="text/css">

<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="JavaScript">
<!--
	var chk_Flg;
	chk_Flg = false;
	//************************************************************
	//  [�@�\]  �y�[�W���[�h������
	//  [����]
	//  [�ߒl]
	//  [����]
	//************************************************************
	function window_onload(){

        document.frm.target="topFrame";
        document.frm.action="sei0600_topDisp.asp";
        document.frm.submit();
	return true;
	}

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){

        // �������S�C�����̌�����������
//        if( getLengthB(document.frm.Syoken.value) > "200" ){
//            window.alert("�S�C�����̗��͑S�p100�����ȓ��œ��͂��Ă�������");
//            document.frm.Syoken.focus();
//            return ;
//        }
        

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

        document.frm.action="sei0600_upd.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Cansel(){

        //document.frm.action="default2.asp";
        //document.frm.target="main";
        document.frm.action="default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();
    
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
	function f_MoveCur(p_inpNm,p_frm,i) {
		if (event.keyCode == 13){		//�����ꂽ�L�[��Enter(13)�̎��ɓ����B
			i++;
			if (i > <%=m_rCnt%>) {i = 1;} //i���ő�l�𒴂���ƁA�͂��߂ɖ߂�B
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
<%
	dim i
	i = 0
	
	'//NN�Ή�
	If session("browser") = "IE" Then
		w_sInputClass = "class='num'"
	Else
		w_sInputClass = ""
	End If

%>
</head>
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">
<center>
<table border="1" cellpadding="1" cellspacing="1" class="hyo">
<!--
			<tr>
				<th class="header" rowspan="2"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
				<th nowrap class="header" rowspan="2">����</th>
				<th nowrap class="header" colspan="2">����</th>
				<th nowrap class="header" colspan="2">������</th>
			</tr>
			<tr>
				<th nowrap class="header2"><font size="1">����</font></th>
				<th nowrap class="header2"><font size="1">�݌v</font></th>
				<th nowrap class="header2"><font size="1">����</font></th>
				<th nowrap class="header2"><font size="1">�݌v</font></th>
			</tr>
-->
		<% For i = 1 to m_rCnt 
		        Call gs_cellPtn(w_cell)

				'//���ۗݐϏ��敪���ݐς̂Ƃ��́A�O�����̌��Ȑ�����ъ�����������Ă���B
				If cint(m_iSyubetu) = C_K_KEKKA_RUISEKI_KEI then 
					call f_GetKesskiSu(m_iSikenKBN,m_sGakuseiNo(i),w_iKessekiSu,w_iKibikiSu)
					m_iKessekiRui(i) = m_iKessekiRui(i) + w_iKessekiSu
					m_iKibikiRui(i)  = m_iKibikiRui(i)  + w_iKibikiSu
				End If

		        If m_iKesseki(i) = 0 then m_iKesseki(i) = m_iKessekiRui(i)
		        If m_iKibiki(i) = 0 then m_iKibiki(i) = m_iKibikiRui(i)

		%>
            <TR>
                    <TD CLASS="<%=w_cell%>" width="50"><%=m_sGakusekiNo(i)%><input type="hidden" name="txtGAKUSEINO_<%=i%>" value="<%=m_sGakuseiNo(i)%>"></TD>
                    <TD CLASS="<%=w_cell%>" width="150"><%=m_sGakuSimei(i)%></TD>
                    <TD CLASS="<%=w_cell%>" width="35" align="center"><input type="text" <%=w_sInputClass%> name="txtKESSEKI_<%=i%>" value='<%=m_iKesseki(i)%>' size=2 maxlength=3 onKeyDown="f_MoveCur('txtKESSEKI_',this.form,<%=i%>)"></TD>
                    <TD CLASS="<%=w_cell%>" width="35" align="right"><%=m_iKessekiRui(i)%></TD>
                    <TD CLASS="<%=w_cell%>" width="35" align="center"><input type="text" <%=w_sInputClass%> name="txtKIBIKI_<%=i%>" value='<%=m_iKibiki(i)%>' size=2 maxlength=3 onKeyDown="f_MoveCur('txtKIBIKI_',this.form,<%=i%>)"></TD>
                    <TD CLASS="<%=w_cell%>" width="35" align="right"><%=m_iKibikiRui(i)%></TD>
            </TR>
		<% Next %>
        </td>
    </TR>
</TABLE>

<br>
	<table width="50%">
	<tr>
		<td align="center"><input type="button" class="button" value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
		<input type="button" class="button" value="�L�����Z��" onclick="javascript:f_Cansel()"></td>
	</tr>
	</table>

	<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
	<input type="hidden" name="txtSikenKBN" value="<%=m_iSikenKBN%>">
	<input type="hidden" name="txtCnt" value="<%=m_rCnt%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
