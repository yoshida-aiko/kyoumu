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
' �C      ��: 2002/06/11 ���V ���D
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n

    '�s�����I��p��Where����
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '�����R�[�h
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
	
	
	'------------- ���V �ǉ� 2002/06/07 -----------------
	Public  m_iTokuketu() '���ʌ���
	Public	m_iTokuCnt    '
	'----------------------------------------------------

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()
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
        If gf_OpenDatabase() <> 0 Then
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
        If f_Nyuryokudate() = 1 Then
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
        If f_Gakusei() <> 0 Then m_bErrFlg = True : Exit Do
		
		'//�����o���W�v�l�f�[�^�擾
        If f_GetKessekiData(m_KskRs, m_iSikenKBN, m_sGakunen, m_sClass, w_sKaisibi, w_sSyuryobi, "") <> 0 Then 
        	m_bErrFlg = True
        	Exit Do
        end if
        
		'//�W�v�l�f�[�^�̉��H�擾
        If f_Kesseki(m_KskRs) <> 0 Then m_bErrFlg = True : Exit Do
		
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

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()
	m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp      = C_PAGE_LINE
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm    = request("txtClassNm")
	m_iSikenKBN    = request("txtSikenKBN")
End Sub

'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_Nyuryokudate()
	dim w_date
	
	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1
	
	w_date = gf_YYYY_MM_DD(date(),"/")
	
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
		
		If gf_GetRecordset(m_DRs, w_sSQL) <> 0 Then
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
'*  [�@�\]  �w���f�[�^���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_Gakusei()
	dim w_Rs,w_sSQL,w_iRet,w_Rs2
	dim w_sSikenKBN,i,w_iCnt,w_vStr
	Dim w_Type
	
	On Error Resume Next
	
	Err.Clear
	f_Gakusei = 1
	
	'---------------------KANAZAWA 2002/6/10--------------------------------------------
	w_sSQL = ""	
	w_sSQL = w_sSQL & vbCrLf & "  SELECT "	
	w_sSQL = w_sSQL & vbCrLf & "  M01_SYOBUNRUIMEI "	
	w_sSQL = w_sSQL & vbCrLf & "  FROM "	
	w_sSQL = w_sSQL & vbCrLf & "   M01_KUBUN "	
	w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo	
	w_sSQL = w_sSQL & vbCrLf & "  AND M01_DAIBUNRUI_CD = " & C_M01_DAIBUNRUI150 '���ʌ���	
	w_sSQL = w_sSQL & vbCrLf & "  ORDER BY M01_SYOBUNRUI_CD "	
	
	If gf_GetRecordset(w_Rs2, w_sSQL) <> 0 Then
        m_bErrFlg = True
		Exit function
    End If
    
    
	m_iTokuCnt = gf_GetRsCount(w_Rs2)
	
	'----------------------------------------------------------------------------------
	
  Do
	select case cint(m_iSikenKBN)
		case C_SIKEN_ZEN_TYU '�O������
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_Z AS KIBIKI "
			w_Type = "_ZT"
		case C_SIKEN_ZEN_KIM '�O������
			w_sSikenKBN = "T13_KESSEKI_KIMATU_Z AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_KIMATU_Z AS KIBIKI "
			w_Type = "_ZK"
		case C_SIKEN_KOU_TYU '�������
			w_sSikenKBN = "T13_KESSEKI_TYUKAN_K AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_KIBIKI_TYUKAN_K AS KIBIKI "
			w_Type = "_KT"
		case C_SIKEN_KOU_KIM '��������i�w�N���j
			w_sSikenKBN = "T13_SUMKESSEKI AS KESSEKI,"
			w_sSikenKBN = w_sSikenKBN & "T13_SUMKIBTEI AS KIBIKI"
			w_Type = ""
	End select
	
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI,T11_GAKUSEI_NO,T13_GAKUSEKI_NO,"
    
    for w_num=1 to 10
	    w_sSql = w_sSql & " T13_TOKUKETU" & w_num & w_Type &" as T13_TOKUKETU" & w_num & " ,"
    next
    
    w_sSQL = w_sSQL & 		w_sSikenKBN
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI,T13_GAKU_NEN"
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND T11_GAKUSEI_NO = T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & " ORDER BY T13_GAKUSEKI_NO "
	
	If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
		Exit do
    End If
    
	m_rCnt = gf_GetRsCount(w_Rs)
	
	'//�z��̍쐬
    Redim m_iKesseki(m_rCnt)        		'���Ȑ�
    Redim m_iKibiki(m_rCnt)					'��������
    Redim m_iKessekiRui(m_rCnt)  			'���ȏW�v�l
    Redim m_iKibikiRui(m_rCnt)				'�������W�v�l
    Redim m_sGakuseiNo(m_rCnt)				'�w���ԍ��i5�N�Ԕԍ��j
    Redim m_sGakusekiNo(m_rCnt)				'�w�Дԍ��i1�N�Ԕԍ��j
    Redim m_sGakuSimei(m_rCnt)				'�w������
	Redim m_iTokuketu(m_iTokuCnt,m_rCnt)	'���ʌ���
	w_Rs.MoveFirst
	
	i = 1
	w_iCnt = 0
	
	Do Until w_Rs.EOF
		m_iKesseki(i) = cint(gf_SetNull2Zero(w_Rs("KESSEKI")))
		m_iKibiki(i)	= cint(gf_SetNull2Zero(w_Rs("KIBIKI")))
		m_sGakuseiNo(i) = w_Rs("T11_GAKUSEI_NO")
		m_sGakusekiNo(i) = w_Rs("T13_GAKUSEKI_NO")
		m_sGakuSimei(i) = w_Rs("T11_SIMEI")
		m_iKessekiRui(i) = 0
		m_iKibikiRui(i) = 0
		'--------------------------------2002/6/10 kanazawa----------------------------
		For w_iCnt = 1 To m_iTokuCnt
			w_vStr = "T13_TOKUKETU" & w_iCnt
			m_iTokuketu(w_iCnt,i) = cint(gf_SetNull2Zero(w_Rs(w_vStr)))
		Next 
		'------------------------------------------------------------------------------
		i = i + 1
		w_Rs.MoveNext
	Loop

	f_Gakusei = 0 '����I��
	exit do
  Loop

    Call gf_closeObject(w_Rs)
	Call gf_closeObject(w_Rs2)


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
        
        //------------------- kanazawa 2002/6/11 ------------------------------------------------------------------------------
        
		// ���������͍��ڂ̒l����������
		var i;
		var w_iNum;
		var w_oKesseki;
		var w_oTokuketu;
		
		for (i = 1; i<= <%=m_rCnt%>; i++){
			w_oKesseki = eval("document.frm.txtKESSEKI_" + i );
			
			if (!f_CheckNum(w_oKesseki)) {
				alert("���l����͂��ĉ������B");
				return false;
				break;
			}	else {
					for (w_iNum = 1; w_iNum<= <%=m_iTokuCnt%>; w_iNum++) {
						w_oTokuketu = eval("document.frm.txtKIBIKI" + w_iNum + "_" + i );
							if (!f_CheckNum(w_oTokuketu)) {
								alert("���p��������͂��ĉ������B");
								return false;
								break;
							}
					}
				}
			
		} //next i;
		
		//------------------------------------------------------------------------------------------------------------------------
		
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return false;
        };
        //'--------------------------------------- kanazawa 2002/6/12 ---------------------------------------------------------------
		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("�o�^���Ă��܂��E�E�E�E�@�@���΂炭���҂���������")%>"
		//'--------------------------------------------------------------------------------------------------------------------------
		
        document.frm.target="main";
        document.frm.action="sei0600_upd.asp";
        document.frm.submit();
    };
    
// ------------------ kanazawa 2002/6/10 ------------------------------    
//************************************************************
//  [�@�\]  �ȈՐ��l�^�`�F�b�N
//************************************************************
	function f_CheckNum(pFromName){
		var wFromName;
		
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
//----------------------------------------------------------------------	
	
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
                    <TD CLASS="<%=w_cell%>" width="80"><%=m_sGakusekiNo(i)%><input type="hidden" name="txtGAKUSEINO_<%=i%>" value="<%=m_sGakuseiNo(i)%>"></TD>
                    <TD CLASS="<%=w_cell%>" width="150"><%=m_sGakuSimei(i)%></TD>
                    <TD CLASS="<%=w_cell%>" width="35" align="center"><input type="text" <%=w_sInputClass%> name="txtKESSEKI_<%=i%>" value='<%=m_iKesseki(i)%>' size=2 maxlength=3 onKeyDown="f_MoveCur('txtKESSEKI_',this.form,<%=i%>)"></TD>
                    <TD CLASS="<%=w_cell%>" width="35" align="right"><%=m_iKessekiRui(i)%></TD>
                    <!-- '------------------ Kanazawa 2002/6/10 ---------------------------- -->
                    <% dim w_iIndex %>
                    <% For w_iIndex = 1 To m_iTokuCnt %>
                    	<TD CLASS="<%=w_cell%>" width="35" align="center"><input type="text" <%=w_sInputClass%> name="txtKIBIKI<%=w_iIndex%>_<%=i%>" value='<%=m_iTokuketu(w_iIndex,i)%>' size=2 maxlength=3 onKeyDown="f_MoveCur('txtKIBIKI'+<%=w_iIndex%>+'_',this.form,<%=i%>)"></TD>
                    <% Next %>
					<!-- '----------------------------------------------------------------- -->
            </TR>
		<% Next %>
        </td>
    </TR>
</TABLE>

<br>
	<table width="50%">
	<tr>
		<td align="center"><input type="button" class="button" value="�@�o�@�^�@" onclick="javascript:f_Touroku()">
		<input type="button" class="button" value="�L�����Z��" onclick="javascript:f_Cansel()"></td>
	</tr>
	</table>

	<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
	<input type="hidden" name="txtSikenKBN" value="<%=m_iSikenKBN%>">
	<input type="hidden" name="txtCnt" value="<%=m_rCnt%>">
	<input type="hidden" name="txtTokuCnt" value="<%=m_iTokuCnt%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
