<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����������ꗗ
' ��۸���ID : web/web0360/web0360_edt.asp
' �@      �\: ���k�̕����������̍X�V
'-------------------------------------------------------------------------
' ��      ��:   txtMode			:�������[�h
'               txtClubCd		:����CD
'               GAKUSEI_NO		:�w��NO
'               cboGakunenCd	:�w�N
'               cboClassCd		:�N���XNO
'               txtTyuClubCd	:���w�Z����CD
'
' ��      �n:	txtClubCd		:����CD
'               cboGakunenCd	:�w�N
'               cboClassCd		:�N���XNO
'               txtTyuClubCd	:���w�Z����CD
' ��      ��:
'           �����k�̕��������̍X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2001/08/22 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
	Public m_iSyoriNen			'//�N�x
	Public m_iKyokanCd			'//��������
	Public m_sClubCd			'//�N���uCD
	Public m_iGakunen           '//�w�N
	Public m_iClassNo           '//�N���XNO
	Public m_sTyuClubCd			'//���w�Z�N���uCD
	Public m_sMode				'//�������[�h

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

	Dim w_iRet			  '// �߂�l
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="�����������ꗗ"
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

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'//�ϐ�������
		Call s_ClearParam()

		'// Main���Ұ�SET
		Call s_SetParam()

'//�f�o�b�O
'Call s_DebugPrint()

		'// ���k�̕������̍X�V
		w_iRet = f_ClubUpdate()
		If w_iRet <> 0 Then
			m_bErrFlg = True
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
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

	m_iSyoriNen  = ""
	m_iKyokanCd  = ""
	m_sClubCd    = ""
	m_iGakunen   = ""
	m_iClassNo   = ""
	m_sTyuClubCd = ""
	m_sMode      = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

	m_iSyoriNen  = Session("NENDO")
	m_iKyokanCd  = Session("KYOKAN_CD")
	m_sClubCd    = Request("txtClubCd")
	m_iGakunen   = Request("cboGakunenCd")	'//�w�N
	m_iClassNo   = Request("cboClassCd")	'//�N���X
	m_sTyuClubCd = replace(Request("txtTyuClubCd"),"@@@","")	'//���w�Z�N���uCD
	m_sMode      = Request("txtMode")	'//�N���X

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iSyoriNen  = " & m_iSyoriNen  & "<br>"
	response.write "m_iKyokanCd  = " & m_iKyokanCd  & "<br>"
	response.write "m_sClubCd    = " & m_sClubCd	& "<br>"
	response.write "m_iGakunen   = " & m_iGakunen   & "<br>"
	response.write "m_iClassNo   = " & m_iClassNo   & "<br>"
	response.write "m_sTyuClubCd = " & m_sTyuClubCd & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �����S���̓������X�V
'********************************************************************************
Function f_UpdNyububi()

	f_UpdNyububi = False
	On Error Resume Next
	Err.Clear

	Dim i
	Dim wFieldName

	w_sGakuseiNo   = split(replace(Request("hidGakuseiNo")," ",""),",")
	w_iGakusekiCnt = UBound(w_sGakuseiNo)
	wFieldName     = split(replace(Request("hidFieldName")," ",""),",")
	w_sNyububi     = split(replace(Request("txtNyububiC")," ",""),",")
	w_Taibubi      = split(replace(Request("txtTaibubi")," ",""),",")
	w_TaibuFlg     = split(replace(Request("hidTaibuFlg")," ",""),",")


'response.write Request("hidGakuseiNo") & "<BR>"
'response.write Request("hidFieldName") & "<BR>"
'response.write Request("txtNyububiC") & "<BR><BR>"

	i = 0
	Do Until i > w_iGakusekiCnt

		'//���������X�V
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN SET"

		if w_sNyububi(i) = "0000000000" then
			w_sSQL = w_sSQL & vbCrLf & "  T13_CLUB_" & wFieldName(i) & " = null"
			w_sSQL = w_sSQL & vbCrLf & " ,T13_CLUB_" & wFieldName(i) & "_NYUBI = null"
			w_sSQL = w_sSQL & vbCrLf & " ,T13_CLUB_" & wFieldName(i) & "_TAIBI = null"
			w_sSQL = w_sSQL & vbCrLf & " ,T13_CLUB_" & wFieldName(i) & "_FLG   = null"
		Else
			w_sSQL = w_sSQL & vbCrLf & " 	T13_CLUB_" & wFieldName(i) & "_NYUBI = '" & gf_YYYY_MM_DD(w_sNyububi(i),"/") & "'"
			if w_TaibuFlg(i) then
				w_sSQL = w_sSQL & vbCrLf & " 	,T13_CLUB_" & wFieldName(i) & "_TAIBI = '" & gf_YYYY_MM_DD(w_Taibubi(i),"/") & "'"
			End if
		End if

		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  	    T13_NENDO      =  " & cInt(m_iSyoriNen)
		w_sSQL = w_sSQL & vbCrLf & "  	AND T13_GAKUSEI_NO = '" & w_sGakuseiNo(i) & "'"
		w_sSQL = w_sSQL & vbCrLf & "  	AND T13_GAKUSEI_NO = '" & w_sGakuseiNo(i) & "'"

'response.write w_sSQL & "<BR>"
'response.write iRet & "<BR>"
		iRet = gf_ExecuteSQL(w_sSQL)
		If iRet <> 0 Then
'response.end
'response.write "rollllllllllllllllllllllllllllllllllllll" & "<BR>"
			'//۰��ޯ�
			Call gs_RollbackTrans()
			Exit Function
		End If

		i = i + 1
	Loop

'response.end

	f_UpdNyububi = True

End Function


'********************************************************************************
'*  [�@�\]  ���k�̕��������̍X�V
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_ClubUpdate()

	Dim w_sSQL
	Dim w_Rs
	Dim w_iKekka

	On Error Resume Next
	Err.Clear

	f_ClubUpdate = 1

	Do 

		'================
		'//�w��No���擾
		'================
		w_sGakuseiNo = split(replace(Request("GAKUSEI_NO")," ",""),",")
		w_iGakusekiCnt = UBound(w_sGakuseiNo)

		'================
		'//���������擾
		'================
		w_sNyububi = split(replace(Request("hidNyububi")," ",""),",")

		'================
		'//�ޕ������擾
		'================
		w_sTaibubi = split(replace(Request("hidTaibubi")," ",""),",")

		'=================
		'//��ݻ޸��݊J�n
		'=================
		Call gs_BeginTrans()

		'// �����S���̓������X�V
		if CStr(m_sMode) = "DELETE" then Call f_UpdNyububi()

		'====================================
		'//�I�����ꂽ���k�̐l�������������s
		'====================================
		For i=0 To w_iGakusekiCnt

			'//�X�V�׸ޏ�����
			w_bClub1 = False
			w_bClub2 = False

			'=====================================
			'//���݂̐��k�̃N���u�󋵂��擾
			'=====================================
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " SELECT "
			w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_1 "
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2"
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_TAIBI"
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_TAIBI"
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_FLG"
			w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_FLG"
			w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
			w_sSQL = w_sSQL & vbCrLf & " WHERE "
			w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen)
			w_sSQL = w_sSQL & vbCrLf & "  AND T13_GAKU_NEN.T13_GAKUSEI_NO='" & w_sGakuseiNo(i) & "'"

			iRet = gf_GetRecordset(rs, w_sSQL)
			If iRet <> 0 Then
				Call gs_RollbackTrans()
				'//۰��ޯ�
				'ں��޾�Ă̎擾���s
				f_ClubUpdate = 99
				Exit Do
			End If

			'//�f�[�^����̎�
			If rs.EOF = False Then

				'=====================================
				'//�������[�h�ɂ�菈����U�蕪����
				'=====================================
				Select Case m_sMode
					Case "INSERT"
						'//�o�^�����̏ꍇ�i�����j
						Call f_InsDataSet(rs("T13_CLUB_1"),rs("T13_CLUB_2"),w_bClub1,w_bClub2,w_sClubCd,rs("T13_CLUB_1_FLG"),rs("T13_CLUB_2_FLG"),rs("T13_CLUB_1_TAIBI"),rs("T13_CLUB_2_TAIBI"))

					Case "DELETE"
						'//�폜�����̏ꍇ�i�ޕ��j
						Call f_DelDataSet(rs("T13_CLUB_1"),rs("T13_CLUB_2"),w_bClub1,w_bClub2,w_sClubCd)

					Case Else
						'//�������[�h�擾���s
						m_sErrMsg = "�������[�h������܂���(�V�X�e���G���[)"
						Exit Do

				End Select

			Else
				'//���k��񂪂Ȃ����͂Ȃ����߁A�����͒ʂ�Ȃ�
				'//۰��ޯ�
				Call gs_RollbackTrans()
				m_sErrMsg = "�f�[�^�̍X�V�Ɏ��s���܂����B"
				Exit Do
			End If

			'================
			'//�X�V�������s
			'================
			If w_bClub1 = True Or w_bClub2 = True Then
				'//���������X�V
				w_sSQL = ""
				w_sSQL = w_sSQL & vbCrLf & " UPDATE T13_GAKU_NEN"
				w_sSQL = w_sSQL & vbCrLf & " SET"

				'//�N���u1���X�V
				If w_bClub1 = True Then
					
					'�����̏ꍇ
					If m_sMode = "INSERT" Then	
						w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_1_FLG=1"	'1=����
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_1='" & w_sClubCd & "'"
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_1_NYUBI='" & gf_YYYY_MM_DD(w_sNyububi(i),"/") & "'"
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_1_TAIBI=Null"

					'�ޕ��̏ꍇ
					Else
						w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_1_FLG=2"	'2=�ޕ�
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_1_TAIBI='" & gf_YYYY_MM_DD(w_sTaibubi(i),"/") & "'"
					End If

				End If

				'//�N���u2���X�V
				If w_bClub2 = True Then
					'�����̏ꍇ
					If m_sMode = "INSERT" Then	
						w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_2_FLG=1"	'1=����
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_2='" & w_sClubCd & "'"
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_2_NYUBI='" & gf_YYYY_MM_DD(w_sNyububi(i),"/") & "'"
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_2_TAIBI=Null"

					'�ޕ��̏ꍇ
					Else
						w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLUB_2_FLG=2"	'2=�ޕ�
						w_sSQL = w_sSQL & vbCrLf & " ,T13_GAKU_NEN.T13_CLUB_2_TAIBI='" & gf_YYYY_MM_DD(w_sTaibubi(i),"/") & "'"
					End If
				End If

				w_sSQL = w_sSQL & vbCrLf & " WHERE "
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen)
				w_sSQL = w_sSQL & vbCrLf & "  AND T13_GAKU_NEN.T13_GAKUSEI_NO='" & w_sGakuseiNo(i) & "'"
				'//�X�V�������s
				iRet = gf_ExecuteSQL(w_sSQL)
				If iRet <> 0 Then
					'//۰��ޯ�
					Call gs_RollbackTrans()
					f_ClubUpdate = 99
					Exit Do
				End If

			End If

			'//ں��޾��CLOSE
			Call gf_closeObject(rs)
		Next

		'//�Я�
		Call gs_CommitTrans()

		'//����I��
		f_ClubUpdate = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �o�^���A�N���u1���X�V���邩�A�܂��̓N���u2���X�V���邩�𒲍�����
'*  [����]  p_T13_CLUB1:T13_CLUB_1
'*          p_T13_CLUB2:T13_CLUB_2
'*	[�ߒl]  p_bClub1=True : Club1�X�V�� p_bClub1=False : Club1�X�V�s��
'*		    p_bClub2=True : Club2�X�V�� p_bClub2=False : Club2�X�V�s��
'*          p_sClubCd:�o�^����CD��Ԃ�
'*  [����]  
'********************************************************************************
Function f_InsDataSet(p_T13_CLUB1,p_T13_CLUB2,p_bClub1,p_bClub2,p_sClubCd,p_iFlg1,p_iFlg2,p_sTaiBi1,p_sTaiBi2)

		'//������
		p_bClub1 = False
		p_bClub2 = False
		p_sClubCd = ""

		'�����Ƃ������̏ꍇ�A�o�^�s��
		If gf_SetNull2String(p_iFlg1) = "1" And gf_SetNull2String(p_iFlg2) = "1" Then
'response.write "�����Ƃ������̏ꍇ�A�o�^�s��"
			Exit Function
		End If

		'//����N���u�őޕ����Ă����ꍇ�A�N���u1�ɓo�^
		If gf_SetNull2String(p_T13_CLUB1) = m_sClubCd And gf_SetNull2String(p_iFlg1) = "2" Then
'response.write "����N���u�őޕ����Ă����ꍇ�A�N���u1�ɓo�^"
			p_bClub1 = True
			p_bClub2 = False
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//����N���u�őޕ����Ă����ꍇ�A�N���u2�ɓo�^
		If gf_SetNull2String(p_T13_CLUB2) = m_sClubCd And gf_SetNull2String(p_iFlg2) = "2" Then
'response.write "����N���u�őޕ����Ă����ꍇ�A�N���u2�ɓo�^"
			p_bClub1 = False
			p_bClub2 = True
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//�N���u1���󂫂̏ꍇ�A�N���u1�ɓo�^
		If gf_SetNull2String(p_T13_CLUB1) = "" Then
'response.write "�N���u1���󂫂̏ꍇ�A�N���u1�ɓo�^"
			p_bClub1 = True
			p_bClub2 = False
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//�N���u2�̂݋󂫂̏ꍇ�A�N���u2�ɓo�^
		If gf_SetNull2String(p_T13_CLUB2) = "" Then
'response.write "�N���u2�̂݋󂫂̏ꍇ�A�N���u2�ɓo�^"
			p_bClub1 = False
			p_bClub2 = True
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//�����Ƃ��Ⴄ�N���u�łǂ��炩�ޕ����Ă����ꍇ�A�N���u1�ɓo�^
		If gf_SetNull2String(p_iFlg1) = "2" And gf_SetNull2String(p_iFlg2) <> "2" Then
'response.write "�����Ƃ��Ⴄ�N���u�łǂ��炩�ޕ����Ă����ꍇ�A�N���u1�ɓo�^"
			p_bClub1 = True
			p_bClub2 = False
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//�����Ƃ��Ⴄ�N���u�łǂ��炩�ޕ����Ă����ꍇ�A�N���u2�ɓo�^
		If gf_SetNull2String(p_iFlg2) = "2" And gf_SetNull2String(p_iFlg1) <> "2" Then
'response.write "�����Ƃ��Ⴄ�N���u�łǂ��炩�ޕ����Ă����ꍇ�A�N���u2�ɓo�^"
			p_bClub1 = False
			p_bClub2 = True
			p_sClubCd = m_sClubCd
			Exit Function
		End If

		'//�����Ƃ��Ⴄ�N���u�ŗ����Ƃ��ޕ����Ă����ꍇ
		If gf_SetNull2String(p_iFlg1) = "2" And gf_SetNull2String(p_iFlg2) = "2" Then
'response.write "�����Ƃ��Ⴄ�N���u�ŗ����Ƃ��ޕ����Ă����ꍇ"

			'��ɑޕ��������A�N���u1�ɓo�^
			If gf_SetNull2String(p_sTaiBi1) < gf_SetNull2String(p_sTaiBi2) Then
'response.write "��ɑޕ��������A�N���u1�ɓo�^"
				p_bClub1 = True
				p_bClub2 = False
				p_sClubCd = m_sClubCd
				Exit Function
			End If

			'��ɑޕ��������A�N���u2�ɓo�^
			If gf_SetNull2String(p_sTaiBi1) > gf_SetNull2String(p_sTaiBi2) Then
'response.write "��ɑޕ��������A�N���u2�ɓo�^"
				p_bClub1 = False
				p_bClub2 = True
				p_sClubCd = m_sClubCd
				Exit Function
			End If
		End If
'response.write "�G���[�I�I�I�I�I�I�I�I�I�I�I�I"

End Function

'********************************************************************************
'*  [�@�\]  �폜���A�폜�ΏۃN���u���N���u1���A�N���u2���𒲍�����
'*  [����]  p_T13_CLUB1:T13_CLUB_1
'*          p_T13_CLUB2:T13_CLUB_2
'*	[�ߒl]  p_bClub1=True : Club1�X�V�� p_bClub1=False : Club1�X�V�s��
'*		    p_bClub2=True : Club2�X�V�� p_bClub2=False : Club2�X�V�s��
'*          p_sClubCd:�o�^����CD��Ԃ�
'*  [����]  
'********************************************************************************
Function f_DelDataSet(p_T13_CLUB1,p_T13_CLUB2,p_bClub1,p_bClub2,p_sClubCd)

		'//������
		p_bClub1 = False
		p_bClub2 = False
		p_sClubCd = ""

		'//����N���u���ޕ��Ώہi�N���u1�j
		If gf_SetNull2String(p_T13_CLUB1) = m_sClubCd Then
			p_bClub1 = True
			p_bClub2 = False

			'�N���u�����폜�����ɑޕ����Ɠ����ޕ��t���O=2�ɂ���ׂɁA�N���u�R�[�h��Ԃ��B�@2001/12/11 �ɓ�
			'p_sClubCd = ""
			p_sClubCd = m_sClubCd
		Else

			'//�A����N���u���ޕ��Ώہi�N���u2�j
			If gf_SetNull2String(p_T13_CLUB2) = m_sClubCd Then
				p_bClub1 = False
				p_bClub2 = True
				'�N���u�����폜�����ɑޕ����Ɠ����ޕ��t���O=2�ɂ���ׂɁA�N���u�R�[�h��Ԃ��B�@2001/12/11 �ɓ�
				'p_sClubCd = ""
				p_sClubCd = m_sClubCd
			Else
				p_bClub1 = False
				p_bClub2 = False
			End If

		End If

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
	<title>�����������ꗗ</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>

	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//  [�@�\]  �y�[�W���[�h������
	//  [����]
	//  [�ߒl]
	//  [����]
	//************************************************************
	function window_onload() {

		<%If m_sMode = "DELETE" Then%>

			alert("<%= "�o�^���I�����܂���" %>");

			//�ޕ��o�^���A������ʂɖ߂�
			//��t���[���ĕ\��
			parent.topFrame.location.href="./web0360_top.asp?txtClubCd=<%=m_sClubCd%>"
			//���t���[���ĕ\��
			parent.main.location.href="./web0360_main.asp?txtClubCd=<%=m_sClubCd%>"

		<%Else%>

			alert("<%= C_TOUROKU_OK_MSG %>");

			//�V�K�o�^���A�o�^��ʂɖ߂�
			var wArg
			wArg="?"
			wArg=wArg + "cboGakunenCd=<%=m_iGakunen%>"
			wArg=wArg + "&cboClassCd=<%=m_iClassNo%>"
			wArg=wArg + "&txtTyuClubCd=<%=m_sTyuClubCd%>"
			wArg=wArg + "&txtClubCd=<%=m_sClubCd%>"

			//��t���[���ĕ\��
			parent.topFrame.location.href="./web0360_insTop.asp"+wArg;
			//���t���[���ĕ\��
			parent.main.location.href="./web0360_insMain.asp"+wArg;
		<%End If%>

        return;
}

	//-->
	</SCRIPT>
	</head>
	<body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">

	<input type="hidden" name="txtClubCd" value="<%=m_sClubCd%>">

	</form>
	</body>
	</html>
<%
End Sub
%>