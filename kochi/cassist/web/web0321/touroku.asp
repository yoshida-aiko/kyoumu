<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �g�p���ȏ��o�^
' ��۸���ID : web/WEB0321/default.asp
' �@	  �\: �g�p���ȏ��̓o�^���s��
'-------------------------------------------------------------------------
' ��	  ��:�����R�[�h 	��		SESSION���i�ۗ��j
' ��	  ��:�Ȃ�
' ��	  �n:�����R�[�h 	��		SESSION���i�ۗ��j
' ��	  ��:
'			���t���[���y�[�W
'-------------------------------------------------------------------------
' ��	  ��: 2001/07/05 �≺�@�K��Y
' ��	  �X: 2001/07/23 �{���@��
' ��	  �X: 2001/07/31 �ɓ��@���q
' ��	  �X: 2001/08/01 �O�c�@�q�j
' ��	  �X: 2001/08/18 �ɓ��@���q ���N�x�̊w����񂪂Ȃ����͎��N�x�̓��͂��o���Ȃ��悤�ɂ���
' ��	  �X: 2001/12/01 �c���@��K �����̏�������w�Ȃ݂̂�ύX�ł���悤�ɏC��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n

	Public m_sGakkiWhere	'�w���̏���
	Public m_sGakkaWhere	'�w�ȃR���{�̏���
	Public m_sKamokuWhere	'�Ȗڂ̏���
	Public m_sKamokuOption	'�Ȗڂ̃I�v�V����
	Public m_sCourseWhere	'�ȖڃR�[�X�̏���
	Public m_sCourseOption	'�ȖڃR�[�X�̃I�v�V����
	Public m_bErrFlg		'�װ�׸�
	Public m_iNendo 		'�N�x
	Public m_sKyokan_CD 	'����CD
	Public m_iMax
	Public m_iDsp
	Public m_sPageCD
	Public m_sTitle 		''�V�K�o�^�E�C���̕\���p
	Public m_sDBMode		''DB�ւ̍X�VӰ��
	Public m_sMode			''��ʂ̕\����Ӱ��
	Public m_sKengen	''����(FULLorNOMAL)
	
	''�ް��\���p
	Public m_sNo
	Public m_sNendo
	Public m_sGakkiCD
	Public m_sGakunenCD
	Public m_sGakkaCD
	Public m_sKamokuCD
	Public m_sCourseCD
	Public m_sKyokan_NAME		'����
	Public m_sKyokasyo_NAME 	'���ȏ�
	Public m_sSyuppansya		'�o�Ŏ�
	Public m_sTyosya			'���Җ�
	Public m_sSidousyo			'�w����
	Public m_sKyokanyo			'�����p
	Public m_sBiko				'���l

	Public m_sNendoOption
	Public m_bJinendoGakki		'//���N�x�̊w����񂪂��邩�ǂ���

	Public m_sSyozokuGakka		'//2001/12/01 Add ���O�C�����������̏�������w��

	Public m_sGetSQL			'2001/12/01 Add

'///////////////////////////���C������/////////////////////////////

	'Ҳ�ٰ�ݎ��s
	Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

Sub Main()
'********************************************************************************
'*	[�@�\]	�{ASP��Ҳ�ٰ��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	Dim w_iRet				'// �߂�l

	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="�A�E��}�X�^�o�^"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False
	m_iDsp = C_PAGE_LINE

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

		'// �l��ϐ��ɓ����
		Call s_SetParam()


		'// �X�V�p���ް���\������
		if m_sMode = "Kousin" then
			if f_GetData() = False then
				exit do
			end if
		end if

		'// �����̖��̂��擾����
		if f_GetData_Kyokan() = False then
			exit do
		end if

		'//���N�x��񂪂��邩�`�F�b�N
		w_iRet = f_GetJinendoGakki(m_bJinendoGakki)
		If w_iRet  = False Then
			m_bErrFlg = True
				exit do
		End If

		'�w���Ɋւ���WHRE���쐬����
		Call f_MakeGakkiWhere() 
		'�w�ȂɊւ���WHRE���쐬����
		Call f_MakeGakkaWhere()
		'�w�ȃR�[�X�Ɋւ���WHRE���쐬����
		Call f_MakeCourseWhere()
		'�ȖڂɊւ���WHRE���쐬����
		Call f_MakeKamokuWhere()

		'//�������擾
		w_iRet = gf_GetKengen_WEB0320(m_sKengen)
		If w_iRet <> 0 Then
			m_bErrFlg = true
			w_sMsg = "����������܂���B"
			Exit Do
		End If

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
'*	[�@�\]	���N�x�̊w����񂪂��邩�ǂ����`�F�b�N����
'*	[����]	�Ȃ�
'*	[�ߒl]	p_bJinendoGakki=true:�w����񂠂�
'*			p_bJinendoGakki=false:�w�����Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetJinendoGakki(p_bJinendoGakki)
	Dim w_iRet				'// �߂�l
	Dim w_sSQL				'// SQL��
	dim w_Rs

	on error resume next
	err.clear

	f_GetJinendoGakki = False
	p_bJinendoGakki = False

	'//���N�x�̊w����񂪂��邩�ǂ���
	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI"
	w_sSQL = w_sSQL & vbCrLf & " FROM M01_KUBUN"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_NENDO=" & cint(SESSION("NENDO"))+1
	w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & C_KAISETUKI

	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		m_bErrFlg = True
		Exit Function
	End If

	'//�f�[�^����������
	If w_Rs.EOF = False Then
		p_bJinendoGakki = True
	End If

	Call gf_closeObject(w_Rs)

	f_GetJinendoGakki = True

End Function

'********************************************************************************
'*	[�@�\]	�l��ϐ��ɓ����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_SetParam()

	m_iNendo	 = Session("NENDO")
	m_sMode 	 = Request("txtMode")		':���[�h

	'//���پ��
	if m_sMode = "Touroku" then
		m_sTitle = "�V�K�o�^"
	elseif m_sMode = "Kousin" then
		m_sTitle = "�C��"
	else
		m_sTitle = Request("txtTitle")	'//�����[�h��
	end if

	''DB�̓o�^����Ӱ�ނ̐ݒ�
	if m_sTitle ="�V�K�o�^" then
		m_sDBMode="Insert"
		m_sNendoOption = ""
	else
		m_sDBMode="Update"
		m_sNendoOption = "DISABLED"
	end if

	'//�ꗗ�\�����y�[�W��ۑ�
	m_sPageCD	 = Request("txtPageCD")

	''�ް��\���p
	if m_sMode = "Touroku" then
		m_sNendo = Session("NENDO")
		m_sNo = ""		''�X�V�pNo�i�[
		m_sKyokan_CD = session("KYOKAN_CD")

	elseif m_sMode = "Kousin" then
		m_sNendo = Session("NENDO")
		m_sNo = Request("txtUpdNo") 	''�X�V�pNo�i�[
		m_sKyokan_CD = Request("SKyokanCd1")
	else	'//�����[�h��
		m_sNendo  = Request("txtNendo")
		m_sNo = Request("txtUpdNo") 	''�X�V�pNo�i�[
		m_sKyokan_CD = Request("SKyokanCd1")
	end if

	'm_sKyokan_CD = Session("KYOKAN_CD")
	'm_sKyokan_CD = Request("SKyokanCd1")

	m_sGakkiCD	 = Request("txtGakkiCD")
	m_sGakunenCD = Request("txtGakunenCD")


	m_sGakkaCD	 = Request("txtGakkaCD")
	m_sKamokuCD  = Request("txtKamokuCD")
	m_sCourseCD  = Request("txtCourseCD")
	m_sKyokan_NAME	= Request("txtKyokanMei")		'����
	m_sKyokasyo_NAME  = Request("txtKyokasyoName")	'���ȏ�
	m_sSyuppansya  = Request("txtSyuppansya")		'�o�Ŏ�
	m_sTyosya  = Request("txtTyosya")				'���Җ�
	m_sSidousyo  = Request("txtSidousyo")			'�w����
	m_sKyokanyo  = Request("txtKyokanyo")			'�����p
	m_sBiko  = trim(Request("txtBiko")) 				  '���l

End Sub

'********************************************************************************
'*	[�@�\]	�f�o�b�O�p
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iNendo		= " & m_iNendo			& "<br>"
	response.write "m_sMode			= " & m_sMode			& "<br>"
	response.write "m_sTitle		= " & m_sTitle			& "<br>"
	response.write "m_sDBMode		= " & m_sDBMode			& "<br>"
	response.write "m_sPageCD		= " & m_sPageCD			& "<br>"
	response.write "m_sNendo		= " & m_sNendo			& "<br>"
	response.write "m_sNo			= " & m_sNo				& "<br>"
	response.write "m_sKyokan_CD	= " & m_sKyokan_CD		& "<br>"
	response.write "m_sGakkiCD		= " & m_sGakkiCD		& "<br>"
	response.write "m_sGakunenCD	= " & m_sGakunenCD		& "<br>"
	response.write "m_sGakkaCD		= " & m_sGakkaCD		& "<br>"
	response.write "m_sKamokuCD		= " & m_sKamokuCD		& "<br>"
	response.write "m_sCourseCD		= " & m_sCourseCD		& "<br>"
	response.write "m_sKyokan_NAME	= " & m_sKyokan_NAME	& "<br>"
	response.write "m_sKyokasyo_NAME= " & m_sKyokasyo_NAME	& "<br>"
	response.write "m_sSyuppansya	= " & m_sSyuppansya		& "<br>"
	response.write "m_sTyosya		= " & m_sTyosya			& "<br>"
	response.write "m_sSidousyo		= " & m_sSidousyo		& "<br>"
	response.write "m_sKyokanyo		= " & m_sKyokanyo		& "<br>"
	response.write "m_sBiko			= " & m_sBiko			& "<br>"

End Sub

'********************************************************************************
'*	[�@�\]	�����̖��̂��擾����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
function f_GetData_Kyokan()
	Dim w_iRet				'// �߂�l
	Dim w_sSQL				'// SQL��
	dim w_Rs

	f_GetData_Kyokan = False

	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & " M04.M04_NENDO "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKAN_CD "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN M04 "
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "    M04_NENDO = " &  m_iNendo & " AND "
	w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN_CD = '" & m_sKyokan_CD & "'"

	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		m_bErrFlg = True
		Exit Function
	Else
		'�y�[�W���̎擾
		m_iMax = gf_PageCount(w_Rs,m_iDsp)
	End If

	m_sKyokan_NAME = ""
	If w_Rs.EOF = False Then
		m_sKyokan_NAME = w_Rs("M04_KYOKANMEI_SEI") & "  " & w_Rs("M04_KYOKANMEI_MEI")
	End If

	w_Rs.close

	f_GetData_Kyokan = True

end function

'********************************************************************************
'*	[�@�\]	�X�V���̕\���ް����擾����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
function f_GetData()
	Dim w_iRet				'// �߂�l
	Dim w_sSQL				'// SQL��
	Dim w_Rs

	f_GetData = False

	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & " T47.T47_NENDO "			''�N�x
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKI_KBN "		''�w���敪
'	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_NO"				''No
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKUNEN " 		''�w�N
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKA_CD "		''�w��
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_COURSE_CD "		''�������
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KAMOKU "			''�Ȗں���
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKASYO "		''���ȏ���
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SYUPPANSYA "		''�o�Ŏ�
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_TYOSYA "			''����
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKANYOUSU " 	''�����p��
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SIDOSYOSU "		''�w������
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_BIKOU "			''���l
	w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKAN "			 ''����
	w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKAMEI "
	w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
	w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "    T47_KYOKASYO T47 "
	w_sSQL = w_sSQL & vbCrLf & "    ,M02_GAKKA M02 "
	w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
	w_sSQL = w_sSQL & vbCrLf & "    ,M04_KYOKAN M04 "
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M02.M02_NENDO(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_GAKKA_CD  = M02.M02_GAKKA_CD(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M03.M03_NENDO(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M04.M04_NENDO(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = M04.M04_KYOKAN_CD(+) AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO = " & Request("KeyNendo") & " AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = '" & m_sKyokan_CD & "' AND "
	w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NO = " & m_sNo & ""

response.write(w_sSQL & "<BR>")
	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		m_bErrFlg = True
		Exit Function
	Else
		'�y�[�W���̎擾
		m_iMax = gf_PageCount(w_Rs,m_iDsp)
	End If
response.write("Set<BR>")

	m_sNendo   = gf_HTMLTableSTR(w_Rs("T47_NENDO"))
	m_sGakkiCD	 = gf_HTMLTableSTR(w_Rs("T47_GAKKI_KBN"))
	m_sGakunenCD = gf_HTMLTableSTR(w_Rs("T47_GAKUNEN"))
	m_sGakkaCD	 = gf_HTMLTableSTR(w_Rs("T47_GAKKA_CD"))
	m_sKamokuCD  = gf_HTMLTableSTR(w_Rs("T47_KAMOKU"))
	m_sCourseCD  = gf_HTMLTableSTR(w_Rs("T47_COURSE_CD"))
	m_sKyokasyo_NAME  = gf_HTMLTableSTR(w_Rs("T47_KYOKASYO"))		'���ȏ�
	m_sSyuppansya  = gf_HTMLTableSTR(w_Rs("T47_SYUPPANSYA"))		'�o�Ŏ�
	m_sTyosya  = gf_HTMLTableSTR(w_Rs("T47_TYOSYA"))				'���Җ�
	m_sSidousyo  = gf_HTMLTableSTR(w_Rs("T47_SIDOSYOSU"))			'�w����
	m_sKyokanyo  = gf_HTMLTableSTR(w_Rs("T47_KYOKANYOUSU")) 		'�����p
	m_sBiko  = gf_HTMLTableSTR(w_Rs("T47_BIKOU"))					'���l

	m_sKyokan_CD = gf_HTMLTableSTR(w_Rs("T47_KYOKAN"))

	w_Rs.close

	f_GetData = True
response.write("f_GetData<BR>")

end function


'********************************************************************************
'*	[�@�\]	�w���R���{�Ɋւ���WHRE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub f_MakeGakkiWhere()
Dim w_sNendo
	m_sGakkiWhere=""

	'//�V�K�o�^���A���N�x��񂪂���Ƃ��͎��N�x���g�p�B�Ȃ����͓��N�x���g�p�B
'	If m_bJinendoGakki = True Then
'		w_sNendo = cint(m_iNendo) + 1
'	Else
'		w_sNendo = cint(m_iNendo)
'	End If

	w_sNendo = cint(request("txtNendo"))

	'm_sGakkiWhere = " M01_DAIBUNRUI_CD = 51  AND "
	m_sGakkiWhere = " M01_DAIBUNRUI_CD = " & C_KAISETUKI & " AND "
	m_sGakkiWhere = m_sGakkiWhere & " M01_SYOBUNRUI_CD <> 3 AND "	'<--"�J�݂��Ȃ�"�ȊO
	If m_sMode = "Touroku" Then
	  m_sGakkiWhere = m_sGakkiWhere & " M01_NENDO = " & w_sNendo  & ""
	Else
	  m_sGakkiWhere = m_sGakkiWhere & " M01_NENDO = " & m_sNendo & ""
	End If

'response.write m_sGakkiWhere & "<BR>"

End Sub

'********************************************************************************
'*	[�@�\]	�w�ȃR���{�Ɋւ���WHRE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub f_MakeGakkaWhere()
	Dim w_sNendo

	'2001/12/01 Add ---->
	Dim w_sSQL				'//SQL��
	Dim w_iRet				'//�߂�l

	Dim w_oRecord			'//�����w�Ȏ擾�̂���

	'//�����w�Ȃ̎擾
	w_sSQL = ""
	w_sSQL = w_sSQL & "SELECT "
	w_sSQL = w_sSQL & "M04_GAKKA_CD "
	w_sSQL = w_sSQL & "From "
	w_sSQL = w_sSQL & "M04_KYOKAN "
	w_sSQL = w_sSQL & "Where "
	w_sSQL = w_sSQL & "M04_NENDO = " & m_iNendo & " "
	w_sSQL = w_sSQL & "And "
	w_sSQL = w_sSQL & "M04_KYOKAN_CD = '" & Session("KYOKAN_CD") & "'"

	w_iRet = gf_GetRecordset(w_oRecord, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit Sub
	End If

	If w_oRecord.EOF <> True Then
		m_sSyozokuGakka = w_oRecord("M04_GAKKA_CD")
	Else
		m_sSyozokuGakka =""
	End If

	'//����
	w_oRecord.Close
	Set w_oRecord = Nothing

	'2001/12/01 Add <----

	'//�V�K�o�^���A���N�x��񂪂���Ƃ��͎��N�x���g�p�B�Ȃ����͓��N�x���g�p�B
'	If m_bJinendoGakki = True Then
'		w_sNendo = cint(m_iNendo) + 1
'	Else
'		w_sNendo = cint(m_iNendo)
'	End If

	w_sNendo = cint(request("txtNendo"))

	m_sGakkaWhere=""

	If m_sMode = "Touroku" Then
		m_sGakkaWhere = " M02_NENDO = " & w_sNendo	& ""
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD <> '00' "
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD = '" & m_sSyozokuGakka & "' "	'2001/12/01 Mod
	Else
		m_sGakkaWhere = " M02_NENDO = " & m_sNendo & ""
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD <> '00' "
		m_sGakkaWhere = m_sGakkaWhere & " AND M02_GAKKA_CD = '" & m_sSyozokuGakka & "' "	'2001/12/01 Mod
	End If

End Sub

'********************************************************************************
'*	[�@�\]	�w�ȃR�[�X�R���{�Ɋւ���WHRE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub f_MakeCourseWhere()
Dim w_sNendo

	'w_sNendo = cint(m_iNendo) + 1

	'//�V�K�o�^���A���N�x��񂪂���Ƃ��͎��N�x���g�p�B�Ȃ����͓��N�x���g�p�B
'	If m_bJinendoGakki = True Then
'		w_sNendo = cint(m_iNendo) + 1
'	Else
'		w_sNendo = cint(m_iNendo)
'	End If

	w_sNendo = cint(request("txtNendo"))

	m_sCourseWhere=""
	m_sCourseOption=""

	If m_sMode = "Touroku" Then
		m_sCourseOption = " DISABLED "
		m_sCourseWhere = " M20_NENDO = " & w_sNendo  & ""
		Exit Sub
	End If

	''�w�Ȗ��I�����́A�w�Ⱥ���͖��I��
	'if m_sGakkaCD = "@@@" then
	if m_sGakkaCD = "@@@" Or m_sGakkaCD = "" then
		m_sCourseOption = " DISABLED "
		m_sCourseWhere = " M20_NENDO = " & w_sNendo  & ""
		m_sCourseCD = "@@@"
		Exit Sub
	end if

	''�S�w�Ȃ̎��́A�w�Ⱥ���͎g�p�s��
	if cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) then
		m_sCourseOption = " DISABLED "
		m_sCourseWhere = " M20_NENDO = " & w_sNendo  & ""
		m_sCourseCD = "@@@"
		Exit Sub
	end if


''	If m_sGakkaCD = 99 Then
''		m_sCourseWhere= " M20_NENDO = " & m_sNendo & " AND "
''		m_sCourseWhere = m_sCourseWhere & " M20_GAKUNEN =  " & m_sGakunenCD & ""
''		m_sCourseWhere = m_sCourseWhere & " Group By M20_GAKKA_CD , M20_COURSE_CD, M20_COURSEMEI "
''	Else
		m_sCourseWhere = " M20_NENDO = " & m_sNendo & " AND "
		m_sCourseWhere = m_sCourseWhere & " M20_GAKKA_CD = " & m_sGakkaCD & " AND "
		m_sCourseWhere = m_sCourseWhere & " M20_GAKUNEN =  " & m_sGakunenCD & ""
		m_sCourseWhere = m_sCourseWhere & " Group By M20_GAKKA_CD , M20_COURSE_CD, M20_COURSEMEI "
''	End IF

End Sub

'********************************************************************************
'*	[�@�\]	�ȖڃR���{�Ɋւ���WHRE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub f_MakeKamokuWhere()

	m_sKamokuWhere=""
	m_sKamokuOption=""

	'//�V�K�o�^���A���N�x��񂪂���Ƃ��͎��N�x���g�p�B�Ȃ����͓��N�x���g�p�B
'	If m_bJinendoGakki = True Then
'		w_sNendo = cint(m_iNendo) + 1
'	Else
'		w_sNendo = cint(m_iNendo)
'	End If

	w_sNendo = cint(request("txtNendo"))

	m_sGetSQL = ""
	m_sGetSQL = m_sGetSQL & "Select "
	m_sGetSQL = m_sGetSQL & "Distinct "
	m_sGetSQL = m_sGetSQL & "T15_KAMOKU_CD, "
	m_sGetSQL = m_sGetSQL & "T15_KAMOKUMEI "
	m_sGetSQL = m_sGetSQL & "From "
	m_sGetSQL = m_sGetSQL & "T15_RISYU, "
	m_sGetSQL = m_sGetSQL & "T27_TANTO_KYOKAN "
	m_sGetSQL = m_sGetSQL & "Where "

	If m_sGakunenCD <> "" Then
		'//�w�N���w�肳��Ă���ꍇ
		m_sGetSQL = m_sGetSQL & "T15_NYUNENDO = " & (cint(w_sNendo) - cint(m_sGakunenCD) + 1) & " "
	Else
		'//�w�N���w�肳��Ă��Ȃ��ꍇ
		m_sGetSQL = m_sGetSQL & "T15_NYUNENDO = " & cint(w_sNendo) & " "
	End If

	If m_sGakkaCD <> "" Then
		'//�w�Ȃ��w�肳��Ă���ꍇ
		If cstr(m_sGakkaCD) = cstr(C_CLASS_ALL) Then
			'�S�w�Ȃ̏ꍇ
			m_sGetSQL = m_sGetSQL & " AND T15_KAMOKU_KBN = " & C_KAMOKU_IPPAN

			If m_sCourseCD <> "@@@" AND m_sCourseCD <> "" Then
				m_sGetSQL = m_sGetSQL & " AND T15_COURSE_CD = " & m_sCourseCD &""
			End IF

			m_sGetSQL = m_sGetSQL & " AND T15_KAISETU" & m_sGakunenCD & "="  & m_sGakkiCD

		Else
			'�ʂ̊w��
			m_sGetSQL = m_sGetSQL & " AND T15_GAKKA_CD = '" & m_sGakkaCD &"' "
			m_sGetSQL = m_sGetSQL & " AND T15_KAISETU" & m_sGakunenCD & " = " & m_sGakkiCD
			m_sGetSQL = m_sGetSQL & " AND T15_KAMOKU_KBN <> " & C_KAMOKU_IPPAN & " "

			If cstr(gf_SetNull2String(m_sCourseCD)) <> "@@@" AND trim(cstr(gf_SetNull2String(m_sCourseCD))) <> "" Then
				m_sGetSQL = m_sGetSQL & " AND T15_COURSE_CD = " & m_sCourseCD &""
			End If

		End If

	End If
	
	m_sGetSQL = m_sGetSQL & " AND "
	m_sGetSQL = m_sGetSQL & "T27_NENDO = " & w_sNendo & " "
	m_sGetSQL = m_sGetSQL & " AND "
	m_sGetSQL = m_sGetSQL & "T15_KAMOKU_CD = T27_KAMOKU_CD "
	m_sGetSQL = m_sGetSQL & " AND "
	m_sGetSQL = m_sGetSQL & "T27_KYOKAN_CD = " & Session("KYOKAN_CD")

	If m_sGakunenCD <> "" Then
		'//�w�N���w�肳��Ă���ꍇ
		m_sGetSQL = m_sGetSQL & " AND "
		m_sGetSQL = m_sGetSQL & "T27_GAKUNEN = " & m_sGakunenCD & " "
	End If

	m_sGetSQL = m_sGetSQL & " Group By T15_NYUNENDO , T15_KAMOKU_CD , T15_KAMOKUMEI "

	''�V�K�o�^��
	If m_sMode = "Touroku" Then
		m_sKamokuOption = " DISABLED "
		Exit Sub
	End If

	''�w�Ȗ��I�����́A�Ȗڂ͖��I��
	if m_sGakkaCD = "@@@" Or m_sGakkaCD = "" then
		m_sKamokuOption = " DISABLED "
		m_sKamokuCD = "@@@"
		Exit Sub
	end if

End Sub


'****************************************************
'[�@�\] �f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'		(���X�g�_�E���{�b�N�X�I��\���p)
'[����] pData1 : �f�[�^�P
'		pData2 : �f�[�^�Q
'[�ߒl] f_Selected : "SELECTED" OR ""
'					
'****************************************************
Function f_Selected(pData1,pData2)

	f_Selected = ""

	If IsNull(pData1) = False And IsNull(pData2) = False Then
		If trim(cStr(pData1)) = trim(cstr(pData2)) Then
			f_Selected = "selected" 
		Else
		End If
	End If

End Function


Sub showPage()
'********************************************************************************
'*	[�@�\]	HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
%>

<html>

<head>
<!-- <%= m_sGetSQL %> -->


<title>�g�p���ȏ��o�^</title>

	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	//************************************************************
	//	[�@�\]	�g�p���ȏ��o�^
	//	[����]	p_iPage :�\���Ő�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_touroku(){

		// ���͒l������
		iRet = f_CheckData();
		if( iRet != 0 ){
			return;
		}

		document.frm.txtKyokanMei.value=document.frm.SKyokanNm1.value

		document.frm.action="./kakunin.asp";
		document.frm.target="";
		document.frm.txtMode.value = "<%=m_sDBMode%>";
		document.frm.submit();
	}

	//************************************************************
	//	[�@�\]	�i�H���C�����ꂽ�Ƃ��A�ĕ\������
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_ReLoadMyPage(){

		document.frm.action="./touroku.asp";
		document.frm.target="";
		document.frm.txtMode.value = "Reload";
		document.frm.submit();
	
	}

	//************************************************************
	//	[�@�\]	���C���y�[�W�֖߂�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Back(){

		document.frm.action="./default.asp";
		document.frm.target="";
		document.frm.txtMode.value = "Back";
		document.frm.submit();
	
	}

	//************************************************************
	//	[�@�\]	���͒l������
	//	[����]	�Ȃ�
	//	[�ߒl]	0:����OK�A1:�����װ
	//	[����]	���͒l��NULL�����A�p���������A�����������s��
	//			���n�ް��p���ް������H����K�v������ꍇ�ɂ͉��H���s��
	//************************************************************
	function f_CheckData() {
	
	
		// ������NULL����������
		// ���w�ȃR�[�h
		if( f_Trim(document.frm.txtGakkaCD.value) == "@@@" ){
			window.alert("�w�Ȃ��I������Ă��܂���");
			if(document.frm.txtGakkaCD.length!=1){
			document.frm.txtGakkaCD.focus();
			}
			return 1;
		}

		// ������NULL����������
		// ���ȖڃR�[�h
		if( f_Trim(document.frm.txtKamokuCD.value) == "@@@" ){
			window.alert("�Ȗڂ��I������Ă��܂���");
			if(document.frm.txtKamokuCD.length!=1){
				document.frm.txtKamokuCD.focus();
			}
			return 1;
		}

		// ������NULL����������
		// �����ȏ���
		if( f_Trim(document.frm.txtKyokasyoName.value) == "" ){
			window.alert("���ȏ��������͂���Ă��܂���");
			document.frm.txtKyokasyoName.focus();
			return 1;
		}

		// ���������ȏ����̌�����������
		if( getLengthB(document.frm.txtKyokasyoName.value) > "80" ){
			window.alert("���ȏ����̗��͑S�p40�����ȓ��œ��͂��Ă�������");
			document.frm.txtKyokasyoName.focus();
			return 1;
		}

		// �������o�ŎЖ��̌�����������
		if( getLengthB(document.frm.txtSyuppansya.value) > "40" ){
			window.alert("�o�ŎЖ��̗��͑S�p20�����ȓ��œ��͂��Ă�������");
			document.frm.txtSyuppansya.focus();
			return 1;
		}

		// ���������Җ��̌�����������
		if( getLengthB(document.frm.txtTyosya.value) > "40" ){
			window.alert("���Җ��̗��͑S�p20�����ȓ��œ��͂��Ă�������");
			document.frm.txtTyosya.focus();
			return 1;
		}

		// �����������p�����̒l����������
		if(f_Trim(document.frm.txtKyokanyo.value)!=""){
			//���l�`�F�b�N
			if( isNaN(document.frm.txtKyokanyo.value)){
				window.alert("�����͔��p�����œ��͂��Ă�������");
				document.frm.txtKyokanyo.focus();
				return 1;
			}else{
				//���`�F�b�N
				if( getLengthB(document.frm.txtKyokanyo.value) > "3" ){
					window.alert("������3���ȓ��œ��͂��Ă�������");
					document.frm.txtKyokanyo.focus();
					return 1;
				}
			}
		}

		// �������w���p�����̒l����������
		if(f_Trim(document.frm.txtSidousyo.value)!=""){
			//���l�`�F�b�N
			if( isNaN(document.frm.txtSidousyo.value)){
				window.alert("�����͔��p�����œ��͂��Ă�������");
				document.frm.txtSidousyo.focus();
				return 1;
			}else{
				//���`�F�b�N
				if( getLengthB(document.frm.txtSidousyo.value) > "3" ){
					window.alert("������3���ȓ��œ��͂��Ă�������");
					document.frm.txtSidousyo.focus();
					return 1;
				}
			}
		}

		// ���������l�̌�����������
		if( getLengthB(document.frm.txtBiko.value) > "80" ){
			window.alert("���l�̗��͑S�p40�����ȓ��œ��͂��Ă�������");
			document.frm.txtBiko.focus();
			return 1;
		}

		return 0;
	}

	//************************************************************
	//	[�@�\]	�����Q�ƑI����ʃE�B���h�E�I�[�v��
	//	[����]
	//	[�ߒl]
	//	[����]
	//************************************************************
	function KyokanWin(p_iInt,p_sKNm) {
		var obj=eval("document.frm."+p_sKNm)
		var w_gak = document.frm.txtGakkaCD.value
		URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"&txtGakka="+w_gak+"";
		nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=600,top=0,left=0");
		nWin.focus();
		return true;	
	}
	//************************************************************
	//	[�@�\]	�N���A�{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function fj_Clear(){
		//���������󔒂ɂ���
		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}

	//-->
	</script>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">

	</head>
	<body>
	<form name="frm" action="" target="" method="post">

<%'call s_DebugPrint%>

	<center>
	<% call gs_title("�g�p���ȏ��o�^",m_sTitle) %>
	<br>
<table border="0" cellpadding="1" cellspacing="1" width="540">
	<tr>
		<td align="left">
			<table width="100%" border=1 CLASS="hyo">
				<tr>
				<th height="16" width="75" class=header nowrap>�N�@�x</th>
				<td height="16" width="430" class=detail nowrap>
				<%If m_sDBMode="Update" Then%>
					<%=Request("KeyNendo")%>
					<input type="hidden" name="txtNendo" value="<%=Request("KeyNendo")%>">
				<%Else%>
					<select name="txtNendo" onchange='javascript:f_ReLoadMyPage()'	>

						<%'//�V�K�o�^���A���N�x��񂪂���Ƃ��͎��N�x���g�p�B�Ȃ����͓��N�x���g�p�B
						If m_bJinendoGakki = True Then
							'//�V�K�o�^���A���N�x��񂪂���Ƃ��͎��N�x���g�p�B�Ȃ����͓��N�x���g�p�B
							If m_sMode = "Touroku" Then
								w_sNendo = cint(m_iNendo) + 1
							Else
								w_sNendo = m_sNendo
							End If
						%>
							<option VALUE="<%= m_iNendo + 1 %>" <%= f_Selected(cstr(w_sNendo),cstr(cint(m_iNendo+1)))%>><%= m_iNendo + 1 %>
							<option VALUE="<%= m_iNendo %>" 	<%= f_Selected(cstr(w_sNendo),cstr(m_iNendo))%>><%= m_iNendo %>
						<%Else%>
							<option VALUE="<%= m_iNendo %>" 	<%= f_Selected(cstr(m_sNendo),cstr(m_iNendo))%>><%= m_iNendo %>
						<%End If%>
					</select><span class=hissu>*</span>
				<%End If%>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class="header" nowrap>�w�@��</th>
				<td height="16" width="430" class="detail" nowrap>

				<%'���ʊ֐�����w���Ɋւ���R���{�{�b�N�X���o�͂���
				If m_sMode = "Touroku" Then
						call gf_ComboSet("txtGakkiCD",C_CBO_M01_KUBUN,m_sGakkiWhere,"onchange = 'javascript:f_ReLoadMyPage()'",False,0)
					Else
						call gf_ComboSet("txtGakkiCD",C_CBO_M01_KUBUN,m_sGakkiWhere,"onchange = 'javascript:f_ReLoadMyPage()'",False,m_sGakkiCD)
				End If
				%><span class=hissu>*</span>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>�w�@�N</th>
				<td height="16" width="430" class=detail nowrap>
					<select name="txtGakunenCD" onchange = 'javascript:f_ReLoadMyPage()'>
						<option Value="1" <%= f_Selected( 1 ,m_sGakunenCD) %>>1�N
						<option Value="2" <%= f_Selected( 2 ,m_sGakunenCD) %>>2�N
						<option Value="3" <%= f_Selected( 3 ,m_sGakunenCD) %>>3�N
						<option Value="4" <%= f_Selected( 4 ,m_sGakunenCD) %>>4�N
						<option Value="5" <%= f_Selected( 5 ,m_sGakunenCD) %>>5�N
					</select><span class=hissu>*</span>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>�w�@��</th>
				<td height="16" width="430" class=detail nowrap>
				<%	'���ʊ֐�����w�ȂɊւ���R���{�{�b�N�X���o�͂���
					call f_ComboSet_Gakka("txtGakkaCD",C_CBO_M02_GAKKA,m_sGakkaWhere,"style='width:175px;' onchange = 'javascript:f_ReLoadMyPage()'",True,m_sGakkaCD)%>
				<span class=hissu>*</span><img src="../../image/sp.gif" width="10">
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>�R�[�X</font></th>
				<td height="16" width="430" class=detail>
				<%	'���ʊ֐�����w�ȃR�[�X�Ɋւ���R���{�{�b�N�X���o�͂���
					call gf_ComboSet("txtCourseCD",C_CBO_M20_COURSE,m_sCourseWhere,"style='width:175px;' onchange = 'javascript:f_ReLoadMyPage()'" & m_sCourseOption,True,m_sCourseCD)%>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>�ȁ@��</font></th>
				<td height="16" width="430" class=detail>
<!--
<%= m_sGakunenCD %>
<%= m_sGakkaCD %>
-->
				<%	'���ʊ֐�����ȖڂɊւ���R���{�{�b�N�X���o�͂���
					'�w�N������ �w�Ȃ����͂���Ă��Ȃ��Ƃ��́ADISABLED�ƂȂ�
					call f_ComboSet("txtKamokuCD",C_CBO_T15_RISYU,m_sKamokuWhere,"style='width:175px;'" & m_sKamokuOption,True,m_sKamokuCD)%>
				<span class=hissu>*</span>
				</td>
				</tr>
				<tr>

				<th height="16" width="80" class=header nowrap>����</font></th>
				<td height="16" width="430" class=detail nowrap>
					<input type="text" class="text" name="SKyokanNm1" VALUE='<%=m_sKyokan_NAME%>' readonly size="30">
					<input type="hidden" name="SKyokanCd1" VALUE='<%=m_sKyokan_CD%>'>
					<%
					'//�ō������҂̂ݗ��p�҂̕ύX���Ƃ���
					If m_sKengen = C_ACCESS_FULL Then%>
						<input type="button" class="button" value="�I��" onclick="KyokanWin(1,'SKyokanNm1')">
						<input type="button" class="button" value="�N���A" onClick="fj_Clear()">
					<%End If%>
				</td>
				</tr>

				<tr>
				<th height="16" width="80" class=header nowrap>���ȏ���</font></th>
				<td height="16" width="430" class=detail nowrap>
				<textarea cols="56" rows="3" Name="txtKyokasyoName" Value="<%= m_sKyokasyo_NAME %>"><%= m_sKyokasyo_NAME %></textarea>
				<span class=hissu>*</span><font size=2><BR>�i�S�p40�����ȓ��j</font>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>�o�Ŏ�</font></th>
				<td height="16" width="430"  class=detail nowrap>
				<input type="text" size="56" Name="txtSyuppansya" Value="<%= m_sSyuppansya %>"><BR><font size=2>�i�S�p20�����ȓ��j</font>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>���Җ�</font></th>
				<td height="16" width="430" class=detail nowrap>
				<input type="text" size="56" Name="txtTyosya" Value="<%= m_sTyosya %>"><BR><font size=2>�i�S�p20�����ȓ��j</font>
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>�����p</font>
				</th>
				<td height="16" width="430" class=detail nowrap>
				<input type="text" size="3" Name="txtKyokanyo" Value="<%= m_sKyokanyo %>" maxlength="3">��
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>�w����</font>
				</th>
				<td height="16" width="430" class=detail nowrap>
				<input type="text" size="3" Name="txtSidousyo" Value="<%= m_sSidousyo %>"  maxlength="3">��
				</td>
				</tr>

				<tr>
				<th height="16" width="75" class=header nowrap>���@�l</font></th>
				<td height="16" width="430" class=detail nowrap>
				<textarea cols="56" rows="3" Name="txtBiko"  Value="<%= trim(m_sBiko) %>"><%= trim(m_sBiko)%></textarea><font size=2>�i�S�p40�����ȓ��j</font>
				</td>
				</TR>
			</TABLE>
			<table width=75%><tr><td align=right><span class=hissu>*��͕K�{���ڂł��B</span></td></tr></table>
		</td>
	</TR>
</TABLE>
		<table border="0" width=300>
		<tr>
		<td valign="top" align=left>
		<input type="button" class=button value="�@�o�@�^�@" OnClick="f_touroku()">
			<img src="../../image/sp.gif" width="30" height="1">
		</td>
		<td valign="top" align=right>
		<input type="Button" class=button value="�L�����Z��" OnClick="f_Back()">
		</td>
		</tr>
		</table>

		</center>

		<input type="hidden" name="txtMode" value="Touroku">
		<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
		<input type="hidden" name="txtUpdNo" value="<%= m_sNo %>">
		<input type="hidden" name="txtTitle" value="<%= m_sTitle %>">

		<input type="hidden" name="KeyNendo" value="<%=Request("KeyNendo")%>">
		<input type="hidden" Name="txtKyokanMei" Value="<%= m_sKyokan_NAME %>">

	</form>
	</body>
	</html>

<%
End Sub

Function f_ComboSet_Gakka(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' �@	�\:ComboBox�Z�b�g
' ��	�l:OK=True/NG=False
' ��	��:p_oCombo - ComboBox
'		   p_sTableName - �e�[�u����
'		   p_sWhere - Where����(WHERE��͗v��Ȃ�)
'		   p_sSelectOption - <SELECT>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
'		   p_bWhite - �擪�ɋ󔒂����邩
'		   p_sSelectCD - �W���I�����������R�[�h(""�Ȃ�I���Ȃ�)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��	�l:�����w�Ȃ���ʑ����w�Ȃ̏ꍇ�͑S�w�Ȃ���
'*************************************************************************************
	Dim w_sId			'ID�t�B�[���h��
	Dim w_sName 		'���̃t�B�[���h��
	Dim w_sTableName	'���̃e�[�u����
	Dim w_rst

	f_ComboSet_Gakka = False

	do 
	''�}�X�^����SELECT����t�B�[���h�����擾
	If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
		Exit Do
	End If

	''�}�X�^SELECT
	If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
		Exit Do
	End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
	Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

	'�󔒂�Option�̑��
	If p_bWhite Then
		response.Write " <Option Value="&C_CBO_NULL&">�@�@�@�@�@ "& Chr(13)
	End If

	''EOF�łȂ���΁A�f�[�^���Z�b�g
	If Not w_rst.EOF Then
		Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
	End If

	'// ��ʑ����w�Ȃ̏ꍇ�͑S�w�Ȃ�I���\
	If m_sSyozokuGakka = "00" Then
		response.write(" <Option Value='" & C_CLASS_ALL & "'")
		If CStr(p_sSelectCD) = CStr(C_CLASS_ALL) Then
			response.write " Selected "
		End If
		response.Write(">" & "�S�w��" & Chr(13))
	End If

	Response.write("</select>" & chr(13))

	If Not w_rst Is Nothing Then
		w_rst.Close
		Set w_rst = Nothing
	End If
   
	f_ComboSet_Gakka = True
	Exit Do
	Loop
End Function

'/*****************************

Public Function f_ComboSet(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
	Dim w_sId			'ID�t�B�[���h��
	Dim w_sName 		'���̃t�B�[���h��
	Dim w_sTableName	'���̃e�[�u����
	Dim w_rst

	do
		'//�f�[�^�̎擾
		If gf_GetRecordset(w_rst, m_sGetSQL) <> 0 Then
			Exit Function
		End If

		If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
		Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

		'�󔒂�Option�̑��
		If p_bWhite Then
			response.Write " <Option Value="&C_CBO_NULL&">�@�@�@�@�@ "& Chr(13)
		End If

		''EOF�łȂ���΁A�f�[�^���Z�b�g
		If Not w_rst.EOF Then
			Call s_MstDataSet(p_sCombo, w_rst, "T15_KAMOKU_CD", "T15_KAMOKUMEI", p_sSelectCD)
		End If

		Response.write("</select>" & chr(13))

		If Not w_rst Is Nothing Then
			w_rst.Close
			Set w_rst = Nothing
		End If

		f_ComboSetf_ComboSet = True
		Exit Do
	Loop

End Function



%>
