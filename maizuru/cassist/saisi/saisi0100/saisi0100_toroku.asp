<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �Ď����ѓo�^
' ��۸���ID : saisi/saisi0100/saisi0100_toroku.asp
' �@      �\: �Ď������т̓��͉��
'-------------------------------------------------------------------------
' ��      ��    
'               
' ��      ��
' ��      �n
'           
'           
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2003/02/17 ���
' ��      �X: 2003/03/03 ���@�s���i���ɂ�T16,T17���X�V���Ȃ��悤�ɕύX
'           : 2003/03/06 ���@���͂��ꂽ�l�͑S��T120�ɍX�V����悤�ɕύX
' �ρ@�@�@�X�F2003/07/25 �A�c�@�������[�v & �o�^�ŗ����Ă����I���̑����X�E�E�E�̂��ߏC��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
Dim m_Rs		
Dim m_RsHyoka

Const mC_HYOKA_NO = 2
Const mC_HYOKA_KBN = 0

Const mC_KAISETU_Z   = 1
Const mC_HYOKA_CD_OK = 0
Const mC_HYOKA_CD_NO = 1

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

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�w�b�_�[�f�[�^"
    w_sMsg=""
    w_sRetURL="../../default.asp"
    w_sTarget="_parent"

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'// ���ѓo�^�̏ꍇ
		if Request("hidMode") = "update" then
			if wf_UpdateSeiseki() = false then
				m_bErrFlg = True
				m_sErrMsg = "���т̓o�^�Ɏ��s���܂����B"
				Exit Do
			end if
'Response.write "����"
'Response.end
			'REDIRECT����
			response.redirect "saisi0100_show.asp"
			response.end

		else
			'// �Ď����Y���w�����擾
			if wf_GetStudent() = false then
				m_bErrFlg = True
				m_sErrMsg = "�w�����̎擾�Ɏ��s���܂����B"
				Exit Do
			end if
		end if

		Exit Do
	Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

	'//�����\��
    Call showPage()

    '// �I������
    Call gs_CloseDatabase()

End Sub

function wf_UpdateSeiseki()
'********************************************************************************
'*  [�@�\]  ���т̔���A�o�^
'*  [����]  �Ȃ�
'*  [�ߒl]  true , false
'*  [����]  
'********************************************************************************

	'�ϐ��̐錾
	Dim w_sSql			'SQL���i�[�G���A
	Dim w_iRet			'�߂�l�i�[�G���A
	Dim i				'�C���f�b�N�X
	Dim w_bNendo		'True:���N�x�@false:�ߋ�
	Dim w_iSeiseki(5)	'���ъi�[�G���A
	Dim w_iGakusei		'�w���ԍ�
	Dim w_iNendo		'�N�x
	Dim w_iMotoTen		'���̓_��

	wf_UpdateSeiseki = false
	
	'// �]������擾
	if wf_GetHyoka() = false then
		m_bErrFlg = True
		Exit Function
	end if

	for i = 1 to Request("hidCount")

		'�ϐ��̊i�[
		w_iGakusei = Request("hidGakusei" & i)
		w_iNendo   = Request("hidNendo" & i)
		w_iMotoTen = Request("hidSeiseki" & i)
		'�z��̊i�[
		w_iSeiseki(1) = Request("txtSeiseki1_" & i)
		w_iSeiseki(2) = Request("txtSeiseki2_" & i)
		w_iSeiseki(3) = Request("txtSeiseki3_" & i)
		w_iSeiseki(4) = Request("txtSeiseki4_" & i)
		w_iSeiseki(5) = Request("txtSeiseki5_" & i)

'�f�o�b�O�v�����g�p		
'Response.Write "�ԍ���" & Request("hidGakusei" & i) & "<br>"
'Response.Write "�N�x��" & Request("hidNendo" & i) & "<br>"
'Response.Write "���_��" & Request("hidSeiseki" & i) & "<br>"
'Response.Write "����1��" & Request("txtSeiseki1_" & i) & "<br>"
'Response.Write "����2��" & Request("txtSeiseki2_" & i) & "<br>"
'Response.Write "����3��" & Request("txtSeiseki3_" & i) & "<br>"
'Response.Write "����4��" & Request("txtSeiseki4_" & i) & "<br>"
'Response.Write "����5��" & Request("txtSeiseki5_" & i) & "<br>"
'response.end

		'���N�x/�ߋ��N�x�̔���
		if Cint(Request("hidNendo" & i)) = Cint(Session("NENDO")) then
			w_bNendo = true
		else
			w_bNendo = false
		end if
'response.write w_bNendo & "<br>"

		'���N�x�Ȗڂ̏ꍇ�iT16 & T120�j
		if w_bNendo = true then
			if wf_UpdateT16(w_iSeiseki,w_iGakusei,w_iNendo,w_iMotoTen) = false then
				m_bErrFlg = True
				Exit Function
			end if
		'�ߋ��N�x�Ȗڂ̏ꍇ�iT17 & T120�j
		else
			if wf_UpdateT17(w_iSeiseki,w_iGakusei,w_iNendo,w_iMotoTen) = false then
				m_bErrFlg = True
				Exit Function
			end if
		end if

	next	

	wf_UpdateSeiseki = true

end function

function wf_UpdateT16(p_iSeiseki,p_iGakusei,p_iNendo,p_iMotoTen)
'********************************************************************************
'*  [�@�\]  T16�̐��уf�[�^�X�V
'*  [����]  p_iSeiseki	:�_���i�z��j
'*          p_iGakusei	:�w���ԍ�
'*          p_iNendo	:�N�x
'*          p_iMotoTen	:���̓_���iT16_HTEN�j
'*  [�ߒl]  true , false
'*  [����]  
'********************************************************************************

	Dim i						'�C���f�b�N�X
	Dim w_sSql					'SQL�i�[�G���A
	Dim w_Rs					'ADODB���R�[�h�Z�b�g
	Dim w_sHyokaMei				'M08_HYOKA_SYOBUNRUI_MEI
	Dim w_iHyotei				'M08_HYOTEI
	Dim w_iHyokaCd				'M08_HYOKA_SYOBUNRUI_RYAKU
	Dim w_iSeiseki				'���сi�_���i�[�G���A�j
	Dim w_sAryJugyoJikan(4)		'�����֘A

	Dim w_bAryUpdateFlg(5)		'T120�̗����X�V�t���O
	Dim w_sAryHyokaMei(5)		'T120�̗�����HYOKA
	Dim w_iAryHyotei(5)			'T120�̗�����HYOTEI

	wf_UpdateT16 = false

	'�ϐ��̏�����
	i = 5

	'*****************************************
	'* �X�V�Ώۉ񐔂̎擾(T16�ɍX�V�����l) *
	'*****************************************
	do until i = 0
		if i <> 1 then
			if p_iSeiseki(i) <> "" then
				w_iSeiseki = p_iSeiseki(i)
				exit do
			end if
		else
			if p_iSeiseki(i) <> "" then
				w_iSeiseki = p_iSeiseki(i)
				exit do
			else
				'���͂��Ȃ������̂ŃR�R�ŏI��
				wf_UpdateT16 = true
				Exit Function
			end if
		end if
		i = i - 1
	loop

	'**************************************
	'* ���͂��ꂽ�l�̔���y�сA�t���O���� *
	'**************************************
	'�z��̏�����
	for i = 1 to 5
		w_bAryUpdateFlg(i) = false		'T120�X�V�p�̃t���O
		w_sAryHyokaMei(i)  = ""			'T120_HYOKAMEI
		w_iAryHyotei(i)    = 0			'T120_HYOTEI
	next

	'�J�E���^�̐ݒ�
	i = 5

	do until i = 0

		'���͂�����
		if p_iSeiseki(i) <> "" then
			'�]��
			m_RsHyoka.MoveFirst
			Do until m_RsHyoka.EOF
				if Cint(p_iSeiseki(i)) >= Cint(m_RsHyoka("M08_MIN")) then
					w_sAryHyokaMei(i) = m_RsHyoka("M08_HYOKA_SYOBUNRUI_MEI")		'HYOKA�i�[
					w_iAryHyotei(i)   = m_RsHyoka("M08_HYOTEI")						'HYOTEI�i�[
					exit do
				end if
				m_RsHyoka.Movenext
			loop
			w_bAryUpdateFlg(i) = true
		end if
		i = i - 1
	loop

	'******************
	'* ��{�f�[�^�擾 *
	'******************
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T16_KAISETU, "
	w_sSql = w_sSql & "		T16_HAITOTANI, "
	w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_Z, "
	w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K "
	w_sSql = w_sSql & "	FROM "
	w_sSql = w_sSql & "		T16_RISYU_KOJIN "
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T16_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T16_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T16_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

	Set w_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(w_Rs, w_sSql)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	'************
	'* �]������ *
	'************
	m_RsHyoka.MoveFirst
	w_sHyokaMei = ""
	Do until m_RsHyoka.EOF
		if Cint(w_iSeiseki) >= Cint(m_RsHyoka("M08_MIN")) then
			w_sHyokaMei = m_RsHyoka("M08_HYOKA_SYOBUNRUI_MEI")
			w_iHyotei   = m_RsHyoka("M08_HYOTEI")
			w_iHyokaCd  = m_RsHyoka("M08_HYOKA_SYOBUNRUI_RYAKU")
		end if
		m_RsHyoka.Movenext
	loop

	'******************
	'* �����X�V�̏��� *
	'******************
	'�O���̏ꍇ
	if Cint(w_Rs("T16_KAISETU")) = Cint(mC_KAISETU_Z) then
		if not isNull(w_Rs("T16_J_JUNJIKAN_KIMATU_Z")) and Cint(w_Rs("T16_J_JUNJIKAN_KIMATU_Z")) > 0 then
			w_sAryJugyoJikan(1) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sAryJugyoJikan(2) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sAryJugyoJikan(3) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sAryJugyoJikan(4) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
		else
			w_sAryJugyoJikan(1) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(2) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(3) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(4) = Cint(w_Rs("T16_HAITOTANI")) * 30
		end if
	'����̏ꍇ
	else
		if not isNull(w_Rs("T16_J_JUNJIKAN_KIMATU_K")) and Cint(w_Rs("T16_J_JUNJIKAN_KIMATU_K")) > 0 then
			w_sAryJugyoJikan(1) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sAryJugyoJikan(2) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sAryJugyoJikan(3) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sAryJugyoJikan(4) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
		else
			w_sAryJugyoJikan(1) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(2) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(3) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sAryJugyoJikan(4) = Cint(w_Rs("T16_HAITOTANI")) * 30
		end if
	end if

	'���i�̏ꍇ�̂�T16���X�V
	if Cint(w_iHyokaCd) <> Cint(mC_HYOKA_CD_NO) then

		'***********
		'* T16�X�V *
		'***********
		w_sSql = ""
		w_sSql = w_sSql & " UPDATE "
		w_sSql = w_sSql & "		T16_RISYU_KOJIN "
		w_sSql = w_sSql & " SET "

		'�O��/����ꍇ�킯
		if Cint(w_Rs("T16_KAISETU")) = Cint(mC_KAISETU_Z) then
			w_sSql = w_sSql & "		T16_HTEN_KIMATU_Z  = " & w_iSeiseki & ", "
			w_sSql = w_sSql & "		T16_HYOKA_KIMATU_Z = '" & w_sHyokaMei & "', "
			w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_Z = " & w_iHyotei & ", "
			w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_Z = '" & f_GetNowDate() & "', "
			w_sSql = w_sSql & "		T16_HYOKA_FUKA_KBN = 0, "
			w_sSql = w_sSql & "		T16_TANI_SUMI = T16_HAITOTANI, "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_Z = " & w_sAryJugyoJikan(4) & ", "
			'����̍���
			w_sSql = w_sSql & "		T16_HTEN_KIMATU_K = " & w_iSeiseki & ", "
			w_sSql = w_sSql & "		T16_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
			w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_K = " & w_iHyotei & ", "
			w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_K = " & w_sAryJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_K = " & w_sAryJugyoJikan(4) & ", "
		else
			w_sSql = w_sSql & "		T16_HTEN_KIMATU_K = " & w_iSeiseki & ", "
			w_sSql = w_sSql & "		T16_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
			w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_K = " & w_iHyotei & ", "
			w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "
			w_sSql = w_sSql & "		T16_HYOKA_FUKA_KBN = 0, "
			w_sSql = w_sSql & "		T16_TANI_SUMI = T16_HAITOTANI, "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_K = " & w_sAryJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_K = " & w_sAryJugyoJikan(4) & ", "
		end if

		w_sSql = w_sSql & " 	T16_UPD_DATE = '" & f_GetNowDate() & "', "
		w_sSql = w_sSql & " 	T16_UPD_USER = '" & Session("LOGIN_ID") & "'"
		w_sSql = w_sSql & "	WHERE "
		w_sSql = w_sSql & "		T16_NENDO = " & p_iNendo & " AND "
		w_sSql = w_sSql & "		T16_GAKUSEI_NO = '" & p_iGakusei & "' AND "
		w_sSql = w_sSql & "		T16_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"
'Response.Write w_sSql & "<br>"
'response.end
		'���̓_����荂���ꍇ�̂�T16�X�V
		if Cint(p_iMotoTen) < Cint(w_iSeiseki) then
			w_iRet = gf_ExecuteSQL(w_sSql)
'Response.Write w_sSql & "<br>"
		end if

		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Function
		End If

	end if

	'**************************
	'* T120�̐��уf�[�^�̍X�V *
	'**************************
	if wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,w_sAryHyokaMei,w_iAryHyotei,w_bAryUpdateFlg,w_iHyokaCd) = false then
		m_bErrFlg = True
		Exit Function
	end if

	w_Rs.close

	wf_UpdateT16 = true

end function

function wf_UpdateT17(p_iSeiseki,p_iGakusei,p_iNendo,p_iMotoTen)
'********************************************************************************
'*  [�@�\]  T17�̐��уf�[�^�X�V
'*  [����]  p_iSeiseki	:�_���i�z��j
'*          p_iGakusei	:�w���ԍ�
'*          p_iNendo	:�N�x
'*          p_iMotoTen	:���̓_���iT16_HTEN�j
'*  [�ߒl]  true , false
'*  [����]  
'********************************************************************************
on error resume next

	Dim i						'�C���f�b�N�X
	Dim w_sSql					'SQL�i�[�G���A
	Dim w_Rs					'ADODB���R�[�h�Z�b�g
	Dim w_sHyokaMei				'M08_HYOKA_SYOBUNRUI_MEI
	Dim w_iHyotei				'M08_HYOTEI
	Dim w_iHyokaCd				'M08_HYOKA_SYOBUNRUI_RYAKU
	Dim w_iSeiseki				'���сi�_���i�[�G���A�j
	Dim w_sAryJugyoJikan(4)		'�����֘A

	Dim w_bAryUpdateFlg(5)		'T120�̗����X�V�t���O
	Dim w_sAryHyokaMei(5)		'T120�̗�����HYOKA
	Dim w_iAryHyotei(5)			'T120�̗�����HYOTEI

	wf_UpdateT17 = false

	'�ϐ��̏�����
	i = 5

	'*****************************************
	'* �X�V�Ώۉ񐔂̎擾(T17�ɍX�V�����l) *
	'*****************************************
	do until i = 0
		if i <> 1 then
			if p_iSeiseki(i) <> "" then
				w_iSeiseki = p_iSeiseki(i)
				exit do
			end if
		else
			if p_iSeiseki(i) <> "" then
				w_iSeiseki = p_iSeiseki(i)
				exit do
			else
				'���͂��Ȃ������̂ŃR�R�ŏI��
				wf_UpdateT17 = true
				Exit Function
			end if
		end if
		i = i - 1
	loop

	'**************************************
	'* ���͂��ꂽ�l�̔���y�сA�t���O���� *
	'**************************************
	'�z��̏�����
	for i = 1 to 5
		w_bAryUpdateFlg(i) = false		'T120�X�V�p�̃t���O
		w_sAryHyokaMei(i)  = ""			'T120_HYOKAMEI
		w_iAryHyotei(i)    = 0			'T120_HYOTEI
	next

	'�J�E���^�̐ݒ�
	i = 5

	do until i = 0

		'���͂�����
		if p_iSeiseki(i) <> "" then
			'�]��
			m_RsHyoka.MoveFirst
			Do until m_RsHyoka.EOF
				if Cint(p_iSeiseki(i)) >= Cint(m_RsHyoka("M08_MIN")) then
					w_sAryHyokaMei(i) = m_RsHyoka("M08_HYOKA_SYOBUNRUI_MEI")		'HYOKA�i�[
					w_iAryHyotei(i)   = m_RsHyoka("M08_HYOTEI")						'HYOTEI�i�[
					exit do
				end if
				m_RsHyoka.Movenext
			loop

			w_bAryUpdateFlg(i) = true
		end if
		i = i - 1
	loop

	'******************
	'* ��{�f�[�^�擾 *
	'******************
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T17_KAISETU "
	w_sSql = w_sSql & "		T17_HAITOTANI, "
	w_sSql = w_sSql & "		T17_J_JUNJIKAN_KIMATU_K "
	w_sSql = w_sSql & "	FROM "
	w_sSql = w_sSql & "		T17_RISYUKAKO_KOJIN "
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T17_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T17_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T17_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

	Set w_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(w_Rs, w_sSql)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	'************
	'* �]������ *
	'************
	m_RsHyoka.MoveFirst
	w_sHyokaMei = ""
	Do until m_RsHyoka.EOF
		if Cint(w_iSeiseki) >= Cint(m_RsHyoka("M08_MIN")) then
			w_sHyokaMei = m_RsHyoka("M08_HYOKA_SYOBUNRUI_MEI")
			w_iHyotei   = m_RsHyoka("M08_HYOTEI")
			w_iHyokaCd  = m_RsHyoka("M08_HYOKA_SYOBUNRUI_RYAKU")
			exit do
		end if
		m_RsHyoka.Movenext
	loop

	'******************
	'* �����X�V�̏��� *
	'******************
	'����̂�
	if not isNull(w_Rs("T17_J_JUNJIKAN_KIMATU_K")) and Cint(w_Rs("T17_J_JUNJIKAN_KIMATU_K")) > 0 then
		w_sAryJugyoJikan(1) = w_Rs("T17_J_JUNJIKAN_KIMATU_K")
		w_sAryJugyoJikan(2) = w_Rs("T17_J_JUNJIKAN_KIMATU_K")
		w_sAryJugyoJikan(3) = w_Rs("T17_J_JUNJIKAN_KIMATU_K")
		w_sAryJugyoJikan(4) = w_Rs("T17_J_JUNJIKAN_KIMATU_K")
	else
		w_sAryJugyoJikan(1) = Cint(w_Rs("T17_HAITOTANI")) * 30
		w_sAryJugyoJikan(2) = Cint(w_Rs("T17_HAITOTANI")) * 30
		w_sAryJugyoJikan(3) = Cint(w_Rs("T17_HAITOTANI")) * 30
		w_sAryJugyoJikan(4) = Cint(w_Rs("T17_HAITOTANI")) * 30
	end if

	'���i�̏ꍇ
	if Cint(w_iHyokaCd) <> Cint(mC_HYOKA_CD_NO) then

		'***********
		'* T17�X�V *
		'***********
		w_sSql = ""
		w_sSql = w_sSql & " UPDATE "
		w_sSql = w_sSql & "		T17_RISYUKAKO_KOJIN "
		w_sSql = w_sSql & " SET "
		'2003.09.02 �O���ƌ���̐��т𓯎��ɍX�V����悤�ɏC�� ITO
		w_sSql = w_sSql & "		T17_HTEN_KIMATU_Z  = " & w_iSeiseki & ", "
		w_sSql = w_sSql & "		T17_HYOKA_KIMATU_Z = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T17_HYOTEI_KIMATU_Z = " & w_iHyotei & ", "
		w_sSql = w_sSql & "		T17_KOUSINBI_KIMATU_Z = '" & f_GetNowDate() & "', "

		w_sSql = w_sSql & "		T17_HTEN_KIMATU_K  = " & w_iSeiseki & ", "
		w_sSql = w_sSql & "		T17_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T17_HYOTEI_KIMATU_K = " & w_iHyotei & ", "
		w_sSql = w_sSql & "		T17_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "

		w_sSql = w_sSql & "		T17_HYOKA_FUKA_KBN = 0, "
		w_sSql = w_sSql & "		T17_TANI_SUMI = T17_HAITOTANI, "

		'2003.09.02 �O���ƌ���̐��т𓯎��ɍX�V����悤�ɏC�� ITO
		w_sSql = w_sSql & "		T17_SOJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(1) & ", "
		w_sSql = w_sSql & "		T17_JUNJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(2) & ", "
		w_sSql = w_sSql & "		T17_J_JUNJIKAN_KIMATU_Z = " & w_sAryJugyoJikan(3) & ", "
		w_sSql = w_sSql & "		T17_KEKA_KIMATU_Z = " & w_sAryJugyoJikan(4) & ", "

		w_sSql = w_sSql & "		T17_SOJIKAN_KIMATU_K = " & w_sAryJugyoJikan(1) & ", "
		w_sSql = w_sSql & "		T17_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(2) & ", "
		w_sSql = w_sSql & "		T17_J_JUNJIKAN_KIMATU_K = " & w_sAryJugyoJikan(3) & ", "
		w_sSql = w_sSql & "		T17_KEKA_KIMATU_Z = " & w_sAryJugyoJikan(4) & ", "

		w_sSql = w_sSql & " 	T17_UPD_DATE = '" & f_GetNowDate() & "', "
		w_sSql = w_sSql & " 	T17_UPD_USER = '" & Session("LOGIN_ID") & "'"
		w_sSql = w_sSql & "	WHERE "
		w_sSql = w_sSql & "		T17_NENDO = " & p_iNendo & " AND "
		w_sSql = w_sSql & "		T17_GAKUSEI_NO = '" & p_iGakusei & "' AND "
		w_sSql = w_sSql & "		T17_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

		'���Ƃ̓_����荂���ꍇ�̂ݍX�V
		if Cint(p_iMotoTen) < Cint(w_iSeiseki) then
			w_iRet = gf_ExecuteSQL(w_sSql)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				Exit Function
			End If
		end if

	end if

	'**************************
	'* T120�̐��уf�[�^�̍X�V *
	'**************************
	if wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,w_sAryHyokaMei,w_iAryHyotei,w_bAryUpdateFlg,w_iHyokaCd) = false then
		m_bErrFlg = True
		Exit Function
	end if

	w_Rs.close

	wf_UpdateT17 = true

end function

function wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,p_sAryHyokaMei,p_iAryHyotei,p_bAryUpdateFlg,p_iHyokaCd)
'********************************************************************************
'*  [�@�\]  T120_SAISIKEN�̐��уf�[�^�X�V
'*  [����]  p_iSeiseki		:�_���i�z��j
'*          p_iGakusei		:�w���ԍ�
'*          p_iNendo		:�N�x
'*          p_sAryHyokaMei	:�]�����́i�z��j
'*          p_iAryHyotei	:�]��i�z��j
'*          p_bAryUpdateFlg	:�X�V�t���O�i�z��j	< True �ōX�V
'*          p_iHyokaCd		:�]���R�[�h			< ���۔���Ɏg�p
'*  [�ߒl]  true , false
'*  [����]  
'********************************************************************************

	Dim i			'�C���f�b�N�X
	Dim w_sSql		'SQL���i�[�G���A
	Dim w_iRet		'�߂�l�i�[�G���A

	wf_UpdateT120 = false

	'**************
	'* T120�̍X�V *
	'**************
	w_sSql = ""
	w_sSql = w_sSql & " UPDATE "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "	SET "
	'�X�V�t���O
	for i = 1 to 5

'		if p_bAryUpdateFlg(i) = true then
			if i = 1 then
				w_sSql = w_sSql & " 	T120_SAISI_SEISEKI " & f_IntUpdSQL(p_iSeiseki(i)) & ","
				w_sSql = w_sSql & " 	T120_HYOKA         " & f_StrUpdSQL(p_sAryHyokaMei(i)) & ","
				w_sSql = w_sSql & " 	T120_HYOTEI        " & f_NumUpdSQL(p_iAryHyotei(i)) & ","
			else
				w_sSql = w_sSql & " 	T120_SAISI_SEISEKI" & i & f_IntUpdSQL(p_iSeiseki(i)) & ","
				w_sSql = w_sSql & " 	T120_HYOKA"  & i & f_StrUpdSQL(p_sAryHyokaMei(i))  & ","
				w_sSql = w_sSql & " 	T120_HYOTEI"  & i & f_NumUpdSQL(p_iAryHyotei(i)) & ","
			end if
'		end if

	next

	'���i�̏ꍇ
	if Cint(p_iHyokaCd) = Cint(mC_HYOKA_CD_OK) then
		w_sSql = w_sSql & " 	T120_SYUTOKU_NENDO = " & Session("NENDO") & ", "
		w_sSql = w_sSql & " 	T120_SYUTOKU_FLG = 1, "
		w_sSql = w_sSql & " 	T120_HYOKA_FUKA_KBN = 0, "
	else
		w_sSql = w_sSql & " 	T120_SYUTOKU_NENDO = NULL, "
		w_sSql = w_sSql & " 	T120_SYUTOKU_FLG = 0, "
		w_sSql = w_sSql & " 	T120_HYOKA_FUKA_KBN = 1, "
	end if
	w_sSql = w_sSql & " 	T120_UPD_DATE = '" & f_GetNowDate() & "', "
	w_sSql = w_sSql & " 	T120_UPD_USER = '" & Session("LOGIN_ID") & "'"
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T120_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T120_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T120_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

	w_iRet = gf_ExecuteSQL(w_sSql)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	wf_UpdateT120 = true

end function


function wf_GetHyoka()
'********************************************************************************
'*  [�@�\]  ���т̔���A�o�^
'*  [����]  �Ȃ�
'*  [�ߒl]  true ,false
'*  [����]  
'********************************************************************************

	'�ϐ��̐錾
	Dim w_sSql
	Dim w_iRet

	wf_GetHyoka = false

	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		M08_HYOKA_SYOBUNRUI_MEI, "
	w_sSql = w_sSql & "		M08_HYOTEI, "
	w_sSql = w_sSql & "		M08_HYOKA_SYOBUNRUI_RYAKU, "
	w_sSql = w_sSql & "		M08_MIN "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		M08_HYOKAKEISIKI "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & "		M08_NENDO = " & Session("NENDO") & " AND "
	w_sSql = w_sSql & "		M08_HYOKA_TAISYO_KBN = " & mC_HYOKA_KBN & " AND "
	w_sSql = w_sSql & "		M08_HYOUKA_NO = " & mC_HYOKA_NO
	w_sSql = w_sSql & " ORDER BY "
	w_sSql = w_sSql & "		M08_HYOKA_SYOBUNRUI_CD ASC "

	Set m_RsHyoka = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(m_RsHyoka, w_sSQL)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	wf_GetHyoka = true

end function

function wf_GetStudent()
'********************************************************************************
'*  [�@�\]  ���C���w���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  true ,false
'*  [����]  
'********************************************************************************

	'�ϐ��̐錾
	Dim w_sSql
	Dim w_iRet

	wf_GetStudent = false
	
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	
	'��ʂɕ\�����鍀��
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_MISYU_GAKUNEN,  "
	w_sSql = w_sSql & "		M05_CLASS.M05_CLASSMEI, "
	w_sSql = w_sSql & "		T11_GAKUSEKI.T11_SIMEI, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_NENDO, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_JYUKOKAISU, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SEISEKI, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI2, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI3, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI4, "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SAISI_SEISEKI5, "

	'Hidden����
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_GAKUSEI_NO "
	
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN, "
	w_sSql = w_sSql & "		T11_GAKUSEKI, "
	w_sSql = w_sSql & "		T13_GAKU_NEN, "
	w_sSql = w_sSql & "		M05_CLASS "
'	w_sSql = w_sSql & "		M08_HYOKAKEISIKI "
	w_sSql = w_sSql & " WHERE "
	
	'TABLE�̌�������
	w_sSql = w_sSql & "			T120_SAISIKEN.T120_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO "
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_NENDO = T13_GAKU_NEN.T13_NENDO "
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_GAKUSEI_NO = T13_GAKU_NEN.T13_GAKUSEI_NO "
	w_sSql = w_sSql & "		AND T13_GAKU_NEN.T13_NENDO = M05_CLASS.M05_NENDO "
	w_sSql = w_sSql & "		AND T13_GAKU_NEN.T13_GAKUNEN = M05_CLASS.M05_GAKUNEN "
	w_sSql = w_sSql & "		AND T13_GAKU_NEN.T13_CLASS = M05_CLASS.M05_CLASSNO "
'	w_sSql = w_sSql & "		AND M08_NENDO = T120_NENDO"			'���C�N�x�̕]��
'	w_sSql = w_sSql & " 	AND M08_HYOUKA_NO = 2 "
'	w_sSql = w_sSql & "		AND M08_HYOKA_TAISYO_KBN = 0 "
'	w_sSql = w_sSql & "		AND M08_HYOKA_SYOBUNRUI_CD = 4 "
	'���̑�����
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "' "
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_KYOUKAN_CD = '" & Session("KYOKAN_CD") & "' "
	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_MISYU_GAKUNEN = " & Request("hidMISYU_GAKUNEN") 
'�����Ή��p�i��ŊO��
	w_sSql = w_sSql & " 	AND ( T120_SYUTOKU_NENDO Is Null or T120_SYUTOKU_NENDO = " & Session("NENDO") & " ) "
'	w_sSql = w_sSql & "		AND T120_SAISIKEN.T120_HYOKA_FUKA_KBN <> 2 "
	w_sSql = w_sSql & " 	AND NOT T120_SEISEKI Is Null "
	w_sSql = w_sSql & " 	AND T120_TAISYO_FLG = 1 "
'	w_sSql = w_sSql & " 	AND T120_SEISEKI <= M08_MAX "
'	w_sSql = w_sSql & " 	AND T120_SEISEKI >= M08_MIN "

	w_sSql = w_sSql & "	ORDER BY"
	w_sSql = w_sSql & "		T13_GAKUNEN,"	
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "

'response.write w_sSql & "<br>"

	Set m_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

'Response.write gf_GetRsCount(m_Rs) & "<br>"

	wf_GetStudent = true
	
end function

function f_GetNowDate()
'--------------------------------------------------------------------------------
'-  [�@�\]  ���݂̓��t���擾
'-  [����]  �Ȃ�
'-  [�ߒl]  YYYY/MM/DD
'-  [����]  
'--------------------------------------------------------------------------------
	Dim wResult

	f_GetNowDate = ""

	wResult = gf_fmtZero(Year(Date()),4) & "/" & gf_fmtZero(Month(Date()),2) & "/" & gf_fmtZero(Day(Date()),2)

	f_GetNowDate = wResult

end function

'/*******************************************************************************
'/*          UPDATE SQL ���l�^�f�[�^�̍쐬
'/*                  pv_string�@�@update SQL �f�[�^
'/*
'/*               �߂�l      =�f�[�^
'/********************************************************************************
Function f_IntUpdSQL(pString)

'***�ϐ��錾***
    Dim ls_String
    Dim ls_Length

'***�ϊ�����***
    '�����񒆂�'����������A''�ɕϊ�����
    pString = Replace(pString, "'", "''")
  
    If gf_IsNull(pString) Then              'NULL�̎�
        f_IntUpdSQL = " = NULL "
    Else
        ls_String = pString
        f_IntUpdSQL = " = " & pString
    End If

End Function

'/*******************************************************************************
'/*          UPDATE SQL ������f�[�^�̍쐬
'/*                  pv_string�@�@update SQL �f�[�^
'/*
'/*               �߂�l      ='������f�[�^'
'/********************************************************************************
Function f_StrUpdSQL(pString)

'***�ϐ��錾***
    Dim ls_String
    Dim ls_Length

'***�ϊ�����***
    '�����񒆂�'����������A''�ɕϊ�����
    pString = Replace(pString, "'", "''")
    
    If gf_IsNull(pString) Then              'NULL�̎�
        f_StrUpdSQL = " = NULL "
    Else
        ls_String = pString
        f_StrUpdSQL = " = '" & pString & "'"
    End If

End Function

'/*******************************************************************************
'/*          UPDATE SQL ������f�[�^�̍쐬
'/*                  pv_string�@�@update SQL �f�[�^
'/*
'/*               �߂�l      ='������f�[�^'
'/********************************************************************************
Function f_NumUpdSQL(pString)

'***�ϐ��錾***
    Dim ls_String
    Dim ls_Length

'***�ϊ�����***
    '�����񒆂�'����������A''�ɕϊ�����
    pString = Replace(pString, "'", "''")
    
    If gf_IsNull(pString) Then              'NULL�̎�
        f_NumUpdSQL = " = NULL "
    ElseIf pString = 0 Then
		f_NumUpdSQL = " = NULL "
	Else
        ls_String = pString
        f_NumUpdSQL = " = " & pString
    End If

End Function

sub showPage()
'********************************************************************************
'*  [�@�\]  HTML�̕\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	'�ϐ��̐錾
	Dim w_iCount
	Dim w_sCellClass
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../common/style.css" type="text/css">
<title>�V�����y�[�W 1</title>

<script language="JavaScript">
<!--

//================================================
//	���M����
//================================================
function jf_Submit() {

	if (!jf_CheckValue()) {
		return;
	}

	if (!confirm("�Ď����̐��т�o�^���܂��B��낵���ł����H")) {
		return;
	}

	document.frm.hidMode.value = "update";
	document.frm.action = "./saisi0100_toroku.asp";
	document.frm.target = "fMain";
	document.frm.submit();

}

//================================================
//	�߂鏈��
//================================================
function jf_Back() {

	location.href = "saisi0100_show.asp";
	return;

}
//================================================
//	�l�̃`�F�b�N
//================================================
function jf_CheckValue() {

	var i,j				//�C���f�b�N�X
	var w_iValueCnt;	//���k���i�[�G���A
	var w_oText;		//TEXTBOX�I�u�W�F�N�g�i�[�G���A
	
	w_iValueCnt = Number(document.frm.hidCount.value);

	for (i=1;i<=w_iValueCnt;i++) {
		
		for (j=1;j<=5;j++) {

			//�I�u�W�F�N�g�̎擾
			w_oText = eval("document.frm.txtSeiseki" + j + "_" + i);

			if (w_oText.value != "") {
				//���l�^�`�F�b�N
				if (isNaN(w_oText.value)) {
					alert("���т͐����œ��͂��Ă��������B");
					w_oText.focus();
					return false;
				}
				//100�ȉ��`�F�b�N
				if (Number(w_oText.value) > 100) {
					alert("���т�100�𒴂��Ă��܂��B");
					w_oText.focus();
					return false;
				}
			}
		}
	}
	
	return true;
}

//-->
</script>

</head>

<body>

<form name="frm" method="post">

<center>
<br>

<table border="1" class="hyo">
	<tr>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="16"><font color="#FFFFFF">���C�w�N</font></td>
		<td width="70"  class="CELL2"   height="16" align="center"><%=Request("hidMISYU_GAKUNEN")%></td>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="16"><font color="#FFFFFF">�ȁ@�@��</font></td>
		<td width="200" class="CELL2"   height="16" align="center"><%=Request("hidKAMOKU_MEI")%></td>
	</tr>
</table>

<br>
<br>

<table border="1" class="hyo">

	<!-- TABLE�w�b�_�� -->
	<tr>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">�w�N</font></td>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">�N���X</font></td>
		<td width="200" class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">���@�@�@��</font></td>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">�N�x</font></td>
		<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">�󌱉�</font></td>
		<td width="40"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">����</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">1��</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">2��</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">3��</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">4��</font></td>
		<td width="30"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">5��</font></td>
	</tr>


	<!-- TABLE���X�g�� -->
<%
	'�J�E���^�̏�����
	w_iCount = 0
	
	'TD��Class�̏�����
	w_sCellClass = "CELL2"

	do until m_Rs.EOF
		w_iCount = w_iCount + 1
%>
	<tr>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_MISYU_GAKUNEN")%><br></td>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("M05_CLASSMEI")%><br></td>
		<td width="200" class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T11_SIMEI")%><br></td>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_NENDO")%><br></td>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_JYUKOKAISU")%><br></td>
		<td width="40"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="hidden" name="hidGakusei<%=w_iCount%>"  value="<%=m_Rs("T120_GAKUSEI_NO")%>">
			<input type="hidden" name="hidNendo<%=w_iCount%>"    value="<%=m_Rs("T120_NENDO")%>">
			<input type="hidden" name="hidSeiseki<%=w_iCount%>"    value="<%=m_Rs("T120_SEISEKI")%>">
			<%=m_Rs("T120_SEISEKI")%><br>
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki1_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI")%>"  size="3" style="ime-mode:disabled" maxlength="3">
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki2_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI2")%>" size="3" style="ime-mode:disabled" maxlength="3">
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki3_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI3")%>" size="3" style="ime-mode:disabled" maxlength="3">
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki4_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI4")%>" size="3" style="ime-mode:disabled" maxlength="3">
		</td>
		<td width="30"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="text"   name="txtSeiseki5_<%=w_iCount%>" value="<%=m_Rs("T120_SAISI_SEISEKI5")%>" size="3" style="ime-mode:disabled" maxlength="3">
		</td>
	</tr>

<%
		m_Rs.MoveNext
		
		if w_sCellClass = "CELL2" then
			w_sCellClass = "CELL1"
		else
			w_sCellClass = "CELL2"
		end if
		
	loop	
%>
</table>
<br>

<table>
	<tr>
		<td><input type="button" value=" �o�@�^ " onclick="jf_Submit();"></td>
		<td><input type="button" value=" �߁@�� " onclick="jf_Back();"></td>
	</tr>
</table>

</center>

<!-- ���� -->
<input type="hidden" name="hidCount" value="<%=w_iCount%>">
<input type="hidden" name="hidKAMOKU_CD" value="<%=Request("hidKAMOKU_CD")%>">
<input type="hidden" name="hidMode">

</form>

</body>

</html>
<%
end sub
%>