<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �ǎ����ѓo�^
' ��۸���ID : saisi/saisi0200/saisi0200_toroku.asp
' �@      �\: �ǎ����̓o�^���
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
' ��      �X: 2003/03/07 ���@��u�񐔂�\�����Ȃ��l�ɕύX
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
			response.redirect "saisi0200_show.asp"
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
	Dim w_sSql
	Dim w_iRet
	Dim i
	Dim w_bNendo	'True:���N�x�@false:�ߋ�
	Dim w_iSeiseki
	Dim w_iGakusei
	Dim w_iNendo

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
		'�z��̊i�[
		w_iSeiseki = Request("txtSeiseki" & i)

'�f�o�b�O�v�����g�p		
'Response.Write "�ԍ���" & Request("hidGakusei" & i) & "<br>"
'Response.Write "�N�x��" & Request("hidNendo" & i) & "<br>"
'Response.Write "����1��" & Request("txtSeiseki" & i) & "<br>"

		'T16 & T120 �̍X�V
		if wf_UpdateT16(w_iSeiseki,w_iGakusei,w_iNendo) = false then
			m_bErrFlg = True
			Exit Function
		end if

'Response.Write "-------------------------------------------------------------------------------------------------------------------<br>"
	next	
'response.end
	wf_UpdateSeiseki = true

end function

function wf_UpdateT16(p_iSeiseki,p_iGakusei,p_iNendo)
'********************************************************************************
'*  [�@�\]  T16�̐��уf�[�^�X�V
'*  [����]  p_iSeiseki :���сi���͒l�j
'*          p_iGakusei :�w���ԍ�
'*          p_iNendo   :�N�x
'*  [�ߒl]  true , false
'*  [����]  
'********************************************************************************

	Dim i						'�C���f�b�N�X
	Dim w_sSql					'SQL�i�[�G���A
	Dim w_Rs					'ADODB���R�[�h�Z�b�g
	Dim w_sHyokaMei				'M08_HYOKA_SYOBUNRUI_MEI
	Dim w_iHyotei				'M08_HYOTEI
	Dim w_iHyokaCd				'M08_HYOKA_SYOBUNRUI_RYAKU
	Dim w_sJugyoJikan(5)		'�����֘A

	wf_UpdateT16 = false

	'**********************************
	'* �_���������ĂȂ��ꍇ�������Ȃ� *
	'**********************************
	if p_iSeiseki = "" then
		wf_UpdateT16 = true
		exit function
	end if

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
		if Cint(p_iSeiseki) >= Cint(m_RsHyoka("M08_MIN")) then
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
			w_sJugyoJikan(1) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sJugyoJikan(2) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sJugyoJikan(3) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
			w_sJugyoJikan(4) = w_Rs("T16_J_JUNJIKAN_KIMATU_Z")
		else
			w_sJugyoJikan(1) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(2) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(3) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(4) = Cint(w_Rs("T16_HAITOTANI")) * 30
		end if
	'����̏ꍇ
	else
		if not isNull(w_Rs("T16_J_JUNJIKAN_KIMATU_K")) and Cint(w_Rs("T16_J_JUNJIKAN_KIMATU_K")) > 0 then
			w_sJugyoJikan(1) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sJugyoJikan(2) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sJugyoJikan(3) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
			w_sJugyoJikan(4) = w_Rs("T16_J_JUNJIKAN_KIMATU_K")
		else
			w_sJugyoJikan(1) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(2) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(3) = Cint(w_Rs("T16_HAITOTANI")) * 30
			w_sJugyoJikan(4) = Cint(w_Rs("T16_HAITOTANI")) * 30
		end if
	end if

	'***********
	'* T16�X�V *
	'***********
	w_sSql = ""
	w_sSql = w_sSql & " UPDATE "
	w_sSql = w_sSql & "		T16_RISYU_KOJIN "
	w_sSql = w_sSql & " SET "

	'�O��/����ꍇ�킯
	if Cint(w_Rs("T16_KAISETU")) = Cint(mC_KAISETU_Z) then
		w_sSql = w_sSql & "		T16_HTEN_KIMATU_Z  = " & p_iSeiseki & ", "
		w_sSql = w_sSql & "		T16_HYOKA_KIMATU_Z = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_Z = '" & w_sHyotei & "', "
		w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_Z = '" & f_GetNowDate() & "', "
		w_sSql = w_sSql & "		T16_HTEN_KIMATU_K  = " & p_iSeiseki & ", "
		w_sSql = w_sSql & "		T16_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "
		'���i�̏ꍇ
		if Cint(w_iHyokaCd) = Cint(mC_HYOKA_CD_OK) then
			w_sSql = w_sSql & "		T16_HYOKA_FUKA_KBN = 0, "
			w_sSql = w_sSql & "		T16_TANI_SUMI = T16_HAITOTANI, "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_Z = " & w_sJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_Z = " & w_sJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_Z = " & w_sJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_Z = " & w_sJugyoJikan(4) & ", "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_K = " & w_sJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_K = " & w_sJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K = " & w_sJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_K = " & w_sJugyoJikan(4) & ", "
		end if
	else
		w_sSql = w_sSql & "		T16_HTEN_KIMATU_K = " & p_iSeiseki & ", "
		w_sSql = w_sSql & "		T16_HYOKA_KIMATU_K = '" & w_sHyokaMei & "', "
		w_sSql = w_sSql & "		T16_HYOTEI_KIMATU_K = '" & w_sHyotei & "', "
		w_sSql = w_sSql & "		T16_KOUSINBI_KIMATU_K = '" & f_GetNowDate() & "', "
		'���i�̏ꍇ
		if Cint(w_iHyokaCd) = Cint(mC_HYOKA_CD_OK) then
			w_sSql = w_sSql & "		T16_HYOKA_FUKA_KBN = 0, "
			w_sSql = w_sSql & "		T16_TANI_SUMI = T16_HAITOTANI, "
			w_sSql = w_sSql & "		T16_SOJIKAN_KIMATU_K = " & w_sJugyoJikan(1) & ", "
			w_sSql = w_sSql & "		T16_JUNJIKAN_KIMATU_K = " & w_sJugyoJikan(2) & ", "
			w_sSql = w_sSql & "		T16_J_JUNJIKAN_KIMATU_K = " & w_sJugyoJikan(3) & ", "
			w_sSql = w_sSql & "		T16_KEKA_KIMATU_K = " & w_sJugyoJikan(4) & ", "
		end if
	end if
	w_sSql = w_sSql & " 	T16_UPD_DATE = '" & f_GetNowDate() & "', "
	w_sSql = w_sSql & " 	T16_UPD_USER = '" & Session("LOGIN_ID") & "'"
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T16_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T16_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T16_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

'response.write w_sSql & "<br>"

	w_iRet = gf_ExecuteSQL(w_sSql)
'Response.Write w_sSql & "<br>"


	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	'**************************
	'* T120�̐��уf�[�^�̍X�V *
	'**************************
	if wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,w_sHyokaMei,w_iHyotei,w_iHyokaCd) = false then
		m_bErrFlg = True
		Exit Function
	end if

	w_Rs.close

	wf_UpdateT16 = true

end function

function wf_UpdateT120(p_iSeiseki,p_iGakusei,p_iNendo,p_sHyokaMei,p_iHyotei,p_iHyokaCd)
'********************************************************************************
'*  [�@�\]  T120_SAISIKEN�̐��уf�[�^�X�V
'*  [����]  �Ȃ�
'*  [�ߒl]  true , false
'*  [����]  
'********************************************************************************

	Dim w_sSql
	Dim w_iRet

	wf_UpdateT120 = false

	'**************
	'* T120�̍X�V *
	'**************
	w_sSql = ""
	w_sSql = w_sSql & " UPDATE "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "	SET "
	w_sSql = w_sSql & "		T120_HYOKA = '" & p_sHyokaMei & "', "
	w_sSql = w_sSql & "		T120_HYOTEI = " & p_iHyotei & ", "
	w_sSql = w_sSql & "		T120_SEISEKI = " & p_iSeiseki & ", "
	'���i�̏ꍇ
	if Cint(p_iHyokaCd) = Cint(mC_HYOKA_CD_OK) then
		w_sSql = w_sSql & " 	T120_SYUTOKU_NENDO = " & Session("NENDO") & ", "
		w_sSql = w_sSql & " 	T120_SYUTOKU_FLG = 1, "
		w_sSql = w_sSql & " 	T120_HYOKA_FUKA_KBN = 0, "
	else
		w_sSql = w_sSql & " 	T120_HYOKA_FUKA_KBN = 1, "
	end if
	w_sSql = w_sSql & " 	T120_UPD_DATE = '" & f_GetNowDate() & "', "
	w_sSql = w_sSql & " 	T120_UPD_USER = '" & Session("LOGIN_ID") & "'"
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T120_NENDO = " & p_iNendo & " AND "
	w_sSql = w_sSql & "		T120_GAKUSEI_NO = '" & p_iGakusei & "' AND "
	w_sSql = w_sSql & "		T120_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "'"

	w_iRet = gf_ExecuteSQL(w_sSql)
'Response.Write w_sSql & "<br>"

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
	w_sSql = w_sSql & "		M08_HYOKA_SYOBUNRUI_CD DESC "
	
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

	'Hidden����
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_GAKUSEI_NO "
	
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN, "
	w_sSql = w_sSql & "		T11_GAKUSEKI, "
	w_sSql = w_sSql & "		T13_GAKU_NEN, "
	w_sSql = w_sSql & "		M05_CLASS "
	w_sSql = w_sSql & " WHERE "
	
	'TABLE�̌�������
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_NENDO = T13_GAKU_NEN.T13_NENDO AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_GAKUSEI_NO = T13_GAKU_NEN.T13_GAKUSEI_NO AND "
	w_sSql = w_sSql & "		T13_GAKU_NEN.T13_NENDO = M05_CLASS.M05_NENDO AND "
	w_sSql = w_sSql & "		T13_GAKU_NEN.T13_GAKUNEN = M05_CLASS.M05_GAKUNEN AND "
	w_sSql = w_sSql & "		T13_GAKU_NEN.T13_CLASS = M05_CLASS.M05_CLASSNO AND "
	
	'���̑�����
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_NENDO = " & Session("NENDO") & " AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_KAMOKU_CD = '" & Request("hidKAMOKU_CD") & "' AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_KYOUKAN_CD = '" & Session("KYOKAN_CD") & "' AND "
'�����Ή��p�i��ŊO��
'	w_sSql = w_sSql & "		T120_SAISIKEN.T120_HYOKA_FUKA_KBN <> 2 AND "
	w_sSql = w_sSql & "		T120_SAISIKEN.T120_SEISEKI Is Null "

	w_sSql = w_sSql & "	ORDER BY"
	w_sSql = w_sSql & "		T13_GAKUNEN,"	
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "


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
'-----------------------------------------------------------------
'	���݂̓��t���擾	�߂�l�FYYYY/MM/DD
'-----------------------------------------------------------------
	Dim wResult

	f_GetNowDate = ""

	wResult = gf_fmtZero(Year(Date()),4) & "/" & gf_fmtZero(Month(Date()),2) & "/" & gf_fmtZero(Day(Date()),2)

	f_GetNowDate = wResult

end function

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

	if (!confirm("�ǎ����̐��т�o�^���܂��B��낵���ł����H")) {
		return;
	}

	document.frm.hidMode.value = "update";
	document.frm.action = "./saisi0200_toroku.asp";
	document.frm.target = "fMain";
	document.frm.submit();

}

//================================================
//	�߂鏈��
//================================================
function jf_Back() {

	location.href = "saisi0200_show.asp";
	return;

}
//================================================
//	�l�̃`�F�b�N
//================================================
function jf_CheckValue() {

	var i				//�C���f�b�N�X
	var w_iValueCnt;	//���k���i�[�G���A
	var w_oText;		//TEXTBOX�I�u�W�F�N�g�i�[�G���A
	
	w_iValueCnt = Number(document.frm.hidCount.value);

	for (i=1;i<=w_iValueCnt;i++) {
		
		//�I�u�W�F�N�g�̎擾
		w_oText = eval("document.frm.txtSeiseki" + i);

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
		<!--<td width="70"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">�󌱉�</font></td>-->
		<td width="40"  class="header3" align="center" bgcolor="#666699" height="24"><font color="#FFFFFF">����</font></td>
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
		<!--<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_JYUKOKAISU")%><br></td>-->
		<td width="40"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="hidden" name="hidGakusei<%=w_iCount%>"   value="<%=m_Rs("T120_GAKUSEI_NO")%>">
			<input type="hidden" name="hidNendo<%=w_iCount%>"     value="<%=m_Rs("T120_NENDO")%>">
			<input type="text"   name="txtSeiseki<%=w_iCount%>" size="3" style="ime-mode:disabled" maxlength="3">
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