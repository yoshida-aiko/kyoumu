<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �g�p���ȏ��o�^
' ��۸���ID : web/WEB0321/WEB0321_main.asp
' �@	  �\: �g�p���ȏ��̓o�^���s��
'-------------------------------------------------------------------------
' ��	  ��:�����R�[�h 	��		SESSION���i�ۗ��j
' ��	  ��:�Ȃ�
' ��	  �n:�����R�[�h 	��		SESSION���i�ۗ��j
' ��	  ��:
'			���t���[���y�[�W
'-------------------------------------------------------------------------
' ��	  ��: 2001/08/01 �O�c �q�j
' ��	  �X: 2001/08/22 �ɓ� ���q ������I���ł���悤�ɕύX
' ��	  �X: 2001/12/01 �c�� ��K �����w�Ȃ݂̂�ύX����悤�ɏC��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public	m_bErrFlg			'�װ�׸�
	Public	m_iNendo			'�N�x
	Public	m_sKyokan_CD		'����CD
	Public	m_sPageCD			':�\���ϕ\���Ő��i�������g����󂯎������j
	Public	m_iMax
	Public	m_Rs
	Public	w_sSQL
	Public	m_iDsp
	Public	m_iDisp 		':�\�������̍ő�l���Ƃ�
	Public	m_sGakka		 '�w�Ȗ���

	Public m_iGakunen
	Public m_sGakkaCd

	Public	m_sKyokanNm		'//���O�C��������

	Public m_sSyozokuGakka		'//2001/12/01 Add ���O�C�����������̏�������w��
	Public m_sKamokuCD()		'//2001/12/01 Add �S������Ȗڂ̈ꗗ

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
	w_sMsgTitle="�g�p���ȏ��o�^"
	w_sMsg=""
	w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""


	On Error Resume Next
	Err.Clear

	m_bErrFlg = False
	m_iDsp = C_PAGE_LINE

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
		session("PRJ_No") = "WEB0321"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'// �l��ϐ��ɓ����
		Call s_SetParam()

		'// �\���p�ް����擾����
		if f_GetData() = False then
			exit do
		end if

		'// �S���Ȗ��ް����擾����
		if f_GetTantoKamoku() = False then
			exit do
		end if

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
	Call gf_closeObject(m_Rs)
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[�@�\]	�l��ϐ��ɓ����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_SetParam()

	m_iNendo	 = request("txtNendo")					   '�N�x
	m_iGakunen = trim(replace(Request("txtGakunenCd"),"@@@",""))
	m_sGakkaCd = trim(replace(Request("txtGakkaCD"),"@@@",""))
	m_iDisp = C_PAGE_LINE		'�P�y�[�W�ő�\����

	'// BLANK�̏ꍇ�͍s���ر
	If Request("txtMode") = "" Then
		m_sPageCD = 1
	Else
		m_sPageCD = INT(Request("txtPageCD"))	':�\���ϕ\���Ő��i�������g����󂯎������j
	End If
	If m_sPageCD = 0 Then m_sPageCD = 1

End Sub

'********************************************************************************
'*	[�@�\]	�\���ް����擾����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
function f_GetData()
	Dim w_sSQL				'// SQL��
	Dim w_iRet				'// �߂�l

	Dim w_oRecord			'//2001/12/01 Add �����w�Ȏ擾�̂���

	f_GetData = False

	'2001/12/01 Add ---->
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
		Exit function
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


'	 w_sSQL = w_sSQL & vbCrLf & " SELECT "
'	 w_sSQL = w_sSQL & vbCrLf & " T47.T47_NENDO "			 ''�N�x
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKI_KBN "		 ''�w���敪
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_NO"				 ''No
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKUNEN "		 ''�w�N
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKA_CD "		 ''�w��
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KAMOKU " 		 ''�Ȗں���
'''  w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKAN "
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKASYO "		 ''���ȏ���
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SYUPPANSYA " 	 ''�o�Ŏ�
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_TYOSYA " 		 ''����
'''  w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKUSEISU "
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKANYOUSU "	 ''�����p��
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SIDOSYOSU "		 ''�w������
'	 w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_BIKOU "			 ''���l
'''  w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_NENDO "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKA_CD "
'	 w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKAMEI "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_NENDO "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKU_CD "
'	 w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_NENDO "
'''  w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKAN_CD "
'	 w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
'	 w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
'	 w_sSQL = w_sSQL & vbCrLf & " FROM "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47_KYOKASYO T47 "
'	 w_sSQL = w_sSQL & vbCrLf & "    ,M02_GAKKA M02 "
'	 w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
'	 w_sSQL = w_sSQL & vbCrLf & "    ,M04_KYOKAN M04 "
'	 w_sSQL = w_sSQL & vbCrLf & " WHERE "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M02.M02_NENDO(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_GAKKA_CD  = M02.M02_GAKKA_CD(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M03.M03_NENDO(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M04.M04_NENDO(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = M04.M04_KYOKAN_CD(+) AND "
'	 w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO = " & m_iNendo & " "
'	 'w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = '" & m_sKyokan_CD & "' "
'	 w_sSQL = w_sSQL & vbCrLf & " ORDER BY T47.T47_GAKKA_CD "



	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "  T47_KYOKASYO.T47_GAKKI_KBN "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_NO "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_GAKUNEN "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_GAKKA_CD "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KAMOKU "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KYOKAN "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KYOKASYO "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_SYUPPANSYA "
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_TYOSYA"
	w_sSQL = w_sSQL & vbCrLf & " FROM "
	w_sSQL = w_sSQL & vbCrLf & "  T47_KYOKASYO"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "  T47_KYOKASYO.T47_NENDO=" & m_iNendo

	If m_iGakunen <> "" Then
		w_sSQL = w_sSQL & vbCrLf & "  AND T47_KYOKASYO.T47_GAKUNEN=" & m_iGakunen
	End If

'2001/12/01 Mod ---->
'	If m_sGakkaCd <> "" Then
'		w_sSQL = w_sSQL & vbCrLf & "  AND T47_KYOKASYO.T47_GAKKA_CD='" & m_sGakkaCd & "'"
'	End If

	w_sSQL = w_sSQL & vbCrLf & "  AND T47_KYOKASYO.T47_GAKKA_CD='" & m_sSyozokuGakka & "'"

'2001/12/01 Mod <----

	w_sSQL = w_sSQL & vbCrLf & " ORDER BY "
	w_sSQL = w_sSQL & vbCrLf & "  T47_KYOKASYO.T47_GAKUNEN"
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_GAKKA_CD"
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KAMOKU"
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KYOKAN"
	w_sSQL = w_sSQL & vbCrLf & "  ,T47_KYOKASYO.T47_KYOKASYO "

'response.write("<BR>w_sSQL = " & w_sSQL)

	Set m_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		m_bErrFlg = True
		Exit function
	Else
		'�y�[�W���̎擾
		m_iMax = gf_PageCount(m_Rs,m_iDsp)
	End If

	f_GetData = True

End Function

'********************************************************************************
'*	[�@�\]	�w�Ȃ̗��̂��擾
'*	[����]	p_sGakkaCd : �w��CD
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetGakkaNm_R(p_sGakkaCd)
	Dim w_sSQL				'// SQL��
	Dim w_iRet				'// �߂�l
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakkaNm_R = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKARYAKSYO"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_GAKKA_CD='" & p_sGakkaCd & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M02_GAKKARYAKSYO")
		End If 

		Exit do 
	Loop

	'//�߂�l���Z�b�g
	f_GetGakkaNm_R = w_sName

	'//RS Close
	Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*	[�@�\]	�Ȗږ��̂��擾
'*	[����]	p_sGakkaCd : �w��CD
'*			p_sKamokuCd
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_GetKamokuNm(p_sGakkaCd,p_sKamokuCd)
	Dim w_sSQL				'// SQL��
	Dim w_iRet				'// �߂�l
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetKamokuNm = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & m_iNendo

		if cstr(gf_HTMLTableSTR(p_sGakkaCd)) <> cstr(C_CLASS_ALL) then
			w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & p_sGakkaCd & "'"
		End If
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_sKamokuCd & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("T15_KAMOKUMEI")
		End If 

		Exit do 
	Loop

	'//�߂�l���Z�b�g
	f_GetKamokuNm = w_sName

	'//RS Close
	Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*	[�@�\]	�ڍׂ�\��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub S_syousai()

	Dim w_iCnt
	Dim w_i
	Dim w_cell

	Dim w_lCnt			'�J�E���^
	Dim w_bTantoYes		'�S�����Ă���

	w_iCnt	= 0
	w_i 	= 0
	w_cell = ""

	Dim w_sCurGakkaCD		'2001/12/01 Add �������̊w�Ȃb�c

	Do While not m_Rs.EOF

		w_i = w_i + 1

		call gs_cellPtn(w_cell)
%>

		<Tr>
		<Td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_GAKUNEN")) %>�N</Td>
<%
		If CStr(gf_HTMLTableSTR(m_Rs("T47_GAKKA_CD"))) = CStr(C_CLASS_ALL) Then
			m_sGakka = "�S�w��"
			w_sCurGakkaCD = ""								'2001/12/01 Add
		Else
			m_sGakka = f_GetGakkaNm_R(m_Rs("T47_GAKKA_CD"))
			w_sCurGakkaCD = CStr(m_Rs("T47_GAKKA_CD"))		'2001/12/01 Add
		End If

		w_bTantoYes= False

		For w_lCnt = 0 To Ubound(m_sKamokuCD)
			If m_sKamokuCD(w_lCnt) = CStr(m_Rs("T47_KAMOKU")) Then
				w_bTantoYes = True
				Exit For
			End if
		Next

		'2001/12/01 Mod ---->
		If w_bTantoYes = True Then
		'== �S�w�Ȃ���������w�Ȃ̏ꍇ ==
%>
<!-- <%= m_sGakka %> -->
		<Td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_sGakka) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(f_GetKamokuNm(m_Rs("T47_GAKKA_CD"),m_Rs("T47_KAMOKU"))) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(gf_GetKyokanNm(m_iNendo,m_Rs("T47_KYOKAN"))) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><A HREF='javascript:f_LinkClick(<%=m_Rs("T47_NO")%>);'><%=gf_HTMLTableSTR(m_Rs("T47_KYOKASYO")) %></A></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_SYUPPANSYA")) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_TYOSYA")) %></Td>
		<Td align="center" width="30"  class=<%=w_cell%>><input class=button type="button" value=">>" onclick="javascript:f_Update(<%=gf_HTMLTableSTR(m_Rs("T47_NO")) %>)"></Td>
		<Td align="center" width="30"  class=<%=w_cell%>><input type="checkbox" name="deleteNO" value="<%=gf_HTMLTableSTR(m_Rs("T47_NO")) %>"></Td>

<%
		Else
		'== �S�w�Ȃł���������w�Ȃł��Ȃ��ꍇ ==
%>
<!-- <%= m_sGakka %> -->

		<Td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_sGakka) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(f_GetKamokuNm(m_Rs("T47_GAKKA_CD"),m_Rs("T47_KAMOKU"))) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(gf_GetKyokanNm(m_iNendo,m_Rs("T47_KYOKAN"))) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><A HREF='javascript:f_LinkClick(<%=m_Rs("T47_NO")%>);'><%=gf_HTMLTableSTR(m_Rs("T47_KYOKASYO")) %></A></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_SYUPPANSYA")) %></Td>
		<Td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_TYOSYA")) %></Td>
		<Td align="center" width="30"  class=<%=w_cell%>>�@</Td>
		<Td align="center" width="30"  class=<%=w_cell%>>�@</Td>

<%
		End If

		m_Rs.MoveNext

		If w_iCnt >= C_PAGE_LINE-1 Then
			Exit Sub
		Else
			w_iCnt = w_iCnt + 1
		End If
	Loop

	m_iDisp= w_i

End sub

Sub showPage()
'********************************************************************************
'*	[�@�\]	HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
	Dim w_bFlg				'// �ް��L��
	Dim w_bNxt				'// NEXT�\���L��
	Dim w_bBfr				'// BEFORE�\���L��
	Dim w_iNxt				'// NEXT�\���Ő�
	Dim w_iBfr				'// BEFORE�\���Ő�
	Dim w_iCnt				'// �ް��\������
	Dim w_pageBar			'�y�[�WBAR�\���p
	
	On Error Resume Next
	Err.Clear

	'�y�[�WBAR�\��
	Call gs_pageBar(w_pageBar)

	Dim w_iRecordCnt		'//���R�[�h�Z�b�g�J�E���g

	On Error Resume Next
	Err.Clear

	w_iCnt	= 1
	w_bFlg	= True

%>

	<html>

	<head>

	<title>�g�p���ȏ��o�^</title>
<!-- <%= m_sSyozokuGakka %>-->
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

	<!--

	//************************************************************
	//	[�@�\]	�ꗗ�\�̎��E�O�y�[�W��\������
	//	[����]	p_iPage :�\���Ő�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_PageClick(p_iPage){

		document.frm.action="";
		document.frm.target="";
		document.frm.txtMode.value = "PAGE";
		document.frm.txtPageCD.value = p_iPage;
		document.frm.submit();
	
	}

	//************************************************************
	//	[�@�\]	�X�V��ʂ�
	//	[����]	p_iPage :�\���Ő�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Update(p_No){

		document.frm.action="./touroku.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.txtMode.value = "Kousin";
		document.frm.txtUpdNo.value = p_No;
		document.frm.submit();

	}

	//************************************************************
	//	[�@�\]	�폜�y�[�W��
	//	[����]	p_iPage :�\���Ő�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Delete(){

		if (f_chk()==1){
		alert( "�폜�̑ΏۂƂȂ鋳�ȏ����I������Ă��܂���" );
		return;
		}

		document.frm.action="del_kakunin.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.txtMode.value = "DELETE";
		document.frm.submit();
	}
	//************************************************************
	//	[�@�\]	���X�g�ꗗ�̃`�F�b�N�{�b�N�X�̊m�F
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_chk(){

		var i;
		i = 0;

		//0���̂Ƃ�
		if (document.frm.txtDisp.value<=0){
			return 1;
			}

		//1���̂Ƃ�
		if (document.frm.txtDisp.value==1){
			if (document.frm.deleteNO.checked == false){
				return 1;
			}else{
				return 0;
				}
		}else{
		//����ȊO�̎�
		var checkFlg
			checkFlg=false

		do { 
			
			if(document.frm.deleteNO[i].checked == true){
				checkFlg=true
				break;
			 }

		i++; }	while(i<document.frm.txtDisp.value);
			if (checkFlg == false){
				return 1;
				}
		}
		return 0;
	}

	//************************************************************
	//	[�@�\]	�����N�N���b�N
	//	[����]
	//	[�ߒl]
	//	[����]
	//************************************************************
	function f_LinkClick(p_No){
		document.frm.txtUpdNo.value = p_No;
		document.frm.action="view.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
		document.frm.submit();
	}

	//-->
	</SCRIPT>
	<link rel=stylesheet href="../../common/style.css" type=text/css>

	</head>

	<body>

	<center>
	<br>

	<form name="frm" action="touroku.asp" target="" Method="POST">

	<%
	'�f�[�^�Ȃ��̏ꍇ
	If m_Rs.EOF Then
	%>
		<br><br><br>
		<span class="msg">�Ώۃf�[�^�͑��݂��܂���B��������͂��Ȃ����Č������Ă��������B</span>
	<%Else%>


		<span class="msg"><font size="2">�����ȏ������N���b�N����Əڍד��e���Q�Ƃł��܂��B</font></span>
	<%Call gs_pageBar(m_Rs,m_sPageCD,m_iDsp,w_pageBar)%>

		<table width=90%>
			<Tr><Td><%=w_pageBar %></Td></Tr>

			<Tr><Td>
				<table border="1" width="90%" class=hyo>
				<Tr>
					<Th width="70"	class=header nowrap>�w�N</Th>
					<Th width="70"	class=header nowrap>�w��</Th>
					<Th width="110" class=header nowrap>�Ȗ�</Th>
					<Th width="110" class=header nowrap>������</Th>
					<Th width="150" class=header nowrap>���ȏ���</Th>
					<Th width="90"	class=header nowrap>�o�Ŏ�</Th>
					<Th width="90"	class=header nowrap>����</Th>
					<Th width="30"	class=header >�C��</Th>
					<Th width="30"	class=header>�폜</Th>
				</Tr>

					<% S_syousai() %>
				<Tr>
					<Td colspan=9 align=right bgcolor=#9999BD>
					<input class=button type=button value="�~�폜" Onclick="f_Delete()"></Td>
				</Tr>

				</table>
			</Td></Tr>
<!--
<% = Ubound(m_sKamokuCD) %>
<%
		For w_lCnt = 0 To Ubound(m_sKamokuCD)
			response.write(m_sKamokuCD(w_lCnt))
		Next
%>
-->
			<Tr><Td><%=w_pageBar %></Td></Tr>
		<table>
	<%End If%>

	<!--�l�n�p-->
	<input type="hidden" name="txtMode" value="Touroku">
	<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
	<input type="hidden" name="txtDisp" value="<%= m_iDisp %>">
	<input type="hidden" name="txtUpdNo" value="">
	<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">

	<input type="hidden" name="KeyNendo" value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokan_CD%>">
	<input type="hidden" name="SKyokanCd1" value="<%=m_sKyokan_CD%>">

	<input type="hidden" name="txtGakunenCd" value="<%= Request("txtGakunenCd") %>">
	<input type="hidden" name="txtGakkaCD"	 value="<%= Request("txtGakkaCD") %>">

	</form>
	</center>
	</body>
	</html>

<%
End Sub


Function f_GetTantoKamoku()
'********************************************************************************
'*	[�@�\]�@�S���������ǂ����`�F�b�N
'*	[����]�@ �Ȃ�
'*	[�ߒl]�@True:�S�����������Ă���AFalse:�S�����������Ă��Ȃ�
'*	[����]
'********************************************************************************
	Dim w_iRet			'�߂�l
	Dim w_sSQL			'SQL
	Dim w_oRecord		'���R�[�h
	Dim w_lCnt			'���R�[�h�J�E���g

	f_GetTantoKamoku = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "     T27_KAMOKU_CD"
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & "     T27_TANTO_KYOKAN T27"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "     T27_NENDO = " & SESSION("NENDO") & " "
	If request("txtGakunen") <> "" Then
		w_sSQL = w_sSQL & " AND "
		w_sSQL = w_sSQL & " T27_GAKUNEN = " & request("txtGakunen") & " "
	End If
	w_sSQL = w_sSQL & " AND "
	w_sSQL = w_sSQL & " T27_KYOKAN_CD = '" & SESSION("KYOKAN_CD") & "' "
	w_sSQL = w_sSQL & " GROUP BY T27_KAMOKU_CD"
	w_sSQL = w_sSQL & " ORDER BY T27_KAMOKU_CD"

	Set w_oRecord = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset_OpenStatic(w_oRecord, w_sSQL)

	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
'		response.write(w_sSQL)
		Exit Function
	End If

	'//�S�����Ă��Ȃ��ꍇ
	If w_oRecord.EOF = True Then
		ReDim m_sKamokuCD(0)

		f_GetTantoKamoku = True
		Exit Function
	End If

	w_lCnt = gf_GetRsCount(w_oRecord)
'	w_oRecord.MoveFirst

	ReDim m_sKamokuCD(w_lCnt)

	'// �S���Ȗڂ̕ێ�
	For w_lCnt = 0 To Ubound(m_sKamokuCD) - 1
		m_sKamokuCD(w_lCnt) = CStr(w_oRecord("T27_KAMOKU_CD"))

		w_oRecord.MoveNext
	Next

	w_oRecord.Close
	Set w_oRecord = Nothing

	f_GetTantoKamoku = True
'response.write("True<BR>")

End Function






%>