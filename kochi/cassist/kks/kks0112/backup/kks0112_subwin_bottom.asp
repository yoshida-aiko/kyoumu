<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0111/kks0111_detail_bottom.asp
' �@      �\: ���y�[�W ���Əo�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
'             TUKI           '//��
' ��      ��:
' ��      �n: NENDO          '//�����N
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��s���o�����͂�\��
'           ���o�^�{�^���N���b�N��
'               ���͏���o�^����
'-------------------------------------------------------------------------
' ��      ��: 2002/05/07 shin
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public m_bErrFlg			'//�G���[�t���O
	Public m_bNoDataFlg			'//�f�[�^�Ȃ��t���O
	
	'�擾�����f�[�^�����ϐ�
	Public m_iSyoriNen      '//�����N�x
	
	Public m_sGakunenCd
	Public m_sClassCd
	Public m_sFromDate
	Public m_sToDate
	Public m_sGakusekiNo
	
	Public m_sKamokuCd		'//�ȖڃR�[�h
	
	Public m_sSyubetu		'//���
	Public m_iMonth
	
	Public m_AryJigen()		'//
	Public m_AryState()		
	
	Public m_Count
	
	Public m_Rs				'//���R�[�h�Z�b�g
	Public m_JigenCount
	Public m_AryXCount
	Public m_StudentCount
	
	Public m_sGakki
	Public m_sZenki_Start
	Public m_sKouki_Start
	Public m_sKouki_End
	
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
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���Əo������"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"
	
	On Error Resume Next
    Err.Clear

    m_bErrFlg = False
	m_bNoDataFlg = false
	
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
		
		'// ���Ұ�SET
        Call s_SetParam()
		
		'//�O���E��������擾
		if gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End) <> 0 then
			m_bErrFlg = True
			exit do
		end if
		
		'//�������̎擾
		if not f_HeadData() then
			m_bErrFlg = True
			exit do
		end if
		
		'//������񂪂Ȃ��Ƃ�
		if m_bNoDataFlg = true then
			Call showWhitePage("�I�����ꂽ�����ł́A���Ə��͂���܂���")
			exit do
		end if
		
		'//���k���ۏ��̎擾
		if not f_Get_KekkaData() then
			m_bErrFlg = True
			exit do
		end if
		
		'//���k��񂪂Ȃ��Ƃ�
		if m_bNoDataFlg = true then
			Call showWhitePage("���k��񂪂���܂���")
			exit do
		end if
		
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
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()
	
	m_iSyoriNen = 0
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	
	m_sKamokuCd = ""
    
    m_sSyubetu = ""
	m_iMonth = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen = Session("NENDO")
	
	m_sGakunenCd = request("hidGakunen")
	m_sClassCd = request("hidClassNo")
	
    m_sKamokuCd = request("hidKamokuCd")
    
    m_sSyubetu = request("hidSyubetu")
	m_iMonth = request("sltMonth")
    
End Sub
'********************************************************************************
'*  [�@�\]  �w�b�_���擾
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
function f_HeadData()
	Dim w_sSQL
	Dim w_iRet
	Dim w_sSDate,w_sEDate
	Dim w_num
	
	On Error Resume Next
	Err.Clear
	
	f_HeadData = false
	
	w_sSDate = ""
	w_sEDate = ""
	
	Call f_GetTukiRange(w_sSDate,w_sEDate)
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T21_HIDUKE, "
	w_sSQL = w_sSQL & "  T21_JIGEN "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T21_SYUKKETU "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "     T21_HIDUKE >='" & w_sSDate & "' "
	w_sSQL = w_sSQL & " and T21_HIDUKE <'" & w_sEDate & "' "
	w_sSQL = w_sSQL & " and T21_NENDO =" & m_iSyoriNen 
	w_sSQL = w_sSQL & " and T21_GAKUNEN =" & m_sGakunenCd 
	w_sSQL = w_sSQL & " and T21_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & " and T21_KAMOKU ='" & m_sKamokuCd & "'"
	w_sSQL = w_sSQL & " and T21_SYUKKETU_KBN in(1,2,3)"
	
	w_sSQL = w_sSQL & " group by T21_HIDUKE,T21_JIGEN"
	w_sSQL = w_sSQL & " order by T21_HIDUKE,T21_JIGEN"
	
	w_iRet = gf_GetRecordset(m_Rs,w_sSQL)
	
	if w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	m_JigenCount = gf_GetRsCount(m_Rs)
	
	if m_JigenCount = 0 then
		m_bNoDataFlg = true
		f_HeadData = true
		exit function
	end if
	
	ReDim Preserve m_AryJigen(m_JigenCount,3)
	
	for w_num = 0 to m_JigenCount-1
		
		m_AryJigen(w_num,0) = day(m_Rs("T21_HIDUKE"))
		m_AryJigen(w_num,1) = m_Rs("T21_JIGEN")
		m_AryJigen(w_num,2) = m_Rs("T21_HIDUKE")
		
		m_Rs.movenext
	next
	
	f_HeadData = true
	
end function


'********************************************************************************
'*	[�@�\]	���̌����������쐬(7���c�@"MONTH>=2001/07/01 AND MONTH<2001/08/01" �Ƃ��Ďg�p)
'*	[����]	�Ȃ�
'*	[�ߒl]	p_sSDate
'*			p_sEDate
'*	[����]	
'********************************************************************************
Function f_GetTukiRange(p_sSDate,p_sEDate)
	Dim w_iNen
	
	p_sSDate = ""
	p_sEDate = ""
	
	if 4 <= cInt(m_iMonth) and cInt(m_iMonth) <=12 then
		w_iNen = cint(m_iSyoriNen)
		
		'//�J�n��
		If cint(month(m_sZenki_Start)) = Cint(m_iMonth) Then
			p_sSDate = m_sZenki_Start
		Else
			p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iMonth),2) & "/01"
		End If
		
		'//�I����
		If cint(month(m_sKouki_Start)) = Cint(m_iMonth) Then
			p_sEDate = m_sKouki_Start
		Else 
			If Cint(m_iMonth) = 12 Then
				p_sEDate = cstr(w_iNen+1) & "/01/01"
			Else
				p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iMonth+1),2) & "/01"
			End If
		End If
		
	Else
		'//����̔N
		If cint(m_iMonth) <=4 Then
			w_iNen = cint(m_iSyoriNen) + 1
		Else
			w_iNen = cint(m_iSyoriNen)
		End If
		
		'//�J�n��
		If cint(month(m_sKouki_Start)) = Cint(m_iMonth) Then
			p_sSDate = m_sKouki_Start
		Else
			p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iMonth),2) & "/01"
		End If
		
		'//�I����
		If cint(month(m_sKouki_End)) = Cint(m_iMonth) Then
			'p_sEDate = m_sKouki_End
			p_sEDate = DateAdd("d",1,m_sKouki_End)
		Else 
			If Cint(m_sTuki) = 12 Then
				p_sEDate = cstr(w_iNen+1) & "/01/01"
			Else
				p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iMonth+1),2) & "/01"
			End If
		End If
		
	End If
	
End Function

'********************************************************************************
'*	[�@�\]	���k���擾,���ۥ�x�������̎擾
'*	[����]	
'*	[�ߒl]	true:���擾���� false:���s
'*	[����]	
'********************************************************************************
function f_Get_KekkaData()
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_num,w_Jnum
	
	On Error Resume Next
	Err.Clear
	
	f_Get_KekkaData = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T13.T13_GAKUSEKI_NO,"
	w_sSQL = w_sSQL & "  T11.T11_SIMEI "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T13_GAKU_NEN T13,"
	w_sSQL = w_sSQL & "  T11_GAKUSEKI T11 "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  T13.T13_GAKUSEI_NO = T11.T11_GAKUSEI_NO "
	w_sSQL = w_sSQL & "  and T13.T13_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and T13.T13_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T13.T13_CLASS =" & m_sClassCd
	
	w_sSQL = w_sSQL & "  group by T11.T11_SIMEI,T13.T13_GAKUSEKI_NO "
	w_sSQL = w_sSQL & "  order by T13.T13_GAKUSEKI_NO "
	
	w_iRet = gf_GetRecordset(m_Rs_Student,w_sSQL)
	
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	m_StudentCount = gf_GetRsCount(m_Rs_Student)		'//���k��
	
	'//�f�[�^�Ȃ�
	if m_StudentCount = 0 then
		m_bNoDataFlg = true
		f_Get_KekkaData = true
		exit function
	end if
	
	m_AryXCount = 2 + m_JigenCount						'//(�w��NO+���k����) + ������
	
	ReDim Preserve m_AryState(m_AryXCount,m_StudentCount)	'//���ۥ�x�����Z�b�g�z��
	
	for w_num = 0 to m_StudentCount
		
		m_AryState(0,w_num) = m_Rs_Student(0)		'�w��NO
		m_AryState(1,w_num) = m_Rs_Student(1)		'���k����
		
		for w_Jnum = 0 to m_JigenCount - 1
			'w_Jnum�����̌��ۋ敪
			if not f_SetKekka(m_Rs_Student(0),m_AryJigen(w_Jnum,1),m_AryJigen(w_Jnum,2),m_AryState(2+w_Jnum,w_num)) then exit function
		next
		
		m_Rs_Student.movenext
		
	next
	
	f_Get_KekkaData = true
	
end function

'********************************************************************************
'*	[�@�\]	���ۥ�x�����̎擾
'*	[����]	p_GakusekiNo���w��NO
'*			p_Jigen������
'*			p_Type��C_KEKKA:���ې�,C_TIKOKU:�x����
'*			p_Kikan��C_ZENKI:�O��,C_KOUKI:���,C_KOUKI:�O�������ȊO
'*
'*	[�ߒl]	0:���擾���� 99:���s�Ap_KekkaNum�����ۥ�x����
'*	[����]	
'********************************************************************************
function f_SetKekka(p_GakusekiNo,p_Jigen,p_Hiduke,p_KekkaType)
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_Rs
	Dim w_KekkaName
	
	On Error Resume Next
	Err.Clear
	
	f_SetKekka = false
	
	p_KekkaNum = 0
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T21_SYUKKETU_KBN, "
	w_sSQL = w_sSQL & "  T21_JIKANSU, "
	w_sSQL = w_sSQL & "  M01_SYOBUNRUIMEI_R "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T21_SYUKKETU, "
	w_sSQL = w_sSQL & "  M01_KUBUN "
	
	w_sSQL = w_sSQL & " where T21_NENDO = " & m_iSyoriNen
	
	w_sSQL = w_sSQL & "  and M01_DAIBUNRUI_CD =" & C_KESSEKI
	w_sSQL = w_sSQL & "  and M01_NENDO =" & m_iSyoriNen
	w_sSQL = w_sSQL & "  and T21_SYUKKETU_KBN = M01_SYOBUNRUI_CD(+) "
	
	w_sSQL = w_sSQL & "  and T21_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T21_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & "  and T21_GAKUSEKI_NO ='" & p_GakusekiNo & "'"
	w_sSQL = w_sSQL & "  and T21_JIGEN =" & p_Jigen
	w_sSQL = w_sSQL & "  and T21_HIDUKE ='" & p_Hiduke & "'"
	
	w_iRet = gf_GetRecordset(w_Rs,w_sSQL)
	
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	'Dim w_IdouType,w_KubunName
	'w_IdouType = cint(gf_SetNull2Zero(gf_Get_IdouChk(p_GakusekiNo,p_Hiduke,m_iSyoriNen,w_KubunName)))
	
	if cInt(gf_SetNull2Zero(w_Rs(0))) <> cInt(C_KETU_KEKKA) then
		p_KekkaType = w_Rs(2)
	else
		p_KekkaType = w_Rs(1) & w_Rs(2)
	end if
	
	f_SetKekka = true
	
end function 

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	Dim w_Class
	Dim w_Jnum,w_num
	
	w_Class = ""
	
    On Error Resume Next
    Err.Clear
	
%>
    <html>
    <head>
    <title>���Əo������</title>
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
		
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    	<table>
			<tr>
				<td valign="top" nowrap>
					<table class="hyo"  border="1">
						<% for w_num = 0 to m_StudentCount-1 %>
							<% Call gs_cellPtn(w_Class) %>
							<tr>
								<td width="80" class="<%=w_Class%>" align="center" nowrap><%=m_AryState(0,w_num)%></td>
								<td width="130"  class="<%=w_Class%>" align="center" nowrap><%=m_AryState(1,w_num)%></td>
								
								<%for w_Jnum = 0 to m_JigenCount-1 %>
									<td width="40" class="<%=w_Class%>" align="center" nowrap><%=gf_HTMLTableSTR(m_AryState(2+w_Jnum,w_num))%></td>
								<%next%>
							</tr>
							
						<% next%>
						
					</table>
				</td>
			</tr>
			
        </table>
	</form>
    </center>
    </body>
    </html>
<%
End Sub

'********************************************************************************
'*	[�@�\]	��HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
	<html>
	<head>
	<title>���Əo������</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	
	//************************************************************
	//	[�@�\]	�y�[�W���[�h������
	//	[����]
	//	[�ߒl]
	//	[����]
	//************************************************************
	function window_onload() {
	}
	//-->
	</SCRIPT>
	
	</head>
	<body LANGUAGE="javascript" onload="return window_onload()">
	<form name="frm" mothod="post">
	
	<center>
	<br><br><br>
		<span class="msg"><%=Server.HTMLEncode(p_Msg)%></span>
	</center>
	
	<input type="hidden" name="txtMsg" value="<%=Server.HTMLEncode(p_Msg)%>">
	</form>
	</body>
	</html>
<%
End Sub
%>
