<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo���\��(�ڍ�)
' ��۸���ID : kks/kks0112/kks0112_subwin_top.asp
' �@      �\: ��y�[�W ���Əo�������X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
'             
' ��      ��:
' ��      �n: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
' ��      ��:
'            
'-------------------------------------------------------------------------
' ��      ��: 2002/05/07 shin
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	
	Public m_bErrFlg		'//�G���[�t���O
	
	Public m_iSyoriNen		'//�����N�x
	Public m_sGakunenCd		'//�w�N
	Public m_sClassCd		'//�N���XCD
	
	Public m_sKamokuCd		'//�ȖڃR�[�h
    
    Public m_iMonth			'//��
	
	Public m_sGakki
	Public m_sZenki_Start
	Public m_sKouki_Start
	Public m_sKouki_End
	
	Public m_Rs				'//���R�[�h�Z�b�g
	
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
	
    Do
        '// �ް��ް��ڑ�
        If gf_OpenDatabase() <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            w_sMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            'm_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
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
		
		'//�w�b�_���擾
		if not f_HeadData() then
			m_bErrFlg = True
			exit do
		end if
		
		if m_Rs.EOF then
			'Call showWhitePage("�ΏۂƂȂ�A���Ə�񂪂���܂���")
			exit do
		end if
		
		'//�y�[�W�\��
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
	
	m_iSyoriNen = 0
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	
    m_sKamokuCd = ""
    
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
	Dim w_sSDate,w_sEDate
	
	On Error Resume Next
	Err.Clear
	
	f_HeadData = false
	
	w_sSDate = ""
	w_sEDate = ""
	
	'//�w�茎�̊J�n���A�I�������Q�b�g
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
	
	w_sSQL = w_sSQL & " group by T21_HIDUKE,T21_JIGEN "
	w_sSQL = w_sSQL & " order by T21_HIDUKE,T21_JIGEN "
	
	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	m_Rs.movefirst
	
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
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
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
    //************************************************************
    function window_onload() {
		//�X�N���[����������
		//parent.init();
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    	<BR>
		
		<table>
			<tr>
				<td valign="bottom" nowrap>
					<table class="hyo" border="1">
						<tr>
							<th width="80" class="header" rowspan="2" nowrap>�w�Дԍ�</th>
							<th width="80" class="header" rowspan="2" nowrap>����</th>
							
							<% if not (m_Rs is nothing) then %>
								
								<th width="40" class="header" align="center" nowrap>���t</th>
								
								<% do until m_Rs.EOF %>
									<th width="40" class="header" align="center" nowrap><%=day(m_Rs("T21_HIDUKE"))%></th>
									
									<% m_Rs.movenext %>
								<% loop %>
							<% end if %>
							
						</tr>
						
						<tr>
							<% if not (m_Rs is nothing) then %>
								<% m_Rs.movefirst %>
								
								<th width="40" class="header" align="center" nowrap>����</th>
								
								<% do until m_Rs.EOF %>
									<th width="40" class="header" align="center" nowrap><%=m_Rs("T21_JIGEN")%></th>
									
									<% m_Rs.movenext %>
								<% loop %>
							<% end if %>	
						</tr>
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
	<body LANGUAGE=javascript onload="return window_onload()">
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
