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
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public m_bErrFlg           '�װ�׸�
	
	'�擾�����f�[�^�����ϐ�
	Public m_iSyoriNen      '//�����N�x
	
	Public m_sGakunenCd
	Public m_sClassCd
	Public m_sFromDate
	Public m_sToDate
	Public m_sGakusekiNo
	
	Public m_AryKekkaMei()  '//���ۖ��̊i�[�z��
	
	Public m_Count
	
	Public m_Rs				'//���R�[�h�Z�b�g
	Public m_bDataNon		'//�ڍ׃f�[�^�t���O
	Public m_RecCount		'//���R�[�h����
	
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
		
		if not f_SetKekka() then
			m_bErrFlg = True
			Exit Do
		end if
		
		'�ڍ׃f�[�^���Ȃ�
		if m_bDataNon = true then
			Call showWhitePage("�ڍ׃f�[�^�́A����܂���")
			exit do
		end if
		
		if not f_Get_SyukketuKbn() then
			m_bErrFlg = True
			Exit Do
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
	m_sGakunenCd = ""
	m_sClassCd = ""
	m_sGakusekiNo = ""
	
	m_Count = 0
	
    m_iSyoriNen = ""
    
	m_bDataNon = false
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	m_sGakunenCd = request("Nen")
	m_sClassCd = request("Class")
	
	m_sFromDate = request("FromDate")
	m_sToDate = request("ToDate")
	m_sGakusekiNo = request("GakusekiNo")
	
    m_iSyoriNen = Session("NENDO")
    
End Sub

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
function f_SetKekka()
	
	Dim w_sSQL
	Dim w_iRet
	
	On Error Resume Next
	Err.Clear
	
	f_SetKekka = false
	
	p_KekkaNum = 0
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T21_HIDUKE, "
	w_sSQL = w_sSQL & "  T21_JIGEN, "
	w_sSQL = w_sSQL & "  T21_KYOKAN, "
	
	w_sSQL = w_sSQL & "  M03_KAMOKUMEI, "
	w_sSQL = w_sSQL & "  M04_KYOKANMEI_SEI, "
	w_sSQL = w_sSQL & "  M04_KYOKANMEI_MEI, "
	w_sSQL = w_sSQL & "  T21_SYUKKETU_KBN, "
	w_sSQL = w_sSQL & "  T21_JIKANSU "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T21_SYUKKETU, "
	w_sSQL = w_sSQL & "  M03_KAMOKU, "
	w_sSQL = w_sSQL & "  M04_KYOKAN "
	
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "      T21_SYUKKETU.T21_KAMOKU = M03_KAMOKU.M03_KAMOKU_CD "
	w_sSQL = w_sSQL & "  and T21_SYUKKETU.T21_KYOKAN = M04_KYOKAN.M04_KYOKAN_CD(+) "
	
	w_sSQL = w_sSQL & "  and T21_SYUKKETU.T21_NENDO = M03_KAMOKU.M03_NENDO "
	w_sSQL = w_sSQL & "  and T21_SYUKKETU.T21_NENDO = M04_KYOKAN.M04_NENDO "
	
	w_sSQL = w_sSQL & "  and T21_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T21_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & "  and T21_GAKUSEKI_NO ='" & m_sGakusekiNo & "'"
	
	w_sSQL = w_sSQL & "  and T21_HIDUKE >='" & m_sFromDate & "'"
	w_sSQL = w_sSQL & "  and T21_HIDUKE <='" & m_sToDate & "'"
	
	w_sSQL = w_sSQL & "  and (T21_SYUKKETU_KBN =" & C_KETU_TIKOKU
	w_sSQL = w_sSQL & "       or T21_SYUKKETU_KBN = " & C_KETU_SOTAI
	w_sSQL = w_sSQL & "       or T21_SYUKKETU_KBN = " & C_KETU_KEKKA
	w_sSQL = w_sSQL & "      )"
	
	w_sSQL = w_sSQL & " order by T21_HIDUKE,T21_JIGEN "
	
	w_iRet = gf_GetRecordset(m_Rs,w_sSQL)
	
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	'�f�[�^���Ȃ�
	if m_Rs.EOF then m_bDataNon = true
	
	m_RecCount = gf_GetRsCount(m_Rs)
	
	f_SetKekka = true
	
end function 

'********************************************************************************
'*	[�@�\]	�o���敪���̂̎擾(�z��ɃZ�b�g)
'*	[����]	�Ȃ�
'*	[�ߒl]	
'*	[����]	
'********************************************************************************
function f_Get_SyukketuKbn()
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_Rs_Kekka
	
	On Error Resume Next
	Err.Clear
	
	f_Get_SyukketuKbn = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  M01_SYOBUNRUIMEI,M01_SYOBUNRUI_CD "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  M01_KUBUN "
	
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "      M01_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and M01_DAIBUNRUI_CD = " & C_KESSEKI
	w_sSQL = w_sSQL & " order by  M01_SYOBUNRUI_CD "
	
	w_iRet = gf_GetRecordset(w_Rs_Kekka,w_sSQL)
	
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	m_Count = gf_GetRsCount(w_Rs_Kekka) - 1
	
	ReDim Preserve m_AryKekkaMei(2,m_Count)
	
	for w_num = 0 to m_Count
		m_AryKekkaMei(0,w_num) = w_Rs_Kekka("M01_SYOBUNRUI_CD")
		m_AryKekkaMei(1,w_num) = w_Rs_Kekka("M01_SYOBUNRUIMEI")
		w_Rs_Kekka.movenext
	Next
	
	f_Get_SyukketuKbn = true
	
end function 

'********************************************************************************
'*	[�@�\]	�o�����̎擾
'*	[����]	p_SyukketuKbn:�o���敪
'*	[�ߒl]	�o������
'*	[����]	
'********************************************************************************
function f_Set_SyukketuMei(p_SyukketuKbn)
	Dim w_num
	
	for w_num = 0 to m_Count
		
		if cInt(m_AryKekkaMei(0,w_num)) = cInt(p_SyukketuKbn) then
			f_Set_SyukketuMei = m_AryKekkaMei(1,w_num)
			exit function
		end if
		
	Next
	
	f_Set_SyukketuMei = ""
	
end function


'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	Dim w_Name			'���O
	Dim w_SyukketuName	'�o������
	Dim w_Class			'td class�Z�b�g
	
	w_Name = ""
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
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    	<table>
	        <tr>
	        	<td>
	        		<table width="540" class="hyo"  border="1" >
	        			<% if not (m_Rs is nothing) then %>
					    	<% Do until m_Rs.EOF %>
								<% Call gs_cellPtn(w_Class) %>
								
								<tr>
									<td width="130" class="<%=w_Class%>" align="center" nowrap><%=m_Rs("T21_HIDUKE")%></td>
									<td width="60"  class="<%=w_Class%>" align="center" nowrap><%=m_Rs("T21_JIGEN")%></td>
									<td width="150" class="<%=w_Class%>" align="center" nowrap><%=m_Rs("M03_KAMOKUMEI")%></td>
									
									<%
										w_Name = ""
										
										if cstr(m_Rs("T21_KYOKAN")) = "1" then
											w_Name = "����"
										else
											w_Name = m_Rs("M04_KYOKANMEI_SEI") & "�@" & m_Rs("M04_KYOKANMEI_MEI")
										end if
									%>
									
									<td width="120" class="<%=w_Class%>" align="center" nowrap><%=w_Name%></td>
									
									<%
										w_SyukketuName = ""
										
										if cInt(m_Rs("T21_SYUKKETU_KBN")) = cInt(C_KETU_KEKKA) then
											w_SyukketuName = gf_SetNull2String(m_Rs("T21_JIKANSU")) & f_Set_SyukketuMei(m_Rs("T21_SYUKKETU_KBN"))
										else
											w_SyukketuName = f_Set_SyukketuMei(m_Rs("T21_SYUKKETU_KBN"))
										end if
									%>
									
						            <td width="80" class="<%=w_Class%>" align="center" nowrap><%=w_SyukketuName%></td>
			            		</tr>
			            		
			            		<% m_Rs.movenext %>
			            	<% Loop %>
						<% end if %>
	            		
	            	</table>
	            </td>
	        </tr>
	        
	        <% if m_RecCount >= 20 then %><!-- 20���ȏゾ�ƕ\������ -->
	        <tr>
				<td align="center"><input type="button" value="����" onClick="javascript:parent.close();"></td>
			</tr>
			<% end if %>
			
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
