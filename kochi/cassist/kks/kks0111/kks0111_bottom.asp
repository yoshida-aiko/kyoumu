<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo���Q��
' ��۸���ID : kks/kks0111/kks0111_bottom.asp
' �@	  �\: ���y�[�W ���Əo�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��	  ��: NENDO 		 '//�����N
'			  GAKUNEN		 '//�w�N
'			  CLASSNO		 '//�׽No
' ��	  ��:
' ��	  �n: NENDO 		 '//�����N
'			  GAKUNEN		 '//�w�N
'			  CLASSNO		 '//�׽No
' ��	  ��:
'			�������\��
'				���������ɂ��Ȃ��s���o�����͂�\��
'			���o�^�{�^���N���b�N��
'				���͏���o�^����
'-------------------------------------------------------------------------
' ��	  ��: 2002/05/07 shin
' ��	  �X: 2015.03.19 kiyomoto Win7�Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
	Const C_KEKKA = 1		'//���ې�
	Const C_TIKOKU = 2		'//�x����
	
	Const C_ZENKI = 1		'//�O��
	Const C_KOUKI = 2		'//���
	Const C_OTHER = 3		'//�O�������ȊO
	
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public	m_bErrFlg		'//�G���[�t���O
	
	Public m_iSyoriNen		'//�����N�x
	
	Public m_sGakki 		'//�w��
	Public m_sZenki_Start	'//�O���J�n��
	Public m_sKouki_Start	'//����J�n��
	Public m_sKouki_End 	'//����I����
	
	'ں��ރZ�b�g
	Public m_Rs_Student
	Public m_Rs_Jigen
	
	Public m_StudentCount	'//���k��
	Public m_AryKekka()		'//���ۥ�x�����Z�b�g�z��
	Public m_AryKei()		'//�O�������ʂ̌��ۥ�x�����Z�b�g�z��
	
	Public m_JigenCount		'//������
	Public m_AryXCount		'//�z��̗�
	
	Public m_sGakunenCd
	Public m_sClassCd
	Public m_sFromDate
	Public m_sToDate
	
	Public m_bDataNon		'�f�[�^���݃t���O
	
'///////////////////////////���C������/////////////////////////////

	'Ҳ�ٰ�ݎ��s
	Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*	[�@�\]	�{ASP��Ҳ�ٰ��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub Main()
	
	Dim w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget
	
	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���Əo������"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_top"
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = False
	m_bDataNon = false
	
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
		
		'//�ϐ�������
		Call s_ClearParam()
		
		'//�p�����[�^SET
		Call s_SetParam()
		
		'//�O���E��������擾
		if gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End) <> 0 then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//�����N�x�̎������擾
		If not f_Get_JigenData() Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//������񂪂Ȃ��Ƃ�(M07_JIGEN)
		if m_bDataNon = true then
			Call showWhitePage("������񂪂���܂���")
			Exit Do
		end if
		
		'//���k�A���ہA�x�����̎擾
		If not f_Get_KekkaData() Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//���k��񂪂Ȃ��Ƃ�(T13_GAKU_NEN,T11_GAKUSEKI)
		if m_bDataNon = true then
			Call showWhitePage("���k��񂪂���܂���")
			Exit Do
		end if
		
		'// �f�[�^�\���y�[�W��\��
		Call showPage()
		
		Exit Do
	Loop
	
	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
	End If
	
	'// �I������
	Call gf_closeObject(m_Rs_Student)
	Call gf_closeObject(m_Rs_Jigen)
	
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[�@�\]	�ϐ�������
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_ClearParam()
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	m_sFromDate = ""
	m_sToDate = ""
	
	m_iSyoriNen = 0
	
	m_sGakki	= ""
	m_sZenki_Start = ""
	m_sKouki_Start = ""
	m_sKouki_End = ""
	
End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_SetParam()
	
	m_sGakunenCd = request("cboGakunenCd")
	m_sClassCd = request("cboClassCd")
	m_sFromDate = gf_YYYY_MM_DD(request("txtFromDate"),"/")
	m_sToDate = gf_YYYY_MM_DD(request("txtToDate"),"/")
	
	m_iSyoriNen = Session("NENDO")
	
End Sub

'********************************************************************************
'*	[�@�\]	�����N�x�̎������̎擾
'*	[����]	
'*	[�ߒl]	true:���� false:���s
'*	[����]	
'********************************************************************************
function f_Get_JigenData()
	Dim w_sSQL
	
	On Error Resume Next
	Err.Clear
	
	f_Get_JigenData = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  MAX(M07_JIKAN) "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  M07_JIGEN "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  M07_NENDO = " & m_iSyoriNen
	
	If gf_GetRecordset(m_Rs_Jigen,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	'//�f�[�^�Ȃ�
	if m_Rs_Jigen.EOF then
		m_bDataNon = true
		f_Get_JigenData = true
		exit function
	end if
	
	m_JigenCount = cInt(m_Rs_Jigen(0))
	
	f_Get_JigenData = true
	
end function

'********************************************************************************
'*	[�@�\]	���k���擾,���ۥ�x�������̎擾
'*	[����]	
'*	[�ߒl]	true:���擾���� false:���s
'*	[����]	
'********************************************************************************
function f_Get_KekkaData()
	
	Dim w_sSQL
	Dim w_num,w_Jnum
	Dim w_KekkaNum
	
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
	
	If gf_GetRecordset(m_Rs_Student,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	'//�f�[�^�Ȃ�
	if m_Rs_Student.EOF then
		m_bDataNon = true
		f_Get_KekkaData = true
		exit function
	end if
	
	m_StudentCount = gf_GetRsCount(m_Rs_Student) - 1		'���k��
	
	m_AryXCount = 2 + m_JigenCount * 2						'(�w��NO+���k����) + ������ * 2
	
	ReDim Preserve m_AryKekka(m_AryXCount,m_StudentCount)	'//���ۥ�x�����Z�b�g�z��
	ReDim Preserve m_AryKei(4,m_StudentCount)				'//�O�������ʂ̌��ۥ�x�����Z�b�g�z��
	
	for w_num = 0 to m_StudentCount
		
		m_AryKekka(0,w_num) = m_Rs_Student(0)		'�w��NO
		m_AryKekka(1,w_num) = m_Rs_Student(1)		'���k����
		
		for w_Jnum = 1 to m_JigenCount
			'w_Jnum�����̌��ې��̎擾
			if not f_SetKekka(m_Rs_Student(0),w_Jnum,C_KEKKA,C_OTHER,m_AryKekka(w_Jnum*2+1,w_num)) then exit function
			
			'w_Jnum�����̒x������ސ��̎擾
			if not f_SetKekka(m_Rs_Student(0),w_Jnum,C_TIKOKU,C_OTHER,m_AryKekka(w_Jnum*2+2,w_num)) then exit function
		next
		
		'�O���̌��ې��̎擾
		if not f_SetKekka(m_Rs_Student(0),0,C_KEKKA,C_ZENKI,m_AryKei(0,w_num)) then exit function
		
		'�O���̒x������ސ��̎擾
		if not f_SetKekka(m_Rs_Student(0),0,C_TIKOKU,C_ZENKI,m_AryKei(1,w_num)) then exit function
		
		'����̌��ې��̎擾
		if not f_SetKekka(m_Rs_Student(0),0,C_KEKKA,C_KOUKI,m_AryKei(2,w_num)) then exit function
		
		'����̒x������ސ��̎擾
		if not f_SetKekka(m_Rs_Student(0),0,C_TIKOKU,C_KOUKI,m_AryKei(3,w_num)) then exit function
		
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
function f_SetKekka(p_GakusekiNo,p_Jigen,p_Type,p_Kikan,p_KekkaNum)
	
	Dim w_sSQL
	Dim w_Rs
	
	On Error Resume Next
	Err.Clear
	
	f_SetKekka = false
	
	p_KekkaNum = 0
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  DECODE(sum(T21_JIKANSU),0,'',sum(T21_JIKANSU)) "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T21_SYUKKETU "
	w_sSQL = w_sSQL & " where T21_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and T21_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T21_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & "  and T21_GAKUSEKI_NO ='" & p_GakusekiNo & "'"
	
	select case p_Kikan
		case C_OTHER	'����(�O�������ȊO)
			w_sSQL = w_sSQL & "  and T21_JIGEN =" & p_Jigen
			w_sSQL = w_sSQL & "  and T21_HIDUKE >='" & m_sFromDate & "'"
			w_sSQL = w_sSQL & "  and T21_HIDUKE <='" & m_sToDate & "'"
			
		case C_ZENKI	'�O��
			w_sSQL = w_sSQL & "  and T21_HIDUKE >='" & m_sZenki_Start & "'"
			w_sSQL = w_sSQL & "  and T21_HIDUKE <'"  & m_sKouki_Start & "'"
			
		case C_KOUKI	'���
			w_sSQL = w_sSQL & "  and T21_HIDUKE >='" & m_sKouki_Start & "'"
			w_sSQL = w_sSQL & "  and T21_HIDUKE <='"  & m_sKouki_End & "'"
			
	end select
	
	if p_Type = C_KEKKA then
		'���ې�
		w_sSQL = w_sSQL & "  and (T21_SYUKKETU_KBN =" & C_KETU_KEKKA & " or T21_SYUKKETU_KBN = " & C_KETU_KEKKA_1 & ")"
	elseif p_Type = C_TIKOKU then
		'�x������ސ�
		w_sSQL = w_sSQL & "  and (T21_SYUKKETU_KBN =" & C_KETU_TIKOKU & " or T21_SYUKKETU_KBN = " & C_KETU_SOTAI & ")"
	end if
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	p_KekkaNum = w_Rs(0)
	
	f_SetKekka = true
	
end function 

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

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	dim w_str	'�\�����b�Z�[�W
	
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
		
		//�X�N���[����������
		parent.init();
		
		if(location.href.indexOf('#')==-1){
			document.frm.target = "topFrame";
			document.frm.action = "kks0111_middle.asp"
			document.frm.submit();
		}
		return;
	}
	
    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Touroku(){
		parent.frames["main"].f_Touroku();
		return;
    }
	
    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Cancel(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default.asp"
    }
	
	//************************************************************
    //  [�@�\]  �ڍ׃{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //	
    //************************************************************
    function f_Detail(pGAKUSEI_NO,pName){
		var PositionX,PositionY,w_position;
		
		url = "kks0111_detail.asp";
		url = url + "?GakusekiNo=" + pGAKUSEI_NO;
		url = url + "&FromDate=<%=m_sFromDate%>";
		url = url + "&ToDate=<%=m_sToDate%>";
		url = url + "&Nen=<%=m_sGakunenCd%>";
		url = url + "&Class=<%=m_sClassCd%>";
		
		w   = 800;
		h   = 600;
		
		PositionX = window.screen.availWidth  / 2 - w / 2;
		PositionY = window.screen.availHeight / 2 - h / 2;
		
		w_position = ",left=" + PositionX + ",top=" + PositionY;
		
		opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
		if (w > 0)
			opt = opt + ",width=" + w;
		if (h > 0)
			opt = opt + ",height=" + h;
		
		opt = opt + w_position;
		
		newWin = window.open(url,"detail_subwin", opt);
	}
	
	//************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Back(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default.asp"
    }
    
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    
		<!-- 2015.03.19 Upd Start kiyomoto-->
    	<!--<table width=800>-->
    	<table>
		<!-- 2015.03.19 Upd End kiyomoto-->
        <tr>
        	<td valign="top" align="center" nowrap>
        		<table class="hyo"  border="1">
        			
        			<%dim i%>
        			<%for i=0 to m_StudentCount%>
						<% Call gs_cellPtn(w_Class) %>
        					<tr>
					            <td class="<%=w_Class%>" align="center" width="100" height="27" nowrap><%=m_AryKekka(0,i)%></td>
					            <td class="<%=w_Class%>" align="left" width="100" height="27" nowrap><%=trim(m_AryKekka(1,i))%></td>
					            <td class="<%=w_Class%>" align="center" width="50" height="27" nowrap><input type="button" name="btnDetail" value="�ڍ�" onClick="javascript:f_Detail('<%=m_AryKekka(0,i)%>','<%=m_AryKekka(1,i)%>');"></td>
					            
					            <% Dim j%>
					            <% for j = 3 to m_AryXCount %>
					            	<td class="<%=w_Class%>" align="center" width="20" height="27" nowrap><%=gf_SetNull2String(m_AryKekka(j,i))%></td>
		            			<% next %>	
							</tr>
		            <%next%>
            	</table>
            </td>
            
            <td width="10" height="27" valign="top" nowrap><br></td>
            
            <td align="center" width="120" valign="top" nowrap>
				
				<table width="120" class="hyo" border="1">
		            <% w_Class = "" %>
		            
		            <% Dim w_kei_num %>
		            <% for w_kei_num=0 to m_StudentCount %>
		            	<% Call gs_cellPtn(w_Class) %>
						
			            <tr>
			            	<td class="<%=w_Class%>" align="center" width="30" height="27" nowrap><%=gf_SetNull2String(m_AryKei(0,w_kei_num))%></td>
				            <td class="<%=w_Class%>" align="center" width="30" height="27" nowrap><%=gf_SetNull2String(m_AryKei(1,w_kei_num))%></td>
				            <td class="<%=w_Class%>" align="center" width="30" height="27" nowrap><%=gf_SetNull2String(m_AryKei(2,w_kei_num))%></td>
				            <td class="<%=w_Class%>" align="center" width="30" height="27" nowrap><%=gf_SetNull2String(m_AryKei(3,w_kei_num))%></td>
			            </tr>
		            	
		            <% next %>
	            </table>
			</td>
            
        </tr>
        
        </table>
		
		<table>
			<tr>
				<td align="center" nowrap>
					<input class="button" type="button" onclick="javascript:f_Back();" value=" �߁@�� ">
				</td>
			</tr>
	    </table>
		
	<INPUT type="hidden" name="txtFromDate"	value = "<%=m_sFromDate%>">
	<INPUT type="hidden" name="txtToDate"	value = "<%=m_sToDate%>">
	<INPUT type="hidden" name="cboGakunenCd"	value = "<%=m_sGakunenCd%>">
	<INPUT type="hidden" name="cboClassCd"	value = "<%=m_sClassCd%>">
	
	<INPUT type="hidden" name="JigenCount"  value="<%=m_JigenCount%>">
	
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>