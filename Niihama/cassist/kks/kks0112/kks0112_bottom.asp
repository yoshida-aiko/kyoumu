<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0112/kks0112_bottom.asp
' �@	  �\: ���y�[�W ���Əo�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��	  ��: 
'			  
'			  
'			  
'			  
' ��	  ��: 
' ��	  �n: 
'			  
'			  
'			  
'			  
' ��	  ��:
'			�������\��
'				���������ɂ��Ȃ����k�ꗗ��\��
'			���o�^�{�^���N���b�N��
'				���͏���o�^����
'			���߂�{�^���N���b�N��
'				�O�y�[�W�ɖ߂�
'-------------------------------------------------------------------------
' ��	  ��: 2002/05/16 shin
' ��	  �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public	m_bErrFlg			'�װ�׸�

	'�擾�����f�[�^�����ϐ�
	Public m_iSyoriNen		'//�����N�x
	Public m_sGakunen		'//�w�N
	Public m_sClassNo		'//�׽NO
	
	Public m_sKamokuCd		'//�ۖ�CD
	Public m_iKamokuKbn		'//���Ǝ��(C_JIK_JUGYO)
	
	'ں��ރZ�b�g
	Public m_Rs_Student		'//recordset���k
	
	Public m_sDate
	Public m_iJigen
	
	Public m_Count
	
	Const C_SELECT = 1	'//�I���ȖڂőI��
	
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
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'//�ϐ�������
		Call s_ClearParam()
		
		'// ���Ұ�SET
		Call s_SetParam()
		
		'//���k���擾
		if not f_Get_Student() Then
			m_bErrFlg = True
			Exit Do
		End If
		
		if m_Rs_Student.EOF Then
			Call showWhitePage("�ΏۂƂȂ�A���k��񂪂���܂���")
			Exit Do
		End If
		
		'// �f�[�^�\���y�[�W��\��
		Call showPage_bottom()
		
		Exit Do
	Loop
	
	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
	End If
	
	'// �I������
	Call gf_closeObject(m_Rs_Student)
	
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[�@�\]	�ϐ�������
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_ClearParam()

	m_iSyoriNen = 0
	
	m_sGakunen	= 0
	m_sClassNo	= 0
	
	m_sKamokuCd = ""
	m_iKamokuKbn	= 0
	
	m_sDate = ""
	m_iJigen = ""
	
End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen = Session("NENDO")
	
	m_iKamokuKbn	= cint((Request("hidSyubetu")))
	
	m_sGakunen	= trim(Request("hidGakunen"))
	m_sClassNo	= trim(Request("hidClassNo"))
	m_sKamokuCd = trim(Request("hidKamokuCd"))
	
	m_sDate = gf_YYYY_MM_DD(trim(Request("txtDate")),"/")
	m_iJigen = trim(Request("sltJigen"))
	
End Sub

'********************************************************************************
'*	[�@�\]	���k���擾,���ۥ�x�������̎擾
'*	[����]	
'*	[�ߒl]	true:���擾���� false:���s
'*	[����]	
'********************************************************************************
function f_Get_Student()
	Dim w_sSQL
	
	On Error Resume Next
	Err.Clear
	
	f_Get_Student = false
	
	if m_iKamokuKbn = C_JIK_JUGYO then
		'//�ʏ���Ƃ̂Ƃ�
		w_sSQL = ""
		w_sSQL = w_sSQL & " select "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & " 	T11.T11_SIMEI,"
		w_sSQL = w_sSQL & " 	T21.T21_JIKANSU, "
		w_sSQL = w_sSQL & " 	T21.T21_SYUKKETU_KBN, "
		w_sSQL = w_sSQL & " 	T16.T16_HISSEN_KBN as HISSEN_KBN , "
		w_sSQL = w_sSQL & " 	T16.T16_SELECT_FLG as SELECT_FLG "
		
		w_sSQL = w_sSQL & " from "
		w_sSQL = w_sSQL & " 	T13_GAKU_NEN T13, "
		w_sSQL = w_sSQL & " 	T11_GAKUSEKI T11, "
		w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN T16 ,"
		w_sSQL = w_sSQL & " 	( "
		w_sSQL = w_sSQL & " 	 select * from T21_SYUKKETU "
		w_sSQL = w_sSQL & " 	 where T21_HIDUKE ='" & m_sDate & "' "
		w_sSQL = w_sSQL & " 	 and T21_JIGEN =" & m_iJigen & " ) T21 "
		
		w_sSQL = w_sSQL & " where "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEI_NO = T11.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " and T13.T13_GAKUSEI_NO = T16.T16_GAKUSEI_NO "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = T21.T21_NENDO(+) "
		w_sSQL = w_sSQL & " and T13.T13_GAKUNEN = T21.T21_GAKUNEN(+) "
		w_sSQL = w_sSQL & " and T13.T13_CLASS = T21.T21_CLASS(+) "
		w_sSQL = w_sSQL & " and T13.T13_GAKUSEKI_NO = T21.T21_GAKUSEKI_NO(+) "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = T16.T16_NENDO "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & " and T13.T13_GAKUNEN = " & m_sGakunen
		w_sSQL = w_sSQL & " and T13.T13_CLASS = " & m_sClassNo
		w_sSQL = w_sSQL & " and T16.T16_KAMOKU_CD ='" & m_sKamokuCd & "'"
		
		w_sSQL = w_sSQL & " group by "
		w_sSQL = w_sSQL & " 	T11.T11_SIMEI,"
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO,"
		w_sSQL = w_sSQL & " 	T21.T21_JIKANSU,"
		w_sSQL = w_sSQL & " 	T21.T21_SYUKKETU_KBN,"
		w_sSQL = w_sSQL & " 	T16.T16_HISSEN_KBN,"
		w_sSQL = w_sSQL & " 	T16.T16_SELECT_FLG "
		w_sSQL = w_sSQL & " order by "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO "
	else
		'//���ʊ���
		w_sSQL = ""
		w_sSQL = w_sSQL & " select "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & " 	T11.T11_SIMEI,"
		w_sSQL = w_sSQL & " 	T21.T21_JIKANSU, "
		w_sSQL = w_sSQL & " 	T21.T21_SYUKKETU_KBN, "
		w_sSQL = w_sSQL & " 	1 as HISSEN_KBN , "
		w_sSQL = w_sSQL & " 	0 as SELECT_FLG "
		
		w_sSQL = w_sSQL & " from "
		w_sSQL = w_sSQL & " 	T13_GAKU_NEN T13, "
		w_sSQL = w_sSQL & " 	T11_GAKUSEKI T11, "
		w_sSQL = w_sSQL & " 	T34_RISYU_TOKU T34 ,"
		w_sSQL = w_sSQL & " 	( "
		w_sSQL = w_sSQL & " 	 select * from T21_SYUKKETU "
		w_sSQL = w_sSQL & " 	 where T21_HIDUKE ='" & m_sDate & "' "
		w_sSQL = w_sSQL & " 	 and T21_JIGEN =" & m_iJigen & " ) T21 "
		
		w_sSQL = w_sSQL & " where "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEI_NO = T11.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " and T13.T13_GAKUSEI_NO = T34.T34_GAKUSEI_NO "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = T21.T21_NENDO(+) "
		w_sSQL = w_sSQL & " and T13.T13_GAKUNEN = T21.T21_GAKUNEN(+) "
		w_sSQL = w_sSQL & " and T13.T13_CLASS = T21.T21_CLASS(+) "
		w_sSQL = w_sSQL & " and T13.T13_GAKUSEKI_NO = T21.T21_GAKUSEKI_NO(+) "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = T34.T34_NENDO "
		
		w_sSQL = w_sSQL & " and T13.T13_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & " and T13.T13_GAKUNEN = " & m_sGakunen
		w_sSQL = w_sSQL & " and T13.T13_CLASS = " & m_sClassNo
		w_sSQL = w_sSQL & " and T34.T34_TOKUKATU_CD ='" & m_sKamokuCd & "'"
		
		w_sSQL = w_sSQL & " group by "
		w_sSQL = w_sSQL & " 	T11.T11_SIMEI,"
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO,"
		w_sSQL = w_sSQL & " 	T21.T21_JIKANSU,"
		w_sSQL = w_sSQL & " 	T21.T21_SYUKKETU_KBN "
		w_sSQL = w_sSQL & " order by "
		w_sSQL = w_sSQL & " 	T13.T13_GAKUSEKI_NO "
		
	end if
	
	If gf_GetRecordset(m_Rs_Student,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	f_Get_Student = true
	
end function

'********************************************************************************
'*	[�@�\]	�o�����̎擾
'*	[����]	p_SyukketuKbn:�o���敪
'*	[�ߒl]	�o������
'*	[����]	
'********************************************************************************
function f_Set_SyukketuMei(p_SyukketuKbn,p_JikanNum)
	Dim w_num
	Dim w_KubunName
	
	if gf_SetNull2String(p_SyukketuKbn) = "" then 
		f_Set_SyukketuMei = ""
		exit function
	end if
	
	'//�o�ȋ敪���擾
	if not gf_GetKubunName(19,p_SyukketuKbn,m_iSyoriNen,w_KubunName) then
		f_Set_SyukketuMei = ""
		exit function
	end if
	
	if cint(p_SyukketuKbn) = 1 then
		f_Set_SyukketuMei = p_JikanNum & w_KubunName
	else
		f_Set_SyukketuMei = w_KubunName
	end if
	
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
	</head>
	
	<body LANGUAGE="javascript">
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
'*	[�@�\]	HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub showPage_bottom()
	Dim w_Class
	Dim w_num
	Dim w_SyukketuName
	Dim w_SelectFlg
	Dim w_HissenFlg
	Dim w_IdouType,w_KubunName
	
	w_num = 1
	
	On Error Resume Next
	Err.Clear
	
%>
	<html>
	<head>
	<title>�s���p�o������</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<!--#include file="../../Common/jsCommon.htm"-->
	
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	
	//************************************************************
	//	[�@�\]	�y�[�W���[�h������
	//	[����]
	//	[�ߒl]
	//	[����]
	//************************************************************
	function window_onload() {
		//�X�N���[����������
		parent.init();
		
		if(location.href.indexOf('#')==-1){
			//�w�b�_����\��submit
			document.frm.target = "topFrame";
			document.frm.action = "kks0112_middle.asp?<%=Request.Form.Item%>"
			document.frm.submit();
		}
		
		return;
		
	}
	
	//************************************************************
	//	[�@�\]	�o�^�{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Insert(){
		if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
		   return false;
		}
		
		//�w�b�_���󔒕\��
		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("�o�^���Ă��܂��E�E�E�E�@�@���΂炭���҂���������")%>"
		
		//���X�g����submit
		document.frm.target = "main";
		document.frm.action = "kks0112_edit.asp";
		document.frm.submit();
	}
	
    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
    //  [����]  
    //  [�ߒl]  
    //  [����]
    //************************************************************
    function f_Back(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default.asp";
    }
	
	//************************************************************
	//	[�@�\]	�o������
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//************************************************************
	function chg(chgInp){
		var w_num,i,w_StateValue,w_KekkaNum;
		var w_txtKekka;
		w_txtKekka = eval("parent.topFrame.document.frm.txtKekka");
		
		w_num = 0;
		w_KekkaNum = w_txtKekka.value;
		
		//�ǂ̃��W�I�{�^�����I������Ă��邩�`�F�b�N
		for(i=0;i<4;i++){
			if(parent.topFrame.document.frm.rdoType[i].checked == true){
				w_num = i + 1;
				if(w_num == 1 && w_KekkaNum == ""){
					f_InpChkErr("���ې�����͂��ĉ�����",w_txtKekka);
					return false;
				}else if(w_num == 1 && (w_KekkaNum < 1 || w_KekkaNum > 9 || isNaN(w_KekkaNum))){
					f_InpChkErr("���ې����s���ł�",w_txtKekka);
					return false;
				}
				break;
			}
		}
		
		switch(w_num){
			case 0 : alert("���͂������u���͋敪�v��I����A�Y������w���̏o���󋵗����N���b�N���ĉ������B");
					 return false;
					 break;
					 
			case 1 : w_StateValue = w_KekkaNum + "����";
					 break;
					 
			case 2 : w_StateValue = "�x��";
					 w_KekkaNum = 1;
					 break;
					 
			case 3 : w_StateValue = "����";
					 w_KekkaNum = 1;
					 break;
					 
			case 4 : w_StateValue = ""; 
					 w_KekkaNum = 0;
					 w_num = 0;
					 break;
					 
			default: break;
		}
		
		chgInp.value = w_StateValue;
		
		var ob = new Array();
		ob[0] = eval("document.frm.hid"+chgInp.name);
		ob[0].value = w_num;
		
		ob[1] = eval("document.frm.hidJikan"+chgInp.name);
		ob[1].value = w_KekkaNum;
		
	}
	
	//************************************************************
    //  [�@�\]  ���̓`�F�b�N�G���[����alert,focus,select����
    //************************************************************
    function f_InpChkErr(p_AlertMsg,p_Object){
		alert(p_AlertMsg);
		p_Object.focus();
		p_Object.select();
	}
	
	//-->
	</SCRIPT>
	
	</head>
	<body LANGUAGE="javascript" onload="window_onload()">
	<form name="frm" method="post">
	
	<center>
		<table width="545">
			<tr>
				<td align="center" valign="top" nowrap>
					<table class="hyo"	border="1" width="300">
						
						<%
						Do until m_Rs_Student.EOF
							Call gs_cellPtn(w_Class)
							
							w_HissenFlg = cint(m_Rs_Student("HISSEN_KBN"))	'�K�C�E�I���t���O
							w_SelectFlg = cint(m_Rs_Student("SELECT_FLG"))	'�I�����Ƃ�I�����Ă��邩�t���O
							
							if (w_HissenFlg = C_HISSEN_HIS) or (w_HissenFlg = C_HISSEN_SEN and w_SelectFlg = C_SELECT) then
						%>
								<tr>
									<td class="<%=w_Class%>" width="80" align="center" nowrap><%=m_Rs_Student("T13_GAKUSEKI_NO")%></td>
									<input type="hidden" name="hidGakusekiNo" value="<%=m_Rs_Student("T13_GAKUSEKI_NO")%>">
									
									<td class="<%=w_Class%>" width="150" nowrap><%=m_Rs_Student("T11_SIMEI")%></td>
									<td class="<%=w_Class%>" width="70" height="28" align="center" nowrap>
										<% 
											'�o����
											w_SyukketuName = ""
											w_SyukketuName = f_Set_SyukketuMei(m_Rs_Student("T21_SYUKKETU_KBN"),m_Rs_Student("T21_JIKANSU"))
											
											'�ٓ����擾
											w_IdouType = cint(gf_SetNull2Zero(gf_Get_IdouChk(m_Rs_Student("T13_GAKUSEKI_NO"),m_sDate,m_iSyoriNen,w_KubunName)))
										%>
										
										<% if w_IdouType = 0 or w_IdouType = C_IDO_FUKUGAKU or w_IdouType = C_IDO_TEI_KAIJO then %>
											<input type="button" class="<%=w_Class%>" name="State<%=m_Rs_Student("T13_GAKUSEKI_NO")%>" style="border-style:none;text-align:center;" tabindex="-1" value="<%=w_SyukketuName%>" onclick="return chg(this);">
										<% else %>
											<font color="red"><%=w_KubunName%></font>
										<% end if %>
										
										<input type="hidden" name='hidState<%=m_Rs_Student("T13_GAKUSEKI_NO")%>' value='<%=gf_SetNull2Zero(m_Rs_Student("T21_SYUKKETU_KBN"))%>'>
										<input type="hidden" name='hidJikanState<%=m_Rs_Student("T13_GAKUSEKI_NO")%>' value='<%=gf_SetNull2Zero(m_Rs_Student("T21_JIKANSU"))%>'>
									</td>
								</tr>
						<%	
							end if
							
							w_num = w_num + 1
							m_Rs_Student.movenext
						loop
						
						%>
						
					</table>
				</td>
			</tr>
			
			<tr>
				<td align="center" valign="top" nowrap>
					<table>
						<tr>
							<td nowrap><input type="button" name="btnInsert" value="�@�o�@�^�@" onClick="f_Insert();"></td>
							<td nowrap><input type="button" name="btnBack" value="�@�߁@��@" onClick="f_Back();"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</center>
	
	<input type="hidden" name="hidGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="hidClassNo" value="<%=m_sClassNo%>">
	<input type="hidden" name="hidKamokuCd" value="<%=m_sKamokuCd%>">
	
	<input type="hidden" name="hidDate" value="<%=m_sDate%>">
	<input type="hidden" name="hidJigen" value="<%=m_iJigen%>">
	
	<input type="hidden" name="hidKamokuName" value="<%=request("hidKamokuName")%>">
	<input type="hidden" name="hidClassName" value="<%=request("hidClassName")%>">
	<input type="hidden" name="hidSyubetu" value="<%=m_iKamokuKbn%>">
	
	</form>
	</body>
	</html>
<%
End Sub
%>
