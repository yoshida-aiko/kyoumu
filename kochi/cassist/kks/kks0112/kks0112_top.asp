<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0112/kks0112_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
'            
'            
' ��      ��:
'           �������\��
'               �ȖڃR���{�{�b�N�X�F���O�C���҂̒S���Ȗ�
'               [���ۓ���]
'					���Ɠ��F�V�X�e�����t
'					����  �F�����}�X�^���擾
'               [���ۈꗗ�Q��]
'					�w�茎�F
'           ���I���{�^���N���b�N��
'               �J�����_�[���o��
'           �����̓{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ����Ƃ̏o�����͉�ʂ�\��
'           ���\���{�^���N���b�N��
'               �T�u�E�B���h�E�Ŏw�肵�������̏o���󋵂�\��
'-------------------------------------------------------------------------
' ��      ��: 2002/05/16 shin
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    
    Public m_iSyoriNen          '//�N�x
    Public m_iKyokanCd          '//��������
    
    Public m_sGakki             '//�w��
    Public m_sZenki_Start		'//�O���J�n��
    Public m_sKouki_Start		'//����J�n��
    Public m_sKouki_End			'//����I����
    
	Public m_Rs_Jigen			'//����
	Public m_Rs_Subject			'//�Ȗ�
	
	Public m_JigenCount			'//������
    
    Public m_Month				'//���݂̌�
    '�G���[�n
    Public m_bErrFlg           '�װ�׸�
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
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If
		
        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))
		
        '//�l�̏�����
		Call s_ClearParam()
		
        '//�ϐ��Z�b�g
        Call s_SetParam()
		
		'//�O���E��������擾
		if gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End) <> 0 then
			m_bErrFlg = True
        	Exit Do
		end if
		
		'//���O�C�������̒S���Ȗڂ̎擾
		if not f_GetSubject() then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//���ƃf�[�^���擾�ł��Ȃ��Ƃ�
		if m_Rs_Subject.EOF then
			Call showWhitePage("���ƃf�[�^������܂���")
			Exit Do
		end if
		
		'//�������̎擾
		if not f_Get_JigenData() then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//�������擾�ł��Ȃ��Ƃ�
		if m_Rs_Jigen.EOF then
			Call showWhitePage("���������擾�ł��܂���")
			Exit Do
		end if
		
        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop
	
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
    End If
    
    '// �I������
    Call gf_closeObject(m_Rs_Jigen)
	Call gf_closeObject(m_Rs_Subject)
	
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
    m_iKyokanCd = ""
    m_Month = 0
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")
	m_Month = month(date())
End Sub

'********************************************************************************
'*  [�@�\]  ���O�C�������̎󎝋��Ȃ��擾(�N�x�A����CD�A�w�����)
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSubject()
	Dim w_sSQL
    
    On Error Resume Next
    Err.Clear
	
    f_GetSubject = false
	
	'�ʏ�A���w����։Ȗڎ擾
	w_sSQL = ""
	w_sSQL = w_sSQL & "select "
	w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & "from"
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & "		,M03_KAMOKU "
	w_sSQL = w_sSQL & "		,M100_KAMOKU_ZOKUSEI "
	w_sSQL = w_sSQL & "where "
	w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_iKyokanCd & "'"
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M03_KAMOKU_CD "
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI = " & C_JIK_JUGYO
	
	w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	M03_ZOKUSEI_CD = M100_ZOKUSEI_CD "
	
	w_sSQL = w_sSQL & "	and	M100_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	M100_SYUKKETSU_FLG = 0 "
	
	w_sSQL = w_sSQL & "union "
	
	'���ʊ����擾
	w_sSQL = w_sSQL & "select "
	w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M41_MEISYO as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & "from "
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & "		,M41_TOKUKATU "
	w_sSQL = w_sSQL & "		,M100_KAMOKU_ZOKUSEI "
	w_sSQL = w_sSQL & "where "
	w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_iKyokanCd & "'"
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M41_TOKUKATU_CD "
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI = " & C_JIK_TOKUBETU
	
	w_sSQL = w_sSQL & "	and	M41_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	M41_ZOKUSEI_CD = M100_ZOKUSEI_CD "
	
	w_sSQL = w_sSQL & "	and	M100_NENDO =" & cint(m_iSyoriNen)
	w_sSQL = w_sSQL & "	and	M100_SYUKKETSU_FLG = 0 "
	
	w_sSQL = w_sSQL & "order by GAKUNEN,CLASS,KAMOKU_KBN "
	
	If gf_GetRecordset(m_Rs_Subject,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit function
	End If
	
	f_GetSubject = true
    
End Function

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
		exit function
	End If
	
	m_JigenCount = cInt(m_Rs_Jigen(0))
	
	f_Get_JigenData = true
	
end function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	Dim w_num
	Dim w_ClassName
%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>���Əo������</title>
	
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		if(location.href.indexOf('#')==-1){
			//�w�b�_����\��submit
			document.frm.target = "main";
			document.frm.action = "white.asp?data_flg=OK"
			document.frm.submit();
		}
    }
    
    //************************************************************
    //  [�@�\]  ���̓{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Insert(){
		if(!f_InpChk()){ return false; }
		
		f_SetHidden();
		
    	document.frm.action="WaitAction.asp";
        document.frm.target="main";
        document.frm.submit();
    }
    //************************************************************
    //  [�@�\]  ���̓`�F�b�N
    //  [����]  
    //  [�ߒl]  
    //  [����]
    //
    //************************************************************
    function f_InpChk(){
		var obj = eval("document.frm.txtDate");
		
		//���J�n��
        //NULL�`�F�b�N
        if(f_Trim(obj.value) == ""){
            f_InpChkErr("���Ɠ������͂���Ă��܂���",obj);
            return false;
        }
        
        //�^�`�F�b�N
        if(IsDate(obj.value) != 0){
        	f_InpChkErr("���Ɠ��̓��t���s���ł�",obj);
        	return false;
        }
        
        //�O���J�n��<=���Ɠ�<=����I�����̃`�F�b�N
        if(DateParse("<%=m_sZenki_Start%>",obj.value) < 0 || DateParse(obj.value,"<%=m_sKouki_End%>") < 0){
			f_InpChkErr("���Ɠ��ɂ́A�O���J�n���Ȍ�A����I�����ȑO�̓��t����͂��Ă�������",obj);
			return false;
		}
		
        return true;
		
	}
	
	//************************************************************
    //  [�@�\]  ���̓`�F�b�N�G���[����alert,focus,select����
    //************************************************************
    function f_InpChkErr(p_AlertMsg,p_Object){
		alert(p_AlertMsg);
		p_Object.focus();
		p_Object.select();
	}
	
    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){
		var PositionX,PositionY,w_position;
		
		var vl = document.frm.sltKamoku.value.split('#@#');
		
		url = "kks0112_subwin.asp";
		url = url + "?hidSyubetu=" + vl[0];
		url = url + "&hidKamokuCd=" + vl[1];
		url = url + "&hidGakunen=" + vl[2];
		url = url + "&hidClassNo=" + vl[3];
		url = url + "&sltMonth=" + document.frm.sltMonth.value;
		
		w   = window.screen.availWidth;
		h   = window.screen.availHeight-30;
		
		PositionX = window.screen.availWidth  / 2 - w / 2;
		PositionY = 0; //window.screen.availHeight / 2 - h / 2;
		
		w_position = ",left=" + PositionX + ",top=" + PositionY;
		
		opt = "directoris=0,location=0,menubar=0,status=0,toolbar=0,resizable=no";
		opt = opt + ",width=" + w + ",height=" + h;
		opt = opt + w_position;
		
		nWin = window.open(url,"kks0112_subwin",opt);
	}
	
    //************************************************************
    //  [�@�\]  �R���{�̔N�A�N���X�A�Ȗڂ��΂炵�ăZ�b�g����
    //************************************************************
    function f_SetHidden(){
		var vl = document.frm.sltKamoku.value.split('#@#');
		
		//�ʏ�E���ʎ���(��ʁA�ۖں��ށA�w�N�A�׽NO���擾)
		document.frm.hidSyubetu.value = vl[0];
        document.frm.hidKamokuCd.value = vl[1];
        document.frm.hidGakunen.value = vl[2];
        document.frm.hidClassNo.value = vl[3];
		
		document.frm.hidClassName.value = vl[4];
        document.frm.hidKamokuName.value = vl[5];
	}
	//************************************************************
    //  [�@�\]  �G���^�[�L�[����
    //************************************************************
	function f_EnterClick(p_Type){
		if(event.keyCode==13){
			if(p_Type == "INSERT"){
				f_Insert();
			}else{
				f_Search();
			}
		}
	}
	
	//-->
    </SCRIPT>
	
	</head>
	<body LANGUAGE="javascript" onload="return window_onload();">
	<% call gs_title("���Əo������","���@��") %>
	<form name="frm" method="post">
	<center>
	<table border="0">
		<tr>
			<td align="right" class="search" nowrap>
				<table border="0">
					<tr>
						<td nowrap>�Ȗ�</td>
						<td colspan="7" nowrap>
							<select name="sltKamoku" style="width:200px;">
								<% 
									do until m_Rs_Subject.EOF
										w_ClassName = ""
										w_ClassName = gf_GetClassName(m_iSyoriNen,m_Rs_Subject("GAKUNEN"),m_Rs_Subject("CLASS"))
										
								%>
										<option value="<%=CStr(cint(m_Rs_Subject("KAMOKU_KBN")) & "#@#" & m_Rs_Subject("KAMOKU_CD") & "#@#" & m_Rs_Subject("GAKUNEN") & "#@#" & m_Rs_Subject("CLASS") & "#@#" & w_ClassName & "#@#" & m_Rs_Subject("KAMOKU_NAME"))%>"><%=m_Rs_Subject("GAKUNEN") & "�N&nbsp;&nbsp;" & w_ClassName & "&nbsp;&nbsp;&nbsp;" & m_Rs_Subject("KAMOKU_NAME") %>
								<%
										m_Rs_Subject.movenext
									loop
								%>
								
							</select>
						</td>
					</tr>
					
				    <tr><td colspan="7" height="10"><img src="../../image/sp_black.gif" width="100%" height="1"></td></tr>
				    
				    <tr>
				    	<th class="header" colspan="4" align="center">���ۓ���</td>
				    	
				    	<td rowspan="4"><img src="../../image/sp_black.gif" width="1" height="80"></td>
				    	
				    	<th class="header" colspan="4" align="center">���ۈꗗ�Q��</td>
				    </tr>
				    
				    <tr>
				    	<td>���Ɠ�</td>
				    	<td>
				    		<input type="text" name="txtDate" value="<%=gf_YYYY_MM_DD(date(),"/")%>" onKeyDown="f_EnterClick('INSERT');">
				    		<input type="button" class="button" onClick="fcalender('txtDate')" value="�I��">
				    	</td>
				    	
				    	<td>����</td>
				    	<td>
				    		<select name="sltJigen">
				    		
				    		<% for w_num=1 to m_JigenCount %>
				    			<option value="<%=w_num%>"><%=w_num%>
				    		<% next %>
				    		
				    	</td>
				    	
				    	<td>�w�茎</td>
				    	<td><select name="sltMonth"  onKeyDown="f_EnterClick('DISP');">
				    			<option value="4"  <%=gf_iif(m_Month = 4,"selected","")%>  >4
				    			<option value="5"  <%=gf_iif(m_Month = 5,"selected","")%>  >5
				    			<option value="6"  <%=gf_iif(m_Month = 6,"selected","")%>  >6
				    			<option value="7"  <%=gf_iif(m_Month = 7,"selected","")%>  >7
				    			<option value="8"  <%=gf_iif(m_Month = 8,"selected","")%>  >8
				    			<option value="9"  <%=gf_iif(m_Month = 9,"selected","")%>  >9
				    			<option value="10" <%=gf_iif(m_Month = 10,"selected","")%> >10
				    			<option value="11" <%=gf_iif(m_Month = 11,"selected","")%> >11
				    			<option value="12" <%=gf_iif(m_Month = 12,"selected","")%> >12
				    			<option value="1"  <%=gf_iif(m_Month = 1,"selected","")%>  >1
				    			<option value="2"  <%=gf_iif(m_Month = 2,"selected","")%>  >2
				    			<option value="3"  <%=gf_iif(m_Month = 3,"selected","")%>  >3
				    		</select>
				    	</td>
				    </tr>
				    
				    <tr>
				    	
				    </tr>
				    
				    <tr>
				    	<td colspan="4" align="center" nowrap>
							<input class="button" type="button" onclick="javascript:f_Insert();" value="�@���@�́@">
						</td>
				    	
						<td colspan="2" align="center" nowrap>
							<input class="button" type="button" onclick="javascript:f_Search();" value="�@�\�@���@">
						</td>
					</tr>
			    </table>
				
		    </td>
	    </tr>
    </table>
	
    <!--�l�n���p-->
    <input type="hidden" name="Tuki_Zenki_Start" value="<%=m_sZenki_Start%>">
    <input type="hidden" name="Tuki_Kouki_Start" value="<%=m_sKouki_Start%>">
    <input type="hidden" name="Tuki_Kouki_End"   value="<%=m_sKouki_End%>">
    
    <INPUT TYPE="hidden" name="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE="hidden" name="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
    
    <INPUT TYPE="hidden" name="hidGakunen"   value = "">
    <INPUT TYPE="hidden" name="hidClassNo"   value = "">
    <INPUT TYPE="hidden" name="hidKamokuCd" value = "">
    <INPUT TYPE="hidden" name="hidSyubetu"   value = "">
	
    <INPUT TYPE="hidden" name="hidClassName" value = "">
	<INPUT TYPE="hidden" name="hidKamokuName"   value = "">
	
    <input TYPE="hidden" name="txtURL" VALUE="kks0112_bottom.asp">
    <input TYPE="hidden" name="txtMsg" VALUE="���΂炭���҂���������">
	
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