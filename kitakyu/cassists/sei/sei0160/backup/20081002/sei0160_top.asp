<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ������w���ѓo�^
' ��۸���ID : sei/sei0160/sei0160_top.asp
' �@      �\: ��y�[�W ������w���ѓo�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:
'           :
' ��      ��:
' ��      �n:
'           :
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ�������w���ѓo�^��ʂ�\��������
'-------------------------------------------------------------------------
' ��      ��: 2007/04/11 ��c
' �C      ��: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Dim  m_bErrFlg           '�װ�׸�
	
	Dim m_iNendo             '�N�x
	Dim m_sKyokanCd          '�����R�[�h

	Dim m_sGakunen		 '�w�N
	Dim m_sClass		 '�N���X
	Dim m_sBunruiCD		 '���ރR�[�h
	Dim m_sBunruiNM		 '���ޖ���
	Dim m_sTani		 '�P��

	Dim m_sClassWhere	'�N���X�擾Where������

	Dim m_bNoData		 ''���͉Ȗڂ��Ȃ��Ƃ�True
	Dim gRs
	
	
'///////////////////////////���C������/////////////////////////////
	
	Call Main()
	
'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()
	Dim w_iRet              '// �߂�l
    Dim w_sSQL              	'// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="������w���ѓo�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"
	
    On Error Resume Next
    Err.Clear
	
    m_bErrFlg = false
	
    	Do
		'//�ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If
		
		'//�l���擾
		call s_SetParam()
		

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
	        '//�N���X�R���{�Ɋւ���WHERE���쐬����
        	Call s_MakeClassWhere() 

		'//���O�C�������̒S��������w�Ȗڂ̎擾
		if not f_GetNintei() then Exit Do
		
		If gRs.EOF Then
			m_bNoData = True
		Else
			m_bNoData = False
		End If

		'// �y�[�W��\��
		Call showPage()
		
		m_bErrFlg = true
		Exit Do
	Loop
	
	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If not m_bErrFlg Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// �I������
	Call gf_closeObject(gRs)
	Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	
    m_iNendo    = session("NENDO")		'�N�x
    m_sKyokanCd = session("KYOKAN_CD")		'�����R�[�h

    m_sGakunen  = request("txtGakunen")         '�w�N
    m_sClass    = request("txtClass")           '�N���X
    m_sBunruiCD = request("txtBunruiCd")	'���ރR�[�h
    m_sBunruiNm = request("txtBunruiNm")	'���ޖ���
    m_sTani     = request("txtTani")		'�P��

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeClassWhere()
    
    m_sClassWhere = "" 

    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iNendo  		   '//�����N�x

    if m_sGakunen <> "@@@" then
    	m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_sGakunen)    '//�w�N
    end if

'response.write " m_sClassWhere=" & m_sClassWhere & "<BR>" 

End Sub

'********************************************************************************
'*  [�@�\]  ���O�C�������̔F��Ȗڂ��擾(�N�x�A����CD���)
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetNintei()
	Dim w_sSQL
    Dim w_sJiki
    Dim w_Rs
    Dim w_sMinNendo 
    
    On Error Resume Next
    Err.Clear
	
    f_GetNintei = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	 M110_BUNRUI_CD     AS BUNRUI_CD "
	w_sSQL = w_sSQL & " 	,M110_BUNRUI_MEISYO AS BUNRUI_NM "
	w_sSQL = w_sSQL & " 	,M110_TANI          AS TANI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & "     M110_NINTEI_H "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "		M110_NENDO =" & m_iNendo
	w_sSQL = w_sSQL & "	AND	M110_KYOKAN_CD ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & " ORDER BY M110_BUNRUI_CD "
'response.write w_sSQL
'response.end		
	If gf_GetRecordset(gRs,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit function
	End If
	
	f_GetNintei = true
    
End Function

'********************************************************************************
'*  HTML���o��
'********************************************************************************
Sub showPage()
	Dim w_TukuName
	Dim w_SubjectDisp
	Dim w_SubjectValue
	
	On Error Resume Next
    Err.Clear
	
%>
	<html>
	<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//  [�@�\]  �������ύX���ꂽ�Ƃ��A�ĕ\������
	//************************************************************
	function f_ReLoadMyPage(){
		// �I�����ꂽ�R���{�̒l���
		f_SetData();
		document.frm.txtClass.value="";

		document.frm.action="sei0160_top.asp";
		document.frm.target="topFrame";
		document.frm.submit();
	}
	
	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���̏���
	//************************************************************
	function f_Search(){

		//���̓`�F�b�N
		if(!f_InpCheck()){
			return false;
		}

		// �I�����ꂽ�R���{�̒l���
		f_SetData();

		document.frm.action="sei0160_bottom.asp";
		document.frm.target="main";
		document.frm.submit();
	}
	//************************************************
	//	���̓`�F�b�N
	//************************************************
	function f_InpCheck(){
		var w_length;
		var ob;

		//�w�N
		ob = eval("document.frm.txtGakunen");
		if(ob.value =="@@@"){
			alert("�w�N��I�����Ă�������");
			ob.focus();
			ob.select();
			return false;
		}

		//�N���X
		ob = eval("document.frm.txtClass");
		if(ob.value =="@@@"){
			alert("�N���X��I�����Ă�������");
			ob.focus();
			ob.select();
			return false;
		}

		//�N���X
		ob = eval("document.frm.sltSubject");
		if(ob.value =="@@@"){
			alert("�Ȗڂ�I�����Ă�������");
			ob.focus();
			ob.select();
			return false;
		}

		return true;
	}
	
	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���ɑI�����ꂽ�f�[�^���
	//************************************************************
	function f_SetData(){
		//�f�[�^�擾
		var vl = document.frm.sltSubject.value.split('#@#');
		
		//�I�����ꂽ�f�[�^���(����CD�A���ޖ��́A�P�ʂ��擾)
		document.frm.txtBunruiCd.value=vl[0];
		document.frm.txtBunruiNm.value=vl[1];
		document.frm.txtTani.value=vl[2];
	}
	
	
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>
	
    	<body LANGUAGE="javascript">
	
	<center>
	<form name="frm" METHOD="post">
	
	<% call gs_title(" ������w���ѓo�^ "," �o�@�^ ") %>
	<br>
	
	<table border="0">
		<tr><td valign="bottom">
			
			<table border="0" width="100%">
				<tr><td class="search">
					
					<table border="0">
						<tr valign="middle">
							<td align="left" nowrap>�w�@�N</td>
							<td>
								<select name="txtGakunen" style="width:110px;" onchange ="javascript:f_ReLoadMyPage()" >
									<option value="@@@" selected >  </option>
									<% For I = 1 To 2 
										if cstr(I) = cstr(m_sGakunen) then %>
									<option value="<%=I%>" selected > <%=I%>�N</option>
										<% Else %>
									<option value="<%=I%>"> <%=I%>�N</option>
										<% end if 
								           Next %>
								</select>
							</td>

							<td>&nbsp;</td>
							
							<td align="left" nowrap>�N �� �X</td>
							
							<!-- '�w�N���I������Ă��Ȃ��ꍇ�́A���͕s�ɂ��� -->
							<td>
								<%IF m_sGakunen <> "@@@" and m_sGakunen <> "" then 
								 	call gf_ComboSet("txtClass",C_CBO_M05_CLASS,m_sClassWhere," style='width:200px;'",True,m_sClass) 
							 	else %>
								<select name="txtClass" DISABLED style="width:200px;">
									<option value="@@@">�@�@�@�@�@�@�@</option>
								</select>
							    <% end if %>
							</td>
	                    			</tr>
						
						<tr>
							<td align="left" nowrap>�Ȗ�</td>
							<td align="left">
								<% if not gRs.EOF then %>
								<select name="sltSubject" style="width:250px;">
									<% 
									do until gRs.EOF
										
										'�ȖڃR���{�\����������
										w_SubjectDisp = gf_SetNull2String(gRs("BUNRUI_NM")) 

										'�ȖڃR���{VALUE��������
										w_SubjectValue = ""
										w_SubjectValue = w_SubjectValue & gRs("BUNRUI_CD")   & "#@#"
										w_SubjectValue = w_SubjectValue & gRs("BUNRUI_NM")   & "#@#"
										w_SubjectValue = w_SubjectValue & gRs("TANI")  

								
										if cstr(m_sBunruiCD) = gf_SetNull2String(gRs("BUNRUI_CD")) then %>
									<option value="<%=w_SubjectValue%>" selected > <%=w_SubjectDisp%></option>
										<% Else %>
									<option value="<%=w_SubjectValue%>"> <%=w_SubjectDisp%></option>
										<% end if 
										gRs.movenext
									loop 
									%>
								</select>
								<% 	'' eof �̂Ƃ��̃��b�Z�[�W��ǉ�
								   else %>
									�F�S�������̓��͉Ȗڂ͂���܂���
								<!--select name="sltSubject">
									<option value="@@@">�@�@�@�@�@�@�@</option>
								</select-->
								<% end if %>
							</td>

							
							<td colspan="7" align="right">
							<% 'EOF �̂Ƃ��̏�����ǉ�
								If Not m_bNoData Then %>
								<input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
							<% 	Else %>
							<% 	end if %>
							</td>
						</tr>
					</table>
					
				</td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
	
	<input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtGakuNo"    value="<%=w_sGakunen%>">
	<input type="hidden" name="txtClassNo"   value="<%=m_sClass%>">
	<input type="hidden" name="txtBunruiCd"  value="<%=m_sBunruiCD%>">
	<input type="hidden" name="txtBunruiNm"  value="<%=m_sBunruiNM%>">
	<input type="hidden" name="txtTani"      value="<%=m_sTaniD%>">
	</form>
	</center>
	</body>
	</html>
<%
End Sub

'********************************************************************************
'*	��HTML���o��
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
	<html>
	<head>
	<title>������w���ѓo�^</title>
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
%>