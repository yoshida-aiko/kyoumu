<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�
' ��۸���ID : gak/gak0350_11/top.asp
' �@      �\: ��y�[�W �w�Ѓf�[�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�Ȃ�
' ��      �n:�����N�x       ��      SESSION���i�ۗ��j
'			txtMode				   :���샂�[�h
'           txtGakunen             :�w�N
'           txtGakkaCD             :�w��
'           txtClass               :�N���X
'           txtName                :����
'           txtGakusekiNo          :�w�Дԍ�
'           txtGakuseiNo           :�w���ԍ�
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X      �w�N��\��
'                                  �w�Ȃ�\��
'                                  �N���X��\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�����������ɂ��Ȃ��w������\��������
'-------------------------------------------------------------------------
' ��      ��: 2006/04/26 �F��
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg			   '�װ�׸�
    
    '�I��p��Where����
    Public s_sGakkaWhere		   '�w�Ȃ̒��o����
    Public m_sClassWhere		   '�N���X�̒��o����
       
    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen      	   ':�����N�x
    Public  m_TxtMode      	       ':���샂�[�h
    Public  m_sGakunen             ':�w�N
    Public  m_sGakkaCD             ':�w��
    Public  m_sClass               ':�N���X
    Public  m_sName                ':����
    Public  m_sGakusekiNo          ':�w�Дԍ�
    Public  m_sGakuseiNo           ':�w���ԍ�
	
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

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�w����񌟍�"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_TxtMode=request("txtMode")

    '// ���Ұ�SET
	if m_TxtMode = "" then
        	Call s_IntParam()
	else
        	Call s_SetParam()
	end if

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
        
        '�w�ȃR���{�Ɋւ���WHERE���쐬����
        Call s_MakeGakkaWhere() 
        
        '�N���X�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeClassWhere() 
        
        '// �y�[�W��\��
        Call showPage()
        
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [�@�\]  �p�����[�^������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_IntParam()

	m_iSyoriNen = cint(Session("Nendo"))	'�����N�x
	m_sGakunen=""            				'�w�N
    m_sGakkaCD=""             				'�w��
    m_sClass=""               				'�N���X
    m_sName=""                				'����
    m_sGakusekiNo=""          				'�w�Дԍ�
    m_sGakuseiNo=""           				'�w���ԍ�
 	
End Sub


'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen    = cint(Session("Nendo"))			 '�����N�x
    m_sGakunen     = request("txtGakunen")           '�w�N
    m_sGakkaCD     = request("txtGakka")             '�w��
    m_sClass       = request("txtClass")             '�N���X
    m_sName        = request("txtName")              '����
    m_sGakusekiNo  = request("txtGakusekiNo")        '�w�Дԍ�
    m_sGakuseiNo   = request("txtGakuseiNo")         '�w���ԍ�
 	
End Sub

'********************************************************************************
'*  [�@�\]  �w�ȃR���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeGakkaWhere()

    s_sGakkaWhere = ""
    s_sGakkaWhere = s_sGakkaWhere & " M02_NENDO = " & m_iSyoriNen & " AND "
    s_sGakkaWhere = s_sGakkaWhere & " M02_GAKKA_CD <> '00'"

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeClassWhere()
    
    m_sClassWhere = "" 
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyoriNen 		
    
    if m_sGakunen <> "@@@" then
        	m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_sGakunen)    
	end if

End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
%>

<html>

<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
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

    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_BackClick(){

        document.frm.action="../../menu/kensaku.asp";
        document.frm.target="_parent";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �������s�{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

		document.frm.action="./main.asp";
        document.frm.target="fMain";
        document.frm.txtMode.value = "Search";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �w�N, �N�x���I�����ꂽ�Ƃ��A�ĕ\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="top.asp";
        document.frm.target="fTop";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\] �N���A�{�^���������ꂽ�Ƃ�
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function jf_Clear(){
        document.frm.txtGakunen.value = "";
        document.frm.txtGakka.value = "";
        document.frm.txtClass.value = "";
        document.frm.txtMode.value = "";
        
    }
    //-->
    </SCRIPT>

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="center">
<%call gs_title("�w����񌟍�","��@��")%>
<form action="./main.asp" method="post" name="frm" target="fMain">

 	<input type="hidden" name="txtMode" width="100%" value="<%=m_TxtMode %>">

	<table cellspacing="0" cellpadding="0" border="0" width="100%" >
		<tr>
			<td valign="top" align="center">
					<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td class=search valign ="top">
								<table border="0" bgcolor="#E4E4ED" cellpadding="0" cellspacing="0">
									<tr>
										<td valign="top">
											<table border="0">
												<tr>
													<td nowrap height="16">�w�@�N
													</td>
													<td>
														<select name="txtGakunen" style="width:110px;" onchange ="javascript:f_ReLoadMyPage()" >
															<option value="@@@" selected >  </option>
															<% For I = 1 To 5 
																if cstr(I) = cstr(m_sGakunen) then %>
																	<option value="<%=I%>" selected > <%=I%>�N</option>
																<% Else %>
																	<option value="<%=I%>"> <%=I%>�N</option>
																<% end if 
														Next %>
														</select>
													</td>
													<td nowrap height="16">�w�@��
													</td>
													<td>
														<!-- 2015.03.20 Upd width:110->180 -->
														<% call gf_ComboSet("txtGakka",C_CBO_M02_GAKKA,s_sGakkaWhere," style='width:180px;'",True,m_sGakkaCD) %>
													</td>
													<td nowrap height="16">�N �� �X
													</td>
												
													<!-- '�w�N���I������Ă��Ȃ��ꍇ�́A���͕s�ɂ��� -->
													<td>
														<%IF m_sGakunen <> "@@@" and m_sGakunen <> "" then 
															<!-- 2015.03.20 Upd width:110->180 -->
								 							call gf_ComboSet("txtClass",C_CBO_M05_CLASS,m_sClassWhere," style='width:180px;'",True,m_sClass) 
							 							else %>
															<!-- 2015.03.20 Upd width:110->180 -->
															<select name="txtClass" DISABLED style="width:180px;" ID="Select1">
															<option value="@@@">�@�@�@�@�@�@�@</option>
															</select>
														<% end if %>
													</td>
													<td>
														<input type="button" class=button value="�N�@���@�A" onclick="jf_Clear()" ID="Button1" NAME="Button1">
														<input type="button" class=button value=" �\�@�� " onClick="javascript:f_Search()" ID="Button2" NAME="Button2">
													</td>
													<!-- '�w�N���I������Ă��Ȃ��ꍇ�́A���͕s�ɂ��� -->
												</tr>
											</table>
										</td>
										<td valign="top" nowrap>
										</td>
									</tr>
									<tr>
										<td align="right">
												<input type="hidden" name ="txtDisp" value="100" >
										</td>
									</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
</div>

</body>
</html>

<%
End Sub
%>

