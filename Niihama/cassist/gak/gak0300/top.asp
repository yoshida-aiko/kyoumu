<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�
' ��۸���ID : gak/gak0300/top.asp
' �@      �\: ��y�[�W �w�Ѓf�[�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtHyoujiNendo         :�\���N�x
'           txtGakunen             :�w�N
'           txtGakkaCD             :�w��
'           txtClass               :�N���X
'           txtName                :����
'           txtGakusekiNo          :�w�Дԍ�
'           txtSeibetu             :����
'           txtGakuseiNo           :�w���ԍ�
'           txtIdou                :�ٓ�
'           txtTyuClub             :���w�Z�N���u
'           txtClub                :���݃N���u
'           txtRyoseiKbn           :��
'           txtMode                :���샂�[�h
'                               BLANK   :�����\��
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X - �\���N�x��\��
'                                �w�N��\��
'                                �w�Ȃ�\��
'                                �N���X��\��
'                                ���w�Z�N���u��\��
'                                ���݃N���u��\��
'                                �����敪��\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�����������ɂ��Ȃ��w������\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/23 ��Ŷ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�����I��p��Where����
    Public s_sGakkaWhere        '�w�Ȃ̒��o����
    Public m_sBukatuWhere      	'�N���u�̒��o����
    Public m_sClassWhere       	'�N���X�̒��o����
    Public m_sRyoseiKbnWhere    '�����敪�̒��o����
    Public m_sSeibetuWhere      '���ʂ̒��o����
    Public m_sIdouWhere         '�ٓ��̒��o����
    
    '�擾�����f�[�^�����ϐ�
    Public  m_TxtMode      	       ':���샂�[�h
    Public  m_iSyoriNen      	   ':�����N�x
    Public  m_iHyoujiNendo         ':�\���N�x
    Public  m_sGakunen             ':�w�N
    Public  m_sGakkaCD             ':�w��
    Public  m_sClass               ':�N���X
    Public  m_sName                ':����
    Public  m_sGakusekiNo          ':�w�Дԍ�
    Public  m_sSeibetu             ':����
    Public  m_sGakuseiNo           ':�w���ԍ�
    Public  m_sIdou                ':�ٓ�
    Public  m_sTyuClub             ':���w�Z�N���u
    Public  m_sClub                ':���݃N���u
    Public  m_sRyoseiKbn           ':��
	Public  m_sTyugaku			   ':�o�g���w�Z

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

    m_iSyoriNen = Session("NENDO")
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
        
        '�N���u�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeBukatuWhere() 

        '�N���X�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeClassWhere() 

        '���ʃR���{�Ɋւ���WHERE���쐬����
        Call s_MakeSeibetuWhere() 

        '���R���{�Ɋւ���WHERE���쐬����
        Call s_MakeRyoseiKbnWhere() 

        '�ٓ��R���{�Ɋւ���WHERE���쐬����
        Call s_MakeIdouWhere() 

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
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_IntParam()
'response.write " s_IntParam <BR>" 

	m_iHyoujiNendo =m_iSyoriNen		'�\���N�x
    m_sGakunen=""            		'�w�N
    m_sGakkaCD=""             		'�w��
    m_sClass=""               		'�N���X
    m_sName=""                		'����
    m_sGakusekiNo=""          		'�w�Дԍ�
    m_sSeibetu=""            		'����
    m_sGakuseiNo=""           		'�w���ԍ�
    m_sIdou =""               		'�ٓ�
    m_sTyuClub =""            		'���w�Z�N���u
    m_sClub=""                		'���݃N���u
    m_sRyoseiKbn=""           		'��
	m_sTyugaku = ""					'�o�g���w�Z

End Sub


'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
'response.write " s_SetParam <BR>" 

    m_iHyoujiNendo = Session("NENDO")     	 '�\���N�x
    m_sGakunen     = request("txtGakunen")           '�w�N
    m_sGakkaCD     = request("txtGakka")             '�w��
    m_sClass       = request("txtClass")             '�N���X
    m_sName        = request("txtName")              '����
    m_sGakusekiNo  = request("txtGakusekiNo")        '�w�Дԍ�
    m_sSeibetu     = request("txtSeibetu")           '����
    m_sGakuseiNo   = request("txtGakuseiNo")         '�w���ԍ�
    m_sIdou        = request("txtIdou")              '�ٓ�
    m_sTyuClub     = request("txtTyuClub")           '���w�Z�N���u
    m_sClub        = request("txtClub")              '���݃N���u
    m_sRyoseiKbn   = request("txtRyoseiKbn")         '��
	m_sTyugaku     = request("txtTyugaku")			 '�o�g���w�Z

'response.write " m_sGakunen = " & m_sGakunen & "<BR>"
End Sub


'********************************************************************************
'*  [�@�\]  �w�ȃR���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeGakkaWhere()
    
    s_sGakkaWhere = ""
    s_sGakkaWhere = m_sGakkaWhere & " M02_NENDO = " & m_iHyoujiNendo  '//�\���N�x

End Sub


'********************************************************************************
'*  [�@�\]  �N���u�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeBukatuWhere()
    
    m_sBukatuWhere = ""
    
    m_sBukatuWhere = m_sBukatuWhere & " M17_NENDO =" & m_iHyoujiNendo  '//�\���N�x
'response.write " m_sBukatuWhere=" & m_sBukatuWhere & "<BR>" 

End Sub


'********************************************************************************
'*  [�@�\]  �N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeClassWhere()
    
    m_sClassWhere = "" 

        		
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iHyoujiNendo  			'//�\���N�x
    
    if m_sGakunen <> "@@@" then
    
    	m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_sGakunen)    '//�w�N

	end if
'response.write " m_sClassWhere=" & m_sClassWhere & "<BR>" 

End Sub


'********************************************************************************
'*  [�@�\]  ���R���{�Ɋւ���WHERE���쐬����i�����敪�j
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeRyoseiKbnWhere()

    m_sRyoseiKbnWhere = ""
    

    m_sRyoseiKbnWhere = m_sRyoseiKbnWhere & " M01_NENDO = " & m_iHyoujiNendo  	'//�\���N�x
    m_sRyoseiKbnWhere = m_sRyoseiKbnWhere & " AND M01_DAIBUNRUI_CD = 23 "  	' //�����敪
    
'response.write " m_sRyoseiKbnWhere=" & m_sRyoseiKbnWhere & "<BR>" 

End Sub


'********************************************************************************
'*  [�@�\]  ���ʃR���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeSeibetuWhere()

    m_sSeibetuWhere = ""
    

    m_sSeibetuWhere  = m_sSeibetuWhere  & " M01_NENDO = " & m_iHyoujiNendo  	'//�\���N�x
    m_sSeibetuWhere  = m_sSeibetuWhere  & " AND M01_DAIBUNRUI_CD = 1 "			'//����
    
'response.write " m_sSeibetuWhere =" & m_sSeibetuWhere  & "<BR>" 

End Sub

'********************************************************************************
'*  [�@�\]  �ٓ��R���{�Ɋւ���WHERE���쐬����i�����敪�j
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeIdouWhere()

    m_sIdouWhere = ""
    
    m_sIdouWhere = m_sIdouWhere & " M01_NENDO = " & m_iHyoujiNendo  	'//�\���N�x
    m_sIdouWhere = m_sIdouWhere & " AND M01_DAIBUNRUI_CD = 9 "			'//�ݐЈٓ��敪
    
End Sub

'********************************************************************************
'*  [�@�\]  �\�������R���{�̍쐬
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function s_MakeDispCmb()

	dim sel1,sel2,sel3
	
	sel1 = ""
	sel2 = ""
	sel3 = ""

	select case request("txtDisp")
		case "50"
			sel2 = " selected"
		case "100"
			sel3 = " selected"
		case else ' = "10"
			sel1 = " selected"
	End Select

    s_MakeDispCmb = ""
    
    s_MakeDispCmb = s_MakeDispCmb & "<Select name ='txtDisp'>"
    s_MakeDispCmb = s_MakeDispCmb & "<option value='10'"&sel1&"> 10</option>"
    s_MakeDispCmb = s_MakeDispCmb & "<option value='50'"&sel2&"> 50</option>"
    s_MakeDispCmb = s_MakeDispCmb & "<option value='100'"&sel3&">100</option>"
    s_MakeDispCmb = s_MakeDispCmb & "</Select>"
    
End Function

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
        document.frm.txtName.value = "";
        document.frm.txtGakusekiNo.value = "";
        document.frm.txtSeibetu.value = "";
        document.frm.txtGakuseiNo.value = "";
        document.frm.txtTyuClub.value = "";
        document.frm.txtClub.value = "";
        document.frm.txtRyoseiKbn.value = "";
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

	<table cellspacing="0" cellpadding="0" border="0" width="100%">
		<tr><td valign="top" align="center">

			<table border="0" cellpadding="0" cellspacing="0"><tr><td class=search>

				<table border="0" bgcolor="#E4E4ED" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top">
						<table border="0">
<!--							<tr><td nowrap height="16">�\���N�x</td>
								<td>
									<select name="txtHyoujiNendo" style="width:110px;" onchange ="javascript:f_ReLoadMyPage()" >
									<% For I = 0 To 4 
										w_iNen = m_iSyoriNen -I 
										if w_iNen = cint(m_iHyoujiNendo) then %>
										<option value="<%=w_iNen%>" selected > <%=w_iNen%>�N�x</option>
									<% Else %>
										<option value="<%=w_iNen%>"> <%=w_iNen%>�N�x</option>
									<% end if 
									   Next %>
									</select>
								</td>
							</tr>
//-->
							<tr><td nowrap height="16">�w�@�N</td>
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
							</tr>
							<tr><td nowrap height="16">�w�@��</td>
							<td>
								<% call gf_ComboSet("txtGakka",C_CBO_M02_GAKKA,s_sGakkaWhere," style='width:110px;'",True,m_sGakkaCD) %>
							</td>
							</tr>				
							<tr><td nowrap height="16">�N �� �X</td>
							
							<!-- '�w�N���I������Ă��Ȃ��ꍇ�́A���͕s�ɂ��� -->
							<td>
								<%IF m_sGakunen <> "@@@" and m_sGakunen <> "" then 
								 	call gf_ComboSet("txtClass",C_CBO_M05_CLASS,m_sClassWhere," style='width:110px;'",True,m_sClass) 
							 	else %>
									<select name="txtClass" DISABLED style="width:110px;">
									<option value="@@@">�@�@�@�@�@�@�@</option>
									</select>
							    <% end if %>
							</td>
							</tr>
							<tr><td nowrap height="16">���@��</td>
							<td>
								<% call gf_ComboSet("txtSeibetu",C_CBO_M01_KUBUN,m_sSeibetuWhere," style='width:110px;'",True,m_sSeibetu) %>
							</td>
							</tr>
						</table>
					</td>
					<td valign="top">
						<table border="0">
							<tr><td nowrap height="16">����(�S�p�J�i)</td>
								<td>
									<input nowrap type="text" size="15" name="txtName" maxlength="60" value="<%=m_sName %>">
								</td>
								</tr>
								<tr><td nowrap height="16"><%=gf_GetGakuNomei(m_iHyoujiNendo,C_K_KOJIN_1NEN)%></td>
								<td>
									<input type="text" size="15" name="txtGakusekiNo" maxlength="5" value="<%=m_sGakusekiNo %>">
								</td>
							</tr>
							<tr><td nowrap height="16"><%=gf_GetGakuNomei(m_iHyoujiNendo,C_K_KOJIN_5NEN)%></td>
							<td>
								<input type="text" size="15" name="txtGakuseiNo" maxlength="10" value="<%=m_sGakuseiNo %>">
							</td>
							</tr>
						</table>
					</td>

					<td valign="top" nowrap>
						<table border="0">
							<tr><td nowrap height="16">�o�g���w�Z</td>
								<td>
									<input nowrap type="text" size="15" name="txtTyugaku" maxlength="60" value="<%=m_sTyugaku %>">
								</td>
							</tr>
							<tr><td nowrap height="16">���w�Z�N���u</td>
								<td>
									<% call gf_ComboSet("txtTyuClub",C_CBO_M17_BUKATUDO,m_sBukatuWhere," style='width:140px;'",True,m_sTyuClub) %>
								</td>
							</tr>
							<tr><td nowrap height="16">���݃N���u</td>
								<td>
									<% call gf_ComboSet("txtClub",C_CBO_M17_BUKATUDO,m_sBukatuWhere," style='width:140px;'",True,m_sClub) %>
								</td>
							</tr>
							<tr><td nowrap height="16">��</td>
								<td>					
									<% call gf_ComboSet("txtRyoseiKbn",C_CBO_M01_KUBUN,m_sRyoseiKbnWhere," style='width:140px;'",True,m_sRyoseiKbn) %>
								</td>
							</tr>
<!--
							<tr>
								<td align="right">
									<input type="checkbox" name ="CheckImage" value="image" >�摜��������
								</td>
								<td>
									<input type="button" class=button value="�N�@���@�A" onclick="jf_Clear()" >
									<input type="button" class=button value="�@�\���@" onClick="javascript:f_Search()">
								</td>
							</tr>
//-->
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<input type="checkbox" name ="CheckImage" value="image" >�摜��������
					</td>
					<td>�@�\�������@ �@ �@<%=s_MakeDispCmb()%>��
					</td>
					<td align="right">
					<input type="button" class=button value="�N�@���@�A" onclick="jf_Clear()" >
					<input type="button" class=button value=" �\�@�� " onClick="javascript:f_Search()">
				</td></tr>
				</table>

			</td>
		</tr>
		</table>
	</td></tr>
<!--
	<tr><td align="right">
		<input class=button type="button" value="�@�\���@" onClick="javascript:f_Search()">
	</td></tr>
//-->
</table>
</form>
</div>

</body>
</html>

<%
End Sub
%>

