<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0111/kks0111_main.asp
' �@      �\: ���y�[�W ���Əo�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
'             TUKI           '//��
' ��      ��:
' ��      �n: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
'             TUKI           '//��
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��s���o�����͂�\��
'           ���o�^�{�^���N���b�N��
'               ���͏���o�^����
'-------------------------------------------------------------------------
' ��      ��: 
' ��      �X: 2015.03.19 kiyomoto Win7�Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�擾�����f�[�^�����ϐ�
    Public m_iSyoriNen      '//�����N�x
    
	Public m_JigenCount		'//������
	
	Public m_sGakunenCd		'//�w�N
	Public m_sClassCd		'//�N���XCD
	Public m_sFromDate		'//kks0111_top.asp�œ��͂������Ԃ̎n�܂�
	Public m_sToDate		'//kks0111_top.asp�œ��͂������Ԃ̏I���
	Public m_sClassName		'//�N���X��
	
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
	
	m_JigenCount = 0
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	m_sFromDate = ""
	m_sToDate = ""
	
	m_iSyoriNen = ""
    
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	
	m_iSyoriNen = Session("NENDO")
	
	m_JigenCount = request("JigenCount")
	
	m_sGakunenCd = request("cboGakunenCd")
	m_sClassCd = request("cboClassCd")
	m_sFromDate = gf_YYYY_MM_DD(request("txtFromDate"),"/")
	m_sToDate = gf_YYYY_MM_DD(request("txtToDate"),"/")
	
	m_sClassName = gf_GetClassName(m_iSyoriNen,m_sGakunenCd,m_sClassCd)
		
End Sub

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
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		//�X�N���[����������
		parent.init();
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
    
    <%call gs_title("�o����","�Q��")%>
    	<table>
			<tr>
				<td nowrap>
			        <table class="hyo" border="1" width="300">
			            <tr>
							<th class="header" width="50"  align="center" nowrap>�N���X</th>
							<td class="detail" width="150" align="left" nowrap>�@<%=m_sGakunenCd%>�N�@<%=m_sClassName%>�ȁ@</td>
						</tr>
						
						<tr>
							<th class="header" width="50"  align="center" nowrap>���t</th>
							<td class="detail" align="left" nowrap>�@<%=m_sFromDate%>�@�`�@<%=m_sToDate%>�@</td>
						</tr>
					</table>
				</td>
			</tr>
			
			<tr>
				<td align="center" nowrap>
					<table>
						<tr>
							<td valign="bottom"align="center" nowrap>
				        	    <input class="button" type="button" onclick="javascript:f_Back();" value=" �߁@�� ">
				        	</td>
						</tr>
			        </table>
				</td>
			</tr>
        </table>
		
		<!-- 2015.03.19 Upd Start kiyomoto-->
        <!--<table width=800>-->
        <table>
		<!-- 2015.03.19 Upd End kiyomoto-->
        <tr>
            <td align="center" nowrap>
	            <table class="hyo"  border="1">
				     <tr>
		                <th class="header"  rowspan="2" width="100" align="center" nowrap><font color="#ffffff">
		                    <%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></font>
		                </th>
						<th class="header" width="100" align="center" rowspan="2" nowrap><font color="#ffffff">����</font></th>
						<th class="header" width="50" align="center" rowspan="2" nowrap><font color="#ffffff">�ڍ�</font></th>
						
						<%Dim w_num%>
						<%for w_num = 1 to m_JigenCount%>
							<th class="header" width="50" align="center" colspan="2" nowrap><font color="#ffffff"><%=w_num%></font></th>
						<%next%>
					</tr>
					
					<tr>
						<%for w_num = 1 to m_JigenCount%>
							<th class="header" width="20" align="center" nowrap><font color="#ffffff">��</font></th>
							<th class="header" width="20" align="center" nowrap><font color="#ffffff">�x</font></th>
						<%next%>
					</tr>
				</table>
			</td>
        	
        	<td width="10" nowrap><br></td>
        	
        	<td align="center" width="120" nowrap>
				
	            <table width="120" class="hyo" border="1">
		            <tr>
		                <th colspan="2" class="header" align="center" width="60" nowrap><font color="#ffffff">�O��</font></th>
		                <th colspan="2" class="header" align="center" width="60" nowrap><font color="#ffffff">���</font></th>
		            </tr>
		            <tr>
		                <th class="header" width="30" align="center" nowrap><font color="#ffffff">��</font></th>
		                <th class="header" width="30" align="center" nowrap><font color="#ffffff">�x</font></th>
		                <th class="header" width="30" align="center" nowrap><font color="#ffffff">��</font></th>
		                <th class="header" width="30" align="center" nowrap><font color="#ffffff">�x</font></th>
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
%>
