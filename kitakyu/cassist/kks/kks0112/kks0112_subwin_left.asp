<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo���\��(�ڍ�)
' ��۸���ID : kks/kks0112/kks0112_subwin_left.asp
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
	Public m_sClassName		'//�N���X��
	
	Public m_sKamokuCd		'//�ȖڃR�[�h
    Public m_sKamokuName	'//�Ȗږ�
    
    Public m_iKamokuKbn		'//���
	Public m_iMonth			'//��
	
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
            w_sMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'//�ϐ�������
		Call s_ClearParam()
		
		'// ���Ұ�SET
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
	
	m_iSyoriNen = 0
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	m_sClassName = ""
	
    m_sKamokuCd = ""
    m_sKamokuName = ""
    
    m_iKamokuKbn = 0
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
	m_sClassName = gf_GetClassName(m_iSyoriNen,m_sGakunenCd,m_sClassCd)
	
    m_sKamokuCd = request("hidKamokuCd")
    m_iKamokuKbn = cint(request("hidSyubetu"))
    m_sKamokuName = gf_GetKamokuMei(m_iSyoriNen,m_sKamokuCd,m_iKamokuKbn)
    
	m_iMonth = request("sltMonth")
	
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
	
	<!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    
	//************************************************************
    //  [�@�\]  �y�[�W�N���[�Y
    //************************************************************
    function f_Close(){
		parent.close();
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript">
    <form name="frm" method="post">
    <center>
		<table>
			<tr><td><br><br></td></tr>
			
			<tr>
				<td align="center" colspan="2"><font size="+2"><%=m_iMonth%>��</font></td>
			</tr>
			
			<tr><td><br></td></tr>
			
			<tr>
				<td>
					<table class="hyo" border="1">
						<tr>
							<th class="header" width="45">�w�N</th>
							<td class="detail" width="120" align="left">&nbsp;<%=m_sGakunenCd%>�N</td>
						</tr>
					</table>
				<td>
			</tr>
			
			<tr><td><br></td></tr>
			
			<tr>
				<td>
					<table class="hyo" border="1">
						<tr>	
							<th class="header" width="45">�N���X</th>
							<td class="detail" width="120" align="left">&nbsp;<%=m_sClassName%>��</td>
						</tr>
					</table>
				<td>
			</tr>
			
			<tr><td><br></td></tr>
			
			<tr>
				<td>
					<table class="hyo" border="1">
						<tr>
							<th class="header" width="45">�Ȗ�</th>
							<td class="detail" width="120" align="left">&nbsp;<%=m_sKamokuName%></td>
						</tr>
					</table>
				<td>
			</tr>
			
			<tr><td><br></td></tr>
			
			<tr>
				<td align="center" colspan="2"><input type="button" value="����" onClick="f_Close();" name="btnClose"></td>
			</tr>
		</table>
		
	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>
