<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0112/kks0112_middle.asp
' �@      �\: ��y�[�W �O�y�[�W�̌���������\��
'-------------------------------------------------------------------------
' ��      ��: 
'             
'             
'             
'             
' ��      ��:
' ��      �n: 
'             
'             
'             
'             
' ��      ��:
'           �������\��
'               ���t�F�O�y�[�W�̌���������\��
'               �����F�O�y�[�W�̌���������\��
'               �ȖځF�O�y�[�W�̌���������\��
'               �N���X�F�O�y�[�W�̌���������\��
'               ���͋敪�F���ہA�x���A���ށA�N���A�̃��W�I�{�^��
'           ���o�^�{�^���N���b�N��
'               ���͏���o�^����
'           ���߂�{�^���N���b�N��
'               �O�y�[�W�ɖ߂�
'-------------------------------------------------------------------------
' ��      ��: 2002/05/16 shin
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '//�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '//�ϐ�
    Public m_iSyoriNen      '//�����N�x
    Public m_sDate			'//���t
	Public m_iJigen			'//������
	Public m_iGakunen		'//�w�N
	Public m_sClassName		'//�N���X��
	Public m_sKamokuName	'//�Ȗږ�
	
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
            w_sMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
        '//�ϐ�������
        Call s_ClearParam()
		
        '// ���Ұ�SET
        Call s_SetParam()
		
		Call showPage_middle()
		
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
	
	m_iJigen = 0
	m_iGakunen = 0
	m_iSyoriNen = 0
	m_sClassName = ""
	m_sKamokuName = ""
	m_sDate = ""
    
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	m_sDate = trim(Request("txtDate"))
	m_iJigen = trim(Request("sltJigen"))
	
	m_iGakunen = trim(Request("hidGakunen"))
	m_sClassName = trim(Request("hidClassName"))
	m_sKamokuName = trim(Request("hidKamokuName"))
	
    m_iSyoriNen = Session("NENDO")
    
End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage_middle()
	dim w_str	'�\�����b�Z�[�W

    On Error Resume Next
    Err.Clear
	
	w_str = "<span class='CAUTION'>�� ���͂������u���͋敪�v��I����A�Y������w���̏o���󋵗����N���b�N���ĉ������B<BR></span>" & vbCrLf
	
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
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Insert(){
		parent.frames["main"].f_Insert();
		return;
    }
	
    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
    //  [����]  
    //  [�ߒl]  
    //  [����]
    //************************************************************
    function f_Back(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default.asp"
    }
	
    //-->
    </SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    <%call gs_title("���Əo������","�o�@�^")%>
    	<table height="160">
    		<tr>
				<td>
	        		<table class="hyo" border="1" width="550">
	            		<tr>
							<th nowrap class="header" width="65" align="center">���t</th>
			                <td nowrap class="detail" width="120" align="center"><%=m_sDate%></td>
			                <th nowrap class="header" width="70" align="center">����</th>
			                <td nowrap class="detail" width="30" align="center"><%=m_iJigen%></td>
			                <th nowrap class="header" width="65" align="center">�Ȗ�</th>
			                <td nowrap class="detail" width="150" align="center"><%=m_sKamokuName%></td>
						</tr>

						
						<tr>
							<th nowrap class="header" width="65"  align="center">�N���X</th>
			                <td nowrap class="detail" width="120"  align="center"><%=m_iGakunen & "�N " & m_sClassName & "�� " %></td>
			                <th nowrap class="header" width="70"  align="center">���͋敪</th>
			                
			                <td nowrap class="detail" width=""  align="center" colspan="3">
			                	<input type="radio" name="rdoType" value="1" checked>
			                	<input type="text" name="txtKekka" size="2" maxlength="2" value="1">����
			                	
			                	<input type="radio" name="rdoType" value="2" >�x��
			                	<input type="radio" name="rdoType" value="3" >����
			                	<input type="radio" name="rdoType" value="4" >�N���A
			                </td>
			            </tr>
	        		</table>
				</td>
			</tr>
			
			<tr>
				<td align="center">
					<table>
						<tr>
							<td><input type="button" name="btnInsert" value="�@�o�@�^�@" onClick="f_Insert();"></td>
							<td><input type="button" name="btnBack" value="�@�߁@��@" onClick="f_Back();"></td>
						</tr>
	      			</table>
				</td>
			</tr>
			
			
			<tr>
				<td align="center">
					<table>
						<tr>
							<td><%=w_str%></td>
						</tr>
	      			</table>
				</td>
			</tr>
			
			<tr>
				<td align="center">
					<table class="hyo" border="1" width="300">
						<tr>
							<th nowrap class="header" width="80"  align="center">�w�Дԍ�</th>
							<th nowrap class="header" width="150"  align="center">���@��</th>
							<th nowrap class="header" width="70"  align="center">��</th>
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
