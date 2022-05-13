<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo���\��(�ڍ�)
' ��۸���ID : kks/kks0111/kks0111_detail_top.asp
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
	Public m_sClassCd		'//�N���X
	Public m_sGakusekiNo	'//�w��NO
	Public m_sName			'//����
	Public m_sZaisekiName	'//�ݐЏ󋵖�
	
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
	Dim w_iRet			'// �߂�l
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
            'm_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'//�ϐ�������
		Call s_ClearParam()
		
		'// ���Ұ�SET
        Call s_SetParam()
		
		'//�ݐЋ敪�̃`�F�b�N����
		if not f_ZaisekiChk() then
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
    Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()
	
	m_sGakunenCd = 0
	m_sClassCd = 0
	m_sName = ""
	
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
	
	m_sGakunenCd = request("Nen")
	m_sClassCd = request("Class")
	m_sGakusekiNo = request("GakusekiNo")
	
	m_sName = f_GetName(m_iSyoriNen,m_sGakusekiNo)
    
End Sub
'********************************************************************************
'*  [�@�\]  �w�����̎擾
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
function f_GetName(p_SyoriNen,p_GakusekiNo)
	Dim w_iRet
    Dim w_sSQL,w_Rs
    
	f_GetName = ""
	
	On Error Resume Next
    Err.Clear
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T11_SIMEI "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T13_GAKU_NEN,"
	w_sSQL = w_sSQL & "  T11_GAKUSEKI "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  T13_GAKUSEI_NO = T11_GAKUSEI_NO "
	w_sSQL = w_sSQL & "  and T13_NENDO = " & p_SyoriNen
	w_sSQL = w_sSQL & "  and T13_GAKUSEKI_NO =" & p_GakusekiNo
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit function
	End If
	
	f_GetName = w_Rs(0)
	
end function

'********************************************************************************
'*  [�@�\]  �ݐЋ敪�̃`�F�b�N����
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
function f_ZaisekiChk()
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_Rs_Zaiseki
	Dim w_ZaisekiKbn
	
	On Error Resume Next
	Err.Clear
	
	f_ZaisekiChk = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  T13_ZAISEKI_KBN "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  T13_GAKU_NEN "
	
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "      T13_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and T13_GAKUNEN =" & m_sGakunenCd
	w_sSQL = w_sSQL & "  and T13_CLASS =" & m_sClassCd
	w_sSQL = w_sSQL & "  and T13_GAKUSEKI_NO ='" & m_sGakusekiNo & "'"
	
	w_iRet = gf_GetRecordset(w_Rs_Zaiseki,w_sSQL)
	
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	if not w_Rs_Zaiseki.EOF then
		w_ZaisekiKbn = cInt(w_Rs_Zaiseki("T13_ZAISEKI_KBN"))
		
		if w_ZaisekiKbn <> C_ZAI_ZAIGAKU then
			'�ݐВ��łȂ��Ƃ��A�ݐЋ敪�����擾
			if not f_Get_ZaisekiName(w_ZaisekiKbn,m_sZaisekiName) then exit function
		else
			m_sZaisekiName = ""
		end if
		
	end if
	
	f_ZaisekiChk = true
	
	
end function

'********************************************************************************
'*	[�@�\]	�ݐЋ敪���̂̎擾
'*	[����]	
'*	[�ߒl]	
'*	[����]	
'********************************************************************************
function f_Get_ZaisekiName(p_ZaisekiKbn,w_sZaisekiName)
	
	Dim w_sSQL
	Dim w_iRet
	Dim w_Rs_Zaiseki
	
	On Error Resume Next
	Err.Clear
	
	f_Get_ZaisekiName = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  M01_SYOBUNRUIMEI "
	
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  M01_KUBUN "
	
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "      M01_NENDO = " & m_iSyoriNen
	w_sSQL = w_sSQL & "  and M01_DAIBUNRUI_CD = " & C_ZAISEKI
	w_sSQL = w_sSQL & "  and M01_SYOBUNRUI_CD = " & p_ZaisekiKbn
	
	w_iRet = gf_GetRecordset(w_Rs_Zaiseki,w_sSQL)
	
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	w_sZaisekiName = w_Rs_Zaiseki("M01_SYOBUNRUIMEI")
	
	f_Get_ZaisekiName = true
	
end function 


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
    //************************************************************
    function window_onload() {
		
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    <%call gs_title("�o����","�Q��")%>
    <%Do %>
        
        <table>
        	<tr>
				<td class="search" nowrap>
					<table>
						<tr>
							<th class="header">�w��NO</th>
							<td><%=m_sGakusekiNo%></td>
							
							<th class="header">����</th>
							<td><%=m_sName%></td>
							
							<td><font color="#FF0000"><%=m_sZaisekiName%></font></td>
						</tr>
					</table>
				</td>
			</tr>
			
			<tr>
				<td align="center"><input type="button" value="����" onClick="javascript:parent.close();"></td>
			</tr>
		</table>
		
		
		<table>
	        <tr>
	        	<td>
	        		<table width="540" class="hyo"  border="1">
	        			<tr>
							<th class="header" width="130" align="center" nowrap>���t</th>
				            <th class="header" width="60"  align="center" nowrap>����</th>
				            <th class="header" width="150" align="center" nowrap>�Ȗ�</th>
				            <th class="header" width="120" align="center" nowrap>���͋���</th>
				            <th class="header" width="80"  align="center" nowrap>��</th>
	            		</tr>
	            	</table>
	            </td>
	        </tr>
        </table>
		
        <%Exit Do%>
    <%Loop%>
	
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>
