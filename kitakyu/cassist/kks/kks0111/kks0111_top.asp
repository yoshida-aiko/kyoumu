<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo���\��
' ��۸���ID : kks/kks0111/kks0111_top.asp
' �@      �\: ���Əo���̌����y�[�W
'-------------------------------------------------------------------------
' ��      ��:�N�x           ��      SESSION("NENDO")
'            
' ��      ��:
' ��      �n:
'            
'            
' ��      ��:
'           �������\��
'               �w�N�̃R���{�{�b�N�X��1�N��\��
'               �N���X�̃R���{�{�b�N�X��CLASSNO��1��\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ����Əo���ꗗ��\��������
'-------------------------------------------------------------------------
' ��      ��: 2002/05/07 shin
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const C_FIRST_DISP_GAKUNEN = 1   '//�����\���̎��̊w�N(1�N)
	
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	Public m_iSyoriNen		'//�����N�x
	Public m_sGakki			'//�w��
	Public m_sZenki_Start	'//�O���J�n��
	Public m_sKouki_Start	'//����J�n��
	Public m_sKouki_End		'//����I����
	
	Public m_iGakunenCd		'//�w�NCD
	Public m_Date			'//�V�X�e�����t
	
	Public m_sGakunenWhere	'//�w�N�R���{��WHERE��
	Public m_sClassWhere	'//�N���X�R���{��WHERE��
	
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
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
	Dim w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget
    
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���Əo���ꗗ"
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
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()
	m_Date = ""
	
	m_iGakunenCd	= 0
    m_iSyoriNen		= 0
    
    m_sGakki		= ""
	m_sZenki_Start	= ""
	m_sKouki_Start	= ""
	m_sKouki_Start	= ""
	
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	m_Date = gf_YYYY_MM_DD(date(),"/")				'//�V�X�e�����t���Z�b�g
	
	m_iGakunenCd = cInt(request("cboGakunenCd"))	'�����[�h���ɃZ�b�g(�V�K���́A"")
	m_iSyoriNen = Session("NENDO")
	
End Sub

'********************************************************************************
'*  [�@�\]  �w�N�R���{��WHERE���̍쐬
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_GakunenWhere
	
	m_sGakunenWhere = ""
	m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyorinen
	m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"
	
End Sub

'********************************************************************************
'*  [�@�\]  �N���X�R���{��WHERE���̍쐬
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClassWhere
	
	m_sClassWhere = ""
	m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen
	
	If m_iGakunenCd = 0 Then
		'//�����\������1�N1�g��\������
		m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & C_FIRST_DISP_GAKUNEN
	Else
		m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & m_iGakunenCd
	End If
	
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
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>���Əo���ꗗ</title>
	
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
	//************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){
		if(!f_InpChk()){ return false; }
		
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
		var ob = new Array();
		ob[0] = eval("document.frm.txtFromDate");
		ob[1] = eval("document.frm.txtToDate");
		
		//���J�n��
        //NULL�`�F�b�N
        if(f_Trim(ob[0].value) == ""){
            f_InpChkErr("�J�n�������͂���Ă��܂���",ob[0]);
            return false;
        }
        
        //�^�`�F�b�N
        if(IsDate(ob[0].value) != 0){
        	f_InpChkErr("�J�n���̓��t���s���ł�",ob[0]);
        	return false;
        }
        
        //�O���J�n��<=�J�n��<=����I�����̃`�F�b�N
        if(DateParse("<%=m_sZenki_Start%>",ob[0].value) < 0 || DateParse(ob[0].value,"<%=m_sKouki_End%>") < 0){
			f_InpChkErr("�J�n���ɂ́A�O���J�n���Ȍ�A����I�����ȑO�̓��t����͂��Ă�������",ob[0]);
			return false;
		}
        
        //���I����
        //NULL�`�F�b�N
        if(f_Trim(ob[1].value) == ""){
			f_InpChkErr("�I���������͂���Ă��܂���",ob[1]);
			return false;
        }
        
        //�^�`�F�b�N
        if(IsDate(ob[1].value) != 0){
			f_InpChkErr("�I�����̓��t���s���ł�",ob[1]);
        	return false;
        }
        
        //�O���J�n��<=�I����<=����I�����̃`�F�b�N
        if(DateParse("<%=m_sZenki_Start%>",ob[1].value) < 0 || DateParse(ob[1].value,"<%=m_sKouki_End%>") < 0){
			f_InpChkErr("�I�����ɂ́A�O���J�n���Ȍ�A����I�����ȑO�̓��t����͂��Ă�������",ob[1]);
			return false;
		}
        
        //�����Ԃ̎擾��������
        if(DateParse(ob[0].value,ob[1].value) < 0){
        	f_InpChkErr("�J�n���ƏI�����𐳂������͂��Ă�������",ob[0]);
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
    //  [�@�\]  �w�N��ύX������(�N���X�����Z�b�g���Ȃ�������)
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){
		document.frm.action = "kks0111_top.asp";
        document.frm.target = "topFrame";
        document.frm.submit();
		
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript">
    <%call gs_title("�o����","�Q��")%>
    <form name="frm" method="post">
	
	<center>
    <table border="0">
	    <tr>
		    <td align="right" class="search" nowrap>
				
			    <table border="0">
					<tr>
						<td align="left" nowrap>�w�N</td>
						<td align="left" nowrap>
							<%  
								call s_GakunenWhere()	'�w�N�R���{��WHERE��
								
								call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange='javascript:f_ReLoadMyPage();' style='width:40px;' ",False,m_iGakunenCd)
							%>
							
							<font>�N</font>
							
							<font>�@�@�@�@�@�@�@�@�N���X</font>
							<%
								call s_ClassWhere()		'�N���X�R���{��WHERE��
								
								call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"style='width:80px;' ",False,"")
							%>
						</td>
						<td align="left" nowrap><br></td>
					</tr>
					
					<tr>
						<td align="left" nowrap>���t</td>
						<td nowrap>
							<input type="text" name="txtFromDate" value="<%=m_Date%>">
							<input type="button" class="button" onclick="fcalender('txtFromDate')" value="�I��">�@�`�@
							
							<input type="text" name="txtToDate" value="<%=m_Date%>">
							<input type="button" class="button" onclick="fcalender('txtToDate')" value="�I��">
							
						</td>
						
						<td valign="bottom" align="right" nowrap>
							<input class="button" type="button" onclick="javascript:f_Search();" value="�@�\�@���@">
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	
    <!--�l�n���p-->
    <input type="hidden" name="txtURL" VALUE="kks0111_bottom.asp">
    <input type="hidden" name="txtMsg" VALUE="���΂炭���҂���������">
	
	</center>
    </form>
    
    </body>
    </html>
<%
End Sub
%>