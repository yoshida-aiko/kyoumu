<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �󂫎��ԏ�񌟍�
' ��۸���ID : web/web0350/web0350_top.asp
' �@      �\: �����y�[�W	 �󂫎��ԏ�񌟍����s��
'-------------------------------------------------------------------------
' ��      ��:
' ��      ��:
' ��      �n:
' ��      ��:
'           
'-------------------------------------------------------------------------
' ��      ��: 2001/08/17 ���i
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  s_sGakkaWhere		'�w�Ȃ�WHERE��
    Public  m_iJMax				'�����ő吔

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

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�󂫎��ԏ�񌟍�"
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
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '�w�ȃR���{�Ɋւ���WHERE���쐬����
        Call s_MakeGakkaWhere() 

        '//�ő厞�������擾
        Call gf_GetJigenMax(m_iJMax)
		if m_iJMax = "" Then
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
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// �I������
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �w�ȃR���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeGakkaWhere()
    
    s_sGakkaWhere = ""
    s_sGakkaWhere = m_sGakkaWhere & " M02_NENDO = " & Session("NENDO")  '//�\���N�x

End Sub

Function f_JigenCbo(p_name)
	Dim i,w_val,w_iRet
	i=0::w_val = "":w_iRet = 0
	f_JigenCbo = ""

	w_iRet = gf_GetJigenMax(w_iJMax)

	If w_iRet <> 0 then 
		f_JigenCbo = f_JigenCbo & vbCrLf & "<SELECT name='" & p_name & "' disabled>"
			f_JigenCbo = f_JigenCbo & vbCrLf & "<option></option>"
			f_JigenCbo = f_JigenCbo & vbCrLf & "</SELECT>"
	Else
			f_JigenCbo = f_JigenCbo & vbCrLf & "<SELECT name='" & p_name & "'>"
		For i = 1 to cint(w_iJMax)
			w_val = right("  "&i,2)
			f_JigenCbo = f_JigenCbo & vbCrLf & "<option value='" & i & "'>" & w_val & "</option>"
		Next
			f_JigenCbo = f_JigenCbo & vbCrLf & "</SELECT>"
	End If
	

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    On Error Resume Next
    Err.Clear

	'//�����̓��t���擾����
	w_Date = gf_YYYY_MM_DD(date(),"/")

%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
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

		// ���t�^�`�F�b�N
		if (IsDate(document.frm.txtDay.value) == 1 ){
			window.alert("���t�̓��͂Ɍ�肪����܂�");
			document.frm.txtDay.focus();
			return ;
		}

		// �J�n���������w��
		strSt = document.frm.txtJigenSt.selectedIndex;
		strEd = document.frm.txtJigenEd.selectedIndex;
			if (strSt > strEd){
			    alert("�J�n�������傫���l��I�����ĉ�����");
				document.frm.txtJigenEd.focus();
				return ;
			}
		
        document.frm.action="web0350_main.asp";
        document.frm.target="main";
        document.frm.submit();
        
    }

    //-->
    </SCRIPT>

    </head>
    <body>
    <%call gs_title("�󂫎��ԏ�񌟍�","��@��")%>

	<div align="center">

	<table border="0">
        <tr>
	        <td class="search">

				<table border="0">
    <form name="frm" method="post">
					<tr>
						<td nowrap>���@�t</td>
						<td nowrap><input type="text" name="txtDay" value="<%=w_Date%>"></td>
						<td nowrap><input type="button" value="�I��" onclick="fcalender('txtDay')"></td>
						<td nowrap>�󂫎���</td>
						<td nowrap><%=f_JigenCbo("txtJigenSt")%>������</td>
						<td nowrap><%=f_JigenCbo("txtJigenEd")%>���̊�</td>
<!--
						<td nowrap><input type="text" name="txtJigenSt" size="3">������</td>
						<td nowrap><input type="text" name="txtJigenEd" size="3">���̊�</td>
-->
					</tr>
					<tr>
						<td nowrap>�w�@��</td>
						<td nowrap><% call gf_ComboSet("txtGakka",C_CBO_M02_GAKKA,s_sGakkaWhere," style='width:115px;'",True,m_sGakkaCD) %></td>
						<td nowrap colspan="4" align="right"><input class="button" type="reset" value=" �N�@���@�A ">
							<input class="button" type="button" value="�@�\�@���@" onClick="javascript:f_Search();"></td>
					</tr>
    </form>
				</table>

	        </td>
		</tr>
	</table>
	<BR>
<table><tr><td>
<span class="msg"><font size="2">
	�� ���ׂ����󂫎��Ԃ���͂��āA�u�\���v�������Ă�������<BR>
	�� ���̊Ԃ��ׂĂɁA�󂫎��Ԃ̂��鋳���̈ꗗ���\������܂�
</font></span>
</td></tr></table>

    </div>
    </body>
    </html>
<%
End Sub
%>