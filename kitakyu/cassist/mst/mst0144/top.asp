<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A�E��}�X�^�o�^
' ��۸���ID : mst/mst0133/top.asp
' �@      �\: ��y�[�W �A�E��}�X�^�̓o�^���s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroCD2     :�i�H�R�[�h
'           txtSingakuCD2       :�i�w�R�[�h
'           txtSinroName        :�A�E�於�́i�ꕔ�j
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��A�E���\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/22 �≺�@�K��Y
' ��      �X: 2001/07/13 �J�e�@�ǖ�
' ��      �X: 2001/08/22 �ɓ��@���q�@�Ǝ�敪�ǉ��Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    '�s�����I��p��Where����
    Public m_sSinroWhere            '�i�H�̏���
    Public m_sSingakuWhere      '�i�w�R���{�̏���
    Public m_sSingakuOption     '�i�w�R���{�̃I�v�V����
    Public m_sSyusyokuName
    Public m_sSinroCD
    Public m_sSingakuCD
    Public m_iNendo
    
    Public m_sMode

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
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E��}�X�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

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

        '// ���Ұ�SET
        Call s_SetParam()
        '�i�H�Ɋւ���WHRE���쐬����
        Call f_MakeSinroWhere() 
        '�i�H�Ɋւ���WHRE���쐬����
        Call f_MakeSingakuWhere()   
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
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_sSinroCD = Request("txtSinroCD")      ':�i�H�R�[�h
	If m_sSinroCD = "@@@" Then
		m_sSinroCD = ""
	End If
    m_sSingakuCD = Request("txtSingakuCD")  ':�i�w�R�[�h
    m_sSyusyokuName = Request("txtSyusyokuName")
    m_sMode = "search"
    m_iNendo = Session("nendo")

End Sub


Sub f_MakeSinroWhere()
'********************************************************************************
'*  [�@�\]  �i�H�R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sSinroWhere=""

    m_sSinroWhere = " M01_DAIBUNRUI_CD = " & C_SINRO & "  AND "
    m_sSinroWhere = m_sSinroWhere & " M01_NENDO = " & m_iNendo & ""


End Sub

Sub f_MakeSingakuWhere()
'********************************************************************************
'*  [�@�\]  �i�w�R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sSingakuWhere=""
    m_sSingakuOption=""

'---2001/08/22 ito �Ǝ�敪�ǉ��Ή�
	'// �i�w
    If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU Then
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_SINGAKU & "  AND "
        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
	'// �A�E
	ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = " & C_GYOSYU_KBN & "  AND "
        m_sSingakuWhere = m_sSingakuWhere & " M01_NENDO = " & m_iNendo & ""
	'// ���̑�
    Else
        m_sSingakuWhere= " M01_DAIBUNRUI_CD = 0 "
        m_sSingakuOption = " DISABLED "
    End IF

End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_sSelectSinroCd
Dim w_sSelectSingakuCd

%>


<html>

<head>

<title>�A�E��}�X�^�o�^</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �i�H���C�����ꂽ�Ƃ��A�ĕ\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){
        document.frm.action="top.asp";
        document.frm.target="_self";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

        document.frm.action = "./main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Touroku(){

        document.frm.action = "syusei.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "Sinki";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>


<link rel=stylesheet href="../../common/style.css" type=text/css>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" method="POST"  onSubmit="return false" onClick="return false">
<%call gs_title("�i�H����o�^","��@��")%>

    <table border="0">
    <tr>
    <td>

        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left" class=search>

                <table border="0" cellpadding="1" cellspacing="1">
                <tr>
                <td Nowrap align="left">
		            �i�H�敪<img src="../../image/sp.gif" width="15">
					<% '���ʊ֐�����i�H�Ɋւ���R���{�{�b�N�X���o�͂���i�N�x�����j
					    call gf_ComboSet("txtSinroCD",C_CBO_M01_KUBUN,m_sSinroWhere," onchange = 'javascript:f_ReLoadMyPage();' ",True,m_sSinroCD) %>

                </td>
                <td Nowrap align="left">
					<%
					'If cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU Then
					'	w_sTitle = "�Ǝ�敪"
					'Else
					'	w_sTitle = "�i�w�敪"
					'End If
					%>
					�@��ʋ敪
					<% '���ʊ֐�����i�w�Ɋւ���R���{�{�b�N�X���o�͂���i�N�x�A�i�H�敪�������j�i�i�H�敪�����͂���Ă��Ȃ��Ƃ��́ADISABLED�ƂȂ�j
					   call gf_ComboSet("txtSingakuCD",C_CBO_M01_KUBUN,m_sSingakuWhere,"style='width:100px;' " & m_sSingakuOption,True,m_sSingakuCd) %>

                </td>
                </tr>

                <tr>
                <td align="left" colspan="2">
	                �i�H�於��
	                <input type="text" name="txtSyusyokuName" size="20" Value="<%=m_sSyusyokuName%>" maxlength="60">   <!--'//2001/07/31�C��-->
	                <font size="2">���i�H�於�̂̈ꕔ�Ō������܂�</font>
                </td>
                </tr>
                <tr>
                <td Nowrap align="right" colspan="2">
			    <input class=button type="button" value="�@�\�@���@" onClick="javascript:f_Search()">
                </td>
                </tr>
                </table>

        </td>
        </tr>
        </table>

    </td>
    <td valign="top">
    <a href="javascript:f_Touroku()" onClick="javascript:f_Touroku()">�V�K�o�^�͂�����</a><br><img src="../../image/sp.gif" height="10"><br>
    </td>
  </tr>
</table>
<input type="hidden" name="txtMode" value="<%=m_sMode%>">
<input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
</form>

</center>

</body>

</html>



<%
    '---------- HTML END   ----------
End Sub
%>
