<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ȓ����o�^
' ��۸���ID : gak/sei0600/sei0600_topDisp.asp
' �@      �\: 
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'               �㕔��ʕ\���̂�
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/09/26 �J�e �ǖ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n

    '�s�����I��p��Where����
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '�����R�[�h
    Public m_sGakunen       '�w�N
    Public m_sClass         '�N���X
    Public m_sClassNm       '�N���X��
    Public m_sGakusei()     '�w���̔z��
    Public m_sGakka     '�w���̏����w��
    Public m_sShiken

    Public  m_GRs
    Public  m_Rs
    Public  m_iMax          '�ő�y�[�W
    Public  m_iDsp          '�ꗗ�\���s��

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
    w_sMsgTitle="���ȓ����o�^"
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

	  '//���������擾
            If f_GetSiken(m_sShiken) <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If

		Call f_Gakusei()

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

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm    = request("txtClassNm")
	m_sShiken    = request("txtSikenKBN")

End Sub

'********************************************************************************
'*  [�@�\]  �����R���{���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSiken(p_sShiken)
    Dim w_sSQL,w_Rs

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & " M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_SYOBUNRUI_CD = " & cint(p_sShiken)

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If
	p_sShiken = w_Rs("M01_SYOBUNRUIMEI")

        f_GetSiken = 0
        Exit Do
    Loop
	Call gf_closeObject(w_Rs)

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

%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type="text/css">

<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="JavaScript">
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
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){
        parent.main.f_Touroku();
    }

	//************************************************************
	//	[�@�\]	�L�����Z���{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//************************************************************
	function f_Cansel(){

        //�����y�[�W��\��
        parent.document.location.href="default.asp"
	
	}


//-->
</SCRIPT>

</head>
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">
<center>
<%call gs_title("���ȓ����o�^","�o�@�^")%>
<table border="0" width="300" class=hyo align="center">
	<tr>
		<th width="300" class="header2" colspan="2"><%=m_sShiken%></th>
	</tr>
	<tr>
		<th width="50" class="header">�N���X</th>
		<td width="250" align="center" class="detail"><%=m_sGakunen%>�N�@<%=m_sClassNm%></td>
	</tr>
</table>
<br>
<div align="center"><span class=CAUTION>���u�݌v�v�́A�����o�����̓��j���[�����X���͂��ꂽ��L�����܂ł̌��ȏ󋵂ł��B<br>
</span></div>
	<table width="50%">
	<tr>
		<td align="center"><input type="button" class="button" value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
		<input type="button" class="button" value="�L�����Z��" onclick="javascript:f_Cansel()"></td>
	</tr>
	</table>
<table border="1" cellpadding="1" cellspacing="1" class="hyo">
			<tr>
				<th nowrap class="header" rowspan="2" width="50"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
				<th nowrap class="header" rowspan="2" width="150">�� ��</th>
				<th nowrap class="header" colspan="2"><font size="2">�� ��</font></th>
				<th nowrap class="header" colspan="2"><font size="2">���ʌ���</font></th>
			</tr>
			<tr>
				<th nowrap class="header2" width="43"><font size="1">����</font></th>
				<th nowrap class="header2" width="43"><font size="1">�݌v</font></th>
				<th nowrap class="header2" width="43"><font size="1">����</font></th>
				<th nowrap class="header2" width="43"><font size="1">�݌v</font></th>
			</tr>
</table>
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
