<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �Ď����ѓo�^
' ��۸���ID : saisi/saisi0100/saisi0100_show.asp
' �@      �\: �Ď���u�҈ꗗ �Ȗڈꗗ
'-------------------------------------------------------------------------
' ��      ��    
'               
' ��      ��
' ��      �n
'           
'           
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2003/02/20  ���
' ��      �X: 2003/03/06  ���@���X�g�A�b�v�p��SQL�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
'�G���[�n
Public  m_bErrFlg          '�װ�׸�

Dim m_Rs		'recordset

Dim m_iNendo             '�N�x
Dim m_sKyokanCd          '�����R�[�h

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

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�Ď���u�҈ꗗ"
    w_sMsg=""
    w_sRetURL="../default.asp"
    w_sTarget="_parent"

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'//�l���擾
		call s_SetParam()

		'// ���C���Ȗڎ擾
		if wf_GetStudent() = false then
			m_bErrFlg = True
            m_sErrMsg = "�Ď��Ȗڂ̎擾�Ɏ��s���܂����B"
            Exit Do
		end if



        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

	'//�����\��
    Call showPage()

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
	
	gDisabled = ""
	
    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
	
End Sub


function wf_GetStudent()
'********************************************************************************
'*  [�@�\]  ���C���Ȗڎ擾
'*  [����]  �Ȃ�
'*  [�ߒl]  true, false
'*  [����]  
'********************************************************************************

	'�ϐ��̐錾
	Dim w_sSql
	Dim w_iRet

	wf_GetStudent = false

	'���N�x�f�[�^
	w_sSql = ""

	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	T120_MISYU_GAKUNEN, "
	w_sSql = w_sSql & " 	T120_KAMOKU_CD, "
	w_sSql = w_sSql & " 	T120_KAMOKUMEI "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & " 	T120_SAISIKEN, "
	w_sSql = w_sSql & " 	M08_HYOKAKEISIKI "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	    T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'����
	w_sSql = w_sSql & " 	AND ( T120_SYUTOKU_NENDO Is Null or T120_SYUTOKU_NENDO = " & Session("NENDO") & " ) "
	w_sSql = w_sSql & " 	AND M08_NENDO = T120_NENDO"			'���C�N�x�̕]��
	w_sSql = w_sSql & " 	AND M08_HYOUKA_NO = 2 "
	w_sSql = w_sSql & " 	AND M08_HYOKA_TAISYO_KBN = 0 "
	w_sSql = w_sSql & " 	AND M08_HYOKA_SYOBUNRUI_CD = 4 "
	w_sSql = w_sSql & " 	AND T120_SEISEKI <= M08_MAX "
	w_sSql = w_sSql & " 	AND T120_SEISEKI >= M08_MIN "
	w_sSql = w_sSql & " GROUP BY "
	w_sSql = w_sSql & " 	T120_MISYU_GAKUNEN, "
	w_sSql = w_sSql & " 	T120_KAMOKU_CD, "
	w_sSql = w_sSql & " 	T120_KAMOKUMEI "

	Set m_Rs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		Exit Function
	End If

	wf_GetStudent = true

end function

sub showPage()
'********************************************************************************
'*  [�@�\]  ��ʂ̕\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

'�ϐ��̐錾
	Dim w_iTableKbn
	Dim w_sCellClass

%>
<html>

<head>
<meta http-equiv="Content-Language" content="ja">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../common/style.css" type="text/css">
<title>�Ď���u�҈ꗗ</title>

<script language="JavaScript">
<!--

//===============================================
//	���M����
//===============================================
function jf_Submit(p_iNumber,p_iGakunen,p_sName) {

	document.frm.hidKAMOKU_CD.value = p_iNumber;
	document.frm.hidMISYU_GAKUNEN.value = p_iGakunen;
	document.frm.hidKAMOKU_MEI.value = p_sName;
	document.frm.action = "saisi0100_toroku.asp";
	document.frm.submit();
	return

}

//-->
</script>

</head>

<body>

<form name="frm">
<center>
<br>
<br>
<br>
<table border="1" class="hyo" >

	<!-- TABLE�w�b�_�� -->
	<tr>
		<th width="70"  class="header3" align="center" height="24">���C�w�N</th>
		<th width="200" class="header3" align="center" height="24">�ȁ@�@�@��</th>
		<th width="70"  class="header3" align="center" height="24">��@�@�@��</th>
	</tr>
      
	<!-- TABLE���X�g�� -->
<%

	'TD��CLASS�̏�����
	w_sCellClass = "CELL2"

	do until m_Rs.EOF
%>
    <tr>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_MISYU_GAKUNEN")%></td>
		<td width="200" class="<%=w_sCellClass%>" align="left" height="24">�@<%=m_Rs("T120_KAMOKUMEI")%></td>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24">
			<input type="button" value=" �\�@�� " onclick="jf_Submit('<%=m_Rs("T120_KAMOKU_CD")%>','<%=m_Rs("T120_MISYU_GAKUNEN")%>','<%=m_Rs("T120_KAMOKUMEI")%>')">
		</td>
    </tr>
<%
		m_Rs.MoveNext
		
		if w_sCellClass = "CELL2" then
			w_sCellClass = "CELL1"
		else
			w_sCellClass = "CELL2"
		end if
		
	loop
%>
</table>

<!-- �����i�[�G���A -->
<input type="hidden" name="hidKAMOKU_CD">
<input type="hidden" name="hidKAMOKU_MEI">
<input type="hidden" name="hidMISYU_GAKUNEN">

</form>

</body>

</html>
<%
end sub
%>