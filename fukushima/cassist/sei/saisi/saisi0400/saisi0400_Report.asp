<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �ǎ���u�҈ꗗ
' ��۸���ID : saisi/saisi0300/saisi0300_Report.asp
' �@      �\: �ǎ���u�҈ꗗ �Ȗڈꗗ
'-------------------------------------------------------------------------
' ��      ��    
'               
' ��      ��
' ��      �n
'           
'           
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2003/02/20  ����
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
Dim m_Rs		'recordset

Dim m_iNendo				'�N�x
Dim m_sKyokanCd				'�����R�[�h

dim m_sKamokuCD				'�Ȗ�CD
dim m_iGakunen				'�w�N
dim m_sKamokuMei			'�Ȗږ�

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

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	
	gDisabled = ""
	
    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
	
	m_sKamokuCD = request("hidKAMOKU_CD")
	m_iGakunen = cint(request("hidMISYU_GAKUNEN"))
	m_sKamokuMei = request("hidKAMOKU_MEI")
	
End Sub


function wf_GetStudent()
'********************************************************************************
'*  [�@�\]  ���C���Ȗڎ擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	'�ϐ��̐錾
	Dim w_sSql
	Dim w_iRet

	wf_GetStudent = false

	'���N�x�f�[�^
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T13_GAKUNEN,"
	w_sSql = w_sSql & "		M05_CLASSMEI,"
	w_sSql = w_sSql & "		T13_GAKUSEKI_NO,"
	w_sSql = w_sSql & "		T11_SIMEI,"
	w_sSql = w_sSql & "		T120_NENDO, "
	w_sSql = w_sSql & "		T120_JYUKO_FLG, "
	w_sSql = w_sSql & "		T120_JYUKOKAISU, "
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "
	
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "		,T16_RISYU_KOJIN "
	w_sSql = w_sSql & "		,T13_GAKU_NEN "
	w_sSql = w_sSql & "		,M05_CLASS "
	w_sSql = w_sSql & "		,T11_GAKUSEKI "
	
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'����
	w_sSql = w_sSql & "		AND T120_KAMOKU_CD = '" & m_sKamokuCD & "' "	'�Ȗ�CD
	w_sSql = w_sSql & "		AND	T120_SYUTOKU_NENDO Is Null "				'���C��
	w_sSql = w_sSql & " 	AND T120_GAKUSEI_NO  = T16_GAKUSEI_NO"			
	w_sSql = w_sSql & "		AND T120_KAMOKU_CD = T16_KAMOKU_CD"
	w_sSql = w_sSql & "		AND T120_NENDO = T16_NENDO "
'	w_sSql = w_sSql & "	AND T16_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'���ѕs�i�]��=*�j
	w_sSql = w_sSql & "		AND T120_SEISEKI Is Null "
	'�����N�x��T13��T120����
	w_sSql = w_sSql & "		AND T13_NENDO = " & m_iNendo
	w_sSql = w_sSql & "		AND T13_GAKUSEI_NO = T120_GAKUSEI_NO"
	'�N���XM����
	w_sSql = w_sSql & "		AND M05_NENDO = T13_NENDO"	
	w_sSql = w_sSql & "		AND M05_GAKUNEN = T13_GAKUNEN"	
	w_sSql = w_sSql & "		AND M05_CLASSNO = T13_CLASS"	
	'T11��T13����
	w_sSql = w_sSql & "		AND T11_GAKUSEI_NO = T13_GAKUSEI_NO"

	w_sSql = w_sSql & "	UNION"
	
	'�ߋ��f�[�^
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T13_GAKUNEN,"
	w_sSql = w_sSql & "		M05_CLASSMEI,"
	w_sSql = w_sSql & "		T13_GAKUSEKI_NO,"
	w_sSql = w_sSql & "		T11_SIMEI,"
	w_sSql = w_sSql & "		T120_NENDO, "
	w_sSql = w_sSql & "		T120_JYUKOKAISU, "
	w_sSql = w_sSql & "		T120_JYUKO_FLG, "
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "
	
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "		,T17_RISYUKAKO_KOJIN "
	w_sSql = w_sSql & "		,T13_GAKU_NEN "
	w_sSql = w_sSql & "		,M05_CLASS "
	w_sSql = w_sSql & "		,T11_GAKUSEKI "
	
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'����
	w_sSql = w_sSql & "		AND T120_KAMOKU_CD = '" & m_sKamokuCD & "' "	'�Ȗ�CD
	w_sSql = w_sSql & "		AND	T120_SYUTOKU_NENDO Is Null "				'���C��
	w_sSql = w_sSql & " 	AND T120_GAKUSEI_NO  = T17_GAKUSEI_NO"			
	w_sSql = w_sSql & "		AND T120_KAMOKU_CD = T17_KAMOKU_CD"
	w_sSql = w_sSql & "		AND T120_NENDO = T17_NENDO "
'	w_sSql = w_sSql & "		AND T17_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'���ѕs�i�]��=*�j
	w_sSql = w_sSql & "		AND T120_SEISEKI Is Null "
	
	'�����N�x��T13��T120����
	w_sSql = w_sSql & "		AND T13_NENDO = " & m_iNendo
	w_sSql = w_sSql & "		AND T13_GAKUSEI_NO = T120_GAKUSEI_NO"
	'�N���XM����
	w_sSql = w_sSql & "		AND M05_NENDO = T13_NENDO"	
	w_sSql = w_sSql & "		AND M05_GAKUNEN = T13_GAKUNEN"	
	w_sSql = w_sSql & "		AND M05_CLASSNO = T13_CLASS"	
	'T11��T13����
	w_sSql = w_sSql & "		AND T11_GAKUSEI_NO = T13_GAKUSEI_NO"

	w_sSql = w_sSql & "	ORDER BY"
	w_sSql = w_sSql & "		T13_GAKUNEN,"	
	w_sSql = w_sSql & "		T13_CLASS, "
	w_sSql = w_sSql & "		T13_SYUSEKI_NO1 "

		
    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ں��޾�Ă̎擾���s
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
	Dim w_iJukoFlg
	Dim w_sCellClass
	Dim w_sJyuko 
	
	
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



//================================================
//	�߂鏈��
//================================================
function jf_Back() {

	location.href = "saisi0300_show.asp";
	return;

}

//-->
</script>

</head>

<body>

<form name="frm">
<center>
<br>
<table border="1" class="hyo">
	<tr>
		<td width="70"  class="header3" align="center"  height="16"><font color="#FFFFFF">���C�w�N</font></td>
		<td width="70"  class="CELL2"   height="16" align="center"><%=m_iGakunen%></td>
		<td width="70"  class="header3" align="center"  height="16"><font color="#FFFFFF">�ȁ@�@��</font></td>
		<td width="200" class="CELL2"   height="16" align="center"><%=m_sKamokuMei%></td>
	</tr>
</table>

<br>
<br>
<table border="1" class="hyo" >

	<!-- TABLE�w�b�_�� -->
	<tr>
		<td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">�w�N</font></td>
		<td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">�N���X</font></td>
		<td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">�w�Дԍ�</font></td>
		<td width="200" class="header3" align="center" height="24"><font color="#FFFFFF">���@�@�@��</font></td>
		<td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">���C�N�x</font></td>
		<td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">�󌱉�</font></td>
		<td width="70"  class="header3" align="center" height="24"><font color="#FFFFFF">�󌱓͏o</font></td>
	</tr>


	<!-- TABLE���X�g�� -->
<%

	'TD��CLASS�̏�����
	w_sCellClass = "CELL2"

	do until m_Rs.EOF
	
	'��u�t���O�`�F�b�N
	w_iJukoFlg = cint(gf_SetNull2Zero(m_Rs("T120_JYUKO_FLG")))		'cint���Ȃ��ƃG���[�ɂȂ�
	IF w_iJukoFlg = 1 then
		w_sJyuko = "��"
	ELSE
		w_sJyuko = "�@"
	END IF
	
%>
   <tr>
		<td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=gf_HTMLTableSTR(m_Rs("T13_GAKUNEN"))%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=gf_HTMLTableSTR(m_Rs("M05_CLASSMEI"))%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEKI_NO"))%></font></td>
        <td width="200" class="<%=w_sCellClass%>" align="left"   height="24">�@<%=gf_HTMLTableSTR(m_Rs("T11_SIMEI"))%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=m_Rs("T120_NENDO")%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=gf_SetNull2Zero(m_Rs("T120_JYUKOKAISU"))%></font></td>
        <td width="70"  class="<%=w_sCellClass%>" align="center" height="24"><%=w_sJyuko%></font></td>    </tr>
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

<table>
	<tr>
		<td><input type="button" value=" �߁@�� " onclick="jf_Back();"></td>
	</tr>
</table>

</center>

</form>

</body>

</html>
<%
end sub
%>