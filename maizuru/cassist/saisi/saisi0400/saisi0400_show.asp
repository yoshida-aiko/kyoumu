<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �ǎ���u�҈ꗗ
' ��۸���ID : saisi/saisi0400/saisi0400_show.asp
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
' ��      �X: 2003/02/27  ���
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
    w_sMsgTitle="�ǎ���u�҈ꗗ"
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
            m_sErrMsg = "�ǎ��Ȗڂ̎擾�Ɏ��s���܂����B"
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
'	w_sSql = ""
'	w_sSql = w_sSql & " SELECT DISTINCT"
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, "
'	w_sSql = w_sSql & "	T16_KAMOKU_KBN,"
'	w_sSql = w_sSql & "	T16_HISSEN_KBN,"
'	w_sSql = w_sSql & "	T16_COURSE_CD,"
'	w_sSql = w_sSql & "	T16_SEQ_NO,"
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, "
'	w_sSql = w_sSql & "	T120_KAMOKUMEI "
'	
'	w_sSql = w_sSql & " FROM "
'	w_sSql = w_sSql & "	T120_SAISIKEN "
'	w_sSql = w_sSql & "	,T16_RISYU_KOJIN "
'	
'	w_sSql = w_sSql & "	WHERE "
'	w_sSql = w_sSql & "	T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'����
'	w_sSql = w_sSql & "	AND	T120_SYUTOKU_NENDO Is Null "				'���C��
'	w_sSql = w_sSql & " AND T120_GAKUSEI_NO  = T16_GAKUSEI_NO"			
'	w_sSql = w_sSql & "	AND T120_KAMOKU_CD = T16_KAMOKU_CD"
'	w_sSql = w_sSql & "	AND T120_NENDO = T16_NENDO "
''	w_sSql = w_sSql & "	AND T16_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'���ѕs�i�]��=*�j
'	w_sSql = w_sSql & "	AND T120_SEISEKI Is Null "
'	
'	w_sSql = w_sSql & "	GROUP BY"
'	w_sSql = w_sSql & "	T16_KAMOKU_KBN,"
'	w_sSql = w_sSql & "	T16_HISSEN_KBN,"
'	w_sSql = w_sSql & "	T16_COURSE_CD,"
'	w_sSql = w_sSql & "	T16_SEQ_NO,"
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, "
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, "
'	w_sSql = w_sSql & "	T120_KAMOKUMEI "
'
'	
'	w_sSql = w_sSql & "	UNION"
'	
'	'�ߋ��f�[�^
'	w_sSql = w_sSql & " SELECT DISTINCT"
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, "
'	w_sSql = w_sSql & "	T17_KAMOKU_KBN,"
'	w_sSql = w_sSql & "	T17_HISSEN_KBN,"
'	w_sSql = w_sSql & "	T17_COURSE_CD,"
'	w_sSql = w_sSql & "	T17_SEQ_NO,"
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, "
'	w_sSql = w_sSql & "	T120_KAMOKUMEI "
'	
'	w_sSql = w_sSql & " FROM "
'	w_sSql = w_sSql & "	T120_SAISIKEN "
'	w_sSql = w_sSql & "	,T17_RISYUKAKO_KOJIN "
'	
'	w_sSql = w_sSql & "	WHERE "
'	w_sSql = w_sSql & "	T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'����
'	w_sSql = w_sSql & "	AND	T120_SYUTOKU_NENDO Is Null "				'���C��
'	w_sSql = w_sSql & " AND T120_GAKUSEI_NO  = T17_GAKUSEI_NO"			
'	w_sSql = w_sSql & "	AND T120_KAMOKU_CD = T17_KAMOKU_CD"
'	w_sSql = w_sSql & "	AND T120_NENDO = T17_NENDO "
''	w_sSql = w_sSql & "	AND T17_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'���ѕs�i�]��=*�j
'	w_sSql = w_sSql & "	AND T120_SEISEKI Is Null "
'
'	w_sSql = w_sSql & "	GROUP BY"
'	w_sSql = w_sSql & "	T17_KAMOKU_KBN,"
'	w_sSql = w_sSql & "	T17_HISSEN_KBN,"
'	w_sSql = w_sSql & "	T17_COURSE_CD,"
'	w_sSql = w_sSql & "	T17_SEQ_NO,"
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, "
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, "
'	w_sSql = w_sSql & "	T120_KAMOKUMEI "
'	
'	w_sSql = w_sSql & "	ORDER BY 1,2,3,4,5"		'�w�N�A�Ȗڋ敪�A�K�I�敪�A�R�[�X�R�[�h�ASEQNO

	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		T120_MISYU_GAKUNEN, "
	w_sSql = w_sSql & "		T120_KAMOKU_CD, "
	w_sSql = w_sSql & "		T120_KAMOKUMEI "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T120_SAISIKEN "
	w_sSql = w_sSql & "	WHERE "
	w_sSql = w_sSql & "		    T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'����
	w_sSql = w_sSql & "		AND	T120_SYUTOKU_NENDO Is Null "					'���C��
	w_sSql = w_sSql & "		AND	T120_NENDO = " & Session("NENDO")				'�N�x
	w_sSql = w_sSql & "		AND	T120_SEISEKI Is Null "					'�����_��
	w_sSql = w_sSql & " GROUP BY "
	w_sSql = w_sSql & "		T120_MISYU_GAKUNEN, "
	w_sSql = w_sSql & "		T120_KAMOKU_CD, "
	w_sSql = w_sSql & "		T120_KAMOKUMEI "
		
    Set m_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

    If w_iRet <> 0 Then
    'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

'Response.write gf_GetRsCount(m_Rs) & "<br>"
'Response.Write session("KYOKAN_CD") & "<br>"

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
<title>�ǎ���u�҈ꗗ</title>

<script language="JavaScript">
<!--

//===============================================
//	���M����
//===============================================
function jf_Submit(p_iNumber,p_iGakunen,p_sName) {

	document.frm.hidKAMOKU_CD.value = p_iNumber;
	document.frm.hidMISYU_GAKUNEN.value = p_iGakunen;
	document.frm.hidKAMOKU_MEI.value = p_sName;
	document.frm.action = "saisi0400_Report.asp";
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