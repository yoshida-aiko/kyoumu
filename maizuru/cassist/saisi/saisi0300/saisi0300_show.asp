<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �Ď���u�҈ꗗ
' ��۸���ID : saisi/saisi0300/saisi0300_show.asp
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
'	w_sSql = w_sSql & " SELECT DISTINCT" & vbcrlf
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, " & vbcrlf
'	w_sSql = w_sSql & "	T16_KAMOKU_KBN," & vbcrlf
'	w_sSql = w_sSql & "	T16_HISSEN_KBN," & vbcrlf
'	w_sSql = w_sSql & "	T16_COURSE_CD," & vbcrlf
'	w_sSql = w_sSql & "	T16_SEQ_NO," & vbcrlf
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, " & vbcrlf
'	w_sSql = w_sSql & "	T120_KAMOKUMEI, "
'	w_sSql = w_sSql & "	MAX(M08_MAX) "			'�s�]���̓_����������
'	
'	w_sSql = w_sSql & " FROM " & vbcrlf
'	w_sSql = w_sSql & "	T120_SAISIKEN " & vbcrlf
'	w_sSql = w_sSql & "	,T16_RISYU_KOJIN " & vbcrlf
'	w_sSql = w_sSql & "	,M08_HYOKAKEISIKI " & vbcrlf
'	
'	w_sSql = w_sSql & "	WHERE " & vbcrlf
'	w_sSql = w_sSql & "	T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'����
'	w_sSql = w_sSql & "	AND	T120_SYUTOKU_NENDO Is Null "				'���C��
'	w_sSql = w_sSql & " AND T120_GAKUSEI_NO  = T16_GAKUSEI_NO"			 & vbcrlf
'	w_sSql = w_sSql & "	AND T120_KAMOKU_CD = T16_KAMOKU_CD" & vbcrlf
'	w_sSql = w_sSql & "	AND T120_NENDO = T16_NENDO " & vbcrlf
'	w_sSql = w_sSql & "	AND T16_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'���ѕs�i�]��=*�j
'	'�J�ݎ����ɂ���ĕ]���_���敪(�O���J�݂�T16_HTEN_KIMATU_Z,����T16_HTEN_KIMATU_K)
'	w_sSql = w_sSql & " AND DECODE(T16_KAISETU ," & C_KAI_ZENKI & ",T16_HTEN_KIMATU_Z,T16_HTEN_KIMATU_K)  IS NOT NULL"	 & vbcrlf
'	w_sSql = w_sSql & "	AND DECODE(T16_KAISETU ," & C_KAI_ZENKI & ",T16_HTEN_KIMATU_Z,T16_HTEN_KIMATU_K) <= M08_MAX" & vbcrlf
'	w_sSql = w_sSql & "	AND DECODE(T16_KAISETU ," & C_KAI_ZENKI & ",T16_HTEN_KIMATU_Z,T16_HTEN_KIMATU_K) >= M08_MIN" & vbcrlf
'	w_sSql = w_sSql & "	AND M08_NENDO = T120_NENDO"			'���C�N�x�̕]��
'	w_sSql = w_sSql & " AND M08_HYOUKA_NO = 2" & vbcrlf
'	w_sSql = w_sSql & "	AND M08_HYOKA_TAISYO_KBN = 0" & vbcrlf
'	w_sSql = w_sSql & "	AND M08_HYOKA_SYOBUNRUI_RYAKU = '1'"		'�s�̉Ȗ�
'	
'	w_sSql = w_sSql & "	GROUP BY" & vbcrlf
'	w_sSql = w_sSql & "	T16_KAMOKU_KBN," & vbcrlf
'	w_sSql = w_sSql & "	T16_HISSEN_KBN," & vbcrlf
'	w_sSql = w_sSql & "	T16_COURSE_CD," & vbcrlf
'	w_sSql = w_sSql & "	T16_SEQ_NO," & vbcrlf
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, " & vbcrlf
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, " & vbcrlf
'	w_sSql = w_sSql & "	T120_KAMOKUMEI " & vbcrlf
'
'	w_sSql = w_sSql & "	UNION" & vbcrlf
'	
'	'�ߋ��f�[�^
'	w_sSql = w_sSql & " SELECT DISTINCT" & vbcrlf
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, " & vbcrlf
'	w_sSql = w_sSql & "	T17_KAMOKU_KBN," & vbcrlf
'	w_sSql = w_sSql & "	T17_HISSEN_KBN," & vbcrlf
'	w_sSql = w_sSql & "	T17_COURSE_CD," & vbcrlf
'	w_sSql = w_sSql & "	T17_SEQ_NO," & vbcrlf
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, " & vbcrlf
'	w_sSql = w_sSql & "	T120_KAMOKUMEI, " & vbcrlf
'	w_sSql = w_sSql & "	MAX(M08_MAX) "			'�s�]���̓_����������
'	
'	w_sSql = w_sSql & " FROM " & vbcrlf
'	w_sSql = w_sSql & "	T120_SAISIKEN " & vbcrlf
'	w_sSql = w_sSql & "	,T17_RISYUKAKO_KOJIN " & vbcrlf
'	w_sSql = w_sSql & "	,M08_HYOKAKEISIKI " & vbcrlf
'	
'	w_sSql = w_sSql & "	WHERE " & vbcrlf
'	w_sSql = w_sSql & "	T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"		'����
'	w_sSql = w_sSql & "	AND	T120_SYUTOKU_NENDO Is Null "				'���C��
'	w_sSql = w_sSql & " AND T120_GAKUSEI_NO  = T17_GAKUSEI_NO"			 & vbcrlf
'	w_sSql = w_sSql & "	AND T120_KAMOKU_CD = T17_KAMOKU_CD" & vbcrlf
'	w_sSql = w_sSql & "	AND T120_NENDO = T17_NENDO " & vbcrlf
'	w_sSql = w_sSql & "	AND T17_HYOKA_FUKA_KBN = " & C_HYOKA_FUKA_SESEKI	'���ѕs�i�]��=*�j
'	'�ߋ��͌�������]���_
'	w_sSql = w_sSql & " AND T17_HTEN_KIMATU_K  IS NOT NULL"	 & vbcrlf
'	w_sSql = w_sSql & "	AND T17_HTEN_KIMATU_K <= M08_MAX" & vbcrlf
'	w_sSql = w_sSql & "	AND T17_HTEN_KIMATU_K >= M08_MIN" & vbcrlf
'	w_sSql = w_sSql & "	AND M08_NENDO = T120_NENDO"			'���C�N�x�̕]��
'	w_sSql = w_sSql & " AND M08_HYOUKA_NO = 2" & vbcrlf
'	w_sSql = w_sSql & "	AND M08_HYOKA_TAISYO_KBN = 0" & vbcrlf
'	w_sSql = w_sSql & "	AND M08_HYOKA_SYOBUNRUI_RYAKU = '1'"		'�s�̉Ȗ�
'	
'	w_sSql = w_sSql & "	GROUP BY" & vbcrlf
'	w_sSql = w_sSql & "	T17_KAMOKU_KBN," & vbcrlf
'	w_sSql = w_sSql & "	T17_HISSEN_KBN," & vbcrlf
'	w_sSql = w_sSql & "	T17_COURSE_CD," & vbcrlf
'	w_sSql = w_sSql & "	T17_SEQ_NO," & vbcrlf
'	w_sSql = w_sSql & "	T120_MISYU_GAKUNEN, " & vbcrlf
'	w_sSql = w_sSql & "	T120_KAMOKU_CD, " & vbcrlf
'	w_sSql = w_sSql & "	T120_KAMOKUMEI " & vbcrlf
'	
'	w_sSql = w_sSql & "	ORDER BY 1,2,3,4,5"		'�w�N�A�Ȗڋ敪�A�K�I�敪�A�R�[�X�R�[�h�ASEQNO

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
    'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function 'GOTO LABEL_MAIN_END
    End If

'Response.write w_sSQL & "<br>"
'Response.end

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
	document.frm.action = "saisi0300_Report.asp";
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