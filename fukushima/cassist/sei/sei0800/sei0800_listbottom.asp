<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���юQ�Ɓi�������j
' ��۸���ID : sei/sei0800/default.asp
' �@      �\: 
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2003/05/13 �A�c
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

	Public  m_iNendo   			'�N�x
	Public  m_sKyokanCd			'���O�C������
	Public  m_bErrFlg			'�װ�׸�
	Public  m_Rs
	Public  m_RecCnt			'���R�[�h�J�E���g
	Dim     m_iGakunen
    Dim     m_iClass

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

	Dim w_sWinTitle
	Dim w_sMsgTitle
	Dim w_sMsg
	Dim w_sRetURL
	Dim w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���юQ��"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_parent"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// �ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'// �����`�F�b�N�Ɏg�p
'		Session("PRJ_No") = "SEI0800"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(Session("PRJ_No"))

		'//���Ұ�SET
		Call s_SetParam()

		'// �w���ꗗ�擾�A�N�Z�X�`�F�b�N
		If Not f_GetStudent() Then m_bErrFlg = True : Exit Do

		'// �Y���҂����Ȃ��ꍇ
		If m_Rs.EOF Then
			Call gs_showWhitePage("�w���f�[�^�����݂��܂���B","���юQ��")
			Exit Do
		End If

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
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()

    m_iNendo    = Session("NENDO")
    m_sKyokanCd = Session("KYOKAN_CD")
	m_iGakunen  = Request("cboGakunenCD")
	m_iClass    = Request("cboClassCD")

End Sub

Function f_GetStudent()
'********************************************************************************
'*  [�@�\]  ���O�C���������S������N���X�̊w���ꗗ���擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  True / False
'*  [����]  
'********************************************************************************

	On Error Resume Next
	Err.Clear

	Dim w_sSQL

	f_GetStudent = False

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T13_GAKUSEI_NO,  "
	w_sSQL = w_sSQL & " 	T13_GAKUSEKI_NO, "
	w_sSQL = w_sSQL & " 	T13_GAKUNEN, "
	w_sSQL = w_sSQL & " 	T11_SIMEI,   "
	w_sSQL = w_sSQL & " 	M05_CLASSMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI, "
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN, "
	w_sSQL = w_sSQL & " 	M05_CLASS "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M05_NENDO      =  " & m_iNendo   & " AND "
	w_sSQL = w_sSQL & " 	M05_GAKUNEN    =  " & m_iGakunen & " AND "
	w_sSQL = w_sSQL & " 	M05_CLASSNO    =  " & m_iClass   & " AND "
'	w_sSQL = w_sSQL & " 	M05_TANNIN     = '" & m_sKyokanCd & "' AND"
	w_sSQL = w_sSQL & " 	T13_NENDO      =  M05_NENDO      AND "
	w_sSQL = w_sSQL & " 	T13_GAKUNEN    =  M05_GAKUNEN    AND "
	w_sSQL = w_sSQL & " 	T13_GAKKA_CD   =  M05_GAKKA_CD   AND "
	w_sSQL = w_sSQL & " 	T13_GAKUSEI_NO =  T11_GAKUSEI_NO     "
	w_sSQL = w_sSQL & " ORDER BY  T13_GAKUSEI_NO     "

	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit Function

	'//ں��ރJ�E���g�擾
	m_RecCnt = gf_GetRsCount(m_Rs)

	f_GetStudent = True

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

	Dim w_sCell
	Dim w_i

	w_sCell = "CELL1"
	w_i     = 0
%>
<html>

<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--

	//************************************************************
	//  [�@�\]  �t�H�[�����[�h
	//************************************************************
	function jf_window_onload(){
		with(document.frm){
			target = "topFrame";
			action = "sei0800_listtop.asp";
			submit();
		}
	}

	//************************************************************
	//  [�@�\]  �\���{�^������
	//************************************************************
	function jf_Submit(p_i){
		with(document.frm){
			var w_Obj1 = eval("hidGakNo" + p_i);
			var w_Obj2 = eval("hidGakNM" + p_i);
			hidGakuseiNo.value = w_Obj1.value;
			hidGakuseiNM.value = w_Obj2.value;
			target = "<%=C_MAIN_FRAME%>";
			action = "sei0800_resultdef.asp";
			submit();
		}
	}

	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>

<body LANGUAGE="javascript" onload="jf_window_onload();">
	<center>
	<form name="frm" METHOD="post">
	<!-- TABLE���X�g�� -->
	<table class="hyo" align="center" border="1">

<%
	Do While Not m_Rs.EOF
		w_sCell = gf_IIF(w_sCell="CELL1","CELL2","CELL1")
		w_i = w_i + 1
%>
						<tr>
							<th width="20"  class="header3" align="center" nowrap><%=w_i%></th>
							<td width="100" class="<%=w_sCell%>" align="center" nowrap><%=m_Rs("T13_GAKUSEI_NO")%></td>
							<td width="100" class="<%=w_sCell%>" align="center" nowrap><%=m_Rs("T13_GAKUSEKI_NO")%></td>
							<td width="250" class="<%=w_sCell%>" align="left"   nowrap>�@<%=m_Rs("T11_SIMEI")%></td>
							<td width="60"  class="<%=w_sCell%>" align="center" nowrap><input type="button" value="�\�@��" class="button" onclick="jf_Submit(<%=w_i%>);" style="width:100%;"></td>
							<input type="hidden" name="hidGakNo<%=w_i%>" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<input type="hidden" name="hidGakNM<%=w_i%>" value="<%=m_Rs("T11_SIMEI")%>">
						</tr>

<%
		m_Rs.MoveNext
	Loop
%>

	</table>
	<center>

	<input type="hidden" name="hidGakuseiNo">
	<input type="hidden" name="hidGakuseiNM">
	<input type="hidden" name="hidGakunen" value="<%=m_iGakunen%>">
	<input type="hidden" name="hidClass"   value="<%=m_iClass%>">

	</form>
</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
