<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���͎������ѓo�^
' ��۸���ID : sei/sei0500/sei0500_middle.asp
' �@      �\: �������e��\������
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/09/07 ���`�i�K
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_iNendo			'�N�x
    Public  m_sKyokanCd	        '�����R�[�h
    Public  m_sSiKenCd	        '�����R�[�h
    Public  m_sGakuNo	        '�w�N
    Public  m_sClassNo	        '�w��
    Public  m_sKamokuCd	        '�ȖڃR�[�h
                                
    Public  m_Kaisi 			'���ѓ��͊��ԁi�͂��߁j
    Public  m_Syuryo			'���ѓ��͊��ԁi�����j
    Public  m_Kamokumei			'�Ȗږ�

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
	w_sMsgTitle="���͎������ѓo�^"
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

		'//���ԃf�[�^�̎擾
        w_iRet = f_Nyuryokudate()
		If w_iRet <> 0 Then 
			Exit Do
		End If

		'//�Ȗږ����擾
		w_iRet = f_GetKamokumei()
		If w_iRet <> 0 Then 
			Exit Do
		End If

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

Sub s_SetParam()
'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSiKenCd	= Cint(request("txtShikenCd"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")

End Sub

Function f_Nyuryokudate()
'********************************************************************************
'*	[�@�\]	���͊��Ԏ擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1

	Do
'2001/12/17 Mod ---->
'        w_sSQL = ""
'        w_sSQL = w_sSQL & vbCrLf & "  SELECT "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SEISEKI_KAISI, "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SEISEKI_SYURYO "
'        w_sSQL = w_sSQL & vbCrLf & "  FROM  "
'        w_sSQL = w_sSQL & vbCrLf & "    M28_SIKEN_KAMOKU M28 "
'        w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND " '�����敪(���͎����̂�)
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_NENDO        =  " & m_iNendo & "  AND "         '�����N�x
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SIKEN_CD     =  " & m_sSiKenCd & "  AND "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SIKEN_KAMOKU = '" & m_sKamokuCd & "' AND "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_GAKUNEN      =  " & m_sGakuNo & "  AND "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_CLASS        =  " & m_sClassNo & "  AND "
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SEISEKI_KAISI  <= '" & gf_YYYY_MM_DD(Date, "/") & "' AND"
'        w_sSQL = w_sSQL & vbCrLf & "    M28.M28_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(Date, "/") & "' "
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SEISEKI_KAISI, "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SEISEKI_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & "  FROM  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND "	'�����敪(���͎����̂�)
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO        =  " & m_iNendo	& "  AND "		'�����N�x
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_CD     =  " & m_sSiKenCd  & "  AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KAMOKU = '" & m_sKamokuCd & "'  "
'		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_GAKUNEN      =  " & m_sGakuNo   & "  AND "
'		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_CLASS        =  " & m_sClassNo  & "  AND "
'		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_SEISEKI_KAISI  <= '" & gf_YYYY_MM_DD(date(),"/") & "' AND"
'		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "
'2001/12/17 Mod <----

'response.write w_sSQL & "<br>"

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If Not w_Rs.EOF Then
'			m_Kaisi  = w_Rs("M28_SEISEKI_KAISI")
'			m_Syuryo = w_Rs("M28_SEISEKI_SYURYO")
			m_Kaisi  = w_Rs("M27_SEISEKI_KAISI")
			m_Syuryo = w_Rs("M27_SEISEKI_SYURYO")
		End If

	    Call gf_closeObject(w_Rs)

		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function

Function f_GetKamokumei()
'********************************************************************************
'*	[�@�\]	�Ȗږ��擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	On Error Resume Next
	Err.Clear
	f_GetKamokumei = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_KAMOKUMEI "
		w_sSQL = w_sSQL & vbCrLf & "  FROM "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO        =  " & m_iNendo	& " AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_CD     =  " & m_sSiKenCd  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KAMOKU = '" & m_sKamokuCd & "' "

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_GetKamokumei = 99
			m_bErrFlg = True
			Exit Do 
		End If

		if Not w_Rs.Eof then
			m_Kamokumei = w_Rs("M27_KAMOKUMEI")
		End if

	    Call gf_closeObject(w_Rs)

		f_GetKamokumei = 0
		Exit Do
	Loop

End Function


'********************************************************************************
'*  [�@�\]  �������ԓ����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
Function f_GetSikenJikan()

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_GetSikenJikan = ""
	p_KamokuName = ""

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_KAMOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_KAISI_JIKOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_SYURYO_JIKOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI.T26_SIKENBI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T26_SIKEN_JIKANWARI "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO = T26_SIKEN_JIKANWARI.T26_CLASS "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_GAKUNEN = M05_CLASS.M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_NENDO = M05_CLASS.M05_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_KBN=" & Cint(m_sSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_SIKEN_CD='0' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_GAKUNEN=" & cint(m_sGakuNo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKKA_CD='" & m_sGakkaCd & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T26_SIKEN_JIKANWARI.T26_KAMOKU='" & m_sKamokuCd & "'"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
			f_GetSikenJikan = 99
            Exit Do
        End If

		If Not w_Rs.EOF Then
			m_sKaisiT = w_Rs("T26_KAISI_JIKOKU")
			m_sSyuryoT = w_Rs("T26_SYURYO_JIKOKU")
			m_sSikenbi = w_Rs("T26_SIKENBI")
		End If

		f_GetSikenJikan = 0
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
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

		//�X�N���[����������
		parent.init();
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

	//************************************************************
	//	[�@�\]	�y�[�X�g�{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//************************************************************
	function f_Paste(pType){

		parent.main.document.frm.PasteType.value=pType;

		//submit�ŉ�ʂ��J���ƃE�B���h�E�̃X�e�[�^�X���ݒ�ł��Ȃ����ߤ
		//��U��y�[�W���J���Ă���A�V�E�B���h�E�ɑ΂���submit����B
		nWin=open("","Paste","location=no,menubar=no,resizable=yes,scrollbars=no,scrolling=no,status=no,toolbar=no,width=300,height=600,top=0,left=0");
		parent.main.document.frm.target="Paste";
		parent.main.document.frm.action="sei0500_paste.asp";
		parent.main.document.frm.submit();
	
	}
	//-->
	</SCRIPT>
	</head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	<% call gs_title(" ���͎������ѓo�^ "," �o�@�^ ") %>
	<center>
	<table border="0" cellpadding="0" cellspacing="0"><tr><td align="center">
		<table border=1 class=hyo>
			<tr>
				<th class=header align="center" colspan="2">���͎������ѓ��͊���</th>
			</tr>
			<tr>
				<th class=header width="96"  align="center">���ѓ��͊���</th>
				<td class=detail width="360" align="center"><%=m_Kaisi%> �` <%=m_Syuryo%></td>
			</tr>
			<tr>
				<th class=header width="96"  align="center">�Ȗ�</th>
				<td class=detail width="360" align="center"><%=m_sGakuNo%>�N�@<%= gf_GetClassName(m_iNendo,m_sGakuNo,m_sClassNo) %>�@<%= m_Kamokumei %></td>

			</tr>
		</table>
	</td></td>
	<tr><td align="center"><span class=msg>�� ���т���͂��āA�o�^�{�^���������Ă��������B</span><br>
	���w�b�_�̕����F���u<FONT COLOR="#99CCFF">����</FONT>�v�̂悤�ɂȂ��Ă��镔�����N���b�N����ƁAExcel�\��t���p�̉�ʂ��J���܂��B</span></td></tr>
	</table>

	<table width=50%>
		<tr>
			<td align=center nowrap>
				<input type=button class=button value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
				<input type=button class=button value="�L�����Z��" onclick="javascript:f_Cansel()"></td>
		</tr>
	</table>

	<table >
		<tr>
			<td valign="top">
				<table class="hyo" border=1 align="center" width="280">
					<tr>
						<th class="header" width="50"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
						<th class="header" width="200">���@��</th>
						<th class="header" width="30" nowrap onClick="f_Paste('Seiseki')"><Font COLOR="#99CCFF">����</Font></th>
					</tr>
				</table>
			</td>
			<td valign="top">
				<table class="hyo" border=1 align="center" width="280">
					<tr>
						<th class="header" width="50"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
						<th class="header" width="200">���@��</th>
						<th class="header" width="30">����</th>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	</FORM>
	</center>
	</body>
	</html>
<%
End sub

%>