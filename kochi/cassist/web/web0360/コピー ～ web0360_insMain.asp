<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����������ꗗ
' ��۸���ID : web/web0360/web0360_main.asp
' �@      �\: ������\��
'-------------------------------------------------------------------------
' ��      ��:   txtClubCd		:����CD
'               cboGakunenCd	:�w�N
'               cboClassCd		:�N���XNO
'               txtTyuClubCd	:���w�Z����CD
'
' ��      �n:   txtMode			:�������[�h
'               txtClubCd		:����CD
'               GAKUSEI_NO		:�w��NO
'               cboGakunenCd	:�w�N
'               cboClassCd		:�N���XNO
'               txtTyuClubCd	:���w�Z����CD
' ��      ��:
'           �������\��
'               �󔒃y�[�W��\��
'           ���\���{�^���������ꂽ�ꍇ
'               �E���������ɂ��Ȃ������k�ꗗ��\������
'               �E��������Ƃ����܂��Ă��鐶�k�̓����o�^�͕s�Ƃ���(�I���`�F�b�N�{�b�N�X��\�����Ȃ�)
'               �E���łɓo�^�Ώە����ɓ������Ă��鐶�k�̓����o�^�͕s�Ƃ���(�I���`�F�b�N�{�b�N�X��\�����Ȃ�)
'-------------------------------------------------------------------------
' ��      ��: 2001/08/22 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	Public m_iSyoriNen			'//�N�x
	Public m_iKyokanCd			'//��������
	Public m_sClubCd			'//�N���uCD
	Public m_iGakunen           '//�w�N
	Public m_iClassNo           '//�N���XNO
	Public m_sTyuClubCd			'//���w�Z�N���uCD

    'ں��ރZ�b�g
	Public m_Rs					'//�����ꗗں��޾��
	Public m_iRsCnt				'//ں��ރJ�E���g

	'�G���[�n
	Public  m_bErrFlg           '�װ�׸�
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

	Dim w_iRet  			'// �߂�l

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="�����������ꗗ"
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

		'//�l�̏�����
		Call s_ClearParam()

		'//�ϐ��Z�b�g
		Call s_SetParam()

'//�f�o�b�O
'Call s_DebugPrint()

		'//���k�ꗗ�̎擾
		w_iRet = f_GetSeitoData()
		If w_iRet <> 0 Then
			m_bErrFlg = True
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

	'//ں��޾��CLOSE
	Call gf_closeObject(m_Rs)

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

	m_iSyoriNen  = ""
	m_iKyokanCd  = ""
	m_sClubCd  	= ""
	m_iGakunen   = ""
	m_iClassNo   = ""
	m_sTyuClubCd = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

	m_iSyoriNen  = Session("NENDO")
	m_iKyokanCd  = Session("KYOKAN_CD")
	m_sClubCd    = Request("txtClubCd")
	m_iGakunen   = Request("cboGakunenCd")	'//�w�N
	m_iClassNo   = gf_cboNull(Request("cboClassCd"))	'//�N���X
	m_sTyuClubCd = replace(Request("txtTyuClubCd"),"@@@","")	'//���w�Z�N���uCD
	Session("HyoujiNendo") = m_iSyoriNen
End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iSyoriNen  = " & m_iSyoriNen  & "<br>"
	response.write "m_iKyokanCd  = " & m_iKyokanCd  & "<br>"
	response.write "m_sClubCd    = " & m_sClubCd    & "<br>"
	response.write "m_iGakunen   = " & m_iGakunen   & "<br>"
	response.write "m_iClassNo   = " & m_iClassNo   & "<br>"
	response.write "m_sTyuClubCd = " & m_sTyuClubCd & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  ���k�ꗗ���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSeitoData()

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetSeitoData = 1

	Do

		'//���k�ꗗ
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI.T11_SIMEI"
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI.T11_TYU_CLUB"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO"
		w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_GAKUNEN=" & m_iGakunen
		
		if m_iClassNo <> "" then
				w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_CLASS=" & m_iClassNo
		End if
		If m_sTyuClubCd <> "" Then
			w_sSql = w_sSql & vbCrLf & "  AND  T11_GAKUSEKI.T11_TYU_CLUB='" & m_sTyuClubCd & "'"
		End If

		w_sSql = w_sSql & vbCrLf & " ORDER BY "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUNEN, T13_GAKU_NEN.T13_CLASS, T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write w_sSQL & "<br>"
		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_GetSeitoData = 99
			Exit Do
		End If

		'//ں��ރJ�E���g�擾
		'//�������擾
		m_iRsCnt = 0
		If m_Rs.EOF = False Then
			m_iRsCnt = gf_GetRsCount(m_Rs)
		End If

		'//����I��
		f_GetSeitoData = 0
		Exit Do
	Loop


End Function

'********************************************************************************
'*  [�@�\]  �N���X�����擾
'*  [����]  p_iGakuNen:�w�N,p_iClassNo:�N���XNO
'*  [�ߒl]  f_GetClassName:�N���X��
'*  [����]  
'********************************************************************************
Function f_GetClassName(p_iGakuNen,p_iClassNo)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassName = ""
	w_sClassName = ""

	Do
		'�N���X�}�X�^���f�[�^���擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_CLASSMEI"
		w_sSql = w_sSql & vbCrLf & "  ,M05_CLASS.M05_GAKKA_CD"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN= " & p_iGakuNen
		w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_CLASSNO= "   & p_iClassNo

'response.write w_sSQL & "<br>"

		'//�f�[�^�擾
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			Exit Do
		End If

		If rs.EOF = False Then
			w_sClassName = rs("M05_CLASSMEI")
			'w_sGakkaCd = rs("M05_GAKKA_CD")
		End If

		Exit Do
	Loop

	'//�߂�l���
	f_GetClassName = w_sClassName

	'//ں���CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  ���������擾����
'*  [����]  p_sClubCd:����CD
'*  [�ߒl]  f_GetClubName�F������
'*  [����]  
'********************************************************************************
Function f_GetClubName(p_sClubCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubName = ""
	w_sClubName = ""

	Do

		'//����CD����̎�
		If trim(gf_SetNull2String(p_sClubCd)) = "" Then
			Exit Do
		End If

		'//���������擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD=" & p_sClubCd

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//������
			w_sClubName = rs("M17_BUKATUDOMEI")
		End If

		Exit Do
	Loop

	'//�߂�l���
	f_GetClubName = w_sClubName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

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
	<title>�����������ꗗ</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//  [�@�\]  �y�[�W���[�h������
	//  [����]
	//  [�ߒl]
	//  [����]
	//************************************************************
	function window_onload() {
		<%If m_Rs.EOF = false Then%>
//			parent.topFrame.document.frm.btnShow.disabled=true;
//			parent.topFrame.document.frm.cboGakunenCd.disabled=true;
//			parent.topFrame.document.frm.cboClassCd.disabled=true;
//			parent.topFrame.document.frm.txtTyuClubCd.disabled=true;
		<%End If%>
	}
	//************************************************************
	//  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_Touroku(){

		// ���͒l������
		iRet = f_CheckData();
		if( iRet != 0 ){
			return;
		}

		if (!confirm("�o�^���Ă���낵���ł����H")) {
		return ;
		}

		//���X�g����submit
		document.frm.txtMode.value = "INSERT";
		document.frm.target = "main";
		document.frm.action = "./web0360_edt.asp"
		document.frm.submit();
		return;
	}

	//************************************************************
	//  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_Cancel(){

		//���ʎg�p�Ƃ���
//		parent.topFrame.document.frm.btnShow.disabled=false;
//		parent.topFrame.document.frm.cboGakunenCd.disabled=false;
//		parent.topFrame.document.frm.cboClassCd.disabled=false;
//		parent.topFrame.document.frm.txtTyuClubCd.disabled=false;

		//�󔒃y�[�W��\��
		parent.main.location.href="default3.asp?txtClubCd=<%=m_sClubCd%>"

	}

	//************************************************************
	//  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_Back(){
		//�L�����Z�����A������ʂɖ߂�
		//��t���[���ĕ\��
		parent.topFrame.location.href="./web0360_top.asp?txtClubCd=<%=m_sClubCd%>"
		//���t���[���ĕ\��
		parent.main.location.href="./web0360_main.asp?txtClubCd=<%=m_sClubCd%>"

	}

    //************************************************************
    //  [�@�\]  �`�F�b�N�����`�F�b�N����Ă��邩
    //  [����]  �Ȃ�
    //  [�ߒl]  0:����OK�A1:�����װ
    //************************************************************
    function f_CheckData() {

		//�`�F�b�N�������擾
		var iMax = document.frm.chkMax.value
		if (iMax==0){
			//alert("No Avairable")
			return 1;
		}

		if(iMax==1){
			if(document.frm.GAKUSEI_NO.checked==false){
				alert("�o�^���鐶�k���I������Ă��܂���")
				return 1;
			}
		}else{

			var i
			var w_bCheck = 1
			for (i = 0; i < iMax; i++) {
				if(document.frm.GAKUSEI_NO[i].checked==true){
					w_bCheck = 0
					break;
				}
			};

			if(w_bCheck == 1){
				alert("�o�^���鐶�k���I������Ă��܂���")
				return 1;
			};
		};

        return 0;
    }

    //************************************************************
    //  [�@�\]  �ڍ׃{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_detail(pGAKUSEI_NO){

			url = "/cassist/gak/gak0310/kojin.asp?hidGAKUSEI_NO=" + pGAKUSEI_NO;
			w   = 700;
			h   = 630;

			wn  = "SubWindow";
			opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
			if (w > 0)
				opt = opt + ",width=" + w;
			if (h > 0)
				opt = opt + ",height=" + h;
			newWin = window.open(url, wn, opt);

//		document.frm.hidGAKUSEI_NO.value = pGAKUSEI_NO;
//		document.forms[0].submit();
    }
	//-->
	</SCRIPT>

	</head>
	<body LANGUAGE=javascript onload="return window_onload()">
	<center>
	<form name="frm" method="post">

	<%
	Do
		'//�f�[�^�Ȃ��̏ꍇ
		If m_Rs.EOF Then%>
			<br><br>
			<span class="msg">�Ώۃf�[�^�͑��݂��܂���B��������͂��Ȃ����Č������Ă��������B</span>
			<br><br>
			<!--
			<input class="button" type="button" onclick="javascript:f_Cancel();" value="�L�����Z��">
			&nbsp;&nbsp;&nbsp;&nbsp;
			-->
			<input class="button" type="button" onclick="javascript:f_Back();" value="�@�߁@��@">

			<%Exit Do
		End If
	%>

		<br>

		<table>
			<tr>
			<td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="�@�o�@�^�@"></td>
			<td ><input class="button" type="button" onclick="javascript:f_Cancel();"  value="�L�����Z��"></td>
			<td ><input class="button" type="button" onclick="javascript:f_Back();"    value="�@�߁@��@"></td>
			</tr>
		</table>

		<span class="msg">�������o�^���鐶�k��I�����A�u�o�^�v�{�^���������Ă��������B</span><br>
		<span class="msg">�����łɓ������Ă��鐶�k�͓o�^�ł��܂���B</span>

		<!--���X�g��-->
		<table >
			<tr><td valign="top">

				<table class=hyo border="1" bgcolor="#FFFFFF">

					<!--�w�b�_-->
					<tr>
						<th nowrap class="header" width="15"  align="center">�I<br>��</th>
						<th nowrap class="header" width="40"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
						<th nowrap class="header" width="120" align="center">����</th>
						<th nowrap class="header" width="100" align="center">������</th>
<!--
						<th nowrap class="header" width="100" align="center">���w����</th>
-->
					</tr>

			<%
			j = 0
			w_iCnt = INT(m_iRsCnt/2 + 0.9)
			Do until m_Rs.EOF 

				'//���ټ�Ă̸׽���Z�b�g
				Call gs_cellPtn(w_Class)
				i = i + 1
				%>
					<tr>
						<td nowrap class="<%=w_Class%>" width="15"  align="center" rowspan="2">

						<%
                        '//���łɃN���u1�ƃN���u2�Ƀf�[�^�����鐶�k�́A�N���u�̐V�K�o�^�͕s�Ƃ���
						'If gf_SetNull2String(m_Rs("T13_CLUB_1")) <> "" AND gf_SetNull2String(m_Rs("T13_CLUB_2")) <> "" Then
                        '//���łɃN���u1�ƃN���u2�Ƀf�[�^�����藼���Ƃ��������̐��k�́A�N���u�̐V�K�o�^�͕s�Ƃ��� 2001/12/11 �ɓ�
						If gf_SetNull2String(m_Rs("T13_CLUB_1")) <> "" AND gf_SetNull2String(m_Rs("T13_CLUB_2")) <> "" AND gf_SetNull2String(m_Rs("T13_CLUB_1_FLG")) = "1" AND gf_SetNull2String(m_Rs("T13_CLUB_2_FLG")) = "1" Then
                        %>
							<br>
						<%Else%>

							<%
							'//���łɓo�^�ΏۃN���u�ɏ������Ă��鐶�k�́A�X�V�s�Ƃ���
							'If (gf_SetNull2String(m_Rs("T13_CLUB_1")) = m_sClubCd) Or (gf_SetNull2String(m_Rs("T13_CLUB_2")) = m_sClubCd) Then
							If (gf_SetNull2String(m_Rs("T13_CLUB_1")) = m_sClubCd And gf_SetNull2String(m_Rs("T13_CLUB_1_FLG")) = "1") Or (gf_SetNull2String(m_Rs("T13_CLUB_2")) = m_sClubCd And gf_SetNull2String(m_Rs("T13_CLUB_2_FLG")) = "1") Then
							%>
								<br>
							<%Else
								j = j + 1
								%>
								<input type="checkbox" name="GAKUSEI_NO" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<%End If%>

						<%End If%>

						</td>
						<td nowrap class="<%=w_Class%>" width="40"  align="left"   rowspan="2"><%=m_Rs("T13_GAKUSEKI_NO")%><br></td>
						<td nowrap class="<%=w_Class%>" width="120" align="left"   rowspan="2"><a href="#" onClick="f_detail(<%=m_Rs("T13_GAKUSEI_NO")%>)"><%=m_Rs("T11_SIMEI")%></a><br></td>
						<td class="<%=w_Class%>" width="100" align="left" >
							
							<%
							'�������Ȃ�\������
							If gf_SetNull2String(m_Rs("T13_CLUB_1_FLG")) = "1" Then 
							%>
								<%=gf_SetNull2Haifun(f_GetClubName(m_Rs("T13_CLUB_1")))%>
							<%
							End If
							%>
							<br>
						</td>

						</tr>
						<tr>
						<td nowrap class="<%=w_Class%>" width="100" align="left">
							<%
							'�������Ȃ�\������
							If gf_SetNull2String(m_Rs("T13_CLUB_2_FLG")) = "1" Then
							%>
								<%=gf_SetNull2Haifun(f_GetClubName(m_Rs("T13_CLUB_2")))%>
							<%
							End If
							%>
							<br>
						</td>
					</tr>

				<%If i =  w_iCnt And m_iRsCnt <> 1 Then
					'//���ټ�Ă̸׽��������
					w_Class = ""
				%>
				</table>
				</td>

				<td valign="top">
				<table class="hyo" border="1" >
					<!--�w�b�_-->
					<tr>
						<th nowrap class="header" width="15"  align="center">�I<br>��</th>
						<th nowrap class="header" width="40"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
						<th nowrap class="header" width="120" align="center">����</th>
						<th nowrap class="header" width="100" align="center">������</th>
<!--
						<th nowrap class="header" width="100" align="center">���w����</th>
-->
					</tr>
				<%End If%>

				<%m_Rs.MoveNext%>
			<%Loop%>

				</table>
				</td></tr>
			</table>
			<br>

			<table>
				<tr>
					<td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="�@�o�@�^�@"></td>
					<td ><input class="button" type="button" onclick="javascript:f_Cancel();" value="�L�����Z��"></td>
					<td ><input class="button" type="button" onclick="javascript:f_Back();" value=" �� �� �� "></td>
				</tr>
			</table>

		<%Exit Do%>
	<%Loop%>

	<!--�l�n���p-->
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
	<input type="hidden" name="txtClubCd" value="<%=m_sClubCd%>">
	<input type="hidden" name="chkMax"    value="<%=j%>">
	<input type="hidden" name="cboGakunenCd" value="<%=m_iGakunen%>">
	<input type="hidden" name="cboClassCd"   value="<%=m_iClassNo%>">
	<input type="hidden" name="txtTyuClubCd"   value="<%=m_sTyuClubCd%>">

	</form>
	</center>
	</body>
	</html>
<%
End Sub 
%>

