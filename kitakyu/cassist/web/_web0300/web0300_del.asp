<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ʋ����\��
' ��۸���ID : web/web0300/web0300_lst.asp
' �@      �\: ��������\��
'-------------------------------------------------------------------------
' ��      ��:   NENDO           '//�N�x
'               KYOKAN_CD       '//����CD
'				txtMode			:�������[�h
'				hidJigen		:����
'				hidDay			:���ɂ�
'				hidYear			:�N
'				hidMonth		:��
'				hidKyositu		:����CD
'				hidKyosituName	:��������
'
' ��      �n:	txtMode			:�������[�h
'				hidJigen		:����
'				hidDay			:���ɂ�
'				hidYear			:�N
'				hidMonth		:��
'				hidKyositu		:����CD
'				hidKyosituName	:��������
' ��      ��:
'           �������\��
'               �����I�����ꂽ�f�[�^�ꗗ��\��
'-------------------------------------------------------------------------
' ��      ��: 2001/08/08 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	Public m_iSyoriNen          '//�N�x
	Public m_iKyokanCd          '//��������

	Public m_sYear   			'//�N
	Public m_sMonth				'//��
	Public m_sDay    			'//��
	Public m_iKyosituCd			'//����CD
	Public m_iKaijyoCnt			'//�����`�F�b�N�{�b�N�X�J�E���g
	Public m_sMode				'//�������[�h
	Public m_iJigen				'//����
	Public m_sMokuteki			'//�ړI
	Public m_sBiko				'//���l
	Public m_sKyosituName		'//��������

    'ں��ރZ�b�g
    Public m_Rs					'//ں��޾��

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

    Dim w_iRet              '// �߂�l

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���ʋ����\��"
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

		Select Case m_sMode

			Case "DISP"

				'//�\���p�f�[�^�擾
				w_iRet = f_GetDispData()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//��ʂ�\��
				Call showPage()

			Case "DELETE"
				'//�f�[�^Delete
				w_iRet = f_DeleteData()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//�폜����I����
				Call showWhitePage()

			Case Else

		End Select

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
    m_sYear      = ""
    m_sMonth     = ""
    m_sDay       = ""
	m_iKyosituCd = ""
	m_sMode      = ""
	m_iJigen     = ""
	m_sMokuteki  = ""
	m_sBiko      = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen  = Session("NENDO")
    'm_iKyokanCd  = Session("KYOKAN_CD")
	''m_iKyokanCd  = Request("YoyakKyokanCd")
    m_iKyokanCd  = Request("SKyokanCd1")

    m_sYear      = Request("hidYear")
    m_sMonth     = Request("hidMonth")
    m_sDay       = Request("hidDay")
	m_iKyosituCd = Request("hidKyositu")
	m_sMode      = Request("txtMode")
	m_iJigen     = Request("hidJigen")
	m_sKyosituName = Request("hidKyosituName")

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_sMode      = " & m_sMode      & "<br>"
    response.write "m_iSyoriNen  = " & m_iSyoriNen  & "<br>"
    response.write "m_iKyokanCd  = " & m_iKyokanCd  & "<br>"
    response.write "m_sYear      = " & m_sYear      & "<br>"
    response.write "m_sMonth     = " & m_sMonth     & "<br>"
    response.write "m_sDay       = " & m_sDay       & "<br>"
    response.write "m_iKyosituCd = " & m_iKyosituCd & "<br>"
    response.write "m_iJigen     = " & m_iJigen     & "<br>"
    response.write "m_sMokuteki  = " & m_sMokuteki  & "<br>"
    response.write "m_sBiko      = " & m_sBiko      & "<br>"
    response.write "m_sKyosituName= " & m_sKyosituName & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �\���f�[�^���擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDispData()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetDispData = 1

    Do
		'//���t���쐬
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'//�����\��f�[�^�擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & " T58_HIDUKE, "
		w_sSql = w_sSql & vbCrLf & " T58_JIGEN, "
		w_sSql = w_sSql & vbCrLf & " T58_MOKUTEKI, "
		w_sSql = w_sSql & vbCrLf & " T58_BIKO"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & " T58_KYOSITU_YOYAKU"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & " T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND T58_HIDUKE='" & w_sDate & "'"
		w_sSql = w_sSql & vbCrLf & " AND T58_JIGEN IN (" & replace(Request("chkKaijyo")," ","") & ")"
		w_sSql = w_sSql & vbCrLf & " AND T58_KYOSITU=" & m_iKyosituCd
		'w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_iKyokanCd & "'"
		w_sSql = w_sSql & vbCrLf & " ORDER BY T58_JIGEN"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetDispData = 99
            Exit Do
        End If

        '//����I��
        f_GetDispData = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �f�[�^UPDATE
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_DeleteData()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_DeleteData = 1

    Do
		'//���t���쐬
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'//DELETE
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " DELETE T58_KYOSITU_YOYAKU "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & " T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND T58_HIDUKE='" & w_sDate & "'"
		w_sSql = w_sSql & vbCrLf & " AND T58_JIGEN IN (" & replace(Request("chkKaijyo")," ","") & ")"
		w_sSql = w_sSql & vbCrLf & " AND T58_KYOSITU=" & m_iKyosituCd
		'w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_iKyokanCd & "'"

'response.write w_sSQL & "<br>"

		iRet = gf_ExecuteSQL(w_sSQL)
		If iRet <> 0 Then
			'�폜���s
			msMsg = Err.description
			f_DeleteData = 99
			Exit Do
		End If

		'//����I��
		f_DeleteData = 0
		Exit Do
	Loop

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
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>���ʋ����\��</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
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
    //  [�@�\]  �L�����Z���{�^���N���b�N
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_Cancel() {

		document.frm.action="web0300_lst.asp";
		document.frm.target="bottom";
		document.frm.submit();
    }

    //************************************************************
    //  [�@�\]  �폜�{�^���N���b�N
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_Delete(){

        if (!confirm("�\����������܂��B��낵���ł����H")) {
           return ;
        }

		document.frm.txtMode.value="DELETE";
		document.frm.action="web0300_del.asp";
		document.frm.target="bottom";
		document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

<%
'//�f�o�b�O
'Call s_DebugPrint()
%>

	<center>
	<!--<form action="yoyaku.asp">-->
	<img src="img/sp.gif" height="3">

		<table border="1" class="hyo" width="700">
			<tr>
			<th CLASS="header" width="100" nowrap>���t</th>
			<th CLASS="header" width="100" nowrap>����</th>
			<th CLASS="header" width="60" nowrap>����</th>
			<th CLASS="header" width="200">�g�p�ړI</th>
			<th CLASS="header" width="200">���l</th>
			</tr>
			<%Do Until m_Rs.EOF%>
				<tr>
				<td class="detail" width="100" align="center" nowrap><%=gf_fmtWareki(m_Rs("T58_HIDUKE"))%><BR></td>
				<td class="detail" width="100" align="center" nowrap><%=m_sKyosituName%><BR></td>
				<td class="detail" width="60"  align="center" nowrap><%=m_Rs("T58_JIGEN")%>����</td>
				<td class="detail" width="200" ><%=m_Rs("T58_MOKUTEKI")%><BR></td>
				<td class="detail" width="200" ><%=m_Rs("T58_BIKO")%><BR></td>
				</tr>

				<%m_Rs.MoveNext%>
			<%Loop%>

		</table>

		<br>

		<table width="250">
			<tr>
			<td align="center" colspan="2"><font size="2">�ȏ�̗\����������܂��B</font></td>
			</tr>
			<tr>
			<td align="center"><input class="button" type="button" value="�@���@���@" onclick="javascript:f_Delete()"></td>
			<td align="center"><input class="button" type="button" value="�L�����Z��" onclick="javascript:f_Cancel()"></td>
			</tr>
		</table>

	<!--�l�n���p-->
	<input type="hidden" name="txtMode"    value="">
	<input type="hidden" name="chkKaijyo"  value="<%=Request("chkKaijyo")%>">
	<input type="hidden" name="SKyokanCd1"    value="<%=m_iKyokanCd%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">

	<input type="hidden" name="hidDay"     value="<%=m_sDay%>">
	<input type="hidden" name="hidYear"    value="<%=m_sYear %>">
	<input type="hidden" name="hidMonth"   value="<%=m_sMonth%>">
	<input type="hidden" name="hidKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="hidKyosituName" value="<%=m_sKyosituName%>">

	</form>
	</center>
	</body>
	</html>

<%
End Sub

'********************************************************************************
'*  [�@�\]  �󔒃y�[�W
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showWhitePage()
%>
    <html>
    <head>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>���ʋ����\��</title>


    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		alert("�\����������܂���")

		var wArg
		//�J�����_�[�y�[�W���ĕ\��
		wArg = ""
		wArg = wArg + "?TUKI=<%=m_sMonth%>"
		wArg = wArg + "&cboKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidDay=<%=m_sDay%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=m_iKyokanCd%>"

		//�J�����_�[�y�[�W���ĕ\��
		parent.middle.location.href="./calendar.asp"+wArg
		//parent.middle.location.href="./calendar.asp?TUKI=<%=m_sMonth%>&cboKyositu=<%=m_iKyosituCd%>&hidDay=<%=m_sDay%>"

		//���X�g�y�[�W���ĕ\��
		wArg = ""
		wArg = wArg + "?hidDay=<%=m_sDay%>"
		wArg = wArg + "&hidYear=<%=m_sYear%>"
		wArg = wArg + "&hidMonth=<%=m_sMonth%>"
		wArg = wArg + "&hidKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidKyosituName=<%=Server.URLEncode(m_sKyosituName)%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=m_iKyokanCd%>"

		parent.bottom.location.href="./web0300_lst.asp"+wArg

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type="hidden" name="TUKI"       value="<%=m_sMonth%>">
	<input type="hidden" name="cboKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="SKyokanCd1" value="<%=m_iKyokanCd%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">

	<input type="hidden" name="hidDay"     value="<%=m_sDay%>">
	<input type="hidden" name="hidYear"    value="<%=m_sYear %>">
	<input type="hidden" name="hidMonth"   value="<%=m_sMonth%>">
	<input type="hidden" name="hidKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="hidKyosituName" value="<%=m_sKyosituName%>">

	</form>
	</body>
	</html>
<%
End Sub
%>