<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ʋ����\��
' ��۸���ID : web/web0300/web0300_lst.asp
' �@      �\: ��������\��
'-------------------------------------------------------------------------
' ��      ��:   NENDO           '//�N�x
'				YoyakKyokanCd	:�\�񋳊�CD
'				txtMode			:�������[�h
'				hidJigen		:����
'				hidDay			:���ɂ�
'				hidYear			:�N
'				hidMonth		:��
'				hidKyositu		:����CD
'				hidKyosituName	:��������
' ��      �n:
'				YoyakKyokanCd	:�\�񋳊�CD
'				txtMode			:�������[�h
'				hidJigen		:����
'				hidDay			:���ɂ�
'				hidYear			:�N
'				hidMonth		:��
'				hidKyositu		:����CD
'				hidKyosituName	:��������
' ��      ��:
'           �������\��
'               �󔒃y�[�W��\��
'           ���\���{�^���������ꂽ�ꍇ
'               ���������ɂ��Ȃ����������Ԋ���\��
'-------------------------------------------------------------------------
' ��      ��: 2001/08/07 �ɓ����q
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

	Public m_sUserId

	Public m_iKyokanCdUpd

    'ں��ރZ�b�g
    Public m_Rs_Jigen           '//����ں��޾��
    Public m_Rs_Kyositu			'//�����\����

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
Call s_DebugPrint()


		'//�������[�h�ɂ�菈����U�蕪����
		Select Case m_sMode

			'//�V�K�o�^���͗p�t�H�[���\��
			Case "BLANK"
				'//�V�K�o�^��ʂ�\��
				Call showPage()

			'//�C���{�^���N���b�N��,�C����ʂ�\��
			Case "DETAIL"

				'//�\���p�f�[�^�擾
				w_iRet = f_GetDispData()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//��ʂ�\��
				Call showPage()

			'//�����N�N���b�N����ʂ�\��
			Case "DISP"

				'//�\���p�f�[�^�擾
				w_iRet = f_GetDispData()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//��ʂ�\��
				Call showPage()

			'//�V�K�o�^����(�����[�h��)
			Case "INSERT"
				'//�f�[�^INSERT
				w_iRet = f_DataInsert()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//�o�^����I����
				Call showWhitePage(C_TOUROKU_OK_MSG)

			'//�X�V����(�����[�h��)
			Case "UPDATE"
				'//�f�[�^UPDATE
				w_iRet = f_DataUpdate()
				If w_iRet <> 0 Then
					m_bErrFlg = True
					Exit Do
				End If

				'//�X�V����I����
				Call showWhitePage(C_UPDATE_OK_MSG)

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
    Call gf_closeObject(m_Rs_Jigen)

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

	m_sUserId = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen    = Session("NENDO")
	'm_iKyokanCdUpd = Request("YoyakKyokanCd")
	m_iKyokanCdUpd = trim(Request("SKyokanCd1"))
	m_iKyokanCd    = trim(session("KYOKAN_CD"))
    m_sYear        = Request("hidYear")
    m_sMonth       = Request("hidMonth")
    m_sDay         = Request("hidDay")
	m_iKyosituCd   = Request("hidKyositu")
	m_sMode        = Request("txtMode")
	m_iJigen       = Request("hidJigen")
	m_sKyosituName = Request("hidKyosituName")

	m_sUserId = trim(Session("LOGIN_ID"))

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_sMode        = " & m_sMode      & "<br>"
    response.write "m_iSyoriNen    = " & m_iSyoriNen  & "<br>"
    response.write "m_iKyokanCd    = " & m_iKyokanCd  & "<br>"
    response.write "m_sYear        = " & m_sYear      & "<br>"
    response.write "m_sMonth       = " & m_sMonth     & "<br>"
    response.write "m_sDay         = " & m_sDay       & "<br>"
    response.write "m_iKyosituCd   = " & m_iKyosituCd & "<br>"
    response.write "m_iJigen       = " & m_iJigen     & "<br>"
    response.write "m_sMokuteki    = " & m_sMokuteki  & "<br>"
    response.write "m_sBiko        = " & m_sBiko      & "<br>"
    response.write "m_sKyosituName = " & m_sKyosituName & "<br>"

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

    f_GetJigen = 1

    Do
		'//���t���쐬
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'//�����\��f�[�^�擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & " T58_MOKUTEKI, "
		w_sSql = w_sSql & vbCrLf & " T58_KYOKAN_CD, "
		w_sSql = w_sSql & vbCrLf & " T58_BIKO"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & " T58_KYOSITU_YOYAKU"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & " T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND T58_HIDUKE='" & w_sDate & "'"
		w_sSql = w_sSql & vbCrLf & " AND T58_JIGEN=" & cint(m_iJigen)
		w_sSql = w_sSql & vbCrLf & " AND T58_KYOSITU=" & m_iKyosituCd
		'w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_iKyokanCd & "'"

response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetDispData = 99
            Exit Do
        End If

		If rs.EOF = False Then
			m_sMokuteki = rs("T58_MOKUTEKI")
			m_sBiko     = rs("T58_BIKO")
			m_iKyokanCd = rs("T58_KYOKAN_CD")
		End If

        '//����I��
        f_GetDispData = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �f�[�^INSERT
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_DataInsert()

    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_iJigen
	Dim w_iJikanCnt
	Dim w_sRiyosya
	Dim w_sMokuteki
	Dim w_sBiko
	Dim w_sInsUser

    On Error Resume Next
    Err.Clear

    f_DataInsert = 1

    Do
		'//���t���쐬
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'w_sRiyosya = trim(Request("SKyokanCd1"))
		'w_sRiyosya = trim(session("KYOKAN_CD"))
		w_sMokuteki = trim(Request("txtMokuteki"))
		w_sBiko = trim(Request("txtBiko"))
		'w_sInsUser = trim(Session("LOGIN_ID"))

        '//���������擾
        w_iJigen = split(replace(m_iJigen," ",""),",")
        w_iJikanCnt = UBound(w_iJigen)

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

		For i=0 To w_iJikanCnt

			'//INSERT
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & " INSERT INTO T58_KYOSITU_YOYAKU"
			w_sSql = w_sSql & vbCrLf & " ("
			w_sSql = w_sSql & vbCrLf & " T58_NENDO "
			w_sSql = w_sSql & vbCrLf & " ,T58_HIDUKE "
			w_sSql = w_sSql & vbCrLf & " ,T58_YOUBI_CD "
			w_sSql = w_sSql & vbCrLf & " ,T58_JIGEN "
			w_sSql = w_sSql & vbCrLf & " ,T58_KYOSITU "
			w_sSql = w_sSql & vbCrLf & " ,T58_KYOKAN_CD "
			w_sSql = w_sSql & vbCrLf & " ,T58_MOKUTEKI "
			w_sSql = w_sSql & vbCrLf & " ,T58_BIKO "
			w_sSql = w_sSql & vbCrLf & " ,T58_INS_DATE "
			w_sSql = w_sSql & vbCrLf & " ,T58_INS_USER"
			w_sSql = w_sSql & vbCrLf & " ) VALUES ("
			w_sSql = w_sSql & vbCrLf & " "   & m_iSyoriNen
			w_sSql = w_sSql & vbCrLf & " ,'" & w_sDate & "'"
			w_sSql = w_sSql & vbCrLf & " ,"  & Weekday(w_sDate)
			w_sSql = w_sSql & vbCrLf & " ,"  & cint(w_iJigen(i))
			w_sSql = w_sSql & vbCrLf & " ,"  & m_iKyosituCd
			'w_sSql = w_sSql & vbCrLf & " ,'" & w_sRiyosya  & "'"
	'//�����b�c���Ȃ��ꍇ�̓��[�U�[�h�c����͂��� 2002.1.8
	'If m_iKyokanCd <> "" then
	'		w_sSql = w_sSql & vbCrLf & " ,'" & m_iKyokanCd & "'"
	'Else
			w_sSql = w_sSql & vbCrLf & " ,'" & m_sUserId & "'"
	'End If
			w_sSql = w_sSql & vbCrLf & " ,'" & w_sMokuteki & "'"
			w_sSql = w_sSql & vbCrLf & " ,'" & w_sBiko     & "'"
			w_sSql = w_sSql & vbCrLf & " ,'" & gf_YYYY_MM_DD(date(),"/") & "'"
			w_sSql = w_sSql & vbCrLf & " ,'" & w_sInsUser & "'"
			w_sSql = w_sSql & vbCrLf & " )"

response.write w_sSQL & "<br>"
response.end

			iRet = gf_ExecuteSQL(w_sSQL)
			If iRet <> 0 Then
                '//۰��ޯ�
                Call gs_RollbackTrans()
				'�o�^���s
				f_DataInsert = 99
				Exit Do
			End If

		Next

        '//�Я�
        Call gs_CommitTrans()

        '//����I��
        f_DataInsert = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �f�[�^UPDATE
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_DataUpdate()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_DataUpdate = 1

    Do
		'//���t���쐬
		w_sDate = gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/")

		'//UPDATE
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " UPDATE T58_KYOSITU_YOYAKU SET"
		w_sSql = w_sSql & vbCrLf & "  T58_MOKUTEKI='" & trim(Request("txtMokuteki")) & "'"
		w_sSql = w_sSql & vbCrLf & " ,T58_BIKO='" & trim(Request("txtBiko")) & "'"
		w_sSql = w_sSql & vbCrLf & " ,T58_UPD_DATE='" & gf_YYYY_MM_DD(date(),"/") & "'"
		w_sSql = w_sSql & vbCrLf & " ,T58_UPD_USER='" & Session("LOGIN_ID") & "'"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & " T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " AND T58_HIDUKE='" & w_sDate & "'"
		w_sSql = w_sSql & vbCrLf & " AND T58_JIGEN=" & cint(m_iJigen)
		w_sSql = w_sSql & vbCrLf & " AND T58_KYOSITU=" & cint(m_iKyosituCd)
		
		If m_iKyokanCd <> "" then
			w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_iKyokanCd & "'"
		Else
			w_sSql = w_sSql & vbCrLf & " AND T58_KYOKAN_CD='" & m_sUserId & "'"
		End if

'response.write w_sSQL & "<br>"
'response.end

		iRet = gf_ExecuteSQL(w_sSQL)
		If iRet <> 0 Then
			'�o�^���s
			msMsg = Err.description
			f_DataUpdate = 99
			Exit Do
		End If

		'//����I��
		f_DataUpdate = 0
		Exit Do
	Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  ���p�Җ����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  T58_KYOKAN_CD�ɂ́A����CD��USERID(M10)�̂ǂ��炩�������Ă���̂ŁA
'*          �͂��߂ɁA�����}�X�^�����������̂��擾�ł��Ȃ������ꍇ��USER�}�X�^���݂�
'********************************************************************************
Function f_GetName(p_sUserId)
    Dim w_iRet
	Dim w_sUserName

    On Error Resume Next
    Err.Clear

    f_GetName = ""
	w_sUserName = ""

    Do

		'//�����}�X�^���A���������擾����
		w_sUserName = gf_GetKyokanNm(m_iSyoriNen,p_sUserId)

		'//�������̂��擾�ł��Ȃ������ꍇ
		If Trim(w_sUserName) = "" Then
			'//USER�}�X�^���AUSER�����擾����
			w_sUserName = gf_GetUserNm(m_iSyoriNen,p_sUserId)
		End If

        Exit Do
    Loop

    f_GetName = w_sUserName

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
    //  [�@�\]  �o�^�{�^���N���b�N
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_Touroku(){

		// ���͒l������
		iRet = f_CheckData();
		if( iRet != 0 ){
			return;
		}

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		<%If m_sMode="BLANK" Then%>
			document.frm.txtMode.value="INSERT";
		<%Else%>
			document.frm.txtMode.value="UPDATE";
		<%End If%>

		document.frm.action="web0300_detail.asp";
		document.frm.target="bottom";
		document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  ���͒l������
    //  [����]  �Ȃ�
    //  [�ߒl]  0:����OK�A1:�����װ
    //  [����]  ���͒l��NULL�����A�p���������A�����������s��
    //          ���n�ް��p���ް������H����K�v������ꍇ�ɂ͉��H���s��
    //************************************************************
    function f_CheckData() {
    
		// ������NULL����������
		// ���ړI
		//if( f_Trim(document.frm.txtMokuteki.value) == "" ){
		//	window.alert("�ړI�����͂���Ă��܂���");
		//	document.frm.txtMokuteki.focus();
		//	return 1;
		//}

//		// �����l
//		if( f_Trim(document.frm.txtBiko.value) == "" ){
//			window.alert("���l�����͂���Ă��܂���");
//			document.frm.txtBiko.focus();
//			return 1;
//		}

		// ����������������������
		// ���ړI
		if( getLengthB(document.frm.txtMokuteki.value) > "50" ){
			window.alert("�ړI�͑S�p25�����ȓ��œ��͂��Ă�������");
			document.frm.txtMokuteki.focus();
			return 1;
		}

		// �����l
		if( getLengthB(document.frm.txtBiko.value) > "200" ){
			window.alert("���l�͑S�p100�����ȓ��œ��͂��Ă�������");
			document.frm.txtBiko.focus();
			return 1;
		}

        return 0;
    }


    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post" onSubmit="return false">

<%
'//�f�o�b�O
'Call s_DebugPrint()
%>
<br>
	<center>


		<table border="1" class="hyo" width="98%">

		<tr>
		<th CLASS="header" width="90"  nowrap>���t</th>
		<td class="detail" ><%=gf_fmtWareki(gf_YYYY_MM_DD(m_sYear & "/" & m_sMonth & "/" &  m_sDay,"/"))%><BR></td>
		</tr>

		<tr>
		<th CLASS="header" width="90"  nowrap>����</th>
		<%
		If m_sMode="BLANK" Then
			w_sJigen =replace(replace(m_iJigen," ",""),",","�����A") & "����"
		Else
			w_sJigen = m_iJigen & "������"
		End If
		%>
		<td class="detail"><%=w_sJigen%></td>
		</tr>

		<%
		'//�\���݂̂̎�
		If m_sMode="DISP" Then%>
			<tr>
			<th CLASS="header" width="90"  nowrap>���p��</th>
			<td class="detail" ><%=f_GetName(m_iKyokanCd)%><BR></td>
			</tr>
		<%End If%>

		<tr>
		<th CLASS="header" width="90" nowrap>����</th>
		<td class="detail"><%=m_sKyosituName%><BR></td>
		</tr>

		<tr>
		<th CLASS="header" width="90" nowrap>�g�p�ړI</th>
		<%
		'//�\���݂̂̎�
		If m_sMode="DISP" Then%>
			<td class="detail" height="20"><%=m_sMokuteki%><BR></td>
		<%Else%>
			<td class="detail"><input type="text" name="txtMokuteki" value="<%=m_sMokuteki%>" maxlength="50" size="70">
			</td>
		<%End If%>

		</tr>

		<tr>
		<th CLASS="header" width="90" nowrap>���l</th>
		<%
		'//�\���݂̂̎�
		If m_sMode="DISP" Then%>
				<td class="detail" height="40" valign="top"><%=m_sBiko%><BR></td>
			</tr>
			</table>

		<%Else%>
				<td class="detail"><textarea rows="4" cols="50" WRAP="soft" class="text" name="txtBiko" ><%=m_sBiko%></textarea>
				<br><font size=2>�i�S�p100�����ȓ��j</font>
				</td>
			</tr>
			</table>
		<%End If%>

		<br>

		<table width="250">
		<tr>
		<%If m_sMode="DISP" Then%>
			<td align="center"><input class="button" type="button" value="����" onclick="javascript:f_Cancel()"></td>
		<%Else%>
			<td align="center"><input class="button" type="button" value="�@�o�@�^�@" onclick="javascript:f_Touroku()"></td>
			<td align="center"><input class="button" type="button" value="�L�����Z��" onclick="javascript:f_Cancel()"></td>
		<%End If%>
		</tr>
		</table>

	<!--�l�n���p-->
	<input type="hidden" name="txtMode"       value="">
	<input type="hidden" name="hidJigen"      value="<%=m_iJigen%>">
	<input type="hidden" name="YoyakKyokanCd" value="<%=m_iKyokanCd%>">
	<input type="hidden" name="SKyokanCd1"    value="<%=Request("SKyokanCd1")%>">
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
Sub showWhitePage(p_sOkMsg)
%>
    <html>
    <head>
    <meta>
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
		alert("<%=p_sOkMsg%>")

		var wArg

		//�J�����_�[�y�[�W���ĕ\��
		wArg = ""
		wArg = wArg + "?TUKI=<%=m_sMonth%>"
		wArg = wArg + "&cboKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidDay=<%=m_sDay%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=Server.URLEncode(Request("SKyokanCd1"))%>"

		parent.middle.location.href="./calendar.asp"+wArg

		//���X�g�y�[�W���ĕ\��
		wArg = ""
		wArg = wArg + "?hidDay=<%=m_sDay%>"
		wArg = wArg + "&hidYear=<%=m_sYear%>"
		wArg = wArg + "&hidMonth=<%=m_sMonth%>"
		wArg = wArg + "&hidKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidKyosituName=<%=Server.URLEncode(m_sKyosituName)%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=Server.URLEncode(Request("SKyokanCd1"))%>"

		parent.bottom.location.href="./web0300_lst.asp"+wArg

    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type="hidden" name="TUKI"       value="<%=m_sMonth%>">
	<input type="hidden" name="cboKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">
	<input type="hidden" name="SKyokanCd1" value="<%=Server.HTMLEncode(Request("SKyokanCd1"))%>">

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