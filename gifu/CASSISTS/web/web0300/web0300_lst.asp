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
'				hidDay    		:���ɂ�
'				hidYear    		:�N
'				hidMonth   		:��
'				hidKyositu 		:����CD
'
' ��      �n:	txtMode			:�������[�h
'				hidJigen		:����
'				YoyakKyokanCd	:�\�񋳊�CD
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
' ��      ��: 2001/07/19 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
'
'	Const C_ACCESS_FULL   = "FULL"		'//�A�N�Z�X����FULL�A�N�Z�X��
'	Const C_ACCESS_NORMAL = "NORMAL"	'//�A�N�Z�X�������
'	Const C_ACCESS_VIEW   = "VIEW"		'//�A�N�Z�X�����Q�Ƃ̂�

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	Public m_iSyoriNen			'//�N�x
	Public m_iKyokanCd			'//��������

	Public m_sYear   			'//�N
	Public m_sMonth			  	'//��
	Public m_sDay   			'//��
	Public m_iKyosituCd			'//����CD
	Public m_iKaijyoCnt			'//�����`�F�b�N�{�b�N�X�J�E���g
	Public m_iYoyakCnt			'//�\��`�F�b�N�{�b�N�X�J�E���g
	Public m_sKyosituName		'//��������

	Public m_sUserId

    'ں��ރZ�b�g
    Public m_Rs_Jigen       	'//����ں��޾��
    Public m_Rs_Kyositu			'//�����\����

    Public m_bUpdate_OK			'//�\��A�X�V�s�����׸�
    Public m_sKengen

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

		'//�������擾
		w_iRet = gf_GetKengen_web0300(m_sKengen)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'//�������A�\�����e��ς���
        Call s_SetViewInfo()

'//�f�o�b�O
'Call s_DebugPrint()

		'//�������擾
		w_iRet = f_GetKyousituName()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//�������̎擾
        w_iRet = f_GetJigen()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '// �����\��󋵂̎擾
        w_iRet = f_GetKyosituInfo()
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
	
	m_sUserId = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen  = Session("NENDO")
    'm_iKyokanCd  = Request("SKyokanCd1")
   'm_iKyokanCd  = SESSION("KYOKAN_CD")
    m_sYear      = Request("hidYear")
    m_sMonth     = Request("hidMonth")
    m_sDay       = Request("hidDay")
	m_iKyosituCd = Request("hidKyositu")

'	m_sUserId    = SESSION("LOGIN_ID")
	m_iKyokanCd  = SESSION("LOGIN_ID")

End Sub

'********************************************************************************
'*  [�@�\]  �������A�\�����e��ύX����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetViewInfo()

	m_bUpdate_OK = False

	'//�Q�Ƃ̂݉\�ȏꍇ
	If m_sKengen = C_ACCESS_VIEW Then
		m_bUpdate_OK = False
	Else
		'//������FULL�A�N�Z�X�܂��́A��ʂ̏ꍇ
		m_bUpdate_OK = True
	End If

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen  = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd  = " & m_iKyokanCd & "<br>"
    response.write "m_sYear      = " & m_sYear     & "<br>"
    response.write "m_sMonth     = " & m_sMonth    & "<br>"
    response.write "m_sDay       = " & m_sDay      & "<br>"
    response.write "m_iKyosituCd = " & m_iKyosituCd      & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKyousituName()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetKyousituName = 1

    Do
		'//�������擾
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU.M06_KYOSITUMEI"
		w_sSql = w_sSql & vbCrLf & " FROM M06_KYOSITU"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU.M06_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M06_KYOSITU.M06_KYOSITU_CD=" & m_iKyosituCd

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKyousituName = 99
            Exit Do
        End If

		If rs.EOF = False Then
			m_sKyosituName = rs("M06_KYOSITUMEI")
		End If

        '//����I��
        f_GetKyousituName = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �������̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetJigen()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen = 1

    Do

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M07_JIKAN"
		w_sSql = w_sSql & vbCrLf & " FROM M07_JIGEN"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "      M07_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " GROUP BY M07_JIKAN"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs_Jigen, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetJigen = 99
            Exit Do
        End If

        '//����I��
        f_GetJigen = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �����\��󋵂̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKyosituInfo()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetKyosituInfo = 1

    Do

		w_sDate = m_sYear & "/" & gf_fmtZero(m_sMonth,2) & "/" &  gf_fmtZero(m_sDay,2)

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "   T58.T58_JIGEN"
		w_sSql = w_sSql & vbCrLf & "  ,T58.T58_KYOKAN_CD"
		w_sSql = w_sSql & vbCrLf & "  ,T58.T58_MOKUTEKI"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T58_KYOSITU_YOYAKU T58"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T58.T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND T58.T58_HIDUKE='" & w_sDate & "' "
		w_sSql = w_sSql & vbCrLf & "  AND T58.T58_KYOSITU=" & m_iKyosituCd

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(m_Rs_Kyositu, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKyosituInfo = 99
            Exit Do
        End If

        '//����I��
        f_GetKyosituInfo = 0
        Exit Do
    Loop

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
'*  [�@�\]  �\�񋳎��f�[�^��\������
'*  [����]  p_Jigen	�F����
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_KyousituData(p_Jigen,p_sClass)

	Dim w_sJigen
	Dim w_sMokuteki
	Dim w_sTourokusya
	Dim w_sTourokuCD

	w_sMokuteki = ""
	w_sTourokusya = ""
	w_sTourokuCD = ""
	w_btnUpdate = "<br>"
	w_chkKaijyo ="<br>"

	w_bYoyak = False

	Do

		If m_Rs_Kyositu.EOF = false then

			Do Until m_Rs_Kyositu.EOF

				'//�擾���������ɋ����\�񂪓����Ă��邩
				If clng(p_Jigen) = clng(m_Rs_Kyositu("T58_JIGEN")) Then

					w_bYoyak = True
					w_sJigen      = p_Jigen
					w_sMokuteki   = "<A href='javascript:f_LinkClick(" & p_Jigen & ");'>" & m_Rs_Kyositu("T58_MOKUTEKI")  & "</A>"
					w_sTourokusya = f_GetName(m_Rs_Kyositu("T58_KYOKAN_CD"))

						'//�A�N�Z�X�������Q�ƈȊO�̏ꍇ�A�{�l�֘A�̃f�[�^�̏C���E�폜���\�ƂȂ�
						If m_sKengen <> C_ACCESS_VIEW Then

							'//���݂̗��p�҂Ɠo�^����Ă��闘�p�҂������ꍇ�͏C���{�^���y�щ����`�F�b�N�{�b�N�X��\��
'Response.Write "kyoukan=[" & m_Rs_Kyositu("T58_KYOKAN_CD") & "]" & "[" & m_iKyokanCd & "]"
							If NOT ISNULL(m_Rs_Kyositu("T58_KYOKAN_CD")) AND NOT ISNULL(m_iKyokanCd) Then
								If cstr(m_Rs_Kyositu("T58_KYOKAN_CD")) = cstr(m_iKyokanCd) Or cstr(m_Rs_Kyositu("T58_KYOKAN_CD")) = m_sUserId Then
									w_chkKaijyo = "<input type='checkbox' name='chkKaijyo' value='" & p_Jigen & "' >"
									w_btnUpdate = "<input type='button'   name='btnUpdate' value='�C��' class='button' onclick='javascript:f_UpdClick(" & p_Jigen & ")'>"
									m_iKaijyoCnt = m_iKaijyoCnt + 1
								End If
							End If

						End If

					Exit Do

				End If

				m_Rs_Kyositu.MoveNext
			Loop

			m_Rs_Kyositu.MoveFirst
		End If

		Exit Do
	Loop

	'//���łɗ\�񂳂�Ă��邩
	If w_bYoyak = True Then
		'//�\�񂪂���Ƃ��\��{�^����\��
		w_btnYoyak = "<br>"
	Else 
		'//�\�񂪂Ȃ����͗\��{�^���\��
		w_btnYoyak  = "<input type='checkbox' name='hidYoyak' & value='" & trim(p_Jigen) & "'>"
		w_sMokuteki = "��</font>"
		w_sTourokusya = "�\"
		w_sJigen   = p_Jigen
		m_iYoyakCnt = m_iYoyakCnt + 1
	End If

	%>
	<td class="<%=p_sClass%>" align="center" height="25" width="50" nowrap><%=w_sJigen%></td>
	<td class="<%=p_sClass%>" align="left"><%=w_sMokuteki%></td>
	<td class="<%=p_sClass%>" align="center" nowrap><%=w_sTourokusya%></td>

	<%'//�����ɂ��\���𐧌�%>

	<%If m_bUpdate_OK = True then%><td class="<%=p_sClass%>" align="center" nowrap><%=w_btnYoyak%></td><%End If%>
	<%If m_bUpdate_OK = True then%><td class="<%=p_sClass%>" align="center" nowrap><%=w_btnUpdate%></td><%End If%>
	<%If m_bUpdate_OK = True then%><td class="<%=p_sClass%>" align="center" nowrap><%=w_chkKaijyo%></td><%End If%>

<%
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
    //  [�@�\]  �����{�^���N���b�N
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_KaijyoClick() {

		//�`�F�b�N�������擾
		var iMax = document.frm.hidKaijyoCnt.value
		if (iMax==0){
			//alert("No Avairable")
			return;
		}

		if(iMax==1){
			if(document.frm.chkKaijyo.checked==false){
				alert("��������f�[�^���I������Ă��܂���")
				return;
			}
		}else{

			var i
			var w_bCheck = 1
			for (i = 0; i < iMax; i++) {
				if(document.frm.chkKaijyo[i].checked==true){
					w_bCheck = 0
					break;
				}
			};

			if(w_bCheck == 1){
				alert("��������f�[�^���I������Ă��܂���")
				return;
			};
		};

		document.frm.txtMode.value="DISP";
		document.frm.action="web0300_del.asp";
		document.frm.target="bottom";
		document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �C���{�^���N���b�N
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_UpdClick(p_Jigen){

		// document.frm.YoyakKyokanCd.value="imawaka"

		document.frm.hidJigen.value=p_Jigen;
		document.frm.txtMode.value="DETAIL";
		document.frm.action="web0300_detail.asp";
		document.frm.target="bottom";
		document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �����N�N���b�N
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_LinkClick(p_Jigen){

		// document.frm.YoyakKyokanCd.value=p_sKyokanCd;

		document.frm.hidJigen.value=p_Jigen;
		document.frm.txtMode.value="DISP";
		document.frm.action="web0300_detail.asp";
		document.frm.target="bottom";
		document.frm.submit();
    }

    //************************************************************
    //  [�@�\]  �\��{�^���N���b�N
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
	function f_btnYoyakClick(){

		//�`�F�b�N�������擾
		var iMax = document.frm.hidYoyakCnt.value
		if (iMax==0){
			//alert("No Avairable")
			return;
		}

		//�`�F�b�N�{�b�N�X���I������Ă��邩�`�F�b�N
		//�I������Ă����hidJigen�Ɋi�[
		if(iMax==1){
			if(document.frm.hidYoyak.checked==false){
				alert("�\�񂷂鎞�����I������Ă��܂���")
				return;
			}else{
				document.frm.hidJigen.value = document.frm.hidYoyak.value
			};
		}else{

			var i
			for (i = 0; i < iMax; i++) {
				if(document.frm.hidYoyak[i].checked==true){
					if(document.frm.hidJigen.value==""){
						document.frm.hidJigen.value = document.frm.hidYoyak[i].value
					}else{
						document.frm.hidJigen.value = document.frm.hidJigen.value+","+document.frm.hidYoyak[i].value
					};

				};
			};

			if(document.frm.hidJigen.value==""){
				alert("�\�񂷂鎞�����I������Ă��܂���")
				return;
			};
		};

		document.frm.txtMode.value="BLANK";
		document.frm.action="web0300_detail.asp";
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
	<img src="img/sp.gif" height="3">
	<%Do%>

		<table border="1" width="98%" class="hyo">
		<tr>
		<th class=header width="50">����</th>
		<th class=header>�g�p�ړI</th>
		<th class=header>���p��</th>

		<%If m_bUpdate_OK = True then%><th class=header>�\��</th><%End If%>
		<%If m_bUpdate_OK = True then%><th class=header>�C��</th><%End If%>
		<%If m_bUpdate_OK = True then%><th class=header>����</th><%End If%>

		</tr>

		<%

		'//�����`�F�b�N�{�b�N�X�J�E���g��������
		m_iKaijyoCnt = 0

		'//�\��`�F�b�N�{�b�N�X�J�E���g��������
		m_iYoyakCnt = 0

		Do Until m_Rs_Jigen.EOF%>

			<%if f_LenB(m_Rs_Jigen("M07_JIKAN")) < 3 then %>

				<tr>
				<%
				'//���ټ�Ă̸׽���Z�b�g
				Call gs_cellPtn(w_Class)

				'//�ڍ׃f�[�^�\��
				Call f_KyousituData(m_Rs_Jigen("M07_JIKAN"),w_Class)
				%>
				</tr>

			<%End If%>

			<%m_Rs_Jigen.MoveNext%>
		<%Loop%>

	    <tr>
		<%If m_bUpdate_OK = True then%>
		    <td colspan="4" align=right bgcolor=#9999BD>
				<input class=button type=button value="�\��" onclick="javascript:f_btnYoyakClick()">
			</td>
		<%End If%>


		<%If m_bUpdate_OK = True then%>
		    <td colspan="2" align=right bgcolor=#9999BD>
				<input class=button type=button value="����" onclick="javascript:f_KaijyoClick()">
			</td>
		<%End If%>

	    </tr>

		</table>

		<table width="98%" border=0>
			<tr>
			<td align="right">
				<span class="msg"><font size="2">

				<%If m_sKengen = C_ACCESS_VIEW Then%>
					���\����̏ڍׂ́A�g�p�ړI���N���b�N����Ɗm�F�ł��܂��B
				<%Else%>
					�����łɗ\�񂳂�Ă��鎞���ɂ͗\��ł��܂���B&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
					�\����̏ڍׂ́A�g�p�ړI���N���b�N����Ɗm�F�ł��܂��B<br>
					���C���E�����͓o�^�҂̂݉\�ł��B&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%End If%>

				</font></span>
			</td>
		<%Exit Do%>
	<%Loop%>

	<!--�l�n�p-->
	<input type="hidden" name="hidYoyakCnt"    value="<%=m_iYoyakCnt%>">
	<input type="hidden" name="hidKaijyoCnt"   value="<%=m_iKaijyoCnt%>">
	<input type="hidden" name="txtMode"        value="">
	<input type="hidden" name="hidJigen"       value="">
	<input type="hidden" name="YoyakKyokanCd"  value="">
	<input type="hidden" name="SKyokanNm1"     value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">
	<input type="hidden" name="SKyokanCd1"     value="<%=m_iKyokanCd%>">

	<input type="hidden" name="hidDay"         value="<%=m_sDay%>">
	<input type="hidden" name="hidYear"        value="<%=m_sYear %>">
	<input type="hidden" name="hidMonth"       value="<%=m_sMonth%>">
	<input type="hidden" name="hidKyositu"     value="<%=m_iKyosituCd%>">
	<input type="hidden" name="hidKyosituName" value="<%=m_sKyosituName%>">

	</form>
	</center>
	</body>
	</html>

<%
End Sub
%>
