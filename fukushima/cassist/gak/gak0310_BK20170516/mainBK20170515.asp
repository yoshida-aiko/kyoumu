<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�����
' ��۸���ID : gak/gak0310/main.asp
' �@      �\: ���y�[�W �w�Ѓf�[�^�̌������ʂ�\������
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtHyoujiNendo         :�\���N�x
'           txtGakunen             :�w�N
'           txtGakkaCD             :�w��
'           txtClass               :�N���X
'           txtName                :����
'           txtGakusekiNo          :�w�Дԍ�
'           txtSeibetu             :����
'           txtGakuseiNo           :�w���ԍ�
'           txtIdou                :�ٓ�
'           txtTyuClub             :���w�Z�N���u
'           txtClub                :���݃N���u
'           txtRyoseiKbn           :��
'           CheckImage               :�摜�\���w��
'           txtMode                :���샂�[�h
'                               BLANK   :�����\��
'                               SEARCH  :���ʕ\��
' ��      ��:
'           �������\��
'               �^�C�g���̂ݕ\��
'           �����ʕ\��
'               ��y�[�W�Őݒ肳�ꂽ���������ɂ��Ȃ��w������\������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/02
'           : 2002/05/06 BLOB�^�Ή��̈� T09_IMAGE ���@T09_GAKUSEI_NO�ɕύX
' ��      �X: 2011/04/05 iwata �w���ʐ^�f�[�^���@Session����łȂ��A�f�[�^�x�[�X����擾����B
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_TxtMode      	       ':���샂�[�h
	Public  m_iSyoriNen      	   ':�����N�x
    Public  m_iHyoujiNendo         ':�\���N�x
    Public  m_sGakunen             ':�w�N
    Public  m_sGakkaCD             ':�w��
    Public  m_sClass               ':�N���X
    Public  m_sName                ':����
    Public  m_sGakusekiNo          ':�w�Дԍ�
    Public  m_sSeibetu             ':����
    Public  m_sGakuseiNo           ':�w���ԍ�
    Public  m_sIdou                ':�ٓ�
    Public  m_sTyuClub             ':���w�Z�N���u
    Public  m_sClub                ':���݃N���u
    Public  m_sRyoseiKbn           ':��
    Public  m_sCheckImage          ':�摜�\���w��
	Public  m_sTyugaku			   ':�o�g���w�Z

    Public	m_Rs					'recordset
    Public	m_iDsp					'�ꗗ�\���s��

    Public  m_iPageTyu      		':�\���ϕ\���Ő��i�������g����󂯎������j

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
    w_sMsgTitle="�w����񌟍�����"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear


    m_bErrFlg = False

	'//�Z�b�V�������E���샂�[�h�̎擾
	m_iSyoriNen = Session("NENDO")
    m_TxtMode=request("txtMode")

    Do
		if m_TxtMode = "" then
           	Call showPage()
			Exit Do
		End if

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

'2011.04.05 ins
    '// �摜�f�[�^�擾�p oo4o �Z�b�V�����쐬
    Set Session("OraDatabasePh") = OraSession.GetDatabaseFromPool(100)

        '// ���Ұ�SET
        Call s_SetParam()

        '�f�[�^���oSQL���쐬����
        Call s_MakeSQL(w_sSQL)

       '���R�[�h�Z�b�g�̎擾
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
		w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL,m_iDsp)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// �y�[�W��\��
        If m_Rs.EOF Then
            Call showPage_NoData()
        Else

	    '�w�����\��
            Call showPage()

        End If

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '// �I������
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

'2011.04.05 ins
		'** oo4o �ڑ��v�[���p��
	   Session("OraDatabasePh").DestroyDatabasePool

End Sub

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************

'    Session("HyoujiNendo") = request("txtHyoujiNendo")     	'�\���N�x
    Session("HyoujiNendo") = Session("NENDO")		'�\���N�x	'<-- 8/16�C��	��
    m_sGakunen=request("txtGakunen")            	'�w�N
	'�R���{���I����
	If m_sGakunen="@@@" Then
		m_sGakunen=""
	End If
    m_sGakkaCD=request("txtGakka")            	'�w��
	'�R���{���I����
	If m_sGakkaCD="@@@" Then
		m_sGakkaCD=""
	End If

	if m_sGakunen="" then	'�w�N���I������Ă��Ȃ��ꍇ�̓N���X�͑I���ł��܂���
		m_sClass=""
	else
    	m_sClass=request("txtClass")               	'�N���X
		'�R���{���I����
		If m_sClass="@@@" Then
			m_sClass=""
		End If
    end if

	m_sName = gf_Zen2Han(request("txtName"))                	'����(���p�ɕϊ�)

	m_sGakusekiNo=request("txtGakusekiNo")          '�w�Дԍ�
	m_sSeibetu=request("txtSeibetu")            	'����
	'�R���{���I����
	If m_sSeibetu="@@@" Then
		m_sSeibetu=""
	End If
	m_sGakuseiNo=request("txtGakuseiNo")           	'�w���ԍ�
	m_sIdou =request("TxtIdou")               	'�ٓ�
	'�R���{���I����
	If m_sIdou="@@@" Then
		m_sIdou=""
	End If
	m_sTyuClub =request("txtTyuClub")            	'���w�Z�N���u
	'�R���{���I����
	If m_sTyuClub="@@@" Then
		m_sTyuClub=""
	End If
	m_sClub=request("txtClub")                	'���݃N���u
	'�R���{���I����
	If m_sClub="@@@" Then
		m_sClub=""
	End If
	m_sRyoseiKbn=request("txtRyoseiKbn")           	'��
	'�R���{���I����
	If m_sRyoseiKbn="@@@" Then
		m_sRyoseiKbn=""
	End If

	m_iDsp = cint(request("txtDisp"))						':�������X�g�̕\������

    '// BLANK�̏ꍇ�͍s���ر
    If m_TxtMode = "Search" Then
        m_iPageTyu = 1
    Else
        m_iPageTyu = int(Request("txtPageTyu"))     ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If

	m_sCheckImage=request("CheckImage")           	'�摜�\���w��

	m_sTyugaku = request("txtTyugaku")

End Sub


'********************************************************************************
'*  [�@�\]  �w�Ѓf�[�^���oSQL������̍쐬
'*  [����]  p_sSql - SQL������
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub s_MakeSQL(p_sSql)

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " A.T13_GAKUSEKI_NO, "
    p_sSql = p_sSql & " A.T13_GAKUSEI_NO, "
    p_sSql = p_sSql & " A.T13_GAKUNEN, "
    p_sSql = p_sSql & " E.M05_CLASSMEI, "
    p_sSql = p_sSql & " B.T11_SIMEI, "
    p_sSql = p_sSql & " B.T11_SEIBETU, "
    p_sSql = p_sSql & " D.M01_SYOBUNRUIMEI, "
    p_sSql = p_sSql & " C.M02_GAKKARYAKSYO "
    p_sSql = p_sSql & " FROM T13_GAKU_NEN A, T11_GAKUSEKI B, M02_GAKKA C, M01_KUBUN D, M05_CLASS E "
    p_sSql = p_sSql & " WHERE A.T13_NENDO = " & cint(Session("HyoujiNendo")) & ""

    '���������̃Z�b�g
    if m_sGakunen <> "" then        '�w�N
        p_sSql = p_sSql & " AND A.T13_GAKUNEN = " & cint(m_sGakunen)
    end if
    if m_sGakkaCD <> "" then         '�w��
        p_sSql = p_sSql & " AND A.T13_GAKKA_CD = '" & m_sGakkaCD & "'"
    end if
    if m_sClass <> "" then           '�N���X
        p_sSql = p_sSql & " AND A.T13_CLASS = '" & m_sClass & "'"
    end if
    if m_sName <> "" then            '����
        p_sSql = p_sSql & " AND B.T11_SIMEI_KD LIKE '%" & m_sName & "%'"
    end if
    if m_sGakusekiNo <> "" then 	'�w�Дԍ�
    	p_sSql = p_sSql & " AND A.T13_GAKUSEKI_NO LIKE '%" & m_sGakusekiNo & "%'"
    end if
    if m_sSeibetu <> "" then 		'����
    	p_sSql = p_sSql & " AND B.T11_SEIBETU = " & m_sSeibetu
    end if
    if m_sGakuseiNo <> "" then 		'�w���ԍ�
        p_sSql = p_sSql & " AND A.T13_GAKUSEI_NO LIKE '%" & m_sGakuseiNo & "%'"
    end if
    if m_sIdou <> "" then 		'�ٓ�
        p_sSql = p_sSql & " AND ( A.T13_IDOU_KBN_1 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_2 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_3 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_4 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_5 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_6 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_7 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    A.T13_IDOU_KBN_8 = '" & m_sIdou & "' )"
    end if

    if m_sTyuClub <> "" then 		'���w�Z�N���u
	    p_sSql = p_sSql & " AND B.T11_TYU_CLUB = '" & m_sTyuClub & "'"
    end if
    if m_sClub <> "" then 		'���݃N���u
	    p_sSql = p_sSql & " AND ( A.T13_CLUB_1 = '" & m_sClub & "'"
	    p_sSql = p_sSql & " OR A.T13_CLUB_2 = '" & m_sClub & "' ) "
    end if
    if m_sRyoseiKbn <> "" then 		'��
        p_sSql = p_sSql & " AND A.T13_RYOSEI_KBN = '" & m_sRyoseiKbn & "'"
    end if

    if m_sTyugaku <> "" then 		'�o�g���w�Z
        p_sSql = p_sSql & " AND B.T11_TYUGAKKO_CD IN ("
        p_sSql = p_sSql & " SELECT M13_TYUGAKKO_CD "
        p_sSql = p_sSql & " FROM M13_TYUGAKKO "
        p_sSql = p_sSql & " WHERE M13_TYUGAKKOMEI like '%" & m_sTyugaku & "%') "
    end if

    '��������
'    p_sSql = p_sSql & " AND A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO "
'    p_sSql = p_sSql & " AND M02_NENDO(+) = '" & cstr(Session("HyoujiNendo")) & "'"
'    p_sSql = p_sSql & " AND M02_GAKKA_CD(+) = T13_GAKKA_CD "
'    p_sSql = p_sSql & " AND M01_NENDO(+) = '" & cstr(Session("HyoujiNendo")) & "'"
'    p_sSql = p_sSql & " AND M01_DAIBUNRUI_CD(+) = 1 "
'    p_sSql = p_sSql & " AND M01_SYOBUNRUI_CD(+) = T11_SEIBETU "
'    p_sSql = p_sSql & " AND M05_CLASSNO(+) = T13_CLASS "


    p_sSql = p_sSql & " AND A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO(+) "
    p_sSql = p_sSql & " AND A.T13_NENDO = C.M02_NENDO"
    p_sSql = p_sSql & " AND A.T13_NENDO = D.M01_NENDO"
    p_sSql = p_sSql & " AND A.T13_NENDO = E.M05_NENDO"
    p_sSql = p_sSql & " AND A.T13_GAKUNEN = E.M05_GAKUNEN "
    p_sSql = p_sSql & " AND A.T13_CLASS = E.M05_CLASSNO "
    p_sSql = p_sSql & " AND A.T13_GAKKA_CD = C.M02_GAKKA_CD "
    p_sSql = p_sSql & " AND D.M01_DAIBUNRUI_CD = 1 "
    p_sSql = p_sSql & " AND B.T11_SEIBETU = D.M01_SYOBUNRUI_CD "

    p_sSql = p_sSql & " ORDER BY A.T13_GAKUNEN,A.T13_GAKUSEKI_NO "
'    p_sSql = p_sSql & " ORDER BY A.T13_GAKUNEN, D.M01_SYOBUNRUIMEI,"
'    p_sSql = p_sSql & " A.T13_CLASS, A.T13_GAKUSEKI_NO, A.T13_GAKUSEI_NO "
'response.write " p_sSql=" & p_sSql & "<BR>"

End Sub

'********************************************************************************
'*  [�@�\]  �ʐ^�����邩���� (BLOB�^�Ή��̈� T09_IMAGE ���@T09_GAKUSEI_NO�ɕύX�j
'*  [����]  �Ȃ�
'*  [�ߒl]  True: False
'*  [����]
'********************************************************************************
Function f_Photoimg(pGAKUSEI_NO)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_Photoimg = False

	'// NULL�Ȃ甲����(False)
	if trim(pGAKUSEI_NO) = "" then Exit Function

	Do
	    w_sSQL = ""
	    w_sSQL = w_sSQL & " SELECT "
	    w_sSQL = w_sSQL & " T09_GAKUSEI_NO "
	    w_sSQL = w_sSQL & " FROM T09_GAKU_IMG "
	    w_sSQL = w_sSQL & " WHERE T09_GAKUSEI_NO = '" & cstr(pGAKUSEI_NO) & "'"

		iRet = gf_GetRecordset(w_ImgRs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			Exit Do
		End If

		'// EOF�Ȃ甲����(False)
		if w_ImgRs.Eof then	Exit Do

		'//����I��
		f_Photoimg = True
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub showPage_NoData()

%>
	<html>
	<head>
	<title>�w����񌟍�</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
	<link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>

	<body>

	<center>
		<br><br><br>
		<span class="msg">�Ώۃf�[�^�͑��݂��܂���B��������͂��Ȃ����Č������Ă��������B</span>
	</center>

	</body>
    </html>

<%
    '---------- HTML END   ----------
End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub showPage()
	Dim w_pageBar			'�y�[�WBAR�\���p
%>

<html>

<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
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
    //  [�@�\]  �ڍ׃{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_detail(pGAKUSEI_NO){

			url = "kojin.asp?hidGAKUSEI_NO=" + pGAKUSEI_NO;
			w   = 800;
			h   = 600;

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

    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="main.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageTyu.value = p_iPage;
        document.frm.submit();

    }

    //-->
    </SCRIPT>
    </head>

    <body>
	<% if m_TxtMode = "" then %>
		<center>
		<br><br><br>
		<span class="msg">���ڂ�I��ŕ\���{�^���������Ă�������</span>
		</center>
	<% Else %>
	    <div align="center">
	    <form action="kojin.asp" method="post" name="frm" target="_detail">

		<BR>
		<table><tr><td align="center">
		<%
			'�y�[�WBAR�\��
			Call gs_pageBar(m_Rs,m_iPageTyu,m_iDsp,w_pageBar)
		%>
		<%=w_pageBar %>

			<table border="0" width="100%">
				<tr>
					<td align="center">
					<% if m_TxtMode = "" then %>
						<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
							<tr>
								<td width="60">&nbsp</td>
								<td valign="top"></td>
							</tr>
						</table>
					<% else %>
						<% dim w_cell %>

					    <!--  �w�����\���@-->
						<% if m_sCheckImage = "" then %>
								<table border="1" width="600" class=hyo>
									<tr>
										<th height=16 class=header><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
										<th height=16 class=header>�w�N</th>
										<th height=16 class=header>�w��</th>
										<th height=16 class=header>�N���X</th>
										<th height=16 class=header>���@�@��</th>
										<th height=16 class=header>����</th>
										<th height=16 class=header><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_5NEN)%></th>
										<th height=16 class=header>�ڍ�</th>
									</tr>

						        	<%
	'									m_Rs.Movefirst
										w_iCnt = 0
										Do Until m_Rs.EOF or w_iCnt >= m_iDsp
											call gs_cellPtn(w_cell)
											%>
											<tr>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEKI_NO")) %>&nbsp</td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T13_GAKUNEN")) %>&nbsp</td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M02_GAKKARYAKSYO")) %>&nbsp</td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M05_CLASSMEI")) %>&nbsp</td>
												<td align="left"   height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T11_SIMEI")) %></td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M01_SYOBUNRUIMEI")) %></td>
												<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEI_NO")) %>&nbsp</td>
												<td align="center" height="16" class=<%=w_cell%>><input type=button class=button value="�ڍ�" onclick="f_detail('<%= gf_HTMLTableSTR(m_Rs("T13_GAKUSEI_NO")) %>');"></td>
											</tr>
											<%
											w_iCnt = w_iCnt + 1
											m_Rs.MoveNext
										Loop
									%>

								</table>
						<% else %>
						<!--  �w���ʐ^�\���@-->

							<table border="0" cellpadding="0" cellspacing="2">
								<%
									w_iCnt = 1
									Do Until m_Rs.Eof or w_iCnt > m_iDsp
										response.write 	"<tr>"
										i_TdLine = 1					'// ���ɂS���\�����C��
										Do Until m_Rs.Eof or i_TdLine > 4 or w_iCnt > m_iDsp
										%>
											<td align="center" class=search width="150" valign="top">
												<a href="javascript:f_detail('<%= gf_HTMLTableSTR(m_Rs("T13_GAKUSEI_NO")) %>');">
												<%
												'// ��ʐ^�����邩��Ɍ�������
												w_bRet = ""
												w_bRet = f_Photoimg(m_Rs("T13_GAKUSEI_NO"))

												if w_bRet = True then
													' 2011.04.05 upd DispBinary => DispBinaryRec �ɕύX
													%><IMG SRC="DispBinaryRec.asp?gakuNo=<%= m_Rs("T13_GAKUSEI_NO") %>" width="88" height="136" border="0"><%
												Else
													%><IMG SRC="images/Img0000000000.gif" width="100" height="120" border="0"><%
												End if
												%></a><br>
												<table border="0" cellpadding="0" cellspacing="2" width="100%">
													<tr><td bgcolor="#666699" nowrap><font color="#FFFFFF"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></font></td><td><%= m_Rs("T13_GAKUSEKI_NO") %></td></tr>
													<tr><td bgcolor="#666699" nowrap><font color="#FFFFFF">����    </font></td><td><%= trim(m_Rs("T11_SIMEI")) %></td></tr>
												</table>
											</td>
											<%
											i_TdLine = i_TdLine + 1
											w_iCnt = w_iCnt + 1
											m_Rs.MoveNext
										Loop %>
										</tr>
									<% Loop	%>
								</tr>
							</table>

						<% end if %>

					<% end if %>
				</td>
			</tr>
		</table>

		<%=w_pageBar %>
		</td></tr></table>

		</div>
	    <input type="hidden" name="txtMode">
	    <input type="hidden" name="txtPageTyu" value="<%=m_iPageTyu%>">
	    <input type="hidden" name="hidGAKUSEI_NO">

		<%' �������� %>
		<input type="hidden" name="txtHyoujiNendo" value="<%=request("txtHyoujiNendo")%>">
		<input type="hidden" name="txtGakunen"     value="<%=request("txtGakunen")%>">
		<input type="hidden" name="txtGakka"       value="<%=request("txtGakka")%>">
		<input type="hidden" name="txtClass"       value="<%=request("txtClass")%>">
		<input type="hidden" name="txtName"        value="<%=request("txtName")%>">
		<input type="hidden" name="txtGakusekiNo"  value="<%=request("txtGakusekiNo")%>">
		<input type="hidden" name="txtSeibetu"     value="<%=request("txtSeibetu")%>">
		<input type="hidden" name="txtGakuseiNo"   value="<%=request("txtGakuseiNo")%>">
		<input type="hidden" name="TxtIdou"        value="<%=request("TxtIdou")%>">
		<input type="hidden" name="txtTyuClub"     value="<%=request("txtTyuClub")%>">
		<input type="hidden" name="txtClub"        value="<%=request("txtClub")%>">
		<input type="hidden" name="txtRyoseiKbn"   value="<%=request("txtRyoseiKbn")%>">
		<input type="hidden" name="CheckImage"     value="<%=request("CheckImage")%>">
		<input type="hidden" name="txtTyugaku"     value="<%=request("txtTyugaku")%>">
		<input type="hidden" name="txtDisp"        value="<%=request("txtDisp")%>">
		</form>
	<% End if %>
	</body>

    </html>

<%
    '---------- HTML END   ----------
End Sub

%>

