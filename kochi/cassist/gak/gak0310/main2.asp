<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w�Ѓf�[�^��������(�摜�\��
' ��۸���ID : gak/gak0300/main2.asp
' �@      �\: ���y�[�W �w�Ѓf�[�^�̌������ʂ��摜�\������
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
'           txtMode                :���샂�[�h
'                               BLANK   :�����\��
'                               SEARCH  :���ʕ\��
' ��      ��:
'           �������\��
'               �^�C�g���̂ݕ\��
'           �����ʕ\��
'               ��y�[�W�Őݒ肳�ꂽ���������ɂ��Ȃ��w�������摜�\������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�擾�����f�[�^�����ϐ�
    Public  m_TxtMode      	       ':���샂�[�h
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
    
    Public	m_Rs					'recordset
    Public	m_iDsp					'// �ꗗ�\���s��

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
    w_sMsgTitle="�w�Ѓf�[�^��������"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	'//�Z�b�V�������E���샂�[�h�̎擾
    m_TxtMode=request("txtMode")
    
    Do
        '// �ް��ް��ڑ�
		w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

        '//�����\��
        if m_TxtMode = "" then
            Call showPage()
            Exit Do
        end if

        '// ���Ұ�SET
        Call s_SetParam()

        '�f�[�^���oSQL���쐬����
        Call s_MakeSQL(w_sSQL)

        '���R�[�h�Z�b�g�̎擾
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// �y�[�W��\��
        If m_Rs.EOF Then
            Call showPage_NoData()
        Else
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

End Sub

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iHyoujiNendo =request("txtHyoujiNendo")     	'�\���N�x
    m_sGakunen=request("txtGakunen")            	'�w�N
	'�R���{���I����
	If m_sGakunen="@@@" Then
		m_sGakunen=""
	End If
    m_sGakkaCD=request("txtGakka")             		'�w��
	'�R���{���I����
	If m_sGakkaCD="@@@" Then
		m_sGakkaCD=""
	End If

	if m_sGakunen="" then		'�w�N���I������Ă��Ȃ��ꍇ�̓N���X�͑I���ł��܂���
		m_sClass=""
	else 
    	m_sClass=request("txtClass")               		'�N���X
		'�R���{���I����
		If m_sClass="@@@" Then
			m_sClass=""
		End If
    end if 
	m_sName=request("txtName")                		'����
	m_sGakusekiNo=request("txtGakusekiNo")          '�w�Дԍ�
	m_sSeibetu=request("txtSeibetu")            	'����
	'�R���{���I����
	If m_sSeibetu="@@@" Then
		m_sSeibetu=""
	End If
	m_sGakuseiNo=request("txtGakuseiNo")           	'�w���ԍ�
	m_sIdou =request("TxtIdou")               		'�ٓ�
	'�R���{���I����
	If m_sIdou="@@@" Then
		m_sIdou=""
	End If
	m_sTyuClub =request("txtTyuClub")            	'���w�Z�N���u
	'�R���{���I����
	If m_sTyuClub="@@@" Then
		m_sTyuClub=""
	End If
	m_sClub=request("txtClub")                		'���݃N���u
	'�R���{���I����
	If m_sClub="@@@" Then
		m_sClub=""
	End If
	m_sRyoseiKbn=request("txtRyoseiKbn")           	'��
	'�R���{���I����
	If m_sRyoseiKbn="@@@" Then
		m_sRyoseiKbn=""
	End If

End Sub

Sub s_MakeSQL(p_sSql)
'********************************************************************************
'*  [�@�\]  �w�Ѓf�[�^���oSQL������̍쐬
'*  [����]  p_sSql - SQL������
'*  [�ߒl]  �Ȃ� 
'*  [����]  
'********************************************************************************

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " T13_GAKUSEKI_NO, "
    p_sSql = p_sSql & " T13_GAKUSEI_NO, "
    p_sSql = p_sSql & " T13_GAKUNEN, "
    p_sSql = p_sSql & " T13_CLASS, "
    p_sSql = p_sSql & " T11_SIMEI, "
    p_sSql = p_sSql & " T11_SEIBETU, "
    p_sSql = p_sSql & " M01_SYOBUNRUIMEI, "
    p_sSql = p_sSql & " M02_GAKKARYAKSYO "
    p_sSql = p_sSql & " FROM T13_GAKU_NEN, T11_GAKUSEKI, M02_GAKKA, M01_KUBUN "

    p_sSql = p_sSql & " WHERE T13_NENDO = '" & cstr(m_iHyoujiNendo) & "'"

    '���������̃Z�b�g
    if m_sGakunen <> "" then        '�w�N
        p_sSql = p_sSql & " AND T13_GAKUNEN = " & cint(m_sGakunen)
    end if
    if m_sGakkaCD <> "" then         '�w��
        p_sSql = p_sSql & " AND T13_GAKKA_CD = '" & m_sGakkaCD & "'"
    end if
    if m_sClass <> "" then           '�N���X
        p_sSql = p_sSql & " AND T13_CLASS = '" & m_sClass & "'"
    end if
    if m_sName <> "" then            '����
        p_sSql = p_sSql & " AND T11_SIMEI_KD LIKE '%" & m_sName & "%'"
    end if
    if m_sGakusekiNo <> "" then 	'�w�Дԍ�
    	p_sSql = p_sSql & " AND T13_GAKUSEKI_NO LIKE '%" & m_sGakusekiNo & "%'"
    end if
    if m_sSeibetu <> "" then 		'����
    	p_sSql = p_sSql & " AND T11_SEIBETU = " & m_sSeibetu
    end if
    if m_sGakuseiNo <> "" then 		'�w���ԍ�
        p_sSql = p_sSql & " AND T13_GAKUSEI_NO LIKE '%" & m_sGakuseiNo & "%'"
    end if
    if m_sIdou <> "" then 		'�ٓ�
        p_sSql = p_sSql & " AND ( T13_IDOU_KBN_1 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_2 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_3 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_4 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_5 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_6 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_7 = '" & m_sIdou & "'"
        p_sSql = p_sSql & " OR    T13_IDOU_KBN_8 = '" & m_sIdou & "' )"
    end if

    if m_sTyuClub <> "" then 		'���w�Z�N���u
	    p_sSql = p_sSql & " AND T11_TYU_CLUB = '" & m_sTyuClub & "'"
    end if
    if m_sClub <> "" then 		'���݃N���u
	    p_sSql = p_sSql & " AND ( T13_CLUB_1 = '" & m_sClub & "'"
	    p_sSql = p_sSql & " OR T13_CLUB_2 = '" & m_sClub & "' ) "
    end if
    if m_sRyoseiKbn <> "" then 		'��
        p_sSql = p_sSql & " AND T13_RYOSEI_KBN = '" & m_sRyoseiKbn & "'"
    end if
    
    '��������
    p_sSql = p_sSql & " AND T13_GAKUSEI_NO = T11_GAKUSEI_NO "
    p_sSql = p_sSql & " AND M02_NENDO(+) = '" & cstr(m_iHyoujiNendo) & "'"
    p_sSql = p_sSql & " AND M02_GAKKA_CD(+) = T13_GAKKA_CD "
    p_sSql = p_sSql & " AND M01_NENDO(+) = '" & cstr(m_iHyoujiNendo) & "'"
    p_sSql = p_sSql & " AND M01_DAIBUNRUI_CD(+) = 1 "
    p_sSql = p_sSql & " AND M01_SYOBUNRUI_CD(+) = T11_SEIBETU "

    p_sSql = p_sSql & " ORDER BY T13_GAKUNEN, M01_SYOBUNRUIMEI,"
    p_sSql = p_sSql & " T13_CLASS, T13_GAKUSEKI_NO, T13_GAKUSEI_NO "

'response.write " p_sSql=" & p_sSql & "<BR>"

End Sub

Sub showPage_NoData()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
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

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
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
    function f_detail(){

        window.alert("�ڍ�");
		document.forms[0].submit();
    }

    //-->
    </SCRIPT>

    </head>

    <body>

    <center>
    <form action="kojin.asp" method="post" name="frm" target="_detail">

	<% if m_TxtMode = "" then %>
	<table border="0" width="100%">
	
		<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
	<tr>
		<td width="60">&nbsp</td>
		<td valign="top"></td>
	</tr>
	</table>
	
	<% else %>
	
	<tr>
		<td align="center">

		<table border="1" width="600" class=hyo>
		<tr><!-- �P�N�Ԕԍ�  -->
		<th height=16 class=header><%=gf_GetGakuNomei(m_iHyoujiNendo,C_K_KOJIN_1NEN)%></th>
		<th height=16 class=header>�w���ʐ^</th>
		</tr>
		
        <% Do While not m_Rs.EOF %>
		<% dim w_cell
		   call gs_cellPtn(w_cell)%>
		<tr>
		<td align="center" height="16" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEKI_NO")) %>&nbsp</td>
		<td >
<!--
                   <IMG SRC="./DispBinary.asp?gakuNo=0000000001"> -->
                   <!-- <IMG SRC="./DispBinary.asp?gakuNo=<%=gf_HTMLTableSTR(m_Rs("T13_GAKUSEKI_NO")) %>"> -->
                </td>
		</tr>
		
		<%   m_Rs.MoveNext
		Loop %>
		
	    </table>
		</td>
		</tr>
		
		<% end if %>

	</center>

	</body>
    </html>


<%
    '---------- HTML END   ----------
End Sub

%>



















































