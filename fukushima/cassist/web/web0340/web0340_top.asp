<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �l���C�I���Ȗڌ���
' ��۸���ID : web/web0340/web0340_top.asp
' �@      �\: ��y�[�W �l���C�I���Ȗڌ���̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/23 �O�c �q�j
' ��      �X: 2001/08/07 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
' ��      �X: 2015/03/20 ���{ ��H     Win7�Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�s�����I��p��Where����
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '�����R�[�h
    Public m_sKBN           '�敪�R���{�{�b�N�X�ɓ���l
    Public m_sGRP           '�����R���{�{�b�N�X�ɓ���l
    Public m_sKBNWhere      '�N�x�R���{�{�b�N�X�̏���
    Public m_sGRPWhere      '�����R���{�{�b�N�X�̏���
    Public m_sOption        '�����R���{�{�b�N�X�̎g�p�A�s�̔���
    Public m_sGakunen       '�w�N
    Public m_sClass         '�N���X
    Public m_sGakka         '�w�ȃR�[�h
    Public m_rs             '
    Public m_sGakunenWhere      '//�w�N�̏���
    Public m_sGakunenOption     '//�w�N�R���{�̃I�v�V����
    Public m_sClassWhere        '//�N���X�̏���
    Public m_sClassOption       '//�N���X�R���{�̃I�v�V����
    Public m_sKengen

    Public  m_iMax          '�ő�y�[�W
    Public  m_iDsp          '�ꗗ�\���s��

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
    w_sMsgTitle="�l���C�I���Ȗڌ���"
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
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

		'//�������擾
		w_iRet = gf_GetKengen_web0340(m_sKengen)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'//�f�[�^��ϐ��ɃZ�b�g
		Call s_SetParam()

		'//�敪�R���{�Ɋւ���WHERE���쐬����
        w_iRet = f_KBNWhere()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//�O���[�v�R���{�Ɋւ���WHERE���쐬����
        Call f_GRPWhere()

'//�f�o�b�O
'call s_DebugPrint


        '//�w�N�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeGakunenWhere() 

        '//�N���X�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeClassWhere() 

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
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()


    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp = C_PAGE_LINE

	'//�������S�C�̏ꍇ�́A�S�C�N���X�̂ݓo�^���\�Ƃ���
	If m_sKengen = C_WEB0340_ACCESS_TANNIN Then

		'//�S�C�����̏ꍇ�́A�S�C�N���X�̔N�g���擾����
		Call f_Gakunen()
	Else
		'//�S�C�ȊO�̏ꍇ
	    m_sGakunen  = Request("cboGakunenCd")   '//�w�N
	    m_sClass    = Request("cboClassCd")     '//�N���X

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

    response.write "m_iNendo     = " & m_iNendo     & "<br>"
    response.write "m_sKyokanCd  = " & m_sKyokanCd  & "<br>"
    response.write "m_sGakunen   = " & m_sGakunen   & "<br>"
    response.write "m_sClass     = " & m_sClass     & "<br>"
    response.write "m_sGakka     = " & m_sGakka     & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �w�N�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeGakunenWhere()

    m_sGakunenWhere = ""
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iNendo
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

	If m_sKengen = C_WEB0340_ACCESS_TANNIN Then
		m_sGakunenOption = "DISABLED"
	End If

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeClassWhere()

    m_sClassWhere = ""
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iNendo

    If m_sGakunen = "" Then
        '//�����\������1�N1�g��\������
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = 1"
    Else
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_sGakunen)
    End If

	'//�������S�C�̏ꍇ�́A�S�C�N���X�ȊO�̓o�^�͏o���Ȃ�
	If m_sKengen = C_WEB0340_ACCESS_TANNIN Then
		m_sClassOption = "DISABLED"
	End If

End Sub

Function f_KBNWhere()
'********************************************************************************
'*  [�@�\]  �敪�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_KBNWhere = 1

    Do

        m_sKBNWhere = ""
        m_sKBNWhere = m_sKBNWhere & " M01_DAIBUNRUI_CD = " & C_KAMOKU & " AND "
        m_sKBNWhere = m_sKBNWhere & " M01_NENDO        = " & m_iNendo & " AND "
        m_sKBNWhere = m_sKBNWhere & " M01_SYOBUNRUI_CD <> 2 "

        m_sKBN = request("txtKBN")

        If request("txtKBN") = C_CBO_NULL Then m_sKBN = ""

        f_KBNWhere = 0
        Exit Do
    Loop


End Function

Sub f_GRPWhere()
'********************************************************************************
'*  [�@�\]  �O���[�v�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_iNyuNendo

	Dim w_sGroup,w_bFlg

    m_sGRPWhere=""
    m_sOption=""
	w_sGroup = ""
	w_bDispFlg = True

    If m_sKBN <> "" Then

		'=============
		'//�w�Ȃ̎擾
		'=============
		Call f_GetGakka(m_sGakunen,m_sClass)

		'==============================================================
		'//��������勳���̏ꍇ�́A�S���Ȗڊ֘A�̂ݓo�^���\�Ƃ���
		'==============================================================
		If m_sKengen = C_WEB0340_ACCESS_SENMON Then

			'//��勳���̏ꍇ�́A�֘A���Ȃ̉Ȗڂ̂ݓ��͉Ƃ���
			'//�Ȗڂ̃O���[�v���擾(����)
			Call f_GetGroup(m_sGakunen,m_sClass,w_sGroup)

			If Trim(w_sGroup) = "" Then
				w_bDispFlg = False
			End If

		End If

		'==============================================================
		'//�O���[�v�R���{���擾
		'==============================================================
		If w_bDispFlg = True Then

	        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

	        m_sGRPWhere = " T15_HISSEN_KBN = " & C_HISSEN_SEN & " AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_GAKKA_CD = " & m_sGakka & " AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_NYUNENDO = " & w_iNyuNendo & " AND "
	        m_sGRPWhere = m_sGRPWhere & " T15_KAMOKU_KBN = " & cInt(m_sKBN) & " AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_NYUNENDO = T15_NYUNENDO(+) AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_GRP = T15_GRP(+) AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_GAKKA_CD = T15_GAKKA_CD(+) AND "
	        m_sGRPWhere = m_sGRPWhere & " T18_GRP <> " & C_T18_GRP & " "

			If w_sGroup <> "" Then
		        m_sGRPWhere = m_sGRPWhere & " AND T18_GRP IN (" & w_sGroup & ")"
			End If

	        m_sGRPWhere = m_sGRPWhere & " GROUP BY T18_GRP,T18_SYUBETU_MEI "

		Else
	        m_sOption = " DISABLED "
	        m_sGRPWhere  = " T18_GAKKA_CD = 00 "
		End If

    Else
        m_sOption = " DISABLED "
        m_sGRPWhere  = " T18_GAKKA_CD = 00 "
    End IF

End Sub

Sub f_Gakunen()
'********************************************************************************
'*  [�@�\]  �w�N�̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    '//�w�N��N���X�̃f�[�^
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT"
    w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_GAKKA_CD "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     M05_CLASS "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & "     M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "

    Set m_rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(m_rs, w_sSQL,m_iDsp)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Sub
    End If

	If m_rs.EOF = false Then
	    m_sGakka   = m_rs("M05_GAKKA_CD")
	    m_sGakunen = cInt(m_rs("M05_GAKUNEN"))
	    m_sClass   = cInt(m_rs("M05_CLASSNO"))
	End If

   Call gf_closeObject(m_rs)

End Sub

Sub f_GetGakka(p_sGakuNen,p_sClass)
'********************************************************************************
'*  [�@�\]  �w�Ȃ̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	Dim rs
	Dim w_sSQL

    '//�w�N��N���X���w�Ȃ��擾
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT"
    w_sSQL = w_sSQL & "     M05_GAKKA_CD "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     M05_CLASS "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & "     M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "     AND M05_GAKUNEN = " & p_sGakuNen
    w_sSQL = w_sSQL & "     AND M05_CLASSNO = " & p_sClass

    w_iRet = gf_GetRecordset(rs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Sub
    End If

	If rs.EOF = false Then
	    m_sGakka   = rs("M05_GAKKA_CD")
	End If

   Call gf_closeObject(rs)

End Sub

Function f_GetGroup(p_sGakuNen,p_sClass,p_sGroup)
'********************************************************************************
'*  [�@�\]  �w�Ȃ̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	Dim rs
	Dim w_sSQL
	Dim w_sKamoku

	w_sKamoku = ""
	p_sGroup = ""

	Do 

	    '//�����̊֘A�Ȗڂ��擾����
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN.T27_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T27_TANTO_KYOKAN.T27_NENDO= " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_GAKUNEN=" & p_sGakuNen
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_CLASS=" & p_sClass
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_KYOKAN_CD=" & m_sKyokanCd
'response.write w_sSQL
'���C�������t���O�������Ă��鋳���̂ݓ��͉\�@add 2001/10/29 tani
		w_sSQL = w_sSQL & vbCrLf & "  AND T27_TANTO_KYOKAN.T27_MAIN_FLG=" & C_MAIN_KYOKAN_YES

'response.write w_sSQL

	    w_iRet = gf_GetRecordset(rs, w_sSQL)
	    If w_iRet <> 0 Then
	        'ں��޾�Ă̎擾���s
	        m_bErrFlg = True
	        Exit Do
	    End If

		If rs.EOF = True Then
			Exit Do
		Else

			'//�Ȗ�CD���擾
			Do Until rs.EOF
				If w_sKamoku = "" Then
				    w_sKamoku = rs("T27_KAMOKU_CD")
				Else
					w_sKamoku = w_sKamoku  & "," & rs("T27_KAMOKU_CD")
				End If
				rs.MoveNext
			Loop

		End If

		'//�Ȗ�CD���擾�ł��Ȃ��Ƃ�
		If Trim(w_sKamoku) = "" Then
			Exit Do
		End If

		'//�Ȗڂ̃O���[�v�̎�ނ��擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_GRP"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(p_sGakuNen) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD= '" & m_sGakka & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD IN (" & w_sKamoku & ")"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_KBN=" & m_sKBN 
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_HISSEN_KBN=" & C_HISSEN_SEN '//�I���Ȗڂ̂�

	    w_iRet = gf_GetRecordset(rs_K, w_sSQL)
	    If w_iRet <> 0 Then
	        'ں��޾�Ă̎擾���s
	        m_bErrFlg = True
	        Exit Do
	    End If

		If rs_K.EOF Then
			Exit Do
		Else

			Do Until rs_K.EOF

				If p_sGroup = "" Then
				    p_sGroup = rs_K("T15_GRP")
				Else
				    p_sGroup = p_sGroup & "," & rs_K("T15_GRP")
				End If

				rs_K.MoveNext
			Loop

		End If

		Exit Do
	Loop

    '//ں��޾��CLOSE
   Call gf_closeObject(rs)
   Call gf_closeObject(rs_K)

End Function

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

<title>�l���C�I���Ȗڌ���</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [�@�\]  �N�x���C�����ꂽ�Ƃ��A�ĕ\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="web0340_top.asp";
        document.frm.target="top";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

        // ������NULL����������
        // ���N�x
        if( f_Trim(document.frm.txtKBN.value) == "" ){
            window.alert("�敪�̑I�����s���Ă�������");
            document.frm.txtKBN.focus();
            return ;
        }
        // ���N�x
        if( f_Trim(document.frm.txtKBN.value) == "<%=C_CBO_NULL%>" ){
            window.alert("�敪�̑I�����s���Ă�������");
            document.frm.txtKBN.focus();
            return ;
        }
        // ���I�����
<% if m_sOption <> "" then '�I���ł���I����ʂ��Ȃ��Ƃ��́A�t�H�[�J�X���Ȃ� %>
        if( f_Trim(document.frm.txtGRP.value) == "" ){
            window.alert("�I���ł���I����ʂ�����܂���B");
            return ;
        }
        // 
        if( f_Trim(document.frm.txtGRP.value) == "<%=C_CBO_NULL%>" ){
            window.alert("�I���ł���I����ʂ�����܂���B");
            return ;
        }
<% Else 	'�I�����ĂȂ��Ƃ�%>
        // 
        if( f_Trim(document.frm.txtGRP.value) == "" ){
            window.alert("�I����ʂ̑I�����s���Ă�������");
            document.frm.txtGRP.focus();
            return ;
        }
        // 
        if( f_Trim(document.frm.txtGRP.value) == "<%=C_CBO_NULL%>" ){
            window.alert("�I����ʂ̑I�����s���Ă�������");
            document.frm.txtGRP.focus();
            return ;
        }
<% End If %>
		//�w�N�A�N���X���Z�b�g
		document.frm.txtGakunen.value = document.frm.cboGakunenCd.value
		document.frm.txtClass.value =document.frm.cboClassCd.value

        document.frm.action="web0340_main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }
    //************************************************************
    //  [�@�\]  �N���A�{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Clear(){

        document.frm.txtKBN.value = "@@@";
        document.frm.txtGRP.value = "@@@";
    
    }

    //-->
    </SCRIPT>

    <link rel="stylesheet" href="../../common/style.css" type="text/css">

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<center>

<form name="frm" METHOD="post">

<table cellspacing="0" cellpadding="0" border="0" width="100%">
    <tr>
        <td valign="top" align="center">
<%call gs_title("�l���C�I���Ȗڌ���","��@��")%>
<br>
            <table border="0">
                <tr>
                    <td class="search">
                        <table border="0" cellpadding="1" cellspacing="1">
                            <tr>
                                <td align="left">
                                    <table border="0" cellpadding="1" cellspacing="1">

                                        <tr>
                                            <td Nowrap align="left">�w�@�N
											<% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' " &  m_sGakunenOption ,False,m_sGakunen) %>�@�N���X
											<!-- 2015.03.20 Upd width:80->180 -->
											<% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:180px;' " & m_sClassOption,False,m_sClass) %>
                                            </td>
                                        </tr>

                                        <tr>
                                            <td Nowrap align="left">��@��
											<%call gf_ComboSet("txtKBN",C_CBO_M01_KUBUN,m_sKBNWhere,"style='width:120px;' onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sKBN)%>
                                            </td>
                                            <td Nowrap align="left">�@�I�����
											<%call gf_PluComboSet("txtGRP",C_CBO_T18_SEL_SYUBETU,m_sGRPWhere, "style='width:160px;' "& m_sOption,True,m_sGRP)%>
                                            </td>
                                        </tr>
										<tr>
											<td colspan="2" align="right">
									        <input type="button" class="button" value=" �N�@���@�A " onclick="javasript:f_Clear();">
											<input class="button" type="button" value="�@�\�@���@" onClick = "javascript:f_Search()">
											</td>
										</tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>

<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
<input type="hidden" name="txtClass"    value="<%=m_sClass%>">
<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">

</form>

</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
