<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���x���ʉȖڌ���
' ��۸���ID : web/web0390/web0390_top.asp
' �@      �\: ��y�[�W ���x���ʉȖڌ���̌������s��
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
' ��      ��: 2001/10/26 �J�e�@�ǖ�
' ��      �X: 
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
    w_sMsgTitle="���x���ʉȖڌ���"
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
		w_iRet = gf_GetKengen_web0390(m_sKengen)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'�����̒萔
		'C_WEB0390_ACCESS_FULL  
		'C_WEB0390_ACCESS_SENMON
		'C_WEB0390_ACCESS_TANNIN

		'//�f�[�^��ϐ��ɃZ�b�g
		Call s_SetParam()

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
	If m_sKengen = C_WEB0390_ACCESS_TANNIN Then

		'//�S�C�����̏ꍇ�́A�S�C�N���X�̔N�g���擾����
		Call f_Gakunen()
	Else
		'//�S�C�ȊO�̏ꍇ
	    m_sGakunen  = Request("cboGakunenCd")   '//�w�N
		if m_sGakunen = "@@@" OR m_sGakunen = "" then m_sGakunen = "1"
	    m_sClass    = Request("cboClassCd")     '//�N���X
		if m_sClass = "@@@" OR m_sClass = "" then m_sClass = "1"

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

	If m_sKengen = C_WEB0390_ACCESS_TANNIN Then
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
	If m_sKengen = C_WEB0390_ACCESS_TANNIN Then
		m_sClassOption = "DISABLED"
	End If

End Sub

'********************************************************************************
'*  [�@�\]  ���x���ʉȖڃR���{����
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Get_KamokuCbo(p_sKamokuCBO,p_sOption)

    Dim w_sSQL,w_Rs,w_iRet
    Dim w_NyuNen 		'���w�N�x
	Dim WEB0391_Flg

    On Error Resume Next
    Err.Clear

    f_Get_KamokuCbo = 1
    p_sKamokuCBO = ""
    p_sOption = ""
	WEB0391_Flg = false

	m_iNendo = cint(gf_SetNull2Zero(m_iNendo))
	m_sGakunen = cint(gf_SetNull2Zero(m_sGakunen))
	m_sClass = cint(gf_SetNull2Zero(m_sClass))

	If (m_iNendo = 0 OR m_sGakunen = 0 OR m_sClass = 0) then 
		p_sKamokuCBO = "<option value=''>�Ȗڂ�����܂���</option>" & vbCrLf
		p_sOption = " DISABLED"
	    f_Get_KamokuCbo = 0
		exit Function
	End If
        '================
        '//�Ȗڏ��擾
        '================
	w_NyuNen = m_iNendo - cInt(m_sGakunen) + 1

        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_KAMOKU_CD, "
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_KAMOKUMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_LEVEL_FLG"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS M05,T15_RISYU T15"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  M05.M05_GAKKA_CD = T15.T15_GAKKA_CD AND "
        w_sSQL = w_sSQL & vbCrLf & "  M05.M05_NENDO="      & cInt(m_iNendo) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  M05.M05_GAKUNEN="    & cInt(m_sGakunen)  & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  M05.M05_CLASSNO="    & cInt(m_sClass)  & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_KAISETU" 	 & cInt(m_sGakunen)  & " < 3 AND"
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_LEVEL_FLG = 1 AND"
        w_sSQL = w_sSQL & vbCrLf & "  T15.T15_NYUNENDO= "  & w_NyuNen

'response.write w_sSQL & "<br>"
'response.end
	'���R�[�h�Z�b�g�̎擾���s
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_KamokuCbo = 99
            Exit Function
        End If

	'�Ȗڂ����Ȃ������ꍇ
	If w_Rs.EOF Then 
		p_sKamokuCBO = "<option value=''>�Ȗڂ�����܂���</option>" & vbCrLf
		p_sOption = " DISABLED"
	    f_Get_KamokuCbo = 0
		exit Function
	End If

	'�f�[�^������΁A�R���{�{�b�N�X�𐶐�
	Do Until w_Rs.EOF 
		If m_sKengen = C_WEB0390_ACCESS_SENMON then 
			If f_KyokanData(w_Rs("T15_KAMOKU_CD")) = true then 
				p_sKamokuCBO = p_sKamokuCBO & "<option value='" & w_Rs("T15_KAMOKU_CD") & "'>"
				p_sKamokuCBO = p_sKamokuCBO & w_Rs("T15_KAMOKUMEI")
				p_sKamokuCBO = p_sKamokuCBO & "</option>" & vbCrLf
				WEB0391_Flg = true
			End If
		Else
				p_sKamokuCBO = p_sKamokuCBO & "<option value='" & w_Rs("T15_KAMOKU_CD") & "'>"
				p_sKamokuCBO = p_sKamokuCBO & w_Rs("T15_KAMOKUMEI")
				p_sKamokuCBO = p_sKamokuCBO & "</option>" & vbCrLf
		End If
		w_Rs.MoveNext
	Loop

	If m_sKengen = C_WEB0390_ACCESS_SENMON then 
			If WEB0391_Flg = false then 
				p_sKamokuCBO = "<option value=''>�Ȗڂ�����܂���</option>" & vbCrLf
				p_sOption = " DISABLED"
			    f_Get_KamokuCbo = 0
				exit Function
			End If
	End If

    '����I��
    f_Get_KamokuCbo = 0

End Function

Function f_KyokanData(p_sKamokuCD)
'******************************************************************
'�@�@�@�\�F�����̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
	Dim w_rs,w_iRet,w_sSQL

    On Error Resume Next
    Err.Clear
    f_KyokanData = false

    Do


        '//�Ȗڂ̃f�[�^�擾
        w_sSQL = ""
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     T27_KYOKAN_CD"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "     T27_TANTO_KYOKAN"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "     T27_NENDO = " & m_iNendo & " "
        w_sSQL = w_sSQL & vbCrLf & " AND T27_GAKUNEN = " & m_sGakunen & " "
        w_sSQL = w_sSQL & vbCrLf & " AND T27_KAMOKU_CD = '" & p_sKamokuCD & "' "
        w_sSQL = w_sSQL & vbCrLf & " AND T27_KYOKAN_CD = '" & m_sKyokanCd & "' "

'response.write w_sSQL & vbCrLf & "<BR>"

        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, w_sSQL,m_iDsp)

'response.write "w_iRet = " & w_iRet & "<BR>"
'response.write w_rs.EOF & "<BR>"& vbCrLf
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            Exit Do 
        End If

        If w_rs.EOF Then
            'ں��޾�Ă̎擾���s
            Exit Do 
        End If

    f_KyokanData = true

    Exit Do

    Loop

   Call gf_closeObject(w_rs)

End Function

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

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	Dim w_sKamokuCBO	'���x���ʉȖڃR���{
	Dim w_sOption		'���x���ʉȖڃR���{�I�v�V����
	Dim w_iRet
	w_iRet = f_Get_KamokuCbo(w_sKamokuCBO,w_sOption)
    If w_iRet <> 0 Then m_bErrFlg = True : Exit Sub

%>
<html>

<head>

<title>���x���ʉȖڌ���</title>
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

        document.frm.action="web0390_top.asp";
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
        // ��
        if( f_Trim(document.frm.cboKamokuCode.value) == "" ){
            window.alert("�Ȗڂ��I������Ă��܂���B");
//            document.frm.cboKamokuCode.focus();
            return ;
        }

		//�w�N�A�N���X���Z�b�g
		document.frm.txtGakunen.value = document.frm.cboGakunenCd.value
		document.frm.txtClass.value =document.frm.cboClassCd.value

        document.frm.action="web0390_main.asp";
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

		document.frm.cboGakunenCd.value = ""
		document.frm.cboClassCd.value = ""

		f_ReLoadMyPage();
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
<%call gs_title("���x���ʉȖڌ���","��@��")%>
<br>
            <table border="0">
                <tr>
                    <td class="search">
                        <table border="0" cellpadding="1" cellspacing="1">
                            <tr>
                                <td align="left">
                                    <table border="0" cellpadding="1" cellspacing="1">

                                        <tr>
                                            <td Nowrap align="left">�N���X
											<% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' " &  m_sGakunenOption ,false,m_sGakunen) %>�@
											<% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere," style='width:80px;' " & m_sClassOption,false,m_sClass) %>
                                            </td>
                                            <td Nowrap align="left">�ȁ@��
											<Select name="cboKamokuCode" style='width:200px;'<%=w_sOption%>>
												<%=w_sKamokuCBO%>
 											</Select>
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
