<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �g�p���ȏ��o�^
' ��۸���ID : web/web0320/default.asp
' �@      �\: �g�p���ȏ��̏ڍ׊m�F
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/09/04 �ɓ��@���q
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n

    Public m_bErrFlg        '�װ�׸�
    Public m_iNendo         '�N�x
    Public m_sKyokan_CD     '����CD
    Public m_sTitle         ''�V�K�o�^�E�C���̕\���p

    ''�ް��\���p
    Public m_sNo
    Public m_sNendo
    Public m_sGakkiCD
    Public m_sGakunenCD
    Public m_sGakkaCD
    Public m_sKamokuCD
    Public m_sCourseCD
    Public m_sKyokan_NAME       '����
    Public m_sKyokasyo_NAME     '���ȏ�
    Public m_sSyuppansya        '�o�Ŏ�
    Public m_sTyosya            '���Җ�
    Public m_sSidousyo          '�w����
    Public m_sKyokanyo          '�����p
    Public m_sBiko              '���l

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

    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E��}�X�^�o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

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

        '// �l��ϐ��ɓ����
        Call s_SetParam()

		'//�\���f�[�^�擾
		if f_GetData() = False then
			exit do
		end if

        '// �����̖��̂��擾����
        if f_GetData_Kyokan() = False then
            exit do
        end if

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

'********************************************************************************
'*  [�@�\]  �l��ϐ��ɓ����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo     = Session("NENDO")
    m_sMode      = Request("txtMode")       ':���[�h
	m_sTitle = "�Q��"
	m_sNo = Request("txtUpdNo")     ''�X�V�pNo�i�[

	'//�ꗗ�\�����y�[�W��ۑ�
    m_sPageCD    = Request("txtPageCD")

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

	response.write "m_iNendo		= " & m_iNendo			& "<br>"
	response.write "m_sMode			= " & m_sMode			& "<br>"
	response.write "m_sTitle		= " & m_sTitle			& "<br>"
	response.write "m_sDBMode		= " & m_sDBMode			& "<br>"
	response.write "m_sPageCD		= " & m_sPageCD			& "<br>"
	response.write "m_sNendo		= " & m_sNendo			& "<br>"
	response.write "m_sNo			= " & m_sNo				& "<br>"
	response.write "m_sKyokan_CD	= " & m_sKyokan_CD		& "<br>"
	response.write "m_sGakkiCD		= " & m_sGakkiCD		& "<br>"
	response.write "m_sGakunenCD	= " & m_sGakunenCD		& "<br>"
	response.write "m_sGakkaCD		= " & m_sGakkaCD		& "<br>"
	response.write "m_sKamokuCD		= " & m_sKamokuCD		& "<br>"
	response.write "m_sCourseCD		= " & m_sCourseCD		& "<br>"
	response.write "m_sKyokan_NAME	= " & m_sKyokan_NAME	& "<br>"
	response.write "m_sKyokasyo_NAME= " & m_sKyokasyo_NAME	& "<br>"
	response.write "m_sSyuppansya	= " & m_sSyuppansya		& "<br>"
	response.write "m_sTyosya		= " & m_sTyosya			& "<br>"
	response.write "m_sSidousyo		= " & m_sSidousyo		& "<br>"
	response.write "m_sKyokanyo		= " & m_sKyokanyo		& "<br>"
	response.write "m_sBiko			= " & m_sBiko			& "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �����̖��̂��擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
function f_GetData_Kyokan()
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    dim w_Rs

    f_GetData_Kyokan = False

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " M04.M04_NENDO "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKAN_CD "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN M04 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    M04_NENDO = " &  m_iNendo & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN_CD = '" & m_sKyokan_CD & "'"

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    Else
        '�y�[�W���̎擾
        m_iMax = gf_PageCount(w_Rs,m_iDsp)
    End If

	m_sKyokan_NAME = ""
	If w_Rs.EOF = False Then
	    m_sKyokan_NAME = w_Rs("M04_KYOKANMEI_SEI") & "  " & w_Rs("M04_KYOKANMEI_MEI")
	End If

    w_Rs.close

    f_GetData_Kyokan = True

end function

'********************************************************************************
'*  [�@�\]  �X�V���̕\���ް����擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
function f_GetData()
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_Rs

    f_GetData = False

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " T47.T47_NENDO "            ''�N�x
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKI_KBN "       ''�w���敪
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKUNEN "         ''�w�N
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKA_CD "        ''�w��
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_COURSE_CD "       ''�������
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KAMOKU "          ''�Ȗں���
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKASYO "        ''���ȏ���
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SYUPPANSYA "      ''�o�Ŏ�
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_TYOSYA "          ''����
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKANYOUSU "     ''�����p��
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SIDOSYOSU "       ''�w������
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_BIKOU "           ''���l
    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKAN "           ''����
    w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKAMEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    T47_KYOKASYO T47 "
    w_sSQL = w_sSQL & vbCrLf & "    ,M02_GAKKA M02 "
    w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
    w_sSQL = w_sSQL & vbCrLf & "    ,M04_KYOKAN M04 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M02.M02_NENDO(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_GAKKA_CD  = M02.M02_GAKKA_CD(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M03.M03_NENDO(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M04.M04_NENDO(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = M04.M04_KYOKAN_CD(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO = " & Request("KeyNendo") & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NO = " & m_sNo & ""

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    Else
        '�y�[�W���̎擾
        m_iMax = gf_PageCount(w_Rs,m_iDsp)
    End If

    m_sNendo   = gf_HTMLTableSTR(w_Rs("T47_NENDO"))
    m_sGakkiCD   = gf_HTMLTableSTR(w_Rs("T47_GAKKI_KBN"))
    m_sGakunenCD = gf_HTMLTableSTR(w_Rs("T47_GAKUNEN"))
    m_sGakkaCD   = gf_HTMLTableSTR(w_Rs("T47_GAKKA_CD"))
    m_sKamokuCD  = gf_HTMLTableSTR(w_Rs("T47_KAMOKU"))
    m_sCourseCD  = gf_HTMLTableSTR(w_Rs("T47_COURSE_CD"))
    m_sKyokasyo_NAME  = gf_HTMLTableSTR(w_Rs("T47_KYOKASYO"))       '���ȏ�
    m_sSyuppansya  = gf_HTMLTableSTR(w_Rs("T47_SYUPPANSYA"))        '�o�Ŏ�
    m_sTyosya  = gf_HTMLTableSTR(w_Rs("T47_TYOSYA"))                '���Җ�
    m_sSidousyo  = gf_HTMLTableSTR(w_Rs("T47_SIDOSYOSU"))           '�w����
    m_sKyokanyo  = gf_HTMLTableSTR(w_Rs("T47_KYOKANYOUSU"))         '�����p
    m_sBiko  = gf_HTMLTableSTR(w_Rs("T47_BIKOU"))                   '���l

    m_sKyokan_CD = gf_HTMLTableSTR(w_Rs("T47_KYOKAN"))
    w_Rs.close
    f_GetData = True

end function

'********************************************************************************
'*  [�@�\]  �w�Ȃ̗��̂��擾
'*  [����]  p_sGakkaCd : �w��CD
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetGakkaNm(p_sGakkaCd)
    Dim w_sSQL              '// SQL��
    Dim w_iRet              '// �߂�l
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakkaNm = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKAMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_sNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_GAKKA_CD='" & p_sGakkaCd & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M02_GAKKAMEI")
		End If 

		Exit do 
	Loop

	'//�߂�l���Z�b�g
	f_GetGakkaNm = w_sName

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'****************************************************
'[�@�\] �R�[�X���̂��擾
'[����] pData1 : �f�[�^�P
'[�ߒl] f_Selected : "SELECTED" OR ""
'****************************************************
Function f_GetCourseNm()
    Dim w_sSQL              '// SQL��
    Dim w_iRet              '// �߂�l
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetCourseNm = ""
	w_sName = ""

	Do

		If Trim(m_sCourseCD) = "" Then
			Exit Do
		End If

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M20_COURSEMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M20_COURSE "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M20_NENDO         =  " & m_sNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M20_GAKKA_CD  = '" & m_sGakkaCD & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND M20_GAKUNEN   =  " & m_sGakunenCD
		w_sSQL = w_sSQL & vbCrLf & "  AND M20_COURSE_CD = '" & m_sCourseCD & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M20_COURSEMEI")
		End If 

		Exit do 
	Loop

	'//�߂�l���Z�b�g
	f_GetCourseNm = w_sName

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*  [�@�\]  �Ȗږ��̂��擾
'*  [����]  p_sGakkaCd : �w��CD
'*          p_sKamokuCd
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKamokuNm(p_sGakkaCd,p_sKamokuCd)
    Dim w_sSQL              '// SQL��
    Dim w_iRet              '// �߂�l
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetKamokuNm = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & m_sNendo

        if cstr(gf_HTMLTableSTR(p_sGakkaCd)) <> cstr(C_CLASS_ALL) then
			w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & p_sGakkaCd & "'"
		End If
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_sKamokuCd & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("T15_KAMOKUMEI")
		End If 

		Exit do 
	Loop

	'//�߂�l���Z�b�g
	f_GetKamokuNm = w_sName

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

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

<title>�g�p���ȏ��o�^</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  ���C���y�[�W�֖߂�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Back(){
		document.frm.action="./default.asp";
        document.frm.target="";
        document.frm.submit();
    
    }

    //-->
    </script>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">

	</head>
	<body>
	<form name="frm" action="" target="" method="post">

<%'call s_DebugPrint%>

	<center>
	<% call gs_title("�g�p���ȏ��o�^",m_sTitle) %>
	<br>
<table border="0" cellpadding="1" cellspacing="1" width="400">
    <tr>
        <td align="left">
            <table width="100%" border=1 CLASS="hyo">
                <tr>
                <th height="16" width="75" class=header nowrap>�N�@�x</th>
                <td height="16" width="325" class=detail nowrap>
					<%=Request("KeyNendo")%><br>
                </td>
                </tr>

                <tr>
                <th height="16" width="75" class="header" nowrap>�w�@��</th>
				<%Call gf_GetKubunName(C_KAISETUKI,m_sGakkiCD,m_sNendo,w_KubunName)%>
                <td height="16" width="325" class="detail" nowrap><%=w_KubunName%><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>�w�@�N</th>
                <td height="16" width="325" class=detail nowrap><%=m_sGakunenCD%>�N</td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>�w�@��</th>
                <td height="16" width="325" class=detail nowrap><%=f_GetGakkaNm(m_sGakkaCD)%><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>�R�[�X</font></th>
                <td height="16" width="325" class=detail><%=f_GetCourseNm()%><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>�ȁ@��</font></th>
                <td height="16" width="325" class=detail><%=f_GetKamokuNm(m_sGakkaCD,m_sKamokuCD)%><br></td>
                </tr>

                <tr>

                <th height="16" width="80" class=header nowrap>����</font></th>
                <td height="16" width="325" class=detail nowrap><%=m_sKyokan_NAME%><br></td>
                </tr>

                <tr>
                <th height="16" width="80" class=header nowrap>���ȏ���</font></th>
                <td height="16" width="325" class=detail nowrap><%= m_sKyokasyo_NAME %><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>�o�Ŏ�</font></th>
                <td height="16" width="325"  class=detail nowrap><%= m_sSyuppansya %><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>���Җ�</font></th>
                <td height="16" width="325" class=detail nowrap><%= m_sTyosya %><br></td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>�����p</font>
                </th>
                <td height="16" width="325" class=detail nowrap><%= m_sKyokanyo %>��</td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>�w����</font>
                </th>
                <td height="16" width="325" class=detail nowrap><%= m_sSidousyo %>��</td>
                </tr>

                <tr>
                <th height="16" width="75" class=header nowrap>���@�l</font></th>
                <td height="35" width="325" class=detail nowrap valign="top"><%= trim(m_sBiko) %><br></td>
                </TR>
            </TABLE>
        </td>
    </TR>
</TABLE>
		<table border="0" width=300>
		<tr>
		<td valign="top" align=center>
		<input type="Button" class=button value="�L�����Z��" OnClick="f_Back()">
		</td>
		</tr>
		</table>
		</center>
		
	    <input type="hidden" name="txtNendo"     value="<%= Request("txtNendo") %>">
	    <input type="hidden" name="txtGakunenCd" value="<%= Request("txtGakunenCd") %>">
	    <input type="hidden" name="txtGakkaCD"   value="<%= Request("txtGakkaCD") %>">
	    <input type="hidden" name="txtPageCD"    value="<%= Request("txtPageCD") %>">
		
		<input type="hidden" name="txtMode" value="<%=Request("txtMode")%>">
		
		<input type="hidden" name="hidYear" value="<%=request("hidYear")%>">
		<input type="hidden" name="hidGakunen" value="<%=request("hidGakunen")%>">
		<input type="hidden" name="hidGakka" value="<%=request("hidGakka")%>">
	</form>
	</body>
	</html>

<%
End Sub
%>