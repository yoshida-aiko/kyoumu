<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �g�p���ȏ��o�^�@�폜�m�F���
' ��۸���ID : web/WEB0320/del_kakunin.asp
' �@      �\: ���y�[�W �g�p���ȏ��o�^�̍폜�m�F��ʂ��s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H��R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSinroName        :�A�E�於�́i�ꕔ�j
'           txtPageCD       :�\���ϕ\���Ő��i�������g����󂯎������j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtRenrakusakiCD    :�I�����ꂽ�A����R�[�h
'           txtPageCD       :�\���ϕ\���Ő��i�������g�Ɉ����n�������j
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��A�E�E�i�w���\��
'           �����ցA�߂�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ��A�E�E�i�w��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/14 �≺�@�K��Y
' ��      �X: 2001/08/01 �O�c�@�q�j
' ��      �X: 2001/08/22 �ɓ� ���q ������I���ł���悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_sNendo         '// �N�x
    Public  m_sKyokan_CD     ':�����R�[�h
    Public  m_sMode          ':���[�h
'    Public  w_sDelKyokasyoCD
    Public  m_Rs             'recordset
'    Public  w_iDisp          ':�\�������̍ő�l���Ƃ�
    Public  m_sPageCD        ':�y�[�W��
    Public  m_sNo            ':default�̍폜�Ƀ`�F�b�N���ꂽ���̂̔z��
    Public  m_sGakka         '�w�Ȗ���

    '�y�[�W�֌W
    Public  m_cell
    Public  m_iMax      ':�ő�y�[�W
    Public  m_iDsp      '// �ꗗ�\���s��

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
    w_sMsgTitle="�g�p���ȏ��o�^�@�폜�m�F���"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

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

		'//�폜����f�[�^�擾
		w_iRet = f_syousaiitiran()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

        If m_Rs.EOF Then
            '// �y�[�W��\��
            Call showPage_NoData()
        Else
            '// �y�[�W��\��
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
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    'm_sNendo        = request("txtNendo")
    m_sNendo        = request("KeyNendo")
    'm_sKyokan_CD    = request("txtKyokanCd")
    m_sKyokan_CD    = request("SKyokanCd1")

    m_sMode         = Request("txtMode")
    m_sPageCD       = Request("txtPageCD")
	m_sNo           = request("deleteNO")
'    w_iDisp  = Request("txtDisp")           ':�y�[�W�ő�l

End Sub


Function f_syousaiitiran()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_syousaiitiran = 1

    Do

	    w_sSQL = w_sSQL & vbCrLf & " SELECT "
	    w_sSQL = w_sSQL & vbCrLf & "  T47.T47_GAKUNEN "         ''�w�N
	    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_GAKKA_CD "        ''�w��
	    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_KYOKASYO "        ''���ȏ���
	    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_SYUPPANSYA "      ''�o�Ŏ�
	    w_sSQL = w_sSQL & vbCrLf & " ,T47.T47_TYOSYA "          ''����
	    w_sSQL = w_sSQL & vbCrLf & " ,M02.M02_GAKKAMEI "
	    w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI "
	    w_sSQL = w_sSQL & vbCrLf & " FROM "
	    w_sSQL = w_sSQL & vbCrLf & "    T47_KYOKASYO T47 "
	    w_sSQL = w_sSQL & vbCrLf & "    ,M02_GAKKA M02 "
	    w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
	    w_sSQL = w_sSQL & vbCrLf & "    ,M04_KYOKAN M04 "
	    w_sSQL = w_sSQL & vbCrLf & " WHERE "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NO IN (" & Trim(m_sNo) & ") AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M02.M02_NENDO(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_GAKKA_CD  = M02.M02_GAKKA_CD(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M03.M03_NENDO(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO  = M04.M04_NENDO(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = M04.M04_KYOKAN_CD(+) AND "
	    w_sSQL = w_sSQL & vbCrLf & "    T47.T47_NENDO = " & m_sNendo
	    'w_sSQL = w_sSQL & vbCrLf & "    T47.T47_KYOKAN = '" & m_sKyokan_CD & "' "
	    w_sSQL = w_sSQL & vbCrLf & " ORDER BY T47.T47_GAKKA_CD "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If

        f_syousaiitiran = 0

        Exit Do
    Loop
    
    'LABEL_showPage_OPTION_END
End Function

Sub S_syousai()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	Do Until m_Rs.EOF
		Call gs_cellPtn(m_cell)
        %>
        <tr>
	        <td align="center" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_GAKUNEN")) %>�N</td>
        <%
        if cstr(gf_HTMLTableSTR(m_Rs("T47_GAKKA_CD"))) = cstr(C_CLASS_ALL) then
            m_sGakka="�S�w��"
        else
            m_sGakka=gf_HTMLTableSTR(m_Rs("M02_GAKKAMEI"))
        end if
        %>
	        <td align="left" class=<%=m_cell%>><%=m_sGakka %></td>
	        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M03_KAMOKUMEI")) %></td>
	        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_KYOKASYO")) %></td>
	        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("T47_SYUPPANSYA")) %></td>
        </tr>

        <%
    m_Rs.MoveNext
	Loop
End sub

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

    On Error Resume Next
    Err.Clear
%>

<html>
    <head>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  syosai_frm�ւ̃p�����[�^�̎󂯓n��
    //  [����]  p_sSyuseiCD
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Henko(p_sSyuseiCD){

        document.frm.action="syusei.asp";
        document.frm.target="";
        document.frm.txtRenrakusakiCD.value = p_sSyuseiCD;
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    }

    //************************************************************
    //  [�@�\]  �폜�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_delete(){

        if (!confirm("<%=C_SAKUJYO_KAKUNIN%>")) {
           return ;
        }

        document.frm.action="./delete.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "Delete";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Back(){
        document.frm.action="./default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtMode.value = "Back";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
</head>
<body>

<center>

<%

If m_sMode = "DELETE" Then
  m_sSubtitle = "��@��"
End If

call gs_title("�g�p���ȏ��o�^",m_sSubtitle)
%>
<br>
�g�@�p�@���@�ȁ@��
<br><br>
<form name="frm" action="" target="" method="post">
<table border="1" class=hyo width="75%">
    <tr>
	    <th class=header>�w�N</th>
	    <th class=header>�w��</th>
	    <th class=header>�Ȗ�</th>
	    <th class=header>���ȏ���</th>
	    <th class=header>�o�Ŏ�</th>
    </tr>

    <% S_syousai() %>

</table>
<br>
�ȏ�̓��e���폜���܂��B
<br><br>
<table border="0" width="75%">
	<tr>
		<td align=center colspan=5>
		<input type="button" class=button value="�@��@���@" onclick="f_delete()">
		<input type="button" class=button value="�L�����Z��" onclick="f_Back()">
		</td>
	</tr>
</table>
	<input type="hidden" name="txtMode" value="">
	<input type="hidden" name="txtDelKyokasyoCD" value="<%= w_sDelKyokasyoCD %>">
	<input type="hidden" name="txtNendo" value="<%= m_sNendo %>">
	<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
	<input type="hidden" name="txtDisp" value="<%= w_iDisp %>">
    <input type="hidden" name="txtNo" value="<%=m_sNo%>">

    <input type="hidden" name="SKyokanCd1" value="<%=m_sKyokan_CD%>">

</form>

</center>

</body>

</html>





<%
    '---------- HTML END   ----------
End Sub

Sub NoDataPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
    <html>
    <head>
    </head>

    <body>

    <center>
        �폜�̑ΏۂƂȂ�f�[�^���I������Ă��܂���B<br><br><br>
    <input type="button" class=button value="�߁@��" onclick="javascript:history.back()">
    </center>

    </body>

    </html>
<%
End Sub
%>