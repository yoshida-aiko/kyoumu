<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������������o�^
' ��۸���ID : gak/gak0460/gak0460_top.asp
' �@      �\: ��y�[�W �������������o�^�̌������s��
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
' ��      ��: 2001/07/18 �O�c �q�j
' ��      �X: 2001/08/07 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
' ��      �X�F2001/08/30 �ɓ� ���q     ����������2�d�ɕ\�����Ȃ��悤�ɕύX
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
    Public m_sNendo         '�N�x�R���{�{�b�N�X�ɓ���l
    Public m_sGakuNo        '�����R���{�{�b�N�X�ɓ���l
    Public m_sNendoWhere    '�N�x�R���{�{�b�N�X�̏���
    Public m_sGakuNoWhere   '�����R���{�{�b�N�X�̏���
    Public m_sOption        '�����R���{�{�b�N�X�̎g�p�A�s�̔���

    Public m_sGakunen        
    Public m_sClass          
    Public m_sClassNm        

    Public  m_Rs
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
    w_sMsgTitle="�������������o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
	m_sNendo = request("txtNendo")
    m_sGakuNo   = request("txtGakuNo")

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

        w_iRet = f_NendoWhere()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

        Call f_GakuNoWhere()
        
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

Function f_NendoWhere()
'********************************************************************************
'*  [�@�\]  �ݒ�N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_NendoWhere = 1

    Do

        m_sNendoWhere=""
            m_sNendoWhere = " M05_NENDO > " & m_iNendo - 5 & "  AND "
            m_sNendoWhere = m_sNendoWhere & " M05_NENDO <= " & m_iNendo & "  AND "
            m_sNendoWhere = m_sNendoWhere & " M05_TANNIN = '" & m_sKyokanCd & "' "

            m_sNendo = request("txtNendo")

        If request("txtNendo") = C_CBO_NULL Then m_sNendo = ""

        If m_sNendo <> "" Then
            w_sSQL = ""
            w_sSQL = w_sSQL & " SELECT "
            w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
            w_sSQL = w_sSQL & " FROM "
            w_sSQL = w_sSQL & "     M05_CLASS "
            w_sSQL = w_sSQL & " WHERE"
            w_sSQL = w_sSQL & "     M05_NENDO = '" & m_sNendo & "' "
            w_sSQL = w_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "

            Set m_Rs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
            If w_iRet <> 0 Then
                'ں��޾�Ă̎擾���s
                f_NendoWhere = 99
                m_bErrFlg = True
                Exit Do 
            End If

			m_sGakunen	= m_Rs("M05_GAKUNEN")
			m_sClass	= m_Rs("M05_CLASSNO")
			m_sClassNm	= m_Rs("M05_CLASSMEI")

        End If

        f_NendoWhere = 0
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub f_GakuNoWhere()
'********************************************************************************
'*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sGakuNoWhere=""
    m_sOption=""

    If m_sNendo <> "" Then
        If m_Rs.EOF Then
            m_sOption = " DISABLED "
            m_sGakuNoWhere  = " T11_GAKUSEI_NO = '' "
        Else
            m_sGakuNoWhere = " T11.T11_GAKUSEI_NO = T13.T13_GAKUSEI_NO AND "
'            m_sGakuNoWhere = m_sGakuNoWhere & " T11.T11_NYUNENDO = T13.T13_NENDO - T13.T13_GAKUNEN + 1 AND "
            m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_GAKUNEN = " & m_sGakunen & " AND "
            m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_CLASS = " & m_sClass & " AND "
            m_sGakuNoWhere = m_sGakuNoWhere & " T13.T13_NENDO = " & m_sNendo & " "
        End If
    Else
        m_sOption = " DISABLED "
        m_sGakuNoWhere  = " T11_GAKUSEI_NO = '' "
    End IF

End Sub

Sub f_Syosai()
'********************************************************************************
'*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    If m_sNendo = "" Then
%>
        <td width="48" Nowrap>�@</td>
        <td width="96" Nowrap>�@</td>
<%
    Else
%>
        <td width="48" Nowrap align="right"><%=m_sGakunen%>�N</td>
        <td width="96" Nowrap><%=m_sClassNm%></td>
<%
    End If

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
    <title>�������������o�^</title>
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

        document.frm.action="gak0461_top.asp";
        document.frm.target="topFrame";
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
        if( f_Trim(document.frm.txtNendo.value) == "" ){
            window.alert("�N�x�̑I�����s���Ă�������");
            document.frm.txtNendo.focus();
            return ;
        }
        // ���N�x
        if( f_Trim(document.frm.txtNendo.value) == "<%=C_CBO_NULL%>" ){
            window.alert("�N�x�̑I�����s���Ă�������");
            document.frm.txtNendo.focus();
            return ;
        }
        // ���w��
        if( f_Trim(document.frm.txtGakuNo.value) == "" ){
			if (document.frm.txtGakuNo.length == 1) {
	            window.alert("�w��N�x�̊w���̃f�[�^������܂���");
	            document.frm.txtNendo.focus();
			} else {
	            window.alert("�w���̑I�����s���Ă�������");
    	        document.frm.txtGakuNo.focus();
			}
            return ;
        }

        // ���w��
        if( f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
			if (document.frm.txtGakuNo.length == 1) {
	            window.alert("�w��N�x�̊w���̃f�[�^������܂���");
	            document.frm.txtNendo.focus();
			} else {
	            window.alert("�w���̑I�����s���Ă�������");
    	        document.frm.txtGakuNo.focus();
			}
        	    return ;
        }

        document.frm.action="gak0461_main.asp";
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

        document.frm.txtNendo.value = "";
        document.frm.txtGakuNo.value = "";
    
    }

    //-->
    </SCRIPT>

    <link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>
<body>
<center>
<form name="frm" METHOD="post">
<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
<td valign="top" align="center">
<%call gs_title("�������������o�^","�o�@�^")%>
<br>
    <table border="0">
    <tr>
    <td class="search" Nowrap>
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td Nowrap>
        <%call gf_ComboSet("txtNendo",C_CBO_M05_CLASS_N,m_sNendoWhere,"style='width:70px;' onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sNendo)%></td><td>�N�x</td>
        <%Call f_Syosai()%>
        <td Nowrap>�@���@���@</td>
		<td Nowrap>
        <%call gf_PluComboSet("txtGakuNo",C_CBO_T11_GAKUSEKI_N,m_sGakuNoWhere,"style='width:250px;' "& m_sOption,True,m_sGakuNo)%>
        </td>
        </tr>
		<tr>
	        <td colspan="6" align="right">
	        <input type="button" class="button" value=" �N�@���@�A " onclick="javasript:f_Clear();">
	        <input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
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
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
</form>
</center>
</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>
