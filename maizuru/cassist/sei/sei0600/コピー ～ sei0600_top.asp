<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ȓ����o�^
' ��۸���ID : gak/sei0600/sei0600_top.asp
' �@      �\: ��y�[�W �����������o�^�̌������s��
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
' ��      ��: 2001/09/26 �J�e
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
    Public m_iSikenKBN   '�����敪
    Public m_sGakuNo        '�����R���{�{�b�N�X�ɓ���l
    Public m_sGakuNoWhere   '�����R���{�{�b�N�X��where����

    Public  m_Rs
    Public  m_Rs_Siken
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
    w_sMsgTitle="���ȓ����o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakuNo   = request("txtGakuNo")
    m_iSikenKBN   = request("txtSikenKBN")
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

        '//�w�N�̑Ώۂ̃f�[�^�擾
        w_iRet = f_getData()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

        Call f_GakuNoWhere()
        
		'=====================
		'//�����R���{���擾
		'=====================
        w_iRet = f_GetSiken()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'==============================================
		'//�����敪��Ă���(�Ȗڎ擾���Ɏg�p)
		'==============================================
		If Request("txtSikenKBN")  = "" Then

			'//�ŋ߂̎����敪���擾
            w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_SEISEKI_KIKAN,m_rs("M05_GAKUNEN"))
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If

		Else
		    m_iSikenKbn = Request("txtSikenKBN")    '//�R���{�����敪
		End If

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
'*  [�@�\]  �����R���{���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSiken()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & "  M01_SYOBUNRUI_CD"
		w_sSQL = w_sSQL & vbCrLf & " ,M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_SYOBUNRUI_CD < " & cint(C_SIKEN_JITURYOKU)
		w_sSQL = w_sSQL & vbCrLf & "  ORDER BY M01_SYOBUNRUI_CD"

        iRet = gf_GetRecordset(m_Rs_Siken, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If


        f_GetSiken = 0
        Exit Do
    Loop

End Function

Function f_Selected(pData1,pData2)
'****************************************************
'[�@�\] �f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'[����] pData1 : �f�[�^�P
'       pData2 : �f�[�^�Q
'[�ߒl] f_Selected : "SELECTED" OR ""
'****************************************************

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else 
            f_Selected = "" 
        End If
    End If

End Function

Function f_getData()
'********************************************************************************
'*  [�@�\]  �w�N�̑Ώۂ̃f�[�^�擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_getData = 1

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     M05_CLASS "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     M05_NENDO = '" & m_iNendo & "' "
        w_sSQL = w_sSQL & " AND M05_TANNIN = '" & m_sKyokanCd & "' "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            f_getData = 99
            m_bErrFlg = True
            Exit Do 
        End If

        f_getData = 0
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

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

<title>���ȓ����o�^</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

        document.frm.action="sei0600_main.asp";
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
<%call gs_title("���ȓ����o�^","�o�@�^")%>
<br>
    <table border="0">
    <tr>
    <td class="search">
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left">
            <table border="0" cellpadding="1" cellspacing="1">
	                    <tr valign="middle">
	                        <td align="left">�����敪</td>
	                        <td align="left">
								<%If m_Rs_Siken.EOF Then%>
									<select name="txtSikenKBN" style='width:150px;' DISABLED>
										<option value="">�f�[�^������܂���
								<%Else%>
									<select name="txtSikenKBN" style='width:150px;' onchange = 'javascript:f_ReLoadMyPage()'>
									<%Do Until m_Rs_Siken.EOF%>
										<option value='<%=m_Rs_Siken("M01_SYOBUNRUI_CD")%>'  <%=f_Selected(cstr(m_Rs_Siken("M01_SYOBUNRUI_CD")),cstr(m_iSikenKbn))%>><%=m_Rs_Siken("M01_SYOBUNRUIMEI")%>
										<%m_Rs_Siken.MoveNext%>
									<%Loop%>
								<%End If%>
								</select>
							</td>
            <td Nowrap align="center">�@�N���X�@</td>
            <td Nowrap><%=m_Rs("M05_GAKUNEN")%>�N�@<%=m_Rs("M05_CLASSMEI")%></td>
            </tr>
			<tr>
		        <td colspan="4" align="right">
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
</td>
</tr>
</table>
	<input type="hidden" name="txtGakunen" value="<%=m_Rs("M05_GAKUNEN")%>">
	<input type="hidden" name="txtClass" value="<%=m_Rs("M05_CLASSNO")%>">
	<input type="hidden" name="txtClassNm" value="<%=m_Rs("M05_CLASSMEI")%>">
</form>

</center>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
