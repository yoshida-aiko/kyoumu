<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: ���N�Y���҈ꗗ
'* ��۸���ID : han/han0121/ichiran.asp
'* �@      �\: ���y�[�W ���N�Y���҈ꗗ���X�g�\�����s��
'*-------------------------------------------------------------------------
'* ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*           cboGakunenCd      :�w�N�R�[�h
'*           :session("PRJ_No")      '���������̃L�[
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'* ��      ��:
'*           �I�����ꂽ�N���X�̗��N�Y���҈ꗗ��\��
'*-------------------------------------------------------------------------
'* ��      ��: 2001/08/08 �O�c�@�q�j
'* ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_sMsg              'ү����
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iNendo         ':�����N�x
    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iGakunen          ':�w�N�R�[�h
    
    Public  m_Rs                'recordset
    Public  m_sMode             '���[�h
    
    '�y�[�W�֌W
    Public  m_iMax              ':�ő�y�[�W
    Public  m_iDsp              '// �ꗗ�\���s��
    Public  m_iPageCD

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
    Dim w_sWHERE            '// WHERE��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���N�Y���҈ꗗ"
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

		'//���X�g�̏ڍ׃f�[�^�擾
		w_iRet = f_getdate()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

		If m_Rs.EOF Then
	        '// �y�[�W��\��
	        Call NO_showPage()
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
    gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()    '//2001/07/30�ύX

    m_iKyokanCd = Request("txtKyokanCd")          ':�����R�[�h
    m_iNendo = Request("txtNendo")              ':�����N�x
    m_iGakunen = Request("cboGakunenCd")   ':�w�N�R�[�h
    m_sMode = Request("txtMode")

    If m_sMode = "Hyouji" Then
        m_iPageCD = 1
    Else
        m_iPageCD = INT(Request("txtPageCd")) ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
    
End Sub

'********************************************************************************
'*  [�@�\]  ���X�g�̏ڍ׎擾
'*  [����]  
'*  [�ߒl]  0:���擾�����A1:���R�[�h�Ȃ��A99:���s
'*  [����]  
'********************************************************************************
Function f_getdate()
    
    On Error Resume Next
    Err.Clear
    
    f_getdate = 1

    Do

        '// �N���X�}�X�^���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & "A.T48_GAKUSEKI_NO,A.T48_SIMEI,B.M02_GAKKAMEI "
        w_sSQL = w_sSQL & vbCrLf & "FROM "
        w_sSQL = w_sSQL & vbCrLf & "T48_RYUNEN A,M02_GAKKA B "
        w_sSQL = w_sSQL & vbCrLf & "WHERE "
        w_sSQL = w_sSQL & vbCrLf & "A.T48_NENDO = " & m_iNendo & " "
        w_sSQL = w_sSQL & vbCrLf & "AND A.T48_GAKUNEN = " & m_iGakunen & " "
        w_sSQL = w_sSQL & vbCrLf & "AND A.T48_NENDO = B.M02_NENDO(+) "
        w_sSQL = w_sSQL & vbCrLf & "AND A.T48_GAKKA_CD = B.M02_GAKKA_CD(+) "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            f_getdate = 99
            Exit Do 'GOTO LABEL_f_GetClassMei_END
        End If

		f_getdate = 0

        Exit Do
    
    Loop
    

'// LABEL_f_GetClassMei_END
End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_iCnt

w_iCnt = 0
    
    On Error Resume Next
    Err.Clear

    '�y�[�WBAR�\��
    Call gs_pageBar(m_Rs,m_iPageCD,m_iDsp,w_pageBar)

%>

<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
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
        document.frm.txtPageCd.value = p_iPage;
        document.frm.submit();
    
    }
-->
</SCRIPT>
</head>

<body>

<center>
<form name ="frm" method="post">

<table border=0 width="500">
<tr>
<td align="center">
<%=w_pageBar %>

	<table border="1" class=hyo width="<%=C_TABLE_WIDTH%>">
		<tr>
			<th class=header width="140">�w �� ��</th>
			<th class=header width="100"><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
			<th class=header width="120">���@�@��</th>
		</tr>

<%
    Do While not m_Rs.EOF
        '//�e�[�u���Z���w�i�F
        call gs_cellPtn(w_cell)
%>
		<tr>
			<td class=<%=w_cell%>><%=m_Rs("M02_GAKKAMEI")%><BR></td>
			<td class=<%=w_cell%>><%=m_Rs("T48_GAKUSEKI_NO")%><BR></td>
			<td class=<%=w_cell%>><%=m_Rs("T48_SIMEI")%><BR></td>
		</tr>
<%
		m_Rs.MoveNext
		If w_iCnt >= C_PAGE_LINE Then
			Exit Do
		Else
			w_iCnt = w_iCnt + 1
        End If
    Loop
%>
    </table>
<%=w_pageBar %>

</td>
</tr>
</table>
	<input type="hidden" name="txtMode" value="<%=m_sMode%>">
	<input type="hidden" name="txtPageCd" value="<%=m_iPageCD %>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_iKyokanCd%>">
	<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">
	<input type="hidden" name="cboGakunenCd" value="<%=m_iGakunen%>">

</form>
</center>

</body>

</html>
<%
    '---------- HTML END   ----------
End Sub

Sub NO_showPage()
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
%>
