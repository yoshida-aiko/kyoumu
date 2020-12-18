<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����������ꗗ
' ��۸���ID : web/web0360/web0360_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:   txtClubCd		:����CD
'
' ��      �n:   txtClubCd		:����CD
'               cboGakunenCd	:�w�N
'               cboClassCd		:�N���XNO
'               txtTyuClubCd	:���w�Z����CD
' ��      ��:
'           �������\��
'               �w�N�A�N���X�A���w�Z�����̃R���{�{�b�N�X��\��
'-------------------------------------------------------------------------
' ��      ��: 2001/08/22 �ɓ����q
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '//��������
    Public m_iKyokanCd          '//�N�x
    Public m_sClubCd			'//����CD
    Public m_iGakunen			'//�w�N
	Public m_iClassNo           '//�N���XNO
	Public m_sTyuClubCd			'//���w�Z�N���uCD

    '//�R���{�pWhere������
    Public m_sClubWhere
    Public m_sGakunenWhere      '//�w�N�̏���
    Public m_sClassWhere        '//�N���X�̏���

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
    w_sMsgTitle="�����������ꗗ"
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

'//�f�o�b�O
'Call s_DebugPrint()

        '//�w�N�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeGakunenWhere() 

        '//�N���X�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeClassWhere() 

        '//���w�Z�N���u�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeClubWhere() 

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
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
	m_sClubCd = ""
	m_iClassNo   = ""
	m_sTyuClubCd = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")
	m_sClubCd   = Request("txtClubCd")
    m_iGakunen  = Request("cboGakunenCd")   '//�w�N

	m_iClassNo   = Request("cboClassCd")	'//�N���X
	m_sTyuClubCd = replace(Request("txtTyuClubCd"),"@@@","")	'//���w�Z�N���uCD

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sClubCd   = " & m_sClubCd   & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
	response.write "m_iClassNo   = " & m_iClassNo   & "<br>"
	response.write "m_sTyuClubCd = " & m_sTyuClubCd & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �w�N�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeGakunenWhere()

    m_sGakunenWhere = ""
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyorinen
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeClassWhere()

    m_sClassWhere = ""
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen

    If m_iGakunen = "" Then
        '//�����\������1�N1�g��\������
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = 1"
    Else
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_iGakunen)
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �N���u�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeClubWhere()

    m_sClubWhere = ""
    m_sClubWhere = m_sClubWhere & " M17_NENDO =" & m_iSyoriNen  '//�����N�x
    m_sClubWhere = m_sClubWhere & " AND M17_BUJYOKYO_KBN = 0"	'//�������󋵋敪

End Sub

'********************************************************************************
'*  [�@�\]  ���������擾����
'*  [����]  p_sClubCd:����CD
'*  [�ߒl]  f_GetClubName�F������
'*  [����]  
'********************************************************************************
Function f_GetClubName(p_sClubCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubName = ""
	w_sClubName = ""

	Do

		'//����CD����̎�
		If trim(gf_SetNull2String(p_sClubCd)) = "" Then
			Exit Do
		End If

		'//���������擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD=" & p_sClubCd

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//������
			w_sClubName = rs("M17_BUKATUDOMEI")
		End If

		Exit Do
	Loop

	'//�߂�l���
	f_GetClubName = w_sClubName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

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
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>�����������ꗗ</title>

    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
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
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

        document.frm.action="./web0360_insMain.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �w�N���ύX���ꂽ�Ƃ��A�{��ʂ��ĕ\��
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./web0360_insTop.asp";
        document.frm.target="topFrame";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post" onClick="return false;">

    <%call gs_title("�����������ꗗ","�o�^")%>
    <center>

	<br>
    <table bordeer="0">
        <tr>
        <td class="search">
            <table border="0">
            <tr><td align="left">�@�\�@<%=f_GetClubName(m_sClubCd)%>�@�����o�^�@�\</td></tr>
            <tr><td>
                <table border="0" cellpadding="1" cellspacing="1">
	                <tr>
		                <td nowrap align="left" >�w�N</td>
		                <td nowrap align="left">
		                    <% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' ",False,m_iGakunen) %>
		                </td>
		                <td nowrap align="left" width="40" >�N���X</td>
		                <td nowrap align="left" >
		                    <% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"style='width:80px;' " & m_sClassOption,true,m_iClassNo) %>
		                </td>
		                <td nowrap align="left"><br></td>
<!--
					</tr>
					<tr>
		                <td nowrap align="left" colspan="2">���w�Z����</td>
		                <td nowrap align="left" colspan="2">
							<% call gf_ComboSet("txtTyuClubCd",C_CBO_M17_BUKATUDO,m_sClubWhere," style='width:140px;'",True,"") %>
		                </td>
-->
				        <td valign="bottom" align="right">
				        <input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();" name="btnShow">
				        </td>
	                </tr>
                </table>

            </td>
            </tr>
            </table>
        </td>
        </tr>
    </table>
    </center>

    <!--�l�n���p-->
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
	<input type="hidden" name="txtClubCd" value="<%=m_sClubCd%>">

    </form>
    </body>
    </html>
<%
End Sub
%>