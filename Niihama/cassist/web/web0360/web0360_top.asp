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
'
' ��      ��:
'           �������\��
'               �N���u�̃R���{�{�b�N�X��\��
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
    Public m_sClubCd

    '//�R���{�pWhere������
    Public m_sClubWhere

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
'call s_DebugPrint()

        '//�N���u�R���{�Ɋւ���WHERE���쐬����
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
    response.write "m_sClubCd   = " & m_sClubCd & "<br>"

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

	    var n = document.frm.txtClubCd.selectedIndex;
		if(document.frm.txtClubCd.options[n].value=="@@@"){
		    alert("�N���u��I�����Ă�������");
			return;
		}

        document.frm.action="./web0360_main.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">

    <center>
    <%call gs_title("�����������ꗗ","��@��")%>

    <table bordeer="0">
        <tr>
        <td class="search">
            <table border="0">
            <tr>
            <td>

                <table border="0" cellpadding="1" cellspacing="1">
	                <tr>
		                <td nowrap align="left">�N���u��</td>
		                <td nowrap align="left" >
							<% call gf_ComboSet("txtClubCd",C_CBO_M17_BUKATUDO,m_sClubWhere," style='width:140px;'",True,cstr(gf_SetNull2String(m_sClubCd))) %>
						</td>
				        <td valign="bottom" align="right">
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
    </center>

    <!--�l�n���p-->
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">

    </form>
    </body>
    </html>
<%
End Sub
%>