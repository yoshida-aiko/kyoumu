<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����ʎ��Ǝ��Ԉꗗ
' ��۸���ID : jik/jik0200/top.asp
' �@      �\: ��y�[�W �����ʎ��Ǝ��Ԃ̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           :session("PRJ_No")      '���������̃L�[
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           cboKyokaKeiCd   :�Ȗڌn��R�[�h
'           cboKyokanCd     :�����R�[�h
'           txtMode         :���샂�[�h
'                            (BLANK)    :�����\��
'                            Reload     :�����[�h
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͉Ȗڌn��Ƌ�����\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏ��ƈꗗ��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/03 ���{ ����
' ��      �X: 2001/07/30 ���{ ����  �߂��URL�ύX
' �@      �@:                       �萔���Ή��ɂ��C��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen         '�����N�x
    Public  m_iKyokanCd         '�����R�[�h
    Public  m_sKyokanName       '�����R�[�h
    
    '�f�[�^�擾�p

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
    w_sMsgTitle="�����ʎ��Ǝ��Ԉꗗ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False


        '// ���Ұ�SET
        Call s_SetParam()
         

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

'response.write "AB"

        '//�������̂��擾
        w_iRet = f_GetKyokanNm(m_iKyokanCd,m_iSyoriNen,m_sKyokanName)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

'response.write "CD"

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

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iSyoriNen = Session("NENDO")
    'm_iSyoriNen = 2001      '//�e�X�g�p
    m_iKyokanCd = Session("KYOKAN_CD")

    '//����CD
    If Request("SKyokanCd1") = "" Then
        m_iKyokanCd = Session("KYOKAN_CD")
    Else
        m_iKyokanCd = Request("SKyokanCd1")
    End If

End Sub

Function f_GetKyokanNm(p_sCD,p_iNENDO,p_sName)
'********************************************************************************
'*  [�@�\]  �����̎������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sName
'*  [����]  
'********************************************************************************
Dim rs
Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetKyokanNm = 1
    w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M04_KYOKAN_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M04_NENDO = " & p_iNENDO & " "

'response.write w_sSQL

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKyokanNm = 99
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M04_KYOKANMEI_SEI") & "�@" & rs("M04_KYOKANMEI_MEI")
        End If

        f_GetKyokanNm = 0
        Exit Do
    Loop

    p_sName = w_sName

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
<link rel=stylesheet href="../../common/style.css" type=text/css>
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
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){
        if(document.frm.SKyokanNm1.value==""){
			alert("������I�����Ă�������")
			return;
		}
        document.frm.action="main.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �����Q�ƑI����ʃE�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function KyokanWin(p_iInt,p_sKNm) {
		var obj=eval("document.frm."+p_sKNm)

        URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=610,top=0,left=0");
        nWin.focus();
        return true;    
    }

    //************************************************************
    //  [�@�\]  �N���A�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function fj_Clear(){

		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}
    //-->
    </SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" action="./main.asp" target="main" Method="POST">
    <input type="hidden" name="txtMode">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">
    <%call gs_title("�����ʎ��Ǝ��Ԉꗗ","��@��")%>
<%
If m_sMode = "" Then
%>
<br>
    <table border="0">
    <tr>
    <td class=search>
        <table border="0" cellpadding="1" cellspacing="1">
        <tr>
        <td align="left">
            <table border="0" cellpadding="1" cellspacing="1">
            <tr>
	            <td align="left" nowrap>
	            ����
	            </td>
	            <td align="left" nowrap colspan="2">
	                <input type="text" class="text" name="SKyokanNm1" VALUE='<%=m_sKyokanName%>' readonly>
	                <input type="hidden" name="SKyokanCd1" VALUE='<%=m_iKyokanCd%>'>
	                <input type="button" class="button" value="�I��" onclick="KyokanWin(1,'SKyokanNm1')">
					<input type="button" class="button" value="�N���A" onClick="fj_Clear()">
	            </td>
		    <td align="right" valign="bottom">�@
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
<%
End IF
%>
</td>
</tr>
</table>
</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
