<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���w�Z��񌟍�
' ��۸���ID : mst/mst0113/top.asp
' �@      �\: ��y�[�W ���w�Z�}�X�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           :txtQueryKenCd       :�s���{���R�[�h     '/2001/07/30�ǉ�
'           :txtQuerySityoCd     :�s�����R�[�h       '/2001/07/30�ǉ�
'           :txtQueryTyuName     :���w�Z��           '/2001/07/30�ǉ�
'           :txtQueryPageTyu     :�\���ϕ\���Ő�     '/2001/07/30�ǉ�
'           :txtMode             :���[�h             '/2001/07/30�ǉ�
'                                (BLANK)    :�����l
'                                Reload     :�����[�h
'           :txtQueryTyuKbn      :���w�Z�敪         '/2001/07/30�ǉ�
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/30�ǉ�
'           txtKenCd        :���R�[�h
'           txtSityoCd      :�s�����R�[�h
'           txtTyuName      :���w�Z���́i�ꕔ�j
'           txtTyuKbn       :���w�Z�敪
'           txtMode         :���[�h             '/2001/07/30�ǉ�
'                            Search         :����
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ����w�Z��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/15 ���u �m��
' ��      �X: 2001/06/20 �≺ �K��Y�i�d�l�ύX�ɔ����N�x�̍폜�E�X�֔ԍ��̒ǉ��j
'           : 2001/07/26 ���{�@�����i���w�Z�敪�ǉ��j
'           : 2001/07/30 ���{�@�����i�����E���n�ǉ�)
'                                  �i���w�Z���̃e�L�X�g�{�b�N�XMAXLENGTH�ǉ��j
'           : 2001/07/31 ���{�@�����i���w�Z���̈����C���j
'           :                        �֐��������K���Ɋ�ύX
'           : 2001/08/07 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    '�s�����I��p��Where����
    Public m_sKenWhere          '���̏���
    Public m_sSityoWhere        '�s�����R���{�̏���
    Public m_sSityoOption       '�s�����R���{�̃I�v�V����
    Public m_sKenSentakuWhere   '�I�������s���{��
    Public m_sSityoSentakuWhere '�I�������s����
    Public m_sMode              '�I���������[�h
    Public m_sTyuWhere          '���w�Z�敪�̏���
    Public m_sTyuSentakuWhere   '�I���������w�Z�敪
    Public m_sTyuName           '���w�Z���́i�ꕔ�j '/2001/07/31�ǉ�

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
    w_sMsgTitle="���w�Z��񌟍�"
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
        
        m_sMode = request("txtMode")

    If m_sMode = "Search" or m_sMode = "PAGE" Then
        '���Ɋւ���WHRE��Query.String����쐬����
        Call s_QueryKenWhere()  
        '�s�����Ɋւ���WHRE��Query.String����쐬����
        Call s_QuerySityoWhere()
        '���w�Z�Ɋւ���WHRE��Query.String����쐬����
        Call s_QueryTyuWhere()  
    Else

        '���Ɋւ���WHRE���쐬����
        Call s_MakeKenWhere()   
        '�s�����Ɋւ���WHRE���쐬����
        Call s_MakeSityoWhere() 
        '���w�Z�Ɋւ���WHRE���쐬����
        Call s_MakeTyuWhere()   
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

Sub s_MakeKenWhere()'/2001/07/31�ύX
'********************************************************************************
'*  [�@�\]  ���R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sKenWhere=""
    m_sKenSentakuWhere=""
        m_sKenWhere = " M16_NENDO = '" & Session("NENDO") & "' "
        m_sKenSentakuWhere = Request("txtKenCd")
End Sub

Sub s_MakeSityoWhere()'/2001/07/31�ύX
'********************************************************************************
'*  [�@�\]  �s�����R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sSityoWhere=""
    m_sSityoOption=""

    If Request("txtKenCd") <> "" Then
        m_sSityoWhere = "     M12_KEN_CD = '" & Request("txtKenCd") & "' "
        m_sSityoWhere = m_sSityoWhere & " GROUP BY M12_SITYOSON_CD,M12_SITYOSONMEI "
    Else
        m_sSityoOption = " DISABLED "
        m_sSityoWhere  = " M12_Ken_CD = '0' "
    End IF

End Sub

Sub s_MakeTyuWhere()'/2001/07/31�ύX
'********************************************************************************
'*  [�@�\]  ���w�Z�敪�R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sTyuWhere=""
    m_sTyuSentakuWhere=""
    m_sTyuName = ""
        m_sTyuWhere = " M01_NENDO = '" & Session("NENDO") & "' "
        m_sTyuWhere = m_sTyuWhere & " AND M01_DAIBUNRUI_CD = " & C_TYUGAKKO_KBN
        m_sTyuSentakuWhere = Request("txtTyuKbn")
        
        m_sTyuName = Request("txtTyuName")  '/2001/07/31�ǉ�
End Sub


Sub s_QueryKenWhere()'/2001/07/31�ύX
'********************************************************************************
'*  [�@�\]  ���R���{�Ɋւ���WHRE��Query.String����쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    m_sKenWhere=""
    m_sKenSentakuWhere=""

        m_sKenWhere = "     M16_NENDO = '" & Session("NENDO") & "' "
        m_sKenSentakuWhere = Request("txtQueryKenCd")
End Sub

Sub s_QuerySityoWhere()'/2001/07/31�ύX
'********************************************************************************
'*  [�@�\]  �s�����R���{�Ɋւ���WHRE��Query.String����쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sSityoSentakuWhere=""
    m_sSityoWhere=""

    If Request("txtQueryKenCd")<>"" Then
        m_sSityoWhere = "     M12_KEN_CD = '" & Request("txtQueryKenCd") & "' "
        m_sSityoWhere = m_sSityoWhere & " GROUP BY M12_SITYOSON_CD,M12_SITYOSONMEI "
        m_sSityoSentakuWhere = Request("txtQuerySityoCd")
    Else
        m_sSityoOption=" DISABLED "
        m_sSityoWhere = " M12_Ken_CD = '0' "
    End IF

End Sub

Sub s_QueryTyuWhere()'/2001/07/31�ύX
'********************************************************************************
'*  [�@�\]  ���w�Z�R���{�Ɋւ���WHRE��Query.String����쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    m_sTyuWhere=""
    m_sTyuSentakuWhere=""
    m_sTyuName = ""

        m_sTyuWhere = "     M01_NENDO = '" & Session("NENDO") & "' "
        m_sTyuWhere = m_sTyuWhere & " AND M01_DAIBUNRUI_CD = " & C_TYUGAKKO_KBN
        m_sTyuSentakuWhere = Request("txtQueryTyuKbn")
        
        m_sTyuName = Request("txtQueryTyuName") '/2001/07/31�ǉ�
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

<title>���w�Z�}�X�^�Q��</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--
    //************************************************************
    //  [�@�\]  �����C�����ꂽ�Ƃ��A�ĕ\������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./top.asp";
        document.frm.target="top";
        document.frm.txtMode.value = "Reload";
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

        document.frm.action="./main.asp";
        document.frm.target="main";
        document.frm.txtMode.value = "Search";
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

        document.frm.txtTyuKbn.value = "@@@";
        document.frm.txtKenCd.value = "@@@";
        document.frm.txtSityoCd.value = "@@@";
        document.frm.txtTyuName.value = "";
    
    }

    //-->
    </SCRIPT>

    <link rel="stylesheet" href="../../common/style.css" type="text/css">

    </HEAD>

<body>

<center>

<form name="frm" Method="POST" onSubmit="return false" onClick="return false;">
<input type="hidden" name="txtMode" value="">
<div align="center">

<%call gs_title("���w�Z��񌟍�","��@��")%>
<br>
    <table border="0">
    <tr>
	    <td class="search">

			<table border="0" cellpadding="1" cellspacing="1">
				<tr>
					<td Nowrap>��@�@��</td>
					<td Nowrap>
							<%  '���ʊ֐�����w�Z�敪�Ɋւ���R���{�{�b�N�X���o�͂���i�N�x�����j
							        call gf_ComboSet("txtTyuKbn",C_CBO_M01_KUBUN,m_sTyuWhere,"",True,m_sTyuSentakuWhere)
							%>
	                </td>
	                <td Nowrap align="center">�s���{��<img src="../../image/sp.gif" width="15"><!-- <select name="gakunen"> -->
							<%  '���ʊ֐����猧�Ɋւ���R���{�{�b�N�X���o�͂���i�N�x�����j
							        call gf_ComboSet("txtKenCd",C_CBO_M16_KEN,m_sKenWhere,"onchange = 'javascript:f_ReLoadMyPage()' ",True,m_sKenSentakuWhere)
							%>
					</td Nowrap>
	                <td Nowrap align="center">�@�s�����@<!-- <select name="gakka"> -->
							<%  '���ʊ֐�����s�����Ɋւ���R���{�{�b�N�X���o�͂���i�N�x�A���������j�i�������͂���Ă��Ȃ��Ƃ��́ADISABLED�ƂȂ�j
							        call gf_ComboSet("txtSityoCd",C_CBO_M12_SITYOSON,m_sSityoWhere,"style='width:200px;' " & m_sSityoOption,True,m_sSityoSentakuWhere)
							%>
					</td>
	            </tr>
				<tr>
					<td Nowrap>���w�Z����</td>
					<td Nowrap><input type="text" size="20" name="txtTyuName" value="<%=m_sTyuName%>" maxlength="60"></td>
					<td colspan="1" Nowrap><font size="2">�����w�Z���̂̈ꕔ�Ō������܂�</font></td>
					<td valign="bottom" align="right" Nowrap>
			        <input type="button" class="button" value=" �N�@���@�A " onclick="javasript:f_Clear();">
					<input class="button" type="button" value="�@�\�@���@" onClick = "javascript:f_Search()">
					</td>
				</tr>
			</table>

		</td>
	</tr>
</table>


</div>
</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
