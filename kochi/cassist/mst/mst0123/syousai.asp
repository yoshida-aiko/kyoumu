<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����w�Z�}�X�^
' ��۸���ID : mst/mst0123/syossai.asp
' �@      �\: ���y�[�W �����w�Z�}�X�^�̏ڍו\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtKenCd        :���R�[�h
'           txtSityoCd      :�s�����R�[�h
'           txtSyuName      :�����w�Z���́i�ꕔ�j
'           txtPageSyu      :�\���ϕ\���Ő��i�������g����󂯎������j
'           txtSyuCd        :�I�����ꂽ�����w�Z�R�[�h
'           txtSyuKbn       :�����w�Z�敪
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           txtKenCd        :���R�[�h�i�߂�Ƃ��j
'           txtSityoCd      :�s�����R�[�h�i�߂�Ƃ��j
'           txtSyuName      :�����w�Z���́i�߂�Ƃ��j
'           txtPageSyu      :�\���ϕ\���Ő��i�߂�Ƃ��j
'           txtSyuKbn       :�����w�Z�敪�i�߂�Ƃ��j
' ��      ��:
'           �������\��
'               �w�肳�ꂽ�����w�Z�̏ڍ׃f�[�^��\��
'           ���n�}�摜�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ������w�Z�n�}��\������i�ʃE�B���h�E�j
'-------------------------------------------------------------------------
' ��      ��: 2001/06/20 �≺�@�K��Y
' ��      �X: 2001/07/26 ���{�@�����@'DB�ύX�ɔ����C��
'           : 2001/07/31 ���{ ����  �ϐ��������K���Ɋ�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_sKubunCd      ':���R�[�h
    Public  m_sKenCd        ':���R�[�h
    Public  m_sSityoCd      ':�s�����R�[�h
    Public  m_sSyuName      ':�����w�Z���́i�ꕔ�j
    Public  m_iPageSyu      ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_sSyuCd        ':�I�����ꂽ���w�Z�R�[�h
    Public  m_Rs            'recordset
    Public  m_iNendo        ':�N�x      '/2001/07/31�ύX
    Public  m_sMode         ':���[�h
    
    Public  m_iSyuKbn       ':�����w�Z�敪


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
    w_sMsgTitle="�����w�Z��񌟍�"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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

        '// ���Ұ�SET
        Call s_SetParam()

        '�����w�Z�}�X�^���擾
        w_sWHERE = ""
        
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M31.M31_GAKKOMEI "        
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_GAKKORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_TEL "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_YUBIN_BANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_TIZUFILENAME "
        w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M12.M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M16.M16_KENMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM M31_SYUSSINKO M31 "
        w_sSQL = w_sSQL & vbCrLf & " , M16_KEN M16 "
        w_sSQL = w_sSQL & vbCrLf & " , M12_SITYOSON M12 "
        w_sSQL = w_sSQL & vbCrLf & " , M01_KUBUN M01 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & "      M01.M01_DAIBUNRUI_CD = " & C_SYUSSINKO
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_KEN_CD = M16.M16_KEN_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_KEN_CD = M12.M12_KEN_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_SITYOSON_CD = M12.M12_SITYOSON_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M31_GAKKO_CD = '" & m_sSyuCd & "' "
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_NENDO = " & m_iNendo & ""
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_NENDO = M01.M01_NENDO(+) " 
        w_sSQL = w_sSQL & vbCrLf & "  AND M31.M31_GAKKO_KBN = M01.M01_SYOBUNRUI_CD(+) "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
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

    m_sKubunCd = Request("txtKubunCd")      ':�敪�R�[�h
    '�R���{���I����
    If m_sKubunCd="@@@" Then
        m_sKubunCd=""
    End If

    m_sKenCd = Request("txtKenCd")          ':���R�[�h
    '�R���{���I����
    If m_sKenCd="@@@" Then
        m_sKenCd=""
    End If

    m_sSityoCd = Request("txtSityoCd")      ':�s�����R�[�h
    '�R���{���I����
    If m_sSityoCd="@@@" Then
        m_sSityoCd=""
    End If

    m_sSyuName = Request("txtSyuName")      ':�����w�Z���́i�ꕔ�j
    m_sSyuCd = Request("txtSyuCd")          ':�����w�Z�R�[�h

    m_iNendo = Session("NENDO")             ':�N�x      '/2001/07/31�ύX
    m_sMode = request("txtMode")            ':���[�h

    '// BLANK�̏ꍇ�͍s���ر
    If Request("txtMode") = "Search" Then
        m_iPageSyu = 1
    Else
        m_iPageSyu = INT(Request("txtPageSyu"))     ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
    
    m_iSyuKbn = Request("txtSyuKbn")        ':�����w�Z�敪

End Sub

''********************************************************************************
''*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
''*  [����]  �Ȃ�
''*  [�ߒl]  �Ȃ�
''*  [����]  
''********************************************************************************
'Sub s_MapHTML()
'
'    If ISNULL(m_Rs("M13_TIZUFILENAME")) OR m_Rs("M13_TIZUFILENAME")="" Then
'        Response.Write("�o�^����Ă��܂���")
'    Else
'        Response.Write("<a Href=""javascript:f_OpenWindow('" & Session("TYUGAKU_TIZU_PATH") & m_Rs("M13_TIZUFILENAME") & "')"">���Ӓn�}</a>")
'    End If
'    
'End Sub
'
'

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
        document.frm.txtPageSyu.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_GoSyosai(p_sTyuCd){

        document.frm.action="syousai.asp";
        document.frm.target="";
        document.frm.txtMode.value = "Syosai";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    </head>

<body>

<center>

<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
<td valign="top" align="center">

<%call gs_title("�����w�Z��񌟍�","�ځ@��")%>

    <table border="1" class=disp width="400">
        <tr>
            <td class=disph align="left" width="100">�����w�Z��</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_GAKKOMEI")) %></td>
        </tr>
        <!-- tr>
            <td class=disph align="left" width="100">�����w�Z����</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_GAKKORYAKSYO")) %></td>
        </tr -->
        <tr>
            <td class=disph align="left" width="100">�敪</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M01_SYOBUNRUIMEI")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�X�֔ԍ�</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_YUBIN_BANGO")) %></td>
        </tr>
        <!-- tr>
            <td class=disph align="left" width="100">��</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M16_KENMEI")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�s�撬��</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M12_SITYOSONMEI")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z���i�P�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_JUSYO1")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z���i�Q�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_JUSYO2")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z���i�R�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_JUSYO3")) %></td>
        </tr -->

        <tr>
            <td class=disph align="left" width="100">�Z��</td>
            <td class=disp align="left" width="300">
                <%=gf_HTMLTableSTR(m_Rs("M31_JUSYO1")) %><BR>
                <%=gf_HTMLTableSTR(m_Rs("M31_JUSYO2")) %>
                <%=gf_HTMLTableSTR(m_Rs("M31_JUSYO3")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�d�b�ԍ�</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M31_TEL")) %></td>
        </tr>
    </table>

</td>
</tr>
</table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form action="./default.asp" target="<%=C_MAIN_FRAME%>">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtKubunCd" value="<%= m_sKubunCd %>">
        <input type="hidden" name="txtKenCd" value="<%= m_sKenCd %>">
        <input type="hidden" name="txtSityoCd" value="<%= m_sSityoCd %>">
        <input type="hidden" name="txtSyuName" value="<%= m_sSyuName %>">
        <input type="hidden" name="txtPageSyu" value="<%= m_iPageSyu %>">
        <input type="hidden" name="txtSyuCd" value="">
        <input type="hidden" name="txtSyuKbn" value="<%=m_iSyuKbn%>">
    <input type="submit" class=button value="�߁@��">
    </form>
    </td>
    </tr>
    </table>

    </center>


    </body>

    </html>





<%
    '---------- HTML END   ----------
End Sub
%>










