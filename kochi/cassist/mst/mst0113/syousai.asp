<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���w�Z�}�X�^
' ��۸���ID : mst/mst0113/syossai.asp
' �@      �\: ���y�[�W ���w�Z�}�X�^�̏ڍו\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtKenCd        :���R�[�h
'           txtSityoCd      :�s�����R�[�h
'           txtTyuName      :���w�Z���́i�ꕔ�j
'           txtPageTyu      :�\���ϕ\���Ő��i�������g����󂯎������j
'           txtTyuCd        :�I�����ꂽ���w�Z�R�[�h
'           txtTyuKbn       :�I�����ꂽ���w�Z�敪
'           txtMode         :���[�h
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           txtKenCd        :���R�[�h�i�߂�Ƃ��j
'           txtSityoCd      :�s�����R�[�h�i�߂�Ƃ��j
'           txtTyuName      :���w�Z���́i�߂�Ƃ��j
'           txtPageTyu      :�\���ϕ\���Ő��i�߂�Ƃ��j
'           txtTyuKbn       :���w�Z�敪�i�߂�Ƃ��j
'           txtMode         :���[�h
' ��      ��:
'           �������\��
'               �w�肳�ꂽ���w�Z�̏ڍ׃f�[�^��\��
'           ���n�}�摜�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ����w�Z�n�}��\������i�ʃE�B���h�E�j
'-------------------------------------------------------------------------
' ��      ��: 2001/06/16 ���u �m��
' ��      �X: 2001/07/26 ���{�@�����@'DB�ύX�ɔ����C��
'             2001/07/31 ���{  ����  �ϐ��������K���Ɋ�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_sKenCd        ':���R�[�h
    Public  m_sSityoCd      ':�s�����R�[�h
    Public  m_sTyuName      ':���w�Z���́i�ꕔ�j
    Public  m_iPageTyu      ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_sTyuCd        ':�I�����ꂽ���w�Z�R�[�h
    Public  m_Rs            'recordset
    Public  m_iNendo        ':�N�x      '//2001/07/31�ύX
    Public  m_sMode         ':���[�h
    Public  m_iTyuKbn       ':���w�Z�敪
    
    Public  m_iTyuKbnD      ':���w�Z�敪(DB)
    Public  m_iJyoKbnD      ':�w�Z�󋵋敪(DB)
    Public  m_sTyuKbnMei    ':���w�Z�敪��
    Public  m_sJyoKbnMei    ':�w�Z�󋵋敪��
    

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
    w_sMsgTitle="���w�Z��񌟍�"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sMode = request("txtMode")

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

        '���w�Z�}�X�^���擾
        w_sWHERE = ""

        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M13.M13_TYUGAKKO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKOMEI "     
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TEL "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_YUBIN_BANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_GAKKOJYOKYO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TIZUFILENAME "
        w_sSQL = w_sSQL & vbCrLf & " ,M12.M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M16.M16_KENMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM M13_TYUGAKKO M13 "
        w_sSQL = w_sSQL & vbCrLf & " , M16_KEN M16 "
        w_sSQL = w_sSQL & vbCrLf & " , M12_SITYOSON M12 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & "      M13.M13_KEN_CD = M16.M16_KEN_CD (+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_KEN_CD = M12.M12_KEN_CD (+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_SITYOSON_CD = M12.M12_SITYOSON_CD (+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13_TYUGAKKO_CD = '" & m_sTyuCd & "' "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_NENDO = " & m_iNendo & ""

'Response.Write w_sSQL & "<br>"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If
        
        '// DB����敪���擾
        Call s_SetDB()
        Call s_GetTyugakkoKbn()
        Call s_GetJyokyoKbn()
        
        
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

    m_iNendo = Session("NENDO")     ':�N�x

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

    m_sTyuName = Request("txtTyuName")      ':���w�Z���́i�ꕔ�j
    m_sTyuCd = Request("txtTyuCd")      ':���w�Z���́i�ꕔ�j


    '// BLANK�̏ꍇ�͍s���ر
    'If Request("txtMode") = "Search" Then
    '    m_iPageTyu = 1
    'Else
        m_iPageTyu = INT(Request("txtPageTyu"))     ':�\���ϕ\���Ő��i�������g����󂯎������j
    'End If

    m_iTyuKbn = Request("txtTyuKbn")        ':���w�Z�敪

End Sub


'********************************************************************************
'*  [�@�\]  DB�l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetDB()

m_iTyuKbnD = m_Rs("M13_TYUGAKKO_KBN")
m_iJyoKbnD = m_Rs("M13_GAKKOJYOKYO_KBN")

End Sub

'********************************************************************************
'*  [�@�\]  ���w�Z�敪���̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_GetTyugakkoKbn()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    m_sTyuKbnMei = ""
    
    On Error Resume Next
    Err.Clear

    Do
        
        '// �敪�}�X�^ں��޾�Ă��擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & " M01_SYOBUNRUIMEI"
        w_sSQL = w_sSQL & " FROM M01_KUBUN "
        w_sSQL = w_sSQL & " WHERE M01_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD = " & C_TYUGAKKO_KBN
        w_sSQL = w_sSQL & " AND M01_SYOBUNRUI_CD = " & gf_SetNull2Zero(trim(m_iTyuKbnD))

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_sTyuKbnMei = "�@"
            'm_sErrMsg = "ں��޾�Ă̎擾���s"
            Exit Do 
        End If
        
        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            m_sTyuKbnMei = "�@"
            'm_sErrMsg = "�Ώ�ں��ނȂ�"
            Exit Do 
        End If
        
        '// �擾�����l���i�[
        m_sTyuKbnMei = w_Rs("M01_SYOBUNRUIMEI")
        '// ����I��
        Exit Do

    Loop

    gf_closeObject(w_Rs)

End Sub

'********************************************************************************
'*  [�@�\]  �w�Z�󋵋敪���̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_GetJyokyoKbn()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    m_sJyoKbnMei = ""
    
    On Error Resume Next
    Err.Clear

    Do
        
        '// �敪�}�X�^ں��޾�Ă��擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & " M01_SYOBUNRUIMEI"
        w_sSQL = w_sSQL & " FROM M01_KUBUN "
        w_sSQL = w_sSQL & " WHERE M01_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD = " & C_GAKKO_JYOKYO
        w_sSQL = w_sSQL & " AND M01_SYOBUNRUI_CD = " & gf_SetNull2Zero(trim(m_iJyoKbnD))

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_sJyoKbnMei = "�@"
            'm_sErrMsg = "ں��޾�Ă̎擾���s"
            Exit Do 
        End If
        
        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            m_sJyoKbnMei = "�@"
            'm_sErrMsg = "�Ώ�ں��ނȂ�"
            Exit Do 
        End If
        
        '// �擾�����l���i�[
        m_sJyoKbnMei = w_Rs("M01_SYOBUNRUIMEI")
        '// ����I��
        Exit Do

    Loop

    gf_closeObject(w_Rs)

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
        document.frm.txtPageTyu.value = p_iPage;
        document.frm.submit();
    
    }

    function f_OpenWindow(p_Url){
    //************************************************************
    //  [�@�\]  �q�E�B���h�E���I�[�v������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
        var window_location;
        window_location=window.open(p_Url,"window","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,scrolling=no,Width=500,Height=500");
        window_location.focus();
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

<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
<tr>
<td valign="top" align="center">

<%call gs_title("���w�Z��񌟍�","�ځ@��")%>

<img src="../../image/sp.gif" height="10"><br>

    <table border="1" class=disp width="400">
        <tr>
            <td class=disph align="left" width="100">���w�Z��</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_TYUGAKKOMEI")) %></td>
        </tr>
        <!-- tr>
            <td class=disph align="left" width="100">���w����</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_TYUGAKKORYAKSYO")) %></td>
        </tr -->
        <tr>
            <td class=disph align="left" width="100">�敪</td>
            <td class=disp align="left" width="300"><%=m_sTyuKbnMei%></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�w�Z��</td>
            <td class=disp align="left" width="300"><%=m_sJyoKbnMei%></td>
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
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_JUSYO1")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z���i�Q�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_JUSYO2")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z���i�R�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_JUSYO3")) %></td>
        </tr -->

        <tr>
            <td class=disph align="left" width="100">�X�֔ԍ�</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_YUBIN_BANGO")) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z��</td>
            <td class=disp align="left" width="300">
                <%=gf_HTMLTableSTR(m_Rs("M13_JUSYO1"))%><BR>
                <%=gf_HTMLTableSTR(m_Rs("M13_JUSYO2"))%>
                <%=gf_HTMLTableSTR(m_Rs("M13_JUSYO3"))%></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�d�b�ԍ�</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_Rs("M13_TEL")) %></td>
        </tr>
        <!--<tr>
            <td class=disph align="left" width="100">�n�}</td>
            <td class=disp align="left" width="300">
<%
    ' �n�}�̗L����\���E�����N����t�@���N�V����
    'Call s_MapHTML()
%>
            </td>
        </tr>-->
    </table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form action="./default.asp" target="<%=C_MAIN_FRAME%>">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtKenCd" value="<%=m_sKenCd%>">
        <input type="hidden" name="txtSityoCd" value="<%=m_sSityoCd%>">
        <input type="hidden" name="txtTyuName" value="<%=m_sTyuName%>">
        <input type="hidden" name="txtPageTyu" value="<%=m_iPageTyu%>">
        <input type="hidden" name="txtTyuCd" value="">
        <input type="hidden" name="txtTyuKbn" value="<%=m_iTyuKbn%>">
    <input class=button type="submit" value="�߁@��">
    </form>
    </td>
    </tr>
    </table>

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










