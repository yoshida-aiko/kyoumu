<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���w�Z��񌟍�
' ��۸���ID : mst/mst0113/main.asp
' �@      �\: ���y�[�W ���w�Z�}�X�^�̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtKenCd        :���R�[�h
'           txtSityoCd      :�s�����R�[�h
'           txtTyuName      :���w�Z���́i�ꕔ�j
'           txtPageTyu      :�\���ϕ\���Ő��i�������g����󂯎������j
'           txtTyuKbn       :���w�Z�敪
'           txtMode         :���[�h
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           txtTyuCd        :�I�����ꂽ���w�Z�R�[�h
'           txtPageTyu      :�\���ϕ\���Ő��i�������g�Ɉ����n�������j
'           txtMode         :���[�h
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ����w�Z��\��
'           �����ցA�߂�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ����w�Z��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/16 ���u �m��
' ��      �X: 2001/07/31 ���{ ����  �ϐ��������K���Ɋ�ύX
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
    Public  m_iTyuKbn       ':���w�Z�敪
    Public  m_iTyuKbnD      ':���w�Z�敪(DB)
    Public  m_sTyuKbnMei    ':���w�Z�敪��
    Public  m_iPageTyu      ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_iNendo        ':�N�x      '//2001/07/31�ύX
    Public  m_sMode         ':���[�h
    Public  m_Rs            'recordset

    '�y�[�W�֌W
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp                      '// �ꗗ�\���s��

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

        '���w�Z�}�X�^���擾
        w_sWHERE = ""
        
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M13.M13_TYUGAKKO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKOMEI "     
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TEL "
        w_sSQL = w_sSQL & vbCrLf & " ,M13.M13_TYUGAKKO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M12.M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M16.M16_KENMEI "
        'w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM M13_TYUGAKKO M13 "
        w_sSQL = w_sSQL & vbCrLf & " , M16_KEN M16 "
        w_sSQL = w_sSQL & vbCrLf & " , M12_SITYOSON M12 "
        'w_sSQL = w_sSQL & vbCrLf & " , M01_KUBUN M01 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & "      M13.M13_KEN_CD = M16.M16_KEN_CD "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_KEN_CD = M12.M12_KEN_CD "
        w_sSQL = w_sSQL & vbCrLf & "  AND M12.M12_YUBIN_BANGO LIKE '%00' "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_SITYOSON_CD = M12.M12_SITYOSON_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_NENDO = M16.M16_NENDO(+) "
        'w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_NENDO = M01.M01_NENDO(+) "
        'w_sSQL = w_sSQL & vbCrLf & "  AND M01.M01_DAIBUNRUI_CD = " & C_TYUGAKKO_KBN
        'w_sSQL = w_sSQL & vbCrLf & "  AND M13.M13_TYUGAKKO_KBN = M01.M01_SYOBUNRUI_CD(+) "

'response.write w_sSQL

        '���o�����̍쐬
        If m_sKenCd<>"" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M13_KEN_CD = '" & m_sKenCd & "' "
        End If
        If m_sSityoCd<>"" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M13_SITYOSON_CD = '" & m_sSityoCd & "' "
        End If
        If m_sTyuName<>"" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M13_TYUGAKKOMEI Like '%" & m_sTyuName & "%' "
        End If
        If m_iTyuKbn <> "" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M13_TYUGAKKO_KBN = " & m_iTyuKbn
        End If
        
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M13_TYUGAKKO_CD"

'Response.Write w_sSQL & "<br>"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        Else
            '�y�[�W���̎擾
            m_iMax = gf_PageCount(m_Rs,m_iDsp)
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
        w_sSQL = w_sSQL & " AND M01_SYOBUNRUI_CD = " & m_iTyuKbnD
        
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

    m_sMode = Request("txtMode")


    '// BLANK�̏ꍇ�͍s���ر
    If Request("txtMode") = "Search" Then
        m_iPageTyu = 1
    Else
        m_iPageTyu = INT(Request("txtPageTyu"))     ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
    
    m_iTyuKbn = Request("txtTyuKbn")
    '�R���{���I����
    If m_iTyuKbn="@@@" Then
        m_iTyuKbn= ""
    else
        m_iTyuKbn = CInt(m_iTyuKbn)
    End If
    
End Sub

'********************************************************************************
'*  [�@�\]  DB�l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetDB()

m_iTyuKbnD = m_Rs("M13_TYUGAKKO_KBN")

if m_iTyuKbnD = "" Then
    m_iTyuKbnD = 0
end if

End Sub

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
    Dim w_pageBar           '�y�[�WBAR�\���p

    Dim w_iRecordCnt        '//���R�[�h�Z�b�g�J�E���g
    Dim w_iCnt
    Dim w_bFlg
    
    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

    '�y�[�WBAR�\��
    Call gs_pageBar(m_Rs,m_iPageTyu,m_iDsp,w_pageBar)

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

        document.frm.action="main.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageTyu.value = p_iPage;
        document.frm.submit();
    
    }
    
    //************************************************************
    //  [�@�\]  �I���������w�Z�̏ڍׂ�\������B
    //  [����]  p_sTyuCd    :�I���������w�Z�R�[�h
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_GoSyosai(p_sTyuCd){

        document.frm.action="syousai.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtTyuCd.value = p_sTyuCd;
        document.frm.txtMode.value = "<%=m_sMode%>";
        document.frm.txtPageTyu.value = "<%=m_iPageTyu%>";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>

    <link rel=stylesheet href=../../common/style.css type=text/css>

    </head>

    <body>

    <center>
<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr><td align="center">
<br>
<span class=CAUTION>�� ���w�Z�����N���b�N����Əڍׂ��m�F�ł��܂��B</span>
<%=w_pageBar %>
    <table border=1 class=hyo width="100%">
    <COLGROUP WIDTH="15%">
    <COLGROUP WIDTH="15%">
    <COLGROUP WIDTH="40%">
    <COLGROUP WIDTH="15%">
    <COLGROUP WIDTH="15%">
    <tr>
        <th class=header>�敪</th>
        <th class=header>��</th>    
        <th class=header>�s����</th>
        <th class=header>���w�Z��</th>
        <th class=header>�d�b�ԍ�</th>
    </tr>

        <%
        'm_Rs.MoveFirst
        Do While (w_bFlg)
        call gs_cellPtn(w_cell)
        call s_SetDB()
        call s_GetTyugakkoKbn()
        %>
        <tr>
        <td class="<%=w_cell%>" align = "left"><%=m_sTyuKbnMei%></td>
        <td class="<%=w_cell%>" align = "left"><%=m_Rs("M16_KENMEI") %></td>
        <td class="<%=w_cell%>" align = "left"><%=m_Rs("M12_SITYOSONMEI") %></td>
        <td class="<%=w_cell%>" align = "left"><a href="javascript:f_GoSyosai('<%=Trim(m_Rs("M13_TYUGAKKO_CD")) %>')"><%=Trim(m_Rs("M13_TYUGAKKOMEI")) %></a></td>
        <td class="<%=w_cell%>"><font size="2"><%=m_Rs("M13_TEL") %></font></td>
        </tr>
        <%
            m_Rs.MoveNext

            If m_Rs.EOF Then
                w_bFlg = False
            ElseIf w_iCnt >= C_PAGE_LINE Then
                w_bFlg = False
            Else
                w_iCnt = w_iCnt + 1
            End If
        Loop

    'LABEL_showPage_OPTION_END
    %>
        </table>

<%=w_pageBar %>
</td></tr></table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form name ="frm"  Method="POST">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtKenCd" value="<%=m_sKenCd%>">
        <input type="hidden" name="txtSityoCd" value="<%=m_sSityoCd%>">
        <input type="hidden" name="txtTyuName" value="<%=m_sTyuName%>">
        <input type="hidden" name="txtPageTyu" value="<%=m_iPageTyu%>">
        <input type="hidden" name="txtNendo" value="<%= Session("NENDO") %>">
        <input type="hidden" name="txtTyuCd" value="">
        <input type="hidden" name="txtTyuKbn" value="<%=m_iTyuKbn%>">
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
