<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����w�Z�}�X�^
' ��۸���ID : mst/mst0123/main.asp
' �@      �\: ���y�[�W �����w�Z�}�X�^�̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtKubunCd      :�敪�R�[�h
'           txtKenCd        :���R�[�h
'           txtSityoCd      :�s�����R�[�h
'           txtSyuName      :�����w�Z���́i�ꕔ�j
'           txtPageSyu      :�\���ϕ\���Ő��i�������g����󂯎������j
'           txtSyuKbn       :�����w�Z�敪
'           txtMode
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           txtSyuCd        :�I�����ꂽ�����w�Z�R�[�h
'           txtPageSyu      :�\���ϕ\���Ő��i�������g�Ɉ����n�������j
'           txtMode
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ������w�Z��\��
'           �����ցA�߂�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ������w�Z��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/20 �≺ �K��Y
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
    Public  m_iKubunCd      ':�敪�R�[�h    '/2001/07/31�ύX
    Public  m_sKenCd        ':���R�[�h
    Public  m_sSityoCd      ':�s�����R�[�h
    Public  m_sSyuName      ':�����w�Z���́i�ꕔ�j
    Public  m_sNendo        ':�N�x
    Public  m_sMode         ':���[�h
    Public  m_Rs            'recordset
    Public  m_sPageSyu      ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_iSyuKbn       ':���w�Z�敪
    
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
    w_sMsgTitle="�����w�Z��񌟍�"
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

        '�����w�Z�}�X�^���擾
        w_sWHERE = ""
        
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M31.M31_GAKKO_CD  "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_GAKKOMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M31.M31_TEL "
        w_sSQL = w_sSQL & vbCrLf & " ,M12.M12_SITYOSONMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M16.M16_KENMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & " 	M31_SYUSSINKO M31, "
        w_sSQL = w_sSQL & vbCrLf & " 	M16_KEN M16, "
        w_sSQL = w_sSQL & vbCrLf & " 	("
        w_sSQL = w_sSQL & vbCrLf & " 	select * "
        w_sSQL = w_sSQL & vbCrLf & " 	from "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & " 	where "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_DAIBUNRUI_CD  = " & C_SYUSSINKO & " and "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_NENDO = " & m_sNendo
        w_sSQL = w_sSQL & vbCrLf & " 	) M01, "
        w_sSQL = w_sSQL & vbCrLf & " 	("
        w_sSQL = w_sSQL & vbCrLf & " 	select * "
        w_sSQL = w_sSQL & vbCrLf & " 	from "
        w_sSQL = w_sSQL & vbCrLf & " 		M12_SITYOSON "
        w_sSQL = w_sSQL & vbCrLf & " 	where "
        w_sSQL = w_sSQL & vbCrLf & " 		M12_YUBIN_BANGO LIKE '%00' "
        w_sSQL = w_sSQL & vbCrLf & " 	) M12 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_NENDO = " & m_sNendo & " and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_KEN_CD = M16.M16_KEN_CD (+) and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_NENDO = M16.M16_NENDO(+) and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_KEN_CD = M12.M12_KEN_CD (+) and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_SITYOSON_CD = M12.M12_SITYOSON_CD (+) and " 
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_GAKKO_KBN = M01.M01_SYOBUNRUI_CD (+) and "
        w_sSQL = w_sSQL & vbCrLf & " 	M31.M31_NENDO = M01.M01_NENDO (+) "
        
        '���o�����̍쐬
        If m_iKubunCd <> "" Then
            w_sSQL = w_sSQL & " AND M01.M01_SYOBUNRUI_CD = " & m_iKubunCd   '//2001/07/31�ύX
        End If
        
        If m_sKenCd <> "" Then
            w_sSQL = w_sSQL & " AND M16.M16_KEN_CD = '" & m_sKenCd & "' "
        End If
        
        If m_sSityoCd <> "" Then
            w_sSQL = w_sSQL & " AND M12.M12_SITYOSON_CD = '" & m_sSityoCd & "' "
        End If
        
        If m_sSyuName <> "" Then
            w_sSQL = w_sSQL & " AND M31.M31_GAKKOMEI Like '%" & m_sSyuName & "%' "
        End If
        
        If m_iSyuKbn <> "" Then
            w_sSQL = w_sSQL & vbCrLf & " AND M31_GAKKO_KBN = " & m_iSyuKbn
        End If
		
        w_sSQL = w_sSQL & " ORDER BY M31.M31_GAKKO_CD"
		
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

            If m_sMode = "" Then
            '// �y�[�W��\��
            Call NoPage()
            Exit Do
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
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	
    m_sNendo = Session("NENDO")     ':�N�x
	
    m_iKubunCd = Request("txtKubunCd")  ':�敪�R�[�h
    '�R���{���I����
    If m_iKubunCd="@@@" Then
        m_iKubunCd=""
    End If
	
    m_sKenCd = Request("txtKenCd")      ':���R�[�h
    '�R���{���I����
    If m_sKenCd="@@@" Then
        m_sKenCd=""
    End If

    m_sMode = Request("txtMode")        ':���[�h

    m_sSityoCd = Request("txtSityoCd")  ':�s�����R�[�h
    '�R���{���I����
    If m_sSityoCd="@@@" Then
        m_sSityoCd=""
    End If

    m_sSyuName = Request("txtSyuName")  ':���Z���́i�ꕔ�j
	
    '// BLANK�̏ꍇ�͍s���ر
	If m_sMode = "Search" Then
        m_sPageSyu = 1
    Else
        m_sPageSyu = INT(Request("txtPageSyu"))     ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
    
    m_iSyuKbn = Request("txtSyuKbn")
    '�R���{���I����
    If m_iSyuKbn="@@@" Then
        m_iSyuKbn= ""
    elseif m_iSyuKbn = "" Then
        m_iSyuKbn = ""
    else
        m_iSyuKbn = CInt(m_iSyuKbn)
    End If
	
End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage_NoData()
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

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub NoPage()
%>
	<html>
	<head>
    </head>
	<body>
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
    Call gs_pageBar(m_Rs,m_sPageSyu,m_iDsp,w_pageBar)
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
    //  [�@�\]  �����w�Z�����N���b�N�����ꍇ
    //  [����]  p_sSyuCd :�����w�Z�R�[�h
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_GoSyosai(p_sSyuCd){

        document.frm.action="syousai.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtSyuCd.value = p_sSyuCd;
        document.frm.txtMode.value = "<%=m_sMode%>";
        document.frm.txtPageSyu.value = "<%=m_sPageSyu%>";
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
<span class=CAUTION>�� �����w�Z�����N���b�N����Əڍׂ��m�F�ł��܂��B</span>
<%=w_pageBar %>

        <table border="1" class=hyo width="100%">
        <COLGROUP WIDTH="10%">
        <COLGROUP WIDTH="10%">
        <COLGROUP WIDTH="40%">
        <COLGROUP WIDTH="20%">
        <COLGROUP WIDTH="20%">
        <tr>
        <th class=header>�敪</th>
        <th class=header>��</th>
        <th class=header>�s����</th>
        <th class=header>�����w�Z��</th>
        <th class=header>�d�b�ԍ�</th>
        </tr>

        <%
        Do While (w_bFlg)
        call gs_cellPtn(w_cell)
        %>
        <tr>
        <td class="<%=w_cell%>" align="left"><%=m_Rs("M01_SYOBUNRUIMEI") %></td>
        <td class="<%=w_cell%>" align="left"><%=m_Rs("M16_KENMEI") %></td>
        <td class="<%=w_cell%>" align="left"><%=m_Rs("M12_SITYOSONMEI") %></td>
        <td class="<%=w_cell%>" align="left"><a href="javascript:f_GoSyosai('<%=m_Rs("M31_GAKKO_CD") %>')"><%=Trim(m_Rs("M31_GAKKOMEI"))%></a></td>
        <td class="<%=w_cell%>" align="left"><%=m_Rs("M31_TEL") %></td>
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
	%>
        </table>

<%=w_pageBar %>
</td></tr></table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form name ="frm" action="" target="">
        <input type="hidden" name="txtMode" value="">
        <input type="hidden" name="txtKubunCd" value="<%=m_iKubunCd%>">
        <input type="hidden" name="txtKenCd" value="<%=m_sKenCd%>">
        <input type="hidden" name="txtSityoCd" value="<%=m_sSityoCd%>">
        <input type="hidden" name="txtSyuName" value="<%=m_sSyuName%>">
        <input type="hidden" name="txtPageSyu" value="<%=m_sPageSyu%>">
        <input type="hidden" name="txtNendo" value="<%= Session("NENDO") %>">
        <input type="hidden" name="txtSyuCd" value="">
        <input type="hidden" name="txtSyuKbn" value="<%=m_iSyuKbn%>">
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
