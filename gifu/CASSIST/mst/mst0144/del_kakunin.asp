<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A�E��}�X�^
' ��۸���ID : mst/mst0133/main.asp
' �@      �\: ���y�[�W �A�E��}�X�^�̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H��R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSinroName        :�A�E�於�́i�ꕔ�j
'           txtPageCD       :�\���ϕ\���Ő��i�������g����󂯎������j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtRenrakusakiCD    :�I�����ꂽ�A����R�[�h
'           txtPageCD       :�\���ϕ\���Ő��i�������g�Ɉ����n�������j
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��A�E�E�i�w���\��
'           �����ցA�߂�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ��A�E�E�i�w��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/27 �≺�@�K��Y
' ��      �X: 2001/07/13 �J�e�@�ǖ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_sSinroCD      ':�i�H��R�[�h
    Public  m_sSingakuCd        ':�i�w�R�[�h
    Public  m_sSinroCD2     ':�i�H��R�[�h
    Public  m_sSingakuCd2       ':�i�w�R�[�h
    Public  m_sSyusyokuName     ':�A�E�於�́i�ꕔ�j
    Public  m_sPageCD       ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_skubun
    Public  m_Rs            'recordset
    Public  w_iDisp         ':�\�������̍ő�l���Ƃ�
    Public  m_sRenrakusakiCD
    Public  w_i
    w_i     = 1
    Public  w_iThisPgCnt
    Public  w_iSinrosakiCD
    Public  m_sSinroName
    Public  m_iNendo        ':�N�x
    Public  m_sMode

    '�y�[�W�֌W
    Public  m_cell
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
    w_sMsgTitle="�A�E�}�X�^"
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

        '�A�E�}�X�^���擾
        w_sWHERE = ""

        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_NENDO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & "    ,M01_KUBUN M01 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M01_NENDO = " & m_iNendo & " AND "
        w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo & " AND "
If m_sSinroCD <> 1 Then
        w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD (+) = "&C_SINRO&""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN = M01.M01_SYOBUNRUI_CD (+)"
Else
        w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD (+) = "&C_SINGAKU&""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN = M01.M01_SYOBUNRUI_CD (+)"
End If

        '���o�����̍쐬
        If m_sSinroCD<>"" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN =" & m_sSinroCD & " "
        End If
        If m_sSingakuCd<>"" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN =" & m_sSingakuCd & " "
        End If
        If m_sSinroName<>"" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINROMEI Like '%" & m_sSinroName & "%' "
        End If

        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M32.M32_SINRO_CD "

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
'Response.Write "m_iMax:" & m_iMax & "<br>"
        End If

'If m_sRenrakusakiCD = "" Then
'   Call NoDataPage()
'Exit Sub
'End If

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

    m_sMode      = Request("txtMode")
    m_sRenrakusakiCD = Request("txtRenrakusakiCD")  ':�A����R�[�h

    m_sSinroCD2 = Request("txtSinroCD2")        ':�i�H�R�[�h
    '�R���{���I����
    If m_sSinroCD2="@@@" Then
        m_sSinroCD2=""
    End If

'response.write m_sSinroCD2

    m_sSingakuCD2 = Request("txtSingakuCD2")    ':�i�w�R�[�h
    '�R���{���I����
    If m_sSingakuCD2="@@@" Then
        m_sSingakuCD2=""
    End If

    m_sSyusyokuName = Request("txtSyusyokuName")    ':�A�E�於�́i�ꕔ�j

    m_sSinroName = Request("txtSinroName")      ':�A�E�於�́i�ꕔ�j

    '// BLANK�̏ꍇ�͍s���ر
    If Request("txtMode") = "Delete" Then
        m_sPageCD = 1
    Else
        m_sPageCD = INT(Request("txtPageCD"))   ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If

    If m_sSinroCD = "1" Then            ':�w�b�_�[�̋敪���̕ύX
        m_skubun = "�i�w�敪"
    else
        m_skubun = "�i�H�敪"
    End If

    m_iNendo = Session("NENDO")         ':�N�x

    w_iDisp  = Request("txtDisp")           ':�y�[�W�ő�l

End Sub


Sub S_syousaiitiran()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_slink
Dim w_iCnt
Dim i

w_iThisPgCnt = 0
w_slink = "�@"

w_iCnt = 0


For i = 1 to w_iDisp


w_iSinrosakiCD = ""
w_sSQL = ""

If Request("deleteNO" & i) <> "" Then

w_iSinrosakiCD = Request("deleteNO" & i)

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWHERE            '// WHERE��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E�}�X�^"
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

        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_NENDO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & "    ,M01_KUBUN M01 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo & " AND "
If m_sSinroCD <> 1 Then
        w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD (+) = "&C_SINRO&""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN = M01.M01_SYOBUNRUI_CD (+)"
Else
        w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD (+) = "&C_SINGAKU&""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN = M01.M01_SYOBUNRUI_CD (+)"
End If

'response.write w_sSQL

        '���o�����̍쐬
        If m_sSinroCD <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN =" & m_sSinroCD & " "
        End If
        If m_sSingakuCd <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN =" & m_sSingakuCd & " "
        End If
        If m_sSinroName <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINROMEI Like '%" & m_sSinroName & "%' "
        End If
        If w_iSinrosakiCD <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_CD = '" & w_iSinrosakiCD & "' "

        End If

        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M32.M32_SINRO_CD "

'response.write w_sSQL


        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        Else
            '�y�[�W���̎擾
            m_iMax = gf_PageCount(m_Rs,m_iDsp)

'Response.Write "m_iMax:" & m_iMax & "<br>"
        End If

		w_iThisPgCnt = w_iThisPgCnt + 1

        If m_Rs.EOF Then
            '// �y�[�W��\��
            Call showPage_NoData()
        Else
            '// �y�[�W��\��
            Call S_syousai()
        End If
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If
    
    '// �I������
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End If



Next

    'LABEL_showPage_OPTION_END
End sub


Sub S_syousai()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_iCnt
Dim w_cell
w_iCnt  = 0

call gs_cellPtn(m_cell)
        %>

        <tr>
        <td align="center" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M01_SYOBUNRUIMEI")) %></td>
        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M32_SINROMEI")) %></td>
        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M32_DENWABANGO")) %></td>
        <td align="left" class=<%=m_cell%>><%=gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) %></td>
        <input type="hidden" name="deleteNO" value="<%=gf_HTMLTableSTR(m_Rs("M32_SINRO_CD")) %>">
        </tr>

        <%
w_i = w_i + 1
End sub

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

    On Error Resume Next
    Err.Clear
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
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  syosai_frm�ւ̃p�����[�^�̎󂯓n��
    //  [����]  p_sSyuseiCD
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Henko(p_sSyuseiCD){

        document.frm.action="syusei.asp";
        document.frm.target="";
        document.frm.txtRenrakusakiCD.value = p_sSyuseiCD;
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    }

    //************************************************************
    //  [�@�\]  �폜�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_delete(){

        if (!confirm("<%=C_SAKUJYO_KAKUNIN%>")) {
           return ;
        }

        document.frm.action="./delete.asp";
        document.frm.target="_self";
        document.frm.txtMode.value = "Delete";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_back(){

        document.frm.action="./default.asp";
        document.frm.target="fTopMain";
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
</head>
<body>

<center>


<%
If m_sMode = "Delete" Then
  m_sSubtitle = "��@��"
End If

call gs_title("�i�H����o�^",m_sSubtitle)
%>
<br>
�i�@�H�@��@��@��
<br><br>
    <table border="1" class=hyo width="75%">
<form name="frm" action="delete.asp" target="_self" method="post">

    <tr>
    <th class=header>�敪</th>
    <th class=header>�i�H��</th>
    <th class=header>TEL</th>
    <th class=header>URL</th>
    </tr>

    <% S_syousaiitiran() %>

    </table>
<br>
�ȏ�̓��e���폜���܂��B
<br><br>
<table border="0" width=50%>
<tr>
<td align=left>
<input type="button" class=button value="�@��@���@" Onclick="f_delete()">
<input type="hidden" name="txtMode" value="">
<input type="hidden" name="txtRenrakusakiCD" value="<%= m_sRenrakusakiCD %>">
<input type="hidden" name="txtSinroCD2" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD2" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtNendo" value="<%= m_iNendo %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
<input type="hidden" name="txtDisp" value="<%= w_iThisPgCnt %>">
</td>
</form>
<form action="default.asp" target="<%=C_MAIN_FRAME%>" method="post">
<td align=right>
<input type="submit" class=button value="�L�����Z��">
<input type="hidden" name="txtMode" value="search">
<input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
</td>
</form>
</tr>
</table>

</center>

</body>

</html>





<%
    '---------- HTML END   ----------
End Sub

Sub NoDataPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
    <html>
    <head>
    </head>

    <body>

    <center>
        �폜�̑ΏۂƂȂ�f�[�^���I������Ă��܂���B<br><br><br>
    <input type="button" class=button value="�߁@��" onclick="javascript:history.back()">
    </center>

    </body>

    </html>
<%
End Sub
%>