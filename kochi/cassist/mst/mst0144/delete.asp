<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A�E��}�X�^
' ��۸���ID : mst/mst0133/main.asp
' �@      �\: �A�E��}�X�^�̍폜���s��
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
' ��      ��: 2001/06/29 �≺�@�K��Y
' ��      �X: 
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
    Public  m_iDisp         ':�\�������̍ő�l���Ƃ�
    Public  m_sRenrakusakiCD
    Public  m_iNendo        ':�N�x


    '�y�[�W�֌W
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp                      '// �ꗗ�\���s��

'   call gs_viewForm(request.form)
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

        '// ���Ұ�SET
        Call s_SetParam()

        '// �폜�̎��s
        Call S_delete()

        '// �y�[�W��\��
        Call showPage()

End Sub


'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_sMode      = Request("txtMode")
    m_sRenrakusakiCD = Request("txtRenrakusakiCD")      ':�A����R�[�h

    m_sSinroCD2 = Request("txtSinroCD2")            ':�i�H�R�[�h
    '�R���{���I����
    If m_sSinroCD2="@@@" Then
        m_sSinroCD2=""
    End If

    m_sSingakuCD2 = Request("txtSingakuCD2")        ':�i�w�R�[�h
    '�R���{���I����
    If m_sSingakuCD2="@@@" Then
        m_sSingakuCD2=""
    End If

    m_sSyusyokuName = Request("txtSyusyokuName")        ':�A�E�於�́i�ꕔ�j


    '// BLANK�̏ꍇ�͍s���ر
    If Request("txtMode") = "Search" Then
        m_sPageCD = 1
    Else
        m_sPageCD = INT(Request("txtPageCD"))       ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If

    If m_sSinroCD = "1" Then                ':�w�b�_�[�̋敪���̕ύX
        m_skubun = "�i�w�敪"
    else
        m_skubun = "�i�H�敪"
    End If

    m_iDisp = Request("txtDisp")                ':�y�[�W�����ő�l

    m_iNendo = Request("txtNendo")              ':�N�x

End Sub


Sub S_delete()
'********************************************************************************
'*  [�@�\]  �폜���s
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_slink
Dim w_iCnt
Dim i
Dim w_iSinrosakiCD
Dim w_iSinrosakiCDs
Dim w_iSinrosakiCDe
Dim w_iSinrosakiCDStr

w_slink = "�@"

w_iCnt = 0

w_iSinrosakiCD = ""
w_sSQL = ""

w_iSinrosakiCD = Request("deleteNO")
w_iSinrosakiCDs = Split(w_iSinrosakiCD)
for each w_iSinrosakiCDe In w_iSinrosakiCDs
    w_iSinrosakiCDe = "'" & Replace(w_iSinrosakiCDe,",","") & "'"
    w_iSinrosakiCDStr = w_iSinrosakiCDStr & w_iSinrosakiCDe
next
    w_iSinrosakiCDStr = Replace(w_iSinrosakiCDStr,"''","','")
    'response.write w_iSinrosakiCDStr


    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWHERE            '// WHERE��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�i�H����o�^"
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
                Exit Do
                m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            End If

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()

        w_sSQL = w_sSQL & vbCrLf & " delete "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & " M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        '���o�����̍쐬
        w_sSQL = w_sSQL & vbCrLf & " M32.M32_SINRO_CD in (" & w_iSinrosakiCDStr & ")"

'response.write w_sSQL
        if gf_ExecuteSQL(w_sSQL) <> 0 then
            'ں��޾�Ă̎擾���s
            '//۰��ޯ�
            Call gs_RollbackTrans()
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If


    '���ׂĂ̏������������I��
    Exit do
    Loop
    

    '//�Я�
    Call gs_CommitTrans()

    '// �I������
    Call gs_CloseDatabase()
    
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If


    'LABEL_showPage_OPTION_END
End sub

'********************************************************************************
'*  [�@�\]  HTML��\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
%>





<html>
<link rel=stylesheet href=../common/style.css type=text/css>
    <head>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    function gonext() {
    alert('<%=C_SAKUJYO_OK_MSG%>');
            document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>
<body onLoad="gonext();">

<center>
<form action="default.asp" name="frm" target=<%=C_MAIN_FRAME%> method="post">
<img src="../../image/sp.gif" width="20" height="1">
<input type="hidden" name="txtMode" value="search">
<input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD2 %>">
<input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCD2 %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
</form>

</center>

</body>

</html>





<%
    '---------- HTML END   ----------
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
    <input type="button" value="�߁@��" onclick="javascript:history.back()">
    </center>

    </body>

    </html>
<%
End Sub
%>