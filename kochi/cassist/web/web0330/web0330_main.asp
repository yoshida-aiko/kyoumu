<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A���f����
' ��۸���ID : web/web0330/web0330_main.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 2001/08/27 �ɓ����q �R�����g�����X�g�̏㕔�ɕ\������悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp          '// �ꗗ�\���s��
    Public  m_sPageCD       ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_rs
    Dim     m_sNendo
    Dim     m_sKyokanCd
    Dim     m_rCnt          '//���R�[�h����

    '�G���[�n
    Public  m_bErrFlg       '�װ�׸�
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
    w_sMsgTitle="�A���f����"
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

        '// ���Ұ�SET
        Call s_SetParam()

        '//���X�g�̈ꗗ�f�[�^�̏ڍ׎擾
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
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

    '//ں��޾��CLOSE
    Call gf_closeObject(m_rs)
    '// �I������
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_sNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp = C_PAGE_LINE

    If Request("txtPageCD") <> "" Then
        m_sPageCD = INT(Request("txtPageCD"))   ':�\���ϕ\���Ő��i�������g����󂯎������j
    Else
        m_sPageCD = 1   ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
    If m_sPageCD = 0 Then m_sPageCD = 1

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_sNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iDsp = C_PAGE_LINE

    If Request("txtPageCD") <> "" Then
        m_sPageCD = INT(Request("txtPageCD"))   ':�\���ϕ\���Ő��i�������g����󂯎������j
    Else
        m_sPageCD = 1   ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
    If m_sPageCD = 0 Then m_sPageCD = 1

End Sub

Function f_GetData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
        '//���X�g�̕\��
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT DISTINCT"
        m_sSQL = m_sSQL & "     T46_NO,T46_KENMEI,T46_KAISI,T46_SYURYO "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T46_RENRAK "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T46_INS_USER = '" & Session("LOGIN_ID") & "' "

		'���}���u�B�������߂�����\�����Ȃ��B	2001/12/17
		'�{���͕\�����Ȃ���������Ȃ��Ċ������߂��Ă���f�[�^�͍폜����B
        m_sSQL = m_sSQL & " AND T46_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "'"
        m_sSQL = m_sSQL & " AND T46_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_rCnt=gf_GetRsCount(m_rs)

    f_GetData = 0

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub S_syousai()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    Dim w_pageBar           '�y�[�WBAR�\���p

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True
%>
<div align="center"><span class=CAUTION>�� �V�K�o�^�̏ꍇ�͢�V�K�o�^�͂����磂��N���b�N���Ă��������B<br>
										�� �C���̏ꍇ�͢>>����N���b�N���Ă��������B<br>
										�� �������N���b�N����Ƒ��M���e���m�F�ł��܂��B
</span></div>
<br>
<%
    '�y�[�WBAR�\��
    Call gs_pageBar(m_rs,m_sPageCD,m_iDsp,w_pageBar)

%>
<%=w_pageBar %>
</td></tr>
<tr><td>
<table width="100%" border="1" CLASS="hyo">
    <TR>
<!--         <TH CLASS="header" width="40" nowrap>����<br>�ԍ�</TH> -->
        <TH CLASS="header" width="55%" nowrap>���@��</TH>
        <TH CLASS="header" nowrap>���@��</TH>
        <TH CLASS="header" width="16" nowrap>�C��</TH>
        <!--TH CLASS="header" width="16" nowrap>�폜</TH-->
    </TR>

<%	Do While (w_bFlg)
    call gs_cellPtn(w_cell)
	call gs_ColorPtnNN(w_color)
%>
    <TR>
<!--        <TD CLASS="<%=w_cell%>" ALIGN="right"><%=m_rs("T46_NO")%></TD> -->
        <TD CLASS="<%=w_cell%>"><a href="javascript:f_Kakunin(<%=m_rs("T46_NO")%>)"><%=m_rs("T46_KENMEI")%></a></TD>
        <TD CLASS="<%=w_cell%>" ALIGN="center"><%=m_rs("T46_KAISI")%>�`<%=m_rs("T46_SYURYO")%></TD>
        <TD CLASS="<%=w_cell%>" ALIGN="center"><input type="button" value=">>" class=button onclick="javascript:f_Syusei(<%=m_rs("T46_NO")%>)"></TD>
        <!--TD CLASS="<%=w_cell%>" ALIGN="center"><input type="checkbox" name=Delchk value="<%=m_rs("T46_NO")%>"></TD-->
    </TR>
<% m_rs.MoveNext

		If m_rs.EOF Then
		    w_bFlg = False
		ElseIf w_iCnt >= C_PAGE_LINE Then
		    w_bFlg = False
		Else
		    w_iCnt = w_iCnt + 1
		End If

    Loop %>
    <tr>
    <!--td colspan=5 align="right" bgcolor=#9999BD><input class=button type=button value="�~�폜" onclick="javascript:f_delete()"></td-->
    </tr>
 </table>
 </td></tr>
 <tr><td>
<%=w_pageBar %>
<BR>
<!--
<div align="center"><span class=CAUTION>�� �V�K�o�^�̏ꍇ�͢�V�K�o�^�͂����磂��N���b�N���Ă��������B<br>
										�� �C���̏ꍇ�͢>>����N���b�N���Ă��������B<br>
										�� �������N���b�N����Ƒ��M���e���m�F�ł��܂��B
</span></div>
-->
<div align="center"><span class=CAUTION>�� ���b�Z�[�W�́A�\�����Ԃ��߂���Ǝ����I�ɍ폜����܂��B<br>
</span></div>
<%End sub

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
        <span class="msg">�A���f�[�^�͑��݂��܂���B<br>��V�K�o�^�͂����磂��N���b�N���Ă��������B</span>
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
    Dim i
    i=1

%>
<HTML>
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>�A���f����</title>

    <!--#include file="../../Common/jsCommon.htm"-->
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
    //  [�@�\]  �폜�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_delete(){

        if (f_chk()==1){
        alert( "�폜�̑ΏۂƂȂ錏�����I������Ă��܂���" );
        return;
        }

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "web0330_DEL.asp";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  ���X�g�ꗗ�̃`�F�b�N�{�b�N�X�̊m�F
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_chk(){
        var i;
        i = 0;

        //0���̂Ƃ�
        if (document.frm.txtRcnt.value<=0){
            return 1;
            }

        //1���̂Ƃ�
        if (document.frm.txtRcnt.value==1){
            if (document.frm.Delchk.checked == false){
                return 1;
            }else{
                return 0;
                }
        }else{
        //����ȊO�̎�
        var checkFlg
            checkFlg=false

        do { 
            
            if(document.frm.Delchk[i].checked == true){
                checkFlg=true
                break;
             }

        i++; }  while(i<document.frm.txtRcnt.value);
            if (checkFlg == false){
                return 1;
                }
        }
        return 0;
    }

    //************************************************************
    //  [�@�\]  �����{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Kakunin(p_Int){

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "view.asp";
        document.frm.txtNo.value = p_Int;
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �C���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Syusei(p_Int){

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.txtNo.value = p_Int;
        document.frm.txtMode.value = "UPD";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" ACTION="post">
<table width="90%" border="0"><tr><td>
<%
    If m_rs.EOF Then
        Call showPage_NoData()
    Else
        Call S_syousai()
    End If
%>
    <INPUT TYPE=HIDDEN  NAME=txtNo          value="">
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_sNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtPageCD      value="<%= m_sPageCD %>">
    <INPUT TYPE=HIDDEN  NAME=txtRcnt        value="<%=m_rCnt%>">
</td></tr></table>

</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>