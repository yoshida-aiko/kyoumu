<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����Q�ƑI�����
' ��۸���ID : web/web0330/sousin_main.asp
' �@      �\: ���y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 2001/08/08 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp          '// �ꗗ�\���s��
    Public  m_iNendo        '�N�x
    Public  m_sKyokanCd     '��������
    Public  m_sJoukin       '��΋敪
    Public  m_sGakka        '�w�ȋ敪
    Public  m_sKkanKBN      '�����敪
    Public  m_sKkeiKBN      '���Ȍn��敪
    Public  m_rs
    Public  m_sPageCD       ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_iI
    Public  m_sKNm

    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bErrMsg           '�װү����
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
    w_sMsgTitle="�����Q�ƑI�����"
    w_sMsg=""
    w_sRetURL="../../../../default.asp"
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

        '// ���Ұ�SET
        Call s_SetParam()

        '//�f�[�^�̕\��
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
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

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Rs)
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

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sJoukin   = request("Joukin")
    m_sGakka    = request("Gakka")
    m_sKkanKBN  = request("KkanKBN")
    m_sKkeiKBN  = request("KkeiKBN")
    m_iI        = request("txtI")
    m_sKNm      = request("txtKNm")
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
        '//�i�荞�܂ꂽ�����ňꗗ�̕\��
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT "
        m_sSQL = m_sSQL & "     A.M04_KYOKAN_CD, "
        m_sSQL = m_sSQL & "     A.M04_KYOKANMEI_SEI, "
        m_sSQL = m_sSQL & "     A.M04_KYOKANMEI_MEI, "
'        m_sSQL = m_sSQL & "     A.M01_SYOBUNRUIMEI_R AS JOKIN_KBN, "
        m_sSQL = m_sSQL & "     B.M01_SYOBUNRUIMEI AS KYOKAKEIRETU_KBN, "
        m_sSQL = m_sSQL & "     C.M01_SYOBUNRUIMEI AS KYOKAN_KBN,"
        m_sSQL = m_sSQL & "     D.M02_GAKKA_KIGO"
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     M04_KYOKAN A ,M01_KUBUN B,M01_KUBUN C,M02_GAKKA D"
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     A.M04_NENDO = " & m_iNendo & " AND "
        m_sSQL = m_sSQL & "     A.M04_GAKKA_CD = D.M02_GAKKA_CD(+) AND "
'        m_sSQL = m_sSQL & "     M04_KYOKAN.M04_JOKIN_KBN = A.M01_SYOBUNRUI_CD AND "
'        m_sSQL = m_sSQL & "     M04_KYOKAN.M04_NENDO = A.M01_NENDO  AND "
        m_sSQL = m_sSQL & "     A.M04_KYOKAKEIRETU_KBN = B.M01_SYOBUNRUI_CD(+) AND "
        m_sSQL = m_sSQL & "     A.M04_KYOKAN_KBN = C.M01_SYOBUNRUI_CD(+) AND "
        m_sSQL = m_sSQL & "     A.M04_NENDO = B.M01_NENDO AND "
        m_sSQL = m_sSQL & "     A.M04_NENDO = C.M01_NENDO AND "
        m_sSQL = m_sSQL & "     A.M04_NENDO = D.M02_NENDO AND "
'        m_sSQL = m_sSQL & "     A.M01_DAIBUNRUI_CD=17 AND "
        m_sSQL = m_sSQL & "     B.M01_DAIBUNRUI_CD= " & C_KYOKA_KEIRETU &" AND "
        m_sSQL = m_sSQL & "     C.M01_DAIBUNRUI_CD= " & C_KYOKAN &" "

        If m_sGakka <> C_CBO_NULL Then
            m_sSQL = m_sSQL & " AND A.M04_GAKKA_CD = '" & m_sGakka & "' "
        End If
        If m_sKkanKBN <> C_CBO_NULL Then
            m_sSQL = m_sSQL & " AND A.M04_KYOKAN_KBN = " & Cint(m_sKkanKBN) & " "
        End If
        If m_sKkeiKBN <> C_CBO_NULL Then
            m_sSQL = m_sSQL & " AND A.M04_KYOKAKEIRETU_KBN = " & Cint(m_sKkeiKBN) & " "
        End If

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
Dim w_iCnt
Dim w_pageBar           '�y�[�WBAR�\���p

    w_iCnt  = 1
    w_bFlg  = True

    Call gs_pageBar(m_Rs,m_sPageCD,m_iDsp,w_pageBar)


%>
<%=w_pageBar %>
<table width="90%">
<tr><td>

<table border="1" width="100%" class="hyo">
<tr>
    <th width="30%" class="header">����</th>
    <th width="10%" class="header">�w��</th>
    <th width="15%" class="header">���Ȍn</th>
    <th width="45%" class="header">����</th>
</tr>
<%Do While (w_bFlg)
    Call gs_cellPtn(w_cell)
    %>
    <tr><form name="aaa" method="post">
        <td align="center" class="<%=w_cell%>"><%=m_rs("KYOKAN_KBN")%><BR></td>
        <td align="center" class="<%=w_cell%>"><%=m_rs("M02_GAKKA_KIGO")%><BR></td>
        <td align="center" class="<%=w_cell%>"><%=m_rs("KYOKAKEIRETU_KBN")%><BR></td>
        <td align="center" class="<%=w_cell%>">
        <input type="button" class="<%=w_cell%>" name="KNm" value='<%=m_rs("M04_KYOKANMEI_SEI")%>�@<%=m_rs("M04_KYOKANMEI_MEI")%>' onclick="iinSelect(this.form)">
        <input type="hidden" name="KCd" value='<%=gf_HTMLTableSTR(m_Rs("M04_KYOKAN_CD")) %>'>
        </td>
    </form></tr>
<%
    m_rs.MoveNext

        If m_rs.EOF Then
            w_bFlg = False
        ElseIf w_iCnt >= C_PAGE_LINE Then
            w_bFlg = False
        Else
            w_iCnt = w_iCnt + 1
        End If

Loop %>

</table>
</td></tr></table>
<%=w_pageBar %>

<!--<span class="CAUTION">�� �I��������ɂ͋��������N���b�N���Ă��������B</span>-->

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
<link rel="stylesheet" href="../../style.css" type="text/css">
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


%>
<HTML>
<BODY>

<link rel="stylesheet" href="../../style.css" type="text/css">
    <title>�����Q�ƑI�����</title>

    <!--#include file="../../jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
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
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �\�����e�\���p�E�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function iinSelect(p_sct) {

        //�}�����̃t�H�[�����擾
            w_sctNm = p_sct.KNm;
            w_sctNo = p_sct.KCd;

        //�}������
            parent.opener.document.frm.SKyokanNm<%=m_iI%>.value = w_sctNm.value;
            parent.opener.document.frm.SKyokanCd<%=m_iI%>.value = w_sctNo.value;

            document.frm.SKyokanNm.value = w_sctNm.value;
            document.frm.SKyokanCd.value = w_sctNo.value;

        return true;    
        //window.close()

    }

    //************************************************************
    //  [�@�\]  �N���A�{�^�����N���b�N�����ꍇ
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_Clear() {

        document.frm.SKyokanNm.value = "";
        document.frm.SKyokanCd.value = "";

        //�}�����������t�H�[�����擾
            w_NmStr = parent.opener.document.frm.SKyokanNm<%=m_iI%>;
            w_NoStr = parent.opener.document.frm.SKyokanCd<%=m_iI%>;

        //�}������

            w_NmStr.value = document.frm.SKyokanNm.value;
            w_NoStr.value = document.frm.SKyokanCd.value;
        return true;    
    }
    
    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" method="post">
    <INPUT TYPE="HIDDEN" NAME="txtNendo"    value="<%=m_iNendo%>">
    <INPUT TYPE="HIDDEN" NAME="txtKyokanCd" value="<%=m_sKyokanCd%>">
    <INPUT TYPE="HIDDEN" NAME="txtPageCD"   value="<%=m_sPageCD%>">
    <input type="hidden" name="txtI"        value="<%=m_iI%>">
    <input type="hidden" name="txtKNm"      value="<%=m_sKNm%>">

	<table width="50%" class="hyo">
	    <tr>
	        <td align="center" width="40%"><font color="white">���@���@��</font></td>
	        <td align="center" class="detail"><input type="text" class="noBorder" name="SKyokanNm" value="<%=m_sKNm%>" readonly>
	        <input type="hidden" name="SKyokanCd" value=""></td>
	    </tr>
	</table>
	<span class="CAUTION">�� �I��������ɂ͋��������N���b�N���Ă��������B</span>
	<table>
	    <tr>
	        <td colspan="4" align="center" nowrap>
	        <input type="button" value=" �N���A " class="button" onclick="javascript:f_Clear();">
	        <input type="button" value="����" class="button" onclick="javascript:parent.window.close();">
	        </td>
	    </tr>
	</table>
<%
        Call S_syousai()
%>
	<table>
	    <tr>
	        <td colspan="4" align="center" nowrap>
	        <input type="button" value=" �N���A " class="button" onclick="javascript:f_Clear();">
	        <input type="button" value="����" class="button" onclick="javascript:parent.window.close();">
	        </td>
	    </tr>
	</table>
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
