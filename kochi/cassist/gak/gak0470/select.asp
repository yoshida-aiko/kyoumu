<%@Language=VBScript %>
<%
'******************************************************************
'�V�X�e����     �F���������V�X�e��
'���@���@��     �F�e��ψ��o�^
'�v���O����ID   �Fgak/gak0470/select.asp
'�@�@�@�@�\     �F�N���X���\���A�I�����s��
'------------------------------------------------------------------
'���@�@�@��     �F
'�ρ@�@�@��     �F
'���@�@�@�n     �F
'���@�@�@��     �F
'------------------------------------------------------------------
'��@�@�@��     �F2001.07.02    �O�c�@�q�j
'��      �X     : 2001/08/08 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'��      �X     : 2002/04/23 �{��@�@���ږ��u�w�Дԍ��v���Ǘ��}�X�^����擾����悤�ɕύX
'
'******************************************************************
'*******************�@ASP���ʃ��W���[���錾�@**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******�@�� �W �� �[ �� �� ���@********
    '�y�[�W�֌W
Public  m_iMax          ':�ő�y�[�W
Public  m_iDsp                      '// �ꗗ�\���s��
Public  m_bErrFlg       '//�G���[�t���O�iDB�ڑ��G���[���̏ꍇ�ɃG���[�y�[�W��\�����邽�߂̃t���O�j
Public  m_sDebugStr     '//�ȉ��f�o�b�N�p
Dim     m_iNendo        '//�����N�x
Dim     m_sKyokanCd     '//�����R�[�h
Dim     m_sGakunen      '//�w�N
Dim     m_sClass        '//�N���X��
Dim     m_sIinNm        '//�ψ�����
Dim     m_iI            '//default�̃��X�g�̈ʒu
Dim     m_rs            '//���R�[�h�Z�b�g
Dim     m_Irs           '//���R�[�h�Z�b�g�i�ψ��p�j
Dim     m_Grs           '//���R�[�h�Z�b�g�i�w�Дԍ��p�j
Dim     m_rCnt          '//���R�[�h����
'******�@���C�������@********

    'Ҳ�ٰ�ݎ��s
    Call Main()

'******�@�d�@�m�@�c�@********

Sub Main()
'******************************************************************
'�@�@�@�\�F�{ASP��Ҳ�ٰ��
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    '******���ʊ֐�******
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�e��ψ��o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"     
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakunen  = request("GAKUNEN")
    m_sClass    = request("CLASS")
    m_sIinNm    = request("IINNM")
    m_iI        = request("i")
    m_iDsp      = C_PAGE_LINE

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

        '//���X�g�ꗗ�̕\��
        w_iRet = f_getData()
        If w_iRet <> 0 Then
            '�G���[����
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

    '// �I������
    Call gs_CloseDatabase()
End Sub

Function f_getData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_getData = 1

    Do
        '//�w�N��N���X�̃f�[�^
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT"
        m_sSQL = m_sSQL & "     A.T13_GAKUSEI_NO,A.T13_GAKUSEKI_NO,B.T11_SIMEI "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T13_GAKU_NEN A,T11_GAKUSEKI B "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     A.T13_NENDO = '" & m_iNendo & "' "
        m_sSQL = m_sSQL & " AND A.T13_GAKUNEN = '" & m_sGakunen & "' "
        m_sSQL = m_sSQL & " AND A.T13_CLASS = '" & m_sClass & "' "
        m_sSQL = m_sSQL & " AND A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO(+) "
        m_sSQL = m_sSQL & " ORDER BY A.T13_GAKUSEKI_NO "

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_rCnt=gf_GetRsCount(m_rs)
    f_getData = 0

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_Half
Dim w_sKNM
Dim j
j = 0

w_sKNM = request("GName")

    On Error Resume Next
    Err.Clear

	'// ��׳�ް�ɂ�������݂̻��ނ�ς��� 
	w_btnWidth = ""
	if session("browser") = "NN" then
		w_btnWidth = "style='width:200'"
	End if

    '---------- HTML START ----------
    %>
<html>
<head>
<title>�e��ψ��o�^</title>
<link rel=stylesheet href="../../common/style.css" type="text/css">
<script language=javascript>
<!--
        //************************************************************
        //  [�@�\]  �\�����e�\���p�E�B���h�E�I�[�v��
        //  [����]
        //  [�ߒl]
        //  [����]
        //************************************************************
        function iinSelect(p_sct,p_No) {

            //�}�����������t�H�[�����擾
				w_NmStr = eval("opener.document.frm.gakuNm" + p_No);
				w_NoStr = eval("opener.document.frm.gakuNo" + p_No);

            //�}�����̃t�H�[�����擾
                w_sctNm = p_sct.gakuNm;
                w_sctNo = p_sct.gakuNo;

            //�}������
                w_NmStr.value = w_sctNm.value;
                w_NoStr.value = w_sctNo.value;

                document.frm.SearchNm.value = w_sctNm.value;
                document.frm.SearchNo.value = w_sctNo.value;

            return true;    
            //window.close()

        }

        //************************************************************
        //  [�@�\]  ����,�w���R�[�h�̐\�����e�\��
        //  [����]
        //  [�ߒl]
        //  [����]
        //************************************************************
        //function f_SearchSelect(p_sct) {
        //  //�}�����̃t�H�[�����擾
        //      w_sctNm = p_sct.gakuNm;
        //      w_sctNo = p_sct.gakuNo;
        //
        //  //�}������
        //      document.frm.SearchNm.value = w_sctNm.value;
        //      document.frm.SearchNo.value = w_sctNo.value;
        //  return true;    
        //}

        //************************************************************
        //  [�@�\]  �N���A�{�^�����N���b�N�����ꍇ
        //  [����]
        //  [�ߒl]
        //  [����]
        //************************************************************
        function f_Clear(p_No) {

            document.frm.SearchNm.value = "";
            document.frm.SearchNo.value = "";

            //�}�����������t�H�[�����擾
				w_NmStr = eval("opener.document.frm.gakuNm" + p_No);
				w_NoStr = eval("opener.document.frm.gakuNo" + p_No);
                
            //�}������
                w_NmStr.value = "";
                w_NoStr.value = "";
            return true;    
        }

    //-->
    </script>


</head>

<body onload="focus();">
<form name="frm">
<center>
<%call gs_title("�e��ψ��o�^","�Q�@��")%>

<br>

<table border="0" cellpadding="1" cellspacing="1" width="500">
	<tr>
		<td align="center" colspan="2">

		    <table class="hyo">
			    <tr>
				    <td align="center" width="150"><font color="white"><%=m_sIinNm%></font></td>
				    <td align="center" width="250" class="detail"><input type="text" class="CELL2" name="SearchNm" value="<%=w_sKNM%>" readonly><input type="hidden" name="SearchNo" value="<%=gf_fmtZero(m_rs("T13_GAKUSEI_NO"),10)%>"></td>
			    </tr>
		    </table>
			<br>
			<form>
			    <input type="button" value=" �N���A " class="button" onclick="javascript:f_Clear('<%=m_iI%>');">�@
			    <input type="button" value="����" class="button" onclick="javascript:window.close();">
			</form>
			<span class="msg">���I��������ɂ͖��O���N���b�N���A����{�^�����N���b�N���Ă�������</span><br>

		</td>
	</tr>
	<tr>
		<td align="center" width="250" valign="top">

		    <table border="1" class="hyo">
		    <tr>
				<!-- 2002/04/23 miyai -->
				<th width="80" class="header" nowrap><%=gf_GetGakuNomei(Session("NENDO"),C_K_KOJIN_1NEN)%></th>
		        <th width="170" class="header" nowrap>����</th>
		    </tr>
		    <%
		        m_rs.MoveFirst
		        w_Half = gf_Round(m_rCnt / 2,0)
		        Do Until m_rs.EOF
		            Call gs_cellPtn(w_cell)
		            j = j + 1 
		            If w_Half + 1 = j then
		            w_cell = ""
		            Call gs_cellPtn(w_cell)
		    %>
		    </table>

		</td>
		<td align="center" width="250" valign="top">

		    <table border="1" class="hyo">
			    <tr>
					<!-- 2002/04/23 miyai -->
					<th width="80" class="header" nowrap><%=gf_GetGakuNomei(Session("NENDO"),C_K_KOJIN_1NEN)%></th>
				    <th width="170" class="header" nowrap>����</th>
			    </tr>
		    <% End If %>
			    <tr><form>
				    <td class="<%=w_cell%>" align="center"><%=m_rs("T13_GAKUSEKI_NO")%><input type="hidden" name="gakuNo" value="<%=m_rs("T13_GAKUSEI_NO")%>"></td>
				    <td class="<%=w_cell%>" align="left"><input type="button" class="<%=w_cell%>" <%=w_btnWidth%> name="gakuNm" value="<%=gf_SetNull2String(m_rs("T11_SIMEI"))%>" onclick="iinSelect(this.form,'<%=m_iI%>')"></td>
			    </form>
			    </tr>
		    <%
		        m_rs.MoveNext
		        Loop%>

		    </table>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2">

			<br>
			<form>
			    <input type="button" value=" �N���A " class="button" onclick="javascript:f_Clear('<%=m_iI%>');">�@
			    <input type="button" value="����" class="button" onclick="javascript:window.close();">
			</form>

		</td>
	</tr>
</table>

<INPUT TYPE="HIDDEN" NAME="GAKUNEN" VALUE="<%=request("m_sGakunen") %>">
<INPUT TYPE="HIDDEN" NAME="CLASS"   VALUE="<%=request("m_sClass") %>">
<INPUT TYPE="HIDDEN" NAME="IINNM"   VALUE="<%=request("m_sIinNm") %>">
<INPUT TYPE="HIDDEN" NAME="i"       VALUE="<%=request("m_iI") %>">

</center>
</form>
</center>
</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>