<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Ԋ������A��
' ��۸���ID : web/web0330/view.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
'            ���[�h         ��      txtMode
'                                   �V�K = NEW
'                                   �X�V = UPDATE
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iMax           ':�ő�y�[�W
    Public m_iDsp                       '// �ꗗ�\���s��
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '��������
    Public m_stxtMode       '���[�h
    Public m_sNaiyou        '���e
    Public m_sKaisibi       '�J�n��
    Public m_sSyuryoubi     '������
    Public m_sJoukin        '��΋敪
    Public m_sGakka         '�w�ȋ敪
    Public m_sKkanKBN       '�����敪
    Public m_sKkeiKBN       '���Ȍn��敪
    Public m_stxtNo         '�����ԍ�
    Public m_rs
    Public m_sListCd
    Dim    m_rCnt           '//���R�[�h����

    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
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
    w_sMsgTitle="�A�������o�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_stxtMode = request("txtMode")

    m_sNaiyou   = request("txtNaiyou")
    m_sKaisibi  = request("txtKaisibi")
    m_sSyuryoubi= request("txtSyuryoubi")
    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_stxtNo    = request("txtNo")
    m_sListCd = request("chk")
    m_iDsp = C_PAGE_LINE

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
        
        '//�f�[�^�̎擾�A�\��
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

        Call showPage()
        Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Rs)
    '// �I������
    Call gs_CloseDatabase()
End Sub

Function f_GetData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_sSQL
Dim w_Srs           '�ڍחp�̃��R�[�h�Z�b�g

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
        '//�ϐ��̒l���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT DISTINCT"
        w_sSQL = w_sSQL & " T52_NAIYO,T52_KAISI,T52_SYURYO "
        w_sSQL = w_sSQL & "FROM "
        w_sSQL = w_sSQL & " T52_JYUGYO_HENKO "
        w_sSQL = w_sSQL & "WHERE "
        w_sSQL = w_sSQL & " T52_NO = '" & cInt(m_stxtNo) & "'"

        Set w_Srs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_Srs, w_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

	    '//�擾�����l��ϐ��ɑ��
	    m_sNaiyou   = w_Srs("T52_NAIYO")
	    m_sKaisibi  = w_Srs("T52_KAISI")
	    m_sSyuryoubi= w_Srs("T52_SYURYO")

        '//���M����Ă���l�̃f�[�^���擾
		m_sSQL = ""
		m_sSQL = m_sSQL & vbCrLf & " SELECT "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_USER_ID "
		m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_KBN "
		m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_NAME "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_CD "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_GAKKA_CD "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAKEIRETU_KBN "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_KBN"
		m_sSQL = m_sSQL & vbCrLf & " FROM "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN "
		m_sSQL = m_sSQL & vbCrLf & "  ,T52_JYUGYO_HENKO "
		m_sSQL = m_sSQL & vbCrLf & " WHERE "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_KYOKAN_CD = M04_KYOKAN.M04_KYOKAN_CD(+) "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO = M04_KYOKAN.M04_NENDO(+)"
		m_sSQL = m_sSQL & vbCrLf & "  AND T52_JYUGYO_HENKO.T52_NO = " & cInt(m_stxtNo)
		m_sSQL = m_sSQL & vbCrLf & "  AND T52_JYUGYO_HENKO.T52_KYOKAN_CD = M10_USER.M10_USER_ID(+) "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_iNendo
		m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_GAKKA_CD,M04_KYOKAKEIRETU_KBN,M10_USER_NAME"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

    f_GetData = 0

    Exit Do

    Loop

End Function

'********************************************************************************
'*  [�@�\]  �w�ȋL�����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  gf_GetUserNm:
'*  [����]  
'********************************************************************************
Function f_GetGakkaKigoName(p_sGakkaCd)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetGakkaKigoName = ""
	w_sName = ""

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKA_KIGO"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_GAKKA_CD='" & p_sGakkaCd & "'"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M02_GAKKA_KIGO")
        End If

        Exit Do
    Loop

	'//�߂�l���
    f_GetGakkaKigoName = w_sName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_sClass

%>
<HTML>
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>���Ԋ������A��</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  ����{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Close(){

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "default.asp";
        document.frm.submit();

    }
    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" action="post">

<br>
<% 
call gs_title("���Ԋ������A��","�Ɓ@��")
%>

<br>
<font>�o�@�^�@���@�e</font>
<br>
</TD>
</TR>
</TABLE>
<BR>
<div align="center"><span class=CAUTION>�� ���t��̊m�F���s�Ȃ��܂��B<br>
</span></div>

<br>

<table width="500" border=1 CLASS="hyo">
    <TR>
        <TH CLASS="header" width="60">���e</TD>
        <TD CLASS="detail"><%=m_sNaiyou%></TD>
    </TR>
    <TR>
        <TH CLASS="header">����</TD>
        <TD CLASS="detail"><%=m_sKaisibi%>�@�`�@<%=m_sSyuryoubi%></TD>
    </TR>
    <tr>
	    <td colspan=5 align="right" bgcolor=#9999BD><input class=button type="submit" value="����" class=button onclick="javascript:f_Close()"></td>
    </tr>
    <TR>
        <TH CLASS="header" valign="top">���t��</TD>
        <TD CLASS="detail" >
        <table border=1 class=hyo width=100% height=100%>
    <%
        m_rs.MoveFirst
        Do Until m_rs.EOF
    %>
            <tr>

			<%
			'========================================================
			'//�敪���̓��擾
			w_sKyokanKbnName = ""
			w_sKeiretuKbnName = ""
			w_sGakkaKigo = ""

			'//����CD���Z�b�g
			w_sKyokanCd = m_rs("M04_KYOKAN_CD")

			'//�����̎�(����CD����̏ꍇ)
			If LenB(w_sKyokanCd) <> 0 Then
				'//�����敪���̂��擾
				Call gf_GetKubunName(C_KYOKAN,m_rs("M04_KYOKAN_KBN"),m_iNendo,w_sKyokanKbnName)

				'//���Ȍn��敪���̂��擾
				Call gf_GetKubunName(C_KYOKA_KEIRETU,m_rs("M04_KYOKAKEIRETU_KBN"),m_iNendo,w_sKeiretuKbnName)

				w_sGakkaKigo = f_GetGakkaKigoName(m_rs("M04_GAKKA_CD"))
			Else
				'//�����ȊO�̏ꍇUSER�敪���̂�\��
				Call gf_GetKubunName(C_USER,m_rs("M10_USER_KBN"),m_iNendo,w_sKyokanKbnName)
				w_sKeiretuKbnName = "�\"
				w_sGakkaKigo = "�\"
			End If

			'========================================================

            Call gs_cellPtn(w_cell)
			%>

                <td class="CELL2" width=21%><%=w_sKyokanKbnName%><br></td>
                <td class="CELL2" width=21%><%=w_sKeiretuKbnName%><br></td>
                <td class="CELL2" width=6%><%=w_sGakkaKigo%><br></td>
                <td class="CELL2" width=40%><%=m_rs("M10_USER_NAME")%><br></td>

            </tr>
    <%  m_rs.MoveNext
        Loop%>
            </table>
        </TD>
    </TR>
    </TABLE>

    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_iNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>