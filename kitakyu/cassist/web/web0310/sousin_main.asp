<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Ԋ������A��
' ��۸���ID : web/web0310/sousin_main.asp
' �@      �\: ���y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
'            ���[�h         ��      txtMode
'                                   �V�K = NEW
'                                   �X�V = UPDATE
'            ���e           ��      txtNaiyou
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/24 �O�c
' ��      �X: 2001/09/03 �ɓ����q �����ȊO�����p�ł���悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp                      '// �ꗗ�\���s��
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '��������
    Public m_stxtNo         '�����ԍ�
    Public m_stxtMode       '���[�h
    Public m_sNaiyou        '���e
    Public m_sKaisibi       '�J�n��
    Public m_sSyuryoubi     '������
    Public m_sJoukin        '��΋敪
    Public m_sGakka         '�w�ȋ敪
    Public m_sKkanKBN       '�����敪
    Public m_sKkeiKBN       '���Ȍn��敪
    Public m_rs
    Public m_Srs            '�X�V�̍ۂ̑��M��̉ߋ��̃f�[�^�擾�p���R�[�h
    Dim    m_rCnt           '//���R�[�h����

	Public m_sUserKbn		'//USER�敪
	Public m_sSimei			'//����

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
    m_iNendo    = request("txtNendo")
    m_sKaisibi  = request("txtKaisibi")
    m_sSyuryoubi= request("txtSyuryoubi")
    m_sKyokanCd = request("txtKyokanCd")
    m_sJoukin   = request("Joukin")
    m_stxtNo    = request("txtNo")
    m_iDsp = C_PAGE_LINE

    m_sGakka   = Trim(Replace(request("Gakka"),"@@@",""))
    m_sKkanKBN = Trim(Replace(request("KkanKBN"),"@@@",""))
    m_sKkeiKBN = Trim(Replace(request("KkeiKBN"),"@@@",""))
	m_sUserKbn = Trim(Replace(request("UserKbn"),"@@@",""))
	m_sSimei   = request("txtSimei")


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

'//�f�o�b�O
'Call s_DebugPrint


        '//�f�[�^�̕\��
        w_iRet = f_GetData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

        Select Case m_stxtMode
            Case "NEW"
                '// �y�[�W��\��
                Call showPage()
                Exit Do
            Case "UPD"
                '// �y�[�W��\��
                Call UPD_showPage()
                Exit Do
        End Select

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

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_stxtMode  = " & m_stxtMode   & "<br>"
    response.write "m_sKenmei   = " & m_sKenmei    & "<br>"
    response.write "m_sNaiyou   = " & m_sNaiyou    & "<br>"
    response.write "m_sKaisibi  = " & m_sKaisibi   & "<br>"
    response.write "m_sSyuryoubi= " & m_sSyuryoubi & "<br>"
    response.write "m_iNendo    = " & m_iNendo     & "<br>"
    response.write "m_sKyokanCd = " & m_sKyokanCd  & "<br>"
    response.write "m_sGakka    = " & m_sGakka     & "<br>"
    response.write "m_sKkanKBN  = " & m_sKkanKBN   & "<br>"
    response.write "m_sKkeiKBN  = " & m_sKkeiKBN   & "<br>"
    response.write "m_stxtNo    = " & m_stxtNo     & "<br>"
    response.write "m_sUserKbn  = " & m_sUserKbn   & "<br>"
    response.write "m_sSimei    = " & m_sSimei     & "<br>"

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
'
        '//�i�荞�܂ꂽ�����ňꗗ�̕\��
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
		m_sSQL = m_sSQL & vbCrLf & " WHERE "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_KYOKAN_CD = M04_KYOKAN.M04_KYOKAN_CD(+) "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO = M04_KYOKAN.M04_NENDO(+)"
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_iNendo

        If m_sGakka <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_GAKKA_CD= '" & m_sGakka & "' "
        End If

        If m_sKkanKBN <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_KYOKAN_KBN=" & Cint(m_sKkanKBN)
        End If

        If m_sKkeiKBN <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M04_KYOKAN.M04_KYOKAKEIRETU_KBN=" & Cint(m_sKkeiKBN)
        End If

        If m_sUserKbn <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_USER_KBN= " & m_sUserKbn
        End If

        If m_sSimei <> "" Then
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_USER_NAME LIKE '%" & m_sSimei & "%'"
        End If

		m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_GAKKA_CD,M04_KYOKAKEIRETU_KBN,M10_USER_NAME"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

		'//���R�[�h���擾
    	m_rCnt=gf_GetRsCount(m_rs)


        If m_stxtMode = "UPD" Then
            '//���M����Ă���l�̃f�[�^���擾
            m_sSQL = ""
            m_sSQL = m_sSQL & "SELECT "
            m_sSQL = m_sSQL & " T52_KYOKAN_CD "
            m_sSQL = m_sSQL & "FROM "
            m_sSQL = m_sSQL & " T52_JYUGYO_HENKO "
            m_sSQL = m_sSQL & "WHERE "
            m_sSQL = m_sSQL & " T52_NO = '" & cInt(m_stxtNo) & "'"

            Set m_Srs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(m_Srs, m_sSQL,m_iDsp)
            If w_iRet <> 0 Then
                'ں��޾�Ă̎擾���s
                m_bErrFlg = True
                Exit Do 
            End If
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

Sub S_NEW_syousai()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_Half
Dim j
j = 0
%>
<table>
    <tr>
        <td colspan=4 align=center>
            <input type="button" value=" �O �� �� �� " class=button onclick="javascript:f_before()">
            <input type="button" value="�S�ă`�F�b�N" class=button onclick="javascript:f_Allchk()">
            <input type="button" value=" �S�ăN���A " class=button onclick="javascript:f_AllClear()">
            <input type="button" value=" �I �� �� �� " class=button onclick="javascript:f_Skanryo()">
        </td>
    </tr>
</table>
<BR>
<div align="center"><span class=CAUTION>�� ���t�������l�Ƀ`�F�b�N�����Ģ�I��������{�^�����N���b�N���܂��B<br>
										�� �S���ɑ��肽���ꍇ�͢�S�ă`�F�b�N����A�S���̃`�F�b�N���͂��������ꍇ�͢�S�ăN���A����N���b�N���܂��B
</span></div>

    <%If NOT m_rCnt = "1" Then %>

<table border=0 width=100%>
<tr>

<td align="center" width=50% valign="top">

    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>�I��</th>
        <th width=22% class=header>����</th>
        <th width=4% class=header>�w��</th>
        <th width=19% class=header>���Ȍn</th>
        <th width=43% class=header>����</th>
    </tr>
    <%
        m_rs.MoveFirst
        w_Half = gf_Round(m_rCnt / 2 ,0)
        Do Until m_rs.EOF
            Call gs_cellPtn(w_cell)
            j = j + 1 
            If w_Half + 1 = j then
            w_cell = ""
            Call gs_cellPtn(w_cell)
    %>
    </table>
</td>
<td align="center" width=50% valign="top">
    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>�I��</th>
        <th width=22% class=header>����</th>
        <th width=4% class=header>�w��</th>
        <th width=19% class=header>���Ȍn</th>
        <th width=43% class=header>����</th>
    </tr>
    <% End If %>
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
			%>

	        <td class=<%=w_cell%> align="center"><input type=checkbox name=chk value="<%=m_rs("M10_USER_ID")%>"></td>
	        <td class=<%=w_cell%>><%=w_sKyokanKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=w_sGakkaKigo%><BR></td>
	        <td class=<%=w_cell%>><%=w_sKeiretuKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=m_rs("M10_USER_NAME")%><BR></td>
    </tr>
    <% m_rs.MoveNext
        Loop
    Else %>

<table border=0 width=50%>
<tr>

<td align="center" width=100% valign="top">

    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>�I��</th>
        <th width=22% class=header>����</th>
        <th width=4% class=header>�w��</th>
        <th width=19% class=header>���Ȍn</th>
        <th width=43% class=header>����</th>
    </tr>
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
	        <td class=<%=w_cell%> align="center"><input type=checkbox name=chk value="<%=m_rs("M10_USER_ID")%>"></td>
	        <td class=<%=w_cell%>><%=w_sKyokanKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=w_sGakkaKigo%><BR></td>
	        <td class=<%=w_cell%>><%=w_sKeiretuKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=m_rs("M10_USER_NAME")%><BR></td>
    </tr>
<% End If %>

</table></td></tr></table>
<table>
    <tr>
        <td colspan=4 align=center>
            <input type="button" value=" �O �� �� �� " class=button onclick="javascript:f_before()">
            <input type="button" value="�S�ă`�F�b�N" class=button onclick="javascript:f_Allchk()">
            <input type="button" value=" �S�ăN���A " class=button onclick="javascript:f_AllClear()">
            <input type="button" value=" �I �� �� �� " class=button onclick="javascript:f_Skanryo()">
        </td>
    </tr>
</table>

<%End sub

Sub S_UPD_syousai()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_Half
Dim j
j = 0
%>
<table>
    <tr>
        <td colspan=4 align=center>
            <input type="button" value=" �O �� �� �� " class=button onclick="javascript:f_before()">
            <input type="button" value="�S�ă`�F�b�N" class=button onclick="javascript:f_Allchk()">
            <input type="button" value=" �S�ăN���A " class=button onclick="javascript:f_AllClear()">
            <input type="button" value=" �I �� �� �� " class=button onclick="javascript:f_Skanryo()">
        </td>
    </tr>
</table>
<BR>
<div align="center"><span class=CAUTION>�� ���t�������l�Ƀ`�F�b�N�����Ģ�I��������{�^�����N���b�N���܂��B<br>
										�� �S���ɑ��肽���ꍇ�͢�S�ă`�F�b�N����A�S���̃`�F�b�N���͂��������ꍇ�͢�S�ăN���A����N���b�N���܂��B
</span></div>

    <%If NOT m_rCnt = "1" Then %>

<table border=0 width=100%>
<tr>

<td align="center" width=50% valign="top">

    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>�I��</th>
        <th width=22% class=header>����</th>
        <th width=4% class=header>�w��</th>
        <th width=19% class=header>���Ȍn</th>
        <th width=43% class=header>����</th>
    </tr>
    <%
        m_rs.MoveFirst
        w_Half = gf_Round(m_rCnt / 2 ,0)
        Do Until m_rs.EOF
            Call gs_cellPtn(w_cell)
            j = j + 1 
            If w_Half + 1 = j then
            w_cell = ""
            Call gs_cellPtn(w_cell)
    %>
    </table>
</td>
<td align="center" width=50% valign="top">
    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>�I��</th>
        <th width=22% class=header>����</th>
        <th width=4% class=header>�w��</th>
        <th width=19% class=header>���Ȍn</th>
        <th width=43% class=header>����</th>
    </tr>
    <%
            End If 
            m_Srs.MoveFirst
            w_schk=""
            Do Until m_Srs.EOF
                If m_rs("M10_USER_ID") = m_Srs("T52_KYOKAN_CD") Then
                    w_schk=" checked"
                    Exit Do
                End If
            m_Srs.MoveNext
            Loop
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

			%>
	        <td class=<%=w_cell%> align="center"><input type=checkbox name=chk value="<%=m_rs("M10_USER_ID")%>" <%=w_schk%>></td>
	        <td class=<%=w_cell%>><%=w_sKyokanKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=w_sGakkaKigo%><BR></td>
	        <td class=<%=w_cell%>><%=w_sKeiretuKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=m_rs("M10_USER_NAME")%><BR></td>
    </tr>
    <% m_rs.MoveNext
        Loop
    Else %>

<table border=0 width=50%>
<tr>

<td align="center" width=100% valign="top">

    <table width=100% border="1" class=hyo>
    <tr>
        <th width=4% class=header>�I��</th>
        <th width=22% class=header>����</th>
        <th width=4% class=header>�w��</th>
        <th width=19% class=header>���Ȍn</th>
        <th width=43% class=header>����</th>
    </tr>
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
	        <td class=<%=w_cell%> align="center"><input type=checkbox name=chk value="<%=m_rs("M10_USER_ID")%>"  <%=w_schk%>></td>
	        <td class=<%=w_cell%>><%=w_sKyokanKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=w_sGakkaKigo%><BR></td>
	        <td class=<%=w_cell%>><%=w_sKeiretuKbnName%><BR></td>
	        <td class=<%=w_cell%>><%=m_rs("M10_USER_NAME")%><BR></td>
    </tr>
<% End If %>

</table></td></tr></table>
<table>
    <tr>
        <td colspan=4 align=center>
            <input type="button" value=" �O �� �� �� " class=button onclick="javascript:f_before()">
            <input type="button" value="�S�ă`�F�b�N" class=button onclick="javascript:f_Allchk()">
            <input type="button" value=" �S�ăN���A " class=button onclick="javascript:f_AllClear()">�@
            <input type="button" value=" �I �� �� �� " class=button onclick="javascript:f_Skanryo()">�@
           </td>
    </tr>
</table>

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

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>���Ԋ������A��</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  �O��ʂփ{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_before(){

		// �S�ĕ\������Ă��Ȃ��ꍇ�͓���s��
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �I�������{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Skanryo(){

		// �S�ĕ\������Ă��Ȃ��ꍇ�͓���s��
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        if (f_chk()==1){
        alert( "�o�^�̑ΏۂƂȂ鑗�M�҂��I������Ă��܂���" );
        return;
        }

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.txtMode.value = "NEW2";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �S�ă`�F�b�N�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Allchk(){

		// �S�ĕ\������Ă��Ȃ��ꍇ�͓���s��
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        var i;
        i = 0;

        //1���̂Ƃ�
        if (document.frm.txtRcnt.value==1){
            document.frm.chk.checked == "True";
        }else{
        //����ȊO�̎�
        do { 
            if(document.frm.chk[i].checked == false){
                document.frm.chk[i].checked = true;
            }
        i++; }  while(i<document.frm.txtRcnt.value);
        }
        return;
    }

    //************************************************************
    //  [�@�\]  �S�ăN���A�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_AllClear(){

		// �S�ĕ\������Ă��Ȃ��ꍇ�͓���s��
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        var i;
        i = 0;

        //1���̂Ƃ�
        if (document.frm.txtRcnt.value==1){
            document.frm.chk.checked = false;
        }else{
        //����ȊO�̎�
        do { 
            if(document.frm.chk[i].checked == true){
                document.frm.chk[i].checked = false;
            }
        i++; }  while(i<document.frm.txtRcnt.value);
        }
        return;
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
            if (document.frm.chk.checked == false){
                return 1;
            }else{
                return 0;
                }
        }else{
        //����ȊO�̎�
        var checkFlg
            checkFlg=false

        do { 
            
            if(document.frm.chk[i].checked == true){
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
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		parent.frames[0].document.frm.BtnCtrl.value="OK"
    }

    //************************************************************
    //  [�@�\]  �{�^���̐���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
	function f_BtnCtrl(){

		if(parent.frames[0].document.frm.BtnCtrl.value!="OK"){
			return 1;
		}
		return 0;
	}

    //-->
    </SCRIPT>


<body LANGUAGE="javascript" onload="return window_onload()">
<center>
<FORM NAME="frm" method="post">
<%
    If m_rs.EOF Then
        Call showPage_NoData()
    Else
        Call S_NEW_syousai()
    End If
%>

</table>
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="<%=m_stxtMode%>">
    <INPUT TYPE=HIDDEN  NAME=txtKenmei      value="<%=m_sKenmei%>">
    <INPUT TYPE=HIDDEN  NAME=txtNaiyou      value="<%=m_sNaiyou%>">
    <INPUT TYPE=HIDDEN  NAME=txtKaisibi     value="<%=m_sKaisibi%>">
    <INPUT TYPE=HIDDEN  NAME=txtSyuryoubi   value="<%=m_sSyuryoubi%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_iNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtRcnt        value="<%=m_rCnt%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub

Sub UPD_showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_Half
Dim w_schk
Dim j
j = 0

%>
<HTML>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>���Ԋ������A��</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  �O��ʂփ{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_before(){
		// �S�ĕ\������Ă��Ȃ��ꍇ�͓���s��
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �I�������{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Skanryo(){

		// �S�ĕ\������Ă��Ȃ��ꍇ�͓���s��
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        if (f_chk()==1){
        alert( "�o�^�̑ΏۂƂȂ鑗�M�҂��I������Ă��܂���" );
        return;
        }

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "regist.asp";
        document.frm.txtMode.value = "UPD2";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �S�ă`�F�b�N�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Allchk(){

		// �S�ĕ\������Ă��Ȃ��ꍇ�͓���s��
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        var i;
        i = 0;

        //1���̂Ƃ�
        if (document.frm.txtRcnt.value==1){
            document.frm.chk.checked == "True";
        }else{
        //����ȊO�̎�
        do { 
            if(document.frm.chk[i].checked == false){
                document.frm.chk[i].checked = true;
            }
        i++; }  while(i<document.frm.txtRcnt.value);
        }
        return;
    }

    //************************************************************
    //  [�@�\]  �S�ăN���A�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_AllClear(){

		// �S�ĕ\������Ă��Ȃ��ꍇ�͓���s��
		iRet = f_BtnCtrl();
		if( iRet != 0 ){
			return;
		}

        var i;
        i = 0;

        //1���̂Ƃ�
        if (document.frm.txtRcnt.value==1){
            document.frm.chk.checked = false;
        }else{
        //����ȊO�̎�
        do { 
            if(document.frm.chk[i].checked == true){
                document.frm.chk[i].checked = false;
            }
        i++; }  while(i<document.frm.txtRcnt.value);
        }
        return;
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
            if (document.frm.chk.checked == false){
                return 1;
            }else{
                return 0;
                }
        }else{
        //����ȊO�̎�
        var checkFlg
            checkFlg=false

        do { 
            
            if(document.frm.chk[i].checked == true){
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
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		parent.frames[0].document.frm.BtnCtrl.value="OK"
    }

    //************************************************************
    //  [�@�\]  �{�^���̐���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
	function f_BtnCtrl(){

		if(parent.frames[0].document.frm.BtnCtrl.value!="OK"){
			return 1;
		}
		return 0;
	}

    //-->
    </SCRIPT>
<body LANGUAGE="javascript" onload="return window_onload()">
<center>

<FORM NAME="frm" method="post">

<%
    If m_rs.EOF Then
        Call showPage_NoData()
    Else
        Call S_UPD_syousai()
    End If
%>

</table>
    <INPUT TYPE=HIDDEN  NAME=txtNo          value="<%=m_stxtNo%>">
    <INPUT TYPE=HIDDEN  NAME=txtMode        value="<%=m_stxtMode%>">
    <INPUT TYPE=HIDDEN  NAME=txtKenmei      value="<%=m_sKenmei%>">
    <INPUT TYPE=HIDDEN  NAME=txtNaiyou      value="<%=m_sNaiyou%>">
    <INPUT TYPE=HIDDEN  NAME=txtKaisibi     value="<%=m_sKaisibi%>">
    <INPUT TYPE=HIDDEN  NAME=txtSyuryoubi   value="<%=m_sSyuryoubi%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo       value="<%=m_iNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    value="<%=m_sKyokanCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtRcnt        value="<%=m_rCnt%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>