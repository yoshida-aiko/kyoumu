<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �󂫎��ԏ�񌟍�
' ��۸���ID : web/web0350/web0350_main.asp
' �@      �\: �������ʃy�[�W	 �󂫎��ԏ�񌟍����s��
'-------------------------------------------------------------------------
' ��      ��:
' ��      ��:
' ��      �n:
' ��      ��:
'           
'-------------------------------------------------------------------------
' ��      ��: 2001/08/17 ���i
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_bErrFlg           '�װ�׸�
    Public  m_Rs				'ں��޾�ĵ�޼ު��(�󂫎��Ԍ���)
    Public  m_iJMax				'�ő厞����
    Public  mRdiMode			'��������
    Public  mJigenSt			'�J�n����
    Public  mJigenEd			'�I������
    Public  m_SplitCell			'�N���X�w�肪�������z��
    Public  m_StrAkijikan		'html�������Ă�ϐ�

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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���Əo������"
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
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'// �����y�[�W��\��
		if gf_IsNull(request("txtDay")) then
			Call showPageDef()
			Exit do
		End if

		'// ��������
		w_iRet = f_SchAkijikan()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//�ő厞�������擾
        Call gf_GetJigenMax(m_iJMax)
		if m_iJMax = "" Then
		    m_bErrFlg = True
		    Exit Do
		end if

        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()
End Sub

Function f_SchAkijikan()
'******************************************************************
'�@�@�@�\�F����������
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear

    f_SchAkijikan = 1

    Do

		'// ���������擾
		wDay     = request("txtDay")		'<--�@���t
		mJigenSt   = request("txtJigenSt")		'<--�@�J�n����
		mJigenEd   = request("txtJigenEd")		'<--�@�I������
		mRdiMode = request("rdiMode")		'<--�@��������
		wGakkaCD = request("txtGakka")		'<--�@�w�ȃR�[�h

		'// �O��������擾
		Call gf_GetGakkiInfo(w_sGakki,w_sZenki_Start,w_sKouki_Start,w_sKouki_End)
		if (w_sZenki_Start > gf_YYYY_MM_DD(wDay,"/")) AND (gf_YYYY_MM_DD(wDay,"/") < w_sKouki_Start) then
			w_sGakki = C_GAKKI_KOUKI		'<--�@���
		ElseIf (w_sKouki_Start > gf_YYYY_MM_DD(wDay,"/")) AND (gf_YYYY_MM_DD(wDay,"/") < w_sKouki_End) then
			w_sGakki = C_GAKKI_ZENKI		'<--�@�O��
		End if

		
		m_sNomiSQL = ""
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " M04.M04_KYOKAN_CD NOT IN "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	(SELECT "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_KYOKAN "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 FROM  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20_JIKANWARI T20  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 WHERE  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_NENDO     = " & Session("NENDO") & " AND  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_GAKKI_KBN = " & w_sGakki & " AND  "
'	If mJigenSt = mJigenEd then '�J�n�ƏI�����Ⴄ�ꍇ�A���Ԏw��
'        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	    T20.T20_JIGEN     = " & mJigenSt   & " AND  "
'	Else
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	    T20.T20_JIGEN     >= " & mJigenSt   & " AND  "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	    T20.T20_JIGEN     <= " & mJigenEd   & " AND  "
'	End If
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_YOUBI_CD  = " & weekday(wDay)
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	group by "
        m_sNomiSQL = m_sNomiSQL & vbCrLf & " 	 	T20.T20_KYOKAN ) AND "


		'// �w�ȃR�[�h
		if Not wGakkaCD = C_CBO_NULL then
			m_sGakkaSQL = " M04.M04_GAKKA_CD = " & wGakkaCD & " AND "
		End if

        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAN_CD,"
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_GAKKA_CD, "
        m_sSQL = m_sSQL & vbCrLf & " 	M02.M02_GAKKARYAKSYO, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAKEIRETU_KBN, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKANMEI_SEI, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKANMEI_MEI "
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & " 	M02_GAKKA M02, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04_KYOKAN M04 "
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & " 	M02.M02_NENDO     = M04.M04_NENDO     AND "
        m_sSQL = m_sSQL & vbCrLf & " 	M02.M02_GAKKA_CD  = M04.M04_GAKKA_CD  AND "
        m_sSQL = m_sSQL & vbCrLf & 		m_sNomiSQL				'<--�����w���WHERE��
        m_sSQL = m_sSQL & vbCrLf & 		m_sGakkaSQL									'<--�w�ȃR�[�h��WHERE��
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_NENDO     = " & Session("NENDO") & " "
        m_sSQL = m_sSQL & vbCrLf & " GROUP BY "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAN_CD, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_GAKKA_CD, "
        m_sSQL = m_sSQL & vbCrLf & " 	M02.M02_GAKKARYAKSYO, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAKEIRETU_KBN, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKANMEI_SEI, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKANMEI_MEI "
        m_sSQL = m_sSQL & vbCrLf & " ORDER BY "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_GAKKA_CD, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAKEIRETU_KBN, "
        m_sSQL = m_sSQL & vbCrLf & " 	M04.M04_KYOKAN_CD "
'response.write m_sSQL
        w_iRet = gf_GetRecordset(m_Rs,m_sSQL)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

    	f_SchAkijikan = 0
	    Exit Do

    Loop

End Function

Sub showPageDef()
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
    <link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>

	<body>
	<BR>
	<div align="center">
		<br><br><br>
		<span class="msg">���ڂ�I��ŕ\���{�^���������Ă�������</span>
	</div>
	</body>
	</html>
<%
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
    <link rel=stylesheet href="../../common/style.css" type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
	</head>

	<body>
	<BR>
	<div align="center">

	<% if m_Rs.Eof then %>
        <br><br><br>
        <span class="msg">�Ώۃf�[�^�͑��݂��܂���B��������͂��Ȃ����Č������Ă��������B</span>
	<% Else %>

        <table class="hyo" border="1" width="400">
            <tr>
                <th nowrap class="header" width="64" align="center">���@�t</th>
                <td nowrap class="detail" width="100" align="center"><%=request("txtDay")%></td>
                <th nowrap class="header" width="64" align="center">���@��</th>
				<% if mJigenSt = mJigenEd then %>
	                <td nowrap class="detail" width="150" align="center"><%= mJigenSt %>����</td>
				<% Else %>
	                <td nowrap class="detail" width="150" align="center"><%= mJigenSt %>-<%= mJigenEd %>����</td>
				<% End if %>
            </tr>
        </table>
		<BR>

		<table><tr><td>
			<span class="msg"><font size="2">��<%=gf_GetRsCount(m_Rs)%>���̕����󂢂Ă��܂�</font></span>
		</td></tr></table>


		<table >
			<tr><td valign="top">
			<table class=hyo border="1" bgcolor="#FFFFFF">
				<tr>
					<th nowrap class="header">�w�@��</th>
					<th nowrap class="header">���Ȍn��</th>
					<th nowrap class="header">���@���@��</th>
				</tr>
				<% 	Call f_MainHyouji()	%>
			</table>
			</td></tr>
		</table>

	<% End if %>
	</div>
	</body>
	</html>
<%
End Sub


Sub s_jigenSuu()
'********************************************************************************
'*  [�@�\]  ���������쐬�i�w�b�_�[�����j
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	'// �����w�肪�Ȃ�������
	if gf_IsNull(mJigen) then
		i = 1
		Do Until i > Cint(m_iJMax)
			%><th nowrap width="80" class="header"><%= i %>����</th><%
			i = i + 1
		Loop
	Else
		'// �����w�肠��
		Select Case Cint(mRdiMode)
			Case 1
				'// �ȑO
				i = 1
				Do Until i > Cint(mJigen)
					%><th nowrap width="80" class="header"><%= i %>����</th><%
					i = i + 1
				Loop

			Case 2
				'// �̂�
				i = mJigen
				%><th nowrap width="80" class="header"><%= i %>����</th><%
				
			Case 3
				'// �ȍ~
				i = Cint(mJigen)
				Do Until i > m_iJMax
					%><th nowrap width="80" class="header"><%= i %>����</th><%
					i = i + 1
				Loop

		End Select
	End if

End Sub



Function f_MainHyouji()
'********************************************************************************
'*  [�@�\]  �󂫎��Ԃ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �S���A���Ƃ������Ă�����\�����Ȃ�
'********************************************************************************

	Dim w_iRsCnt,w_iCnt,i

	f_MainHyouji = 1
	
	w_iRsCnt = gf_GetRsCount(m_Rs)
	w_iCnt = INT(w_iRsCnt/2 + 0.9)
	
	Do Until m_Rs.Eof
			i = i + 1

		'// �e�[�u���̃N���X�w��
        Call gs_cellPtn(m_cell)

		'// ���Ȍn��擾
		Call gf_GetKubunName(C_KYOKA_KEIRETU,m_Rs("M04_KYOKAKEIRETU_KBN"),Session("NENDO"),wKyoukaKeiretu)

		
	%>
		<tr>
			<td nowrap class="<%=m_cell%>"><%=m_Rs("M02_GAKKARYAKSYO")%></td>
			<td nowrap class="<%=m_cell%>"><%=wKyoukaKeiretu%></td>
			<td nowrap class="<%=m_cell%>"><%=m_Rs("M04_KYOKANMEI_SEI")%>�@<%=m_Rs("M04_KYOKANMEI_MEI")%></td>
		</tr>
	<% If i =  w_iCnt And w_iRsCnt <> 1 Then 
				'//���ټ�Ă̸׽��������
				m_cell = ""
			%>
					</table>
				</td>
				<td valign="top">
					<table class="hyo" border="1" >
						<!--�w�b�_-->

				<tr>
				<th nowrap class="header">�w�@��</th>
				<th nowrap class="header">���Ȍn��</th>
				<th nowrap class="header">���@���@��</th>
			</tr>
	<%End If
	m_Rs.MoveNext
	Loop

	f_MainHyouji = 0
'
End Function

Function f_AkiJikan()
'********************************************************************************
'*  [�@�\]  �󂫎��Ԃ�\���i�j
'*  [����]  �Ȃ�
'*  [�ߒl]  True : False
'*  [����]  
'********************************************************************************

	f_AkiJikan = False
	w_AkinashiFlg = 0

	if gf_IsNull(mJigen) then
		i = 0
		Do Until i => Cint(m_iJMax)
			m_StrAkijikan = m_StrAkijikan & "<td nowrap class='" & m_SplitCell(i) & "'>&nbsp;</td>"
			i = i + 1
		Loop
	Else
		Select Case Cint(mRdiMode)
			Case 1
				'// �ȑO
				i = 0
				Do Until i >= Cint(mJigen)
					m_StrAkijikan = m_StrAkijikan & "<td nowrap class='" & m_SplitCell(i) & "'>&nbsp;</td>"
					if m_SplitCell(i) = "AKIJIKAN" then
						w_AkinashiFlg = w_AkinashiFlg + 1
					End if
					i = i + 1
				Loop

				if w_AkinashiFlg = Cint(mJigen) then
					Exit Function
				End if

			Case 2
				'// �̂�
				i = Cint(mJigen) - 1
				m_StrAkijikan = m_StrAkijikan & "<td nowrap class='" & m_SplitCell(i) & "'>&nbsp;</td>"
				
			Case 3
				'// �ȍ~
				i = Cint(mJigen) - 1
				w_FlgMax = 0
				Do Until i >= (Cint(m_iJMax))
					m_StrAkijikan = m_StrAkijikan & "<td nowrap class='" & m_SplitCell(i) & "'>&nbsp;</td>"
					if m_SplitCell(i) = "AKIJIKAN" then
						w_AkinashiFlg = w_AkinashiFlg + 1
					End if
					i = i + 1
					w_FlgMax = w_FlgMax + 1
				Loop

				if w_AkinashiFlg = w_FlgMax then
					Exit Function
				End if

		End Select
	End if
	m_StrAkijikan = m_StrAkijikan & "</tr>"

	f_AkiJikan = True

End Function
%>
