<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�����
' ��۸���ID : gak/gak0350_11/main.asp
' �@      �\: ���y�[�W �w�Ѓf�[�^�̌������ʂ�\������
'-------------------------------------------------------------------------
' ��      ��:�Ȃ�
' ��      ��:�����N�x       ��      SESSION���i�ۗ��j
'			txtGakunen             :�w�N
'           txtGakkaCD             :�w��
'           txtClass               :�N���X
'           txtName                :����
'           txtGakusekiNo          :�w�Дԍ�
'           txtGakuseiNo           :�w���ԍ�
'           txtMode                :���샂�[�h
' ��      ��:
'           �������\��
'               �^�C�g���̂ݕ\��
'           �����ʕ\��
'               ��y�[�W�Őݒ肳�ꂽ���������ɂ��Ȃ��w���ʐ^��\������
'-------------------------------------------------------------------------
' ��      ��: 2006/04/28 �F��
' ��      �X: 2011/04/05 iwata �w���ʐ^�f�[�^���@Session����łȂ��A�f�[�^�x�[�X����擾����B
' ��      �X: 2017/05/17 ���{ �w���ʐ^�f�[�^Session�쐬�� global.asp�ōs���B
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen      	   ':�����N�x
    Public  m_sGakunen             ':�w�N
    Public  m_sGakkaCD             ':�w��
    Public  m_sClass               ':�N���X
    Public  m_sName                ':����
    Public  m_sGakusekiNo          ':�w�Дԍ�
    Public  m_sGakuseiNo           ':�w���ԍ�
    Public  m_TxtMode      	       ':���샂�[�h

    Public	m_Rs				   'recordset
    Public	m_iDsp				   '�ꗗ�\���s��

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
    w_sMsgTitle="�w����񌟍�����"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	'//���샂�[�h�̎擾
	m_TxtMode=request("txtMode")

    Do
		if m_TxtMode = "" then
           	Call showPage()
			Exit Do
		End if

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

'2011.04.05 ins
    '// �摜�f�[�^�擾�p oo4o �Z�b�V�����쐬
    'Set Session("OraDatabasePh") = OraSession.GetDatabaseFromPool(100)		'2017/05/17 Del Kiyomoto

        '// ���Ұ�SET
        Call s_SetParam()

        '�f�[�^���oSQL���쐬����
        Call s_MakeSQL(w_sSQL)

       '���R�[�h�Z�b�g�̎擾
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
		w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL,m_iDsp)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// �y�[�W��\��
        If m_Rs.EOF Then
            Call showPage_NoData()
        Else

	    '�w�����\��������
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
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

'2011.04.05 ins
		'** oo4o �ڑ��v�[���p��
	   'Session("OraDatabasePh").DestroyDatabasePool	'2017/05/17 Del Kiyomoto


End Sub

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************

	m_iSyoriNen = cint(Session("Nendo"))			'�����N�x

    m_sGakunen=request("txtGakunen")            	'�w�N
	'�R���{���I����
	If m_sGakunen="@@@" Then
		m_sGakunen=""
	End If

    m_sGakkaCD=request("txtGakka")            		'�w��
	'�R���{���I����
	If m_sGakkaCD="@@@" Then
		m_sGakkaCD=""
	End If

	'�w�N���I������Ă��Ȃ��ꍇ�̓N���X�͑I���ł��܂���
	if m_sGakunen="" then
		m_sClass=""
	else
    	m_sClass=request("txtClass")               	'�N���X
		'�R���{���I����
		If m_sClass="@@@" Then
			m_sClass=""
		End If
    end if

	m_sName = gf_Zen2Han(request("txtName"))        '����(���p�ɕϊ�)
	m_sGakusekiNo=request("txtGakusekiNo")          '�w�Дԍ�
	m_iDsp = cint(request("txtDisp"))				'�������X�g�̕\������

End Sub


'********************************************************************************
'*  [�@�\]  �w�Ѓf�[�^���oSQL������̍쐬
'*  [����]  p_sSql - SQL������
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub s_MakeSQL(p_sSql)

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " A.T13_GAKUSEKI_NO, "
    p_sSql = p_sSql & " A.T13_GAKUSEI_NO, "
    p_sSql = p_sSql & " B.T11_SIMEI "
    p_sSql = p_sSql & " FROM T13_GAKU_NEN A, T11_GAKUSEKI B, M02_GAKKA C, M05_CLASS D "
    p_sSql = p_sSql & " WHERE A.T13_NENDO = " & m_iSyoriNen & ""

    '���������̃Z�b�g
    if m_sGakunen <> "" then        '�w�N
        p_sSql = p_sSql & " AND A.T13_GAKUNEN = " & cint(m_sGakunen)
    end if
    if m_sGakkaCD <> "" then         '�w��
        p_sSql = p_sSql & " AND A.T13_GAKKA_CD = '" & m_sGakkaCD & "'"
    end if
    if m_sClass <> "" then           '�N���X
        p_sSql = p_sSql & " AND A.T13_CLASS = '" & m_sClass & "'"
    end if

    '��������
    p_sSql = p_sSql & " AND A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO(+) "
    p_sSql = p_sSql & " AND A.T13_NENDO = C.M02_NENDO"
    p_sSql = p_sSql & " AND A.T13_NENDO = D.M05_NENDO"
    p_sSql = p_sSql & " AND A.T13_GAKUNEN = D.M05_GAKUNEN "
    p_sSql = p_sSql & " AND A.T13_CLASS = D.M05_CLASSNO "
    p_sSql = p_sSql & " AND A.T13_GAKKA_CD = C.M02_GAKKA_CD "

    p_sSql = p_sSql & " ORDER BY A.T13_GAKUNEN,A.T13_GAKUSEKI_NO "

End Sub

'********************************************************************************
'*  [�@�\]  �ʐ^�����邩���� (BLOB�^�Ή��̈� T09_IMAGE ���@T09_GAKUSEI_NO�ɕύX�j
'*  [����]  �Ȃ�
'*  [�ߒl]  True: False
'*  [����]
'********************************************************************************
Function f_Photoimg(pGAKUSEI_NO)
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_Photoimg = False

	'// NULL�Ȃ甲����(False)
	if trim(pGAKUSEI_NO) = "" then Exit Function

	Do
	    w_sSQL = ""
	    w_sSQL = w_sSQL & " SELECT "
	    w_sSQL = w_sSQL & " T09_GAKUSEI_NO "
	    w_sSQL = w_sSQL & " FROM T09_GAKU_IMG "
	    w_sSQL = w_sSQL & " WHERE T09_GAKUSEI_NO = '" & cstr(pGAKUSEI_NO) & "'"

		w_iRet = gf_GetRecordset(w_ImgRs, w_sSQL)

		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			Exit Do
		End If

		'// EOF�Ȃ甲����(False)
		if w_ImgRs.Eof then	Exit Do

		'//����I��
		f_Photoimg = True
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �S�C���̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �S�C��
'*  [����]
'********************************************************************************
Function sGetTannin()
	Dim w_iRet
	Dim w_sSQL
	Dim w_oRs

	sGetTannin = ""

	w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "  M04_KYOKANMEI_SEI ,"
    w_sSQL = w_sSQL & "  M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "  M04_KYOKAN ,"
    w_sSQL = w_sSQL & "  M05_CLASS "
    w_sSQL = w_sSQL & " WHERE "

    '��������
    w_sSQL = w_sSQL & "  M04_KYOKAN.M04_NENDO = M05_CLASS.M05_NENDO "
    w_sSQL = w_sSQL & "  AND M04_KYOKAN.M04_KYOKAN_CD = M05_CLASS.M05_TANNIN "

    '���������̃Z�b�g
    w_sSQL = w_sSQL & "  AND M05_CLASS.M05_NENDO = " & m_iSyoriNen
    '�w�N
    If m_sGakunen <> "" Then
		w_sSQL = w_sSQL & "  AND M05_CLASS.M05_GAKUNEN = " & cint(m_sGakunen)
	Else
		w_sSQL = w_sSQL & "  AND M05_CLASS.M05_GAKUNEN = null"
	End If
	'�w��
    If m_sGakkaCD <> "" Then
		w_sSQL = w_sSQL & "  AND ( M05_CLASS.M05_GAKKA_CD = '" & m_sGakkaCD & "'"
	Else
		w_sSQL = w_sSQL & "  AND ( M05_CLASS.M05_GAKKA_CD = ''"
	End If
	'�N���X
    If m_sClass <> "" Then
		w_sSQL = w_sSQL & "  OR M05_CLASS.M05_CLASSNO = " & m_sClass & " )"
	Else
		w_sSQL = w_sSQL & "  OR M05_CLASS.M05_CLASSNO = null ) "
	End If

    w_iRet = gf_GetRecordset(w_oRs,w_sSQL)

	'�f�[�^�̊l���Ɏ��s�����甲����
    If w_iRet <> 0 Then Exit Function
    '�������擾�ł��Ȃ��ꍇ�͔�����
    If w_oRs.EOF Then Exit Function

	'���������֐��̖߂�l�ɃZ�b�g
	sGetTannin = "�S�C�F"
	sGetTannin = sGetTannin & w_oRs.fields(0).value		'������
	sGetTannin = sGetTannin & "�@"						'�����Ԃ̃X�y�[�X
	sGetTannin = sGetTannin & w_oRs.fields(1).value		'������

end Function

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
	<title>�w����񌟍�</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
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
Sub showPage()
%>

<html>

<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

    }
    //-->
    </SCRIPT>
    </head>

    <body>
	<% if m_TxtMode = "" then %>
		<center>
		<br><br><br>
		<span class="msg">���ڂ�I��ŕ\���{�^���������Ă�������</span>
		</center>
	<% Else %>
	    <div align="center">
	    <form action="" method="post" name="frm" target="">
		<% If sGetTannin <> "" Then %>
		<div align="left"><%= sGetTannin %></div>
		<% End If %>
		<table ID="Table1"><tr><td align="center" >
			<table border="0" width="100%" >
				<tr>
					<td align="center">
					<% if m_TxtMode = "" then %>
						<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
							<tr>
								<td width="60">&nbsp</td>
								<td valign="top"></td>
							</tr>
						</table>
					<% else %>
						<% dim w_cell %>

						<!--  �w���ʐ^�\���@-->
							<table border="0" cellpadding="0" cellspacing="2">
								<%
									w_iCnt = 1
									Do Until m_Rs.Eof or w_iCnt > m_iDsp
										response.write 	"<tr>"
										i_TdLine = 1
										'// ���ɂT���\�����C��
										Do Until m_Rs.Eof or i_TdLine > 5 or w_iCnt > m_iDsp
										%>
											<td align="center" class=search width="150" valign="top">
												<%
												'// ��ʐ^�����邩��Ɍ�������
												w_bRet = ""
												w_bRet = f_Photoimg(m_Rs("T13_GAKUSEI_NO"))

												if w_bRet = True then
													' 2011.04.05 upd DispBinary => DispBinaryRec �ɕύX
													' 2023.11.24 upd DispBinaryRec => DispBinary �ɕύX
													%><IMG SRC="DispBinary.asp?gakuNo=<%= m_Rs("T13_GAKUSEI_NO") %>" width="90" height="120" border="0"><%

												Else
													%><IMG SRC="images/Img0000000000.gif" width="90" height="120" border="0"><%
												End if
												%></a><br>
												<table border="0" cellpadding="0" cellspacing="2" width="100%">
													<tr><td bgcolor="#666699" nowrap><font color="#FFFFFF" size="1"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></font></td><td><font size="1"><%= m_Rs("T13_GAKUSEKI_NO") %></font></td></tr>
													<tr><td bgcolor="#666699" nowrap><font color="#FFFFFF" size="1">����    </font></td><td><font size="1"><%= trim(m_Rs("T11_SIMEI")) %></font></td></tr>
												</table>
											</td>
											<%
											i_TdLine = i_TdLine + 1
											w_iCnt = w_iCnt + 1
											m_Rs.MoveNext
										Loop %>
										</tr>
									<% Loop	%>
								</tr>
							</table>
					<% end if %>
				</td>
			</tr>
		</table>


		</td></tr></table>

		</div>
	    <input type="hidden" name="txtMode">

		<%' �������� %>
		<input type="hidden" name="txtGakunen"     value="<%=request("txtGakunen")%>">
		<input type="hidden" name="txtGakka"       value="<%=request("txtGakka")%>">
		<input type="hidden" name="txtClass"       value="<%=request("txtClass")%>">
		<input type="hidden" name="txtName"        value="<%=request("txtName")%>">
		<input type="hidden" name="txtGakusekiNo"  value="<%=request("txtGakusekiNo")%>">
		<input type="hidden" name="txtGakuseiNo"   value="<%=request("txtGakuseiNo")%>" ID="Hidden1">
		</form>
	<% End if %>
	</body>

    </html>

<%
    '---------- HTML END   ----------
End Sub

%>

