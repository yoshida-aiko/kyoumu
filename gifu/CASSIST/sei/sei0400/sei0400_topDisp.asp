<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����������o�^
' ��۸���ID : gak/sei0400/sei0400_topDisp.asp
' �@      �\: 
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'               �㕔��ʕ\���̂�
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/18 �O�c �q�j
' ��      �X: 2001/08/09 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
' ��      �X�F2001/08/30 �ɓ� ���q     ����������2�d�ɕ\�����Ȃ��悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n

    '�s�����I��p��Where����
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '�����R�[�h
    Public m_sGakuNo        '�����R���{�{�b�N�X�ɓ���l
    Public m_sBeforGakuNo   '�����R���{�{�b�N�X�ɓ���l�̈�l�O
    Public m_sAfterGakuNo   '�����R���{�{�b�N�X�ɓ���l�̈�l��
    Public m_sSsyoken       '��������
    Public m_sBikou         '�l���l
    Public m_sSinro         '�i�H��
    Public m_sSotudai       '�����ۑ�
    Public m_sSkyokan1      '����1
    Public m_sSkyokan2      '����2
    Public m_sSkyokan3      '����3
    Public m_sGakunen       '�w�N
    Public m_sClass         '�N���X
    Public m_sClassNm       '�N���X��
    Public m_sGakusei()     '�w���̔z��
    Public m_sGakka     '�w���̏����w��
    Public m_sShiken

    Public  m_GRs
    Public  m_Rs
    Public  m_iMax          '�ő�y�[�W
    Public  m_iDsp          '�ꗗ�\���s��

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
    w_sMsgTitle="�����������o�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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

        '// ���Ұ�SET
        Call s_SetParam()

	  '//���������擾
            If f_GetSiken(m_sShiken) <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If

		Call f_Gakusei()

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

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakuNo   = request("txtGakuNo")
    m_iDsp      = C_PAGE_LINE
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm    = request("txtClassNm")
	m_sShiken    = request("txtSikenKBN")

	'//�O��OR���փ{�^���������ꂽ��
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

Function f_Gakusei()
'********************************************************************************
'*  [�@�\]  �w���f�[�^���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim i
i = 1


    w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

	'//�w���̏����W
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
'    w_sSQL = w_sSQL & " AND T11_NYUNENDO = " & w_iNyuNendo & " "

    Set m_GRs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(m_GRs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
    End If


    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     A.T11_GAKUSEI_NO "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI A,T13_GAKU_NEN B "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     B.T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND B.T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND B.T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & " ORDER BY B.T13_GAKUSEKI_NO "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
    End If
	w_rCnt=cint(gf_GetRsCount(w_Rs))

	'//�z��̍쐬

		w_Rs.MoveFirst

       Do Until w_Rs.EOF

            ReDim Preserve m_sGakusei(i)
            m_sGakusei(i) = w_Rs("T11_GAKUSEI_NO")
            i = i + 1
            
            w_Rs.MoveNext
            
        Loop

		For i = 1 to w_rCnt

			If m_sGakusei(i) = m_sGakuNo Then

				If i <= 1 Then
					m_sGakuNo      = m_sGakusei(i)
	                m_sAfterGakuNo = m_sGakusei(i+1)
					Exit For
				End If

				If i = w_rCnt Then
					m_sGakuNo      = m_sGakusei(i)
	                m_sBeforGakuNo = m_sGakusei(i-1)
					Exit For
				End If

				m_sGakuNo      = m_sGakusei(i)
                m_sAfterGakuNo = m_sGakusei(i+1)
                m_sBeforGakuNo = m_sGakusei(i-1)
				
				Exit For
			End If

		Next

End Function


Function f_getGakuseki_No()
'********************************************************************************
'*  [�@�\]  �w���̊w��NO���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	Dim rs
	Dim w_sSQL

    On Error Resume Next
    Err.Clear

    f_getGakuseki_No = ""

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     T13_GAKUSEKI_NO"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & "     AND T13_GAKUSEI_NO = '" & m_sGakuNo & "' "

        w_iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            Exit Do 
        End If

		If rs.EOF = False Then
			w_iGakusekiNo = rs("T13_GAKUSEKI_NO")
		End If

        Exit Do
    Loop

	'//�߂�l�Z�b�g
    f_getGakuseki_No = w_iGakusekiNo

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �����R���{���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSiken(p_sShiken)
    Dim w_sSQL,w_Rs

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT"
		w_sSQL = w_sSQL & vbCrLf & " M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & "  FROM"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & "  WHERE M01_NENDO = " & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "    AND M01_SYOBUNRUI_CD = " & cint(p_sShiken)

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If
	p_sShiken = w_Rs("M01_SYOBUNRUIMEI")

        f_GetSiken = 0
        Exit Do
    Loop
	Call gf_closeObject(w_Rs)

End Function

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
<link rel="stylesheet" href="../../common/style.css" type="text/css">

<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="JavaScript">
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
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">
<center>
<%call gs_title("�����������o�^","�o�@�^")%>
<BR>
<table border="0" width="500" class=hyo align="center">
	<tr>
		<th width="500" class="header2" colspan="4"><%=m_sShiken%></th>
	</tr>
	<tr>
		<th width="50" class="header">�N���X</th>
		<td width="150" align="center" class="detail"><%=m_sGakunen%>�N�@<%=m_sClassNm%></td>
		<th width="50" class="header">���@��</th>
		<td width="250" align="left" class="detail">�@( <%=f_getGakuseki_No() & " )�@" & m_GRs("T11_SIMEI")%></td>

	</tr>
</table>
<br>
<div align="center"><span class=CAUTION>�� ��O�֣����֣�̃{�^�����N���b�N�����ꍇ�A���͂��ꂽ���̂��ۑ�����A<br>
										���ݓ��͂���Ă���w���̑O�܂��́A��̊w���̏����͂Ɉڂ�܂��B
</span></div>


</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
