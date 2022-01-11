<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0100_upd.asp
' �@      �\: ���y�[�W ���ѓo�^�̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
' ��      ��:
' ��      �n:
' ��      ��:
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2001/07/27 �O�c �q�j
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<!--#include file="sei0100_upd_func.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
    Const DebugPrint = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Dim     m_sKyokanCd     '//����CD
    Dim     m_iNendo 
    Dim     m_sSikenKBN
    Dim     m_sKamokuCd
    Dim     i_max 
    Dim     m_sGakuNo	'//�w�N
    Dim     m_sGakkaCd	'//�w��

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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���ѓo�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_sKyokanCd     = request("txtKyokanCd")
    m_iNendo        = request("txtNendo")
	m_sSikenKBN     = Cint(request("txtSikenKBN"))
	m_sKamokuCd     = request("KamokuCd")
	i_max           = request("i_Max")
	m_sGakuNo	    = Cint(request("txtGakuNo"))	'//�w�N
	m_sGakkaCd	    = request("txtGakkaCd")			'//�w��

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

        '// ���ѓo�^
        w_iRet = f_Update(m_sSikenKBN)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'//�����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
		'//�O���݂̂̏ꍇ�́A�擾�����f�[�^��������������ɂ��o�^����
		If cint(m_sSikenKBN) = cint(C_SIKEN_ZEN_KIM) Then    'C_SIKEN_ZEN_KIM :�O����������(=2)

			'//�����Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
			w_iRet = f_SikenInfo(w_bZenkiOnly)
			If w_iRet<> 0 Then
				Exit Do
			End If 

			If w_bZenkiOnly = True Then
		        '// ���ѓo�^(�O���݂̂̎����Ȗڂ̏ꍇ)
		        w_iRet = f_Update(C_SIKEN_KOU_KIM)
		        If w_iRet <> 0 Then
		            m_bErrFlg = True
		            Exit Do
		        End If

			End If

		End If

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

Function f_Update(p_sSikenKBN)
'********************************************************************************
'*  [�@�\]  �w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Dim i
Dim w_Today

    On Error Resume Next
    Err.Clear

    f_Update = 99

    Do 
		w_Today = gf_YYYY_MM_DD(m_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		
		'// ���Z�敪�擾(sei0100_upd_func.asp���֐�)
		If Not Incf_SelGenzanKbn() Then Exit Function

		'// ���ہE���Ȑݒ�擾(sei0100_upd_func.asp���֐�)
		If Not Incf_SelM15_KEKKA_KESSEKI() then Exit Function

		'// �ݐϋ敪�擾(sei0100_upd_func.asp���֐�)
		If Not Incf_SelKanriMst(m_iNendo,C_K_KEKKA_RUISEKI) then Exit Function

		For i=1 to i_max

			'// �����Ǝ��Ԏ擾(sei0100_upd_func.asp���֐�)
			Call Incs_GetJituJyugyou(i)

			'// �w�����̏ꍇ�A�Œ᎞�Ԃ��擾����
			if Cint(m_sSikenKBN) = C_SIKEN_KOU_KIM then
				'// �Œ᎞�Ԏ擾(sei0100_upd_func.asp���֐�)
				If Not Incf_GetSaiteiJikan(i) then Exit Function
			End if

            '//T16_RISYU_KOJIN��UPDATE
            w_sSQL = ""
            w_sSQL = w_sSQL & vbCrLf & " UPDATE T16_RISYU_KOJIN SET "

		'Select Case m_sSikenKBN
		Select Case p_sSikenKBN

			Case C_SIKEN_ZEN_TYU
				if request("hidUpdFlg" & i ) Then
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_TYUKAN_Z		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_TYUKAN_Z	= '" & request("Hyoka"&i) & "', "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_TYUKAN_Z		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_TYUKAN_Z	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_TYUKAN_Z		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
				End if
				w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_TYUKAN_Z    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_TYUKAN_Z   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_TYUKAN_Z = " & f_CnvNumNull(m_iJituJyugyou) & ","
				
				w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_TYUKAN_Z = '" & w_Today & "',"
				
			Case C_SIKEN_ZEN_KIM
				if request("hidUpdFlg" & i ) Then
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_KIMATU_Z		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_KIMATU_Z	= '" & request("Hyoka"&i) & "', "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_KIMATU_Z		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_KIMATU_Z	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_KIMATU_Z		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
				End if
				w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_KIMATU_Z    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_KIMATU_Z   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_KIMATU_Z = " & f_CnvNumNull(m_iJituJyugyou) & ","
				
				w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_KIMATU_Z = '" & w_Today & "',"
				
			Case C_SIKEN_KOU_TYU
				if request("hidUpdFlg" & i ) Then
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_TYUKAN_K		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_TYUKAN_K	= '" & request("Hyoka"&i) & "', "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_TYUKAN_K		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_TYUKAN_K	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_TYUKAN_K		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
				End if
				w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_TYUKAN_K    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_TYUKAN_K   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_TYUKAN_K = " & f_CnvNumNull(m_iJituJyugyou) & ","
				
				w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_TYUKAN_K = '" & w_Today & "',"
				
			Case C_SIKEN_KOU_KIM
				if request("hidUpdFlg" & i ) Then
					w_sSQL = w_sSQL & vbCrLf & " 	T16_SEI_KIMATU_K		= " & f_CnvNumNull(request("Seiseki"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_HYOKAYOTEI_KIMATU_K	= '" & request("Hyoka"&i) & "', "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_KIMATU_K		= " & f_CnvNumNull(request("Kekka"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KEKA_NASI_KIMATU_K	= " & f_CnvNumNull(request("KekkaGai"&i)) & ", "
					w_sSQL = w_sSQL & vbCrLf & " 	T16_CHIKAI_KIMATU_K		= " & f_CnvNumNull(request("Chikai"&i)) & ", "
				End if
				w_sSQL = w_sSQL & vbCrLf & " 	T16_SOJIKAN_KIMATU_K    = " & f_CnvNumNull(m_iSouJyugyou)  & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_JUNJIKAN_KIMATU_K   = " & f_CnvNumNull(m_iJunJyugyou)  & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_J_JUNJIKAN_KIMATU_K = " & f_CnvNumNull(m_iJituJyugyou) & ","
				w_sSQL = w_sSQL & vbCrLf & " 	T16_SAITEI_JIKAN        = " & f_CnvNumNull(m_iSaiteiJikan) & ","
				
				w_sSQL = w_sSQL & vbCrLf & " 	T16_KOUSINBI_KIMATU_K = '" & w_Today & "',"
				
				if Not gf_IsNull(m_iKyuSaiteiJikan) Then
					w_sSQL = w_sSQL & vbCrLf & " 	T16_KYUSAITEI_JIKAN = " & f_CnvNumNull(m_iKyuSaiteiJikan) & ","
				End if
		End Select

            w_sSQL = w_sSQL & vbCrLf & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
            w_sSQL = w_sSQL & vbCrLf & "   T16_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T16_NENDO = " & Cint(m_iNendo) & " "
            w_sSQL = w_sSQL & vbCrLf & "    AND T16_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "    AND T16_KAMOKU_CD = '" & Trim(m_sKamokuCd) & "'  "
			
            iRet = gf_ExecuteSQL(w_sSQL)
			
            If iRet <> 0 Then
                '//۰��ޯ�
                msMsg = Err.description
                Exit Do
            End If

		Next

        '//����I��
        f_Update = 0

        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  ���l�^���ڂ̍X�V���̐ݒ�
'*  [����]  �l
'*  [�ߒl]  �Ȃ�
'*  [����]  ���l�������Ă���ꍇ��[�l]�A�����ꍇ��"NULL"��Ԃ�
'********************************************************************************
Function f_CnvNumNull(p_vAtai)

	If Trim(p_vAtai) = "" Then
		f_CnvNumNull = "NULL"
	Else
		f_CnvNumNull = cInt(p_vAtai)
    End If

End Function

'********************************************************************************
'*  [�@�\]  �����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_SikenInfo(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_SikenInfo = 1
	p_bZenkiOnly = false

    Do 

'		'//�����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSQL = w_sSQL & vbCrLf & " T15_RISYU.T15_KAMOKU_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO=" & Cint(m_iNendo)-cint(m_sGakuNo)+1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & Trim(m_sKamokuCd) & "'" 
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAISETU" & m_sGakuNo & "=" & C_KAI_ZENKI	'//�O���J��

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_SikenInfo = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_bZenkiOnly = True
		End If

        f_SikenInfo = 0
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
%>
    <html>
    <head>
    <title>���ѓo�^</title>
    <link rel=stylesheet href="../../common/style.css" type=text/css>

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

		alert("<%=C_TOUROKU_OK_MSG%>");

	    document.frm.target = "main";
	    document.frm.action = "./sei0100_bottom.asp"
	    document.frm.submit();
	    return;

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

	<input type=hidden name=txtNendo    value="<%=trim(Request("txtNendo"))%>">
	<input type=hidden name=txtKyokanCd value="<%=trim(Request("txtKyokanCd"))%>">
	<input type=hidden name=txtSikenKBN value="<%=trim(Request("txtSikenKBN"))%>">
	<input type=hidden name=txtGakuNo   value="<%=trim(Request("txtGakuNo"))%>">
	<input type=hidden name=txtClassNo  value="<%=trim(Request("txtClassNo"))%>">
	<input type=hidden name=txtKamokuCd value="<%=trim(Request("txtKamokuCd"))%>">
	<input type=hidden name=txtGakkaCd  value="<%=trim(Request("txtGakkaCd"))%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>