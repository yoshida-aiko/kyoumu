<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0500/sei0500_upd.asp
' �@      �\: ���y�[�W ���ѓo�^�̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
' ��      ��:
' ��      �n:
' ��      ��:
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2001/09/07 ���`�i�K
' ��      �X: 2016/05/18 Nishimura �ٓ�(�x�w��)�̏ꍇ�X�V�ł��Ȃ���Q�Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���͎������ѓo�^"
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

        '//��ݻ޸��݊J�n
        Call gs_BeginTrans()


		'// ���т��X�V����
		If f_SeisekiUpdate() then
			'//�Я�
			Call gs_CommitTrans()
		Else
			'// ۰��ޯ�
			Call gs_RollbackTrans()
			Exit Do
		End if

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

Function f_SeisekiUpdate()
'********************************************************************************
'*  [�@�\]  ���т��X�V����
'*  [����]  �Ȃ�
'*  [�ߒl]  True : False
'*  [����]  
'********************************************************************************
Dim i

    On Error Resume Next
    Err.Clear
    
    f_SeisekiUpdate = False

	'// ���Ұ��擾
	w_iNendo	= request("txtNendo")
	w_sKyokanCd	= request("txtKyokanCd")
	w_sSiKenCd	= Cint(request("txtShikenCd"))
	w_sGakuNo	= Cint(request("txtGakuNo"))
	w_sClassNo	= Cint(request("txtClassNo"))
	w_sKamokuCd	= request("txtKamokuCd")

	m_rCnt      = Cint(request("hidRecCnt"))

	i = 1
	Do Until i > m_rCnt

		w_iIdoCnt = request("hidIdoCnt" & i )	'Ins 2016/05/18 Nishimura �x�w�҂̏ꍇ�G���[�ɂȂ邽�߁AIF�� ����ǉ�
		IF w_iIdoCnt = 1 Then					'Ins 2016/05/18 Nishimura 

		w_SQL = ""
		w_SQL = w_SQL & vbCrLf & " Update T33_SIKEN_SEISEKI Set "
'2015/10/20 UPDATE URAKAWA NULL�̎�0��o�^���Ȃ��B
'		w_SQL = w_SQL & vbCrLf & "	T33_TOKUTEN		 =  " & Cint(gf_SetNull2Zero(request("Seiseki" & i))) & ", "
		w_SQL = w_SQL & vbCrLf & "	T33_TOKUTEN		 =  '" & gf_SetNull2String(request("Seiseki" & i)) & "', "
		w_SQL = w_SQL & vbCrLf & "	T33_UPD_DATE	 = '" & gf_YYYY_MM_DD(date(),"/") & "',"
		w_SQL = w_SQL & vbCrLf & "	T33_UPD_USER	 = '" & w_sKyokanCd & "'"
		w_SQL = w_SQL & vbCrLf & " WHERE "
		w_SQL = w_SQL & vbCrLf & "	T33_NENDO		 =  " & w_iNendo & " AND"
		w_SQL = w_SQL & vbCrLf & "	T33_SIKEN_KBN	 =  " & C_SIKEN_JITURYOKU & " AND"
		w_SQL = w_SQL & vbCrLf & "	T33_SIKEN_CD	 =  " & w_sSiKenCd & " AND"
		w_SQL = w_SQL & vbCrLf & "	T33_SIKEN_KAMOKU = '" & w_sKamokuCd & "' AND"
		w_SQL = w_SQL & vbCrLf & "	T33_GAKUSEKI_NO  = '" & request("hidGakusekiNo" & i ) & "' AND"
		w_SQL = w_SQL & vbCrLf & "	T33_GAKUNEN  	 =  " & w_sGakuNo & " AND"
		w_SQL = w_SQL & vbCrLf & "	T33_CLASS		 =  " & w_sClassNo

		iRet = gf_ExecuteSQL(w_SQL)
		If iRet <> 0 Then
			m_bErrFlg = True
			msMsg = Err.description
			Exit Function
		End If

		END IF	'Ins 2016/05/18 Nishimura 

		i = i + 1
	Loop

	'//����I��
	f_SeisekiUpdate = True

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
	    document.frm.action = "sei0500_bottom.asp"
	    document.frm.submit();

    }
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">

	<input type=hidden name=txtNendo    value="<%=trim(Request("txtNendo"))%>">
	<input type=hidden name=txtKyokanCd value="<%=trim(Request("txtKyokanCd"))%>">
	<input type=hidden name=txtShikenCd value="<%=trim(Request("txtShikenCd"))%>">
	<input type=hidden name=txtGakuNo   value="<%=trim(Request("txtGakuNo"))%>">
	<input type=hidden name=txtClassNo  value="<%=trim(Request("txtClassNo"))%>">
	<input type=hidden name=txtKamokuCd value="<%=trim(Request("txtKamokuCd"))%>">

	</form>
    </body>
    </html>
<%
End Sub
%>

