<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0112/kks0112_edit.asp
' �@      �\: �O�y�[�W��(kks0112_bottom.asp)�o�^�����o���󋵂�o�^����
'-------------------------------------------------------------------------
' ��      ��: 
'             
'             
'             
'             
' ��      ��: 
' ��      �n: 
'             
'             
'             
'             
' ��      ��: 
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2002/05/16 shin
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Dim m_bErrFlg           '�G���[�t���O
	
    '�擾�����f�[�^�����ϐ�
    Dim m_iSyoriNen
    Dim m_iKyokanCd
    
	Dim m_sGakunen
	Dim m_sClassNo
	Dim m_sKamokuCd
	
	Dim m_sDate
	Dim m_iJigen
	Dim m_sUserId
	Dim m_iKamokuKbn
	
'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()
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
        If gf_OpenDatabase() <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
        '//�ϐ�������
        Call s_ClearParam()
		
        '//Main���Ұ�SET
        Call s_SetParam()
		
        '//���ȕʏo���o�^
        If not f_AbsEdit() Then
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

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()
	m_sUserId = ""
    m_iSyoriNen = ""
    m_iKyokanCd = ""
    
	m_sGakunen = ""
	m_sClassNo = ""
	m_sKamokuCd = ""
	
	m_sDate = ""
	m_iJigen = ""
	m_iKamokuKbn = 0
	
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
	m_sUserId = Session("LOGIN_ID")
	
    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")
    
	m_sGakunen = trim(Request("hidGakunen"))
	m_sClassNo = trim(Request("hidClassNo"))
	m_sKamokuCd = trim(Request("hidKamokuCd"))
	
	m_sDate = gf_YYYY_MM_DD(trim(Request("hidDate")),"/")
	m_iJigen = trim(Request("hidJigen"))
	
	m_iKamokuKbn = cint(request("hidSyubetu"))
	
End Sub

'********************************************************************************
'*  [�@�\]  ���ȕʏo���o�^
'*  [����]  �Ȃ�
'*  [�ߒl]  false:���擾���� true:���s
'*  [����]  �f���[�g��A�C���T�[�g����
'********************************************************************************
Function f_AbsEdit()
	Dim w_sSQL
    Dim w_Rs
    Dim w_iKekka
	Dim w_sUserId
	Dim w_sGakusekiNo,w_iCount
	Dim w_State
	Dim w_JikanNum
	
    On Error Resume Next
    Err.Clear
    
    f_AbsEdit = false
	
	'//�w��NO
	w_sGakusekiNo = split(replace(Request("hidGakusekiNo")," ",""),",")
	
	'//�w����
	w_iCount = UBound(w_sGakusekiNo)
	
	'//��ݻ޸��݊J�n
    Call gs_BeginTrans()
	
	'//delete
	if not f_AbsDelete() then
		'//۰��ޯ�
		Call gs_RollbackTrans()
		exit function
	end if
	
	for i=0 to w_iCount
		w_State = 0
		w_State = gf_SetNull2Zero(trim(request("hidState" & w_sGakusekiNo(i))))
		
		w_JikanNum = 0
		w_JikanNum = gf_SetNull2Zero(trim(request("hidJikanState" & w_sGakusekiNo(i))))
		
		'//�󋵂��I������Ă�����Ainsert
		if w_State <> 0 then
			if not f_AbsInsert(w_sGakusekiNo(i),w_State,w_JikanNum) then
				'//۰��ޯ�
				Call gs_RollbackTrans()
				exit function
			end if
		end if
	next
	
	'//�Я�
	Call gs_CommitTrans()
    
    f_AbsEdit = true
    
End Function

'********************************************************************************
'*  [�@�\]  
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
Function f_AbsInsert(p_GakusekiNo,p_State,p_JikanNum)
	
	Dim w_sSQL
    
    On Error Resume Next
    Err.Clear
    
    f_AbsInsert = false
	
	'if cInt(p_State) <> 1 then p_JikanNum = 0
	
    w_sSQL = ""
    w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T21_SYUKKETU  "
    w_sSQL = w_sSQL & vbCrLf & "   ("
    w_sSQL = w_sSQL & vbCrLf & "  T21_NENDO, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_HIDUKE, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_YOUBI_CD, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_GAKUNEN, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_CLASS, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_GAKUSEKI_NO, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_JIGEN, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_KAMOKU, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_KYOKAN, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU_KBN, "
	w_sSQL = w_sSQL & vbCrLf & "  T21_JIKANSU, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_JIMU_FLG, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_KAMOKU_KBN, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_INS_DATE, "
    w_sSQL = w_sSQL & vbCrLf & "  T21_INS_USER"
    w_sSQL = w_sSQL & vbCrLf & "   )VALUES("
    w_sSQL = w_sSQL & vbCrLf & "    "  & m_iSyoriNen				& " ,"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & m_sDate					& "',"
    w_sSQL = w_sSQL & vbCrLf & "    "  & cint(Weekday(m_sDate))		& ","
    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_sGakunen)			& " ,"
    w_sSQL = w_sSQL & vbCrLf & "    "  & cInt(m_sClassNo)			& " ,"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & p_GakusekiNo				& "',"
    w_sSQL = w_sSQL & vbCrLf & "    "  & m_iJigen					& " ,"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & m_sKamokuCd				& "',"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & Trim(m_iKyokanCd)			& "',"
    w_sSQL = w_sSQL & vbCrLf & "   "   & p_State					& ","
    w_sSQL = w_sSQL & vbCrLf & "   "   & p_JikanNum					& ","
    w_sSQL = w_sSQL & vbCrLf & "   "   & C_JIMU_FLG_NOTJIMU			& ","
    w_sSQL = w_sSQL & vbCrLf & "   "   & m_iKamokuKbn				& ","
    w_sSQL = w_sSQL & vbCrLf & "   '"  & gf_YYYY_MM_DD(date(),"/")	& "',"
    w_sSQL = w_sSQL & vbCrLf & "   '"  & m_sUserId					& "' "
    w_sSQL = w_sSQL & vbCrLf & "   )"
	
	if gf_ExecuteSQL(w_sSQL) <> 0 Then exit function
    
	'//����I��
    f_AbsInsert = true
    
End Function


'********************************************************************************
'*  [�@�\]  
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
Function f_AbsDelete()
	
	Dim w_sSQL
    
    On Error Resume Next
    Err.Clear
    
    f_AbsDelete = false
	
    w_sSQL = ""
    w_sSQL = w_sSQL & " delete from T21_SYUKKETU  "
    w_sSQL = w_sSQL & "  where "
    w_sSQL = w_sSQL & "      T21_NENDO			=  " & m_iSyoriNen
    w_sSQL = w_sSQL & "  and T21_HIDUKE			= '" & m_sDate				& "' "
    w_sSQL = w_sSQL & "  and T21_GAKUNEN		=  " & cInt(m_sGakunen)
    w_sSQL = w_sSQL & "  and T21_CLASS			=  " & cInt(m_sClassNo)
    w_sSQL = w_sSQL & "  and T21_JIGEN			=  " & m_iJigen
	w_sSQL = w_sSQL & "  and T21_KAMOKU			= '" & m_sKamokuCd			& "' "
    w_sSQL = w_sSQL & "  and T21_KYOKAN			= '" & Trim(m_iKyokanCd)	& "' "
    'w_sSQL = w_sSQL & "  and T21_GAKUSEKI_NO	= '" & p_GakusekiNo			& "' "
    
    if gf_ExecuteSQL(w_sSQL) <> 0 Then exit function
    
	'//����I��
    f_AbsDelete = true
    
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
    <title>���Əo������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		
		alert("<%= C_TOUROKU_OK_MSG %>");
		
		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("�ĕ\�����Ă��܂��@���΂炭���҂���������")%>"
		
	    parent.main.document.frm.target = "main";
        parent.main.document.frm.action = "kks0112_bottom.asp"
	    parent.main.document.frm.submit();
	    return;
	}
	
    //-->
    </SCRIPT>
    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
	
	<input TYPE="HIDDEN" NAME="txtURL" VALUE="kks0112_bottom.asp">
    <input TYPE="HIDDEN" NAME="txtMsg" VALUE="<%=Server.HTMLEncode("�ĕ\�����Ă��܂��@���΂炭���҂���������")%>">
	
	<input type="hidden" name="hidGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="hidClassNo" value="<%=m_sClassNo%>">
	<input type="hidden" name="hidKamokuCd" value="<%=m_sKamokuCd%>">
	
	<input type="hidden" name="txtDate" value="<%=m_sDate%>">
	<input type="hidden" name="sltJigen" value="<%=m_iJigen%>">
	
	<input type="hidden" name="hidKamokuName" value="<%=request("hidKamokuName")%>">
	<input type="hidden" name="hidClassName" value="<%=request("hidClassName")%>">
	
	<input type="hidden" name="hidSyubetu" value="<%=m_iKamokuKbn%>">
	
    </form>
    </body>
    </html>
<%
End Sub
%>