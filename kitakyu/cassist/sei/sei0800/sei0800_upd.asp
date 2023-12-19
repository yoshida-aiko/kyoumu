<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �Ď������ѓo�^
' ��۸���ID : sei/sei0800/sei0800_upd.asp
' �@      �\: ���y�[�W �Ď������ѓo�^�̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
' ��      ��:
' ��      �n:
' ��      ��:
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2021/12/23 �g�c�@���ѓo�^��ʂ𗬗p���쐬
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<!--#include file="sei0800_upd_func.asp"-->
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
    Dim     m_SchoolFlg
    Dim     m_SQL
    Dim     hidSeiseki
    Dim     m_UpdateDate
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
    w_sMsgTitle="�Ď������ѓo�^"
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
    m_UpdateDate	= request("txtUpdateDate")			'//�w��

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
Dim w_DataKbnFlg
Dim w_DataKbn
Dim w_Sisekiarray

    On Error Resume Next
    Err.Clear
	
    f_Update = 99
	w_DataKbnFlg = false
	w_DataKbn = 0
    w_Sisekiarray = split(Trim(request("hidSeiseki")),",")
    'response.write  "w_Sisekiarray0:" & w_Sisekiarray(0)
    ' response.end
    Do 
		w_Today = gf_YYYY_MM_DD(m_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		
		m_SchoolFlg = cbool(request("hidSchoolFlg"))
		
		'// ���Z�敪�擾(sei0800_upd_func.asp���֐�)
		If Not Incf_SelGenzanKbn() Then Exit Function
		
		'// ���ہE���Ȑݒ�擾(sei0800_upd_func.asp���֐�)
		If Not Incf_SelM15_KEKKA_KESSEKI() then Exit Function
		
		'// �ݐϋ敪�擾(sei0800_upd_func.asp���֐�)
		If Not Incf_SelKanriMst(m_iNendo,C_K_KEKKA_RUISEKI) then Exit Function

		For i=1 to i_max : Do
            '���т��ۂ̏ꍇ�́A���̊w����
            if w_Sisekiarray(i-1) = 2 then Exit Do
              
            '//�����Ǝ��Ԏ擾(sei0800_upd_func.asp���֐�)
            Call Incs_GetJituJyugyou(i)
            
            '//�w�����̏ꍇ�A�Œ᎞�Ԃ��擾����
            if cInt(m_sSikenKBN) = C_SIKEN_KOU_KIM then
                '//�Œ᎞�Ԏ擾(sei0800_upd_func.asp���֐�)
                If Not Incf_GetSaiteiJikan(i) then Exit Function
            End if
            
            if m_SchoolFlg = true then
                w_DataKbn = 0
                w_DataKbnFlg = false
                
                '//���]���A�]���s�\�̐ݒ�
                if cint(gf_SetNull2Zero(request("hidMihyoka"))) <> 0 then
                    w_DataKbn = cint(gf_SetNull2Zero(request("hidMihyoka")))
                    w_DataKbnFlg = true
                else
                    w_DataKbn = cint(gf_SetNull2Zero(request("chkHyokaFuno" & i)))
                    
                    if w_DataKbn = cint(C_HYOKA_FUNO) then
                        w_DataKbnFlg = true
                    end if
                end if
            end if

            
			'//T16_RISYU_KOJIN��UPDATE
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " UPDATE T16_RISYU_KOJIN SET "
			' w_sSQL = w_sSQL & vbCrLf & "   T16_SEI_KIMATU_K = " & C_GOUKAKUTEN  & ","
            '2023.09.07 Add Kiyomoto �O���I���Ȗڂ͑O���������т��X�V -->
            w_sSQL = w_sSQL & vbCrLf & "   T16_SEI_KIMATU_Z = CASE WHEN T16_KAISETU = " & C_KAI_ZENKI & " THEN 60 "
            w_sSQL = w_sSQL & vbCrLf & "                           ELSE T16_SEI_KIMATU_Z END,"
            w_sSQL = w_sSQL & vbCrLf & "   T16_KOUSINBI_KIMATU_Z = CASE WHEN T16_KAISETU = 1 THEN '"& gf_YYYY_MM_DD(date(),"/") & "'"
            w_sSQL = w_sSQL & vbCrLf & "                           ELSE T16_KOUSINBI_KIMATU_Z END,"
            '2023.09.07 Add Kiyomoto �O���I���Ȗڂ͑O���������т��X�V <--
            w_sSQL = w_sSQL & vbCrLf & "   T16_SEI_KIMATU_K = 60,"
            w_sSQL = w_sSQL & vbCrLf & "   T16_KOUSINBI_KIMATU_K = '" & gf_YYYY_MM_DD(date(),"/") & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T16_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
            w_sSQL = w_sSQL & vbCrLf & "   T16_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T16_NENDO = " & Cint(m_iNendo) & " "
            w_sSQL = w_sSQL & vbCrLf & "    AND T16_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "    AND T16_KAMOKU_CD = '" & Trim(m_sKamokuCd) & "'  "

            If gf_ExecuteSQL(w_sSQL) <> 0 Then
                '//۰��ޯ�
                msMsg = Err.description
                Exit Do
            End If
            ' response.write  "txtGseiNo:" & Trim(request("txtGseiNo"&i)) & "<BR>"
            ' response.write w_sSQL & "<BR>"
        Loop Until 1: Next
		'response.end
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
    <title>�Ď������ѓo�^</title>
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

	   //alert("<%=m_SQL%>");
	    alert("<%=C_TOUROKU_OK_MSG%>");

	    document.frm.target = "main";
	    document.frm.action = "./sei0800_bottom.asp"
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
    <input type=hidden name=txtUpdateDate  value="<%=m_UpdateDate%>">
    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>