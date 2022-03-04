<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���i���Ґ��ѓo�^
' ��۸���ID : sei/sei0900/sei0900_upd.asp
' �@      �\: ���y�[�W ���i���Ґ��ѓo�^�̓o�^�A�X�V
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
' ��      ��:
' ��      �n:
' ��      ��:
'           �����̓f�[�^�̓o�^�A�X�V���s��
'-------------------------------------------------------------------------
' ��      ��: 2022/2/1 �g�c�@�Ď������ѓo�^��ʂ𗬗p���쐬
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<!--#include file="sei0900_upd_func.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
    Const DebugPrint = 0
    Public Const C_GOUKAKUTEN = 60  '���i�_
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_Rs_Hyoka			'�]�����
	
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
    Dim     m_iRisyuKakoNendo   '//�ߔN�x 
    Dim     m_iHaitotani   '//�z���P��
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
    w_sMsgTitle="���i���Ґ��ѓo�^"
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
    m_iRisyuKakoNendo =  request("txtRisyuKakoNendo")
    m_iHaitotani =  request("txtHaitoTani")

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
    ' response.write  "w_Sisekiarray0:" & w_Sisekiarray(0)

    Do 
		w_Today = gf_YYYY_MM_DD(m_iNendo & "/" & month(date()) & "/" & day(date()),"/")
		
		m_SchoolFlg = cbool(request("hidSchoolFlg"))
		
        '// �Ȗڕ]���擾
        w_iRet = f_GetKamokuTensuHyoka(m_iRisyuKakoNendo,m_sKamokuCd)
        If w_iRet<> 0 Then
            Exit Do
        End If

		For i=1 to i_max : Do
            '���т��ۂ̏ꍇ�́A���̊w����
            if w_Sisekiarray(i-1) = 2 then Exit Do
			       
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

			'//T17_RISYUKAKO_KOJIN��UPDATE
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " UPDATE T17_RISYUKAKO_KOJIN SET "
			w_sSQL = w_sSQL & vbCrLf & "   T17_SEI_KIMATU_K = " & C_GOUKAKUTEN  & ","
            w_sSQL = w_sSQL & vbCrLf & "   T17_HYOKA_KIMATU_K = '" &  m_Rs_Hyoka("M08_HYOKA_SYOBUNRUI_MEI")  & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T17_GPA_KIMATU_K = " & m_Rs_Hyoka("M08_HYOTEN_GPA")   & ","
            w_sSQL = w_sSQL & vbCrLf & "   T17_TANI_SUMI = " & m_iHaitotani  & ","
            w_sSQL = w_sSQL & vbCrLf & "   T17_KOUSINBI_KIMATU_K = '" & gf_YYYY_MM_DD(date(),"/") & "',"
            w_sSQL = w_sSQL & vbCrLf & "   T17_UPD_DATE = '" & gf_YYYY_MM_DD(date(),"/") & "', "
            w_sSQL = w_sSQL & vbCrLf & "   T17_UPD_USER = '"  & Trim(Session("LOGIN_ID")) & "' "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "        T17_NENDO = " & Cint(m_iRisyuKakoNendo) & " "
            w_sSQL = w_sSQL & vbCrLf & "    AND T17_GAKUSEI_NO = '" & Trim(request("txtGseiNo"&i)) & "'  "
            w_sSQL = w_sSQL & vbCrLf & "    AND T17_KAMOKU_CD = '" & Trim(m_sKamokuCd) & "'  "

            If gf_ExecuteSQL(w_sSQL) <> 0 Then
                '//۰��ޯ�
                msMsg = Err.description
                Exit Do
            End If
        ' response.write w_sSQL & "<BR>"
        '   response.end
	    Loop Until 1: Next
		' response.end
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
'*  [�@�\]  �Ȗڕ]���擾
'*  [����]  �l
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKamokuTensuHyoka(p_iNendo,p_sKamokuCD)

    f_GetKamokuTensuHyoka = 1
    Dim w_iZokuseiCD         '�Ȗڑ���
    Dim w_iHyokaNo
    Dim w_iRet

    ' �Ȗڑ����擾
	w_iRet = f_GetKamokuZokusei(p_iNendo,p_sKamokuCD,w_iZokuseiCD)
    If w_iRet<> 0 Then
            Exit Function
    End If
    
    '�Ȗڑ�������]��NO�擾
    w_iRet = f_iGetHyokaNo(p_iNendo,w_iZokuseiCD,w_iHyokaNo) 
    If w_iRet<> 0 Then
            Exit Function
    End If
    
    '�]��NO����]���f�[�^�擾
    w_iRet = f_GetTensuHyoka(p_iNendo,w_iZokuseiCD,C_GOUKAKUTEN) 
    If w_iRet<> 0 Then
            Exit Function
    End If

    f_GetKamokuTensuHyoka = 0
End Function

'********************************************************************************
'*  [�@�\]  �Ȗڑ����擾(�ʏ펞)
'*  [����]  p_iNendo - �N�x(IN)
'        �@ p_sKamokuCD - �ȖڃR�[�h(IN)
'           p_iZokuseiCD - �����R�[�h(OUT)
'*  [�ߒl]   
'*  [����]  
'********************************************************************************
Function f_GetKamokuZokusei(p_iNendo,p_sKamokuCD, p_iZokuseiCD)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKamokuZokusei = 1
	p_bZenkiOnly = false

    Do 

'		'//�Ȗڑ����擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSql = w_sSql & vbCrLf & " M03_ZOKUSEI_CD"
        w_sSql = w_sSql & vbCrLf & " FROM"
        w_sSql = w_sSql & vbCrLf & " M03_KAMOKU"
        w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " M03_NENDO=" & Cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & " AND M03_KAMOKU_CD='" & Trim(m_sKamokuCd) & "'" 

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKamokuZokusei = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_iZokuseiCD = w_Rs("M03_ZOKUSEI_CD")
		End If

        f_GetKamokuZokusei = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �]���`��No���擾����
'*  [����]  p_iNendo - �N�x(IN)
'           p_iKamokuZokusei_CD - �Ȗڑ����R�[�h(IN)
'           p_iHYOKA_NO - �]���`��No(OUT)
'*  [�ߒl]   
'*  [����]  
'********************************************************************************
Function f_iGetHyokaNo(p_iNendo,p_iKamokuZokusei_CD,p_iHYOKA_NO)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_iGetHyokaNo = 1
	p_bZenkiOnly = false

    Do 

'		'//�]���`��No���擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSql = w_sSql & vbCrLf & " M100_HYOUKA_NO"
        w_sSql = w_sSql & vbCrLf & " FROM"
        w_sSql = w_sSql & vbCrLf & " M100_KAMOKU_ZOKUSEI"
        w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " M100_NENDO=" & Cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & " AND M100_ZOKUSEI_CD='" & Trim(p_iKamokuZokusei_CD) & "'" 

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_iGetHyokaNo = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_iHYOKA_NO = w_Rs("M100_HYOUKA_NO")
		End If

        f_iGetHyokaNo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �Ȗڕ]���擾(�]��=�̃f�[�^)
'*  [����]  p_iNendo - �N�x(IN)
'           p_iKamokuZokusei_CD - �Ȗڑ����R�[�h(IN)
'           p_iHYOKA_NO - �]���`��No(OUT)
'*  [�ߒl]   
'*  [����]  
'********************************************************************************
Function f_GetTensuHyoka(p_iNendo,p_iHYOKA_NO,p_iTensu)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetTensuHyoka = 1
	p_bZenkiOnly = false

    Do 

'		'//�]���`��No���擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
 		w_sSql = w_sSql & vbCrLf & " M08_HYOKA_SYOBUNRUI_MEI,"
        w_sSql = w_sSql & vbCrLf & " M08_HYOTEI,"
        w_sSql = w_sSql & vbCrLf & " M08_HYOKA_SYOBUNRUI_RYAKU,"
        w_sSql = w_sSql & vbCrLf & " M08_HYOTEN_GPA"
        w_sSql = w_sSql & vbCrLf & " FROM"
        w_sSql = w_sSql & vbCrLf & " M08_HYOKAKEISIKI"
        w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " M08_NENDO=" & Cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & " AND M08_HYOUKA_NO='" & p_iHYOKA_NO & "'" 
        w_sSQL = w_sSQL & vbCrLf & " AND M08_MIN <= '" & p_iTensu & "'" 
        w_sSQL = w_sSQL & vbCrLf & " AND M08_MAX >= '" & p_iTensu & "'" 
        w_sSQL = w_sSQL & vbCrLf & " AND M08_HYOKA_TAISYO_KBN ='" & C_HYOKA_TAISHO_IPPAN & "'" 

' response.write w_sSQL
'response.end
        iRet = gf_GetRecordset(m_Rs_Hyoka, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetTensuHyoka = 99
            Exit Do
        End If

        '//����I��
        f_GetTensuHyoka = 0
        Exit Do
    Loop

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
    <title>���i���Ґ��ѓo�^</title>
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
	    document.frm.action = "./sei0900_bottom.asp"
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
	<input type=hidden name=txtKamokuCd value="<%=trim(Request("txtKamokuCd"))%>">
    <input type="hidden" name="txtKamokuNM" value="<%=trim(Request("txtKamokuNM"))%>"">
    <input type="hidden" name="txtRisyuKakoNendo" value="<%=m_iRisyuKakoNendo%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>