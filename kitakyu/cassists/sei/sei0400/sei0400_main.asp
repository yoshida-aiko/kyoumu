<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������������o�^
' ��۸���ID : gak/sei0400/sei0400_main.asp
' �@      �\: ���y�[�W �������������o�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
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
    Public m_sSyoken
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
    Public m_iSikenKBN
    
    Public  m_GRs,m_DRs
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
    w_sMsgTitle="�������������o�^"
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


'2001/12/05 ITO �C�� ���͊����͐ݒ肵�Ȃ��Ă悢
		'===============================
		'//���ԃf�[�^�̎擾
		'===============================
        'w_iRet = f_Nyuryokudate()
		'If w_iRet = 1 Then
			'// �y�[�W��\��
		'	Call No_showPage("���ѓ��͊��ԊO�ł��B")
		'	Exit Do
		'End If
		'If w_iRet <> 0 Then 
		'	m_bErrFlg = True
		'	Exit Do
		'End If


		Call f_Gakusei()

        '//�f�[�^�擾
        w_iRet = f_getdate()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
         '//�w�Ȃb�c�擾
        w_iRet = f_getGakka()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

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
	m_iSikenKBN    = request("txtSikenKBN")
	'//�O��OR���փ{�^���������ꂽ��
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

Function f_Nyuryokudate()
'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
	dim w_date

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1


	w_date = gf_YYYY_MM_DD(date(),"/")
'	w_date = "2000/06/18"

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_iSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & w_date & "' "

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
'//���ѓ��͊��ԃe�X�g�p

'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '2002/04/30'"
'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '2000/03/01'"

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_

'response.write w_sSQL & "<<<BR>"

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If m_DRs.EOF Then
			Exit Do
		Else
			m_sSikenNm = m_DRs("M01_SYOBUNRUIMEI")
		End If
		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function

Function f_Gakusei()
'********************************************************************************
'*  [�@�\]  �����̎������擾
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
    w_sSQL = w_sSQL & "      T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
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


Function f_KYO_MEI(p_sCD,p_iNENDO)
'********************************************************************************
'*  [�@�\]  �����̎������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_Rs

    If Isnull(p_sCD) Then 
        f_KYO_MEI = "" 
        Exit Function
    End If

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     M04_KYOKAN "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     M04_KYOKAN_CD = '" & p_sCD & "' "
    w_sSQL = w_sSQL & " AND M04_NENDO = " & p_iNENDO & " "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(w_Rs, w_sSQL, m_iDsp)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
    End If

    'f_KYO_MEI = w_Rs("M04_KYOKANMEI_SEI")&"�@"&w_Rs("M04_KYOKANMEI_MEI")
    response.write w_Rs("M04_KYOKANMEI_SEI")&"�@"&w_Rs("M04_KYOKANMEI_MEI")

End Function

Function f_SINRO(p_sCD,p_iNENDO)
'********************************************************************************
'*  [�@�\]  �i�H����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_Rs

    If Isnull(p_sCD) Then 
        f_SINRO = "" 
        Exit Function
    End If

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     M32_SINROMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     M32_SINRO "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     M32_SINRO_CD = '" & p_sCD & "' "
    w_sSQL = w_sSQL & " AND M32_NENDO = " & p_iNENDO & " "

    Set w_Rs = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(w_Rs, w_sSQL, m_iDsp)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
    End If

    'f_SINRO = w_Rs("M32_SINROMEI")
    response.write w_Rs("M32_SINROMEI")

End Function

Function f_getdate()
'********************************************************************************
'*  [�@�\]  �f�[�^�̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	dim w_sSikenKBN

    On Error Resume Next
    Err.Clear
    f_getdate = 1

	select case cint(m_iSikenKBN)
		case C_SIKEN_ZEN_TYU
			w_sSikenKBN = "T13_SYOKEN_TYUKAN_Z"
		case C_SIKEN_ZEN_KIM
			w_sSikenKBN = "T13_SYOKEN_KIMATU_Z"
		case C_SIKEN_KOU_TYU
			w_sSikenKBN = "T13_SYOKEN_TYUKAN_K"
		case C_SIKEN_KOU_KIM
			w_sSikenKBN = "T13_SYOKEN_KIMATU_K"
	End select

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     " & w_sSikenKBN & " as Shoken "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_NENDO = " & m_iNendo & " AND "
        w_sSQL = w_sSQL & "     T13_GAKUSEI_NO = '" & m_sGakuNo & "' "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            f_getdate = 99
            m_bErrFlg = True
            Exit Do 
        End If
        m_sSyoken  = m_Rs("Shoken")
        f_getdate = 0
        Exit Do
    Loop

End Function

Function f_getGakka()
'********************************************************************************
'*  [�@�\]  �w���̏����w�Ȃ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_getGakka = 1

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     T13_GAKKA_CD"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_GAKUSEI_NO = '" & m_sGakuNo & "' "

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            f_getGakka = 99
            m_bErrFlg = True
            Exit Do 
        End If

	m_sGakka = m_Rs("T13_GAKKA_CD")
        f_getGakka = 0
        Exit Do
    Loop

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

Sub No_showPage(p_msg)
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>

	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

	<body>
	<center>
	<br><br><br>
			<span class="msg"><%=p_msg%></span>
	</center>
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
<link rel="stylesheet" href="../../common/style.css" type="text/css">

<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="JavaScript">
<!--
	var chk_Flg;
	chk_Flg = false;
	//************************************************************
	//  [�@�\]  �y�[�W���[�h������
	//  [����]
	//  [�ߒl]
	//  [����]
	//************************************************************
	function window_onload() {

        document.frm.target="topFrame";
        document.frm.action="sei0400_topDisp.asp";
        document.frm.submit();

	}

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(p_FLG){

        // �������S�C�����̌�����������
        if( getLengthB(document.frm.Syoken.value) > "200" ){
            window.alert("�S�C�����̗��͑S�p100�����ȓ��œ��͂��Ă�������");
            document.frm.Syoken.focus();
            return ;
        }
        
	if (chk_Flg == false && p_FLG != 0) {f_Button(p_FLG);return false;} //�ύX���Ȃ��ꍇ�͂��̂܂܎���

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

        document.frm.action="sei0400_upd.asp";
        document.frm.target="main";
		if( p_FLG == 1){
			document.frm.GakuseiNo.value = document.frm.txtBeforGakuNo.value;
		}
		if( p_FLG == 2){
        	document.frm.GakuseiNo.value = document.frm.txtAfterGakuNo.value;
        }
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Cansel(){

        //document.frm.action="default2.asp";
        //document.frm.target="main";
        document.frm.action="default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �O��,���փ{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Button(p_FLG){

        //document.frm.action="default.asp";
        document.frm.action="sei0400_main.asp";
        document.frm.target="main";

		if( p_FLG == 1){
			document.frm.GakuseiNo.value = document.frm.txtBeforGakuNo.value;
		}else{
        	document.frm.GakuseiNo.value = document.frm.txtAfterGakuNo.value;
        }
		document.frm.submit();
    
    }

//-->
</SCRIPT>

</head>
<body LANGUAGE=javascript onload="return window_onload()">
<form name="frm" method="post">
<center>

<br>
<table border="0" width="250">
    <tr>
<%If m_sBeforGakuNo <> "" Then%>
        <td valign="top" align="center">
            <input type="button" value="�@�O�@�ց@" class="button" onclick="javascript:f_Touroku(1)">
        </td>
<%Else%>
        <td valign="top" align="center">
            <input type="button" value="�@�O�@�ց@" class="button" DISABLED>
        </td>
<%End If%>
        <td valign="top" align="center">
            <input type="button" value="�@�o�@�^�@" class="button" onclick="javascript:f_Touroku(0)">
        </td>
        <td valign="top" align="center">
            <input type="button" value="�L�����Z��" class="button" onclick="javascript:f_Cansel()">
        </td>
<%If m_sAfterGakuNo <> "" Then%>
        <td valign="top" align="center">
            <input type="button" value="�@���@�ց@" class="button" onclick="javascript:f_Touroku(2)">
        </td>
<%Else%>
        <td valign="top" align="center">
            <input type="button" value="�@���@�ց@" class="button" DISABLED>
        </td>
<%End If%>
    </tr>
</table>
<br>
<table border="0" cellpadding="1" cellspacing="1" width="520" >
    <tr>
        <td align="left">
            <table width="500" border=1 CLASS="hyo">
                <TR>
                    <TH CLASS="header" width="120">�S�C����</TH>
                    <TD CLASS="detail"><textarea rows="4" cols="50" class="text" name="Syoken" onChange="chk_Flg=true;"><%=m_sSyoken%></textarea><br>
                    <font size="2">�i�S�p100�����ȓ��j</font></TD>
                </TR>
            </TABLE>
        </td>
    </TR>
</TABLE>

<br>

<table border="0" width="250">
    <tr>
<%If m_sBeforGakuNo <> "" Then%>
        <td valign="top" align="center">
            <input type="button" value="�@�O�@�ց@" class="button" onclick="javascript:f_Touroku(1)">
        </td>
<%Else%>
        <td valign="top" align="center">
            <input type="button" value="�@�O�@�ց@" class="button" DISABLED>
        </td>
<%End If%>
        <td valign="top" align="center">
            <input type="button" value="�@�o�@�^�@" class="button" onclick="javascript:f_Touroku(0)">
        </td>
        <td valign="top" align="center">
            <input type="button" value="�L�����Z��" class="button" onclick="javascript:f_Cansel()">
        </td>
<%If m_sAfterGakuNo <> "" Then%>
        <td valign="top" align="center">
            <input type="button" value="�@���@�ց@" class="button" onclick="javascript:f_Touroku(2)">
        </td>
<%Else%>
        <td valign="top" align="center">
            <input type="button" value="�@���@�ց@" class="button" DISABLED>
        </td>
<%End If%>
    </tr>
</table>
	<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">
	<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="txtBeforGakuNo" value="<%=m_sBeforGakuNo%>">
	<input type="hidden" name="txtAfterGakuNo" value="<%=m_sAfterGakuNo%>">
	<input type="hidden" name="GakuseiNo" value="">
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
	<input type="hidden" name="txtSikenKBN" value="<%=m_iSikenKBN%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
