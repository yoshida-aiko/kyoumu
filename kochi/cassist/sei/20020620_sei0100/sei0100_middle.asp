<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0100_middle.asp
' �@      �\: ���y�[�W ���ѓo�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'				�R���{�{�b�N�X�͋󔒂ŕ\��
'			���\���{�^���N���b�N��
'				���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/26 �O�c �q�j
' ��      �X: 2001/08/21 �ɓ� ���q
' ��      �X: 2001/08/21 �ɓ� ���q �w�b�_���؂藣��
' ��      �X: 2002/05/02 �i   �_�l ���ʊ����̒x�����G�N�Z���\��t���Ή���
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bErrMsg           '�װү����

	'�����I��p��Where����
    Public m_iNendo			'�N�x
    Public m_sKyokanCd		'�����R�[�h
    Public m_sSikenKBN		'�����敪
    Public m_sGakuNo		'�w�N
    Public m_sClassNo		'�w��
    Public m_sKamokuCd		'�ȖڃR�[�h
    Public m_sSikenNm		'������
    Public m_sSikenbi		'������
    Public m_sKaisiT		'�������{�J�n����
    Public m_sSyuryoT		'�������{�I������
    Public m_sKamokuNo		'�Ȗږ�
    Public m_sTKyokanCd		'�S���Ȗڂ̋���
	Dim		m_rCnt			'���R�[�h�J�E���g
    Public m_sGakkaCd
	Public m_TUKU_FLG		'�ʏ���ƃt���O
	
    Public m_sGakuNo_s		'�w�N
    Public m_sGakkaCd_s		'�w��
    Public m_sKamokuCd_s	'�ȖڃR�[�h

	Public m_sGetTable			'�ȖڃR���{���쐬�����e�[�u��
	
    Public m_iKamoku_Kbn
    Public m_iHissen_Kbn

	Public	m_Rs
	Public	m_TRs
	Public	m_DRs
	Public	m_SRs
	Public	m_iMax			'�ő�y�[�W
	Public	m_iNKaishi		'���͊J�n��
	Public	m_iNSyuryo		'���͏I����
	Public	m_iKekkaKaishi		'���ȓ��͊J�n��
	Public	m_iKekkaSyuryo		'���ȓ��͏I����


	Public	m_iKikan		'���͊��ԃt���O
	Public	m_bKekkaNyuryokuFlg		'���ۓ��͉\�׸�(True:���͉� / False:���͕s��)
	
	Public m_UpdateDate
	Public m_sFirstGakusekiNo
	
	m_sKaisiT = ""
	m_sSyuryoT = "-"
	m_sSikenbi = ""

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
	w_sMsgTitle="���ѓo�^"
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

	    '// ���Ұ�SET
	    Call s_SetParam()

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'===============================
		'//���ԃf�[�^�̎擾
		'===============================
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			m_iKikan = "NO"	'���ѓ��͊��ԊO�̏ꍇ�́A�\���̂�
		End If
		
		if not f_GetUpdateDate(m_iNendo,m_sKamokuCd,m_sSikenKBN,m_TUKU_FLG,m_sFirstGakusekiNo,m_UpdateDate) then
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

Sub s_SetParam()
'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSikenKBN	= Cint(request("txtSikenKBN"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sGakkaCd	= request("txtGakkaCd")
	m_TUKU_FLG	= request("txtTUKU_FLG")

	m_sGakuNo_s	= Cint(request("txtGakuNo"))
	m_sGakkaCd_s	= request("txtGakkaCd")
	m_sKamokuCd_s	= request("txtKamokuCd")
	
	m_UpdateDate = ""
	m_sFirstGakusekiNo	= request("hidFirstGakusekiNo")
	
End Sub

'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Function f_Nyuryokudate()

	Dim w_sSysDate

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1
	m_bKekkaNyuryokuFlg = False

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI "
		w_sSQL = w_sSQL & vbCrLf & "  ,SYSDATE "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & Cint(m_sGakuNo)
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
'//���ѓ��͊��ԃe�X�g�p

'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_KAISI <= '2002/04/30'"
'		w_sSQL = w_sSQL & vbCrLf & "	AND T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO >= '2000/03/01'"

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_

'response.write w_sSQL & "<BR>"

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If m_DRs.EOF Then
			m_iNKaishi="          "
			m_iNSyuryo="          "
			Exit Do
		Else
			m_sSikenNm = gf_SetNull2String(m_DRs("M01_SYOBUNRUIMEI"))		'��������
			m_iNKaishi = gf_SetNull2String(m_DRs("T24_SEISEKI_KAISI"))		'���ѓ��͊J�n��
			m_iNSyuryo = gf_SetNull2String(m_DRs("T24_SEISEKI_SYURYO"))		'���ѓ��͏I����
			m_iKekkaKaishi = gf_SetNull2String(m_DRs("T24_KEKKA_KAISI"))	'���ۓ��͊J�n
			m_iKekkaSyuryo = gf_SetNull2String(m_DRs("T24_KEKKA_SYURYO"))	'���ۓ��͏I��
			w_sSysDate = Left(gf_SetNull2String(m_DRs("SYSDATE")),10)		'�V�X�e�����t
		End If

		'���͊��ԓ��Ȃ琳��
		If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			f_Nyuryokudate = 0
		End If

		'���ۓ��͉\�׸�
		If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			m_bKekkaNyuryokuFlg = True
		End If

		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  ���C�e�[�u�����Ȗږ��̂��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_KamokuName
'*  [����]  
'********************************************************************************
Function f_GetKamokuName(p_Gakunen,p_GakkaCd,p_KamokuCd)
	
    Dim w_sSQL
    Dim w_Rs
    Dim w_GakkaCd
    Dim w_iRet
	
    On Error Resume Next
    Err.Clear
	
    f_GetKamokuName = ""
	p_KamokuName = ""
	
    Do 

	If m_TUKU_FLG = C_TUKU_FLG_TUJO Then '�ʏ���ƂƓ��ʊ����Ŏ����ς���B
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI AS KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(p_Gakunen) + 1
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & p_GakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_KamokuCd & "'"
	Else 
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M41_MEISYO AS KAMOKUMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM M41_TOKUKATU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M41_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M41_TOKUKATU_CD='" & p_KamokuCd & "'"
	End If

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		If w_Rs.EOF = False Then
			p_KamokuName = w_Rs("KAMOKUMEI")
		End If

        Exit Do
    Loop

    f_GetKamokuName = p_KamokuName

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  ���x���ʂ��ǂ����𒲂ׂ�B
'*  [����]  �Ȃ�
'*  [�ߒl]  ���x���ʁFtrue
'*  [����]  
'********************************************************************************
Function f_LevelChk(p_Gakunen,p_KamokuCd)

    Dim w_sSQL
    Dim w_Rs
    Dim w_GakkaCd
    Dim w_iRet

    On Error Resume Next
    Err.Clear

    f_LevelChk = false
	p_KamokuName = ""
    Do 

		'//�����s���̂Ƃ�
		If trim(p_Gakunen)="" Or  trim(p_KamokuCd) = "" Then
            Exit Do
		End If

		'//�w��CD���擾
'		w_iRet = f_GetGakkaCd(p_Gakunen,p_Class,w_GakkaCd)
		If w_iRet<> 0 Then
            Exit Do
		End If

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(T15_LEVEL_FLG) "
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_NYUNENDO = " & cint(m_iNendo) - cint(p_Gakunen) + 1
'		w_sSQL = w_sSQL & vbCrLf & "  AND T15_GAKKA_CD='" & w_GakkaCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "  AND T15_KAMOKU_CD = '" & p_KamokuCd & "'"


        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

		If w_Rs.EOF = False and cint(w_Rs("MAX(T15_LEVEL_FLG)")) = 1 Then
			f_LevelChk = true
		End If

        Exit Do
    Loop
    Call gf_closeObject(w_Rs)
End Function

'*******************************************************************************
' �@�@�@�\�F�I�񂾎����敪,�Ȗڂ̍X�V�����擾����
' 
' �ԁ@�@�l�F
' �@�@�@�@�@(True)����, (False)���s
' ���@�@���Fp_sSikenKbn			- �����敪
' �@�@�@�@�@p_Nendo				- �N�x
' �@�@�@�@�@p_KamokuCd			- �ȖڃR�[�h
'			p_FirstGakusekiNo	- �w��NO
'			p_TUKU_FLG			- �Ȗڋ敪(0:�ʏ�A1:����)
'			(�߂�l)p_UpdateDate - �Ȗڎ��ѓo�^�̍X�V��
' �@�\�ڍׁF
' ���@�@�l�F
' old_ver �� gf_GetT16UpdDate(m_iNendo,m_sGakuNo_s,m_sGakkaCd_s,m_sKamokuCd_s,"")
'*******************************************************************************
function f_GetUpdateDate(p_Nendo,p_KamokuCd,p_ShikenKbn,p_TUKU_FLG,p_FirstGakusekiNo,p_UpdateDate)
	
	Dim w_Sql,w_Rs
	Dim w_ShikenType
	Dim w_Table
	Dim w_TableName
	Dim w_KamokuName
	
	On Error Resume Next
	Err.Clear
	
	f_GetUpdateDate = false
	
	if cint(p_TUKU_FLG) = cint(C_TUKU_FLG_TUJO) then
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	else
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_KamokuName = "T34_TOKUKATU_CD"
	end if
	
	select case cint(p_ShikenKbn)
		
		case C_SIKEN_ZEN_TYU '�O�����Ԏ���
			w_ShikenType = w_Table & "_KOUSINBI_TYUKAN_Z"
			
		case C_SIKEN_ZEN_KIM '�O����������
			w_ShikenType = w_Table & "_KOUSINBI_KIMATU_Z"
			
		case C_SIKEN_KOU_TYU '������Ԏ���
			w_ShikenType = w_Table & "_KOUSINBI_TYUKAN_K"
			
		case C_SIKEN_KOU_KIM '�����������
			w_ShikenType = w_Table & "_KOUSINBI_KIMATU_K"
			
		case else
			exit function
			
	end select
	
	w_Sql = ""
	w_Sql = w_Sql & " select "
	w_Sql = w_Sql & " 		" & w_ShikenType
	w_Sql = w_Sql & " from "
	w_Sql = w_Sql & " 		" & w_TableName
	w_Sql = w_Sql & " where "
	w_Sql = w_Sql & " 		" & w_Table & "_GAKUSEKI_NO= '"   & p_FirstGakusekiNo   & "' "
	w_Sql = w_Sql & " and	" & w_Table & "_NENDO = " & p_Nendo
	w_Sql = w_Sql & " and	" & w_KamokuName & "= '"   & p_KamokuCd   & "' "
	
	If gf_GetRecordset(w_Rs,w_Sql) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit function
	End If
	
	p_UpdateDate = gf_YYYY_MM_DD(w_Rs(0),"/")
	
	f_GetUpdateDate = true
	
end function


Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_sGakusekiCd
Dim w_sSeiseki
Dim w_sHyoka
Dim w_sKekka
Dim w_sChikai
Dim w_sKekkasu
Dim w_sChikaisu

Dim w_ihalf
Dim i

i = 0

%>
<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

		//�X�N���[����������
		parent.init();

    }

   //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){
        parent.main.f_Touroku();
    }

	//************************************************************
	//	[�@�\]	�L�����Z���{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//************************************************************
	function f_Cansel(){

        //�����y�[�W��\��
        parent.document.location.href="default.asp";
	
	}

	//************************************************************
	//	[�@�\]	�y�[�X�g�{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//************************************************************
	function f_Paste(pType){
		
		parent.main.document.frm.PasteType.value=pType;
		
		//submit�ŉ�ʂ��J���ƃE�B���h�E�̃X�e�[�^�X���ݒ�ł��Ȃ����ߤ
		//��U��y�[�W���J���Ă���A�V�E�B���h�E�ɑ΂���submit����B
		nWin=open("","Paste","location=no,menubar=no,resizable=yes,scrollbars=no,scrolling=no,status=no,toolbar=no,width=300,height=600,top=0,left=0");
		parent.main.document.frm.target="Paste";
		parent.main.document.frm.action="sei0100_paste.asp";
		parent.main.document.frm.submit();
	
	}
	//-->
	</SCRIPT>
	</head>
    <body onload="return window_onload()">
	<table border="0" cellpadding="0" cellspacing="0" height="245" width="100%">
		<tr>
			<td>
				<%
				If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then
					call gs_title(" ���ѓo�^ "," �o�@�^ ")
				Else
					call gs_title(" ���ѓo�^ "," �\�@�� ")
				End If
				%>
			</td>
		</tr>
		<tr>
			<td align="center" nowrap><form name="frm" method="post">
				<table border=1 class=hyo width=670>
					<tr>
						<th class="header3" colspan="6" nowrap align="center">
						���ѓ��͊��ԁ@<%=m_sSikenNm%>�@�@�@�X�V���F<%=m_UpdateDate%>
						</th>
					</tr>
					<tr>
						<th class=header3 width="96"  align="center">���ѓ��͊���</th><td class=detail width="239"  align="center" colspan="2"><%=m_iNKaishi%> �` <%=m_iNSyuryo%></td>
						<th class=header3 width="96"  align="center">���ۓ��͊���</th><td class=detail width="239"  align="center" colspan="2"><%=m_iKekkaKaishi%> �` <%=m_iKekkaSyuryo%></td>
					</tr>
					<tr>
						<th class=header3 width="96"  align="center">���{�Ȗ�</th>
						<%
							If f_LevelChk(m_sGakuNo,m_sKamokuCd) = true then 
								w_str = m_sGakuNo & "�N�@" & gf_GetClassName(m_iNendo,m_sGakuNo,m_sClassNo) & "�@" & f_GetKamokuName(m_sGakuNo,m_sGakkaCd,m_sKamokuCd)
							Else
								w_str = m_sGakuNo & "�N�@" & gf_GetClassName(m_iNendo,m_sGakuNo,m_sClassNo) & "�@" & f_GetKamokuName(m_sGakuNo,m_sGakkaCd,m_sKamokuCd)
							End If
						%>
						<td class=detail colspan="5" align="center"><%=w_str%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">

				<span class=msg2>
				���u���X�v�v�́A���Əo�����̓��j���[�����X���͂��ꂽ��L�����܂ł̏o���󋵂ł��B<br>
				���u�ΏۊO�v�́A�����Ȃǂ̗݌v����͂��Ă��������B<br>
				���w�b�_�̕����F���u<FONT COLOR="#99CCFF">����</FONT>�v�̂悤�ɂȂ��Ă��镔�����N���b�N����ƁAExcel�\��t���p�̉�ʂ��J���܂��B<br>
				<%
				'�ʏ���ƂƓ��ʊ����ŕ\����ς���
				If m_TUKU_FLG = C_TUKU_FLG_TUJO Then
					Select Case m_sSikenKBN
						Case C_SIKEN_ZEN_TYU
							%>�� �]�������N���b�N����ƁA�]���̓��͂��ł��܂��B�i�����E�̏��ŕ\������܂��j<br><%
						Case C_SIKEN_KOU_TYU
							%>�� �]�������N���b�N����ƁA�]���̓��͂��ł��܂��B�i���������E�̏��ŕ\������܂��j<br><%
						Case Else
							response.write "<BR>"
					End Select
				End If
				%>
				</spanc>

				<% If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then %>
					<input type=button class=button value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
				<%End If%>
				<input type=button class=button value="�L�����Z��" onclick="javascript:f_Cansel()">

			</td>
		</tr>
		<tr>
			<td align="center" valign="bottom" nowrap>
				
				<!--�ʏ���ƂƓ��ʊ����ŕ\����ς���B-->
				<% If m_TUKU_FLG = C_TUKU_FLG_TUJO Then %>
					
					<table class="hyo" border="1" align="center" width="710">
						<tr><th class="header3" colspan="13" nowrap align="center">
								�����Ɛ�&nbsp;<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%><input type="text" class="NUM" maxlength="3" style="width:30px" name="txtSouJyugyou" value="<%= Request("hidSouJyugyou") %>"><% Else %><%= Request("hidSouJyugyou") %><% End if%>�@
								�����Ɛ�&nbsp;<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%><input type="text" class="NUM" maxlength="3" style="width:30px" name="txtJunJyugyou" value="<%= Request("hidJunJyugyou") %>"><% Else %><%= Request("hidJunJyugyou") %><% End if%>�@
							</th></tr>                                                                                                                                                 
						<tr>
							<th class="header3" rowspan="2" width="65" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
							<th class="header3" rowspan="2" width="150" nowrap>���@��</th>
							<th class="header3" colspan="4" width="120" nowrap>���ї���</th>
							<th class="header3" rowspan="2" width="50" nowrap onClick="f_Paste('Seiseki')"><FONT COLOR="#99CCFF">����</FONT></th>
							<th class="header3" rowspan="2" width="50" nowrap>�]��</th>
							<th class="header3" colspan="2" width="110" nowrap>�x��</th>
							<th class="header3" colspan="3" width="165" nowrap">����</th>
						</tr>
						
						<tr>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">�O��</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">�O��</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">�㒆</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">�w��</span></th>
							<th class="header2" width="55" nowrap onClick="f_Paste('Chikai')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">����</FONT></span></th>
							<th class="header2" width="55" nowrap><span style="font-size:10px;">���X�v</span></th>
							<th class="header2" width="55" nowrap onClick="f_Paste('Kekka')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">�Ώ�</FONT></span></th>
							<th class="header2" width="55" nowrap onClick="f_Paste('KekkaGai')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">�ΏۊO</FONT></span></th>
							<th class="header2" width="55" nowrap><span style="font-size:10px;">���X�v</span></th>
						</tr>
					</table>
				<% else %>
					<table class="hyo" border=1 align="center" width="710" nowrap>
						<tr>
							<th class="header3" colspan="13" nowrap align="center">
								�����Ɛ�&nbsp;<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%><input type="text" class="NUM" maxlength="5" style="width:30px" name="txtSouJyugyou" value="<%= Request("hidSouJyugyou") %>"><% Else %><%= Request("hidSouJyugyou") %><% End if%>�@
								�����Ɛ�&nbsp;<%If m_iKikan <> "NO" or m_bKekkaNyuryokuFlg Then%><input type="text" class="NUM" maxlength="5" style="width:30px" name="txtJunJyugyou" value="<%= Request("hidJunJyugyou") %>"><% Else %><%= Request("hidJunJyugyou") %><% End if%>�@
							</th>
						</tr>
						<tr>
							<th class="header3" rowspan="2" width="65" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
							<th class="header3" rowspan="2" width="150" nowrap>���@��</th>
							<th class="header3" colspan="4" width="120" nowrap>���ї���</th>
							<th class="header3" rowspan="2" width="50" nowrap onClick="f_Paste('Seiseki')"><FONT COLOR="#99CCFF">����</FONT></th>
							<th class="header3" rowspan="2" width="50" nowrap>�]��</th>
							<th class="header3" rowspan="2" width="100" nowrap onClick="f_Paste('Chikai')"><FONT COLOR="#99CCFF">�x��</FONT></th>
							<th class="header3" colspan="2" width="165" nowrap>����</th>
						</tr>
						<tr>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">�O��</FONT></span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">�O��</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">�㒆</span></th>
							<th class="header2" width="30" nowrap><span style="font-size:10px;">�w��</span></th>
							<th class="header2" width="80" nowrap onClick="f_Paste('Kekka')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">�Ώ�</FONT></span></th>
							<th class="header2" width="85" nowrap onClick="f_Paste('KekkaGai')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">�ΏۊO</FONT></span></th>
						</tr>
					</table>
				<% end if %>
			</td>
		</tr>
	</table>

	</body>
	</html>
<%
End sub

Sub No_showPage()
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
	<SCRIPT language="javascript">
	<!--

	    //************************************************************
	    //  [�@�\]  �y�[�W���[�h������
	    //  [����]
	    //  [�ߒl]
	    //  [����]
	    //************************************************************
	    function window_onload() {

	        //submit
			parent.location.href = "white.asp?txtMsg=���ѓ��͊��ԊO�ł��B"
	        return;
	    }

	//-->
	</SCRIPT>
	</head>

    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">

	<input type="hidden" name="txtMsg" value="���ѓ��͊��ԊO�ł��B">

	</form>
	</body>
	</html>

<%
End Sub

Sub showPage_No()
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
	<SCRIPT language="javascript">
	<!--

	    //************************************************************
	    //  [�@�\]  �y�[�W���[�h������
	    //  [����]
	    //  [�ߒl]
	    //  [����]
	    //************************************************************
	    function window_onload() {

	        //submit
			parent.location.href = "white.asp?txtMsg=�l���C�f�[�^�����݂��܂���B"
	        return;
	    }

	//-->
	</SCRIPT>
	</head>

    <body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" method="post">
	</head>

	<body>
	<br><br><br>
	<center>
		<span class="msg">�l���C�f�[�^�����݂��܂���B</span>
	</center>

	<input type="hidden" name="txtMsg" value="�f�[�^�����݂��܂���B">

	</form>
	</body>
	</html>

<%
End Sub
%>