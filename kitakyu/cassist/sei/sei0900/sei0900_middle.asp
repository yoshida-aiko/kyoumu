<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���i���Ґ��ѓo�^
' ��۸���ID : sei/sei0900/sei0900_middle.asp
' �@      �\: ���y�[�W ���i���Ґ��ѓo�^�̌������s��
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
' ��      ��: 2022/2/1 �g�c�@�Ď������ѓo�^��ʂ𗬗p���쐬
' ��      �X: 
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
    Public m_sKamokuNM		'�Ȗږ�			INS 2017/12/26 Nishimura
    Public m_sSikenNm		'������
    Public m_sSikenbi		'������
    Public m_sKaisiT		'�������{�J�n����
    Public m_sSyuryoT		'�������{�I������
    Public m_sKamokuNo		'�Ȗږ�
    Public m_sTKyokanCd		'�S���Ȗڂ̋���
	Dim		m_rCnt			'���R�[�h�J�E���g
    Public m_sGakkaCd
	Public m_TUKU_FLG		'�ʏ���ƃt���O
	Public m_iRisyuKakoNendo'�ߔN�x
	
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
	
	Dim m_iCount
	Dim m_sMiHyoka
	Dim m_Checked
	Dim m_Disabled
	Dim m_SchoolFlg
	
	m_Checked  = ""
	m_Disabled = ""
	
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
'response.write "middle START" & "<BR>"
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���i���Ґ��ѓo�^"
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

		'===============================
		'//�ŐV�X�V�����擾
		'===============================
		'//�N���X�̊w�N�����т��Ō�ɍX�V���������擾
		If f_GetUpdateDate(m_UpdateDate) <> 0 Then 
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
	' m_sGakuNo	= Cint(request("txtGakuNo"))
	' m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sKamokuNM	= request("txtKamokuNM")
	m_iRisyuKakoNendo = request("txtRisyuKakoNendo")
	m_sGakkaCd	= request("txtGakkaCd")
	m_TUKU_FLG	= request("txtTUKU_FLG")

	m_sGakuNo_s	= Cint(request("txtGakuNo"))
	m_sGakkaCd_s	= request("txtGakkaCd")
	m_sKamokuCd_s	= request("txtKamokuCd")
	
	m_UpdateDate = ""
	m_sFirstGakusekiNo	= request("hidFirstGakusekiNo")
	
	m_iCount = cint(request("i_Max"))
	m_sMiHyoka = request("hidMihyoka")
	m_SchoolFlg = cbool(request("hidSchoolFlg"))
	
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
	' m_bKekkaNyuryokuFlg = False

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SEISEKI_KAISI "
		w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO"
		' w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_KAISI "
		' w_sSQL = w_sSQL & vbCrLf & "  ,T24_SIKEN_NITTEI.T24_KEKKA_SYURYO "
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
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & C_SIKEN_KARISINKYU
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0'"
		w_sSQL = w_sSQL & vbCrLf & "  AND rownum <= 1 "

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
			' m_iKekkaKaishi = gf_SetNull2String(m_DRs("T24_KEKKA_KAISI"))	'���ۓ��͊J�n
			' m_iKekkaSyuryo = gf_SetNull2String(m_DRs("T24_KEKKA_SYURYO"))	'���ۓ��͏I��
			w_sSysDate = Left(gf_SetNull2String(m_DRs("SYSDATE")),10)		'�V�X�e�����t
		End If

		'���͊��ԓ��Ȃ琳��
		If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
			f_Nyuryokudate = 0
		End If

		' '���ۓ��͉\�׸�
		' If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
		' 	m_bKekkaNyuryokuFlg = True
		' End If

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

	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKUMEI AS KAMOKUMEI"
	w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "      T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(p_Gakunen) + 1
	w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_GAKKA_CD='" & p_GakkaCd & "'"
	w_sSQL = w_sSQL & vbCrLf & "  AND T15_RISYU.T15_KAMOKU_CD='" & p_KamokuCd & "'"
	

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

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(T15_LEVEL_FLG) "
		w_sSQL = w_sSQL & vbCrLf & " FROM T15_RISYU "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T15_NYUNENDO = " & cint(m_iNendo) - cint(p_Gakunen) + 1
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

'********************************************************************************
'*  [�@�\]  �ŏI�X�V���̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �ŏI�X�V��
'*  [����]  
'********************************************************************************
Function f_GetUpdateDate(p_UpdateDate)

    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetUpdateDate = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MAX(T17_KOUSINBI_KIMATU_K)"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T17_RISYUKAKO_KOJIN A,T11_GAKUSEKI B,T13_GAKU_NEN C "
		w_sSQL = w_sSQL & " WHERE"
		w_sSQL = w_sSQL & " 	A.T17_NENDO = " & Cint(m_iRisyuKakoNendo) & " "
		w_sSQL = w_sSQL & " AND	A.T17_KAMOKU_CD = '" & m_sKamokuCd & "' "
		w_sSQL = w_sSQL & " AND	A.T17_GAKUSEI_NO = B.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T17_GAKUSEI_NO = C.T13_GAKUSEI_NO "
		w_sSQL = w_sSQL & " AND	A.T17_NENDO = C.T13_NENDO "

		'//�u�����̐��k�͂͂���(C_TIKAN_KAMOKU_MOTO = 1    '�u����)
		w_sSQL = w_sSQL & " AND	A.T17_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO

		'//�K�C���I���Ȗڂ̂����I�����Ă���w���݂̂��擾����		'INS 2019/03/06 ����
		w_sSQL = w_sSQL & " AND	( T17_HISSEN_KBN = " & C_HISSEN_HIS
		w_sSQL = w_sSQL & "       OR (T17_HISSEN_KBN = " & C_HISSEN_SEN & " AND T17_SELECT_FLG = 1) "
		w_sSQL = w_sSQL & " 	) "
		w_sSQL = w_sSQL & " AND T17_HYOKA_FUKA_KBN NOT IN(" & C_HYOKA_FUKA_KEKKA &  "," & C_HYOKA_FUKA_BOTH & ") "
	
        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetUpdateDate = 99
            Exit Do
        End If
		
		' response.write "w_sSQL:" & w_sSQL & "<BR>"
		' response.end
		'//�߂�l���
		If w_Rs.EOF = False Then
			p_UpdateDate = w_Rs("MAX(T17_KOUSINBI_KIMATU_K)")
		End If
		' response.write "p_UpdateDate" & p_UpdateDate & "<BR>"

        f_GetUpdateDate = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function


'********************************************************************************
'*  [�@�\]  ���]���̐ݒ�
'********************************************************************************
Sub setHyokaType()
	
	'�Ȗڂ����]��
	if cint(gf_SetNull2Zero(m_sMiHyoka)) = cint(C_MIHYOKA) then
		m_Checked = "checked"
	end if
	
	'���͊��ԊO
	if m_iKikan = "NO" then
		m_Disabled = "disabled"
	end if
	
End Sub

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
Dim w_sInputClass

Dim w_ihalf
Dim i

i = 0

'//NN�Ή�
If session("browser") = "IE" Then
	w_sInputClass = "class='num'"
Else
	w_sInputClass = ""
End If

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
		parent.main.document.frm.action="sei0900_paste.asp";
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
					call gs_title(" ���i���Ґ��ѓo�^ "," �o�@�^ ")
				Else
					call gs_title(" ���i���Ґ��ѓo�^ "," �\�@�� ")
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
						<th class=header3 width="96"  align="center">���i�����ѓ��͊���</th><td class=detail width="239"  align="center" colspan="2"><%=m_iNKaishi%> �` <%=m_iNSyuryo%></td>
					</tr>
					<tr>

						<th class=header3 width="96"  align="center">���{�Ȗ�</th>
						<td class=detail colspan="5" align="center"><%=m_sKamokuNM%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<span class=msg2>
				<%
				'�ʏ����
				Select Case m_sSikenKBN
					Case C_SIKEN_ZEN_TYU
						%>�� �]�������N���b�N����ƁA�]���̓��͂��ł��܂��B�i�����E�̏��ŕ\������܂��j<br><%
					Case C_SIKEN_KOU_TYU
						%>�� �]�������N���b�N����ƁA�]���̓��͂��ł��܂��B�i���������E�̏��ŕ\������܂��j<br><%
					Case Else
						response.write "<BR>"
				End Select
				
				%>
				</span>
				
				<%If m_iKikan <> "NO" Then%>
					<input type=button class=button value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
				<%End If%>
				<input type=button class=button value="�L�����Z��" onclick="javascript:f_Cansel()">
				
			</td>
		</tr>
		<tr>
			<td align="center" valign="bottom" nowrap>
				<table class="hyo" border="1" align="center" width="<%= gf_IIF(m_SchoolFlg,760,710) %>">
					<tr>
						<th class="header3" colspan="14" nowrap align="center">

						</th>
					</tr>                                                                                                                                                 
					
					<tr>
						<th class="header3" rowspan="2" width="65"  nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
						<th class="header3" rowspan="2" width="150" nowrap>���@��</th>
						<th class="header3" rowspan="2" width="50"  nowrap >����</th>
						<th class="header3" rowspan="2" width="50"  nowrap>�]��</th>
						<% if m_SchoolFlg then %>
							<th class="header3" rowspan="2" width="50"  nowrap>�]��<br>�s�\</th>
						<% end if %>
						
					</tr>
					
					<tr>
					</tr>
				</table>

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