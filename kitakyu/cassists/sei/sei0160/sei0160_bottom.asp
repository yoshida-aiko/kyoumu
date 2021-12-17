<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ������w���ѓo�^
' ��۸���ID : sei/sei0100/sei0160_bottom.asp
' �@      �\: ���y�[�W ������w���ѓo�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���
'           :�N�x		��		SESSION���
' ��      ��:�Ȃ�
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2007/04/11 ��c
' ��      �X: 2008/09/30 �����@�k��B����̏ꍇ�@���ۂ̕]���͏o�͂��Ȃ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Dim m_bErrFlg				'//�װ�׸�
    
    Const C_ERR_GETDATA = "�f�[�^�̎擾�Ɏ��s���܂���"
    
    Dim m_iNendo				'//�N�x
    Dim m_sKyokanCd				'//�����R�[�h
    Dim m_sGakunen				'//�w�N
    Dim m_sClass				'//�N���X

    Dim m_sBunruiCD		 		'//���ރR�[�h
    Dim m_sBunruiNM		 		'//���ޖ���
    Dim m_sTani		 			'//�P��

    Dim m_lDataCount,m_uData()			'//�]���f�[�^
    Dim m_rCnt					'//���R�[�h�J�E���g
    Dim m_Rs
	
    Dim m_iSeisekiInpType
	Public m_sGakkoNO       '�w�Z�ԍ�  INS 2008/09/30
	Public m_iMaxTem        '���ۍő�_��  INS 2008/09/30

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
	Dim w_iRet
	Dim w_sSQL
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
	'Message�p�̕ϐ��̏�����
	w_sWinTitle = "�L�����p�X�A�V�X�g"
	w_sMsgTitle = "������w���ѓo�^"
	w_sMsg = ""
	w_sRetURL = C_RetURL & C_ERR_RETURL
	w_sTarget = ""
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = false
	
	Do
		'//�ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//�s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'//���Ұ�SET
		Call s_SetParam()

		'//���ѓ��͕��@�̎擾(0:�_��[C_SEISEKI_INP_TYPE_NUM]�A1:����[C_SEISEKI_INP_TYPE_STRING]�A2:���ہA�x��[C_SEISEKI_INP_TYPE_KEKKA])
		if not gf_GetKamokuSeisekiInp(m_iNendo,m_sBunruiCd,C_KAMOKUBUNRUI_NINTEI,m_iSeisekiInpType) then 
			m_bErrFlg = True
			Exit Do
		end if

		'//�w�Z�ԍ��̎擾
		if Not gf_GetGakkoNO(m_sGakkoNO) then
			m_bErrFlg = True
			Exit Do
		end if
		
		'//���сA�w���f�[�^�擾
		If not f_GetStudent() Then m_bErrFlg = True : Exit Do
		
		If m_Rs.EOF Then
			Call gs_showWhitePage("�l���C�f�[�^�����݂��܂���B","������w���ѓo�^")
			Exit Do
		End If

		'//�]���f�[�^�̎擾
		if not f_GetKamokuHyokaData(m_iNendo,m_sBunruiCd,C_KAMOKUBUNRUI_NINTEI,m_lDataCount,m_uData) then
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
		
		if w_sMsg = "" then w_sMsg = C_ERR_GETDATA
		
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// �I������
	Call gf_closeObject(m_Rs)
	
	Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()
	
    	m_iNendo    	= session("NENDO")		'�N�x
    	m_sKyokanCd 	= session("KYOKAN_CD")		'�����R�[�h

	m_sGakunen  	= request("txtGakunen")         '�w�N
	m_sClass    	= request("txtClass")           '�N���X
	m_sBunruiCD 	= request("txtBunruiCd")	'���ރR�[�h
	m_sBunruiNm 	= request("txtBunruiNm")	'���ޖ���
	m_sTani     	= request("txtTani")		'�P��

End Sub

'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'********************************************************************************
Function f_GetStudent()
	
	Dim w_sSQL
	
	On Error Resume Next
	Err.Clear
	
	f_GetStudent = false

	'//�������ʂ̒l���ꗗ��\��
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " T13_GAKUSEI_NO  AS GAKUSEI_NO,"
	w_sSQL = w_sSQL & " T13_GAKUSEKI_NO AS GAKUSEKI_NO,"
	w_sSQL = w_sSQL & " T13_GAKKA_CD    AS GAKKA_CD,"
	w_sSQL = w_sSQL & " T13_CLASS    �@ AS CLASS,"
	w_sSQL = w_sSQL & " T13_COURCE_CD   AS COURCE_CD,"

	w_sSQL = w_sSQL & " T11_SIMEI       AS SIMEI, "

    	w_sSQL = w_sSQL & " T100_HAITOTANI  AS HAITOTANI,"
    	w_sSQL = w_sSQL & " T100_HYOTEI     AS HYOTEI,"
    	w_sSQL = w_sSQL & " T100_HYOKA      AS HYOKA,"
    	w_sSQL = w_sSQL & " T100_NINTEIBI   AS NINTEIBI,"
    	w_sSQL = w_sSQL & " T100_SYUTOKU_NENDO AS SYUTOKU_NENDO,"
    	w_sSQL = w_sSQL & " T100_SEISEKI    AS SEISEKI,"
    	w_sSQL = w_sSQL & " T100_HYOKA_FUKA_KBN AS FUKA_KBN"

	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI,"
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN, "
	w_sSQL = w_sSQL & " 	T100_RISYU_NINTEI "
	
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "     T13_NENDO    = " & Cint(m_iNendo)
	w_sSQL = w_sSQL & " AND	T13_GAKUNEN  = " & Cint(m_sGakunen)
	w_sSQL = w_sSQL & " AND	T13_CLASS    = " & Cint(m_sClass)
    	w_sSQL = w_sSQL & " AND T13_GAKUSEI_NO = T11_GAKUSEI_NO"
    	w_sSQL = w_sSQL & " AND T13_GAKUSEI_NO = T100_GAKUSEI_NO(+)"
    	w_sSQL = w_sSQL & " AND T100_BUNRUI_CD(+) = '" & m_sBunruiCd & "'"

    
    	w_sSQL = w_sSQL & " ORDER BY"
    	w_sSQL = w_sSQL & " T13_GAKUSEKI_NO"

	
'response.write w_sSQL
'response.end


	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit function
	
	'//ں��ރJ�E���g�擾
	m_rCnt = gf_GetRsCount(m_Rs)
	
	f_GetStudent = true
	
End Function

'********************************************************************************
'*  [�@�\]  �e�[�u���T�C�Y�̃Z�b�g
'********************************************************************************
Sub s_SetTableWidth(p_TableWidth)
	
	p_TableWidth = 610
	
End Sub

'*******************************************************************************
' �@�@�@�\�F�Ȗڕ]�����X�g�擾(������w��p)
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_sKamokuCD - �ȖڃR�[�h(IN)
'               �i�F��Ȗڂ̏ꍇ�͕��ރR�[�h���w�肷��j
'           p_sKamokuBunrui - �Ȗڕ��ރR�[�h(IN)
'               C_KAMOKUBUNRUI_TUJYO = �ʏ�Ȗ�
'               C_KAMOKUBUNRUI_NINTEI = �F��Ȗ�
'               C_KAMOKUBUNRUI_TOKUBETU = ���ʉȖ�
'           p_lDataCount -  �]���f�[�^����(OUT)
'           p_uData - �]���f�[�^(OUT)
'
' �@�\�ڍׁF�w��̉ȖڃR�[�h�Ɠ_������p_uData()�ɕ]���A�]��A���_�Ȗڂ�ݒ肷��
'           p_uData()�͓��I�z���call���Ő錾���邱�ƁB�i�錾�͊֐����ōs���j
'           p_uData()�̌�����p_lDataCount�ɃZ�b�g�����B�܂��A�z��C���f�b�N�X��
'           1 �` p_lDataCount�܂ł��L���B
' ���@�@�l�F�_���]���̉ȖڑΏ�
'           call��
'           ret = gf_GetKamokuHyokaData(m_iNendo, w_KamokuCD, C_KAMOKUBUNRUI_TUJYO, w_lConut, w_udata())
'
'           gf_GetKamokuHyokaData�@���쐬�Af_GetHyokaData��Call����悤�ɕύX
'*******************************************************************************
Function f_GetKamokuHyokaData(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_lDataCount,p_uData)
    Dim w_iZokuseiCD         '�Ȗڑ���
    Dim w_iHyokaNo
    
    On Error Resume Next
    
    f_GetKamokuHyokaData = False
    
    '�Ȗڑ����擾
    If Not gf_GetKamokuZokusei(p_iNendo, p_sKamokuCD, p_sKamokuBunrui, w_iZokuseiCD) Then
        Exit Function
    End If
    '�Ȗڑ�������]��NO�擾
    w_iHyokaNo = gf_iGetHyokaNo(w_iZokuseiCD, p_iNendo)
    
    '�]��NO����]���f�[�^�擾
    If Not f_GetHyokaData(p_iNendo, w_iHyokaNo, p_lDataCount, p_uData) Then
        Exit Function
    End If
	
    '�]��NO����]���f�[�^�擾�@INS 2008/09/30����
	IF m_sGakkoNO = cstr(C_NCT_KITAKYU) then
		If Not f_GetKekkaMaxTen(p_iNendo, w_iHyokaNo) Then
	        Exit Function
	    End If
	end if
	

    f_GetKamokuHyokaData = True
             
End Function


'*******************************************************************************
' �@�@�@�\�F�Ȗڕ]���擾(������w��p)
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_iHyokaNo - �]��NO(IN)
'           p_lDataCount - ����(OUT)
'           p_uData - �]���f�[�^(OUT)
'
' �@�\�ڍׁF�_������]��NO��p_uData�ɕ]���A�]��A���_�Ȗڂ�ݒ肷��
' ���@�@�l�F�]��NO�����łɕ������Ă���ꍇ�ɂ͒���call����
'           �]��NO��������Ȃ��Ƃ��́Agf_GetKamokuTensuHyoka��call
'           f_GetHyokaData�@���쐬�AM08_HYOKA_TAISYO_KBN�@��C_HYOKA_TAISHO_HOKA�ɕύX
'*******************************************************************************
Function f_GetHyokaData(p_iNendo,p_iHyokaNo,p_lDataCount,p_uData)
    Dim w_oRecord
    Dim w_sSql
    Dim w_lIdx
    
    On Error Resume Next
    
    f_GetHyokaData = False
    
    p_lDataCount = 0
    
    w_sSql = ""
    w_sSql = w_sSql & " SELECT"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_MEI,"
    w_sSql = w_sSql & " 	M08_HYOTEI,"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_RYAKU"
    
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & " 	M08_HYOKAKEISIKI"
    
    w_sSql = w_sSql & " WHERE"
    w_sSql = w_sSql & " 	M08_HYOUKA_NO = " & p_iHyokaNo      		'�]��NO
    w_sSql = w_sSql & " AND M08_NENDO = " & p_iNendo        			'�N�x
    w_sSql = w_sSql & " AND M08_HYOKA_TAISYO_KBN = " & C_HYOKA_TAISHO_HOKA     	'����w

	' INS 2008/09/30 ���� ���_�L���̓R���{�ɃZ�b�g���Ȃ�
	IF m_sGakkoNO = cstr(C_NCT_KITAKYU) then
		w_sSql = w_sSql & " AND M08_HYOKA_SYOBUNRUI_RYAKU = 0 "
	END IF
    
    w_sSql = w_sSql & " ORDER BY"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_CD"
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '�Ȗ�M�Ȃ����G���[
    If w_oRecord.EOF Then
        Exit Function
    End If
    
    p_lDataCount = gf_GetRsCount(w_oRecord)
    
    '�z��f�[�^�錾
    ReDim p_uData(p_lDataCount,3)
    w_lIdx = 0
    
    Do Until w_oRecord.EOF
        
        '�f�[�^�Z�b�g
        p_uData(w_lIdx,0) = w_oRecord("M08_HYOKA_SYOBUNRUI_MEI")	'�]��
        p_uData(w_lIdx,1) = w_oRecord("M08_HYOTEI")			'�]��
        p_uData(w_lIdx,2) = w_oRecord("M08_HYOKA_SYOBUNRUI_RYAKU")	
        
        w_lIdx = w_lIdx + 1
        w_oRecord.MoveNext
    Loop
    
    Call gf_closeObject(w_oRecord)
    
    f_GetHyokaData = True

End Function

'*******************************************************************************
' �@�@�@�\�F���ۉȖڂ�Max�_��(������w��p)
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_iHyokaNo - �]��NO(IN)
'
' �@�\�ڍׁF���_�Ȗڂ̍ő�_�����擾
' ���@�@�l�F
'*******************************************************************************
Function f_GetKekkaMaxTen(p_iNendo,p_iHyokaNo)

    Dim w_oRecord
    Dim w_sSql
    Dim w_lIdx
    
    On Error Resume Next
    
    f_GetKekkaMaxTen = False
    
	m_iMaxTem = 0   
    w_sSql = ""
    w_sSql = w_sSql & " SELECT"
    w_sSql = w_sSql & " 	MAX(M08_MAX) MAX_TEN"
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & " 	M08_HYOKAKEISIKI"
    w_sSql = w_sSql & " WHERE"
    w_sSql = w_sSql & " 	M08_HYOUKA_NO = " & p_iHyokaNo      		'�]��NO
    w_sSql = w_sSql & " AND M08_NENDO = " & p_iNendo        			'�N�x
    w_sSql = w_sSql & " AND M08_HYOKA_TAISYO_KBN = " & C_HYOKA_TAISHO_HOKA     	'����w
    w_sSql = w_sSql & " AND M08_HYOKA_SYOBUNRUI_RYAKU = 1 "
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '�Ȗ�M�Ȃ����G���[
    If w_oRecord.EOF Then
        Exit Function
    End If
    
    
    '�f�[�^�Z�b�g
    m_iMaxTem = w_oRecord("MAX_TEN")
    
    Call gf_closeObject(w_oRecord)

f_GetKekkaMaxTen = true

End Function


'********************************************************************************
'*  [�@�\]  HTML���o��
'********************************************************************************
Sub showPage()
	DIm w_cell
	DIm w_Padding
	DIm w_Padding2

	Dim i
	Dim ii

	DIm w_sValue	''�R���{VALUE����

	Dim w_sInputClass
	
	Dim w_Disabled
	Dim w_Disabled2
	Dim w_TableWidth

	'�����ݒ�
	w_Padding = "style='padding:2px 0px;'"
	w_Padding2 = "style='padding:2px 0px;font-size:10px;'"
	
	i = 1

	'//NN�Ή�
	If session("browser") = "IE" Then
		w_sInputClass  = "class='num'"
	Else
		w_sInputClass = ""
	End If
	
	'//�e�[�u���T�C�Y�̃Z�b�g
	Call s_SetTableWidth(w_TableWidth)
	
%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--
	//************************************************************
	//  [�@�\]  �y�[�W���[�h������
	//************************************************************
	function window_onload() {
//		//�X�N���[����������
//		parent.init();
//		
//		document.frm.target = "topFrame";
//		document.frm.action = "sei0160_top.asp";
//		document.frm.submit();
	}
	//************************************************************
	//  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
	//************************************************************
	function f_Touroku(){
		if(!f_InpCheck()){
		//	alert("���͒l���s���ł�");
			return false;
		}

		if(!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return false;}
		
		//�w�b�_���󔒕\��
		parent.topFrame.document.location.href="white.asp";
		
		//�o�^����
		document.frm.action="sei0160_upd.asp";
		document.frm.target="main";
		document.frm.submit();
	}
	
	//************************************************************
	//	[�@�\]	�L�����Z���{�^���������ꂽ�Ƃ�
	//************************************************************
	function f_Cancel(){
		parent.document.location.href="default.asp";
	}

	//************************************************
	//	���̓`�F�b�N
	//************************************************
	function f_InpCheck(){
		var w_length;
		var ob;
		
		w_length = document.frm.elements.length;
		
		for(i=0;i<w_length;i++){
			ob = eval("document.frm.elements[" + i + "]")
			
			if(ob.type=="text"){
				ob = eval("document.frm." + ob.name);
				
				if(!f_CheckNum(ob)){
						alert("���͒l���s���ł�");
						return false;}
			}


		}


    	<% IF m_sGakkoNO = cstr(C_NCT_KITAKYU) then %>
		for(i=1;i < <%=m_rCnt%> + 1 ;i++){

				ob = eval("document.frm.Seiseki" + i);
				if(ob.value.length == 0){
				
				}
				else
				{
				if(!f_CheckKekka(ob)){
					
					alert(<%=m_iMaxTem%> + "�_�ȉ��͓��͂ł��܂���");
					return false;	
				}
				}
		}
		<% END IF%>


		return true;
	}

	//************************************************************
	//  [�@�\]  ���l�^�`�F�b�N
	//************************************************************
	function f_CheckNum(pFromName){
		var wFromName,w_len;
		
		wFromName = eval(pFromName);
		
		if(isNaN(wFromName.value)){
			wFromName.focus();
			wFromName.select();
			return false;
		}else{
			//���`�F�b�N
			if(wFromName.name.indexOf("Seiseki") != -1){
				if(wFromName.value > 100){
					wFromName.focus();
					wFromName.select();
					return false;
				}
			}

			//�C���N�x�́A4���܂�
			if(wFromName.name.indexOf("SyuNendo") != -1){
				w_len = 4;
			}else{
				w_len = 3;
			}
			
			if(wFromName.value.length > w_len){
				wFromName.focus();
				wFromName.select();
				return false;
			}

			//�}�C�i�X���`�F�b�N
			var wStr = new String(wFromName.value)
			if (wStr.match("-")!=null){
				wFromName.focus();
				wFromName.select();
				return false;
			}
		}

		return true;
	}

	//************************************************************
	//  [�@�\]  ���ۓ_�`�F�b�N INS2008/09/30����
	//************************************************************
	function f_CheckKekka(pFromName){
		var wFromName,w_len;

		var wFromName,w_len;
		
		wFromName = eval(pFromName);
		
		if(isNaN(wFromName.value)){
			wFromName.focus();
			wFromName.select();
			return false;
		}else{

			if(wFromName.value <= <%=m_iMaxTem%>){
				wFromName.focus();
				wFromName.select();
				return false;
			}
		}
		return true;
	}

	
	//************************************************
	//Enter �L�[�ŉ��̓��̓t�H�[���ɓ����悤�ɂȂ�
	//�����Fp_inpNm	�Ώۓ��̓t�H�[����
	//    �Fp_frm	�Ώۃt�H�[��
	//�@�@�Fi		���݂̔ԍ�
	//�ߒl�F�Ȃ�
	//���̓t�H�[�������Axxxx1,xxxx2,xxxx3,�c,xxxxn 
	//�̖��O�̂Ƃ��ɗ��p�ł��܂��B
	//************************************************
	function f_MoveCur(p_inpNm,p_frm,i){
		if (event.keyCode == 13){		//�����ꂽ�L�[��Enter(13)�̎��ɓ����B
			i++;
			
			//���͉\�̃e�L�X�g�{�b�N�X��T���B����������t�H�[�J�X���ڂ��ď����𔲂���B
	        for (w_li = 1; w_li <= 99; w_li++) {
				
				if (i > <%=m_rCnt%>) i = 1; //i���ő�l�𒴂���ƁA�͂��߂ɖ߂�B
				inpForm = eval("p_frm."+p_inpNm+i);
				
				//���͉\�̈�Ȃ�t�H�[�J�X���ڂ��B
				if (typeof(inpForm) != "undefined") {
					inpForm.focus();			//�t�H�[�J�X���ڂ��B
					inpForm.select();			//�ڂ����e�L�X�g�{�b�N�X����I����Ԃɂ���B
					break;
				//���͕t���Ȃ玟�̍��ڂ�
				} else{
					i++
				}
	        }
		}else{
			return false;
		}
		return true;
	}

	//************************************************
	//	�]���R���{���͎��̏���
	//	
	//************************************************
	function f_ChgHyoka(w_num){
		var ob = new Array();
		ob[0] = eval("document.frm.sltHyoka" + w_num);

		ob[1] = eval("document.frm.hidHyoka" + w_num);
		ob[2] = eval("document.frm.hidHyotei" + w_num);
		ob[3] = eval("document.frm.hidHyokaFukaKbn" + w_num);
		ob[4] = eval("document.frm.SyuNendo" + w_num);

		if(ob[0].value.length == 0||ob[0].value =="@@@"){
			ob[1].value = "";
			ob[2].value = "";
			ob[3].value = "";
			ob[4].value = "";
		}else{
			var vl = ob[0].value.split('#@#');
			
			ob[1].value = vl[0];
			ob[2].value = vl[1];
			ob[3].value = vl[2];
			
			//���i�ŏC���N�x�������͂̂Ƃ������N�x��\������
			if(ob[3].value=="0"){
				if(ob[4].value==""){
					ob[4].value ="<%=m_iNendo%>";
				}
			}else{
				ob[4].value ="";
			}
		}
	}

	//************************************************
	//	�����N�x���͎��̏���
	//	
	//************************************************
	function f_ChkSyoriNendo(w_num){
		var ob = new Array();
		ob[0] = eval("document.frm.SyuNendo" + w_num);
		ob[1] = eval("document.frm.hidHyokaFukaKbn" + w_num);

		if(ob[1].value=="0"){
			//�]���s�敪��"��" �̂Ƃ������N�x�͂�\������
			if(ob[0].value==""){
				ob[0].value ="<%=m_iNendo%>";
			}
		}else{
			//�]���s�敪��"��" �ȊO�̂Ƃ������N�x�͓��͂ł��Ȃ�
			ob[0].value ="";
		}

	}
	
	//-->
	</SCRIPT>
	</head>
	<body LANGUAGE="javascript" onload="window_onload();">
	<center>
	<form name="frm" method="post">
	
	<table width="<%=w_TableWidth%>">
	<tr>
	<td>
	
	<table class="hyo" align="center" width="<%=w_TableWidth%>" border="1">
		<tr>
			<th class="header" width="80" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
			<th class="header" width="300"���@��</th>
			<th class="header" width="50" nowrap>����</th>
			<th class="header" width="50" nowrap>�]��</th>
			<th class="header" width="100" nowrap>�C���N�x</th>
		</tr>
	<%
	m_Rs.MoveFirst

	i = 0
	
	Do Until m_Rs.EOF

		i = i + 1
		
		Call gs_cellPtn(w_cell)
	%>
			
		<tr>
			<td class="<%=w_cell%>" align="center" width="80"  nowrap <%=w_Padding%>><%=m_Rs("GAKUSEKI_NO")%></td>
			<input type="hidden" name="txtGsekiNo<%=i%>"   value="<%=m_Rs("GAKUSEKI_NO")%>">
			<input type="hidden" name="txtGseiNo<%=i%>"    value="<%=m_Rs("GAKUSEI_NO")%>">
			<input type="hidden" name="txtGakkaCD<%=i%>"   value="<%=m_Rs("GAKKA_CD")%>">
			<input type="hidden" name="txtClass<%=i%>"     value="<%=m_Rs("CLASS")%>">
			<input type="hidden" name="txtCorceCD<%=i%>"   value="<%=m_Rs("COURCE_CD")%>">
			<td class="<%=w_cell%>" align="left" width="300" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%></td>
				

			<!-- ����  -->
			<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>
                        	<input type="text" <%=w_sInputClass%>  name="Seiseki<%=i%>" value="<%=trim(m_Rs("SEISEKI"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);">
                        </td>

				
			<!-- �]�� -->
			<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>
				<select name="sltHyoka<%=i%>" onchange ="f_ChgHyoka(<%=i%>);">
					<Option Value="@@@">

		<% 	For ii = 0 to  m_lDataCount-1 

				'�R���{VALUE��������
				w_sValue = ""
				w_sValue = w_sValue & m_uData(ii,0) & "#@#"
				w_sValue = w_sValue & m_uData(ii,1) & "#@#"
				w_sValue = w_sValue & m_uData(ii,2) 
		%>
				<% if trim(m_Rs("HYOKA")) = gf_SetNull2String(m_uData(ii,0)) then %>
					<option value="<%=w_sValue%>" selected> <%=m_uData(ii,0)%></option>
				<% Else %>
					<option value="<%=w_sValue%>"> <%=m_uData(ii,0)%></option>
				<% end if 
		    	NEXT 
		%>
				</select>
			</td>

			<!-- �擾�N�x -->
			<td class="<%=w_cell%>" align="center" width="100" nowrap <%=w_Padding%>>
	                	<input type="text" <%=w_sInputClass%> name="SyuNendo<%=i%>" value="<%=trim(m_Rs("SYUTOKU_NENDO"))%>" size=4 maxlength=4 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);"onBlur ="f_ChkSyoriNendo(<%=i%>);">
                        </td>

			<input type="hidden" name="hidHyoka<%=i%>"           value="<%=m_Rs("HYOKA")%>">
			<input type="hidden" name="hidHyotei<%=i%>"          value="<%=m_Rs("HYOTEI")%>">
			<input type="hidden" name="hidHyokaFukaKbn<%=i%>"    value="<%=m_Rs("FUKA_KBN")%>">
		</tr>
	<%
		m_Rs.MoveNext
	Loop
	%>
			
			
	</table>
	
	</td>
	</tr>
	
	<tr>
	<td align="center">
	<table>
		<tr>
			<td align="center" align="center" colspan="13">
					<input type="button" class="button" value="�@�o�@�^�@" onClick="f_Touroku();">�@
					<input type="button" class="button" value="�L�����Z��" onClick="f_Cancel();">
			</td>
		</tr>
	</table>
	</td>
	</tr>

	</table>
	
	
	<input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">

	<input type="hidden" name="txtGakunen"   value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass"     value="<%=m_sClass%>">
	<input type="hidden" name="txtBunruiCd"  value="<%=m_sBunruiCD%>">
	<input type="hidden" name="txtBunruiNm"  value="<%=m_sBunruiNM%>">
	<input type="hidden" name="txtTani"      value="<%=m_sTani%>">


	<input type="hidden" name="i_Max"        value="<%=i%>">
	
	</form>
	</center>
	</body>
	</html>
<%
End sub
%>