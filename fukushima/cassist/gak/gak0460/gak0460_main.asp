<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������������o�^
' ��۸���ID : gak/gak0460/gak0460_main.asp
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
' ��      �X�F2002/10/08 �A�c �k��Y   �S�C�����A���i���̍��ڂ�ǉ�
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
    Public m_sTanninSyoken  '�S�C����
    Public m_sTanninBikou   '���i��
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
	
    Public  m_GRs
    Public  m_Rs
    Public  m_iMax          '�ő�y�[�W
    Public  m_iDsp          '�ꗗ�\���s��
	
	Public m_sNendo         '�N�x�R���{�{�b�N�X�ɓ���l
	Public m_sGakkoNO       '�w�Z�ԍ�
	
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
        If gf_OpenDatabase() <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '// ���Ұ�SET
        Call s_SetParam()

		Call f_Gakusei()
		
        '//�f�[�^�擾
        If f_getdate() <> 0 Then m_bErrFlg = True : Exit Do
        
        '//�w�Ȃb�c�擾
        If f_getGakka() <> 0 Then m_bErrFlg = True : Exit Do

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
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
    m_iNendo    = cint(session("NENDO"))
    m_sKyokanCd = session("KYOKAN_CD")
    m_sGakuNo   = request("txtGakuNo")
    m_iDsp      = C_PAGE_LINE
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm  = request("txtClassNm")
	
	m_sNendo    = request("txtNendo")
	
	'//�O��OR���փ{�^���������ꂽ��
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

'********************************************************************************
'*  [�@�\]  �����̎������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_Gakusei()

Dim i
i = 1

    w_iNyuNendo = Cint(m_sNendo) - Cint(m_sGakunen) + 1
    'w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

	'//�w���̏����W
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & " T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
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
    w_sSQL = w_sSQL & "     B.T13_NENDO = " & m_sNendo & " "
    'w_sSQL = w_sSQL & "     B.T13_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & " AND B.T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND B.T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & " ORDER BY B.T13_GAKUSEKI_NO "
	
	If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then
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

'********************************************************************************
'*  [�@�\]  �f�[�^�̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_getdate()

    On Error Resume Next
    Err.Clear
    f_getdate = 1
	
	if Not gf_GetGakkoNO(m_sGakkoNO) then
        m_bErrFlg = True
		exit function
	end if
	
    Do
		
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT  "
        w_sSQL = w_sSQL & "     T11_SOGOSYOKEN,T11_KOJIN_BIK,T11_SINRO,T11_SOTUKEN_DAI, "
        w_sSQL = w_sSQL & "     T11_SOTU_KYOKAN_CD1,T11_SOTU_KYOKAN_CD2,T11_SOTU_KYOKAN_CD3 "
        
        if m_sGakkoNO = cstr(C_NCT_KUMAMOTO) then
	        w_sSQL = w_sSQL & "    ,T13_TANNINSYOKEN "
        	w_sSQL = w_sSQL & "    ,T13_TANNIN_BIK"
        end if
        
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T11_GAKUSEKI, "
		w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
		
		if m_sGakkoNO = cstr(C_NCT_KUMAMOTO) then
			w_sSQL = w_sSQL & "     AND T13_NENDO = " & m_sNendo
			w_sSQL = w_sSQL & "     AND T13_GAKUSEI_NO = T11_GAKUSEI_NO "
		end if
		
		Set m_Rs = Server.CreateObject("ADODB.Recordset")
        
        If gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp) <> 0 Then
            'ں��޾�Ă̎擾���s
            f_getdate = 99
            m_bErrFlg = True
            Exit Do 
        End If

        m_sSsyoken  = m_Rs("T11_SOGOSYOKEN")
        m_sBikou    = m_Rs("T11_KOJIN_BIK")
        m_sSinro    = m_Rs("T11_SINRO")
        m_sSotudai  = m_Rs("T11_SOTUKEN_DAI")
        m_sSkyokan1 = m_Rs("T11_SOTU_KYOKAN_CD1")
        m_sSkyokan2 = m_Rs("T11_SOTU_KYOKAN_CD2")
        m_sSkyokan3 = m_Rs("T11_SOTU_KYOKAN_CD3")

		if m_sGakkoNO = cstr(C_NCT_KUMAMOTO) then
			m_sTanninSyoken = m_Rs("T13_TANNINSYOKEN")
	        m_sTanninBikou  = m_Rs("T13_TANNIN_BIK")
		end if

        f_getdate = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �w���̏����w�Ȃ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_getGakka()

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
        document.frm.action="gak0460_topDisp.asp";
        document.frm.submit();

	}

    //************************************************************
    //  [�@�\]  �i�H��I����ʃE�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function NewWin(p_iFLG,p_sSNm) {

		var obj=eval("document.frm."+p_sSNm)
        URL = "../../mst/mst0133/default.asp?txtFLG="+p_iFLG+"&txtSNm="+escape(obj.value)+"";
        //URL = "../../mst/mst0133/default.asp?txtFLG="+p_iFLG+"&txtSNm="+p_sSNm+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,toolbar=no,width=560,height=600,top=0,left=0");
        nWin.focus();
        return true;    
    }

    //************************************************************
    //  [�@�\] �N���A�{�^���������ꂽ�Ƃ�
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function jf_Clear(pTextName,pHiddenName){
        eval("document.frm."+pTextName).value = "";
        eval("document.frm."+pHiddenName).value = "";
    }

    //************************************************************
    //  [�@�\]  ���������Q�ƑI����ʃE�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function KyokanWin(p_iInt,p_sKNm) {
		
		var obj=eval("document.frm."+p_sKNm)
        URL = "../../Common/com_select/SEL_KYOKAN/default.asp";
        URL = URL + "?txtI="+p_iInt;
        URL = URL + "&txtKNm="+escape(obj.value);
        URL = URL + "&txtGakka=<%=m_sGakka%>";
        //URL = URL + "&hidNendo=<%=m_sNendo%>";
        
        //URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+p_sKNm+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=yes,scrollbars=yes,status=no,toolbar=no,width=550,height=650,top=0,left=0");
        nWin.focus();
        return true;    
    }
    
    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(p_FLG){
	
	if (chk_Flg == false && p_FLG != 0) {f_Button(p_FLG);return false;} //�ύX���Ȃ��ꍇ�͂��̂܂܎���

        // ���������������̌�����������
        if( getLengthB(document.frm.SGSyoken.value) > "200" ){
            window.alert("���������̗��͑S�p100�����ȓ��œ��͂��Ă�������");
            document.frm.SGSyoken.focus();
            return ;
        }
        // �������l���l�̌�����������
        if( getLengthB(document.frm.Bikou.value) > "80" ){
            window.alert("�l���l�̗��͑S�p40�����ȓ��œ��͂��Ă�������");
            document.frm.Bikou.focus();
            return ;
        }
<%If m_sGakunen = 5 Then%>
        // �����������_��̌�����������
        if( getLengthB(document.frm.SRondai.value) > "80" ){
            window.alert("�����_��̗��͑S�p40�����ȓ��œ��͂��Ă�������");
            document.frm.SRondai.focus();
            return ;
        }
<%End If%>
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

        document.frm.action="gak0460_upd.asp";
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
        document.frm.action="gak0460_main.asp";
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
<form name="frm" method="post" onClick="return false;">
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
				<% if m_sGakkoNO = cstr(C_NCT_KUMAMOTO) then %>
	                <TR>
	                    <TH CLASS="header" width="120">�S�C����</TH>
	                    <TD CLASS="detail"><textarea rows="4" cols="50" class="text" name="TanninSyoken" onChange="chk_Flg=true;"><%=m_sTanninSyoken%></textarea><br>
	                    <font size="2">�i�S�p100�����ȓ��j</font></TD>
	                </TR>
	                <TR>
	                    <TH CLASS="header" width="120">���i��</TH>
	                    <TD CLASS="detail"><textarea rows="2" cols="50" class="text" name="TanninBikou" onChange="chk_Flg=true;"><%=m_sTanninBikou%></textarea><br>
	                    <font size="2">�i�S�p40�����ȓ��j</font></TD>
	                </TR>
				<% end if %>

                <TR>
                    <TH CLASS="header" width="120">��������</TH>
                    <TD CLASS="detail"><textarea rows="4" cols="50" class="text" name="SGSyoken" onChange="chk_Flg=true;"><%=m_sSsyoken%></textarea><br>
                    <font size="2">�i�S�p100�����ȓ��j</font></TD>
                </TR>
                <TR>
                    <TH CLASS="header" width="120">���@�l</TH>
                    <TD CLASS="detail"><textarea rows="2" cols="50" class="text" name="Bikou" onChange="chk_Flg=true;"><%=m_sBikou%></textarea><br>
                    <font size="2">�i�S�p40�����ȓ��j</font></TD>
                </TR>
<%If m_sGakunen = 5 Then%>
                <!--TR>
                    <TH CLASS="header" width="120">���ƌ�̐i�H</TH>
                    <TD CLASS="detail">
                    <input type="text" class="text" name="SinroNm" VALUE='<%Call f_SINRO(m_sSinro,m_iNendo)%>' size="50" readonly style="width:260px;" onChange="chk_Flg=true;">
                    <input type="hidden" name="SinroCd" VALUE='<%=m_sSinro%>'>
                    <input type="button" class="button" value="�I��" onclick="NewWin(1,'SinroNm')">
                    <input type="button" class="button" value="�N���A" onclick="jf_Clear('SinroNm','SinroCd')">
                </TR-->
                <TR>
                    <TH CLASS="header" width="120">�����_��</TH>
                    <TD CLASS="detail"><textarea rows="2" cols="50" class="text" name="SRondai" onChange="chk_Flg=true;"><%=m_sSotudai%></textarea><br>
                    <font size="2">�i�S�p40�����ȓ��j</font></TD>
                </TR>
                <TR>
                    <TH CLASS="header" nowrap width="120">��������1</TH>
                    <TD CLASS="detail">
                    <!--input type="text" class="text" name="SKyokanNm1" VALUE='<%Call f_KYO_MEI(m_sSkyokan1,m_iNendo)%>' size="24" readonly onChange="chk_Flg=true;"-->
                    <input type="text" class="text" name="SKyokanNm1" VALUE='<%Call f_KYO_MEI(m_sSkyokan1,m_sNendo)%>' size="24" readonly onChange="chk_Flg=true;">
                    <input type="hidden" name="SKyokanCd1" VALUE='<%=m_sSkyokan1%>'>
                    <input type="button" class="button" value="�I��" onclick="KyokanWin(1,'SKyokanNm1')">
                    <input type="button" class="button" value="�N���A" onclick="jf_Clear('SKyokanNm1','SKyokanCd1')"></td>
                </TR>
                <TR>
                    <TH CLASS="header" nowrap width="120">��������2</TH>
                    <TD CLASS="detail">
                    <!--input type="text" class="text" name="SKyokanNm2" VALUE='<%Call f_KYO_MEI(m_sSkyokan2,m_iNendo)%>' size="24" readonly onChange="chk_Flg=true;"-->
                    <input type="text" class="text" name="SKyokanNm2" VALUE='<%Call f_KYO_MEI(m_sSkyokan2,m_sNendo)%>' size="24" readonly onChange="chk_Flg=true;">
                    <input type="hidden" name="SKyokanCd2" VALUE='<%=m_sSkyokan2%>'>
                    <input type="button" class="button" value="�I��" onclick="KyokanWin(2,'SKyokanNm2')">
                    <input type="button" class="button" value="�N���A" onclick="jf_Clear('SKyokanNm2','SKyokanCd2')"></td>
                </TR>
                <TR>
                    <TH CLASS="header" nowrap width="120">��������3</TH>
                    <TD CLASS="detail">
                    <!--input type="text" class="text" name="SKyokanNm3" VALUE='<%Call f_KYO_MEI(m_sSkyokan3,m_iNendo)%>' size="24" readonly onChange="chk_Flg=true;"-->
                    <input type="text" class="text" name="SKyokanNm3" VALUE='<%Call f_KYO_MEI(m_sSkyokan3,m_sNendo)%>' size="24" readonly onChange="chk_Flg=true;">
                    <input type="hidden" name="SKyokanCd3" VALUE='<%=m_sSkyokan3%>'>
                    <input type="button" class="button" value="�I��" onclick="KyokanWin(3,'SKyokanNm3')">
                    <input type="button" class="button" value="�N���A" onclick="jf_Clear('SKyokanNm3','SKyokanCd3')"></td>
                </TR>
<%End If%>
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
	<input type="hidden" name="txtNendo" value="<%=m_sNendo%>">
	<!--input type="hidden" name="txtNendo" value="<%=m_iNendo%>"-->
	<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">
	<input type="hidden" name="txtGakunen" value="<%=m_sGakunen%>">
	<input type="hidden" name="txtBeforGakuNo" value="<%=m_sBeforGakuNo%>">
	<input type="hidden" name="txtAfterGakuNo" value="<%=m_sAfterGakuNo%>">
	<input type="hidden" name="GakuseiNo" value="">
	<input type="hidden" name="txtClass" value="<%=m_sClass%>">
	<input type="hidden" name="txtClassNm" value="<%=m_sClassNm%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
