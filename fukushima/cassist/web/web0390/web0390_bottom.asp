<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���x���ʉȖڌ���
' ��۸���ID : web/web0390/web0390_main.asp
' �@      �\: ���y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/10/26 �J�e�@�ǖ�
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp          '// �ꗗ�\���s��
    Public  m_sPageCD       ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_Grs           '�w���p���R�[�h�Z�b�g
    Public  m_KSrs          '�Ȗڐ��̃��R�[�h�Z�b�g
    Public  m_GrsCntMax     '�w���p���R�[�h��
'    Public  m_rs            '���R�[�h�Z�b�g
    Dim     m_iNendo        '//�N�x
    Dim     m_sKyokanCd     '//�����R�[�h
    Dim     m_sGakunen      '//�w�N
    Dim     m_sClass        '//�N���X
	Dim		m_sKamokuCD		'//�ȖڃR�[�h
    Dim     m_GrCnt         '//�w���̃��R�[�h�J�E���g
    Dim     m_cell          '�z�F�̐ݒ�
	Dim 	m_sKengen		'//����
    Dim     m_iSTani        
	Dim		m_sRisyuJotai	'���C��ԃt���O add 2001/10/25
	Dim 	m_sLKyokan()	'�I�����ꂽ���x���ʉȖڂ̒S������
	Dim 	m_iLKyokanCnt()	'�S��������I��ł���l�̐�
    Dim     i               
    Dim     j               

    '�G���[�n
    Public  m_bErrFlg       '�װ�׸�
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

    Dim	 w_iRet              '// �߂�l
    Dim  w_Krs           '�Ȗڗp���R�[�h�Z�b�g
    Dim  w_KrCnt         '//�Ȗڂ̃��R�[�h�J�E���g

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���x���ʉȖړo�^"
    w_sMsg=""
    w_sRetURL=C_RetURL & C_ERR_RETURL
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
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
        '// ���Ұ�SET
        Call s_SetParam()


		'//�������擾
		w_iRet = gf_GetKengen_web0390(m_sKengen)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'�����̒萔
		'C_WEB0390_ACCESS_FULL  
		'C_WEB0390_ACCESS_SENMON
		'C_WEB0390_ACCESS_TANNIN

		'���C��ԋ敪���擾(���C�����肵�Ă邩�ǂ����j
		'C_K_RIS_MAE = 0        '�m�菈���O
		'C_K_RIS_ATO = 1        '�m�菈����
		if f_GetKanriM(m_iNendo,C_K_RIS_JOUTAI,m_sRisyuJotai) <> 0 then 
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
	        m_bErrFlg = True
	        Call w_sMsg("�Ǘ��}�X�^�̗��C��ԋ敪������܂���B")
	        Exit Do
		end if

'-----------------------------------------------------
'm_sRisyuJotai = "1" 'test�p
'-----------------------------------------------------

        '//�����̏��擾
        w_iRet = f_KyokanData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

		If Ubound(m_sLKyokan) = 0 Then
			Call showPage_NoData()
	        Exit Do
		End If

        '//�w���̏��擾
        w_iRet = f_GakuseiData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

        '// �y�[�W��\��
        Call showPage()

        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Grs)
    '// �I������
    Call gs_CloseDatabase()
    
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
response.write w_sMsg
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sGakunen  = request("txtGakunen")
    m_sClass    = request("txtClass")
    m_iDsp      = C_PAGE_LINE
	m_sKamokuCD = request("cboKamokuCode")

End Sub

Function f_KyokanData()
'******************************************************************
'�@�@�@�\�F�����̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear
    f_KyokanData = 1
	i = 0
	m_bErrFlg = false

    Do

        '//�Ȗڂ̃f�[�^�擾
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT "
        m_sSQL = m_sSQL & vbCrLf & "     T27_KYOKAN_CD"
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T27_TANTO_KYOKAN T27"
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T27_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T27_GAKUNEN = " & m_sGakunen & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T27_KAMOKU_CD = '" & m_sKamokuCD & "' "
        m_sSQL = m_sSQL & vbCrLf & " GROUP BY T27_KYOKAN_CD"
        m_sSQL = m_sSQL & vbCrLf & " ORDER BY T27_KYOKAN_CD"

'response.write m_sSQL & "<BR>"
'response.end
        Set w_Krs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset_OpenStatic(w_Krs, m_sSQL)

'response.write "w_iRet = " & w_iRet & "<BR>"
'response.write w_Krs.EOF & "<BR>"

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
			w_sMsg = "�����f�[�^���擾�ł��܂���B"
            Exit Do 
        End If

    w_KrCnt=cint(gf_GetRsCount(w_Krs))
'response.write w_KrCnt

	Redim m_sLKyokan(w_KrCnt)
	Redim m_iLKyokanCnt(w_KrCnt)

	w_Krs.MoveFirst
	Do Until w_Krs.EOF
		m_sLKyokan(i) = w_Krs("T27_KYOKAN_CD")
		m_iLKyokanCnt(i) = 0

	if m_sKengen = C_WEB0390_ACCESS_SENMON then
		If m_sKyokanCd = w_Krs("T27_KYOKAN_CD") then
			m_sMain = i
		End If
	End If

		i = i + 1
		w_Krs.MoveNext
	Loop
'response.end
    f_KyokanData = 0

    Exit Do

    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(w_Krs)

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If

End Function

Function f_GakuseiData()
'******************************************************************
'�@�@�@�\�F�w���̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
	Dim w_sSQL

    On Error Resume Next
    Err.Clear
    f_GakuseiData = 1

    Do
        '//�w���̃f�[�^�擾
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_LEVEL_KYOUKAN AS L_KYOKAN, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO AS GAKUSEI,"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO AS GAKUSEKI,"
        w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI AS SIMEI"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN,T13_GAKU_NEN,T11_GAKUSEKI"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_NENDO = "          & cInt(m_iNendo) 		& " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_HAITOGAKUNEN = "   & cInt(m_sGakunen)  	& " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_KAMOKU_CD = '"     & m_sKamokuCd       	& "' AND "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLASS = "     	    &  m_sClass       		& " AND"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_ZAISEKI_KBN < "   	& C_ZAI_SOTUGYO		  	& " AND"
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_GAKUSEI_NO = T13_GAKU_NEN.T13_GAKUSEI_NO AND"
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_HAITOGAKUNEN = T13_GAKU_NEN.T13_GAKUNEN AND"
        w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_NENDO = T13_GAKU_NEN.T13_NENDO AND"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO "
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI "

'response.write m_sSQL & "<BR>"

        Set m_Grs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset_OpenStatic(m_Grs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_GrCnt=gf_GetRsCount(m_Grs)

    f_GakuseiData = 0

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
'    If m_bErrFlg = True Then
'        w_sMsg = gf_GetErrMsg()
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
'		response.end
'    End If

End Function

'********************************************************************************
'*  [�@�\]  �Ǘ��}�X�^���f�[�^���擾
'*  [����]  p_iNendo	�N�x
'*  �@�@�@  p_iNo		�����ԍ�
'*  [�ߒl]  p_iKanri	�Ǘ��f�[�^
'*  [����]  �Ǘ��}�X�^���f�[�^���擾����B
'********************************************************************************
Function f_GetKanriM(p_iNendo,p_iNo,p_sKanri)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKanriM = 0
    p_sKanri = ""

    Do 

		'//�Ǘ��}�X�^��藚�C��ԋ敪���擾
		'//���C��ԋ敪(C_K_RIS_JOUTAI = 28)
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_NENDO=" & cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M00_KANRI.M00_NO=" & C_K_RIS_JOUTAI	'���C��ԋ敪(=28)

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            f_GetKanriM = iRet
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			'//Public Const C_K_RIS_MAE = 0    '����O
			'//Public Const C_K_RIS_ATO = 1    '�����
			p_sKanri = w_Rs("M00_KANRI")
		End If

        f_GetKanriM = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

Sub showPage_NoData()
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
	<SCRIPT language="javascript">
	<!--
    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		parent.location.href = "white.asp?txtMsg=���x���ʉȖڂ̃f�[�^������܂���B"
        return;
    }
	//-->
	</SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>
    </center>
	<form name="frm" method="post">

	<input type="hidden" name="txtMsg" value="���x���ʉȖڂ̃f�[�^������܂���B">

	</form>
    </body>
    </html>

<%
    '---------- HTML END   ----------
End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim n
Dim w_sChg

    On Error Resume Next
    Err.Clear

i = 0
n = 0
%>
<HTML>


<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>���x���ʉȖڌ���</title>

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

        //�w�b�_��submit
        document.frm.target = "middle";
        document.frm.action = "web0390_middle.asp";
        document.frm.submit();
        return;

    }

    //************************************************************
    //  [�@�\]  �{�^����VALUE�̕ύX
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Chenge(p_iS,p_iK){
		var w_sBtn;
		var w_sTmp;
		var w_sBNm ;		//�����ꂽ�{�^���̗�̖��O
		var w_sKYOCntNm;	//middle�̑I���w�����̃t�H�[���̖��O
		var w_sKYOCnt;		//middle�̑I���w�����̃t�H�[��
		var w_sKYONm; 		//�I�����������̋����b�c
		var w_sLKYONm;		//�w���̑I�����������̋����b�c�����Ƃ���
		var w_sLKYO_OLDNm;	//�w���̑I�����Ă��������̋����b�c�����Ƃ���
		
        w_sBNm = "document.frm.K"+p_iS+"_";
		w_sLKYO_OLDNm = eval("document.frm.L_KYOKAN_OLD"+p_iS);
        w_sKYOCntNm = "parent.middle.document.frm.KYOKAN";
		w_sKYONm = eval("document.frm.KyokanuCd"+p_iK);
		w_sLKYONm = eval("document.frm.L_KYOKAN"+p_iS);
		//������
		w_sCnt = <%=UBound(m_sLKyokan)%>;
		w_sBtn = eval(w_sBNm+p_iK);

		//���܂őI�����Ă������̂̃J�E���g�����炷(�I�����Ă����ꍇ�j
		if (w_sLKYO_OLDNm.value != 999) {
			w_sKYOCnt = eval(w_sKYOCntNm + w_sLKYO_OLDNm.value);
			w_sKYOCnt.value = parseInt(w_sKYOCnt.value) - 1;
		}
		//�I���������̂��������Ƃ�
		if (w_sBtn.value == "��") {
			w_sBtn.value = "�@";
			w_sLKYONm.value = "";
			eval(w_sLKYO_OLDNm).value = 999;
		} else {

		//�I��������
			//��U�A�S�Ắ����폜
			for ( i=0;i<=w_sCnt-1;i++) {

				w_sTmp = eval(w_sBNm+i)
				w_sTmp.value = "�@";
			}

			//�I���������̂Ɋۂ�����
			w_sBtn.value = "��";

			//���A�I�����Ă������̂̃J�E���g�𑝂₷
			w_sKYOCnt = eval(w_sKYOCntNm + p_iK);
			w_sKYOCnt.value = parseInt(w_sKYOCnt.value) + 1;

			//�I�����������̃R�[�h�������
			eval(w_sLKYO_OLDNm).value = p_iK;
			w_sLKYONm.value = w_sKYONm.value;
		}
        return;
    }
    
    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Cansel(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default2.asp";

    
    }
    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        document.frm.action="web0390_upd.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>

	<center>

	<body onload="return window_onload()">
	<FORM NAME="frm" method="post">

	    <%
		'//�B���t�B�[���h�ɉȖ�CD�Ɗe�Ȗڂ̒P�ʐ����i�[(�o�^���Ɏg�p����)

        Do Until n > UBound(m_sLKyokan)
		    %>
	        <input type=hidden name=KyokanuCd<%=n%> value="<%=m_sLKyokan(n)%>">
		    <%
	        n = n + 1
        Loop%>
	<table class=hyo border=1>

	    <%
	        m_Grs.MoveFirst
	        Do Until m_Grs.EOF
	            Call gs_cellPtn(m_cell)
		        i = i + 1
		        j = 0
				w_iChkNo = 999
			    %>
			    <tr>
			        <td class=<%=m_cell%> width="50"><%=m_Grs("GAKUSEKI")%>
			        <input type=hidden name=gakuNo<%=i%> value="<%=m_Grs("GAKUSEI")%>"></td>
			        <td class=<%=m_cell%> width="120"><%=m_Grs("SIMEI")%>
			        <input type=hidden name=gakuNm<%=i%> value="<%=m_Grs("SIMEI")%>">
			        <input type=hidden name=L_KYOKAN<%=i%> value="<%=m_Grs("L_KYOKAN")%>"></td>
			    <%

				For n = 0 to UBound(m_sLKyokan)-1

					'�Ή����鋳����I�����Ă���΁A��������
					If m_sLKyokan(n) = m_Grs("L_KYOKAN") then
						w_sChk = "��" 
						w_iChkNo = n 
						m_iLKyokanCnt(n) = m_iLKyokanCnt(n) + 1
					else 
						w_sChk = "�@"
					End If

			If cint(m_sRisyuJotai) = C_K_RIS_ATO then 
				'�m�菈����----------------------------------------------------- 
				w_sChg = ""
				
			Else
				'�m�菈���O----------------------------------------------------- 
				If m_sKengen <> C_WEB0390_ACCESS_SENMON then
					'�������S�������̂݃��[�h�łȂ�----------------------------------------------------- 
					w_sChg = "onclick='f_Chenge(""" & i & """,""" & n & """)'"
				Else
					'�������S�������̂݃��[�h----------------------------------------------------- 
					'm_sLKyokan(i)�Ƌ����b�c����v(�ύX�ł���)----------------------------------------------------- 
					If m_sLKyokan(n) = m_sKyokanCd then 
						w_sChg = "onclick='f_Chenge(""" & i & """,""" & n & """)'"
					Else 
						w_sChg = ""
					End If
				End If
			End If

			%>
			        <td class=<%=m_cell%>   width="90">
			        <input type="button" class="<%=m_cell%>" name="K<%=i%>_<%=n%>" value="<%=w_sChk%>" <%=w_sChg%> style="text-align:center" >
					</td>
			<% 
		        Next
				%>
			        <input type=hidden name=L_KYOKAN_OLD<%=i%> value="<%=w_iChkNo%>">
				    </tr>
				<%
				m_Grs.MoveNext
	        Loop%>
	</table>
	<% If cint(m_sRisyuJotai) = C_K_RIS_MAE then %>
	<table>
	    <tr>
	        <td align=center><input type=button class=button value="�@�o�@�^�@" onclick="javascript:f_Touroku()"></td>
	        <td align=center><input type=button class=button value="�L�����Z��" onclick="javascript:f_Cansel()"></td>
	    </tr>
	</table>
	<% End If %>

	<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="cboKamokuCode"      value="<%=m_sKamokuCD%>">

	<input type="hidden" name="txtGakunen"  value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass"    value="<%=m_sClass%>">
	<input type="hidden" name="txtRisyu"      value="<%=m_sRisyuJotai%>">
	<input type="hidden" name="txtGakuMax"      value="<%=m_GrCnt%>">

<% '������I�������w�������B���Ď�����
	For n=0 to UBound(m_sLKyokan)
%>
	<input type="hidden" name="txtLKCnt<%=n%>"      value="<%=m_iLKyokanCnt(n)%>">
<%
    Next
%>

	</FORM>
	</center>
	</BODY>
	</HTML>
<%
End Sub
%>