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
' ��      �X�F2001/08/30 �ɓ� ���q     ����������2�d�ɕ\�����Ȃ��悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�����I��p��Where����
    Public m_iNendo         '�N�x
    Public m_sKyokanCd      '�����R�[�h
    Public m_sNendo         '�N�x�R���{�{�b�N�X�ɓ���l
    Public m_sGakuNo        '�����R���{�{�b�N�X�ɓ���l
    Public m_sBeforGakuNo   '�����R���{�{�b�N�X�ɓ���l�̈�l�O
    Public m_sAfterGakuNo   '�����R���{�{�b�N�X�ɓ���l�̈�l��
    Public m_sGakunen       '�w�N
    Public m_sClass         '�N���X
    Public m_sClassNm       '�N���X��
    Public m_sGakusei()     '�w���̔z��

    Public  m_TRs           
    Public  m_GRs           
    Public  m_URs
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

        '//�f�[�^�擾
        w_iRet = f_Gakusei()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		If m_GRs.EOF Then
			Call NO_Showpage()
			Exit Do
		End If

        '//�f�[�^�擾
        w_iRet = f_getdate()
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

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_sNendo    = request("txtNendo")
    m_sGakuNo   = request("txtGakuNo")
	m_sGakunen  = Cint(request("txtGakunen"))
	m_sClass    = Cint(request("txtClass"))
	m_sClassNm  = request("txtClassNm")
    m_iDsp      = C_PAGE_LINE

	'//�O��OR���փ{�^���������ꂽ��
	If Request("GakuseiNo") <> "" Then
	    m_sGakuNo   = Request("GakuseiNo")
	End If

End Sub

Function f_Gakusei()
'********************************************************************************
'*  [�@�\]  �����̎������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim i
i = 1


    w_iNyuNendo = Cint(m_sNendo) - Cint(m_sGakunen) + 1

	'//�w���̏����W
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT  "
    w_sSQL = w_sSQL & "     T11_SIMEI "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T11_GAKUSEKI "
    w_sSQL = w_sSQL & " WHERE"
    w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "
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
    w_sSQL = w_sSQL & " AND B.T13_GAKUNEN = " & m_sGakunen & " "
    w_sSQL = w_sSQL & " AND B.T13_CLASS = " & m_sClass & " "
    w_sSQL = w_sSQL & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
'    w_sSQL = w_sSQL & " AND A.T11_NYUNENDO = B.T13_NENDO - B.T13_GAKUNEN + 1"
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

Function f_getdate()
'********************************************************************************
'*  [�@�\]  �f�[�^�̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_getdate = 1

    Do
        '//�s������,�����Z��擾���i��,�l���l�̃f�[�^�擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
'        w_sSQL = w_sSQL & "     T11_KODOSYOKEN,T11_SYUMITOKUGI,T11_TYOSA_BIK "
        w_sSQL = w_sSQL & "     T11_TYOSA_BIK "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T11_GAKUSEKI "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T11_GAKUSEI_NO = '" & m_sGakuNo & "' "

        Set m_TRs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_TRs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            f_getdate = 99
            m_bErrFlg = True
            Exit Do 
        End If

        '//���ʊ����̋L�^,�N�������̃f�[�^�擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     T13_TOKUKATU_DET,T13_NENSYOKEN,T13_NENSYOKEN2,T13_NENSYOKEN3 "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T13_GAKU_NEN "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     T13_GAKUSEI_NO = '" & m_sGakuNo & "' "
        w_sSQL = w_sSQL & " AND T13_NENDO = " & m_sNendo & " "
        Set m_URs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_URs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            f_getdate = 99
            m_bErrFlg = True
            Exit Do 
        End If

        f_getdate = 0
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

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
        w_sSQL = w_sSQL & "     T13_NENDO = " & m_sNendo
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
    <title>�������������o�^</title>
<link rel=stylesheet href="../../common/style.css" type=text/css>

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
        document.frm.action="gak0461_topDisp.asp";
        document.frm.submit();

	}

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(p_FLG){


        // �������s�������̌�����������
//        if( getLengthB(document.frm.KSyoken.value) > "100" ){
//            window.alert("�s�������̗��͑S�p50�����ȓ��œ��͂��Ă�������");
//            document.frm.KSyoken.focus();
//            return ;
//        }
//        // ����������̌�����������
//        if( getLengthB(document.frm.SyumiTokugi.value) > "200" ){
//            window.alert("����̗��͑S�p100�����ȓ��œ��͂��Ă�������");
//            document.frm.SyumiTokugi.focus();
//            return ;
//        }
        // �������w����Q�l�ƂȂ鏔�����̌�����������
        if( getLengthB(document.frm.NSyoken.value) > "100" ){
            window.alert("�w����Q�l�ƂȂ鏔�����̗��͑S�p50�����ȓ��œ��͂��Ă�������");
            document.frm.NSyoken.focus();
            return ;
        }
        // �������w����Q�l�ƂȂ鏔�����̌�����������
        if( getLengthB(document.frm.NSyoken2.value) > "100" ){
            window.alert("�w����Q�l�ƂȂ鏔�����̗��͑S�p50�����ȓ��œ��͂��Ă�������");
            document.frm.NSyoken2.focus();
            return ;
        }
        // �������w����Q�l�ƂȂ鏔�����̌�����������
        if( getLengthB(document.frm.NSyoken3.value) > "100" ){
            window.alert("�w����Q�l�ƂȂ鏔�����̗��͑S�p50�����ȓ��œ��͂��Ă�������");
            document.frm.NSyoken3.focus();
            return ;
        }
        // ���������ʊ����̌�����������
        if( getLengthB(document.frm.Tokukatu.value) > "100" ){
            window.alert("���ʊ����̗��͑S�p50�����ȓ��œ��͂��Ă�������");
            document.frm.Tokukatu.focus();
            return ;
        }
        // ���������l�̌�����������
//        if( getLengthB(document.frm.Bikou.value) > "200" ){
//            window.alert("���l�̗��͑S�p100�����ȓ��œ��͂��Ă�������");
//            document.frm.Bikou.focus();
//            return ;
//        }

	if (chk_Flg == false && p_FLG != 0) {f_Button(p_FLG);return false;} //�ύX���Ȃ��ꍇ�͂��̂܂܎���

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        document.frm.action="gak0461_upd.asp";
        document.frm.target="main";
        //document.frm.target="<%=C_MAIN_FRAME%>";
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
        document.frm.action="gak0461_main.asp";
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
<table border="0" cellpadding="1" cellspacing="1" width="520">
    <tr>
        <td align="left">
            <table width="100%" border=1 CLASS="hyo">
<%
'--------------�@�s�������@����Z�@�폜�@--------------
'                <TR>
'                    <TH CLASS="header" width="120" nowrap>�s������</TH>
'                    <TD CLASS="detail"><textarea rows=2 cols=50 class=text name="KSyoken"><%=m_TRs("T11_KODOSYOKEN")%***></textarea><br>
'                    <font size=2>�i�S�p50�����ȓ��j</font></TD>
'                </TR>
'                <TR>
'                    <TH CLASS="header" width="120" nowrap>�����Z<BR>�擾���i��</TH>
'                    <TD CLASS="detail"><textarea rows=4 cols=50 class=text name="SyumiTokugi"><%='m_TRs("T11_SYUMITOKUGI")%****></textarea><br>
'                    <font size=2>�i�S�p100�����ȓ��j</font></TD>
'                </TR>
%>
                <TR>
                    <TH CLASS="header" width="120" nowrap>�w����Q�l<BR>�ƂȂ鏔����</TH>
                    <TD CLASS="detail">
                        <table>
                            <TR align="center">
                                <TH CLASS="header" width="120" align="left" nowrap>(1)�w�K�ɂ����������<BR>(2)�s���̓����A���Z��</TH>
                            </TR>
                            <TR>
                                <TD CLASS="detail">
<!--2015/03/18 UPDATE URAKAWA-->
<!--<textarea rows=2 cols=50 class=text name="NSyoken" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN")%></textarea><br>-->
<textarea rows=4 cols=50 class=text name="NSyoken" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN")%></textarea><br>
                    <font size=2>�i�S�p50�����ȓ��j</font>
                                </TD>
                            </TR>
                            <TR>
                                <TH CLASS="header" width="120" align="left" nowrap>(3)�������A�{�����e�B�A������<BR>(4)�擾���i�A���蓙</TH>
                            </TR>
                            <TR>
                                <TD CLASS="detail">
<!--2015/03/18 UPDATE URAKAWA-->
<!--<textarea rows=2 cols=50 class=text name="NSyoken2" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN2")%></textarea><br>-->
<textarea rows=4 cols=50 class=text name="NSyoken2" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN2")%></textarea><br>
                    <font size=2>�i�S�p50�����ȓ��j</font>
                                </TD>
                            </TR>
                            <TR>
                                <TH CLASS="header" width="120" align="left" nowrap>(5)���̑�</TH>
                            </TR>
                            <TR>
                                <TD CLASS="detail">
<!--2015/03/18 UPDATE URAKAWA-->
<!--<textarea rows=2 cols=50 class=text name="NSyoken3" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN3")%></textarea><br>-->
<textarea rows=4 cols=50 class=text name="NSyoken3" onChange="chk_Flg=true;"><%=m_URs("T13_NENSYOKEN3")%></textarea><br>
                    <font size=2>�i�S�p50�����ȓ��j</font>
                                </TD>
                            </TR>
                        </table>
                    </TD>
                </TR>
                <TR>
                    <TH CLASS="header" width="120" nowrap>���ʊ���<BR>�̋L�^</TH>
<!--2015/03/18 UPDATE URAKAWA-->
<!--                    <TD CLASS="detail"><textarea rows=2 cols=50 class=text name="Tokukatu" onChange="chk_Flg=true;"><%=m_URs("T13_TOKUKATU_DET")%></textarea><br> -->
                    <TD CLASS="detail"><textarea rows=8 cols=50 class=text name="Tokukatu" onChange="chk_Flg=true;"><%=m_URs("T13_TOKUKATU_DET")%></textarea><br>
                    <font size=2>�i�S�p50�����ȓ��j</font></TD>
                </TR>

<!--2015/03/18 DELETE URAKAWA-->
<!--            <TR>
                    <TH CLASS="header" width="120" nowrap>���@�l</TH>
                    <TD CLASS="detail"><textarea rows=4 cols=50 class=text name="Bikou" onChange="chk_Flg=true;"><%=m_TRs("T11_TYOSA_BIK")%></textarea><br>
                    <font size=2>�i�S�p100�����ȓ��j</font></TD>
                </TR>
-->
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
Sub NO_Showpage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
    <html>
    <head>
    <title>�������������o�^</title>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <script language="javascript">
    </script>
    </head>
    <body>
    <center>
<br><br><br><br><br>
        <span class="msg">�I�����ꂽ�w���̒������������o�^�̃f�[�^������܂���B</span>


    </center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub
%>
