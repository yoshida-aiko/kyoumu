<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A���f����
' ��۸���ID : web/web0330/regist.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
'            ���[�h         ��      txtMode
'                                   �V�K = NEW
'                                   �X�V = UPDATE
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 2001/09/01 �ɓ����q �����ȊO�����p�ł���悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iMax           ':�ő�y�[�W
    Public m_iDsp                       '// �ꗗ�\���s��
    Public m_stxtMode           '���[�h
    Public m_rs
    Public m_sSQL
    Public m_sNendo         '�N�x
    Public m_sKyokanCd      '��������
    Public m_stxtNo         '�����ԍ�
    Public m_sKenmei        '����
    Public m_sNaiyou        '���e
    Public m_sKaisibi       '�J�n��
    Public m_sSyuryoubi     '������
    Public m_sListCd

    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
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

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A���f����"
    w_sMsg=""
    w_sRetURL=C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_stxtMode = request("txtMode")

    m_sKenmei   = request("txtKenmei")
    m_sNaiyou   = request("txtNaiyou")
    m_sKaisibi  = request("txtKaisibi")
    m_sSyuryoubi= request("txtSyuryoubi")
    m_sNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_stxtNo    = request("txtNo")
    m_iDsp = C_PAGE_LINE

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

        '//���[�h�ɂ���ď��̎擾
        Select Case m_stxtMode
            Case "NEW"
        '// �y�[�W��\��

			m_sKaisibi = gf_YYYY_MM_DD(Trim(date()),"/")

            Call showPage()
            Exit Do
            Case "UPD"
            w_iRet = f_GetData()
            If w_iRet <> 0 Then
                '�ް��ް��Ƃ̐ڑ��Ɏ��s
                m_bErrFlg = True
                Exit Do
            End If
            Case "NEW2"
            w_iRet = f_NUgetData()
            If w_iRet <> 0 Then
                '�ް��ް��Ƃ̐ڑ��Ɏ��s
                m_bErrFlg = True
                Exit Do
            End If
            Case "UPD2"
            w_iRet = f_NUgetData()
            If w_iRet <> 0 Then
                '�ް��ް��Ƃ̐ڑ��Ɏ��s
                m_bErrFlg = True
                Exit Do
            End If
        End Select

            '// �y�[�W��\��
            Call showPage()
            Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ں��޾��CLOSE
    Call gf_closeObject(m_rs)
    '// �I������
    Call gs_CloseDatabase()
End Sub

Function f_GetData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_rs

    On Error Resume Next
    Err.Clear
    f_GetData = 1

    Do
        '//�ϐ��̒l���擾
        m_sSQL = ""
        m_sSQL = m_sSQL & "SELECT DISTINCT"
        m_sSQL = m_sSQL & " T46_KENMEI,T46_NAIYO,T46_KAISI,T46_SYURYO "
        m_sSQL = m_sSQL & "FROM "
        m_sSQL = m_sSQL & " T46_RENRAK "
        m_sSQL = m_sSQL & "WHERE "
        m_sSQL = m_sSQL & " T46_NO = '" & cInt(m_stxtNo) & "'"

        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

        '//�擾�����l��ϐ��ɑ��
        m_sKenmei   = w_rs("T46_KENMEI")
        m_sNaiyou   = w_rs("T46_NAIYO")
        m_sKaisibi  = w_rs("T46_KAISI")
        m_sSyuryoubi= w_rs("T46_SYURYO")

        If m_stxtMode = "UPD" Then

			m_sSQL = ""
			m_sSQL = m_sSQL & vbCrLf & " SELECT "
			m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_USER_ID "
			m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_KBN "
			m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_NAME "
			m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_CD "
			m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_GAKKA_CD "
			m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAKEIRETU_KBN "
			m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_KBN"
			m_sSQL = m_sSQL & vbCrLf & " FROM "
			m_sSQL = m_sSQL & vbCrLf & "  M10_USER "
			m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN "
			m_sSQL = m_sSQL & vbCrLf & "  ,T46_RENRAK "
			m_sSQL = m_sSQL & vbCrLf & " WHERE "
			m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_KYOKAN_CD = M04_KYOKAN.M04_KYOKAN_CD(+) "
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO = M04_KYOKAN.M04_NENDO(+)"
			m_sSQL = m_sSQL & vbCrLf & "  AND T46_RENRAK.T46_NO = " & cInt(m_stxtNo)
			m_sSQL = m_sSQL & vbCrLf & "  AND T46_RENRAK.T46_KYOKAN_CD = M10_USER.M10_USER_ID(+) "
			m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_sNendo
			m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_KYOKAKEIRETU_KBN,M04_GAKKA_CD,M10_USER_NAME"

'response.write m_sSQL & "<BR>"

            Set m_rs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)
            If w_iRet <> 0 Then
                'ں��޾�Ă̎擾���s
                m_bErrFlg = True
                Exit Do 
            End If

        End If

        f_GetData = 0

    Exit Do

    Loop
End Function

Function f_NUgetData()
'******************************************************************
'�@�@�@�\�F�f�[�^�̎擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_NUgetData = 1

    m_sListCd = request("chk")

    Do
        '//���t��̃f�[�^�擾

		'//USERID�擾�����^
		w_sUser = ""
		w_sAryUser = split(Replace(Trim(m_sListCd)," ",""),",")
		w_iCnt = UBound(w_sAryUser)

		For i = 0 To w_iCnt
			If w_sUser = "" Then
				w_sUser = "'" & w_sAryUser(i) & "'"
			Else
				w_sUser = w_sUser & ",'" & w_sAryUser(i) & "'"
			End If
		Next

        '//���t��̃f�[�^�擾
        m_sSQL = ""
		m_sSQL = m_sSQL & vbCrLf & " SELECT "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_USER_ID "
		m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_KBN "
		m_sSQL = m_sSQL & vbCrLf & "  ,M10_USER.M10_USER_NAME "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_CD "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_GAKKA_CD "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAKEIRETU_KBN "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN.M04_KYOKAN_KBN"
		m_sSQL = m_sSQL & vbCrLf & " FROM "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER "
		m_sSQL = m_sSQL & vbCrLf & "  ,M04_KYOKAN "
		m_sSQL = m_sSQL & vbCrLf & " WHERE "
		m_sSQL = m_sSQL & vbCrLf & "  M10_USER.M10_KYOKAN_CD = M04_KYOKAN.M04_KYOKAN_CD(+) "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO = M04_KYOKAN.M04_NENDO(+)"
        m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_USER_ID IN (" & w_sUser & ") "
		m_sSQL = m_sSQL & vbCrLf & "  AND M10_USER.M10_NENDO=" & m_sNendo
		m_sSQL = m_sSQL & vbCrLf & "  ORDER BY M10_USER_KBN,M04_KYOKAN_KBN,M04_GAKKA_CD,M04_KYOKAKEIRETU_KBN,M10_USER_NAME"

'response.write m_sSQL & "<BR>"

        Set m_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_rs, m_sSQL,m_iDsp)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

        f_NUgetData = 0

    Exit Do

    Loop

End Function

'********************************************************************************
'*  [�@�\]  �w�ȋL�����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  gf_GetUserNm:
'*  [����]  
'********************************************************************************
Function f_GetGakkaKigoName(p_sGakkaCd)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetGakkaKigoName = ""
	w_sName = ""

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKA_KIGO"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_NENDO=" & m_sNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_GAKKA_CD='" & p_sGakkaCd & "'"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M02_GAKKA_KIGO")
        End If

        Exit Do
    Loop

	'//�߂�l���
    f_GetGakkaKigoName = w_sName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

    Err.Clear

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
<HTML>
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>�A���f����</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  ���M��C���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Syusei(){

        var iRet;
        // ���͒l������
        iRet = f_CheckData();
        if( iRet != 0 ){
            return;
        }
        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "default.asp";
<%If m_stxtMode = "NEW" or m_stxtMode = "NEW2" Then%>
        document.frm.txtMode.value = "NEW";
<%ElseIf m_stxtMode = "UPD" or m_stxtMode = "UPD2" Then%>
        document.frm.txtMode.value = "UPD";
<%End If%>
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Touroku(){

        var iRet;
        // ���͒l������
        iRet = f_CheckData();
        if( iRet != 0 ){
            return;
        }
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "web0330_edt.asp";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_itiran(){

        //���X�g����submit
        document.frm.target = "<%=C_MAIN_FRAME%>" ;
        document.frm.action = "default.asp";
        document.frm.txtMode.value = "";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  ���͒l������
    //  [����]  �Ȃ�
    //  [�ߒl]  0:����OK�A1:�����װ
    //  [����]  ���͒l��NULL�����A�p���������A�����������s��
    //************************************************************
    function f_CheckData() {
    
        // ������NULL����������
        // ������
        if( f_Trim(document.frm.Kenmei.value) == "" ){
            window.alert("���������͂���Ă��܂���");
            document.frm.Kenmei.focus();
            return 1;
        }

        // �����e
        if( f_Trim(document.frm.Naiyou.value) == "" ){
            window.alert("���e�ɉ������͂���Ă��܂���");
            document.frm.Naiyou.focus();
            return 1;
        }

        // ���J�n��
        if( f_Trim(document.frm.Kaisibi.value) == "" ){
            window.alert("�J�n�������͂���Ă��܂���");
            document.frm.Kaisibi.focus();
            return 1;
        }

        // ��������
        if( f_Trim(document.frm.Syuryoubi.value) == "" ){
            document.frm.Syuryoubi.value = document.frm.Kaisibi.value;
        }

        // ���������t����������
        // �� �J�n��
        if( IsDate(document.frm.Kaisibi.value) != 0 ){
            window.alert("���Ԃ̊J�n���̓��t���s���ł�");
            document.frm.Kaisibi.focus();
            return 1;
        }

	    // �� ������
        if( IsDate(document.frm.Syuryoubi.value) != 0 ){
            window.alert("���Ԃ̊������̓��t���s���ł�");
            document.frm.Syuryoubi.focus();
            return 1;
        }

        // ���������e�̌�����������
        if( getLengthB(document.frm.Naiyou.value) > "254" ){
            window.alert("���e�̗��͑S�p127�����ȓ��œ��͂��Ă�������");
            document.frm.Naiyou.focus();
            return 1;
        }

        // �����������̌�����������
        if( getLengthB(document.frm.Kenmei.value) > "40" ){
            window.alert("�����̗��͑S�p20�����ȓ��œ��͂��Ă�������");
            document.frm.Kenmei.focus();
            return 1;
        }

        // ���������Ԃ̎擾������������
        if ( f_Trim(document.frm.Syuryoubi.value) != "" ){
	        if( DateParse(document.frm.Kaisibi.value,document.frm.Syuryoubi.value) < 0){
	            window.alert("�J�n���ƏI�����𐳂������͂��Ă�������");
	            document.frm.Kaisibi.focus();
	            return 1;
	        }
        }
        return 0;
    }

    //-->
    </SCRIPT>

<center>

<FORM NAME="frm" method="post">

<br>

<% If m_stxtMode = "NEW"or m_stxtMode = "NEW2" Then 
    call gs_title("�A���f����","�V�@�K")
   Else 
    call gs_title("�A���f����","�C�@��")
   End If%>

<br>
<font>�o�@�^�@���@�e</font>
<br>
<br>
<%If m_stxtMode = "NEW" Then %>
<div align="center"><span class=CAUTION>�� ���͎�������͂��A����t��I��{�^�����N���b�N���Ă��������B<br>
</span></div>
<%ElseIf m_stxtMode = "UPD" Then%> 
<div align="center"><span class=CAUTION>�� �C�����������ڂ��C�����A���t���ύX�������ꍇ�͢���t��I��{�^�����N���b�N���Ă��������B
</span></div>
<%Else%> 
<div align="center"><span class=CAUTION>�� ���t���ύX�������ꍇ�͢���t��I��{�^�����N���b�N���Ă��������B<br>
										�� �o�^���Ă悯��Γo�^�{�^�����N���b�N���Ă��������B
</span></div>
<%End If%>
</TD>
</TR>
</TABLE>

<br>
<table width="510" border=1 CLASS="hyo">
    <TR>
        <TH CLASS="header" width="60">����</TH>
        <TD CLASS="detail"><input type="text" size="57" name="Kenmei" value="<%=m_sKenmei%>" maxlength=40><br>
        <font size=2>�i�S�p20�����ȓ��j</font></TD>
    </TR>
    <TR>
        <TH CLASS="header" width="60">���e</TH>
        <TD CLASS="detail"><textarea rows=6 cols=40 class=text name="Naiyou"><%=m_sNaiyou%></textarea><br>
        <font size=2>�i�S�p127�����ȓ��j</font></TD>
    </TR>
    <TR>
        <TH CLASS="header" width="60">����</TH>
        <TD CLASS="detail"><input type="text" size="23" name="Kaisibi" value="<%=m_sKaisibi%>" maxlength=10>
        <input type="button" class="button" onclick="fcalender('Kaisibi')" value="�I��">
        �@�`�@<input type="text" size="23" name="Syuryoubi" value="<%=m_sSyuryoubi%>" maxlength=10>
        �@<input type="button" class="button" onclick="fcalender('Syuryoubi')" value="�I��"><br>
        <font size=2>�i���͗�:<%=Date()%>�j</font></TD>
    </TR>
<%
    If m_stxtMode <> "NEW" Then
%>
    <tr>
    <td colspan=2 align=right bgcolor=#9999BD>
	<input type="button" value="�o�@�^" class=button onclick="javascript:f_Touroku()">
	<input type="button" value="�L�����Z��" class=button onclick="javascript:f_itiran()">
	<input class=button type=button value="���t��I��" onclick="javascript:f_Syusei()"></td>
    </tr>
    <TR>
        <TH CLASS="header" valign="top">���t��</TD>
        <TD CLASS="detail" colspan=2>
        <table border=1 class=hyo width=100% height=100%>
<%
    m_rs.MoveFirst
    Do Until m_rs.EOF
%>
		    <TR>

			<%
			'========================================================
			'//�敪���̓��擾
			w_sKyokanKbnName = ""
			w_sKeiretuKbnName = ""
			w_sGakkaKigo = ""

			'//����CD���Z�b�g
			w_sKyokanCd = m_rs("M04_KYOKAN_CD")

			'//�����̎�(����CD����̏ꍇ)
			If LenB(w_sKyokanCd) <> 0 Then
				'//�����敪���̂��擾
				Call gf_GetKubunName(C_KYOKAN,m_rs("M04_KYOKAN_KBN"),m_sNendo,w_sKyokanKbnName)

				'//���Ȍn��敪���̂��擾
				Call gf_GetKubunName(C_KYOKA_KEIRETU,m_rs("M04_KYOKAKEIRETU_KBN"),m_sNendo,w_sKeiretuKbnName)

				w_sGakkaKigo = f_GetGakkaKigoName(m_rs("M04_GAKKA_CD"))
			Else
				'//�����ȊO�̏ꍇUSER�敪���̂�\��
				Call gf_GetKubunName(C_USER,m_rs("M10_USER_KBN"),m_sNendo,w_sKyokanKbnName)
				w_sKeiretuKbnName = "�\"
				w_sGakkaKigo = "�\"
			End If

			'========================================================

            Call gs_cellPtn(w_cell)
			%>

	        <td class="CELL2"><%=w_sKyokanKbnName%><BR></td>
	        <td class="CELL2"><%=w_sKeiretuKbnName%>
				<input type="hidden" name="KCD" value='<%=m_rs("M10_USER_ID")%>'><BR></td>
	        <td class="CELL2"><%=w_sGakkaKigo%><BR></td>
	        <td class="CELL2"><%=m_rs("M10_USER_NAME")%><BR></td>
		    </TR>
<%
    m_rs.MoveNext
    Loop
%>
        </table>
		</td>
<%
    Else
%>
    <tr>
    <td colspan=6 align=right bgcolor=#9999BD>
	<input type="button" value="�L�����Z��" class=button onclick="javascript:f_itiran()">
	<input class=button type=button value="���t��I��" onclick="javascript:f_Syusei()"></td>
<%
    End If
%>
    </tr>

</TABLE>
    <INPUT TYPE=HIDDEN  NAME=txtNo value="<%=m_stxtNo%>">
    <INPUT TYPE=HIDDEN  NAME=txtMode value="<%=m_stxtMode%>">
    <INPUT TYPE=HIDDEN  NAME=txtNendo   VALUE="<%=m_sNendo%>">
    <INPUT TYPE=HIDDEN  NAME=txtListCd      value="<%=m_sListCd%>">
    <INPUT TYPE=HIDDEN  NAME=txtKyokanCd    VALUE="<%=m_sKyokanCd%>">
</FORM>
</center>
</BODY>
</HTML>
<%
End Sub
%>
