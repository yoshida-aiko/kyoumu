<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �l���C�I���Ȗڌ���
' ��۸���ID : web/web0340/web0340_main.asp
' �@      �\: ���y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/25 �O�c
' ��      �X: 2001/08/28 �ɓ����q �w�b�_���؂藣���Ή�
' ��      �X: 2015/08/19 ���{ 1�N�Ԕԍ��̕���50��70�ɕύX
' ��      �X: 2015/08/27 ���� �Ȗڂ̃f�[�^�擾���@�ύX(T15_RISYU��T16_RISYU_KOJIN)
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
    Public  m_Krs           '�Ȗڗp���R�[�h�Z�b�g
    Public  m_KSrs          '�Ȗڐ��̃��R�[�h�Z�b�g
    Dim     m_iNendo        '//�N�x
    Dim     m_sKyokanCd     '//�����R�[�h
    Dim     m_sGakunen      '//�w�N
    Dim     m_sClass        '//�N���X
    Dim     m_sKBN          '//�敪
    Dim     m_sGRP          '//�O���[�v�敪
    Dim     m_KrCnt         '//�Ȗڂ̃��R�[�h�J�E���g
    Dim     m_KSrCnt        '//�Ȗڐ��̃��R�[�h�J�E���g
    Dim     m_cell          '�z�F�̐ݒ�
	Dim		m_sRisyuJotai	'���C��ԃt���O add 2001/10/25
    Dim     i               
    Dim     j               
    Dim     k               

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

    Dim w_iRet              '// �߂�l

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�������o�^"
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

        '//�Ȗڂ̏��擾
        w_iRet = f_KamokuData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

		If m_Krs.EOF Then
			Call showPage_NoData()
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

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Krs)
    '//ں��޾��CLOSE
    'Call gf_closeObject(m_Grs)
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

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sGakunen  = request("txtGakunen")
    m_sClass    = request("txtClass")
    m_sKBN      = Cint(request("txtKBN"))
    m_sGRP      = Cint(request("txtGRP"))
    m_iDsp      = C_PAGE_LINE
    m_sRisyuJotai = Request("txtRisyu")

End Sub

Function f_KamokuData()
'******************************************************************
'�@�@�@�\�F�Ȗڂ̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear
    f_KamokuData = 1

    Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

        '//�Ȗڂ̃f�[�^�擾
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT DISTINCT "
        m_sSQL = m_sSQL & vbCrLf & "     T16_KAMOKUMEI,T16_KAMOKU_CD,T16_HAITOTANI"
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T16_RISYU_KOJIN "
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T16_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_HISSEN_KBN = " & C_HISSEN_SEN & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_HAITOTANI <> " & C_T15_HAITO & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_GRP = " & m_sGRP & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_KAMOKU_KBN = " & m_sKBN & " "
        m_sSQL = m_sSQL & vbCrLf & " AND EXISTS ( SELECT 'X' "
        m_sSQL = m_sSQL & vbCrLf & "              FROM  "
        m_sSQL = m_sSQL & vbCrLf & "                    T11_GAKUSEKI,T13_GAKU_NEN "
        m_sSQL = m_sSQL & vbCrLf & "              WHERE  "
        m_sSQL = m_sSQL & vbCrLf & "                    T13_NENDO = T16_NENDO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_GAKUSEI_NO = T16_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_GAKUSEI_NO = T11_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T11_NYUNENDO = " & w_iNyuNendo & " "
        m_sSQL = m_sSQL & vbCrLf & "             ) "

'response.write m_sSQL & "<BR>"

        Set m_Krs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Krs, m_sSQL,m_iDsp)

'response.write "w_iRet = " & w_iRet & "<BR>"
'response.write m_Krs.EOF & "<BR>"

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_KrCnt=gf_GetRsCount(m_Krs)

    f_KamokuData = 0

    Exit Do

    Loop

End Function

Function f_KamokusuData()
'******************************************************************
'�@�@�@�\�F�Ȗڐ��̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_KamokusuData = 1

    m_KSrCnt=""

    Do
        '//�Ȗڐ��̃f�[�^�擾
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT T16_KAMOKU_CD "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T16_RISYU_KOJIN ,T13_GAKU_NEN "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T16_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & " AND T16_NENDO = T13_NENDO "
        m_sSQL = m_sSQL & " AND T16_GAKUSEI_NO = T13_GAKUSEI_NO "
        m_sSQL = m_sSQL & " AND T16_HAITOGAKUNEN = T13_GAKUNEN "
        m_sSQL = m_sSQL & " AND T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & " AND T16_SELECT_FLG = " & C_SENTAKU_YES & " "
        m_sSQL = m_sSQL & " AND T16_KAMOKU_CD = '" & m_Krs("T16_KAMOKU_CD") & "' "
        m_sSQL = m_sSQL & " AND T16_HAITOGAKUNEN = " & m_sGakunen & " "
        m_sSQL = m_sSQL & " AND T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO & " "
'response.write m_sSQL
        Set m_KSrs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_KSrs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
	    m_KSrCnt=gf_GetRsCount(m_KSrs)

        If m_KSrs.EOF Then
            m_KSrCnt = "0"%>
	        <td class=disph><%=m_Krs("T16_KAMOKUMEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=m_KSrCnt%>" class="CELL2" name=Kamoku<%=i%> readonly></td>
        <%Else%>
	        <td class=disph><%=m_Krs("T16_KAMOKUMEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=m_KSrCnt%>" class="CELL2" name=Kamoku<%=i%> readonly></td>
        <%End If

	    f_KamokusuData = 0

	    Exit Do

    Loop


    '//ں��޾��CLOSE
    Call gf_closeObject(m_KSrs)

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If

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
		parent.location.href = "white.asp?txtMsg=�l���C�I���Ȗڂ̃f�[�^������܂���B"
        return;
    }
	//-->
	</SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>
    </center>
	<form name="frm" method="post">

	<input type="hidden" name="txtMsg" value="�l���C�I���Ȗڂ̃f�[�^������܂���B">

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
Dim w_iKhalf
Dim w_iGhalf
Dim n

    On Error Resume Next
    Err.Clear

i = 0
k = 0
n = 0
%>
<HTML>
<BODY>

<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>�l���C�I���Ȗڌ���</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Cansel(){

        //�󔒃y�[�W��\��
        parent.document.location.href="default2.asp"
    
    }
    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){
        parent.bottom.f_Touroku();
    }
    //-->
    </SCRIPT>

	<center>

	<FORM NAME="frm" method="post">
	<table class=disp border=1>
	    <tr>
	        <th class=header rowspan=2 width=16>�I��</th>
	    <%
	        m_Krs.MoveFirst
	        w_iKhalf = gf_Round(m_KrCnt / 2,0)
	        Do Until m_Krs.EOF
	            i = i + 1 
	            If w_iKhalf + 1 = i then
				    %>
				    </tr>
				    <tr>
				    <%
				End If 
		        Call f_KamokusuData() 
		        m_Krs.MoveNext
	        Loop%>
	    </tr>
	</table>
  <% If cint(m_sRisyuJotai) = C_K_RIS_MAE then %>
	<span class=CAUTION>�� ���肷��Ȗڂ��N���b�N���A��������Ă��������B(�����͊e�w���̊�]���ʁA���͌���)</span>
	<table>
	    <tr>
	        <td align=center><input type=button class=button value="�@�o�@�^�@" onclick="javascript:f_Touroku()"></td>
	        <td align=center><input type=button class=button value="�L�����Z��" onclick="javascript:f_Cansel()"></td>
	    </tr>
	</table>
  <% Else %>
	<BR>
	<table border="0">
	    <tr>
	        <td align=center><FONT size="1">	<BR><BR></FONT></td>
	    </tr>
	</table>
  <% End If %>
	<table class=hyo border=1>
	    <tr>
	        <th class=header width=70 height=34><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
	        <th class=header width=120>���@��</th>
	    <%

	        m_Krs.MoveFirst
	        Do Until m_Krs.EOF
		        n = n + 1
			    %>
		        <th class=header2 width=96 valign=middle><%=m_Krs("T16_KAMOKUMEI")%>
		        <input type=hidden name=kamokuCd<%=n%> value="<%=m_Krs("T16_KAMOKU_CD")%>">
		        <input type=hidden name=Tanisuu<%=n%> value="<%=m_Krs("T16_HAITOTANI")%>"></th>
			    <%
		        m_Krs.MoveNext
	        Loop%>
	    </tr>
	</table>

	</FORM>
	</center>
	</BODY>
	</HTML>
<%
End Sub
%>