<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A�E��}�X�^
' ��۸���ID : mst/mst0144/kakunin.asp
' �@      �\: ���y�[�W �A�E��}�X�^�̏ڍוύX���s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H�R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSinroName        :�i�H���́i�ꕔ�j
'           txtPageSinro        :�\���ϕ\���Ő��i�������g����󂯎������j
'           Sinro_syuseiCD      :�I�����ꂽ�i�H�R�[�h
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H�R�[�h�i�߂�Ƃ��j
'           txtSingakuCd        :�i�w�R�[�h�i�߂�Ƃ��j
'           txtSinroName        :�i�H���́i�߂�Ƃ��j
'           txtPageSinro        :�\���ϕ\���Ő��i�߂�Ƃ��j
' ��      ��:
'           �������\��
'               �w�肳�ꂽ�i�w��E�A�E��̏ڍ׃f�[�^��\��
'           ���n�}�摜�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ��i�w��E�A�E���\������i�ʃE�B���h�E�j
'-------------------------------------------------------------------------
' ��      ��: 2001/07/12 �≺ �K��Y
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�

    Public  m_Rs            'recordset
    Public  m_sDBMode           'DB��Ӱ�ނ̐ݒ�
    Public  m_sDATE
    Public  m_sNendo
    Public  m_sGakkiCD
    Public  m_sGakunenCD
    Public  m_sGakkaCD
    Public  m_sCourseCD
    Public  m_sKamokuCD
    Public  m_sKyokanCD
    Public  m_sKyokanMei
    Public  m_sKyokasyoName
    Public  m_sSyuppansya
    Public  m_sTyosya
    Public  m_sKyokanyo
    Public  m_sSidousyo
    Public  m_sBiko
    Public  m_sMaxNO

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

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E��}�X�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

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

        '// DB�o�^
        if m_sDBMode = "Insert" then
            w_iRet = f_Insert
        else
            w_iRet = f_Update
        end if

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
    call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    Dim strErrMsg

    strErrMsg = ""

    m_sDBMode    = Request("txtMode")       'DBӰ�ނ̎擾
    m_sNendo     = Request("txtNendo")      '�N�x�̎擾
    m_sGakkiCD   = Request("txtGakkiCD")    '�w���̎擾
    m_sGakunenCD     = Request("txtGakunenCD")  '�w�N�̎擾
    m_sGakkaCD   = Request("txtGakkaCD")    '�w�Ȃ̎擾
    m_sCourseCD  = Request("txtCourseCD")   '�R�[�X�̎擾
    m_sKamokuCD  = Request("txtKamokuCD")   '�Ȗڂ̎擾
    m_sKyokanMei     = Request("txtKyokanMei")  '�������̎擾
    m_sKyokasyoName  = Request("txtKyokasyoName")   '���ȏ����̎擾
    m_sSyuppansya    = Request("txtSyuppansya") '�o�ŎЂ̎擾
    m_sTyosya    = Request("txtTyosya")     '���҂̎擾
    m_sKyokanyo  = Request("txtKyokanyo")   '�����p�̎擾
    m_sSidousyo  = Request("txtSidousyo")   '�w�����̎擾
    m_sBiko      = Request("txtBiko")       '�����p�̎擾

    m_sDate = gf_YYYY_MM_DD(date(),"/")

    'm_sKyokanCD = Session("KYOKAN_CD")          ':���[�U�[ID
    m_sKyokanCD = Request("SKyokanCd1")          ':���[�U�[ID

    if m_sDBMode = "Insert" then
        Call f_Max()
        m_sMaxNO = Cint(m_sMaxNO) + 1
    else
        m_sMaxNO = Request("txtUpdNo")
    end if

    If strErrmsg <> "" Then
        ' �G���[��\������t�@���N�V����
        Call err_page(strErrMsg)
        response.end
    End If
'   call s_viewForm(request.form)   '�f�o�b�O�p�@�����̓��e������
End Sub


'********************************************************************************
'*  [�@�\]  �V�K�A����R�[�h�𐶐�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_Max()

Dim w_Rs
Dim w_sSQL
dim f_MaxNO 

    f_MaxNO = 0
    m_sMaxNO = 0

    '// �A����R�[�h���擾�i�A����R�[�hMax�l�j
    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT "
    w_sSQL = w_sSQL & " MAX( "
    w_sSQL = w_sSQL & " T47.T47_NO"
    w_sSQL = w_sSQL & " ) "
    w_sSQL = w_sSQL & " AS MAXNO "
    w_sSQL = w_sSQL & " FROM T47_KYOKASYO T47 "
    w_sSQL = w_sSQL & " WHERE T47_NENDO = " & m_sNendo & " "

'response.write w_sSQL & "<<<BR>"

    w_sRet = gf_GetRecordset(w_Rs, w_sSQL)

    If w_sRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_sMsg = Err.description
        Exit function
    End If

    IF w_Rs.EOF THEN
        f_MaxNO = 0
    Else
        f_MaxNO = gf_SetNull2Zero(w_Rs("MAXNO"))
    End If
    
    m_sMaxNO = f_MaxNO

'    response.write("<BR>m_sMaxNO = " & m_sMaxNO)

End Function

'********************************************************************************
'*  [�@�\]  �V�K�o�^����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
function f_Insert
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��

    f_Insert = False

    w_sSQL = w_sSQL & vbCrLf & " Insert Into "
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKASYO"
    w_sSQL = w_sSQL & "(T47_NENDO,"
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKKI_KBN,"
    w_sSQL = w_sSQL & vbCrLf & " T47_NO,"
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKUNEN,"
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKKA_CD,"
    w_sSQL = w_sSQL & vbCrLf & " T47_COURSE_CD,"
    w_sSQL = w_sSQL & vbCrLf & " T47_KAMOKU,"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKAN,"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKASYO,"
    w_sSQL = w_sSQL & vbCrLf & " T47_SYUPPANSYA,"
    w_sSQL = w_sSQL & vbCrLf & " T47_TYOSYA,"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKANYOUSU,"
    w_sSQL = w_sSQL & vbCrLf & " T47_SIDOSYOSU,"
    w_sSQL = w_sSQL & vbCrLf & " T47_BIKOU,"
    w_sSQL = w_sSQL & vbCrLf & " T47_INS_DATE,"
    w_sSQL = w_sSQL & vbCrLf & " T47_INS_USER,"
    w_sSQL = w_sSQL & vbCrLf & " T47_UPD_DATE,"
    w_sSQL = w_sSQL & vbCrLf & " T47_UPD_USER)"
    w_sSQL = w_sSQL & vbCrLf & " Values"
    w_sSQL = w_sSQL & "(" & m_sNendo & ","
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sGakkiCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " " & m_sMaxNO & ","
    w_sSQL = w_sSQL & vbCrLf & " " & m_sGakunenCD & ","
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sGakkaCD & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sCourseCD & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sKamokuCD & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sKyokanCD & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sKyokasyoName & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sSyuppansya & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sTyosya & "',"
    w_sSQL = w_sSQL & vbCrLf & " " & m_sKyokanyo & ","
    w_sSQL = w_sSQL & vbCrLf & " " & m_sSidousyo & ","
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sBiko & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sDATE & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & m_sDATE & "',"
    w_sSQL = w_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "')"

    w_iRet = gf_ExecuteSQL(w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    End If

    f_Insert = True

End Function

'********************************************************************************
'*  [�@�\]  �X�V����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
function f_Update
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��

    f_Update = False

    w_sSQL = w_sSQL & vbCrLf & " Update T47_KYOKASYO SET "
    w_sSQL = w_sSQL & vbCrLf & " T47_NENDO = " & m_sNendo & ","
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKKI_KBN = '" & m_sGakkiCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKUNEN = " & m_sGakunenCD &","
    w_sSQL = w_sSQL & vbCrLf & " T47_GAKKA_CD = '" & m_sGakkaCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_COURSE_CD = '" & m_sCourseCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_KAMOKU = '" & m_sKamokuCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKAN = '" & m_sKyokanCD & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKASYO = '" & m_sKyokasyoName & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_SYUPPANSYA = '" & m_sSyuppansya & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_TYOSYA = '" & m_sTyosya & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_KYOKANYOUSU = " & m_sKyokanyo & ","
    w_sSQL = w_sSQL & vbCrLf & " T47_SIDOSYOSU = " & m_sSidousyo & ","
    w_sSQL = w_sSQL & vbCrLf & " T47_BIKOU = '" & m_sBiko & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_UPD_DATE = '" & m_sDATE & "',"
    w_sSQL = w_sSQL & vbCrLf & " T47_UPD_USER = '" & Session("LOGIN_ID") & "' "
    w_sSQL = w_sSQL & vbCrLf & " Where T47_NENDO = " & Request("KeyNendo")
'    w_sSQL = w_sSQL & vbCrLf & " and T47_GAKKI_KBN = '" & m_sGakkiCD & "'"
    w_sSQL = w_sSQL & vbCrLf & " and T47_NO = " & m_sMaxNO 

'Response.Write w_sSQL & "<br>"

    w_iRet = gf_ExecuteSQL(w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    End If

    f_Update = True

End Function

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
'Sub s_MapHTML()

'   If ISNULL(m_Rs("M13_TIZUFILENAME")) OR m_Rs("M13_TIZUFILENAME")="" Then
'       Response.Write("�o�^����Ă��܂���")
'   Else
'       Response.Write("<a Href=""javascript:f_OpenWindow('" & Session("TYUGAKU_TIZU_PATH") & m_Rs("M13_TIZUFILENAME") & "')"">���Ӓn�}</a>")
'   End If
    
'End Sub


Sub S_syousai()
'********************************************************************************
'*  [�@�\]  �ڍׂ�\��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

Dim w_slink
Dim w_iCnt

w_iCnt = 0

Do While not m_Rs.EOF

w_slink = "�@"

if m_Rs("M32_SINRO_URL") <> "" Then 
    w_sLink= "<a href='" & gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "'>" 
    w_sLink= w_sLink &  gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "</a>"
End if

        %>
        <%=w_slink%>
        <%
            m_Rs.MoveNext

        Loop

    'LABEL_showPage_OPTION_END
End sub


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

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    function gonext() {
        window.alert('<%=C_TOUROKU_OK_MSG%>');

            document.frm.action = "./default.asp";
            document.frm.target="fTopMain";
            document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>

<body bgcolor="#ffffff" onLoad="setTimeout('gonext()',0000)">

<center>

<Form Name ="frm" Action="">


<input type="hidden" Name="txtMode" Value="">
<input type="hidden" name="SKyokanCd1" value="<%=m_sKyokanCD%>">

</From>
</center>

</body>

</html>


<%
    '---------- HTML END   ----------
End Sub

Sub Nyuryokuzumi()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

%>

    <html>
    <head>
    </head>

    <body>

    <center>
    <font size="2">���͂��ꂽ�A����R�[�h�͂��łɎg�p�ς݂ł�<br><br></font>
    <input type="button" onclick="javascript:history.back()" value="�߁@��">
    </center>
    </body>

    </html>


<%
    '---------- HTML END   ----------
End Sub
%>