<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����o������
' ��۸���ID : kks/kks0170/kks0170_middle.asp
' �@      �\: ���y�[�W ���Əo�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��: SESSION("NENDO")           '//�����N
'             SESSION("KYOKAN_CD")       '//����CD
'             TUKI           '//��
'             cboDate        '//���t
' ��      ��:
' ��      �n: NENDO"        '//�����N
'             KYOKAN_CD     '//����CD
'             GAKUNEN"      '//�w�N
'             CLASSNO"      '//�׽No
'             cboDate"      '//���t
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��S�C�׽���k����\��
'           ���o�^�{�^���N���b�N��
'               ���͏���o�^����
'-------------------------------------------------------------------------
' ��      ��: 2001/07/24 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bTannin       '//�S�C�׸�

    '�擾�����f�[�^�����ϐ�
    Public m_iSyoriNen      '//�����N�x
    Public m_iKyokanCd      '//����CD
    Public m_sDate          '//���t
    Public m_iGakunen       '//�w�N
    Public m_iClassNo       '//�N���XNo
    Public m_sClassNm       '//�N���X����
    Public m_iRsCnt         '//�N���Xں��ސ�
	Public m_sEndDay		'//���͂ł��Ȃ��Ȃ��

    'ں��ރZ�b�g
    Public m_Rs

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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�����o������"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_bTannin = False

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

        '//�ϐ�������
        Call s_ClearParam()

        '// ���Ұ�SET
        Call s_SetParam()

        '// �S�C�N���X���擾
        w_iRet = f_GetClassInfo(m_bTannin)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'���͕s�ɂȂ�����擾
		call gf_Get_SyuketuEnd(m_iGakunen,m_sEndDay)

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
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_sDate     = ""
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_sClassNm = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = SESSION("NENDO")
    m_iKyokanCd = SESSION("KYOKAN_CD")
    m_sDate     = trim(Request("cboDate"))

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub
    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sDate     = " & m_sDate     & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_sClassNm  = " & m_sClassNm  & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �S�C�N���X���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_GetClassInfo(p_bTannin)

    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear
    
    f_GetClassInfo = 1

    Do 

        '// �S�C�N���X���
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_TANNIN"
        w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & m_iKyokanCd & "'"

'response.write w_sSQL & "<BR>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetClassInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            p_bTannin = True 
            m_iGakunen = rs("M05_GAKUNEN")
            m_iClassNo = rs("M05_CLASSNO")
            m_sClassNm = rs("M05_CLASSMEI")
        End If

        f_GetClassInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �S�C�N���X�ꗗ�擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_GetClassList()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetClassList = 1

    Do 

        '// �S�C�N���X���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_IDOU_NUM, "
        w_sSQL = w_sSQL & vbCrLf & "  B.T11_SIMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  B.T11_GAKUSEI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  C.T30_HIDUKE, "
        w_sSQL = w_sSQL & vbCrLf & "  C.T30_SYUKKETU_KBN,"
        '//"�o��"�͕\�����Ȃ�
        w_sSQL = w_sSQL & vbCrLf & "  DECODE(D.M01_SYOBUNRUIMEI_R,'�o','',D.M01_SYOBUNRUIMEI_R) AS SYUKKETU_MEI"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN A"
        w_sSQL = w_sSQL & vbCrLf & "  ,T11_GAKUSEKI B"
        w_sSQL = w_sSQL & vbCrLf & "  ,(SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     T30_HIDUKE,"
        w_sSQL = w_sSQL & vbCrLf & "     T30_SYUKKETU_KBN,"
        w_sSQL = w_sSQL & vbCrLf & "     T30_GAKUSEKI_NO"
        w_sSQL = w_sSQL & vbCrLf & "    FROM T30_KESSEKI"
        w_sSQL = w_sSQL & vbCrLf & "    WHERE T30_HIDUKE='" & m_sDate & "'"
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_GAKUNEN=" & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_CLASS=" & m_iClassNo & ") C"
        w_sSQL = w_sSQL & vbCrLf & "  ,(SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     M01_SYOBUNRUI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "     M01_SYOBUNRUIMEI_R"
        w_sSQL = w_sSQL & vbCrLf & "    FROM M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & "    WHERE "
        w_sSQL = w_sSQL & vbCrLf & "          M01_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "      AND M01_DAIBUNRUI_CD=" & C_KESSEKI & ") D"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        'w_sSQL = w_sSQL & vbCrLf & "      A.T13_NENDO - A.T13_GAKUNEN + 1 = B.T11_NYUNENDO(+) "
        w_sSQL = w_sSQL & vbCrLf & "      A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO "
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_GAKUSEKI_NO = C.T30_GAKUSEKI_NO(+)"
        w_sSQL = w_sSQL & vbCrLf & "  AND C.T30_SYUKKETU_KBN = D.M01_SYOBUNRUI_CD(+)"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_GAKUNEN=" & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_CLASS=" & m_iClassNo
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T13_GAKUSEKI_NO"

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetClassList = 99
            Exit Do
        End If

        '//ں��ރJ�E���g���擾
        m_iRsCnt = gf_GetRsCount(m_Rs)

        f_GetClassList = 0
        Exit Do
    Loop

End Function


'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear

%>
    <html>
    <head>
    <title>�����o������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
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
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Cancel(){
        //�����y�[�W��\��
        //parent.document.location.href="default2.asp"
        document.frm.target = "<%=C_MAIN_FRAME%>";
        document.frm.action = "./default.asp"
        document.frm.submit();
        return;


    }

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Touroku(){

		parent.frames["main"].f_Touroku();
    }


    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">
<%
'//�f�o�b�O
'Call s_DebugPrint()
%>
    <center>
    <form name="frm" method="post">
    <%call gs_title("�����o������","��@��")%>
    <%Do
        '//�S�C�N���X���Ȃ��ꍇ
        If m_bTannin = False Then
        %>
        <br><br>
        <span class="msg">�󎝃N���X������܂���B</span>
        <%
            Exit Do
        End If
		%>

        <table>
		<tr><td>
	        <table class="hyo" border="1" width="400">
	            <tr>
	                <th nowrap class="header" width="64"  align="center">�N���X</th>
	                <td nowrap class="detail" width="50"  align="center"><%=m_iGakunen%>�N</td>
	                <td nowrap class="detail" width="130" align="center"><%=m_sClassNm%></td>
	                <th nowrap class="header" width="100" align="center">���͑Ώۓ�</th>
	                <td nowrap class="detail" width="150" align="center"><%=m_sDate & "(" & gf_GetYoubi(Weekday(m_sDate)) & ")"%></td>
	            </tr>
	        </table>
		</td></tr><tr>
		<td align="center">
<%		if m_sEndDay < m_sDate then %>
            <table>
				<tr>
                <td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="�@�o�@�^�@"></td>
                <td ><input class="button" type="button" onclick="javascript:f_Cancel();" value="�L�����Z��"></td>
				</tr>
            </table>
<% Else %>
            <table>
				<tr>
                <td ><input class="button" type="button" onclick="javascript:f_Cancel();" value=" �߁@�� "></td>
				</tr>
            </table>
<% End If %>
		</td></tr>
        </table>

        <!--���׃w�b�_��-->
        <table >
<%		if m_sEndDay < m_sDate then %>
            <tr>
                <td align="center" colspan=3 valign="bottom">
                    <span class="CAUTION">�� �o���󋵗����N���b�N���āA�o���󋵂���͂��Ă��������B�i�����x��������(�o��)�̏��ŕ\������܂��j</span>
                </td>
            </tr>
<% Else%>
            <tr>
                <td align="center" colspan=3 valign="bottom">
                    <span class="CAUTION">�� �o���󋵂�ύX���邱�Ƃ͂ł��܂���B</span>
                </td>
            </tr>

<% End If %>

            <tr><td valign="top">

                <!--�w�b�_-->
                <table class=hyo border="1" bgcolor="#FFFFFF">
                    <tr>
                        <th nowrap class="header" width="80"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
                        <th nowrap class="header" width="150" align="center">���@��</th>
                        <th nowrap class="header" width="80" align="center">�o����</th>
                    </tr>

            <%If i = w_iCnt Then
                '//���X�g�����s����

                '//���ټ�Ă̸׽��������
				w_Class = ""
                %>
                </table>
                </td><td width="10"></td><td valign="top">
                <!--�w�b�_-->
                <table class="hyo" border="1" >
                    <tr>
                        <th nowrap class="header" width="80"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
                        <th nowrap class="header" width="150" align="center">���@��</th>
                        <th nowrap class="header" width="80" align="center">�o����</th>
            <%End If%>
                </table>
                </td></tr>
            </table>

        <%Exit Do%>
    <%Loop%>

    <!--�l�n���p-->
    <input type="hidden" name="NENDO"     value="<%=m_iSyoriNen%>">
    <input type="hidden" name="KYOKAN_CD" value="<%=m_iKyokanCd%>">
    <input type="hidden" name="GAKUNEN"   value="<%=m_iGakunen%>">
    <input type="hidden" name="CLASSNO"   value="<%=m_iClassNo%>">
    <input type="hidden" name="cboDate"   value="<%=m_sDate%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
'********************************************************************************
'*  [�@�\]  ��HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showWhitePage()
%>
    <html>
    <head>
    <title>�����o������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

    }
    //-->
    </SCRIPT>
    </head>

	<body LANGUAGE=javascript onload="return window_onload()">
    </body>
    </html>
<%
End Sub
%>

