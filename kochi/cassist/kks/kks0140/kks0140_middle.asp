<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �s���o������
' ��۸���ID : kks/kks0140/kks0140_middle.asp
' �@      �\: ���y�[�W �s���o�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��: NENDO     '//�N�x
'             KYOKAN_CD '//����CD
'             GAKUNEN   '//�w�N
'             CLASSNO   '//�N���XNO
'             GYOJI_CD  '//�s��CD
'             GYOJI_MEI '//�s����
'             KAISI_BI  '//�J�n��
'             SYURYO_BI '//�I����
'             SOJIKANSU '//�����Ԑ�
' ��      ��:
' ��      �n: NENDO     '//�N�x
'             KYOKAN_CD '//����CD
'             GAKUNEN   '//�w�N
'             CLASSNO   '//�N���XNO
'             GYOJI_CD  '//�s��CD
'             GYOJI_MEI '//�s����
'             KAISI_BI  '//�J�n��
'             SYURYO_BI '//�I����
'             SOJIKANSU '//�����Ԑ�
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��s���o�����͂�\��
'           ���\���{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ����w�Z��\��������
'           ���o�^�{�^���N���b�N��
'               ���͏���o�^����
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
    Const DebugPrint = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public m_iSyoriNen      '//�����N�x
    Public m_iKyokanCd      '//����CD
    Public m_sGakunen       '//�w�N
    Public m_sClassNo       '//�׽NO
    Public m_sTuki          '//��
    Public m_sGyoji_Cd      '//�s��CD
    Public m_sGyoji_Mei     '//�s����
    Public m_sKaisi_Bi      '//�J�n��
    Public m_sSyuryo_Bi     '//�I����
    Public m_sSoJikan       '//�����Ԑ�
	Public m_sEndDay		'//���͂ł��Ȃ��Ȃ��

    '//ں��ރZ�b�g
    Public m_Rs_M           '//recordset���׏��
    Public m_Rs_G           '//recordset�s���o�����
    Public m_iRsCnt         '//�w�b�_ں��ސ�

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
    w_sMsgTitle="�s���o������"
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


        '//�ϐ�������
        Call s_ClearParam()

        '// ���Ұ�SET
        Call s_SetParam()

'//�f�o�b�O
'Call s_DebugPrint()

        '// ���k���X�g���擾
        w_iRet = f_Get_DetailData()
        If w_iRet <> 0 Then
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
    Call gf_closeObject(m_Rs_M)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen   = ""
    m_iKyokanCd  = ""
    m_sGakunen = ""
    m_sClassNo = ""
    m_sTuki = ""

    m_sGyoji_Cd  = ""
    m_sGyoji_Mei = ""
    m_sKaisi_Bi  = ""
    m_sSyuryo_Bi = ""
    m_sSoJikan   = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = trim(Request("NENDO"))
    m_iKyokanCd = trim(Request("KYOKAN_CD"))
    m_sGakunen  = trim(Request("GAKUNEN"))
    m_sClassNo  = trim(Request("CLASSNO"))
    m_sTuki     = trim(Request("TUKI"))

    m_sGyoji_Cd  = trim(Request("GYOJI_CD"))
    m_sGyoji_Mei = trim(Request("GYOJI_MEI"))
    m_sKaisi_Bi  = trim(Request("KAISI_BI"))
    m_sSyuryo_Bi = trim(Request("SYURYO_BI"))
    m_sSoJikan   = trim(Request("SOJIKANSU"))
    m_sEndDay   = trim(Request("ENDDAY"))

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()

    response.write "<font color=#000000>m_iSyoriNen = " & m_iSyoriNen  & "</font><br>"
    response.write "<font color=#000000>m_iKyokanCd = " & m_iKyokanCd  & "</font><br>"
    response.write "<font color=#000000>m_sGakunen  = " & m_sGakunen   & "</font><br>"
    response.write "<font color=#000000>m_sClassNo  = " & m_sClassNo   & "</font><br>"
    response.write "<font color=#000000>m_sTuki     = " & m_sTuki      & "</font><br>"

    response.write "<font color=#000000>m_sGyoji_Cd = " & m_sGyoji_Cd  & "</font><br>"
    response.write "<font color=#000000>m_sGyoji_Mei= " & m_sGyoji_Mei & "</font><br>"
    response.write "<font color=#000000>m_sKaisi_Bi = " & m_sKaisi_Bi  & "</font><br>"
    response.write "<font color=#000000>m_sSyuryo_Bi= " & m_sSyuryo_Bi & "</font><br>"
    response.write "<font color=#000000>m_sSoJikan  = " & m_sSoJikan   & "</font><br>"

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�ꗗ���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Get_DetailData()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_Get_DetailData = 1

    Do 

        '// ���׃f�[�^
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN," 
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "   T11_GAKUSEKI.T11_SIMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN,T11_GAKUSEKI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN=" & cInt(m_sGakunen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS=" & cInt(m_sClassNo)
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write "<font color=#000000>" & w_sSQL & "</font><BR>"
        iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_DetailData = 99
            Exit Do
        End If

        '//�������擾
        m_iRsCnt = 0
        If m_Rs_M.EOF = False Then
            m_iRsCnt = gf_GetRsCount(m_Rs_M)
        End If

        f_Get_DetailData = 0
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
    <title>�s���p�o������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
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

		//�X�N���[����������
		parent.init();

        <%If m_Rs_M.EOF = True Then%>
			document.location.href="white.htm"
			return;
		<%End If%>
    }

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Cancel(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default.asp"
    }

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){

		parent.frames["main"].f_Touroku();
		return;
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>

    <form name="frm" method="post" >
    <%call gs_title("�s���o������","��@��")%>
	<br>
    <%Do%>
        <%If m_Rs_M.EOF = True Then
			Exit Do
		End If%>

        <table>
		<tr><td>
            <table class=hyo width="590" border="1" >
                <tr>
                    <th nowrap class="header" width="80"  align="center">�s����</th>
                    <td nowrap class="detail" width="200" align="left">�@<%=m_sGyoji_Mei%></td>
                    <th nowrap class="header" width="80"  align="center">������</th>
                    <td nowrap class="detail" width="50"  align="center"><%=m_sSoJikan%></td>
                    <th nowrap class="header" width="80"  align="center">���{��</th>
                    <td nowrap class="detail" width="100" align="center"><%=month(m_sKaisi_Bi) & "/" & day(m_sKaisi_Bi)%>�`<%=month(m_sSyuryo_Bi) & "/" & day(m_sSyuryo_Bi)%></td>
                </tr>
            </table>
		</td></tr><tr>
		<td align="center">
	<% 'If m_sEndDay < m_sSyuryo_Bi then %>
            <table>
                <td ><input class=button type="button" onclick="javascript:f_Touroku();" value="�@�o�@�^�@"></td>
                <td ><input class=button type="button" onclick="javascript:f_Cancel();" value="�L�����Z��"></td>
            </table>
	<% 'Else %>
            <!--table>
                <td ><input class=button type="button" onclick="javascript:f_Cancel();" value=" �߁@�� "></td>
            </table-->
	<% 'End If %>
		</td></tr>
        </table>

        <!--���׃w�b�_��-->

        <table >
            <tr>
				<td valign="top">
		            <table class="hyo"  border="1" >
		               <tr>
		                   <th class="header" width="80"  height="23" align="center"  nowrap><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
		                   <th class="header" width="150" height="23" align="center"  nowrap>���@��</th>
		                   <th class="header" width="80"  height="23" align="center"  nowrap>���ێ���</th>
		               </tr>
		            </table>
	            </td>
			<%If m_iRsCnt <> 1 Then%>
				<td width="10"><br></td>
				<td valign="top" >
	                <table class="hyo"  border="1" >
                        <tr>
                            <th class="header" width="80"  height="23" align="center"  nowrap><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
                            <th class="header" width="150" height="23" align="center"  nowrap>���@��</th>
                            <th class="header" width="80"  height="23" align="center"  nowrap>���ێ���</th>
		               </tr>
                  </table>
                </td>
			<%End If%>
			</tr>
        </table>

        <%
        Exit Do

    Loop%>

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>