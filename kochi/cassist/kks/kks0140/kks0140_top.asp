<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �s���o������
' ��۸���ID : kks/kks0140/kks0140_top.asp
' �@      �\: ��y�[�W �����Ƃ̍s������\������
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
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
'               ���̃R���{�{�b�N�X�͓�����\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��s���ꗗ��\��������
'           ���o�^�{�^���N���b�N��
'               ���͂��ꂽ����o�^����
'-------------------------------------------------------------------------
' ��      ��: 2001/07/03 �ɓ����q
' ��      �X: 2001/12/07 ������ �s���敪�̒ǉ��ɑΉ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '��������
    Public m_iKyokanCd          '�N�x

    Public m_sGakunen           '//�w�N
    Public m_sClassNo           '//�N���XNO
    Public m_sClassMei          '//�N���X��
    Public m_sTuki              '//��
    Public m_Rs
    Public m_sNoTanMsg

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
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '//�l�̏�����
        Call s_ClearParam()

        '//�ϐ��Z�b�g
        Call s_SetParam()

        '//�������擾
        If request("TUKI") <> "" Then
            m_sTuki = request("TUKI")
        Else
            m_sTuki = month(date())
        End If

        '//�N�x�A����CD���S�C�N���X�����擾
        w_iRet = f_GetClassInfo()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//�S�C�N���X�����邩�ǂ���
        If trim(m_sGakunen) = "" AND trim(m_sClassNo) = "" Then
            '//�󎝃N���X���Ȃ���
            m_sNoTanMsg = "�S�C�N���X������܂���"
        Else
            '//�S�C�N���X������Ƃ��̂ݕ\��
            '// �w�b�_���X�g���擾
            w_iRet = f_Get_HeadData()
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If
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
    Call gf_closeObject(m_Rs)
    '// �I������
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
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sClassMei = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")    
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"
    response.write "m_sClassMei = " & m_sClassMei & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  ����CD���A�S�C�N���X�����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetClassInfo()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetClassInfo = 1

    Do
        '�N���X�}�X�^����N���X�����擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "    M05_NENDO,"
        w_sSQL = w_sSQL & "    M05_GAKUNEN,"
        w_sSQL = w_sSQL & "    M05_CLASSNO,"
        w_sSQL = w_sSQL & "    M05_CLASSMEI"
        w_sSQL = w_sSQL & " FROM M05_CLASS"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "       M05_TANNIN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & "   AND M05_NENDO = " & cInt(m_iSyoriNen)

'response.write w_sSQL & "<br>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetClassInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            m_sGakunen  = rs("M05_GAKUNEN")
            m_sClassNo  = rs("M05_CLASSNO")
            m_sClassMei = rs("M05_CLASSMEI")
        End If

        '//����I��
        f_GetClassInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �s�����擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Get_HeadData()

    Dim w_sSQL
    Dim w_Rs
    Dim w_sSDate,w_sEDate

    On Error Resume Next
    Err.Clear
    
    f_Get_HeadData = 1

    Do 

        '// �s���w�b�_�f�[�^
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_MEI, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_KAISI_BI, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_SYURYO_BI, "
        w_sSQL = w_sSQL & vbCrLf & "  max(T32_GYOJI_M.T32_SOJIKANSU) AS T32_SOJIKANSU "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H, "
        w_sSQL = w_sSQL & vbCrLf & "  T32_GYOJI_M "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_CD = T32_GYOJI_M.T32_GYOJI_CD(+) AND "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_NENDO = T32_GYOJI_M.T32_NENDO(+) AND"
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_KYUKA_FLG='0' AND "   '//�����x��FLG (0:�ʏ� 1:�����x�� 2:�j��)
'2001/12/07 �u�����v��ǉ�
'        w_sSQL = w_sSQL & vbCrLf & "  T32_GYOJI_M.T32_COUNT_KBN='0' AND "   '//�J�E���g�敪(0:�s�� 1:���� 2:���̑�)
        w_sSQL = w_sSQL & vbCrLf & "  (T32_GYOJI_M.T32_COUNT_KBN='0' OR "   '//�J�E���g�敪(0:�s�� 1:���� 2:���̑� 3:����)
        w_sSQL = w_sSQL & vbCrLf & "  T32_GYOJI_M.T32_COUNT_KBN='3') AND "

        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  (T32_GYOJI_M.T32_TAISYO_GAKUNEN=" & cInt(m_sGakunen) & " Or T32_GYOJI_M.T32_TAISYO_GAKUNEN=" & C_GAKUNEN_ALL & ") AND "   '//�Ώۊw�N(0:�S�w�N 1-5:1-5�N )
        w_sSQL = w_sSQL & vbCrLf & "  (T32_GYOJI_M.T32_TAISYO_CLASS="   & cInt(m_sClassNo) & " Or T32_GYOJI_M.T32_TAISYO_CLASS=" & C_CLASS_ALL & ")"        '//�ΏۃN���X(99:�S�N���X)
        w_sSQL = w_sSQL & vbCrLf & " AND SUBSTR((T32_GYOJI_M.T32_HIDUKE),6,2)='" & gf_fmtZero(m_sTuki,2) & "'"
        w_sSQL = w_sSQL & vbCrLf & "  GROUP BY"
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_MEI, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_KAISI_BI, "
        w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_SYURYO_BI "

'response.write "<font color=#000000>" & w_sSQL & "<BR>"
        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then

            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_HeadData = 99
            Exit Do
        End If

        '//����I��
        f_Get_HeadData = 0
        Exit Do
    Loop

End Function

'****************************************************
'[�@�\] �f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'       (���X�g�_�E���{�b�N�X�I��\���p)
'[����] pData1 : �f�[�^�P
'       pData2 : �f�[�^�Q
'[�ߒl] f_Selected : "SELECTED" OR ""
'                   
'****************************************************
Function f_Selected(pData1,pData2)

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else 
            f_Selected = "" 
        End If
    End If

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
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>�s���o������</title>

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

        <%
        '//�S�C�N���X���Ȃ��ꍇ
        If m_sNoTanMsg <> "" Then%>
            parent.main.location.href="default2.asp?NoTanMsg=<%=m_sNoTanMsg%>"
        <%End If%>

    }

    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

        //�s�����Ȃ��ꍇ
        if (document.frm.GYOJI.value==""){
            //parent.main.location.href = "default2.asp"
            alert("�s��������܂���");
            return;
        };

        var vl = document.frm.GYOJI.value.split('_')

        document.frm.GYOJI_CD.value  = vl[0];
        document.frm.GYOJI_MEI.value = vl[1];
        document.frm.KAISI_BI.value  = vl[2];
        document.frm.SYURYO_BI.value = vl[3];
        document.frm.SOJIKANSU.value = vl[4];

		//���X�g��ʕ\��
        document.frm.action="./kks0140_bottom.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  ����ύX������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ChangeTuki(){

        //�{��ʂ�submit
        document.frm.target = "topFrame";
        document.frm.action = "./kks0140_top.asp"
        document.frm.submit();
        return;
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">

<%
'//�f�o�b�O
'Call s_DebugPrint()
%>

    <center>
    <%call gs_title("�s���o������","��@��")%>
    <%Do %>

        <%
        '//�S�C�N���X��񂪂Ȃ���
        If m_sGakunen = "" Or m_sClassMei = "" Then
            response.write "<span class=msg>" & m_sNoTanMsg & "<span>"
            Exit Do
        End If
        %>
		<br>
        <table >
        <tr><td class="search">
                    <table cellpadding="1" cellspacing="1">
                        <tr>
                        <td nowrap >�@<%=m_sGakunen%>�N&nbsp;&nbsp;<%=m_sClassMei%></td>
                        <td nowrap >�@<select name="TUKI" onchange="javascript:f_ChangeTuki();" style="width:50px;">
                                <option value="4"  <%=f_Selected("4" ,cstr(m_sTuki))%> >4
                                <option value="5"  <%=f_Selected("5" ,cstr(m_sTuki))%> >5
                                <option value="6"  <%=f_Selected("6" ,cstr(m_sTuki))%> >6
                                <option value="7"  <%=f_Selected("7" ,cstr(m_sTuki))%> >7
                                <option value="8"  <%=f_Selected("8" ,cstr(m_sTuki))%> >8
                                <option value="9"  <%=f_Selected("9" ,cstr(m_sTuki))%> >9
                                <option value="10" <%=f_Selected("10",cstr(m_sTuki))%> >10
                                <option value="11" <%=f_Selected("11",cstr(m_sTuki))%> >11
                                <option value="12" <%=f_Selected("12",cstr(m_sTuki))%> >12
                                <option value="1"  <%=f_Selected("1" ,cstr(m_sTuki))%> >1
                                <option value="2"  <%=f_Selected("2" ,cstr(m_sTuki))%> >2
                                <option value="3"  <%=f_Selected("3" ,cstr(m_sTuki))%> >3
                            </select></td>
						<td>��</td>
						<td>&nbsp;&nbsp;�s��</td>
                        <td nowrap  valign="middle" >

                        <%If m_Rs.EOF Then%>
                            <select name="GYOJI" style='width:200px;' DISABLED>
                                <option value="">�s��������܂���
                        <%Else%>
                            <select name="GYOJI" style='width:200px;'>
                            <%Do Until m_Rs.EOF%>
                                <option value=<%=m_Rs("T31_GYOJI_CD") & "_" & m_Rs("T31_GYOJI_MEI") & "_" & m_Rs("T31_KAISI_BI") & "_" & m_Rs("T31_SYURYO_BI") & "_" & m_Rs("T32_SOJIKANSU")%>>&nbsp;<%=m_Rs("T31_GYOJI_MEI")%>&nbsp;&nbsp;

                                <%m_Rs.MoveNext%>
                            <%Loop%>
                        <%End If%>

                            </select>
                        </td>
						<td valign="bottom" align="right">
			            <input class="button" type="button" onclick="javascript:f_Search();" value="�@�\�@���@">
						</tr>
                    </table>
		        </td>
	        </tr>
        </table>

        <%Exit Do%>
    <%Loop%>

    </center>

    <!--�l�n���p-->
    <INPUT TYPE=HIDDEN NAME="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE=HIDDEN NAME="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
    <INPUT TYPE=HIDDEN NAME="GAKUNEN"   value = "<%=m_sGakunen%>">
    <INPUT TYPE=HIDDEN NAME="CLASSNO"   value = "<%=m_sClassNo%>">
    <INPUT TYPE=HIDDEN NAME="GYOJI_CD"  value = "">
    <INPUT TYPE=HIDDEN NAME="GYOJI_MEI" value = "">
    <INPUT TYPE=HIDDEN NAME="KAISI_BI"  value = "">
    <INPUT TYPE=HIDDEN NAME="SYURYO_BI" value = "">
    <INPUT TYPE=HIDDEN NAME="SOJIKANSU" value = "">

    </form>
    </body>
    </html>
<%
    '---------- HTML END   ----------
End Sub
%>
