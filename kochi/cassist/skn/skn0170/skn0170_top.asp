<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������Ԋ�(�N���X��)
' ��۸���ID : skn/skn0170/skn0170_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:   NENDO           '//:�N
'               KYOKAN_CD       '//����CD
'               cboGakunenCd    '//�w�N
'               cboClassCd      '//�N���X
'               cboSikenKbn     '//�����敪
'               cboSikenCd      '//����CD
' ��      ��:
'           �������\��
'               ���݂̓��t�Ɉ�ԋ߂������ł̎��Ԋ��ꗗ��\������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/19 �ɓ����q
' ��      �X: 2001/08/02 ���{ ����  '�����R���{�\���ύX
'           : 2001/08/09 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Public Const C_FIRST_DISP_GAKUNEN = 1   '//�����\���̎��̊w�N(1�N)

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '//��������
    Public m_iKyokanCd          '//�N�x

    '//�R���{�pWhere������
    Public m_sGakunenWhere      '//�w�N�̏���
    Public m_sClassWhere        '//�N���X�̏���
    Public m_sClassOption       '//�N���X�R���{�̃I�v�V����
    Public m_sSikenWhere        '�����̏���
    Public m_sSikenOption       '�����R���{�̃I�v�V����
    Public m_sSikenCdWhere      '�����R���{�̃I�v�V�����i�����R�[�h�j

    Public m_iGakunen           '//�w�N
    Public m_iClassNo           '//�N���XNO
    Public m_iSikenKbn          '//�����敪
    Public m_sSikenCd           '//����CD
    Public m_sTxtMode           '//���샂�[�h
    Public m_sSikenBrank

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
    w_sMsgTitle="�������Ԋ�(�N���X��)"
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

        '//�w�N�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeGakunenWhere() 

        '//�N���X�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeClassWhere() 

        '//���݂̓��t�Ɉ�ԋ߂������敪���擾
        '//(�����\���͌��݂̓��t�Ɉ�ԋ߂������ł̎��Ԋ��ꗗ��\������)
        If m_sTxtMode = "" Then
'            w_iRet = f_Get_SikenKbn(m_iSikenKbn,m_sSikenCd)
            w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_JISSI_KIKAN,C_FIRST_DISP_GAKUNEN)
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If
        End If
        '//�����R���{�Ɋւ���WHERE���쐬����
        Call s_MakeSikenWhere() 
        
        '//�����R���{(�ǎ���)�Ɋւ���WHERE���쐬����
        Call s_MakeSikenCdWhere() 

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
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_iSikenKbn = ""
    m_sTxtMode = ""

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
    m_iGakunen  = Request("cboGakunenCd")   '//�w�N
    m_iClassNo  = Request("cboClassCd")     '//�N���X
    m_iSikenKbn = Request("cboSikenKbn")    '//�����敪
    m_sTxtMode  = Request("txtMode")
    m_sSikenCd  = Request("cboSikenCd")     '//��������

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
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_iSikenKbn = " & m_iSikenKbn & "<br>"
    response.write "m_sTxtMode  = " & m_sTxtMode  & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �w�N�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeGakunenWhere()

    m_sGakunenWhere = ""
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyorinen
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeClassWhere()

    m_sClassWhere = ""
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen

    If m_iGakunen = "" Then
        '//�����\������1�N1�g��\������
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & C_FIRST_DISP_GAKUNEN
    Else
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_iGakunen)
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �����敪�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeSikenWhere()

    m_sSikenWhere = ""
    m_sSikenWhere = m_sSikenWhere & " M01_NENDO = " & m_iSyorinen
    m_sSikenWhere = m_sSikenWhere & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
    m_sSikenWhere = m_sSikenWhere & " AND M01_SYOBUNRUI_CD <= 4 "						'<!--8/16�C��

End Sub

'********************************************************************************
'*  [�@�\]  �������ރR���{�Ɋւ���WHERE���쐬����i�����R�[�h�j
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeSikenCdWhere()

    m_sSikenCdWhere = ""
    m_sSikenOption = ""

    If cint(m_iSikenKbn) = Cint(C_SIKEN_JITURYOKU) or cint(m_iSikenKbn) = cInt(C_SIKEN_TUISI)  Then
        m_sSikenCdWhere = m_sSikenCdWhere & " M27_NENDO = " & m_iSyoriNen
        m_sSikenCdWhere = m_sSikenCdWhere & " AND M27_SIKEN_KBN = " & m_iSikenKbn

    else
        m_sSikenCdWhere = m_sSikenCdWhere & " M27_NENDO = " & m_iSyoriNen
        m_sSikenCdWhere = m_sSikenCdWhere & " AND M27_SIKEN_KBN = 99"
        m_sSikenOption = " DISABLED "

    End If
End Sub

'********************************************************************************
'*  [�@�\]  ���݂̓��t�Ɉ�ԋ߂������敪���擾(�����\�����̂�)
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �����\���͌��݂̓��t�Ɉ�ԋ߂������ł̎��Ԋ��ꗗ��\������
'********************************************************************************
Function f_Get_SikenKbn(p_iSiken_Kbn,p_sSiken_CD)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_Get_SikenKbn = 1
    p_iSiken_Kbn = ""
    p_sSiken_CD  = ""

    Do
        '���݂̓��t�Ɉ�ԋ߂������敪���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "    T24_SIKEN_KBN,"
        w_sSQL = w_sSQL & "    T24_SIKEN_CD"
        w_sSQL = w_sSQL & " FROM T24_SIKEN_NITTEI"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "       T24_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & "   AND T24_GAKUNEN = " & C_FIRST_DISP_GAKUNEN
        w_sSQL = w_sSQL & "   AND T24_JISSI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"
        w_sSQL = w_sSQL & " ORDER BY T24_JISSI_SYURYO ASC"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_SikenKbn = 99
            Exit Do
        End If

        If rs.EOF = False Then
            p_iSiken_Kbn = rs("T24_SIKEN_KBN")
            p_sSiken_CD  = rs("T24_SIKEN_CD")
        End If

        '//����I��
        f_Get_SikenKbn = 0
        Exit Do
    Loop

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
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>�������Ԋ�(�N���X��)</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
    }

    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

        document.frm.action="./skn0170_main.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �w�N���ύX���ꂽ�Ƃ��A�{��ʂ��ĕ\��
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./skn0170_top.asp";
        document.frm.target="top";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //-->
    </SCRIPT>

    </head>
    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">

<%
'//�f�o�b�O
'Call s_DebugPrint()
%>

    <center>
    <%call gs_title("�������Ԋ�(�N���X��)","��@��")%>
<br>
    <table bordeer="0">
        <tr>
        <td class="search">
            <table border="0">
            <tr>
            <td>
                <table border="0" cellpadding="1" cellspacing="1">
	                <tr valign="middle">

		                <td nowrap align="left">�N���X</td>
		                <td nowrap align="left">
		                    <% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere,"onchange = 'javascript:f_ReLoadMyPage()' style='width:40px;' ",False,m_iGakunen) %>
		                </td>
		                <td align="left">�N</td>
		                <td align="left" >
		                    <% call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere,"style='width:80px;' " & m_sClassOption,False,m_iClassNo) %>
		                </td>
		                <td align="left"><br></td>
	                </tr>
	                <tr>

		                <td nowrap align="left">����</td>
		                <td nowrap align="left" colspan="3">
		                    <% call gf_ComboSet("cboSikenKbn",C_CBO_M01_KUBUN,m_sSikenWhere,  "onchange = 'javascript:f_ReLoadMyPage()' style='width:160px;' ",False,m_iSikenKbn) %>
		                </td>
				        <td valign="bottom" clspan="1" align="right">
				        <input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
				        </td>
<!--
                <td nowrap align="left">
                    <% call gf_ComboSet("cboSikenCd" ,C_CBO_M27_SIKEN,m_sSikenCdWhere,m_sSikenOption & " style='width:120px;' ",True,m_sSikenCd) %>
                </td>
//-->
	                </tr>
                </table>
            </td>
            </tr>
            </table>
        </td>
        </tr>
    </table>
    </center>

    <!--�l�n���p-->
    <INPUT TYPE="HIDDEN" NAME="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE="HIDDEN" NAME="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">

    </form>
    </body>
    </html>
<%
    '---------- HTML END   ----------
End Sub
%>
