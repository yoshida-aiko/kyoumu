<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������ԋ����\��ꗗ
' ��۸���ID : skn/skn0180/skn0180_top.asp
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
' ��      ��: 2001/07/23 �ɓ����q
' ��      �X: 2001/08/02 ���{ ����  '�����R���{�\���ύX
'           : 2001/08/08 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'           : 2001/08/09 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '//��������
    Public m_iKyokanCd          '//�N�x

    '//�R���{�pWhere������
    Public m_sSikenWhere        '�����̏���
    Public m_sSikenOption       '�����R���{�̃I�v�V����
    Public m_sSikenCdWhere      '�����R���{�̃I�v�V�����i�����R�[�h�j
    Public m_sKyokanName        '//��������
    Public m_iSikenKbn          '//�����敪
    Public m_sSikenCd           '//����CD
    Public m_sTxtMode           '//���샂�[�h

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
    w_sMsgTitle="�������ԋ����\��ꗗ"
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

        '//�������̂��擾
        w_iRet = f_GetKyokanNm(m_iKyokanCd,m_iSyoriNen,m_sKyokanName)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '//���݂̓��t�Ɉ�ԋ߂������敪���擾
        '//(�����\���͌��݂̓��t�Ɉ�ԋ߂������ł̎��Ԋ��ꗗ��\������)
        If m_sTxtMode = "" Then
            w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_JISSI_KIKAN,0)
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
    m_iSikenKbn = Request("cboSikenKbn")    '//�����敪
    m_sTxtMode  = Request("txtMode")

    '//����CD
    If Request("SKyokanCd1") = "" Then
        m_iKyokanCd = Session("KYOKAN_CD")
    Else
        m_iKyokanCd = Request("SKyokanCd1")
    End If

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
    response.write "m_iSikenKbn = " & m_iSikenKbn & "<br>"
    response.write "m_sTxtMode  = " & m_sTxtMode  & "<br>"

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
        m_sSikenOption = " DISABLED "

    End If
End Sub

'********************************************************************************
'*  [�@�\]  ���݂̓��t�Ɉ�ԋ߂������敪���擾
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

'********************************************************************************
'*  [�@�\]  �����̎������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sName
'*  [����]  
'********************************************************************************
Function f_GetKyokanNm(p_sCD,p_iNENDO,p_sName)
Dim rs
Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetKyokanNm = 1
    w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M04_KYOKAN_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M04_NENDO = " & p_iNENDO & " "

'response.write w_sSQL

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKyokanNm = 99
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M04_KYOKANMEI_SEI") & "�@" & rs("M04_KYOKANMEI_MEI")
        End If

        f_GetKyokanNm = 0
        Exit Do
    Loop

    p_sName = w_sName

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
    <title>�������ԋ����\��ꗗ</title>

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

        if(document.frm.SKyokanNm1.value==""){
			alert("������I�����Ă�������")
			return;
		}

        document.frm.action="./skn0180_main.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �����R���{���ύX���ꂽ�Ƃ��A�{��ʂ��ĕ\��
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ReLoadMyPage(){

        document.frm.action="./skn0180_top.asp";
        document.frm.target="top";
        document.frm.txtMode.value = "Reload";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �N���A�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function fj_Clear(){

		document.frm.SKyokanNm1.value = "";
		document.frm.SKyokanCd1.value = "";

	}

    //************************************************************
    //  [�@�\]  �����Q�ƑI����ʃE�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function KyokanWin(p_iInt,p_sKNm) {
		var obj=eval("document.frm."+p_sKNm)

        URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=610,top=0,left=0");
        nWin.focus();
        return true;    
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <form name="frm" method="post">

<%
'//�f�o�b�O
'Call s_DebugPrint()
%>

    <center>
    <%call gs_title("�������ԋ����\��ꗗ","��@��")%>
<br>
    <table>
        <tr>
        <td class="search">
            <table border="0">
            <tr>
            <td>
                <table border="0" cellpadding="1" cellspacing="1">
                <tr valign="middle">
                <td align="left" nowrap>
                ����
                </td>
                <td align="left" nowrap>
                    <% call gf_ComboSet("cboSikenKbn",C_CBO_M01_KUBUN,m_sSikenWhere,  "onchange = 'javascript:f_ReLoadMyPage()' style='width:150px;' ",False,m_iSikenKbn) %>
                </td>
<!--
                <td align="left" nowrap>
                    <% call gf_ComboSet("cboSikenCd" ,C_CBO_M27_SIKEN,m_sSikenCdWhere,m_sSikenOption & " style='width:120px;' ",True,"") %>
                </td>
//-->
                </tr>
                <tr valign="middle">
                <td align="left" nowrap>
                ����
                </td>
                <td align="left" nowrap colspan="2">
                    <input type="text" class="text" name="SKyokanNm1" VALUE='<%=m_sKyokanName%>' readonly>
                    <input type="hidden" name="SKyokanCd1" VALUE='<%=m_iKyokanCd%>'>
                    <input type="button" class="button" value="�I��" onclick="KyokanWin(1,'SKyokanNm1')">
					<input type="button" class="button" value="�N���A" onClick="fj_Clear()">
                    <!--<input type="button" class="button" value="�I��" onclick="KyokanWin(1,'<%=m_sKyokanName%>')">-->
                </td>
                </tr>
                </table>
            </td>
            </tr>
			<tr>
		        <td valign="bottom" clspan="1" align="right">
		        <input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
		        </td>
			</tr>
            </table>
        </td>
        </tr>
    </table>

    </center>

    <!--�l�n���p-->
    <INPUT TYPE="HIDDEN" NAME="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
    </form>
    </body>
    </html>
<%
    '---------- HTML END   ----------
End Sub
%>
