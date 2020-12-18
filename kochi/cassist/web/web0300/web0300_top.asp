<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ʋ����\��
' ��۸���ID : web/web0300/web0300_top.asp
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
'               �A�N�Z�X������FULL�̏ꍇ�́A���p�҂�ύX�ł���
'-------------------------------------------------------------------------
' ��      ��: 2001/08/06 �ɓ����q
' ��      �X: 2001/08/07 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////

'	Const C_ACCESS_FULL   = "FULL"		'//�A�N�Z�X����FULL�A�N�Z�X��
'	Const C_ACCESS_NORMAL = "NORMAL"	'//�A�N�Z�X�������
'	Const C_ACCESS_VIEW   = "VIEW"		'//�A�N�Z�X�����Q�Ƃ̂�

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '//��������
    Public m_iKyokanCd          '//�N�x
    Public m_iTuki
    Public m_sLoginId

    '//�R���{�pWhere������
    Public m_sKyosituWhere      '//�����擾����
    Public m_sKyokanName        '//��������
    Public m_sKengen			'//�A�N�Z�X����

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
    w_sMsgTitle="���ʋ����\��"
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

        '//�����R���{�Ɋւ���WHERE���쐬����
        Call s_MakeKyosituWhere() 

		'//�������擾
		w_iRet = gf_GetKengen_web0300(m_sKengen)
		If w_iRet <> 0 Then
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
	m_sLoginId  = ""
    m_iTuki     = ""

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
	m_sLoginId  = trim(Session("LOGIN_ID"))
    m_iTuki     = month(date())

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
    response.write "m_sLoginId  = " & m_sLoginId  & "<br>"
    response.write "m_iTuki     = " & m_iTuki     & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeKyosituWhere()

    m_sKyosituWhere = ""
    m_sKyosituWhere = m_sKyosituWhere & "     M06_NENDO = " & m_iSyorinen
	'�g�p�t���O���g�p�̂��̂����\�� 2001/12/11 
    m_sKyosituWhere = m_sKyosituWhere & " AND M06_SIYO_FLG = '1'"

End Sub

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
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>���ʋ����\��</title>
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
    function f_Show(){

        if(document.frm.SKyokanNm1.value==""){
			alert("���p�҂�I�����Ă�������")
			return;
		}

        document.frm.action="./web0300_main.asp";
        document.frm.target="main";
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

		//����ʂ��󔒕\��
		parent.main.location.href="default2.asp"

		//���p�җ����󔒂ɂ���
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

		//����ʂ��󔒕\��
		parent.main.location.href="default2.asp"

		var obj=eval("document.frm."+p_sKNm)

		//���p�ґI����ʂ�\������
        //URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
        URL = "../../Common/com_select/Sel_User/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
//2015.8.19 UPDATE URAKAWA �\���T�C�Y��ύX
//        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=600,top=0,left=0");
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=570,height=640,top=0,left=0");
        nWin.focus();
        return true;    
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE="javascript" onload="return window_onload()">
    <%call gs_title("���ʋ����\��","��@��")%>

    <form name="frm" method="post" onClick="return false;">
<%
'//�f�o�b�O
'Call s_DebugPrint()
%>

    <center>
        <table border="0">
        <tr>
        <td class="search">
            <table border="0" cellpadding="1" cellspacing="1">
            <tr>
            <td align="left">

                <table border="0" cellpadding="1" cellspacing="1">
                <tr>
                <td Nowrap align="left">����</td>
                <td Nowrap align="left">
                    <% call gf_ComboSet("cboKyositu",C_CBO_M06_KYOSITU,m_sKyosituWhere,"style='width:220px;' ",False,m_iSikenKbn) %>
                </td>
                <td Nowrap align="left">
                    <select name="TUKI">
                        <option value="4"  <%=f_Selected("4" ,cstr(m_iTuki))%> >4
                        <option value="5"  <%=f_Selected("5" ,cstr(m_iTuki))%> >5
                        <option value="6"  <%=f_Selected("6" ,cstr(m_iTuki))%> >6
                        <option value="7"  <%=f_Selected("7" ,cstr(m_iTuki))%> >7
                        <option value="8"  <%=f_Selected("8" ,cstr(m_iTuki))%> >8
                        <option value="9"  <%=f_Selected("9" ,cstr(m_iTuki))%> >9
                        <option value="10" <%=f_Selected("10",cstr(m_iTuki))%> >10
                        <option value="11" <%=f_Selected("11",cstr(m_iTuki))%> >11
                        <option value="12" <%=f_Selected("12",cstr(m_iTuki))%> >12
                        <option value="1"  <%=f_Selected("1" ,cstr(m_iTuki))%> >1
                        <option value="2"  <%=f_Selected("2" ,cstr(m_iTuki))%> >2
                        <option value="3"  <%=f_Selected("3" ,cstr(m_iTuki))%> >3
                    </select>��</td>
                </tr>

                <tr>
                <td Nowrap align="left">���p��</td>
                <td align="left" nowrap colspan="1">
                    <input type="text" class="text" name="SKyokanNm1" VALUE='<%=gf_GetUserNm(m_iSyoriNen,m_sLoginId)%>' readonly>
                    <input type="hidden" name="SKyokanCd1" VALUE='<%=m_sLoginId%>'>
					<%
					'//�ō������҂̂ݗ��p�҂̕ύX���Ƃ���
					If m_sKengen = C_ACCESS_FULL Then%>
	                    <input type="button" class="button" value="�I��" onclick="KyokanWin(1,'SKyokanNm1')">
						<input type="button" class="button" value="�N���A" onClick="fj_Clear()">
					<%End If%>
                </td>
                <td align="right" nowrap colspan="1">
		        <input class="button" type="button" value="�@�\�@���@" onclick="javascript:f_Show()">
                </td>
                </tr>
                </table>
            </td>
            </tr>
            </table>
        </td>
        </tr>
        </table>

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>