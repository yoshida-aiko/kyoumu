<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �i�H���񌟍�
' ��۸���ID : mst/mst0133/syousai.asp
' �@      �\: ���y�[�W �A�E��}�X�^�̏ڍו\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtSinroCD      :�i�H�R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSinroName        :�i�H���́i�ꕔ�j
'           txtPageSinro        :�\���ϕ\���Ő��i�������g����󂯎������j
'           txtSentakuSinroCD       :�I�����ꂽ�i�H�R�[�h
'           txtSentakuSinroKbn       :�I�����ꂽ�i�H�敪
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           txtSinroCD          :�i�H�敪�i�߂�Ƃ��j
'           txtSingakuCd        :�i�w�R�[�h�i�߂�Ƃ��j
'           txtSinroName        :�i�H���́i�߂�Ƃ��j
'           txtPageSinro        :�\���ϕ\���Ő��i�߂�Ƃ��j
' ��      ��:
'           �������\��
'               �w�肳�ꂽ�i�w��E�A�E��̏ڍ׃f�[�^��\��
'           ���n�}�摜�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ��i�w��E�A�E���\������i�ʃE�B���h�E�j
'-------------------------------------------------------------------------
' ��      ��: 2001/06/21 �≺ �K��Y
' ��      �X: 2001/07/25 ���{�@����
'           : 2001/07/31 ���{�@����     �ϐ��������K���Ɋ�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_iSinroCD          ':�i�H�敪  '/2001/07/31�ύX
    Public  m_iSingakuCd        ':�i�w�敪
    Public  m_sSyusyokuName     ':�i�H���́i�ꕔ�j
    Public  m_iPageCD           ':�\���ϕ\���Ő��i�������g����󂯎������j'/2001/07/31�ύX
    Public  m_Rs                'recordset
    Public  m_iNendo            ':�N�x
    Public  m_sSentakuSinroCD   ':�R���{�{�b�N�X�őI�����ꂽ�i�HCD
    Public  m_sMode             ':���[�h
    Public  m_iSentakuSinroKbn  ':�I�����ꂽ�i�H�敪
    
    Public m_sKbn               ':�敪
    Public m_sSinromei          ':�i�H��
    Public m_sSinromeiKan       ':�i�H���i�J�i�j
    Public m_sSinromeiRya       ':�i�H���i���́j
    Public m_sJyusyo1           ':�Z���i�P�j
    Public m_sJyusyo2           ':�Z���i�Q�j
    Public m_sJyusyo3           ':�Z���i�R�j
    Public m_sTel               ':�i�H��d�b�ԍ�
    Public m_sYubin             ':�i�H��X�֔ԍ�
    Public m_sUrl               ':URL
    Public m_iGyosyuKbn         ':�Ǝ�敪
    Public m_iSihonkin          ':���{���i�P�ʁF���~�j
    Public m_iSihonkinY         ':���{���i�P�ʁF�~�j
    Public m_iJyugyoin_Suu      ':�]�ƈ���
    Public m_iSyoninkyu         ':���C��
    Public m_sBiko              ':���l
    Public m_iSinroKbn          ':�i�H�敪
    Public m_sKbnName			':��ʖ���

    'Public Const C_SYORYAKU_KETA = 4    '//�\�����ɏȗ����錅���i���{���j
    

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
    Dim w_sWHERE            '// WHERE��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�i�H���񌟍�"
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

        '// �l�̏�����
        Call s_SetBlank()
        
        '// ���Ұ�SET
        Call s_SetParam()

        '�A�E��}�X�^���擾
        w_sWHERE = ""

        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " M01_1.M01_NENDO "
        w_sSQL = w_sSQL & vbCrLf & " ,M01_1.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_NENDO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI_KANA "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRORYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO1 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO2 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JUSYO3 "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_YUBIN_BANGO "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINGAKU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_GYOSYU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SIHONKIN "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_JYUGYOIN_SUU "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SYONINKYU "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_BIKO "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
        w_sSQL = w_sSQL & vbCrLf & "    ,M01_KUBUN M01_1 "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "        M01_1.M01_NENDO = " & m_iNendo & ""
        w_sSQL = w_sSQL & vbCrLf & "    AND M01_1.M01_DAIBUNRUI_CD = " & C_SINRO & " "
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_NENDO = " & m_iNendo & ""
        w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN = M01_1.M01_SYOBUNRUI_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & "    AND M32_SINRO_CD = '" & m_sSentakuSinroCD & "' "

'Response.Write w_sSQL & "<br>"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        End If
        
        '//DB����l���擾
        Call s_SetDB()
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
'*  [�@�\]  �S�l��������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetBlank()
    
    m_sSentakuSinroCD = ""
    m_iNendo = ""
    m_iSinroCD = ""
    m_iSingakuCd = ""
    m_sMode = ""
    m_sSyusyokuName = ""
    m_iPageCD = ""
    
    m_sKbn = ""
    m_sSinromei = ""
    m_sSinromeiKan = ""
    m_sSinromeiRya = ""
    m_sJyusyo1 = ""
    m_sJyusyo2 = ""
    m_sJyusyo3 = ""
    m_sTel = ""
    m_sYubin = ""
    m_sUrl = ""
    m_iGyosyuKbn = ""
    m_iSihonkin = ""
    m_iJyugyoin_Suu = ""
    m_iSyoninkyu = ""
    m_sBiko = ""
    m_iSinroKbn = ""
    m_iSinroKbnY = ""
    m_iSentakuSinroKbn = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_sSentakuSinroCD = Request("txtSentakuSinroCD")    ':�i�H�R�[�h

    m_iNendo = Session("NENDO")     ':�N�x

    m_iSinroCD = Request("txtSinroCD")      ':�i�H�敪
    '�R���{���I����
    If m_iSinroCD="@@@" Then
        m_iSinroCD=""
    End If

    m_iSingakuCd = Request("txtSingakuCd")      ':�i�w�敪
    '�R���{���I����
    If m_iSingakuCd="@@@" Then
        m_iSingakuCd=""
    End If

    m_sMode = Request("txtMode")        ':���[�h

    m_sSyusyokuName = Request("txtSyusyokuName")    ':�A�E�於�́i�ꕔ�j

    '// BLANK�̏ꍇ�͍s���ر
    If Request("txtMode") = "Search" Then
        m_iPageCD = 1
    Else
        m_iPageCD = INT(Request("txtPageSyusyoku"))     ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
    
    m_iSentakuSinroKbn = CInt(Request("txtSentakuSinroKbn"))    ':�i�H�敪
    
End Sub

'********************************************************************************
'*  [�@�\]  DB����擾�����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetDB()

Dim w_iSihonkin

	if IsNull(m_Rs("M01_SYOBUNRUIMEI")) = False Then
	    m_sKbn = m_Rs("M01_SYOBUNRUIMEI")
	end if

	if IsNull(m_Rs("M32_SINROMEI")) = False Then
	    m_sSinromei = m_Rs("M32_SINROMEI")
	end if

	if IsNull(m_Rs("M32_SINROMEI_KANA")) = False Then
	    m_sSinromeiKan = m_Rs("M32_SINROMEI_KANA")
	end if
	if IsNull(m_Rs("M32_SINRORYAKSYO")) = False Then
	    m_sSinromeiRya = m_Rs("M32_SINRORYAKSYO")
	end if
	if IsNull(m_Rs("M32_JUSYO1")) = False Then
	    m_sJyusyo1 = m_Rs("M32_JUSYO1")
	end if
	if IsNull(m_Rs("M32_JUSYO2")) = False Then
	    m_sJyusyo2 = m_Rs("M32_JUSYO2")
	end if
	if IsNull(m_Rs("M32_JUSYO3")) = False Then
	    m_sJyusyo3 = m_Rs("M32_JUSYO3")
	end if
	if IsNull(m_Rs("M32_DENWABANGO")) = False Then
	    m_sTel = m_Rs("M32_DENWABANGO")
	end if
	if IsNull(m_Rs("M32_YUBIN_BANGO")) = False Then
	    m_sYubin = m_Rs("M32_YUBIN_BANGO")
	end if
	if IsNull(m_Rs("M32_SINRO_URL")) = False Then
	    m_sUrl = m_Rs("M32_SINRO_URL")
	end if

	if IsNull(m_Rs("M32_GYOSYU_KBN")) = False Then
	    m_iGyosyuKbn = m_Rs("M32_GYOSYU_KBN")
	end if

	if IsNull(m_Rs("M32_SIHONKIN")) = False Then
	    m_iSihonkinY = m_Rs("M32_SIHONKIN")
	    w_iSihonkin = CInt(Len(m_iSihonkinY)) - C_SYORYAKU_KETA
	    m_iSihonkin = Mid(m_iSihonkinY,1,w_iSihonkin)
	end if

	if IsNull(m_Rs("M32_JYUGYOIN_SUU")) = False Then
	    m_iJyugyoin_Suu = m_Rs("M32_JYUGYOIN_SUU")
	end if
	if IsNull(m_Rs("M32_SYONINKYU")) = False Then
	    m_iSyoninkyu = m_Rs("M32_SYONINKYU")
	end if
	if IsNull(m_Rs("M32_BIKO")) = False Then
	    m_sBiko = m_Rs("M32_BIKO")
	end if
	if IsNull(m_Rs("M32_SINRO_KBN")) = False Then
	    m_iSinroKbn = m_Rs("M32_SINRO_KBN")
	end if


	'//�i�H�敪OR�Ǝ�敪���̂��擾
	Select case cint(gf_SetNull2Zero(m_Rs("M32_SINRO_KBN")))
		Case C_SINRO_SINGAKU	'//�i�H�敪���i�w�̏ꍇ

			'//�i�w�敪���̂��擾
			w_sKbn = trim(m_Rs("M32_SINGAKU_KBN"))
			If w_sKbn <> "" Then
				Call gf_GetKubunName(C_SINGAKU,m_Rs("M32_SINGAKU_KBN"),m_iNendo,m_sKbnName)
			End If

		Case C_SINRO_SYUSYOKU	'//�i�H�敪���A�E�̏ꍇ

			'//�Ǝ�敪���̂��擾
			w_sKbn = trim(m_Rs("M32_GYOSYU_KBN"))
			If w_sKbn <> "" Then
				Call gf_GetKubunName(C_GYOSYU_KBN,m_Rs("M32_GYOSYU_KBN"),m_iNendo,m_sKbnName)
			End If

		Case C_SINRO_SONOTA	'//�i�H�敪�����̑��̏ꍇ
			m_sKbnName = ""
	End Select


End Sub

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

w_slink = "�@"

if m_Rs("M32_SINRO_URL") <> "" Then 
    w_sLink= "<a href='" & gf_HTMLTableSTR(m_sUrl) & "'>" 
    w_sLink= w_sLink &  gf_HTMLTableSTR(m_sUrl) & "</a>"
End if

        %>
        <%=w_slink%>
        <%
            m_Rs.MoveNext


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

    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="";
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageSinro.value = p_iPage;
        document.frm.submit();
    
    }

    function f_OpenWindow(p_Url){
    //************************************************************
    //  [�@�\]  �q�E�B���h�E���I�[�v������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
        var window_location;
        window_location=window.open(p_Url,"window","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,scrolling=no,Width=500,Height=500");
        window_location.focus();
    }

    //-->
    </SCRIPT>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    </head>

    <body>

    <center>

<%
m_sSubtitle = "�ځ@��"

call gs_title("�i�H���񌟍�",m_sSubtitle)
%>

    <table border=1 class=disp width="400">
        <tr>
            <td class=disph align="left" width="100">����</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sSinromei) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">���́i�J�i�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sSinromeiKan) %></td>
        </tr>

        <tr>
            <td class=disph align="left" width="100">����</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sSinromeiRya) %></td>
        </tr>

        <tr>
            <td class=disph align="left" width="100">�i�H�敪</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sKbn) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">��ʋ敪</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sKbnName) %></td>
        </tr>

        <tr>
            <td class=disph align="left" width="100">�X�֔ԍ�</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sYubin) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z���i�P�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sJyusyo1) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z���i�Q�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sJyusyo2) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�Z���i�R�j</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sJyusyo3) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�d�b�ԍ�</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sTel) %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">URL</td>
            <td class=disp align="left" width="300"><% S_syousai() %></td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">���{��</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_iSihonkin) %>���~</td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">�]�ƈ���</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_iJyugyoin_Suu) %>�l</td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">���C��</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_iSyoninkyu) %>�~</td>
        </tr>
        <tr>
            <td class=disph align="left" width="100">���l</td>
            <td class=disp align="left" width="300"><%=gf_HTMLTableSTR(m_sBiko) %></td>
        </tr>

    </table>


    <br>


    <table border="0">
    <tr>
    <td valign="top">
    <form name ="frm" action="./default.asp" target="<%=C_MAIN_FRAME%>">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtSinroCD" value="<%= m_iSinroCD %>">
        <input type="hidden" name="txtSingakuCD" value="<%= m_iSingakuCd %>">
        <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
        <input type="hidden" name="txtPageCD" value="<%= m_iPageCD %>">
    <input class=button type="submit" value="�߁@��">
    </form>
    </td>
    </tr>
    </table>


    </center>

    </body>

    </html>




<%
    '---------- HTML END   ----------
End Sub
%>