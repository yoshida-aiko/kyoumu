<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �i�H���񌟍�
' ��۸���ID : mst/mst0133/main.asp
' �@      �\: ���y�[�W �A�E��}�X�^�̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtSinroCD      :�i�H�R�[�h
'           txtSingakuCD        :�i�w�R�[�h
'           txtSinroName        :�A�E�於�́i�ꕔ�j
'           txtPageSyusyoku     :�\���ϕ\���Ő��i�������g����󂯎������j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           txtSinroCD      :�i�H�R�[�h             '/2001/07/31�ǉ�
'           txtSingakuCD        :�i�w�R�[�h         '/2001/07/31�ǉ�
'           txtSinroName        :�A�E�於�́i�ꕔ�j
'           txtSentakuSinroCD   :�I�����ꂽ�i�H�R�[�h
'           txtSentakuSinroKbn   :�I�����ꂽ�i�H�敪
'           txtPageSyusyoku     :�\���ϕ\���Ő��i�������g�Ɉ����n�������j
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��A�E�E�i�w���\��
'           �����ցA�߂�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ��A�E�E�i�w��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/18 �≺�@�K��Y
' ��      �X: 2001/07/31 ���{ ����  �����E���n�ǉ�
'           :                       �ϐ��������K���Ɋ�ύX
'           : 2001/08/10 ���{ ����  NN�Ή��ɔ����\�[�X�ύX
'           : 2001/08/22 �ɓ� ���q  ����SQL���ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_iSinroCD          ':�i�H�R�[�h        '/2001/07/31�ύX
    Public  m_iSingakuCd        ':�i�w�R�[�h        '/2001/07/31�ύX
    Public  m_sSyusyokuName     ':�A�E�於�́i�ꕔ�j
    Public  m_iPageCD           ':�\���ϕ\���Ő��i�������g����󂯎������j'/2001/07/31�ύX
    Public  m_skubun            ':�敪����
    Public  m_Rs                'recordset
    Public  m_iNendo            ':�N�x
    Public  m_sMode             ':���[�h
    Public  m_iFLG              ':
    Public  m_sSNm              ':
    'Public  m_sSinroKBN        ':�i�H�敪
    Public  m_iSinroKbn         ':�i�H�敪�R�[�h
    

    '�y�[�W�֌W
    Public  m_iMax              ':�ő�y�[�W
    Public  m_iDsp                      '// �ꗗ�\���s��

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
    m_iDsp = C_PAGE_LINE

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

        '�A�E�}�X�^���擾
        w_sWHERE = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " 	M01.M01_SYOBUNRUIMEI "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINRO_CD "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINROMEI "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_DENWABANGO "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINRO_URL "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINRO_KBN "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_SINGAKU_KBN "
        w_sSQL = w_sSQL & vbCrLf & " 	,M32.M32_GYOSYU_KBN"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & " 	M32_SINRO M32, "
        w_sSQL = w_sSQL & vbCrLf & " 	("
        w_sSQL = w_sSQL & vbCrLf & " 	select * "
        w_sSQL = w_sSQL & vbCrLf & " 	from "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & " 	where "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_DAIBUNRUI_CD  = " & C_SINRO & " and "
        w_sSQL = w_sSQL & vbCrLf & " 		M01_NENDO = " & m_iNendo
        w_sSQL = w_sSQL & vbCrLf & " 	) M01"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & " 	M32.M32_SINRO_KBN = M01.M01_SYOBUNRUI_CD (+) and "
		w_sSQL = w_sSQL & vbCrLf & " 	M32.M32_NENDO = M01.M01_NENDO (+) and "
		w_sSQL = w_sSQL & vbCrLf & " 	M32.M32_NENDO = " & m_iNendo & ""
		
        '���o�����̍쐬
        'If m_sSinroKBN <> "" Then
        If m_iSinroCD <> "" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN =" & m_iSinroCD & " "
        End If
        
        If m_iSingakuCd <> "" Then
			if m_iSinroCD = 1 then
	            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN =" & m_iSingakuCd & " "
			ElseIf m_iSinroCD = 2 then
	            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_GYOSYU_KBN =" & m_iSingakuCd & " "
			End if
        End If
        
        If m_sSyusyokuName<>"" Then
            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINROMEI Like '%" & m_sSyusyokuName & "%' "
        End If

        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M32.M32_SINRO_CD "
		
		Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 'GOTO LABEL_MAIN_END
        Else
            '�y�[�W���̎擾
            m_iMax = gf_PageCount(m_Rs,m_iDsp)
        End If

        If m_Rs.EOF Then
            '// �y�[�W��\��
            Call showPage_NoData()
        Else

            If m_iFLG = "1" Then
                '// �y�[�W��\��
                Call showPage_SHOW()
            Else
                '// �y�[�W��\��
                Call showPage()
            End If
        End If
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo = Session("NENDO")         ':�N�x

    m_iSinroCD = Request("txtSinroCD")      ':�i�H�敪
    '�R���{���I����
    If m_iSinroCD = "@@@" Then
        m_iSinroCD = ""
    End If

    m_iSingakuCd = Request("txtSingakuCd")      ':�i�w�敪
    '�R���{���I����
    If m_iSingakuCd="@@@" Then
        m_iSingakuCd=""
    End If

    m_sMode = Request("txtMode")            ':���[�h

    m_sSyusyokuName = Request("txtSyusyokuName")    ':�A�E�於�́i�ꕔ�j

    If m_sMode = "Search" Then
        m_iPageCD = 1
    Else
        m_iPageCD = INT(Request("txtPageSyusyoku")) ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If

'    If cstr(gf_SetNull2String(m_iSinroCD)) = "1" Then            ':�w�b�_�[�̋敪���̕ύX
'        m_skubun = "�i�w�敪"
'	ElseIf cstr(gf_SetNull2String(m_iSinroCD)) = "2" Then
'        m_skubun = "�Ǝ�敪"
'    else
'        'm_skubun = "�i�H�敪"
'        m_skubun = "��ʋ敪"
'    End If

    m_iDisp = C_PAGE_LINE       '�P�y�[�W�ő�\����

    m_iFLG = request("txtFLG")
    m_sSNm = request("txtSNm")

	if gf_IsNull(request("txtSNm")) then
		m_sSNm = request("SearchNm")
	End if

    m_iSinroKbn = ""
    
End Sub


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

If m_iFLG <> "1" Then

    Do While not m_Rs.EOF

    w_slink = "�@"
    m_iSinroKbn = m_Rs("M32_SINRO_KBN")

    if m_Rs("M32_SINRO_URL") <> "" Then 
        'w_sLink= "<a href='" & gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "'>" 
        w_sLink= "<a href='" & m_Rs("M32_SINRO_URL") & "' target='_site'>" 
        w_sLink= w_sLink &  m_Rs("M32_SINRO_URL") & "</a>"
    End if

        '//�e�[�u���Z���w�i�F
        call gs_cellPtn(w_cellT)
        %>
        <tr>

		<%
		'//������
		w_sKbn = ""
		w_sKbnName = ""

		'//�i�H�敪OR�Ǝ�敪���擾
		Select case cint(gf_SetNull2Zero(m_Rs("M32_SINRO_KBN")))
			Case C_SINRO_SINGAKU	'//�i�H�敪���i�w�̏ꍇ

				'//�i�w�敪���̂��擾
				w_sKbn = trim(m_Rs("M32_SINGAKU_KBN"))
				If w_sKbn <> "" Then
					Call gf_GetKubunName(C_SINGAKU,m_Rs("M32_SINGAKU_KBN"),m_iNendo,w_sKbnName)
				End If

			Case C_SINRO_SYUSYOKU	'//�i�H�敪���A�E�̏ꍇ

				'//�Ǝ�敪���̂��擾
				w_sKbn = trim(m_Rs("M32_GYOSYU_KBN"))
				If w_sKbn <> "" Then
					Call gf_GetKubunName(C_GYOSYU_KBN,m_Rs("M32_GYOSYU_KBN"),m_iNendo,w_sKbnName)
				End If

			Case C_SINRO_SONOTA	'//�i�H�敪�����̑��̏ꍇ

		End Select

		%>

        <td align="left" class=<%=w_cellT%>><%=m_Rs("M01_SYOBUNRUIMEI") %></td>
        <td align="left" class=<%=w_cellT%>><%=w_sKbnName%></td>
        <td align="left" class=<%=w_cellT%>><a href="javascript:f_GoSyosai('<%=m_Rs("M32_SINRO_CD") %>','<%=m_iSinroKbn%>')"><%=trim(m_Rs("M32_SINROMEI")) %></a></td>
        <td align="left" class=<%=w_cellT%>><%=m_Rs("M32_DENWABANGO") %></td>
        <td align="left" class=<%=w_cellT%>><%=w_slink%></td>
        </tr>
        <%
        m_Rs.MoveNext

        If w_iCnt >= C_PAGE_LINE Then
            Exit Do
        Else
            w_iCnt = w_iCnt + 1
        End If
    Loop

Else 

    Do While not m_Rs.EOF
        Call gs_cellPtn(w_cell)

        %>
        <tr>
        <td align="left" class=<%=w_cell%>><%=m_Rs("M01_SYOBUNRUIMEI") %></td>
        <td align="left" class=<%=w_cell%>>
        <input type=button class=<%=w_cell%> name="SinroNm_<%=w_iCnt%>" value='<%=m_Rs("M32_SINROMEI") %>' onclick="iinSelect(<%=w_iCnt%>)">
        <input type=hidden name="SinroCd_<%=w_iCnt%>" value='<%=m_Rs("M32_SINRO_CD") %>'>
        </td>
        <td align="left" class=<%=w_cell%>><%=m_Rs("M32_DENWABANGO") %></td>
        </tr>
        <%
        m_Rs.MoveNext

        If w_iCnt >= C_PAGE_LINE Then
            Exit Do
        Else
            w_iCnt = w_iCnt + 1
        End If
    Loop

End If

    'LABEL_showPage_OPTION_END
End sub

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
   </head>

    <body>

    <center>
        <br><br><br>
        <span class="msg">�Ώۃf�[�^�͑��݂��܂���B��������͂��Ȃ����Č������Ă��������B</span>
    </center>

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
    Dim w_pageBar           '�y�[�WBAR�\���p

    Dim w_iRecordCnt        '//���R�[�h�Z�b�g�J�E���g

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

    '�y�[�WBAR�\��
    Call gs_pageBar(m_Rs,m_iPageCD,m_iDsp,w_pageBar)

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
        document.frm.txtPageSyusyoku.value = p_iPage;
        document.frm.submit();
    
    }
    
    //************************************************************
    //  [�@�\]  �ڍ׃y�[�W��\��
    //  [����]  p_sSinroCD:�i�H�R�[�h
    //          p_sSinroKbn:�i�H�敪
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_GoSyosai(p_sSinroCD,p_sSinroKbn){

        document.frm.action="syousai.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.txtSentakuSinroCD.value = p_sSinroCD;
        document.frm.txtSentakuSinroKbn.value = p_sSinroKbn;
        document.frm.txtMode.value = "Search";
        document.frm.submit();
    
    }
    //-->
    </SCRIPT>
    <link rel=stylesheet href="../../common/style.css" type=text/css>
    </head>

    <body>

    <center>
<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr><td align="center">
<br>
<span class=CAUTION>�� �i�H�����N���b�N����Əڍׂ��m�F�ł��܂��B</span>
<%=w_pageBar %>

        <table border=1 class=hyo width="100%">
        <COLGROUP WIDTH="15%">
        <COLGROUP WIDTH="15%">
        <COLGROUP WIDTH="30%">
        <COLGROUP WIDTH="25%">
        <COLGROUP WIDTH="30%">
        <tr>
        <th class=header>�i�H�敪</th>
        <th class=header>��ʋ敪</th>
        <th class=header>�i�H��</th>
        <th class=header>TEL</th>
        <th class=header>URL</th>
        </tr>

    <% S_syousai() %>

        </table>

<%=w_pageBar %>

</td></tr></table>

    <br>

    <table border="0">
    <tr>
    <td valign="top">
    <form name ="frm" action="" target="">
        <input type="hidden" name="txtMode" value="<%=m_sMode%>">
        <input type="hidden" name="txtSinroCD" value="<%= m_iSinroCD %>">
        <input type="hidden" name="txtSingakuCD" value="<%= m_iSingakuCd %>">
        <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
        <input type="hidden" name="txtPageSyusyoku" value="<%= m_iPageCD %>">
        <input type="hidden" name="txtSentakuSinroCD" value="">
        <input type="hidden" name="txtSentakuSinroKbn" value="">
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

Sub showPage_SHOW()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    Dim w_pageBar           '�y�[�WBAR�\���p

    Dim w_iRecordCnt        '//���R�[�h�Z�b�g�J�E���g

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

    '�y�[�WBAR�\��
    Call gs_pageBar(m_Rs,m_iPageCD,m_iDsp,w_pageBar)

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


        document.frm.action="main.asp";
        document.frm.target="main";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageSyusyoku.value = p_iPage;
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �\�����e�\���p�E�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function iinSelect(p_sct) {

        //�}�����̃t�H�[�����擾
            w_sctNm = eval("document.frm.SinroNm_"+p_sct);
            w_sctNo = eval("document.frm.SinroCd_"+p_sct);

        //�}������
            parent.opener.document.frm.SinroNm.value = w_sctNm.value;
            parent.opener.document.frm.SinroCd.value = w_sctNo.value;

            document.frm.SearchNm.value = w_sctNm.value;
            document.frm.SearchNo.value = w_sctNo.value;

        return true;    
        //window.close()

    }

    //************************************************************
    //  [�@�\]  �N���A�{�^�����N���b�N�����ꍇ
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function f_Clear(p_No) {

        document.frm.SearchNm.value = "";
        document.frm.SearchNo.value = "";

        //�}�����������t�H�[�����擾
            w_NmStr = parent.opener.document.frm.SinroNm;
            w_NoStr = parent.opener.document.frm.SinroCd;

        //�}������

            w_NmStr.value = document.frm.SearchNm.value;
            w_NoStr.value = document.frm.SearchNo.value;
        return true;    
    }
    //-->
    </SCRIPT>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    </head>

    <body>
    <center>
    <form name="frm" method="post">

    <table width="90%" border="0">
        <tr>
            <td align="center">
                <table width="80%" class="hyo">
                    <tr>
                        <td align="center" width="30%"><font color="white">�i�@�H�@��</font></td>
                        <td align="center" class="detail"><input type="text" class="noBorder" name="SearchNm" value="<%=m_sSNm%>" readonly><input type="hidden" name="SearchNo" value=""></td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <span class="CAUTION">�� �I��������ɂ͐i�H�����N���b�N���Ă��������B</span>
    <table border="0" align="center">
    <tr>
    <td valign="top">
        <input type="button" value=" �N���A " class="button" onclick="javascript:f_Clear()">
        <input type="button" value="����" class="button" onclick="javascript:parent.window.close()">
    </td>
    </tr>
    </table>

                <%=w_pageBar %>
                <table border="1" class="hyo" width="100%">
                    <tr>
                        <th class="header" width="10%" nowrap>�i�H�敪</th>
                        <th class="header" width="50%">�i�H��</th>
                        <th class="header" width="40%">TEL</th>
                    </tr>
                    <% S_syousai() %>
                </table>
                <%=w_pageBar %>
            </td>
        </tr>
    </table>

    <table border="0" align="center">
    <tr>
    <td valign="top">
        <input type="button" value=" �N���A " class="button" onclick="javascript:f_Clear()">
        <input type="button" value="����" class="button" onclick="javascript:parent.window.close()">
    </td>
    </tr>
    </table>

	    <input type="hidden" name="txtMode" value="<%=m_sMode%>">
	    <input type="hidden" name="txtSinroCD" value="<%= m_iSinroCD %>">
	    <input type="hidden" name="txtSingakuCD" value="<%= m_iSingakuCd %>">
	    <input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
	    <input type="hidden" name="txtPageSyusyoku" value="<%= m_iPageCD %>">
	    <input type="hidden" name="txtSentakuSinroCD" value="">
	    <input type="hidden" name="txtSentakuSinroKbn" value="">
	    <input type="hidden" name="txtFLG" value="<%=m_iFLG%>">
    </form>

    </center>
    </body>
    </html>



<%
    '---------- HTML END   ----------
End Sub
%>