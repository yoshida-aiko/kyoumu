<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A�E��}�X�^
' ��۸���ID : mst/mst0133/main.asp
' �@      �\: ���y�[�W �A�E��}�X�^�̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H��R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSyusyokuName     :�A�E�於�́i�ꕔ�j
'           txtPageCD       :�\���ϕ\���Ő��i�������g����󂯎������j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtRenrakusakiCD    :�I�����ꂽ�A����R�[�h
'           txtPageCD       :�\���ϕ\���Ő��i�������g�Ɉ����n�������j
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��A�E�E�i�w���\��
'           �����ցA�߂�{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ��A�E�E�i�w��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/06/18 �≺�@�K��Y
' ��      �X: 2001/07/13 �J�e�@�ǖ�
' ��      �X: 2001/08/22 �ɓ��@���q �Ǝ�敪�ǉ��Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_iNendo            '�����N�x
    Public  m_sSinroCD      ':�i�H��R�[�h
    Public  m_sSingakuCd        ':�i�w�R�[�h
    Public  m_sSyusyokuName     ':�A�E�於�́i�ꕔ�j
    Public  m_sPageCD       ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_skubun
    Public  m_Rs            'recordset
    Public  m_iDisp         ':�\�������̍ő�l���Ƃ�
    Public  m_sMode
    '�y�[�W�֌W
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp          '// �ꗗ�\���s��

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
    w_sMsgTitle="�A�E�}�X�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    '// ���Ұ�SET
    Call s_SetParam()

        If m_sMode = "" Then
        '// �y�[�W��\��
        Call NoPage()
    Else
        
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

            '�A�E�}�X�^���擾
            w_sWHERE = ""

            w_sSQL = w_sSQL & vbCrLf & " SELECT "
            w_sSQL = w_sSQL & vbCrLf & " M01.M01_SYOBUNRUIMEI "
            w_sSQL = w_sSQL & vbCrLf & " ,M01.M01_NENDO "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_CD "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINROMEI "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_DENWABANGO "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_URL "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINRO_KBN "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_SINGAKU_KBN "
            w_sSQL = w_sSQL & vbCrLf & " ,M32.M32_GYOSYU_KBN "
            w_sSQL = w_sSQL & vbCrLf & " FROM "
            w_sSQL = w_sSQL & vbCrLf & "    M32_SINRO M32 "
            w_sSQL = w_sSQL & vbCrLf & "    ,M01_KUBUN M01 "
            w_sSQL = w_sSQL & vbCrLf & " WHERE "
            w_sSQL = w_sSQL & vbCrLf & "    M01_NENDO = " & m_iNendo & " AND "
            w_sSQL = w_sSQL & vbCrLf & "    M32_NENDO = " & m_iNendo & " AND "

'---2001/08/22 ito �Ǝ�敪�ǉ��Ή�
	         w_sSQL = w_sSQL & vbCrLf & "    M01.M01_DAIBUNRUI_CD = "&C_SINRO&""
	         w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN = M01.M01_SYOBUNRUI_CD "

            '���o�����̍쐬
            If m_sSinroCD<>"" Then
                w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINRO_KBN =" & m_sSinroCD & " "
            End If

'---2001/08/22 ito �Ǝ�敪�ǉ��Ή�
	        If m_sSingakuCd <> "" Then
				if cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SINGAKU then
		            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINGAKU_KBN =" & m_sSingakuCd & " "
				ElseIf cint(gf_SetNull2Zero(m_sSinroCD)) = C_SINRO_SYUSYOKU then
		            w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_GYOSYU_KBN =" & m_sSingakuCd & " "
				End if
	        End If

            If m_sSyusyokuName<>"" Then
                w_sSQL = w_sSQL & vbCrLf & "    AND M32.M32_SINROMEI Like '%" & m_sSyusyokuName & "%' "
            End If

            w_sSQL = w_sSQL & vbCrLf & " ORDER BY M32.M32_SINRO_CD "

'   Response.Write w_sSQL & "<br>"

            Set m_Rs = Server.CreateObject("ADODB.Recordset")
            w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)

            If w_iRet <> 0 Then
                'ں��޾�Ă̎擾���s
                m_bErrFlg = True
                Exit Do 'GOTO LABEL_MAIN_END
            Else
                '�y�[�W���̎擾
                m_iMax = gf_PageCount(m_Rs,m_iDsp)
'   Response.Write "m_iMax:" & m_iMax & "<br>"
            End If

                If m_Rs.EOF Then
                '// �y�[�W��\��
                Call showPage_NoData()
            Else
                '// �y�[�W��\��
                Call showPage()
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
    End If

End Sub


'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()
    m_sSinroCD = Request("txtSinroCD")              ':�i�H�R�[�h
    If m_sSinroCD="@@@" Then m_sSinroCD=""      '�R���{���I����

    m_sSingakuCd = Request("txtSingakuCd")          ':�i�w�R�[�h
    If m_sSingakuCd="@@@" Then m_sSingakuCd=""  '�R���{���I����

    m_sMode = Request("txtMode")            ':���[�h

    m_iNendo = Session("NENDO")     ':�N�x
    m_sSyusyokuName = Request("txtSyusyokuName")    ':�A�E�於�́i�ꕔ�j

    If Request("txtPageCD") <> "" Then
        m_sPageCD = INT(Request("txtPageCD"))   ':�\���ϕ\���Ő��i�������g����󂯎������j
    Else
        m_sPageCD = 1   ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If
    If m_sPageCD = 0 Then m_sPageCD = 1

    If m_sSinroCD = "1" Then            ':�w�b�_�[�̋敪���̕ύX
        m_skubun = "�i�w�敪"
    else
        m_skubun = "�i�H�敪"
    End If
    
    m_iDisp = C_PAGE_LINE       '�P�y�[�W�ő�\����

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
Dim w_i
Dim w_cell
w_iCnt  = 1
w_i     = 0
w_cell = ""

Do While not m_Rs.EOF
	w_i = w_i + 1

	w_slink = "�@"

	if m_Rs("M32_SINRO_URL") <> "" Then 
	    w_sLink= "<a href='" & gf_HTMLTableSTR(m_Rs("M32_SINRO_URL")) & "' target='_blank'>" 
	    w_sLink= w_sLink &  gf_HTMLTableSTR(trim(m_Rs("M32_SINRO_URL"))) & "</a>"
	End if
	call gs_cellPtn(w_cell)
        %>

        <tr>
        <td align="center" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M01_SYOBUNRUIMEI")) %></td>

		<%
		'//������
		w_sKbn = ""
		w_sKbnName = ""

		'//�i�H�敪OR�Ǝ�敪���̂��擾
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

        <td align="center" class=<%=w_cell%>><%=gf_HTMLTableSTR(w_sKbnName) %></td>

        <td align="left" class=<%=w_cell%>><%=gf_HTMLTableSTR(trim(m_Rs("M32_SINROMEI"))) %></a></td>
        <td align="left" class=<%=w_cell%>><%=gf_HTMLTableSTR(m_Rs("M32_DENWABANGO")) %></td>
        <td align="left" class=<%=w_cell%>><%=w_slink%></td>
        <td align="center" class=<%=w_cell%>><input class=button type="button" value=">>" onclick="javascript:f_Henko('<%=cstr(m_Rs("M32_SINRO_CD")) %>')"></td>
        <td align="center" class=<%=w_cell%>><input type="checkbox" name="deleteNO<%= w_i %>" value="<%=gf_HTMLTableSTR(m_Rs("M32_SINRO_CD")) %>"></td>
        </tr>

        <%
            m_Rs.MoveNext
            If w_iCnt >= C_PAGE_LINE Then
                Exit Sub
            Else
                w_iCnt = w_iCnt + 1
            End If
        Loop

        m_iDisp = w_i

End sub



Sub NoPage()
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

    </head>

    <body>

	<center>
	<br><br><br>
	<span class="msg"><%=C_BRANK_VIEW_MSG%></span>
	</center>

    </body>

    </html>

<%
    '---------- HTML END   ----------
End Sub


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
    
    On Error Resume Next
    Err.Clear

%>

<html>
    <head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
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
        document.frm.target="";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageCD.value = p_iPage;
        document.frm.submit();
    
    }
   

    //************************************************************
    //  [�@�\]  �C����ʂ�\������
    //  [����]  p_sSinroCD
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_GoSyosai(p_sSinroCD){

        document.frm.action="syousai.asp";
        document.frm.target="";
        document.frm.txtPageCD.value = p_sSinroCD;
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  syosai_frm�ւ̃p�����[�^�̎󂯓n��
    //  [����]  p_sSyuseiCD
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Henko(p_sSyuseiCD){

        document.frm.action="syusei.asp";
        document.frm.target="fTopMain";
        document.frm.txtRenrakusakiCD.value = p_sSyuseiCD;
        document.frm.txtMode.value = "Syusei";
        document.frm.submit();
    }


    //************************************************************
    //  [�@�\]  �폜�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Delete(){

    MainFrm = parent.window.frames["main"]
    var i;
    i = 1;

    var checkFlg
        checkFlg=false

    do { 
        obj = eval("document.frm.deleteNO" + i);
        if(obj.checked == true){
        
            checkFlg=true
            break;
         }

    i++; }  while(i<=document.frm.txtDisp.value);
    if (checkFlg == false){
        alert( "�폜�̑ΏۂƂȂ�i�H���I������Ă��܂���" );
    }else{

        document.frm.action="./del_kakunin.asp";
        document.frm.target="fTopMain";
        document.frm.txtMode.value = "Delete";
        document.frm.submit();
        }
    }


    //-->
    </SCRIPT>

    </head>

<body>
<center>

<form name="frm" action="" target="" method="post">
<br>
<table><tr><td align="center" width="800">
	<%
		'�y�[�WBAR�\��
		Call gs_pageBar(m_Rs,m_sPageCD,m_iDsp,w_pageBar)
	%>
	<%=w_pageBar %>

    <table border=1 class=hyo width="100%">
	    <tr>
		    <th class=header width="80">�i�H�敪</th>
		    <th class=header width="80">��ʋ敪</th>
		    <th class=header>�i�@�H�@��</th>
		    <th class=header width="96">�s �d �k</th>
		    <th class=header width="30%">�t �q �k</th>
		    <th class=header width="32">�C��</th>
		    <th class=header width="32">�폜</th>
	    </tr>
	    <% S_syousai() %>
	    <tr>
		    <td colspan=7 align=right bgcolor=#9999BD><input class=button type=button value="�~�폜" Onclick="f_Delete()"></td>
	    </tr>
	</table>

	<%=w_pageBar %>
</td></tr></tabel>

<br>
</center>

<input type="hidden" name="txtMode" value="">
<input type="hidden" name="txtRenrakusakiCD" value="">
<input type="hidden" name="txtSinroCD2" value="<%= m_sSinroCD %>">
<input type="hidden" name="txtSingakuCD2" value="<%= m_sSingakuCd %>">
<input type="hidden" name="txtSinroCD" value="<%= m_sSinroCD %>">
<input type="hidden" name="txtSingakuCD" value="<%= m_sSingakuCd %>">
<input type="hidden" name="txtSyusyokuName" value="<%= m_sSyusyokuName %>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
<input type="hidden" name="txtNendo" value="<%= Session("SYORI_NENDO") %>">
<input type="hidden" name="txtDisp" value="<%= m_iDisp %>">
</form>

</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>