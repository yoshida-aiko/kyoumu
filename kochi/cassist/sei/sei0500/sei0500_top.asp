<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���͎������ѓo�^
' ��۸���ID : sei/sei0500/sei0500_top.asp
' �@      �\: ��y�[�W ���ѓo�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�N�x           ��      SESSION���i�ۗ��j
' ��      ��:

'-------------------------------------------------------------------------
' ��      ��: 2001/09/06 ���`�i�K
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    Public m_iNendo             '�N�x
    Public m_sKyokanCd          '�����R�[�h
    Public m_iSikenCD			'����CD

    Public m_Rs_Siken			'���������擾
    Public m_Rs					'�w�N�A�N���X�A�Ȗڎ擾RS

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

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���͎������ѓo�^"
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

		'//�l���擾
		call s_SetParam()

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

		'//�����R���{���擾
        w_iRet = f_GetSiken()
        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		'//�����R�[�h��NULL��������A�R���{�̂͂��߂̎����R�[�h������
		if gf_IsNull(m_iSikenCd) then m_iSikenCd = m_Rs_Siken("M28_SIKEN_CD")

		if Not gf_IsNull(m_iSikenCd) then

			'//�w�N�E�N���X�E�ȖڃR���{���擾
			w_iRet = f_GetKamoku()
			If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		End if

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
    Call gf_closeObject(m_Rs_Siken)
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

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
    m_iSikenCd  = Request("txtSikenCD")    '//�R���{�����敪

End Sub

'********************************************************************************
'*  [�@�\]  �����R���{���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSiken()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetSiken = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKENMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD "
		w_sSQL = w_sSQL & vbCrLf & "  FROM  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28_SIKEN_KAMOKU M28,  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD         = M27.M27_SIKEN_CD AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KAMOKU     = M27.M27_SIKEN_KAMOKU AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KBN        = M27.M27_SIKEN_KBN AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_NENDO            = M27.M27_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KBN        =  " & C_SIKEN_JITURYOKU & " AND "	'�����敪(���͎����̂�)
		w_sSQL = w_sSQL & vbCrLf & "  	(M28.M28_SEISEKI_KYOKAN1 = '" & m_sKyokanCd & "' OR "		'���͋���1
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN2 = '" & m_sKyokanCd & "' OR "		'���͋���2
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN3 = '" & m_sKyokanCd & "' OR "		'���͋���3
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN4 = '" & m_sKyokanCd & "' OR "		'���͋���4
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN5 = '" & m_sKyokanCd & "' ) AND "	'���͋���5
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_NENDO            =  " & m_iNendo					'�����N�x
		w_sSQL = w_sSQL & vbCrLf & "  GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKENMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD "
'Response.Write w_ssql & "<br>"
        iRet = gf_GetRecordset(m_Rs_Siken, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSiken = 99
            Exit Do
        End If

        f_GetSiken = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �w�N�E�N���X�E�ȖڃR���{���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKamoku()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetKamoku = 1

    Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD,"
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KAMOKU,"
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_GAKUNEN,  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_CLASS,  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_KAMOKUMEI "
		w_sSQL = w_sSQL & vbCrLf & "  FROM  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27, "
		w_sSQL = w_sSQL & vbCrLf & "  	M28_SIKEN_KAMOKU M28 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KAMOKU     = M28.M28_SIKEN_KAMOKU AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_CD         = M28.M28_SIKEN_CD AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KBN        = M28.M28_SIKEN_KBN AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO            = M28.M28_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO            =  " & m_iNendo & " AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_KBN        =  " & C_SIKEN_JITURYOKU & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M28.M28_SIKEN_CD         =  " & m_iSikenCD  & " AND  "
		w_sSQL = w_sSQL & vbCrLf & "  	(M28.M28_SEISEKI_KYOKAN1 = '" & m_sKyokanCd & "' OR "
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN2 = '" & m_sKyokanCd & "' OR "
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN3 = '" & m_sKyokanCd & "' OR "
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN4 = '" & m_sKyokanCd & "' OR "
		w_sSQL = w_sSQL & vbCrLf & "  	 M28.M28_SEISEKI_KYOKAN5 = '" & m_sKyokanCd & "' ) "

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKamoku = 99
            Exit Do
        End If

        f_GetKamoku = 0
        Exit Do
    Loop

End Function

'****************************************************
'[�@�\] �f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'[����] pData1 : �f�[�^�P
'       pData2 : �f�[�^�Q
'[�ߒl] f_Selected : "SELECTED" OR ""
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

    On Error Resume Next
    Err.Clear

%>
	<html>
	<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//  [�@�\]  �������ύX���ꂽ�Ƃ��A�ĕ\������
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_ReLoadMyPage(){

	    document.frm.action="sei0500_top.asp";
	    document.frm.target="topFrame";
	    document.frm.submit();

	}

	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���̏���
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_Search(){

	    // ������NULL����������
	    // ���w�N
	    if( f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
	        window.alert("�w�N�̑I�����s���Ă�������");
	        document.frm.txtGakuNo.focus();
	        return ;
	    }
	    // ���N���X
	    if( f_Trim(document.frm.txtClassNo.value) == "<%=C_CBO_NULL%>" ){
	        window.alert("�N���X�̑I�����s���Ă�������");
	        document.frm.txtClassNo.focus();
	        return ;
	    }

	    // ���Ȗږ�
	    if( f_Trim(document.frm.txtKamokuCd.value) == "<%=C_CBO_NULL%>" ){

			if (document.frm.txtKamokuCd.length ==1){
		        window.alert("�����Ȗڂ�����܂���");
		        return ;
			}else{
		        window.alert("�Ȗڂ̑I�����s���Ă�������");
		        document.frm.txtKamokuCd.focus();
		        return ;
			}
	    }

		// �I�����ꂽ�R���{�̒l���
		iRet = f_SetData();
		if( iRet != 0 ){
			return;
		}

	    document.frm.action="sei0500_bottom.asp";
	    document.frm.target="main";
	    document.frm.submit();

	}

	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���ɑI�����ꂽ�f�[�^���
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_SetData(){

        if (document.frm.cboKamoku.value==""){
            alert("�Ȗڃf�[�^������܂���B")
            return 1;
        };

		//�f�[�^�擾
        var vl = document.frm.cboKamoku.value.split('$$$');

        //�I�����ꂽ�f�[�^���(�w�N�A�N���X�A�Ȗ�CD���擾)
        document.frm.txtGakuNo.value=vl[0];
        document.frm.txtClassNo.value=vl[1];
        document.frm.txtKamokuCd.value=vl[2];

        return 0;
	}

	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>

	<body>
	<center>
	<form name="frm" METHOD="post">

	<% call gs_title(" ���͎������ѓo�^ "," �o�@�^ ") %>
	<br>
	<table border="0">
	    <tr><td valign="bottom">

	        <table border="0" width="100%">
	            <tr><td class="search">

	                <table border="0">
	                    <tr valign="middle">
	                        <td align="left" nowrap>�����敪</td>
	                        <td align="left" colspan="3">
								<%If m_Rs_Siken.EOF Then%>
									<select name="txtSikenCD" style='width:150px;' DISABLED>
										<option value="">�f�[�^������܂���
								<%Else%>
									<select name="txtSikenCD" style='width:150px;' onchange = 'javascript:f_ReLoadMyPage()'>
									<%Do Until m_Rs_Siken.EOF%>
										<option value='<%=m_Rs_Siken("M28_SIKEN_CD")%>'  <%=f_Selected(cstr(m_Rs_Siken("M28_SIKEN_CD")),cstr(m_iSikenCD))%>><%=m_Rs_Siken("M27_SIKENMEI")%>
										<%m_Rs_Siken.MoveNext%>
									<%Loop%>
								<%End If%>
								</select>
							</td>
	                        <td>&nbsp;</td>

	                        <td align="left" nowrap>�Ȗ�</td>
	                        <td align="left">
								<%If m_iSikenCd = "" Then%>
									<select name="cboKamoku" style='width:200px;' DISABLED>
										<option value="">�f�[�^������܂���
								<%Else%>
									<%If m_Rs.EOF Then%>
										<select name="cboKamoku" style='width:200px;' DISABLED>
											<option value="">�Ȗڃf�[�^������܂���
									<%Else%>
										<select name="cboKamoku" style='width:200px;'>
										<%Do Until m_Rs.EOF%>
											<%
											wSikenCd   = m_Rs("M28_SIKEN_CD") 

											'//�\�����e���쐬
											w_Str=""
											w_Str= w_Str & m_Rs("M28_GAKUNEN") & "�N�@"
											w_Str= w_Str & gf_GetClassName(m_iNendo,m_Rs("M28_GAKUNEN"),m_Rs("M28_CLASS")) & "�@"
											w_Str= w_Str & m_Rs("M27_KAMOKUMEI") & "�@"
											%>
											<option value="<%=m_Rs("M28_GAKUNEN")%>$$$<%=m_Rs("M28_CLASS")%>$$$<%=m_Rs("M28_SIKEN_KAMOKU")%>"><%=w_Str%>
											<%m_Rs.MoveNext%>
										<%Loop%>
									<%End If%>
								<%End If%>
								</select>
							</td>
	                    </tr>
						<tr>
					        <td colspan="7" align="right">
					        <input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
					        </td>
						</tr>
	                </table>
	            </td>
				</tr>
	        </table>
	        </td>
	    </tr>
	</table>

	<input type="hidden" name="txtNendo"     value="<%= m_iNendo %>">
	<input type="hidden" name="txtKyokanCd"  value="<%= m_sKyokanCd %>">
	<input type="hidden" name="txtShikenCd"  value="<%= wSikenCd   %>">

	<input type="hidden" name="txtGakuNo"    value="">
	<input type="hidden" name="txtClassNo"   value="">
	<input type="hidden" name="txtKamokuCd"  value="">

	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>