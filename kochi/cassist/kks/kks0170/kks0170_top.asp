<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����o������
' ��۸���ID : kks/kks0140/kks0140_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:cboDate        :���t
'            TUKI           :��
' ��      ��:
'           �������\��
'               ���A���̃R���{�{�b�N�X�͖{����\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��S�C�N���X�ꗗ��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/24 �ɓ����q
' ��      �X: 2001/07/30 �ɓ����q�@�����x�ɂ�\�����Ȃ��悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
	Const C_KYUKA_TYOUKI = 1	'//�����x���׸�(�����׸�)
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '��������
    Public m_iKyokanCd          '�N�x
    Public m_iTuki              '//��
    Public m_sDate              '//���t
    Public m_sDateWhere

    Public m_sAryDay()			'//�����x�ɂƓy���ȊO�̓��t
	Public m_iCnt				'//�����x�ɂƓy���ȊO�̓��t��

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
    w_sMsgTitle="�����o������"
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

        '//�R���{���t���擾(�����x�ɋy�ѓy���j�����������t)
        w_iRet = f_SetDay()
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
    m_iTuki = ""
    m_sDate = ""

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

    '//�������擾
    If request("TUKI") <> "" Then
        m_iTuki = request("TUKI")
    Else
        m_iTuki = month(date())
        m_sDate = gf_YYYY_MM_DD(date(),"/")
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
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"
    response.write "m_sClassMei = " & m_sClassMei & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  ���t�R���{�f�[�^���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_SetDay()

    Dim w_iRet
    Dim w_sSQL
    Dim rs
    Dim w_sSDate
    Dim w_sEDate

    On Error Resume Next
    Err.Clear

    f_SetDay = 1

    Do

        '//1�`3��
        If m_iTuki <= 3 Then
            w_iNen = cint(m_iSyoriNen)+1
        Else
            w_iNen = cint(m_iSyoriNen)
        End If

        '//���̌����������쐬
        '//�J�n��
        w_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iTuki),2) & "/01"

        '//�I����
        If Cint(m_sTuki) = 12 Then
            w_sEDate = cstr(w_iNen+1) & "/01/01"
        Else
            w_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_iTuki+1),2) & "/01"
        End If

        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & " A.T32_HIDUKE"
        w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M A"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & " A.T32_NENDO=" & cInt(m_iSyoriNen)
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE>='" & w_sSDate & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE< '" & w_sEDate & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='" & C_HEIJITU & "'"
        w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE"
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE"

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_SetDay = 99
            Exit Do
        End If

		i = 0

		If rs.EOF = False Then

			Do Until rs.EOF

				w_bKyuka = False
				w_sDate = rs("T32_HIDUKE")

				'//�擾�������t�������x�ɂ��ǂ���(�����x�Ɂcw_bKyuka = True)
				w_iRet = f_GetKyukaInfo(w_sDate,w_bKyuka)
				If w_iRet <> 0 Then
		            'ں��޾�Ă̎擾���s
		            f_SetDay = 99
		            Exit Function
				End If

				'//�����x�ɂłȂ����t���擾
				If w_bKyuka = False Then
					ReDim Preserve m_sAryDay(i)
					m_sAryDay(i) = w_sDate
					i = i + 1
				End If

				rs.MoveNext
			Loop

		End If

		m_iCnt = i-1

        '//����I��
        f_SetDay = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �擾�������t�������x�ɂ��ǂ���
'*  [����]  p_sDate  : ���t
'*  [�ߒl]  p_bKyuka : �����x��=True �����x�ɂłȂ� = False
'*  [����]  
'********************************************************************************
Function f_GetKyukaInfo(p_sDate,p_bKyuka)

    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_sGakunen,w_sClassNo

    On Error Resume Next
    Err.Clear

    f_GetKyukaInfo = 1
	p_bKyuka = False	'//�����x���׸�

    Do

		'//�S�C�N���X�����擾
		iRet = f_GetClassInfo(w_sGakunen,w_sClassNo)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            f_GetKyukaInfo = 99
            Exit Do
        End If

		'//�����x�ɂ��ǂ���
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  Count(T31_GYOJI_H.T31_GYOJI_CD) AS CNT"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H "
		w_sSQL = w_sSQL & vbCrLf & "  ,T32_GYOJI_M "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T31_GYOJI_H.T31_GYOJI_CD = T32_GYOJI_M.T32_GYOJI_CD "
		w_sSQL = w_sSQL & vbCrLf & "  AND T31_GYOJI_H.T31_NENDO = T32_GYOJI_M.T32_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND T31_GYOJI_H.T31_KYUKA_FLG='" & cstr(C_KYUKA_TYOUKI) & "' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T31_GYOJI_H.T31_NENDO=" & cInt(m_iSyoriNen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T32_GYOJI_M.T32_GAKUNEN IN (" & cint(w_sGakunen) & "," & cint(C_GAKUNEN_ALL) & ")"
		w_sSQL = w_sSQL & vbCrLf & "  AND T32_GYOJI_M.T32_CLASS IN (" & cint(w_sClassNo) & "," & cint(C_CLASS_ALL) & ") "
		w_sSQL = w_sSQL & vbCrLf & "  AND T32_GYOJI_M.T32_HIDUKE='" & p_sDate & "'"

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKyukaInfo = 99
            Exit Do
        End If

		'//�����x�Ƀf�[�^���擾�����ꍇ
		If cint(rs("CNT")) > 0 Then
			p_bKyuka = True
		End If

        '//����I��
        f_GetKyukaInfo = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  ����CD���A�S�C�N���X�����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetClassInfo(p_sGakunen,p_sClassNo)

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
			p_sGakunen  = rs("M05_GAKUNEN")
			p_sClassNo  = rs("M05_CLASSNO")
		End If

		'//����I��
		f_GetClassInfo = 0
		Exit Do
	Loop

    '//ں��޾��CLOSE
	Call gf_closeObject(rs)

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
    <title>�����o������</title>

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

    }

    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

		if(document.frm.cboDate.value==""){
			alert("�Ώۓ�������܂���");
			return;
		}

        document.frm.action="./kks0170_bottom.asp";
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
        document.frm.target = "_self";
        document.frm.action = "./kks0170_top.asp"
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
    <%call gs_title("�����o������","��@��")%>
    <table border="0" >
	    <tr>
		    <td colspan="2" align="center"><span class=CAUTION>�� �y���j���y�ђ����x�Ɏ��́A�o�^�ł��܂���</span></td>
	    </tr>
    <tr>
        <td class=search>
            <table border="0" cellpadding="0" cellspacing="0">
                <tr>
				<td>���͑Ώۓ�</td>
                <td nowrap >�@<select name="TUKI" onchange="javascript:f_ChangeTuki();" style="width:50px;">
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
                    </select></td>
				<td>��</td>
                </td>
                <td nowrap >

                    <%If m_iCnt < 0 Then%>
	                    <select name="cboDate"  DISABLED style="width:50px;">
                        <option value="">
                    <%Else%>
	                    <select name="cboDate"  style="width:50px;">
						<%For i = 0 To m_iCnt
                            %>
                            <option value="<%=m_sAryDay(i)%>" <%=f_Selected(m_sAryDay(i) ,m_sDate)%>><%=Day(m_sAryDay(i))%>
                            <%
                        Next
                    End If
                    %>
                    </select></td>
				<td>��</td>
				<td valign="bottom" align="right">
	            <input class="button" type="button" onclick="javascript:f_Search();" value="�@�\�@���@">
				</tr>
            </table>
        </td>
    </tr>
    </table>

    </center>
    </form>
    </body>
    </html>
<%
End Sub
%>
