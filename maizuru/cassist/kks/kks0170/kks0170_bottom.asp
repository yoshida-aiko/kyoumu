<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����o������
' ��۸���ID : kks/kks0170/kks0170_bottom.asp
' �@      �\: ���y�[�W ���Əo�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��: SESSION("NENDO")           '//�����N
'             SESSION("KYOKAN_CD")       '//����CD
'             TUKI           '//��
'             cboDate        '//���t
' ��      ��:
' ��      �n: NENDO"        '//�����N
'             KYOKAN_CD     '//����CD
'             GAKUNEN"      '//�w�N
'             CLASSNO"      '//�׽No
'             cboDate"      '//���t
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��S�C�׽���k����\��
'           ���o�^�{�^���N���b�N��
'               ���͏���o�^����
'-------------------------------------------------------------------------
' ��      ��: 2001/07/24 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
    Const C_SYOBUNRUICD_IPPAN = 4   '//���ȋ敪(0:�o��,1:����,2:�x��,3:����,4:����,�c)

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bTannin       '//�S�C�׸�

    '�擾�����f�[�^�����ϐ�
    Public m_iSyoriNen      '//�����N�x
    Public m_iKyokanCd      '//����CD
    Public m_sDate          '//���t
    Public m_iGakunen       '//�w�N
    Public m_iClassNo       '//�N���XNo
    Public m_sClassNm       '//�N���X����
    Public m_iRsCnt         '//�N���Xں��ސ�
    Public m_sDispMsg       '//�G���[�����b�Z�[�W
	Public m_sEndDay		'//���͂ł��Ȃ��Ȃ��

    'ں��ރZ�b�g
    Public m_Rs

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
    w_sMsgTitle="�����o������"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_bTannin = False

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

        '// �S�C�N���X���擾
        w_iRet = f_GetClassInfo(m_bTannin)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'���͕s�ɂȂ�����擾
		call gf_Get_SyuketuEnd(m_iGakunen,m_sEndDay)

        '//���O�C���������S�C�N���X�������Ă���Ƃ����k���X�g���擾
        If m_bTannin = True Then
            '// ���k���X�g���擾
            w_iRet = f_GetClassList()
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If

			If m_Rs.EOF Then
		        '// �f�[�^�Ȃ��̏ꍇ�A�󔒃y�[�W��\��
		        Call showWhitePage("�N���X��񂪂���܂���")
			Else
		        '// �ڍ׃y�[�W��\��
		        Call showPage()
			End If
		Else
	        '// �S�C�N���X�������Ă��Ȃ��ꍇ�A�󔒃y�[�W��\��
	        Call showWhitePage("�󎝃N���X������܂���B")
        End If

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gf_closeObject(m_Rs)
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
    m_sDate     = ""
    m_iGakunen  = ""
    m_iClassNo  = ""
    m_sClassNm = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = SESSION("NENDO")
    m_iKyokanCd = SESSION("KYOKAN_CD")

    m_sDate     = trim(Request("cboDate"))
    If m_sDate = "" Then
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
    response.write "m_sDate     = " & m_sDate     & "<br>"
    response.write "m_iGakunen  = " & m_iGakunen  & "<br>"
    response.write "m_iClassNo  = " & m_iClassNo  & "<br>"
    response.write "m_sClassNm  = " & m_sClassNm  & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �S�C�N���X���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_GetClassInfo(p_bTannin)

    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear
    
    f_GetClassInfo = 1

    Do 

        '// �S�C�N���X���
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_TANNIN"
        w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & m_iKyokanCd & "'"

'response.write w_sSQL & "<BR>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetClassInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            p_bTannin = True 
            m_iGakunen = rs("M05_GAKUNEN")
            m_iClassNo = rs("M05_CLASSNO")
            m_sClassNm = rs("M05_CLASSMEI")
        End If

        f_GetClassInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �S�C�N���X�ꗗ�擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_GetClassList()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_GetClassList = 1

    Do 

        '// �S�C�N���X���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_GAKUNEN, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  A.T13_IDOU_NUM, "
        w_sSQL = w_sSQL & vbCrLf & "  B.T11_SIMEI, "
        w_sSQL = w_sSQL & vbCrLf & "  B.T11_GAKUSEI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  C.T30_HIDUKE, "
        w_sSQL = w_sSQL & vbCrLf & "  C.T30_SYUKKETU_KBN,"
        '//"�o��"�͕\�����Ȃ�
        w_sSQL = w_sSQL & vbCrLf & "  DECODE(D.M01_SYOBUNRUIMEI_R,'�o','�@',D.M01_SYOBUNRUIMEI_R) AS SYUKKETU_MEI"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN A"
        w_sSQL = w_sSQL & vbCrLf & "  ,T11_GAKUSEKI B"
        w_sSQL = w_sSQL & vbCrLf & "  ,(SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     T30_HIDUKE,"
        w_sSQL = w_sSQL & vbCrLf & "     T30_SYUKKETU_KBN,"
        w_sSQL = w_sSQL & vbCrLf & "     T30_GAKUSEKI_NO"
        w_sSQL = w_sSQL & vbCrLf & "    FROM T30_KESSEKI"
        w_sSQL = w_sSQL & vbCrLf & "    WHERE T30_HIDUKE='" & m_sDate & "'"
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_GAKUNEN=" & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "      AND T30_CLASS=" & m_iClassNo & ") C"
        w_sSQL = w_sSQL & vbCrLf & "  ,(SELECT "
        w_sSQL = w_sSQL & vbCrLf & "     M01_SYOBUNRUI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "     M01_SYOBUNRUIMEI_R"
        w_sSQL = w_sSQL & vbCrLf & "    FROM M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & "    WHERE "
        w_sSQL = w_sSQL & vbCrLf & "          M01_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "      AND M01_DAIBUNRUI_CD=" & C_KESSEKI & ") D"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        'w_sSQL = w_sSQL & vbCrLf & "      A.T13_NENDO - A.T13_GAKUNEN + 1 = B.T11_NYUNENDO(+) "
        w_sSQL = w_sSQL & vbCrLf & "      A.T13_GAKUSEI_NO = B.T11_GAKUSEI_NO "
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_GAKUSEKI_NO = C.T30_GAKUSEKI_NO(+)"
        w_sSQL = w_sSQL & vbCrLf & "  AND C.T30_SYUKKETU_KBN = D.M01_SYOBUNRUI_CD(+)"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_NENDO=" & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_GAKUNEN=" & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T13_CLASS=" & m_iClassNo
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T13_GAKUSEKI_NO"

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(m_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetClassList = 99
            Exit Do
        End If

        '//ں��ރJ�E���g���擾
        m_iRsCnt = gf_GetRsCount(m_Rs)

        f_GetClassList = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �ړ�����̏ꍇ�ړ��󋵂̎擾
'*  [����]  p_Gakusei_No:�w��NO
'*          p_Date      :���Ǝ��{��
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Get_IdouInfo(p_Gakusei_No)

    Dim w_sSQL
    Dim w_Rs
    Dim w_sKubunName

    On Error Resume Next
    Err.Clear

    w_IdoFlg = False

    Do

        '// �ړ����
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_1, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_1, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_2, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_2, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_3, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_3, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_4, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_4, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_5, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_5, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_6, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_6, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_7, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_7, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_8, "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_8"
        w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO='" & p_Gakusei_No & "' AND"
        w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM>0"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            Exit Do
        End If

        If w_Rs.EOF = false Then

            i = 1
            Do Until i>8    '//8�c�ő�ړ���

                If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) = "" Then
                    Exit Do
                End If

                If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) > m_sDate  Then
                    Exit Do
                End If
                i = i + 1
            Loop

            If i = 1 then
                '//�ŏ��̈ړ��������Ɠ���薢���̏ꍇ�A���Ɠ��Ɉړ���Ԃł͂Ȃ�
                w_sKubunName = ""
            Else
                '//�ړ����̏ꍇ�A�ړ��敪�E�ړ����R���擾
                Select Case Trim(w_Rs("T13_IDOU_KBN_" & i-1))
                 Case cstr(C_IDO_FUKUGAKU),cstr(C_IDO_TEI_KAIJO)  '//C_IDO_FUKUGAKU=3:���w�AC_IDO_TEI_KAIJO=5:��w����
                    w_sKubunName = ""
                 Case Else
                    '//�ړ����R���擾(�敪�}�X�^�A�啪��=C_IDO)
                    w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i-1)),m_iSyoriNen,w_sKubunName)
                    If w_bRet<> True Then
                        Exit Do
                    End If
                End Select

            End If

        End If

        Exit Do
    Loop

    f_Get_IdouInfo = w_sKubunName

    Call gf_closeObject(w_Rs)

    Err.Clear

End Function

'********************************************************************************
'*  [�@�\]  �o���敪�Ɩ��̂��擾(javascript�����p)
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  �o�����͂�JAVASCRIPT�쐬
'********************************************************************************
Function f_Get_SYUKETU_KBN(p_MaxNo)

    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear
    
    f_Get_SYUKETU_KBN = 1

    Do 

        '// ���׃f�[�^
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI_R"
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_KESSEKI) & " AND "
        '//C_SYOBUNRUICD_IPPAN = 4  '//���ȋ敪(0:�o��,1:����,2:�x��,3:����,4:����,�c)
        w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD<" & C_SYOBUNRUICD_IPPAN
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY M01_KUBUN.M01_SYOBUNRUI_CD"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_SYUKETU_KBN = 99
            Exit Do
        End If

        i=0
        If rs.EOF = True Then
            response.write ("var ary = new Array(0);")
            response.write ("ary[0] = '�@';")
        Else

            '//ں��ރJ�E���g�擾
            w_iCnt = gf_GetRsCount(rs) - 1
            response.write ("var ary = new Array(" & w_iCnt & ");") & vbCrLf

            Do Until rs.EOF
                If i = 0 Then
                    response.write ("ary[0] = '�@';") & vbCrLf
                Else
                    response.write ("ary[" & rs("M01_SYOBUNRUI_CD") &  "] = '" & rs("M01_SYOBUNRUIMEI_R") & "';") & vbCrLf
                End If

                i=i+1
                rs.MoveNext
            Loop

        End If

        p_MaxNo = w_iCnt

        f_Get_SYUKETU_KBN = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)
    Err.Clear

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
Dim i 
Dim w_sIduoRiyu

    On Error Resume Next
    Err.Clear

%>
    <html>
    <head>
    <title>�����o������</title>
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

        //�w�b�_����\��submit
        document.frm.txtMsg.value = "<%=m_sDispMsg%>";
        document.frm.target = "topFrame";
        document.frm.action = "kks0170_middle.asp"
        document.frm.submit();
        return;

    }

    //************************************************************
    //  [�@�\]  �o������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function chg(chgInp) {

        no = 0;
        <%
        '//�o���敪���擾
        Call f_Get_SYUKETU_KBN(w_MaxNo)
        %>

        str = chgInp.value;
        for(i=0; i<<%=w_MaxNo+1%>; i++){
            if (ary[i]==str){
                break;
            }
        };

        no = i + 1;
        if (no > <%=w_MaxNo%>) no = 0;
        chgInp.value = ary[no];

        //�B���t�B�[���h�Ƀf�[�^���Z�b�g
        var obj=eval("document.frm.hid"+chgInp.name);
        obj.value=no;
        return;
    }
    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Touroku(){

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		//�w�b�_���󔒃y�[�W�\��
		parent.topFrame.location.href="./white.htm"

        //���X�g����submit
        document.frm.target = "main";
        //document.frm.target = "<%=C_MAIN_FRAME%>";
        document.frm.action = "./kks0170_edt.asp"
        document.frm.submit();
        return;
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
        //parent.document.location.href="default2.asp"

        document.frm.target = "<%=C_MAIN_FRAME%>";
        document.frm.action = "./default.asp"
        document.frm.submit();
        return;

    }

    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">

<%
'//�f�o�b�O
'Call s_DebugPrint()
%>
    <center>
    <form name="frm" method="post" onClick="return false;">

    <%Do
        '//�S�C�N���X���Ȃ��ꍇ
        If m_bTannin = False Then
            Exit Do
        End If

        If m_Rs.EOF = False Then
            '//���X�g���s�J�E���g���擾
            w_iCnt = INT(m_iRsCnt/2 + 0.9)
        Else
            '//�f�[�^�Ȃ��̏ꍇ
            Exit Do
        End If%>

        <!--���X�g��-->
        <table >
            <tr><td valign="top">

                <!--�w�b�_-->
                <table class=hyo border="1" bgcolor="#FFFFFF">

        <%
		'���͊��Ԃ��߂��Ă���΁A���͂͂ł��Ȃ��B
		if m_sEndDay < m_sDate then 
				w_tmp =  " onclick='return chg(this)'"
		else
				w_tmp = " DISABLED"
		End If

        i=1
        Do Until m_Rs.EOF

                    '//���ټ�Ă̸׽���Z�b�g
                    Call gs_cellPtn(w_Class) 

                    '//���͋��E�񋖉̔���
                    '//�ړ��󋵂̍l��(T13_IDOU_NUM��1�ȏ�̏ꍇ�͈ړ��󋵂𔻕ʂ���)�ړ����̏ꍇ�́A�o�����͕s��
                    If gf_SetNull2Zero(m_Rs_D("T13_IDOU_NUM")) > 0 Then 
                        w_sIduoRiyu = ""
                        w_sIduoRiyu = f_Get_IdouInfo(m_Rs("T11_GAKUSEI_NO"))
                    End If

                    %>

                    <!--�ڍ�-->
                    <tr>
                        <td nowrap class="<%=w_Class%>" width="80"  align="center"><%=m_Rs("T13_GAKUSEKI_NO")%><input type="hidden" name="GAKUSEKI_NO" value="<%=m_Rs("T13_GAKUSEKI_NO")%>"><br></td>
                        <td nowrap class="<%=w_Class%>" width="150" align="left"><%=m_Rs("T11_SIMEI")%><br></td>
                        <%
                        If w_sIduoRiyu <> "" Then
                            %>
                            <td class="NOCHANGE" width="80" align="center" ><%=gf_SetNull2String(w_sIduoRiyu)%><br>
                            <input type="hidden" name="hidKBN<%=m_Rs("T13_GAKUSEKI_NO")%>" size="2" value="---"></td>
                            <%
                        Else
                            %>
                            <td class="<%=w_Class%>" width="80" align="center">
                            <input type="button" name="KBN<%=m_Rs("T13_GAKUSEKI_NO")%>" value="<%=gf_HTMLTableSTR(m_Rs("SYUKKETU_MEI"))%>" size="20" maxlength="2" class=<%=w_Class%> style="border-style:none" style="text-align:center" <%=w_tmp%>>
                            <input type="hidden" name="hidKBN<%=m_Rs("T13_GAKUSEKI_NO")%>" size="2" value="<%=gf_SetNull2Zero(m_Rs("T30_SYUKKETU_KBN"))%>"></td>
                            <%
                        End If
                        %>

                    </tr>

            <%If i = w_iCnt Then
                '//���X�g�����s����

                '//���ټ�Ă̸׽��������
				w_Class = ""
                %>
                </table>
                </td>
				<td width="10"></td>
				<td valign="top">
                <!--�w�b�_-->
                <table class="hyo" border="1" >

            <%End If

            i = i + 1
            m_Rs.MoveNext%>
        <%Loop%>

                </table>
                </td></tr>
            </table>
            <br>
<%		if m_sEndDay < m_sDate then %>
            <table>
				<tr>
                <td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="�@�o�@�^�@"></td>
                <td ><input class="button" type="button" onclick="javascript:f_Cancel();" value="�L�����Z��"></td>
				</tr>
            </table>
<% Else %>
            <table>
				<tr>
                <td ><input class="button" type="button" onclick="javascript:f_Cancel();" value=" �߁@�� "></td>
				</tr>
            </table>
<% End If %>
        <%Exit Do%>
    <%Loop%>

    <!--�l�n���p-->
    <input type="hidden" name="NENDO"     value="<%=m_iSyoriNen%>">
    <input type="hidden" name="KYOKAN_CD" value="<%=m_iKyokanCd%>">
    <input type="hidden" name="GAKUNEN"   value="<%=m_iGakunen%>">
    <input type="hidden" name="CLASSNO"   value="<%=m_iClassNo%>">
    <input type="hidden" name="cboDate"   value="<%=m_sDate%>">

    <input type="hidden" name="txtMsg"    value="">

    </form>
    </center>
    </body>
    </html>
<%
End Sub

'********************************************************************************
'*  [�@�\]  ��HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
    <html>
    <head>
    <title>�����o������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		//parent.location.href="white.htm"
    }
    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">

	<center>
	<br><br><br>
		<span class="msg"><%=p_Msg%></span>
	</center>

    </body>
    </html>
<%
End Sub
%>
