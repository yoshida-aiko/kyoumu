<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: ���Ǝ��Ԉꗗ
'* ��۸���ID : login/jikanwari.asp
'* �@      �\: ��y�[�W ���Ԋ��}�X�^�̈ꗗ���X�g�\�����s��
'*-------------------------------------------------------------------------
'* ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'* ��      ��:
'*           ���O�C�����������̎��Ǝ��Ԉꗗ��\��
'*-------------------------------------------------------------------------
'* ��      ��: 2001/07/19 ���{ ����
'* ��      �X: 2001/07/25 ���{ ����
'*           : 2001/08/06 ���{ ����     �߂��URL�Atarget�ύX
'*           :                          �ϐ��������K���Ɋ�ύX
'*           : 2001/08/07 ���{ ����     NN�Ή��ɔ����\�[�X�ύX
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_sMsg              'ү����
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iKyokanCd         ':�����R�[�h
    
    Public  m_Rs                'recordset
    Public  m_Rds                'recordset
    
    Public  m_iGakunen          ':�w�N
    Public  m_sClass            ':�N���X
    Public  m_sYobi             ':�\���j��
    Public  m_iYobiCd           ':�j���R�[�h
    Public  m_iJigen            ':����
    Public  m_iKamokuCd         ':�ȖڃR�[�h
    Public  m_sKamoku           ':�Ȗږ�
    Public  m_iKyosituCd        ':�����R�[�h
    Public  m_sKyositu          ':������
    
    Public  m_sCellD             ':�e�[�u���Z���F�i�j���j
    Public  m_iJMax             ':�ő厞����
    
    '�y�[�W�֌W
    Public  m_iMax              ':�ő�y�[�W
    Public  m_iDsp              '// �ꗗ�\���s��

    '�f�[�^�擾�p
    Public  m_iYobiCnt          ':�J�E���g�i�j���j
    Public  m_iJgnCnt           ':�J�E���g�i�����j
    Public  m_iYobiCCnt         ':�J�E���g�i�j���E�e�[�u���F�\���p�j
    Public  m_iDate             ':�����̓��t
    Public  m_sYobiD            ':�����̗j��
    
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
    w_sMsgTitle="HOME"
    w_sMsg=""
    w_sRetURL="../default.asp"     
    w_sTarget="_top"


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

        '// �����`�F�b�N�Ɏg�p
        session("PRJ_No") = C_LEVEL_NOCHK

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '// �l�̏�����
        Call s_SetBlank()

        '// ���Ұ�SET
        Call s_SetParam()
        
        '// ���Ұ�SET
        Call s_SetDate()
        
        '���Ǝ��Ԋ��e�[�u���}�X�^���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_GAKUNEN"
        w_sSQL = w_sSQL & vbCrLf & " ,M05_CLASS.M05_CLASSRYAKU"
        w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_YOUBI_CD"
        w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_JIGEN"
        w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_KYOSITU"
        w_sSQL = w_sSQL & vbCrLf & " ,M03_KAMOKU.M03_KAMOKUMEI"
        'w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_GODO_FLG"
        w_sSQL = w_sSQL & vbCrLf & " FROM T20_JIKANWARI"
        w_sSQL = w_sSQL & vbCrLf & ", M03_KAMOKU"
        w_sSQL = w_sSQL & vbCrLf & ", M05_CLASS"
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND M03_KAMOKU.M03_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND M05_CLASS.M05_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_KYOKAN = '" & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_KAMOKU = M03_KAMOKU.M03_KAMOKU_CD(+) "
        w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN(+) "
        w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO(+) "



response.write "�H����" & "<BR>"
response.end
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            'm_sErrMsg = Err.description
            m_sErrMsg = "���R�[�h�Z�b�g�̎擾�Ɏ��s���܂���"
            Exit Do 'GOTO LABEL_MAIN_END
        Else
            '�y�[�W���̎擾
            'm_iMax = gf_PageCount(m_Rs,m_iDsp)
        End If

        '//�ő厞�������擾
        Call gf_GetJigenMax(m_iJMax)
            if m_iJMax = "" Then
                m_bErrFlg = True
                m_sErrMsg = Err.description
                Exit Do
            end if
        
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
    gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub
'********************************************************************************
'*  [�@�\]  �S���ڂ��󔒂ɏ�����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetBlank()

    m_iKyokanCd = ""
    m_iSyoriNen = ""
    
    m_sYobi = ""
    m_iYobiCd = ""
    m_iJigen = ""
    m_iKamokuCd = ""
    
    m_iYobiCnt = ""
    m_iJgnCnt = ""
    m_iYobiCCnt = ""
    
    m_iDate = ""
    m_sYobiD = ""
    
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")          ':�����R�[�h
    m_iSyoriNen = Session("NENDO")              ':�����N�x
    
End Sub

Sub s_SetDate()
'********************************************************************************
'*  [�@�\]  �����̓��t��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iDate = gf_YYYY_MM_DD(date(),"/")
    m_sYobiD = Weekdayname(Weekday(m_iDate),true)

End Sub

Sub s_SetYobi()
'********************************************************************************
'*  [�@�\]  �j����ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sYobi = Weekdayname(m_iYobiCnt,true)

End Sub


Sub showYobi()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �j����\���i�e�[�u���p�j
'********************************************************************************

if m_iYobiCCnt Mod 2 <> 0 Then
    m_sCellD = ""
end if

call gs_cellPtn(m_sCellD)

    if m_iJgnCnt = 1 Then
        response.write "<td rowspan='" & m_iJMax & "' class='"
        response.write m_sCellD
        response.write "'>" & m_sYobi & "</td>"
    else
    end if
    
End Sub

Sub showClass()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �w�N�E�N���X��\��
'********************************************************************************

    Dim w_sClass

    m_iGakunen = ""
    m_sClass = ""
    
    w_sClass = ""

    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CInt(m_Rs("T20_JIGEN")) = CInt(m_iJgnCnt) Then
            m_iGakunen = m_Rs("T20_GAKUNEN")
            'w_sClass = Right(m_Rs("M05_CLASSRYAKU"),1)
            'm_sClass = m_sClass & w_sClass
            m_sClass = m_sClass & m_Rs("M05_CLASSRYAKU")
            
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst
    
    if m_iGakunen <> "" and m_sClass <> "" Then
        response.write m_iGakunen & "-" & m_sClass
    end if

End Sub

Sub showKamokuMei()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �Ȗږ���\��
'********************************************************************************

    m_sKamoku = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CInt(m_Rs("T20_JIGEN")) = CInt(m_iJgnCnt) Then
            m_sKamoku = m_Rs("M03_KAMOKUMEI")
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_iGakunen <> "" and m_sClass <> "" Then
        response.write m_sKamoku
    end if

End Sub

Sub SetKyositu()
'********************************************************************************
'*  [�@�\]  �l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �����R�[�h��ݒ�
'********************************************************************************

    m_iKyosituCd = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CInt(m_Rs("T20_JIGEN")) = CInt(m_iJgnCnt) Then
            m_iKyosituCd = m_Rs("T20_KYOSITU")
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

End Sub

sub s_Jikanwari()
'********************************************************************************
'*  [�@�\]  ���Ԋ��ύX�f�[�^�̗L���̊m�F
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    w_sSQL = ""
    w_sSQL = w_sSQL & " SELECT * "
    w_sSQL = w_sSQL & " FROM "
    w_sSQL = w_sSQL & "     T52_JYUGYO_HENKO "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & "     T52_KYOKAN_CD = '" & m_iKyokanCd & "' "
    w_sSQL = w_sSQL & " AND T52_KAISI <= '" & gf_YYYY_MM_DD(date(),"/") & "'"
    w_sSQL = w_sSQL & " AND T52_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "'"

    Set m_Rds = Server.CreateObject("ADODB.Recordset")
    w_iRet = gf_GetRecordsetExt(m_Rds, w_sSQL,m_iDsp)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
    End If

    If m_Rds.EOF Then
        Exit Sub
    End If
%>
<a href="#" onclick=NewWin()> �����Ԋ��̕ύX�A��������܂�</a>  
<%

End Sub

'********************************************************************************
'*  [�@�\]  �������̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetKyosituMei()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    Call SetKyositu()
    m_sKyositu = ""
    
    if m_iKyosituCd <> "" Then
        Do
            
            '// �������}�X�^ں��޾�Ă��擾
            w_sSQL = ""
            w_sSQL = w_sSQL & "SELECT"
            w_sSQL = w_sSQL & " M06_KYOSITUMEI"
            w_sSQL = w_sSQL & " FROM M06_KYOSITU "
            w_sSQL = w_sSQL & " WHERE M06_NENDO = " & m_iSyoriNen
            w_sSQL = w_sSQL & " AND M06_KYOSITU_CD = " & m_iKyosituCd
            
            w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
            
            If w_iRet <> 0 Then
                'ں��޾�Ă̎擾���s
                'm_sErrMsg = "���R�[�h�Z�b�g�̎擾�Ɏ��s���܂���"
                'm_bErrFlg = True
                's_SetKyosituMei = 99
                Exit Do 'GOTO LABEL_s_SetKyosituMei_END
            Else
            End If
            
            If w_Rs.EOF Then
                '�Ώ�ں��ނȂ�
                'm_sErrMsg = "�Ώۃ��R�[�h������܂���"
                's_SetKyosituMei = 1
                Exit Do 'GOTO LABEL_s_SetKyosituMei_END
            End If
            
                '// �擾�����l���i�[
                    m_sKyositu = w_Rs("M06_KYOSITUMEI")    '//���������i�[
            '// ����I��
            Exit Do
        
        Loop
        
        gf_closeObject(w_Rs)
    
    end if

response.write m_sKyositu

End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    Dim w_sCellT             '//Table�Z���F

    On Error Resume Next
    Err.Clear

%>


<html>
<head>
<link rel="stylesheet" href="../common/style.css" type="text/css">
<!--#include file="../Common/jsCommon.htm"-->
<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
    //************************************************************
    //  [�@�\]  �\�����e�\���p�E�B���h�E�I�[�v��
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function NewWin() {
        URL = "j_view.asp";
        nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=450,height=450,top=0,left=0");
        return false;   
    }
//-->
</SCRIPT>
</head>
<body>
<center>
<BR>
<font size="3">�����̎��Ԋ�</font>
<BR><BR>
<table border="0" width="90%">
<tr>
<td valign="top" align="center">
    <table border="1" class="hyo" width="100%">
    <COLGROUP WIDTH="5%" ALIGN="center">
    <COLGROUP WIDTH="5%" ALIGN="center">
    <COLGROUP WIDTH="20%" ALIGN="center">
    <COLGROUP WIDTH="35%" ALIGN="center">
    <COLGROUP WIDTH="35%" ALIGN="center">
    <tr>
        <th colspan="2" class="header"><br></th>  
        <th class="header">�N���X</th>
        <th class="header">���Ȗ�</th>
        <th class="header">����</th>
    </tr>
<%

    m_iYobiCCnt = 1
    
    'For m_iYobiCnt = C_YOUBI_MIN to C_YOUBI_MAX
    For m_iYobiCnt = 1 to 7

        For m_iJgnCnt = 1 to m_iJMax
        
        '//�e�[�u���Z���w�i�F
        call gs_cellPtn(w_sCellT)
        '//���Ԋ��e�[�u������j�����擾
        call s_SetYobi()

'response.write "aaaaaa" & "<BR>"
'response.end

            '//���Ԋ��̗j���ƍ����̗j��������̏ꍇ�\��
            if m_sYobi = m_sYobiD Then

'response.write "aaaaaa" & "<BR>"
'response.end


%>
    <tr>
<%
            call showYobi()
%>
        <td class="<%=w_sCellT%>"><%=m_iJgnCnt%></td>


        <td class="<%=w_sCellT%>"><%call showClass()%><br></td>

        <td class="<%=w_sCellT%>"><%call showKamokuMei()%><br></td>
        <td class="<%=w_sCellT%>"><%call s_SetKyosituMei()%><br></td>
    </tr>
<%
            end if
        Next
    m_iYobiCCnt = m_iYobiCCnt + 1   '//�j���J�E���g�i�e�[�u���w�i�F�\���p�j
    Next
%>
    </table>
</td>
</tr>
</table>
<%
    Do Until m_Rs.EOF
%>
<%
m_Rs.MoveNext

    if m_Rs.EOF Then
        Exit Do
    end if
    Loop
%>
<br>
<%Call s_Jikanwari()%>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
