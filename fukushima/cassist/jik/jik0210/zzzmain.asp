<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: �����ʎ��Ǝ��Ԉꗗ
'* ��۸���ID : jik/jik0210/main.asp
'* �@      �\: ���y�[�W ���Ԋ��}�X�^�̈ꗗ���X�g�\�����s��
'*-------------------------------------------------------------------------
'* ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*           cboGakunenCd      :�w�N�R�[�h
'*           cboClassCd      :�N���X�R�[�h
'*           txtMode         :���샂�[�h
'           :session("PRJ_No")      '���������̃L�[
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'* ��      ��:
'*           �I�����ꂽ�N���X�̎��Ǝ��Ԉꗗ��\��
'*-------------------------------------------------------------------------
'* ��      ��: 2001/07/06 ���{ ����
'* ��      �X: 2001/07/30 ���{ ����  �߂��URL�ύX
'*                                  �ϐ��������K���Ɋ�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_sMsg              'ү����
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iGakunen          ':�w�N�R�[�h
    Public  m_iClass            ':�N���X�R�[�h
    
    Public  m_Rs                'recordset
    
    Public  m_sClass            ':�N���X��
    Public  m_sYobi             ':�\���j��
    Public  m_iYobiCd           ':�j���R�[�h
    Public  m_iJigen            ':����
    Public  m_iKamokuCd         ':�ȖڃR�[�h
    Public  m_sKamoku           ':�Ȗږ�
    Public  m_iKyosituCd        ':�����R�[�h
    Public  m_sKyositu          ':������
    Public  m_sKyokan           ':������
    Public  m_iNyuNen           ':���N�x
    Public  m_iCourseCd         ':�R�[�X�R�[�h
    
    Public  m_iJMax             ':�ő厞����
    Public  m_Flg			'���Ԋ��P���ڊm�F�t���O
    
    Public m_iCourse            ':�R�[�X�R�[�h
    
    '�y�[�W�֌W
    Public  m_iMax              ':�ő�y�[�W
    Public  m_iDsp              '// �ꗗ�\���s��

    '�f�[�^�擾�p
    Public  m_iYobiCnt          ':�J�E���g�i�j���j
    Public  m_iJgnCnt           ':�J�E���g�i�����j
    Public  m_iYobiCCnt         ':�J�E���g�i�j���E�e�[�u���F�\���p�j
    
    Public  m_sCellD             ':�e�[�u���Z���F�i�j���j'//2001/07/30�ύX
    
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
    w_sMsgTitle="�N���X�ʎ��Ǝ��Ԉꗗ"
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

        '// �l�̏�����
        Call s_SetBlank()

        '// ���Ұ�SET
        Call s_SetParam()

        '// �N���X���擾
        Call f_GetClassMei()
            if f_GetClassMei <> 0 Then
                Exit Do
            end if
        
        '���Ǝ��Ԋ��e�[�u���}�X�^���擾
        
            w_sSQL = ""
            w_sSQL = w_sSQL & "SELECT"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_YOUBI_CD,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_JIGEN,"
            w_sSQL = w_sSQL & vbCrLf & " M04.M04_KYOKANMEI_SEI,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_KAMOKU, "
            w_sSQL = w_sSQL & vbCrLf & " M06.M06_KYOSITUMEI,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_DOJI_JISSI_FLG,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_TUKU_FLG "
            w_sSQL = w_sSQL & vbCrLf & "FROM "
            w_sSQL = w_sSQL & vbCrLf & "T20_JIKANWARI T20,"
            w_sSQL = w_sSQL & vbCrLf & "M04_KYOKAN M04,"
            w_sSQL = w_sSQL & vbCrLf & "M06_KYOSITU M06 "
            w_sSQL = w_sSQL & vbCrLf & "WHERE "
            w_sSQL = w_sSQL & vbCrLf & "T20.T20_KYOKAN = M04.M04_KYOKAN_CD(+) "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_NENDO = M04.M04_NENDO(+) "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_KYOSITU = M06.M06_KYOSITU_CD(+) "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_NENDO = M06.M06_NENDO(+) "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_DOJI_JISSI_FLG IS NULL "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_NENDO=" & m_iSyoriNen & " "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_GAKUNEN=" & m_iGakunen & " "
            w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_CLASS=" & m_iClass & " "
            w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_YOUBI_CD,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_JIGEN,"
            w_sSQL = w_sSQL & vbCrLf & " M04.M04_KYOKANMEI_SEI,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_KAMOKU, "
            w_sSQL = w_sSQL & vbCrLf & " M06.M06_KYOSITUMEI,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_DOJI_JISSI_FLG,"
            w_sSQL = w_sSQL & vbCrLf & " T20.T20_TUKU_FLG "
            w_sSQL = w_sSQL & vbCrLf & "ORDER BY "
'            w_sSQL = w_sSQL & vbCrLf & "T20.T20_GAKKI_KBN, "
            w_sSQL = w_sSQL & vbCrLf & "T20.T20_YOUBI_CD, "
            w_sSQL = w_sSQL & vbCrLf & "T20.T20_JIGEN "
            
'        Response.Write w_sSQL & vbCrLf &"<br>"
'response.end
        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        'w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
        
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

    m_sYobi = ""
    m_iYobiCd = ""
    m_iJigen = ""
    m_iKamokuCd = ""
    
    m_iYobiCnt = ""
    m_iJgnCnt = ""
    m_iYobiCCnt = ""
    
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()    '//2001/07/30�ύX

    m_iKyokanCd = Session("KYOKAN_CD")          ':�����R�[�h
    'm_iKyokanCd = 000000                       ':�����R�[�h'//�e�X�g�p
    m_iSyoriNen = Session("NENDO")              ':�����N�x
    'm_iSyoriNen = 2002                         ':�����N�x'//�e�X�g�p
    
    m_iGakunen = Request("cboGakunenCd")   ':�w�N�R�[�h
    m_iClass = Request("cboClassCd")       ':�N���X�R�[�h
    
    m_iNyuNen = m_iSyoriNen - m_iGakunen + 1
    
End Sub

'********************************************************************************
'*  [�@�\]  �N���X���̎擾
'*  [����]  
'*  [�ߒl]  0:���擾�����A1:���R�[�h�Ȃ��A99:���s
'*  [����]  
'********************************************************************************
Function f_GetClassMei()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    f_GetClassMei = 0
    m_sClass = ""

    Do

        '// �N���X�}�X�^���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & vbCrLf & "M05_CLASSMEI "
        w_sSQL = w_sSQL & vbCrLf & "FROM "
        w_sSQL = w_sSQL & vbCrLf & "M05_CLASS "
        w_sSQL = w_sSQL & vbCrLf & "WHERE "
        w_sSQL = w_sSQL & vbCrLf & "M05_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & vbCrLf & "AND M05_GAKUNEN = " & m_iGakunen
        w_sSQL = w_sSQL & vbCrLf & "AND M05_CLASSNO = " & m_iClass
        
        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
        
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            f_GetClassMei = 99
            Exit Do 'GOTO LABEL_f_GetClassMei_END
        Else
        End If
        
        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            f_GetClassMei = 1
            Exit Do 'GOTO LABEL_f_GetClassMei_END
        End If

            '// �擾�����l���i�[
            m_sClass = w_Rs("M05_CLASSMEI")
        '// ����I��
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs)

'// LABEL_f_GetClassMei_END
End Function

Function f_getKamokumei(p_iNendo,p_sKamokuCD,p_iGaknen,p_iTUKU,p_iCourseCD) 
'********************************************************************************
'*  [�@�\]  �Ȗږ��̎擾(���łɃR�[�X������Ă��܂��B)
'*  [����]  
'*  [�ߒl]  �Ȗږ�
'*  [����]  2001/9/15
'********************************************************************************
    dim w_sSQL,w_Rs,w_iRet
    
    On Error Resume Next
    Err.Clear
    
    f_getKamokumei = "-"
    p_iCourseCD = 0 
    
  Do

   if p_iTUKU = C_TUKU_FLG_TOKU then '���ʊ����̂Ƃ��́AM41���ʊ����}�X�^���疼�̎擾
    	w_sSQL = ""
    	w_sSQL = w_sSQL & vbCrLf & "SELECT "
    	w_sSQL = w_sSQL & vbCrLf & "M41_MEISYO "
    	w_sSQL = w_sSQL & vbCrLf & "FROM  "
    	w_sSQL = w_sSQL & vbCrLf & "M41_TOKUKATU "
    	w_sSQL = w_sSQL & vbCrLf & "WHERE "
    	w_sSQL = w_sSQL & vbCrLf & "M41_NENDO = " & p_iNendo & " AND "
    	w_sSQL = w_sSQL & vbCrLf & "M41_TOKUKATU_CD = '" & p_sKamokuCD & "' "
    	w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
    	w_sSQL = w_sSQL & vbCrLf & "M41_MEISYO "

   	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
   	If w_iRet <> 0 OR w_Rs.EOF = true Then Exit Do 
		
   	f_getKamokumei = w_Rs("M41_MEISYO")

    Else '���ʂ̎��Ƃ̂Ƃ��́AT15���C���疼�̎擾
    	w_sSQL = ""
    	w_sSQL = w_sSQL & vbCrLf & "SELECT "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKUMEI, "
    	w_sSQL = w_sSQL & vbCrLf & "T15_COURSE_CD"
    	w_sSQL = w_sSQL & vbCrLf & "FROM "
    	w_sSQL = w_sSQL & vbCrLf & " T15_RISYU "
    	w_sSQL = w_sSQL & vbCrLf & "WHERE "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKU_CD = '"&p_sKamokuCD&"' AND "
    	w_sSQL = w_sSQL & vbCrLf & "T15_NYUNENDO = "& p_iNendo - p_iGaknen + 1 &" "
    	w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKUMEI, T15_COURSE_CD"

    	w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
    	If w_iRet <> 0 OR w_Rs.EOF = true Then Exit Do 

    	f_getKamokumei = w_Rs("T15_KAMOKUMEI")
    	p_iCourseCD = w_Rs("T15_COURSE_CD") 

    End if
    Exit Do

   Loop

end Function

Sub s_ShowYobi(p_iJigenMax)    '//2001/07/30�ύX
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

    if m_iJgnCnt <= 1  And m_Flg = 0 Then
	m_Flg = 1
        'response.write "<td rowspan=8 class="
        response.write "<td rowspan=" & p_iJigenMax & " class="
        'call showYobiColor()
        response.write m_sCellD
        response.write ">" & WeekdayName(m_iYobiCnt,True) & "</td>"
    else
    end if
    
End Sub

Function f_ShowKamokuMei()   '//2001/07/30�ύX
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �Ȗږ���\��
'********************************************************************************
dim w_iCourseCD
    m_sKamoku = ""
    f_ShowKamokuMei = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then

		m_sKamoku = f_getKamokumei(m_iSyoriNen,m_Rs("T20_KAMOKU"),m_iGakunen,m_Rs("T20_TUKU_FLG"),w_iCourseCD) 

            if CInt(w_iCourseCD) = 0 Then

'                m_iCourseCd = m_Rs("T15_COURSE_CD")
                Exit Do
            else
                m_sKamoku = "-"
                'm_iCourseCd = m_Rs("T15_COURSE_CD")
            end if
        else
            m_sKamoku = "-"
            'm_iCourseCd = m_Rs("T15_COURSE_CD")
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst
    
    if m_sKamoku = "-" Then
        Call f_GetSentaku()
            'if m_sKamoku = "@@" Then
            if f_GetSentaku = 1 or m_sKamoku = "" Then
                Call s_GetCourse()
            end if
    end if

    if m_iGakunen <> "" and m_iClass <> "" Then
        f_ShowKamokuMei = m_sKamoku
    else
        f_ShowKamokuMei =  "-"
    end if

End Function

'********************************************************************************
'*  [�@�\]  �I���Ȗږ��̎擾
'*  [����]  
'*  [�ߒl]  0:���擾�����A1:���R�[�h�Ȃ��A99:���s
'*  [����]  
'********************************************************************************
Function f_GetSentaku()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    f_GetSentaku = 0
    m_sKamoku = ""
    m_sKyokan = ""

    Do

        '// ���Ǝ��Ԋ��e�[�u���}�X�^���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"

w_sSQL = w_sSQL & vbCrLf & "T20.T20_YOUBI_CD, "
w_sSQL = w_sSQL & vbCrLf & "T20.T20_JIGEN,"
w_sSQL = w_sSQL & vbCrLf & "T20.T20_HYOJI_KYOKAN,"
w_sSQL = w_sSQL & vbCrLf & "T18.T18_SYUBETU_MEI,"
w_sSQL = w_sSQL & vbCrLf & "T20.T20_DOJI_JISSI_FLG,"
w_sSQL = w_sSQL & vbCrLf & "T15.T15_COURSE_CD"
w_sSQL = w_sSQL & vbCrLf & "FROM "
w_sSQL = w_sSQL & vbCrLf & "T20_JIKANWARI T20, "
w_sSQL = w_sSQL & vbCrLf & "T15_RISYU T15,"
w_sSQL = w_sSQL & vbCrLf & "T18_SELECTSYUBETU T18 "
w_sSQL = w_sSQL & vbCrLf & "WHERE "
w_sSQL = w_sSQL & vbCrLf & "T20.T20_KAMOKU = T15.T15_KAMOKU_CD(+) "
w_sSQL = w_sSQL & vbCrLf & "AND T15.T15_NYUNENDO = T18.T18_NYUNENDO(+) "
w_sSQL = w_sSQL & vbCrLf & "AND T15.T15_GRP = T18.T18_GRP(+) "
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_GAKKI_KBN = '1' "
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_NENDO = " & m_iSyoriNen
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_GAKUNEN = " & m_iGakunen
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_CLASS = " & m_iClass
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_YOUBI_CD = " & m_iYobiCnt
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_JIGEN = " & m_iJgnCnt
w_sSQL = w_sSQL & vbCrLf & "AND T20.T20_DOJI_JISSI_FLG Is Not Null "
w_sSQL = w_sSQL & vbCrLf & "AND T15.T15_NYUNENDO = " & m_iNyunen
w_sSQL = w_sSQL & vbCrLf & "GROUP BY "
w_sSQL = w_sSQL & vbCrLf & "T20.T20_YOUBI_CD, "
w_sSQL = w_sSQL & vbCrLf & "T20.T20_JIGEN,"
w_sSQL = w_sSQL & vbCrLf & "T20.T20_HYOJI_KYOKAN,"
w_sSQL = w_sSQL & vbCrLf & "T18.T18_SYUBETU_MEI,"
w_sSQL = w_sSQL & vbCrLf & "T20.T20_DOJI_JISSI_FLG,"
w_sSQL = w_sSQL & vbCrLf & "T15.T15_COURSE_CD"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
'response.write w_sSQL & "<br>"
        
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            'response.write w_iRet & "<br>"
            'm_sErrMsg = "ں��޾�Ă̎擾���s"
            'm_bErrFlg = True
            f_GetSentaku = 99
            Exit Do 'GOTO LABEL_f_GetSentaku_END
        Else
        End If
        
        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            'm_sErrMsg = "�Ώ�ں��ނȂ�"
            f_GetSentaku = 1
            Exit Do 'GOTO LABEL_f_GetSentaku_END
        End If

            '// �擾�����l���i�[
            'm_sKamoku = w_Rs("T18_SYUBETU_MEI") & "@@"
            'm_sKyokan =  w_Rs("T20_HYOJI_KYOKAN") & "@@"
            'm_sKyositu =  "@@"
            'm_iCourseCd = w_Rs("T15_COURSE_CD") & "@@"
            if IsNull(w_Rs("T18_SYUBETU_MEI")) = False Then
                m_sKamoku = w_Rs("T18_SYUBETU_MEI")
            else
                m_sKamoku = ""
            end if
            if IsNull(w_Rs("T20_HYOJI_KYOKAN")) = False Then
                m_sKyokan = w_Rs("T20_HYOJI_KYOKAN")
            else
                m_sKyokan = ""
            end if
            m_sKyositu =  ""
            'm_iCourseCd = w_Rs("T15_COURSE_CD")
        '// ����I��
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs)

'// LABEL_f_GetSentaku_END
End Function


'********************************************************************************
'*  [�@�\]  �R�[�X�ʂ̎擾
'*  [����]  
'*  [�ߒl]  0:���擾�����A1:���R�[�h�Ȃ��A99:���s
'*  [����]  
'********************************************************************************
Sub s_GetCourse()
    
    Dim w_Rs2                '// ں��޾�ĵ�޼ު��
    Dim w_iRet2              '// �߂�l
    Dim w_sSQL2              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    m_iCourse = 0
    m_sKamoku = ""
    m_sKyokan = ""
    m_sKyositu = ""

    Do

        '// ���Ǝ��Ԋ��e�[�u���}�X�^���擾
        w_sSQL2 = ""
        w_sSQL2 = w_sSQL2 & "SELECT"
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_YOUBI_CD, "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_JIGEN,"
'w_sSQL2 = w_sSQL2 & vbCrLf & "T15.T15_KAMOKUMEI,"
'w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_KYOSITU,"
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_HYOJI_KYOKAN as M04_KYOKANMEI_SEI"
w_sSQL2 = w_sSQL2 & vbCrLf & "FROM "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20_JIKANWARI T20, "
w_sSQL2 = w_sSQL2 & vbCrLf & "T15_RISYU T15"
w_sSQL2 = w_sSQL2 & vbCrLf & "WHERE "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_KAMOKU = T15.T15_KAMOKU_CD(+) "
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_GAKKI_KBN = '1' "
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_NENDO = " & m_iSyoriNen
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_GAKUNEN = " & m_iGakunen
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_CLASS = " & m_iClass
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T15.T15_COURSE_CD != '0' "
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T15.T15_NYUNENDO = " & m_iNyunen
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_YOUBI_CD = " & m_iYobiCnt
w_sSQL2 = w_sSQL2 & vbCrLf & "AND T20.T20_JIGEN = " & m_iJgnCnt
w_sSQL2 = w_sSQL2 & vbCrLf & "GROUP BY "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_YOUBI_CD, "
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_JIGEN,"
'w_sSQL2 = w_sSQL2 & vbCrLf & "T15.T15_KAMOKUMEI,"
'w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_KYOSITU,"
w_sSQL2 = w_sSQL2 & vbCrLf & "T20.T20_HYOJI_KYOKAN"

        w_iRet2 = gf_GetRecordset(w_Rs2, w_sSQL2)
'response.write w_sSQL2 & "<br>"
        
        If w_iRet2 <> 0 Then
            'ں��޾�Ă̎擾���s
            'response.write w_iRet2 & "<br>"
            'm_sErrMsg = "ں��޾�Ă̎擾���s"
            'm_bErrFlg = True
            'response.write "?"
            m_iCourse = 99
            Exit Do
        Else
        End If
        
        If w_Rs2.EOF Then
            '�Ώ�ں��ނȂ�
            'm_sErrMsg = "�Ώ�ں��ނȂ�"
            m_iCourse = 1
            'response.write "---"
            Exit Do
        End If

            '// �擾�����l���i�[
            'm_sKamoku = w_Rs2("T15_KAMOKUMEI") & "*"
            if m_iCourse = 0 Then
                m_sKamoku = "�R�[�X��"
                if IsNull(w_Rs2("T20_HYOJI_KYOKAN")) = False Then
                    'm_sKyokan = m_iCourse
                    m_sKyokan = w_Rs2("T20_HYOJI_KYOKAN")
                else
                    m_sKyokan = ""
                end if
                m_sKyositu = w_Rs2("T20_KYOSITU")
            else
                m_sKamoku = "?"
                m_sKyokan = "?"
                m_sKyositu = "?"
            end if
        '// ����I��
        Exit Do
    
    
    Loop
    
    gf_closeObject(w_Rs2)

End Sub

Function f_ShowKyokanMei()   '//2001/07/30�ύX
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  ��������\��
'********************************************************************************

    m_sKyokan = ""
    f_ShowKyokanMei = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then
            m_sKyokan = m_Rs("M04_KYOKANMEI_SEI")
            Exit Do

        else
            'm_sKyokan = ""
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_sKyokan = "" Then
        Call f_GetSentaku()
            if f_GetSentaku = 1 or m_sKyokan = "" Then
            'if f_GetSentaku = 1 or m_sKyokan = "@@" Then
                Call s_GetCourse()
            end if
    end if

    if m_iGakunen <> "" and m_iClass <> "" Then
        f_ShowKyokanMei = m_sKyokan
    else
        f_ShowKyokanMei = "-"
    end if
    
End Function

Sub s_ShowKyosituMei()  '//2001/07/30�ύX
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  ��������\��
'********************************************************************************

    m_sKyositu = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then
            m_sKyositu = m_Rs("M06_KYOSITUMEI")
            Exit Do
        else
            'm_sKyositu = ""
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_sKyositu = "" Then
        Call f_GetSentaku()
            if f_GetSentaku = 1 or m_sKyositu = "" Then
            'if f_GetSentaku = 1 or m_sKyositu = "@@" Then
                Call s_GetCourse()
            end if
    end if

    if m_iGakunen <> "" and m_iClass <> "" Then
        response.write m_sKyositu
    else
        response.write "-"
    end if

End Sub

Function f_KamokuSu(p_iYobiCnt)   '//2001/09/06 add
'********************************************************************************
'*  [�@�\]  �Ȗڐ����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �Ȗږ���\��
'********************************************************************************
    f_KamokuSu = cint(m_iJMax)
    m_Rs.MoveFirst
    Do Until m_Rs.EOF

        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(p_iYobiCnt) and right(cstr(CDbl(m_Rs("T20_JIGEN"))*10),1) <> "0" Then
            f_KamokuSu = f_KamokuSu + 1
        end if
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

'    if m_iGakunen <> "" and m_sClass <> "" Then
'        response.write m_sKamoku
'    end if

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    Dim w_cellT             '//Table�Z���F
    Dim w_sKyokan        '//������
    Dim w_sKamoku       '//�Ȗږ�
    Dim w_iJigenMax      '//�j����row�p
    Dim w_iJgnNo          '//�\���p����

   
    On Error Resume Next
    Err.Clear

%>


<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
</head>

<body>

<center>

<table border=1 class=hyo >
	<tr>
	<th class=header width="64"  align="center">�N���X</th>
	<td class=detail width="50"  align="center"><%=m_iGakunen%>�N</td>
	<td class=detail width="130" align="center"><%=m_sClass%></td>
	</tr>
</table>
<br>
<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr>
<td align="center">

    <table border=1 class=hyo width="100%">
        <COLGROUP WIDTH="5%" ALIGN=center>
        <COLGROUP WIDTH="5%" ALIGN=center>
        <COLGROUP WIDTH="30%" ALIGN=center>
        <COLGROUP WIDTH="30%" ALIGN=center>
        <COLGROUP WIDTH="30%" ALIGN=center>
        <tr>
            <th colspan="2" class=header><br></th>  
            <th class=header>�Ȗ�</th>
            <th class=header>����</th>
            <th class=header>����</th>
        </tr>

<%
m_iYobiCCnt = 1

For m_iYobiCnt = C_YOUBI_MIN to C_YOUBI_MAX
 m_Flg = 0
 w_iJigenMax =  f_KamokuSu(m_iYobiCnt)
    For m_iJgnCnt = 0.5 to m_iJMax step 0.5

	w_sKyokan =  f_ShowKyokanMei()
	w_sKamoku = f_ShowKamokuMei()
	w_iJgnNo = m_iJgnCnt
	if right(cstr(w_iJgnNo*10),1) <> "0" then w_iJgnNo = " "

	If w_sKamoku <> "" OR w_sClass <> "" OR right(cstr(m_iJgnCnt*10),1) = "0" Then 

    call gs_cellPtn(w_cellT)
%>
    <tr>
<%call s_ShowYobi(w_iJigenMax)%>
        <td class=<%=w_cellT%>>
        <%=w_iJgnNo%>
        <br></td>
        <td class=<%=w_cellT%>>
        <%=w_sKamoku%>
        <br></td>
        <td class=<%=w_cellT%>>
        <%=w_sKyokan%>
        <br></td>
        <td class=<%=w_cellT%>>
        <%call s_ShowKyosituMei()%>
        <br></td>
    </tr>
<%

	End If
    Next
m_iYobiCCnt = m_iYobiCCnt + 1   '//�j���J�E���g�i�e�[�u���w�i�F�\���p�j
%>
<%
Next
%>

    </table>

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
