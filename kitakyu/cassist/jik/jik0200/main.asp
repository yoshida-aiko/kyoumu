<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: �����ʎ��Ǝ��Ԉꗗ
'* ��۸���ID : jik/jik0200/main.asp
'* �@      �\: ���y�[�W ���Ԋ��}�X�^�̈ꗗ���X�g�\�����s��
'*-------------------------------------------------------------------------
'* ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*           cboKyokaKeiCd      :�Ȗڌn��R�[�h
'*           cboKyokanCd      :�����R�[�h
'*           txtMode         :���샂�[�h
'           :session("PRJ_No")      '���������̃L�[
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'* ��      ��:
'*           �I�����ꂽ�����̎��Ǝ��Ԉꗗ��\��
'*-------------------------------------------------------------------------
'* ��      ��: 2001/07/03 ���{ ����
'* ��      �X: 2001/07/30 ���{ ���� �߂��URL�ύX
'*                                  �ϐ��������K���Ɋ�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iSKyokanCd        ':�I�������R�[�h
    
    Public  m_Rs                'recordset
	Public  gRs					'��֎��Ԋ�ں��޾��
    
    Public  m_iGakunen          ':�w�N
    Public  m_sClass            ':�N���X
    Public  m_sYobi             ':�\���j��
    Public  m_iYobiCd           ':�j���R�[�h
    Public  m_iJigen            ':����
    Public  m_iKamokuCd         ':�ȖڃR�[�h
    Public  m_sKamoku           ':�Ȗږ�
    Public  m_iKyosituCd        ':�����R�[�h
    Public  m_sKyositu          ':������
    
    Public  m_sCellD             ':�e�[�u���Z���F�i�j���j'//2001/07/30�ύX
    Public  m_iJMax             ':�ő厞����
    Public  m_Flg			'���Ԋ��P���ڊm�F�t���O
    
    '�y�[�W�֌W
    Public  m_iMax              ':�ő�y�[�W
    Public  m_iDsp              '// �ꗗ�\���s��

    '�f�[�^�擾�p
    Public  m_iYobiCnt          ':�J�E���g�i�j���j
    Public  m_iJgnCnt           ':�J�E���g�i�����j
    Public  m_iYobiCCnt         ':�J�E���g�i�j���E�e�[�u���F�\���p�j
    
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
    w_sMsgTitle="�����ʎ��Ǝ��Ԉꗗ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    Do

'response.write "AAA"
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
        
        '���Ǝ��Ԋ��e�[�u���}�X�^���擾
		w_sSQL = ""
		w_sSQL = w_sSQL & "SELECT"
		w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_GAKUNEN"
		w_sSQL = w_sSQL & vbCrLf & " ,M05_CLASS.M05_CLASSRYAKU"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_YOUBI_CD"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_JIGEN"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_KAMOKU"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_KYOSITU"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_TUKU_FLG"
		w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI.T20_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " FROM T20_JIKANWARI"
		w_sSQL = w_sSQL & vbCrLf & ", M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE " 
		w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND T20_GAKKI_KBN = " & Session("GAKKI") & " "
		w_sSQL = w_sSQL & vbCrLf & " AND M05_CLASS.M05_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_KYOKAN = '" & m_iSKyokanCd & "' "
		w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN(+) "
		w_sSQL = w_sSQL & vbCrLf & " AND T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO(+) "
		w_sSQL = w_sSQL & vbCrLf & " Order By "
		w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & " T20_JIKANWARI.T20_CLASS "

'response.write w_sSQL
'response.end

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            m_sErrMsg = "���R�[�h�Z�b�g�̎擾�Ɏ��s���܂���"
            Exit Do 'GOTO LABEL_MAIN_END
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
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  ��֎��Ԋ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  2001/12/20�ǉ�
'********************************************************************************
Function f_GetDaigae(p_iYoubiCD,p_iJigen,p_sKamoku,p_flg)

    On Error Resume Next
    Err.Clear

	f_GetDaigae = False

	Dim wSQL

	'??????????????????????????????????????????????????????
	'              SQL�������i�v�f�o�b�N�j
	'??????????????????????????????????????????????????????
	wSQL = ""
	wSQL = wSQL & " SELECT "
	wSQL = wSQL & " T16_KAMOKUMEI "
	wSQL = wSQL & " FROM T23_DAIGAE_JIKAN ,T16_RISYU_KOJIN "
	wSQL = wSQL & " WHERE "
	wSQL = wSQL & " T23_KYOKAN = '" & m_iKyokanCd & "' AND "
	wSQL = wSQL & " T23_NENDO = " & m_iSyoriNen & " AND "
	wSQL = wSQL & " T23_GAKKI_KBN = " & Session("GAKKI") & " AND "
	wSQL = wSQL & " T23_YOUBI_CD = " & p_iYoubiCD & " AND "
	wSQL = wSQL & " T23_JIGEN = " & p_iJigen & " AND "
	wSQL = wSQL & " T16_NENDO = T23_NENDO AND "
	wSQL = wSQL & " T16_KAMOKU_CD = T23_KAMOKU "

'response.write wSQL
	Set gRs = Server.CreateObject("ADODB.Recordset")
	w_iRet = gf_GetRecordset(gRs, wSQL)

	If w_iRet <> 0 Then
		m_bErrFlg = True
		m_sErrMsg = "1���R�[�h�Z�b�g�̎擾�Ɏ��s���܂���"
		Exit Function
	End If

	'�ȖڃR�[�h������΁A�Ԓl�Ƃ��Ď�������B
	if gRs.EOF then 
		p_sKamoku = p_sKamoku
		p_flg = 0
	Else
		p_sKamoku = gRs("T16_KAMOKUMEI")
		p_flg = 1
	End if	

	'�I�u�W�F�N�g�N���[�Y
	gf_closeObject(gRs)

f_GetDaigae = True

End Function

'********************************************************************************
'*  [�@�\]  �S���ڂ��󔒂ɏ�����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetBlank()

    m_iKyokanCd = ""
    m_iSyoriNen = ""
    
    m_iKyokaKeiKbn = ""
    m_iSKyokanCd = ""
    
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
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")          ':�����R�[�h
    m_iSyoriNen = Session("NENDO")              ':�����N�x
    m_iSKyokanCd = Request("SKyokanCd1")        ':�I�������R�[�h
    
End Sub

Sub s_ShowYobi(p_iJigenMax)    '//2001/07/30�ύX
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �j����\���i�e�[�u���p�j
'********************************************************************************
    On Error Resume Next
    Err.Clear


	if m_iYobiCCnt Mod 2 <> 0 Then
	    m_sCellD = ""
	end if

	call gs_cellPtn(m_sCellD)
    if m_iJgnCnt <= 1  And m_Flg = 0 Then
	m_Flg = 1
        'response.write "<td rowspan=8 class="
        response.write "<td rowspan=" & p_iJigenMax & " class="
'        response.write "<td rowspan=" & m_iJMax & " class="
        'call showYobiColor()
        response.write m_sCellD
        response.write ">" & WeekdayName(m_iYobiCnt,True) & "</td>"
    else
    end if
    
End Sub

Function f_ShowClass()   '//2001/07/30�ύX
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �w�N�E�N���X��\��
'********************************************************************************
    On Error Resume Next
    Err.Clear


    Dim w_sClass
    Dim w_clsMax
    Dim w_clsCnt

    m_iGakunen = ""
    m_sClass = ""
    
    w_sClass = ""
    f_ShowClass = ""
	w_clsCnt = 0
    
	w_clsMax = f_GetClassMax(m_iSyoriNen,m_Rs("T20_GAKUNEN"))

    Do Until m_Rs.EOF

		'�j���Ǝ����������Ԃ͕�����쐬
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then

			'����������ҏW
			If m_iGakunen = "" Then
				f_ShowClass =  f_ShowClass & cstr(m_Rs("T20_GAKUNEN")) & "-" & m_Rs("M05_CLASSRYAKU")
	            m_iGakunen = m_Rs("T20_GAKUNEN")
	            m_sClass = m_Rs("M05_CLASSRYAKU")
				w_clsCnt = 1
			End If

			'�w�N���ς�����當����ҏW
			if cstr(m_iGakunen) <> cstr(m_Rs("T20_GAKUNEN")) then
				If w_clsCnt = w_clsMax Then
					f_ShowClass =  m_iGakunen & "-�S"
				End If

				f_ShowClass =  f_ShowClass & "<BR>" & cstr(m_Rs("T20_GAKUNEN")) & "-" & m_Rs("M05_CLASSRYAKU")
	            m_iGakunen = m_Rs("T20_GAKUNEN")
			Else
				
				'�N���X���ς�����當����ҏW
				if cstr(m_sClass) <> cstr(m_Rs("M05_CLASSRYAKU")) then
					w_clsCnt = w_clsCnt + 1

'response.write w_clsCnt & "=" & w_clsMax
					If w_clsCnt = w_clsMax Then
						f_ShowClass =  m_iGakunen & "-�S"
						w_clsCnt = 1
					Else
						f_ShowClass =  f_ShowClass & "," & cstr(m_Rs("M05_CLASSRYAKU"))
					End If
		            m_sClass = m_Rs("M05_CLASSRYAKU")
				End if
			End if
			
        End if
        
        m_Rs.MoveNext
    Loop

    m_Rs.MoveFirst
    
    'if m_iGakunen <> "" and m_sClass <> "" Then
	'	If f_ShowClass = "" Then
	'        f_ShowClass = f_ShowClass & m_iGakunen & "-" & m_sClass
	'	Else
	'        f_ShowClass = f_ShowClass & "," & m_sClass
	'	End If
    'End if

End Function

Function f_ShowKamokuMei()   '//2001/07/30�ύX
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �Ȗږ���\��
'********************************************************************************
    On Error Resume Next
    Err.Clear

    dim w_iCourseCD
	dim w_sKamoku

    w_sKamoku =""
    
    m_sKamoku = ""
    f_ShowKamokuMei = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then

'            m_sKamoku = f_getKamokumei(m_iSyoriNen,m_Rs("T20_KAMOKU"),m_Rs("T20_GAKUNEN"),m_Rs("T20_TUKU_FLG"),w_iCourseCD) 
		
            m_sKamoku = f_getKamokumei(m_iSyoriNen,m_Rs("T20_KAMOKU"),m_Rs("T20_GAKUNEN"),m_Rs("T20_TUKU_FLG"),w_iCourseCD)

            if CStr(w_sKamoku) <> "" And w_sKamoku <> m_sKamoku then
            	w_sKamoku = w_sKamoku & "<BR>" & m_sKamoku
            Else
            	w_sKamoku = m_sKamoku
            End If

        Else
        End If
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

    if m_iGakunen <> "" and m_sClass <> "" Then
'        f_ShowKamokuMei = m_sKamoku		���������R�����g�A�E�g����΁A�\�������Ȗڂ�1�ɂȂ�܂�
        f_ShowKamokuMei = w_sKamoku
    end if

End Function

Function f_KamokuSu(p_iYobiCnt)   '//2001/09/06 add
'********************************************************************************
'*  [�@�\]  �Ȗڐ����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �Ȗږ���\��
'********************************************************************************
    On Error Resume Next
    Err.Clear

    f_KamokuSu = cint(m_iJMax)
    m_Rs.MoveFirst
    Do Until m_Rs.EOF

        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(p_iYobiCnt) and right(cstr(cDbl(m_Rs("T20_JIGEN"))*10),1) <> "0" Then
            f_KamokuSu = f_KamokuSu + 1
        end if
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

'    if m_iGakunen <> "" and m_sClass <> "" Then
'        response.write m_sKamoku
'    end if

End Function

Sub s_SetKyositu()  '//2001/07/30�ύX
'********************************************************************************
'*  [�@�\]  �l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �����R�[�h��ݒ�
'********************************************************************************

    On Error Resume Next
    Err.Clear

    m_iKyosituCd = ""
    Do Until m_Rs.EOF
        if CInt(m_Rs("T20_YOUBI_CD")) = CInt(m_iYobiCnt) and CDbl(m_Rs("T20_JIGEN")) = CDbl(m_iJgnCnt) Then
            m_iKyosituCd = m_Rs("T20_KYOSITU")
        else
        end if
        
        m_Rs.MoveNext
    Loop
    m_Rs.MoveFirst

End Sub

'********************************************************************************
'*  [�@�\]  �������̎擾
'*  [����]  
'*  [�ߒl]  0:���擾�����A1:���R�[�h�Ȃ��A99:���s
'*  [����]  
'********************************************************************************
Function f_GetKyosituMei()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    Call s_SetKyositu()
    f_GetKyosituMei = 0
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
                'response.write w_iRet & "<br>"
                'm_sErrMsg = "ں��޾�Ă̎擾���s"
                'm_bErrFlg = True
                f_GetKyosituMei = 99
                Exit Do 'GOTO LABEL_f_GetKyosituMei_END
            Else
            End If
            
            If w_Rs.EOF Then
                '�Ώ�ں��ނȂ�
                'm_sErrMsg = "�Ώ�ں��ނȂ�"
                f_GetKyosituMei = 1
                Exit Do 'GOTO LABEL_f_GetKyosituMei_END
            End If
            
                '// �擾�����l���i�[
                    m_sKyositu = w_Rs("M06_KYOSITUMEI")    '//���������i�[
            '// ����I��
            Exit Do
        
        Loop
        
        gf_closeObject(w_Rs)
    
    end if

'// LABEL_f_GetKyosituMei_END
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
    w_i = p_iNendo - p_iGaknen + 1
    	w_sSQL = ""
    	w_sSQL = w_sSQL & vbCrLf & "SELECT "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKUMEI, "
    	w_sSQL = w_sSQL & vbCrLf & "T15_COURSE_CD"
    	w_sSQL = w_sSQL & vbCrLf & "FROM "
    	w_sSQL = w_sSQL & vbCrLf & " T15_RISYU "
    	w_sSQL = w_sSQL & vbCrLf & "WHERE "
    	w_sSQL = w_sSQL & vbCrLf & "T15_KAMOKU_CD = '" & p_sKamokuCD & "' AND "
    	w_sSQL = w_sSQL & vbCrLf & "T15_NYUNENDO = "& p_iNendo - cint(p_iGaknen) + 1 &" "
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

'********************************************************************************
'*  [�@�\]  �N���X�����擾����
'*  [����]  p_iNendo  �F�����N�x
'*          p_iGakuNen�F�w�N
'*  [�ߒl]  f_GetClassMax�F�N���X��
'*  [����]  
'********************************************************************************
Function f_GetClassMax(p_iNendo,p_iGakuNen)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassMax = 0

	Do

		'//�N���X���̎擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  COUNT(M05_CLASSNO) as ClassMax"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M05_NENDO=" & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND M05_GAKUNEN=" & p_iGakuNen

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
'response.write w_sSql&vbCrLf&"<BR>"
'response.write w_iRet
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//�N���X��
			f_GetClassMax = cint(rs("ClassMax"))
		End If

		Exit Do
	Loop

	'//�߂�l���
'	f_GetClassMax = rs("ClassMax")

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

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

    Dim w_cellT             '//Table�Z���F
    Dim w_sClass          '//�N���X��
    Dim w_sKamoku       '//�Ȗږ�
    Dim w_iJigenMax      '//�j����row�p
    Dim w_iJgnNo          '//�\���p����
	Dim w_daigaeF         '//��փt���O

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

<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr>
<td align="center">

    <table border=1 class=hyo width="100%">
    <COLGROUP WIDTH="5%" ALIGN=center>
    <COLGROUP WIDTH="5%" ALIGN=center>
    <COLGROUP WIDTH="20%" ALIGN=center>
    <COLGROUP WIDTH="35%" ALIGN=center>
    <COLGROUP WIDTH="35%" ALIGN=center>
    <tr>
        <th colspan="2" class=header><br></th>  
        <th class=header>�N���X</th>
        <th class=header>���Ȗ�</th>
        <th class=header>����</th>
    </tr>

<%
m_iYobiCCnt = 1

'For m_iYobiCnt = 2 to 6
For m_iYobiCnt = C_YOUBI_MIN to C_YOUBI_MAX

	m_Flg = 0
	w_iJigenMax =  f_KamokuSu(m_iYobiCnt)
    For m_iJgnCnt = 0.5 to m_iJMax step 0.5
		w_sClass =  f_ShowClass()
		w_sKamoku = f_ShowKamokuMei()
		w_iJgnNo = m_iJgnCnt
		w_daigaeF = 0	'��փt���O�̏�����

		If right(cstr(w_iJgnNo*10),1) <> "0" then	'0.5�����̏ꍇ
			w_iJgnNo = " "
		Else										'���ʂ̎��Ԃ̏ꍇ

			'�ʏ�̎��Ԋ��Ȗڂ��Ȃ��Ƃ��A��։Ȗڂ̎��Ԋ��擾
'			If w_sKamoku = "" and w_sClass = "" then 
				If f_GetDaigae(m_iYobiCnt,m_iJgnCnt,w_sKamoku,w_daigaeF) <> true then
					m_bErrFlg = True
					m_sErrMsg = "���R�[�h�Z�b�g�̎擾�Ɏ��s���܂���"
					Exit Sub
				End if

				if w_daigaeF = 1 then
					w_sClass = "��։Ȗ�"
					m_sKyositu = ""
				end if
'			End if

		End If

		If w_sKamoku <> "" OR w_sClass <> "" OR right(cstr(m_iJgnCnt*10),1) = "0" Then 
			call gs_cellPtn(w_cellT)

			'�������̎擾�i��։Ȗڂ������j
			if w_daigaeF <> 1 then 
				call f_GetKyosituMei()
'				    if f_GetKyosituMei = 0 Then
'				        response.write m_sKyositu
'				    else
'				    end if
			end if
			%>
			    <tr>
			<%call s_ShowYobi(w_iJigenMax)%>
			        <td class=<%=w_cellT%>><%=w_iJgnNo%></td>
			        <td class=<%=w_cellT%>><%=w_sClass%><br></td>
			        <td class=<%=w_cellT%>><%=w_sKamoku%><br></td>
			        <td class=<%=w_cellT%>><%=m_sKyositu%><br></td>
			    </tr>
			<%
		End If
    Next
m_iYobiCCnt = m_iYobiCCnt + 1   '//�j���J�E���g�i�e�[�u���w�i�F�\���p�j

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