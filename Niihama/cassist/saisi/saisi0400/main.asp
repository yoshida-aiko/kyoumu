<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍�����
' ��۸���ID : gak/gak0300/main.asp
' �@      �\: ���y�[�W �w�Ѓf�[�^�̌������ʂ�\������
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           txtHyoujiNendo         :�\���N�x
'           txtGakunen             :�w�N
'           txtGakkaCD             :�w��
'           txtClass               :�N���X
'           txtName                :����
'           txtGakusekiNo          :�w�Дԍ�
'           txtSeibetu             :����
'           txtGakuseiNo           :�w���ԍ�
'           txtIdou                :�ٓ�
'           txtTyuClub             :���w�Z�N���u
'           txtClub                :���݃N���u
'           txtRyoseiKbn           :��
'           CheckImage               :�摜�\���w��
'           txtMode                :���샂�[�h
'                               BLANK   :�����\��
'                               SEARCH  :���ʕ\��
' ��      ��:
'           �������\��
'               �^�C�g���̂ݕ\��
'           �����ʕ\��
'               ��y�[�W�Őݒ肳�ꂽ���������ɂ��Ȃ��w������\������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
	
    '�擾�����f�[�^�����ϐ�
    Public  m_TxtMode      	       ':���샂�[�h
    Public  m_PgMode      	       ':�v���O�������[�h
	Public  m_iSyoriNen      	   ':�����N�x
    Public  m_iHyoujiNendo         ':�\���N�x
    Public  m_sGakunen             ':�w�N
    Public  m_sGakkaCD             ':�w��
    Public  m_sClass               ':�N���X
    Public  m_sName                ':����
    Public  m_sGakusekiNo          ':�w�Дԍ�
    Public  m_sSeibetu             ':����
    Public  m_sGakuseiNo           ':�w���ԍ�
    Public  m_sIdou                ':�ٓ�
    Public  m_sTyuClub             ':���w�Z�N���u
    Public  m_sClub                ':���݃N���u
    Public  m_sRyoseiKbn           ':��
    Public  m_sCheckImage          ':�摜�\���w��
	Public  m_sTyugaku			   ':�o�g���w�Z
    Public  m_sMsgTitle            ':����
	  
    Public	m_Rs					'recordset
    Public	m_RsGakka
    
    Public	m_iDsp					'�ꗗ�\���s��

    Public  m_iPageTyu      		':�\���ϕ\���Ő��i�������g����󂯎������j

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
    w_sMsgTitle="���шꗗ"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_iDsp=15

	'//�Z�b�V�������E���샂�[�h�̎擾
	m_iSyoriNen = Session("NENDO")
    m_TxtMode=request("txtMode")
    m_PgMode=request("p_mode")
    
	Select Case m_PgMode
		Case "P_HAN0100"
		    m_sMsgTitle="���шꗗ�\"
		Case "P_KKS0200"
		    m_sMsgTitle="���ۈꗗ�\"
		Case "P_KKS0210"
		    m_sMsgTitle="�x���ꗗ�\"
		Case "P_KKS0220"
		    m_sMsgTitle="�s�����ۈꗗ�\"
		Case Else
	End Select
	w_sMsgTitle = m_sMsgTitle


    Do
		if m_TxtMode = "" then
           	Call showPage()
			Exit Do
		End if

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


        '// ���Ұ�SET
        Call s_SetParam()

        '�f�[�^���oSQL���쐬����
		'�N���X�������łr�p�k�����쐬����B
        Call s_MakeSQL(w_sSQL)
		
		'If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then
		'	m_bErrFlg = True
		'	exit do
		'End If
		
		
       '���R�[�h�Z�b�g�̎擾
        Set m_Rs = Server.CreateObject("ADODB.Recordset")

		w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do     'GOTO LABEL_MAIN_END
        End If

        '// �y�[�W��\��
        If m_Rs.EOF Then
            Call showPage_NoData()
        Else
			'PDF���\��
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
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

'    Session("HyoujiNendo") = request("txtHyoujiNendo")     	'�\���N�x
'    Session("HyoujiNendo") = Session("NENDO")		'�\���N�x	'<-- 8/16�C��	��
'
'	m_iDsp = cint(request("txtDisp"))						':�������X�g�̕\������
'
    '// BLANK�̏ꍇ�͍s���ر
    If m_TxtMode = "Search" Then
        m_iPageTyu = 1
    Else
        m_iPageTyu = int(Request("txtPageTyu"))     ':�\���ϕ\���Ő��i�������g����󂯎������j
    End If

End Sub


'********************************************************************************
'*  [�@�\]  PDF���f�[�^���oSQL������̍쐬
'*  [����]  p_sSql - SQL������
'*  [�ߒl]  �Ȃ� 
'*  [����]  
'********************************************************************************
Sub s_MakeSQL(p_sSql)

	Dim w_iRet
	Dim w_sId
	Dim w_sWhere
	Dim w_iGakunen
	Dim w_iClassNo
	Dim w_sGakka
	

	'����������ID���擾
	w_iRet = f_GetSyori(w_sId)

'response.write "getid" & w_sId


	w_sWhere = ""
	Select Case w_sId
		'//FULL����
        Case C_ID_SEI0200
			'Full�����̏ꍇ��Where�������w�肵�Ȃ��B
		'�w�ȕ� //////////////////////////////////////////////////////////////////////////////////////////////////
		'Add 2002.1.13 okada
        Case C_ID_SEI0210
        
			'//��������������w�Ȃ��擾
			if Not f_GetKyokanGakka(m_iSyoriNen,session("KYOKAN_CD"),w_sGakka) Then
					m_bErrFlg = True
					m_sErrMsg = "�����w�Ȃ̎擾�Ɏ��s���܂����B"
					Exit sub        
			End if
 

			'// �w�Ȃf���������m_RsGakka �̎擾
			if Not f_GetGakkaGrp(m_iSyoriNen,w_sGakka) Then
					m_bErrFlg = True
					m_sErrMsg = "�w�Ȃf�����̎擾�Ɏ��s���܂����B"
					Exit sub				
			End if
			w_sWhere = w_sWhere & "( "
			Do until m_RsGakka.EOF'//�߂�l���
	
				'// �w�N�ƃN���X���擾����B
				w_iClassNo = ""
				'w_iGakunen = ""
'response.write m_RsGakka("M23_GAKUNEN")
'Response.end				
				'�w�N�ƃN���X���擾����
				if Not f_GetClass(m_iSyoriNen,m_RsGakka("M23_GAKUNEN"),w_iClassNo,m_RsGakka("M23_GAKKA_CD")) then
					m_bErrFlg = True
					m_sErrMsg = "�N���X�f�[�^�̎擾�Ɏ��s���܂����B"
					Exit sub				
				end if
				
				w_sWhere = w_sWhere & "(T76_GAKUNEN = " & m_RsGakka("M23_GAKUNEN") & " And T76_CLASS = " & w_iClassNo & ") "				
								
				m_RsGakka.MoveNext
				if Not m_RsGakka.Eof then	'�l�������m����������ꍇ�͂n�q��ǉ�
					w_sWhere = w_sWhere & " or "
				Else
					w_sWhere = w_sWhere & " ) AND "
				End if
			Loop
			
			'//ں��޾��CLOSE
			Call gf_closeObject(m_RsGakka)
		'/////////////////////////////////////////////////////////////////////////////////////////////////////////////	

		'1�N��
        Case C_ID_SEI0221
			w_sWhere = " T76_GAKUNEN = 1 AND "
		'2�N��
        Case C_ID_SEI0222
			w_sWhere = " T76_GAKUNEN = 2 AND "
		'3�N��
        Case C_ID_SEI0223
			w_sWhere = " T76_GAKUNEN = 3 AND "
		'4�N��
        Case C_ID_SEI0224
			w_sWhere = " T76_GAKUNEN = 4 AND "
		'5�N��
        Case C_ID_SEI0225
			w_sWhere = " T76_GAKUNEN = 5 AND "

		'�S�C
        Case C_ID_SEI0230

			'//�S�C�����̏ꍇ�́A���̃��[�U�[���S�C���Ă���N���X
	        '
			If Not f_GetTanninClass(m_iSyoriNen,session("KYOKAN_CD"),w_iGakunen,w_iClassNo) Then
				m_bErrFlg = True
				m_sErrMsg = "�N���X�f�[�^�̎擾�Ɏ��s���܂����B"
				Exit sub
			End If

			w_sWhere = " T76_GAKUNEN = " & w_iGakunen & " AND "
			w_sWhere = w_sWhere & " T76_CLASS = " & w_iClassNo & " AND "

'response.write " Where�� " & w_sWhere
		
		Case Else

	End Select

'response.write " Where�� " & w_sWhere

    p_sSql = ""
    p_sSql = p_sSql & " SELECT "
    p_sSql = p_sSql & " T76_NENDO      , "
    p_sSql = p_sSql & " T76_GAKKI_KBN  , "
    p_sSql = p_sSql & " T76_TYOHYO_ID  , "
    p_sSql = p_sSql & " T76_TYOHYO_NAME, "
    p_sSql = p_sSql & " T76_SIKEN_KBN  , "
    p_sSql = p_sSql & " T76_SIKENMEI   , "
    p_sSql = p_sSql & " T76_GAKUNEN    , "
    p_sSql = p_sSql & " T76_CLASS      , "
    p_sSql = p_sSql & " T76_CLASSMEI   , "
    p_sSql = p_sSql & " T76_PATH       , "
    p_sSql = p_sSql & " T76_FILENAME   , "
    p_sSql = p_sSql & " T76_INS_DATE     "
    p_sSql = p_sSql & " FROM T76_PDF "
    p_sSql = p_sSql & " WHERE "
    p_sSql = p_sSql & w_sWhere
    p_sSql = p_sSql & " T76_NENDO = " & m_iSyoriNen & " AND "
    p_sSql = p_sSql & " T76_TYOHYO_ID = '" & m_PgMode & "' "
    p_sSql = p_sSql & " ORDER BY "
    p_sSql = p_sSql & " T76_TYOHYO_ID, "
    p_sSql = p_sSql & " T76_SIKEN_KBN DESC, "
    p_sSql = p_sSql & " T76_GAKUNEN ,"
    p_sSql = p_sSql & " T76_CLASS "

'response.write p_sSql & "<BR>"

End Sub

'********************************************************************************
'*  [�@�\]  ������ʎ擾
'*  [����]  p_iMenuID�F���ڂ�NO
'*  [�ߒl]  true/false�Ap_iMenuID:����ID
'*  [����]  ������Ă��錠���̌���ID���擾
'********************************************************************************
Function f_GetSyori(p_iMenuID)
	Dim w_sLevel
	Dim w_iRet,w_Rs,w_sSq
	Dim w_Where

	Dim w_iCnt
	
	f_GetSyori = false

	'// Session("LEVEL")��NULL�Ȃ�A�ʂ���
	if gf_IsNull(Trim(Session("LEVEL"))) then Exit Function

	'// Session("LEVEL")��"0"�Ȃ�A�ʂ���
	if Cint(Session("LEVEL")) = Cint(0) then Exit Function

	w_sLevel = "T51_LEVEL" & Trim(Session("LEVEL"))

	'// WHERE���쐬

    Do
		w_sSql = ""
		w_sSql = w_sSql & "Select "
		w_sSql = w_sSql & "T51_ID,"
		w_sSql = w_sSql &  w_sLevel & " "
		w_sSql = w_sSql & "From T51_SYORI_LEVEL "
		w_sSql = w_sSql & "Where "
		w_sSql = w_sSql & w_sLevel & " = 1 AND "
		w_sSql = w_sSql & "T51_ID in ("
		w_sSql = w_sSql & "'" & C_ID_SEI0200 & "',"	'FULL����
		w_sSql = w_sSql & "'" & C_ID_SEI0210 & "',"	'�w�ȕ�
		w_sSql = w_sSql & "'" & C_ID_SEI0221 & "',"	'1�N��
		w_sSql = w_sSql & "'" & C_ID_SEI0222 & "',"	'2�N��
		w_sSql = w_sSql & "'" & C_ID_SEI0223 & "',"	'3�N��
		w_sSql = w_sSql & "'" & C_ID_SEI0224 & "',"	'4�N��
		w_sSql = w_sSql & "'" & C_ID_SEI0225 & "',"	'5�N��
		w_sSql = w_sSql & "'" & C_ID_SEI0230 & "' "	'�S�C
		w_sSql = w_sSql & ") "
		w_sSql = w_sSql & " ORDER BY T51_ID "

'response.write " �����擾p_sSql=" & w_sSql & "<BR>"
		w_iRet = gf_GetRecordset(w_Rs, w_sSql)

		If w_iRet <> 0 Then
		    'ں��޾�Ă̎擾���s
		    Exit Do
		End If

		If w_Rs.EOF = true Then
		    '�Y������
		    Exit Do
		End If

		w_flg = false
		w_Rs.movefirst
		Do Until w_Rs.EOF
			If trim(gf_SetNull2String(w_Rs(w_sLevel))) = "1" then 

				p_iMenuID = trim(gf_SetNull2String(w_Rs("T51_ID")))

				w_flg = true
				Exit Do
			End If

'response.write "OK!" & w_iCnt & "<BR>"

			w_Rs.movenext
		Loop

		w_Rs.close
		Set w_Rs = Nothing

		If w_flg <> true Then
		    '�Ώ�ں��ނȂ�
		    Exit Do
		End If

		f_GetSyori = true

		'// ����I��
		Exit Do

    Loop

End Function

'********************************************************************************
'*  [�@�\]  �N���X�R�[�h�Ɗw�N���擾����
'*  [����]  p_iNendo   �F�����N�x
'*          p_sKyokanCd�F�����R�[�h
'*          p_iGakunen �F�w�N
'*          p_iClassNo �F�N���XNO
'*  [�ߒl]  gf_GetClassName�F�N���X��
'*  [����]  
'********************************************************************************
Function f_GetTanninClass(p_iNendo,p_sKyokanCd,p_iGakunen,p_iClassNo)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetTanninClass = False

	p_iGakunen = 0
	p_iClassNo = 0
	
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " SELECT "
	w_sSql = w_sSql & vbCrLf & "  M05_GAKUNEN,"
	w_sSql = w_sSql & vbCrLf & "  M05_CLASSNO "
	w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
	w_sSql = w_sSql & vbCrLf & " WHERE "
	w_sSql = w_sSql & vbCrLf & "      M05_NENDO=" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "  AND M05_TANNIN=" & p_sKyokanCd

	'//ں��޾�Ď擾
	w_iRet = gf_GetRecordset(rs, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit Function
	End If

	'//�f�[�^���擾�ł����Ƃ�
	If rs.EOF = False Then
		p_iGakunen = rs("M05_GAKUNEN")	'�w�N
		p_iClassNo = rs("M05_CLASSNO")	'�N���XNO
	End If

	'//�߂�l���
	f_GetTanninClass = True

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �����̏�������w�Ȃ��擾
'*  [����]  p_iNendo   �F�����N�x
'*          p_sKyokanCd�F�����R�[�h
'*          p_iGakkaCd �F�w�ȃR�[�h
'*  [�ߒl]  gf_GetClassName�Fp_iGakkaCd �F�w�ȃR�[�h
'*  [����]  
'********************************************************************************
Function f_GetKyokanGakka(p_iNendo,p_sKyokanCd,p_sGakkaCd)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetKyokanGakka = False

	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " SELECT "
	w_sSql = w_sSql & vbCrLf & "  M04_GAKKA_CD "
	w_sSql = w_sSql & vbCrLf & " FROM M04_KYOKAN"
	w_sSql = w_sSql & vbCrLf & " WHERE "
	w_sSql = w_sSql & vbCrLf & "      M04_NENDO=" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "  AND M04_KYOKAN_CD =" & p_sKyokanCd

	'//ں��޾�Ď擾
	w_iRet = gf_GetRecordset(rs, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit Function
	End If

	'//�f�[�^���擾�ł����Ƃ�
	If rs.EOF = False Then
		p_sGakkaCd = rs("M04_GAKKA_CD")	'�w�ȃR�[�h
	End If

	'//�߂�l���
	f_GetKyokanGakka = True

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �w�Ȃb�c����w�Ȃf�������擾
'*  [����]  p_iNendo   �F�����N�x
'*          p_sKyokanCd�F�����R�[�h
'*          p_iGakkaCd �F�w�ȃR�[�h
'*  [�ߒl]  gf_GetClassName�Fp_iGakkaCd �F�w�ȃR�[�h
'*  [����]  
'********************************************************************************
Function f_GetGakkaGrp(p_iNendo,p_sGakkaCd)
	Dim w_iRet
	Dim w_sSQL
	Dim rs
	Dim w_sGakaGrp	

	On Error Resume Next
	Err.Clear

	f_GetGakkaGrp = False

	'�w�Ȃb�c����f�������擾
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " Select "
	w_sSql = w_sSql & vbCrLf & "	M23_GROUP,"
	w_sSql = w_sSql & vbCrLf & "	M23_GAKKA_CD "
	w_sSql = w_sSql & vbCrLf & " From "
	w_sSql = w_sSql & vbCrLf & "	M23_GAKKA_GRP "
	w_sSql = w_sSql & vbCrLf & " Where "
	w_sSql = w_sSql & vbCrLf & "	M23_NENDO =" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "	AND M23_GAKKA_CD =" & p_sGakkaCd
	w_sSql = w_sSql & vbCrLf & " Order By M23_GROUP "

'response.write w_sSql
 
	'//ں��޾�Ď擾
	w_iRet = gf_GetRecordset(rs, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit Function
	End If

	'//�f�[�^���擾�ł����Ƃ�
	If rs.EOF = False Then
		w_sGakaGrp = rs("M23_GROUP")	'�w�ȃR�[�h
	End If
	
'response.write "OK!" & w_sGakaGrp 
'Response.End 	
	
	'���̂f�����ɏ������Ă���w�Ȃ��擾���Ȃ����B
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " Select "
	w_sSql = w_sSql & vbCrLf & "	M23_GROUP,"
	w_sSql = w_sSql & vbCrLf & "	M23_GAKUNEN,"
	w_sSql = w_sSql & vbCrLf & "	M23_GAKKA_CD "
	w_sSql = w_sSql & vbCrLf & " From "
	w_sSql = w_sSql & vbCrLf & "	M23_GAKKA_GRP "
	w_sSql = w_sSql & vbCrLf & " Where "
	w_sSql = w_sSql & vbCrLf & "	M23_NENDO =" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "	AND M23_GROUP =" & w_sGakaGrp
	w_sSql = w_sSql & vbCrLf & " Order By M23_GROUP "

'response.write "OK!" & w_sSql 	
'Response.End 		
	sSet m_RsGakka = Server.CreateObject("ADODB.Recordset")
	'//ں��޾�Ď擾
	w_iRet = gf_GetRecordset(m_RsGakka, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit Function
	End If

	'//�f�[�^���擾�ł����Ƃ� aaaa
	If m_RsGakka.EOF = False Then
	
		'//�߂�l���
		f_GetGakkaGrp = True
		
	End If	
'response.write "OK!" & "A"	
'Response.End 			
 			
	'//ں��޾��CLOSE
	Call gf_closeObject(rs)
'response.write "OK!" & "B"

End Function

'********************************************************************************
'*  [�@�\]  �w�Ȃb�c����N���X�R�[�h���擾����
'*  [����]  p_iNendo   �F�����N�x
'*          p_sKyokanCd�F�����R�[�h
'*          p_iGakunen �F�w�N
'*          p_iClassNo �F�N���XNO
'*  [�ߒl]  gf_GetClassName�F�N���X��
'*  [����]  
'********************************************************************************
Function f_GetClass(p_iNendo,p_iGakunen,p_iClassNo,p_iGakkaCD)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClass = False

	'p_iGakunen = 0
	'p_iClassNo = 0
	
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & " SELECT "
	w_sSql = w_sSql & vbCrLf & "  M05_CLASSNO "
	w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
	w_sSql = w_sSql & vbCrLf & " WHERE "
	w_sSql = w_sSql & vbCrLf & "      M05_NENDO=" & p_iNendo
	w_sSql = w_sSql & vbCrLf & "  AND M05_GAKUNEN=" & p_iGakunen
	w_sSql = w_sSql & vbCrLf & "  AND M05_GAKKA_CD=" & p_iGakkaCD

'response.write w_sSql
	
	'//ں��޾�Ď擾
	w_iRet = gf_GetRecordset(rs, w_sSQL)
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit Function
	End If

	'//�f�[�^���擾�ł����Ƃ�
	If rs.EOF = False Then
		p_iClassNo = rs("M05_CLASSNO")	'�N���XNO
	End If

	'//�߂�l���
	f_GetClass = True

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage_NoData()

%>
	<html>
	<head>
	<title><%=m_sMsgTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
	<link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>

	<body>

	<center>
		<br><br><br>
		<span class="msg">�Ώۃf�[�^�͑��݂��܂���</span>
	</center>

	</body>
    </html>

<%
    '---------- HTML END   ----------
End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	Dim w_pageBar			'�y�[�WBAR�\���p
%>

<html>

<head>
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

    }

    //************************************************************
    //  [�@�\]  �ڍ׃{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_detail(p_sUrl){

			url = p_sUrl;
			w   = 1015;
			h   = 710;
			wn  = "SubWindow";
			//opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
			opt = "left=0,top=0,directoris=no,location=no,menubar=no,scrollbars=yes,status=no,toolbar=no,resizable=yes";
			if (w > 0)
				opt = opt + ",width=" + w;
			if (h > 0)
				opt = opt + ",height=" + h;
			newWin = window.open(url, wn, opt);

    }

    //************************************************************
    //  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_PageClick(p_iPage){

        document.frm.action="main.asp?p_mode=<%=m_PgMode%>";
        document.frm.target="_self";
        document.frm.txtMode.value = "PAGE";
        document.frm.txtPageTyu.value = p_iPage;
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>
    </head>

    <body>
	<% if m_TxtMode = "" then %>
		<center>
		<br><br><br>
		<span class="msg">���ڂ�I��ŕ\���{�^���������Ă�������</span>
		</center>
	<% Else %>
	    <div align="center">
	    <form action="kojin.asp" method="post" name="frm" target="_detail">

		<BR>
		<table><tr><td align="center">
		<%
			'�y�[�WBAR�\��
			Call gs_pageBar(m_Rs,m_iPageTyu,m_iDsp,w_pageBar)
		%>
		<%=w_pageBar %>

			<table border="0" width="100%">
				<tr>
					<td align="center">
					<% if m_TxtMode = "" then %>
						<table border="0" cellpadding="1" cellspacing="1" bordercolor="#886688" width="800">
							<tr>
								<td width="60">&nbsp</td>
								<td valign="top"></td>
							</tr>
						</table>
					<% else %>
						<% dim w_cell %>

					    <!--  PDF���\���@-->

                        <font color="red">�u�ڍׁv�{�^�����N���b�N���āA���шꗗ���\������Ȃ��A���̓G���[���b�Z�[�W���\�������ꍇ�́A<br></font>
                        <font color="#3366CC">Download</font><font color="red">���}�E�X�ŉE�N���b�N���āu�Ώۂ��t�@�C���ɕۑ��v��I�����Ă��������B<br><br></font>

						<table border="1" width="610" class=hyo>
							<tr>
								<th align="center" height=16 class=header width="200"nowrap>���@��</th>
								<th align="center" height=16 class=header width="60"nowrap>�w�N</th>
								<th align="center" height=16 class=header width="80"nowrap>�N���X</th>
								<th align="center" height=16 class=header width="125"nowrap>��@���@��</th>
								<th align="center" height=16 class=header width="45"nowrap>�ڍ�</th>
								<th align="center" height=16 class=header width="100"nowrap >�_�E�����[�h</th>
							</tr>

				        	<%
							w_iCnt = 1
							Do Until m_Rs.EOF or w_iCnt > m_iDsp
								call gs_cellPtn(w_cell)
							%>
								<tr>
									<td align="center" height="16" class=<%=w_cell%> width="200"nowrap><%=gf_HTMLTableSTR(m_Rs("T76_SIKENMEI")) %>&nbsp</td>
									<td align="center" height="16" class=<%=w_cell%> width="60"nowrap><%=gf_HTMLTableSTR(m_Rs("T76_GAKUNEN")) %>&nbsp</td>
									<td align="center" height="16" class=<%=w_cell%> width="80"nowrap><%=gf_HTMLTableSTR(m_Rs("T76_CLASSMEI")) %>&nbsp</td>
									<td align="center" height="16" class=<%=w_cell%> width="125"nowrap><%=gf_HTMLTableSTR(m_Rs("T76_INS_DATE")) %>&nbsp</td>
									<td align="center" height="16" class=<%=w_cell%> width="45"nowrap><input type=button class=button value="�ڍ�" onclick="f_detail('../..<%=gf_HTMLTableSTR(m_Rs("T76_PATH")) %><%=gf_HTMLTableSTR(m_Rs("T76_FILENAME")) %>');"></td>
									<td align="center" height="16" class=<%=w_cell%> width="100"nowrap><a href='../..<%=gf_HTMLTableSTR(m_Rs("T76_PATH")) %><%=gf_HTMLTableSTR(m_Rs("T76_FILENAME")) %>'>Download</a></td>
								</tr>
							<%
								w_iCnt = w_iCnt + 1
								m_Rs.MoveNext
							Loop
							%>

						</table>

					<% end if %>
				</td>
			</tr>
		</table>

		<%=w_pageBar %>
		</td></tr></table>

		</div>
	    <input type="hidden" name="txtMode">
	    <input type="hidden" name="txtPageTyu" value="<%=m_iPageTyu%>">
	    <input type="hidden" name="hidGAKUSEI_NO">

		<%' �������� %>
		</form>
	<% End if %>
	</body>
    </html>


<%
    '---------- HTML END   ----------
End Sub

%>

