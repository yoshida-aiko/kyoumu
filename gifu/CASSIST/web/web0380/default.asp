<%@ Language=VBScript %>
<% Response.Expires = 0%>
<% Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �ٓ��󋵈ꗗ
' ��۸���ID : web/web0380/default.asp
' �@      �\: �ٓ��󋵈ꗗ���o���B
'-------------------------------------------------------------------------
' ��      ��:SESSION(""):�����R�[�h     ��      SESSION���
' ��      ��:�Ȃ�
' ��      �n:SESSION(""):�����R�[�h     ��      SESSION���
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/09/3 �J�e
' ��      �X: 2002/02/20 ���c
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public m_bErrFlg           '�װ�׸�
    Public m_iNendo
    Public m_rs
    Public m_sZenki_Start		'�O������

    Public m_iNowPage			'�����߰�����ް
	Public m_iPagesize			'�\������
	m_iPagesize = C_PAGE_LINE


  '********** �\���p�z�� **********
    Public m_sSimei()		'����
    Public m_sNendo()		'�N�x    
    Public m_sGakuNo()		'�w���ԍ�
    Public m_sGakuseiNo()	'�w�Дԍ�
    Public m_sGakunen()		'�w�N
    Public m_sGakka()		'�w��
    Public m_sClass()		'�N���X(�g)
    Public m_sJiyu()		'�ٓ����R
    Public m_sHiduke()		'���t(�J�n���j
    Public m_sEHiduke()		'���t(�I�����j    
    Public m_sBiko()		'���l

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()
response.end
'///////////////////////////�@�d�m�c�@/////////////////////////////


'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	Dim w_lblMeisyo			'// �P�N�Ԕԍ��̖��̎擾�p
	
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�ٓ��󋵈ꗗ"
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

		'// �����`�F�b�N�Ɏg�p
'		session("PRJ_No") = "WEB0380"

		'// �s���A�N�Z�X�`�F�b�N
'		Call gf_userChk(session("PRJ_No"))

		'// �ϐ�������
		call f_paraSet()

		'// �ٓ��L�̊w���擾
		If f_GetidoGaku() <> true then
			'�f�[�^�擾���s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�̎擾�Ɏ��s���܂����B"
			Exit Do
		End If
				
        If cint(gf_GetRsCount(m_rs)) = 0 Then
            '�ٓ��f�[�^���Ȃ�
	        Call showNoPage()
            Exit Do
        End If

		'�ٓ��󋵂�z��ɑ��
		If f_InsAry() <> true then
			'�f�[�^�擾���s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�̎擾�Ɏ��s���܂����B"
			Exit Do
		End If
					
		'�f�[�^�̃\�[�g
		call s_sortBubble()
						
        '// �y�[�W��\��
        Call showPage()
     
        Exit Do
    Loop

    '// �I������
    Call gf_closeObject(m_rs)
    Call gs_CloseDatabase()

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
End Sub

'*******************************************************************************
' �@�@�@�\�F�ϐ��̏������Ƒ��
' ���@�@���F�Ȃ�
' �@�\�ڍׁF
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
Sub f_paraSet()

	m_iNendo = session("NENDO")

	'// �\���߰�����ް
	m_iNowPage = Request("hidPageNo")
	if gf_IsNull(m_iNowPage) then
		m_iNowPage = 1
	End if

	m_iNowPage = Cint(m_iNowPage)

End Sub


'*******************************************************************************
' �@�@�@�\�F�w�ȃO���[�v�̎擾
' �ԁ@�@�l�FTRUE:OK / FALSE:NG
' ���@�@���Fp_sGakkaGrp - �w�ȃO���[�v
' �@�@�@�@�@p_sNendo - �N�x
' �@�\�ڍׁF�w�ȃO���[�v�̎擾
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/07/27�@�c��
' �ρ@�@�X�F2001/08/28�@�J�e
'*******************************************************************************
Function f_GetidoGaku()

    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt
    
    f_GetidoGaku = False
    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "T11_SIMEI,T13_GAKUSEI_NO,T13_GAKUNEN, "
    w_sSQL = w_sSQL & "T13_GAKUSEKI_NO,T13_NENDO, "    
    w_sSQL = w_sSQL & "M02_GAKKAMEI,M05_CLASSRYAKU as M05_CLASSMEI,T13_IDOU_NUM "
    w_sSQL = w_sSQL & "FROM T13_GAKU_NEN,T11_GAKUSEKI,M02_GAKKA,M05_CLASS "
    w_sSQL = w_sSQL & "WHERE "
    w_sSQL = w_sSQL & " T13_NENDO <= " & m_iNendo & " AND "
    w_sSQL = w_sSQL & " T13_GAKUSEI_NO = T11_GAKUSEI_NO AND "
    w_sSQL = w_sSQL & " T13_NENDO = M02_NENDO AND "
    w_sSQL = w_sSQL & " T13_GAKKA_CD = M02_GAKKA_CD AND "
    w_sSQL = w_sSQL & " T13_NENDO = M05_NENDO AND "
    w_sSQL = w_sSQL & " T13_GAKUNEN = M05_GAKUNEN AND "
    w_sSQL = w_sSQL & " T13_CLASS = M05_CLASSNO AND "
    w_sSQL = w_sSQL & " T13_IDOU_NUM > 0 "
    w_sSQL = w_sSQL & "ORDER BY T13_GAKUSEI_NO ,T13_NENDO"

    '== ں��޾�Ď擾 ==
    w_iRet = gf_GetRecordset_OpenStatic(m_rs, w_sSQL)
    If w_iRet <> 0 Then
        '== �擾����Ȃ������ꍇ ==
        Exit function
    End If
    f_GetidoGaku = True
    Exit Function
    
End Function

'********************************************************************************
'*  [�@�\]  �O���E��������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sGakki		:�w��CD
'*			p_sZenki_Start	:�O���J�n��
'*			p_sKouki_Start	:����J�n��
'*			p_sKouki_End	:����I����
'*  [����]  
'********************************************************************************
Function f_GetZenki_Start(p_iNendo,p_sZenki_Start)

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	p_sZenki_Start = ""

	'�Ǘ��}�X�^����w�������擾
	w_sSQL = ""
	w_sSQL = w_sSQL & vbCrLf & " SELECT "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO, "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NO, "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_KANRI, "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_BIKO"
	w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
	w_sSQL = w_sSQL & vbCrLf & " WHERE "
	w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO=" & p_iNendo & " AND "
	w_sSQL = w_sSQL & vbCrLf & "   (M00_KANRI.M00_NO=" & C_K_ZEN_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_SYURYO & ") "  '//[M00_NO]10:�O���J�n 11:����J�n

	iRet = gf_GetRecordset(rs, w_sSQL)
	If iRet <> 0 Then
	    'ں��޾�Ă̎擾���s
	    m_bErrMsg = Err.description
	    Exit Function
	End If

	If rs.EOF = False Then
	    Do Until rs.EOF
	        If cInt(rs("M00_NO")) = C_K_ZEN_KAISI Then
	            p_sZenki_Start = rs("M00_KANRI")
	        End If
	        rs.MoveNext
	    Loop
	End If

    Call gf_closeObject(rs)

End Function

'*******************************************************************************
' �@�@�@�\�F�z��Ƀf�[�^��}���B
' �ԁ@�@�l�FTRUE:OK / FALSE:NG
' ���@�@���F
' �@�\�ڍׁF�z��ɕ\���p�̃f�[�^��}��
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/28�@�J�e
'*******************************************************************************
Function f_InsAry()

    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt
    Dim w_rs
    
    f_InsAry = False

    w_iCnt = 1
    Do Until m_Rs.EOF

	for i = 1 to cint(m_Rs("T13_IDOU_NUM"))

	    w_sSQL = ""
	    w_sSQL = w_sSQL & "SELECT "	    
	    w_sSQL = w_sSQL & "M01_SYOBUNRUIMEI as IDO_KBN,"
	    w_sSQL = w_sSQL & "T13_IDOU_BI_" & i & " as IDO_BI,"
	    w_sSQL = w_sSQL & "T13_IDOU_ENDBI_" & i & " as IDO_BI_E,"
	    w_sSQL = w_sSQL & "T13_IDOU_BIK_" & i & " as IDO_BIK "
	    w_sSQL = w_sSQL & "FROM T13_GAKU_NEN, M01_KUBUN "
	    w_sSQL = w_sSQL & "WHERE "
'	    w_sSQL = w_sSQL & " T13_NENDO <= " & m_iNendo & " AND "
	    w_sSQL = w_sSQL & " T13_NENDO = " & m_rs("T13_NENDO") & " AND "
	    w_sSQL = w_sSQL & " M01_NENDO = " & m_iNendo & " AND "
	    w_sSQL = w_sSQL & " T13_GAKUSEI_NO = '"& m_rs("T13_GAKUSEI_NO") &"' AND "
	    w_sSQL = w_sSQL & " T13_GAKUSEKI_NO = '"& m_rs("T13_GAKUSEKI_NO") &"' AND "	    
	    w_sSQL = w_sSQL & " M01_DAIBUNRUI_CD = '"& C_IDO &"' AND "
	    w_sSQL = w_sSQL & " M01_SYOBUNRUI_CD = T13_IDOU_KBN_"& i &" "

		'== ں��޾�Ď擾 ==
		w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
		If w_iRet <> 0 Then
			'== �擾����Ȃ������ꍇ ==
			Exit function
		End If

		redim Preserve m_sSimei(w_iCnt)
		redim Preserve m_sNendo(w_iCnt)	    
		redim Preserve m_sGakuNo(w_iCnt)
		redim Preserve m_sGakuseiNo(w_iCnt)
		redim Preserve m_sGakunen(w_iCnt)
		redim Preserve m_sGakka(w_iCnt)
		redim Preserve m_sClass(w_iCnt)
		redim Preserve m_sJiyu(w_iCnt)
		redim Preserve m_sHiduke(w_iCnt)
		redim Preserve m_sEHiduke(w_iCnt)	    
		redim Preserve m_sBiko(w_iCnt)

		'// ����ں��ނƑO��ں��ނ́u�J�n���t�E�ٓ����R�E�w��NO�v�������ł͂Ȃ��ꍇ
		if Cstr(w_rs("IDO_BI") & w_rs("IDO_KBN") & m_rs("T13_GAKUSEI_NO")) <> Cstr(m_sHiduke(w_iCnt-1) & m_sJiyu(w_iCnt-1) & m_sGakuseiNo(w_iCnt-1)) then

			'// �O���J�n�����擾
			Call f_GetZenki_Start(Cint(left(w_rs("IDO_BI"),4)),m_sZenki_Start)

			'// �����N�x�̎��N�x�̑O���J�n���f�[�^�����݂��Ȃ��ꍇ�A4/1���Z�b�g����B
            if m_sZenki_Start = "" AND Cint(left(w_rs("IDO_BI"),4)) > cint(m_inendo) then
                  m_sZenki_Start = m_inendo & "/04/01"
    		end if

			'// �O���J�n�����J�n�����O��������A�N�x��-1����
			if right(gf_YYYY_MM_DD(m_sZenki_Start,"/"),5) > right(gf_YYYY_MM_DD(w_rs("IDO_BI"),"/"),5) then
			    m_sNendo(w_iCnt)	= Cint(left(w_rs("IDO_BI"),4)) - 1
			Else
			    m_sNendo(w_iCnt)	= left(w_rs("IDO_BI"),4)
			End if

			'// �z��ɾ�Ă���
		    m_sSimei(w_iCnt)		= m_rs("T11_SIMEI")
		    m_sGakunen(w_iCnt)		= m_rs("T13_GAKUNEN")
		    m_sGakka(w_iCnt)		= m_rs("M02_GAKKAMEI")
		    m_sClass(w_iCnt)		= m_rs("M05_CLASSMEI")
		    m_sGakuNo(w_iCnt)		= m_rs("T13_GAKUSEKI_NO")
			m_sGakuseiNo(w_iCnt)    = m_rs("T13_GAKUSEI_NO")

		    m_sJiyu(w_iCnt)			= w_rs("IDO_KBN")
		    m_sHiduke(w_iCnt)		= w_rs("IDO_BI")
		    m_sEHiduke(w_iCnt)		= w_rs("IDO_BI_E")
		    m_sBiko(w_iCnt)			= w_rs("IDO_BIK")

			w_iCnt = w_iCnt +1
		End if

	    Call gf_closeObject(w_rs)

	next
	m_rs.MoveNext
    loop
	
	'//ں��޾��CLOSE

    f_InsAry = True
    Exit Function
    
End Function


'*******************************************************************************
' �@�@�@�\�F�o�u���\�[�g
' �ԁ@�@�l�FTRUE:OK / FALSE:NG
' ���@�@���F
' �@�\�ڍׁF���t�Ń\�[�g���܂��B
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/28�@�J�e
' �ρ@�@�X�F2002/02/20�@���c�F���t�E�w�Дԍ����Ƃ���
'*******************************************************************************
sub s_sortBubble() 

    Dim i
    Dim j
	Dim loopc
	    
    For i = 0 To UBound(m_sHiduke) - 1
        For j = i + 1 To UBound(m_sHiduke)
            If m_sHiduke(j) > m_sHiduke(i) Then
                Call s_swap(m_sSimei(i),m_sSimei(j))
                Call s_swap(m_sNendo(i),m_sNendo(j))                
                Call s_swap(m_sGakuNo(i),m_sGakuNo(j))
                Call s_swap(m_sGakunen(i),m_sGakunen(j))
                Call s_swap(m_sGakka(i),m_sGakka(j))
                Call s_swap(m_sClass(i),m_sClass(j))
                Call s_swap(m_sJiyu(i),m_sJiyu(j))
                Call s_swap(m_sHiduke(i),m_sHiduke(j))
                Call s_swap(m_sEHiduke(i),m_sEHiduke(j))                
                Call s_swap(m_sBiko(i),m_sBiko(j))
            End If          
        Next
    Next


    
    For i = 0 To UBound(m_sHiduke) - 1
        For j = i + 1 To UBound(m_sHiduke)
            If m_sHiduke(j) = m_sHiduke(i) Then 
				For loopc = i  To UBound(m_sHiduke)    
					If m_sHiduke(i) = m_sHiduke(loopc) Then
						If m_sGakuNo(i) > m_sGakuNo(loopc) Then
						    Call s_swap(m_sSimei(i),m_sSimei(loopc))
						    Call s_swap(m_sNendo(i),m_sNendo(loopc))
						    Call s_swap(m_sGakuNo(i),m_sGakuNo(loopc))
						    Call s_swap(m_sGakunen(i),m_sGakunen(loopc))
						    Call s_swap(m_sGakka(i),m_sGakka(loopc))
						    Call s_swap(m_sClass(i),m_sClass(loopc))
						    Call s_swap(m_sJiyu(i),m_sJiyu(loopc))
						    Call s_swap(m_sHiduke(i),m_sHiduke(loopc))
						    Call s_swap(m_sEHiduke(i),m_sEHiduke(loopc))                
						    Call s_swap(m_sBiko(i),m_sBiko(loopc))
						End If	 						
					End If
				Next				
            End If            
        Next
    Next    
End Sub


'*******************************************************************************
' �@�@�@�\�F�X���b�v
' �ԁ@�@�l�FTRUE:OK / FALSE:NG
' ���@�@���F�`�C�a
' �@�\�ڍׁF�`�C�a�̒��g�����ւ���B
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/28�@�J�e
'*******************************************************************************
Sub s_swap(a,b) 

dim tmp
	tmp = a
	a = b
	b = tmp
End sub


'********************************************************************************
'*  [�@�\]  �y�[�W�֌W�̕\���p�T�u���[�`��
'*  [����]  p_iRsCnt        �Fں��޶���
'*          p_iPageCd       �F�y�[�W�ԍ�
'*          p_iDsp          �F1�y�[�W�̍ő�\�������B
'*  [�ߒl]  p_pageBar       �F�ł����y�[�W�o�[HTML
'*  [����]  
'********************************************************************************
Sub s_pageBar(p_iRsCnt,p_iPageCd,p_iDsp)

	Dim w_bNxt					'// NEXT�\���L��
	Dim w_bBfr					'// BEFORE�\���L��
	Dim w_iNxt					'// NEXT�\���Ő�
	Dim w_iBfr					'// BEFORE�\���Ő�
	Dim w_iCnt					'// �ް��\������
	Dim w_iMax					'// �ް��\������
	Dim i,w_iSt,w_iEd

	Dim w_iRecordCnt			'//���R�[�h�Z�b�g�J�E���g

	On Error Resume Next
	Err.Clear

	w_iCnt = 1
	w_bFlg = True

	'////////////////////////////////////////
	'�y�[�W�֌W�̐ݒ�
	'////////////////////////////////////////

	'���R�[�h�����擾
	w_iRecordCnt = p_iRsCnt
	w_iMax = int((Cint(p_iRsCnt) / p_iDsp) + 0.9)

	'EOF�̂Ƃ��̐ݒ�
	If Cint(p_iPageCd) >= w_iMax Then
		p_iPageCd = w_iMax
	End If

	'�O�y�[�W�̐ݒ�
	If Cint(p_iPageCd) = 1 Then
		w_bBfr = False
		w_iBfr = 0
	Else
		w_bBfr = True
		w_iBfr = Cint(p_iPageCd) - 1
	End If

	'��y�[�W�̐ݒ�
	If Cint(p_iPageCd) = w_iMax Then
		w_bNxt = False
		w_iNxt = Cint(p_iPageCd)
	Else
		w_bNxt = True
		w_iNxt = Cint(p_iPageCd) + 1
	End If

	'�y�[�W�̃��X�g�̎n��(w_iSt)�ƏI���(w_iEd)����
	'��{�I�ɑI������Ă���y�[�W(p_iPageCd)���^���ɗ���悤�ɂ���B
	w_iEd = Cint(p_iPageCd) + 5
	w_iSt = Cint(p_iPageCd) - 4

	'�y�[�W�̃��X�g��10�Ȃ����A�I���y�[�W�����X�g�̐^���ɂ��Ȃ��Ƃ��B
	If Cint(p_iPageCd) < 5 Then w_iEd = 10
	If w_iEd > w_iMax then w_iEd = w_iMax : w_iSt = w_iMax - 9
	If w_iSt < 1 or w_iMax < 10 then w_iSt = 1

	'////////////////////////////////////////
	'�y�[�W�֌W�̐ݒ�(�����܂�)
	'////////////////////////////////////////

	p_pageBar = ""
	p_pageBar = p_pageBar & vbCrLf & "<table border='0' width='100%'>"
	p_pageBar = p_pageBar & vbCrLf & "<tr>"
	p_pageBar = p_pageBar & vbCrLf & "<td align='left' width='10%'>"

	If w_bBfr = True Then
		p_pageBar = p_pageBar & vbCrLf & "<a href='javascript:f_PageClick("& w_iBfr &");' class='page'>�O��</a>"
	End If

	p_pageBar = p_pageBar & vbCrLf & " </td>"
	p_pageBar = p_pageBar & vbCrLf & "<td align=center width='80%'>"
	p_pageBar = p_pageBar & vbCrLf & " Page�F[ "

	for i = w_iSt to w_iEd
		If i = Cint(p_iPageCd) then 
			p_pageBar = p_pageBar & vbCrLf & "<span class='page'>" & i & "</span>"
		Else
			p_pageBar = p_pageBar & vbCrLf & "<a href='javascript:f_PageClick("& i &");' class='page'>" & i & "</a>"
		End If
	next

	p_pageBar = p_pageBar & vbCrLf & "/" & w_iMax & "] "
	p_pageBar = p_pageBar & vbCrLf & " Results�F" & w_iRecordCnt & "Hits"
	p_pageBar = p_pageBar & vbCrLf & "</td>"
	p_pageBar = p_pageBar & vbCrLf & "<td align='right' width='10%'> "

	If w_bNxt = True Then
		p_pageBar = p_pageBar & vbCrLf & "<a href='javascript:f_PageClick(" & w_iNxt & ")' class='page'>����</a>"
	End If

	p_pageBar = p_pageBar & vbCrLf & "</td>"
	p_pageBar = p_pageBar & vbCrLf & "</tr>"
	p_pageBar = p_pageBar & vbCrLf & "</table>"

	'// �����o��
	response.write p_pageBar

End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  2003.08.07:���c:loop�̏����������܂�ł����ׂɍŏI�f�[�^���\���ł��Ȃ�����
'********************************************************************************
sub s_writecell()
	dim i,w_cell
	Dim w_sMaeSimei

	'// ٰ�߶�����̏����l���擾
	i = ((m_iNowPage - 1) * m_iPagesize)

	'// ٰ�߶�����̍ő�ٰ�ߐ�
	iMax = (m_iNowPage * m_iPagesize)

	Do Until i > (Ubound(m_sHiduke)-1) or i >= iMax
		call gs_cellPtn(w_cell)
		 %>
		<TR>
			<TD class="<%=w_cell%>"><%=m_sNendo(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sHiduke(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sEHiduke(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sJiyu(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sSimei(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sGakuNo(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sGakunen(i)%>-<%=m_sClass(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sGakka(i)%></TD>
			<TD class="<%=w_cell%>"><%=m_sBiko(i)%></TD>
		</TR>
		<%
		i = i + 1
	Loop

End sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showNoPage()
%>
<html>
<head>
    <title>�ٓ��󋵈ꗗ</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
</head>
<body>
<center>
<%call gs_title("�ٓ��󋵈ꗗ","��@��")%>
<BR>
<span class="msg">���݁A�ٓ��҂͂��܂���B</span>
</center>
</body>
</html>
<%
End Sub


'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
%>
<html>
<head>
<title>�ٓ��󋵈ꗗ</title>
<link rel=stylesheet href=../../common/style.css type=text/css>
<SCRIPT LANGUAGE="javascript">
<!--

	//************************************************************
	//  [�@�\]  �ꗗ�\�̎��E�O�y�[�W��\������
	//  [����]  p_iPage :�\���Ő�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//
	//************************************************************
	function f_PageClick(p_iPage){

		document.frm.action = "default.asp";
		document.frm.target = "fTopMain";
		document.frm.hidPageNo.value = p_iPage;
		document.frm.submit();

	}
//-->
</SCRIPT>
</head>
<body>
<center>
<form name="frm" method="post">
<%call gs_title("�ٓ��󋵈ꗗ","��@��")%>

<BR>

<table class="hyo" border="1" width="">
	<tr>
		<th nowrap class="header">�ٓ��󋵈ꗗ</th>
		<td nowrap class="detail" align="center"><%=gf_fmtWareki(date())%>����</td>
	</tr>
</table>

<BR>
<table border=0 cellpaddin="0" cellspacing="0" width="98%">
	<tr><td><% Call s_pageBar((Ubound(m_sHiduke)-1),m_iNowPage,m_iPagesize) %></td></tr>
	<tr>
		<td>
			<table border=1 class="hyo" width="100%">
				<tr>
					<th class="header" width="5%" nowrap>�N�x</th>
					<th class="header" width="10%" nowrap>�J�n���t</th>
					<th class="header" width="10%" nowrap>�I�����t</th>
					<th class="header" nowrap>�ٓ����R</th>
					<th class="header" nowrap>����</th>
					<th class="header" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
					<th class="header" nowrap>�N���X</th>
					<th class="header" nowrap>�w��</th>
					<th class="header" width="15%">���l</th>
				</tr>
				<% call s_writecell() %>
			</table>
		</td>
	</tr>
	<tr><td><% Call s_pageBar((Ubound(m_sHiduke)-1),m_iNowPage,m_iPagesize) %></td></tr>
</table>

<input type="hidden" name="hidPageNo" value="<%= m_iNowPage %>">
</from>

</center>
</body>
</head>
</html>
<%
End Sub
%>