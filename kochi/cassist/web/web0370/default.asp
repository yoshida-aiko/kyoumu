<%@ Language=VBScript %>
<%Response.Expires = 0%>
<%Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w�����ꗗ
' ��۸���ID : web/web0370/default.asp
' �@      �\: �w�����̈ꗗ���o���B
'-------------------------------------------------------------------------
' ��      ��:SESSION(""):�����R�[�h     ��      SESSION���
' ��      ��:�Ȃ�
' ��      �n:SESSION(""):�����R�[�h     ��      SESSION���
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/08/29 �J�e
' ��      �X: 2015/08/27 ���� �ύX���e(�w�Ȗ��擾���̔N�x�A�N���X��擪1���̂ݎg�p����A�����N���X�l���ŁA0�l�̃N���X�͕\�����Ȃ�)
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public m_bErrFlg           '�װ�׸�
    Public m_iNendo
    Public m_iGrp_su
    Public m_ikongoCls()		'�����N���X�t���O
    Public m_sGakkaGrp()	'�w�ȃO���[�v
    Public m_gakka_cd()	'�w�ȃR�[�h
    Public m_gakkamei()	'�w�Ȗ�
    Public m_Qgakkamei()	'���w�Ȗ�
    Public m_Fld_M()	'�S�̐�
    Public m_Fld_F()	'�S�̐��i���q�j
    Public m_Fld_R()	'�S�̐��i���w���j
    Public m_Fld_MK()	'�x�w�Ґ�
    Public m_Fld_FK()	'�x�w�Ґ��i���q�j
    Public m_Fld_RK()	'�x�w�Ґ��i���w���j
    Public m_Cls_M()	'�����N���X
    Public m_Cls_F()		'�����N���X�i���q�j

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()
response.end
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
    w_sMsgTitle="�w�����ꗗ"
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
'		session("PRJ_No") = "WEB0370"

		'// �s���A�N�Z�X�`�F�b�N
'		Call gf_userChk(session("PRJ_No"))

		'// �ϐ�������
		call f_paraSet()
		
		'// �w�ȃO���[�v�擾
		If f_GetGakkaGrp() <> true then
	            '�f�[�^�擾���s
	            m_bErrFlg = True
	            m_sErrMsg = "�w�ȃf�[�^������܂���"
	            Exit Do
		End If
		
        If m_iGrp_su = 0 Then
            '�w�ȃf�[�^���Ȃ�
            m_bErrFlg = True
            m_sErrMsg = "�w�ȃf�[�^������܂���"
            Exit Do
        End If

		call f_arySet()
		
		'// �����N���X�擾
		If f_GetClass() <> true then
	            '�f�[�^�擾���s
	            m_bErrFlg = True
	            m_sErrMsg = "�����N���X�f�[�^������܂���"
	            Exit Do
		End If
		
		'// �f�[�^�̏W�v
		for i = 1 to m_iGrp_su
			if f_GetGakusei(i) <> true then
				'�f�[�^�擾���s
				m_bErrFlg = True
				m_sErrMsg = "�f�[�^������܂���B"
				Exit for
			End If
		next
		
		'// �w�Ȗ��擾
		If f_GetGakkaMei() <> true then
	            '�f�[�^�擾���s
	            m_bErrFlg = True
	            m_sErrMsg = "�w�Ȗ�������܂���B"
	            Exit Do
		End If
		
        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
		response.write w_sMsg
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()
End Sub

Sub f_paraSet()
'*******************************************************************************
' �@�@�@�\�F�ϐ��̏������Ƒ��
' ���@�@���F�Ȃ�
' �@�\�ڍׁF
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
m_iNendo = session("NENDO")
'm_iNendo = 2001

End Sub

Sub f_arySet()
'*******************************************************************************
' �@�@�@�\�F�ϐ��̏������Ƒ��
' ���@�@���F�Ȃ�
' �@�\�ڍׁF
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
Dim i,j
Redim m_gakka_cd(m_iGrp_su)	'�w�ȃR�[�h
Redim m_gakkamei(m_iGrp_su)	'�w�Ȗ�
Redim m_Qgakkamei(m_iGrp_su)	'���w�Ȗ�
Redim m_Fld_M(6,m_iGrp_su)	'�S�̐�
Redim m_Fld_F(6,m_iGrp_su)	'�S�̐��i���q�j
Redim m_Fld_R(6,m_iGrp_su)	'�S�̐��i���w���j
Redim m_Fld_MK(6,m_iGrp_su)	'�x�w�Ґ�
Redim m_Fld_FK(6,m_iGrp_su)	'�x�w�Ґ��i���q�j
Redim m_Fld_RK(6,m_iGrp_su)	'�x�w�Ґ��i���w���j

'/*�@�z��̏������@*/
for j = 0 to m_iGrp_su
	m_gakka_cd(j) = m_sGakkaGrp(j)
	m_gakkamei(j) = ""
	m_Qgakkamei(j) = ""
	for i = 0 to 6 
		m_Fld_M(i,j) = 0
		m_Fld_F(i,j) = 0
		m_Fld_R(i,j) = 0
		m_Fld_MK(i,j) = 0
		m_Fld_FK(i,j) = 0
		m_Fld_RK(i,j) = 0
	next 
next

End Sub

Function f_GetGakkaGrp()
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
    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt
    Dim w_rs
    
    f_GetGakkaGrp = False
    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M23_GROUP "
    w_sSQL = w_sSQL & "FROM M23_GAKKA_GRP "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M23_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M23_GAKKA_CD IS NOT NULL "
    w_sSQL = w_sSQL & "Group By M23_GROUP "
    w_sSQL = w_sSQL & "Order By M23_GROUP "
    w_sSQL = w_sSQL & ""

    '== ں��޾�Ď擾 ==
    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
    If w_iRet <> 0 Then
        '== �擾����Ȃ������ꍇ ==
        Exit function
    End If
    
    If w_rs.eof = True Then
        m_iGrp_su = 0
    
    Else
        '== ں��ތ����̎擾 ==
        m_iGrp_su = cint(gf_GetRsCount(w_rs))
'        m_iGrp_su = w_rs.RecordCount
    End If
    ReDim m_sGakkaGrp(m_iGrp_su)
   
    '== �w�ȃO���[�v�̃f�[�^���Z�b�g���� ==
    For w_iCnt = 1 to m_iGrp_su
        m_sGakkaGrp(w_iCnt) = w_rs("M23_GROUP")
        w_rs.MoveNext
    Next
    
	'//ں��޾��CLOSE
	Call gf_closeObject(w_rs)

    f_GetGakkaGrp = True
    Exit Function
    
End Function

Function f_GetClass()
'*******************************************************************************
' �@�@�@�\�F�����N���X��T���B
' �ԁ@�@�l�FTRUE:OK / FALSE:NG
' ���@�@���F
' �@�\�ڍׁF�����N���X��z��ɓ����
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/28�@�J�e
'*******************************************************************************
    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt,w_max_cls
    Dim w_rs
    
    f_GetClass = False
    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M05_GAKUNEN,Max(M05_CLASSNO) as MAX_CLS,M05_SYUBETU "
    w_sSQL = w_sSQL & "FROM M05_CLASS "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_GAKUNEN IS NOT NULL "
    w_sSQL = w_sSQL & "Group By M05_GAKUNEN,M05_SYUBETU "
    w_sSQL = w_sSQL & ""

    '== ں��޾�Ď擾 ==
    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
   If w_iRet <> 0 Then
        '== �擾����Ȃ������ꍇ ==
        Exit function
    End If
    
    If w_rs.eof = True Then
        w_iCnt = 0
    
    Else
        '== ں��ތ����̎擾 ==
        w_iCnt = cint(gf_GetRsCount(w_rs))
'        m_iGrp_su = w_rs.RecordCount
    End If

    '�����N���X�t���O�̔z��̏�����
    ReDim m_ikongoCls(w_iCnt)
   for each i in m_ikongoCls
    i=0
     next
    '== �����N���X�̃t���O���Z�b�g���� ==
    w_max_cls = 0
    Do Until w_rs.EOF
	w_igak = cint(w_rs("M05_GAKUNEN"))
        m_ikongoCls(w_igak) = w_rs("M05_SYUBETU")

	'�P�w�N�ł������N���X������΁A�t���O�𗧂Ă�B
	If cint(w_rs("M05_SYUBETU")) = 1 then m_ikongoCls(0) = 1

	If w_max_cls < cint(w_rs("MAX_CLS")) then 
		w_max_cls = cint(w_rs("MAX_CLS"))
	End If

        w_rs.MoveNext
    Loop

	'�W�v�p�z��̏�����
	ReDim m_Cls_M(w_iCnt,w_max_cls)
	ReDim m_Cls_F(w_iCnt,w_max_cls)
    
	    For i = 0 to w_iCnt
		    For j = 0 to w_max_cls
			m_Cls_M(i,j) = 0
			m_Cls_F(i,j) = 0
		    Next
	    Next

	'//ں��޾��CLOSE
	Call gf_closeObject(w_rs)


    f_GetClass = True
    Exit Function
    
End Function

Function f_GetClassMei(p_gak,p_cls)
'*******************************************************************************
' �@�@�@�\�F�����N���X���擾�B
' �ԁ@�@�l�FTRUE:OK / FALSE:NG
' ���@�@���F�N���X��
' �@�\�ڍׁF�����N���X�̃N���X�����o���B
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/28�@�J�e
'*******************************************************************************
    Dim w_sSQL
    Dim w_iRet
    Dim w_rs
    
    f_GetClassMei = ""
    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M05_CLASSMEI "
    w_sSQL = w_sSQL & "FROM M05_CLASS "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M05_NENDO = " & m_iNendo & " AND "
    w_sSQL = w_sSQL & "M05_GAKUNEN = " & p_gak & " AND "
    w_sSQL = w_sSQL & "M05_CLASSNO = " & p_cls & "1 "
    w_sSQL = w_sSQL & ""

    '== ں��޾�Ď擾 ==
    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
   If w_iRet <> 0 Then
        '== �擾����Ȃ������ꍇ ==
        Exit function
    End If
    If w_rs.eof = True Then
        f_GetClassMei = ""
    
    Else
	f_GetClassMei = w_rs("M05_CLASSMEI")
    End If
	'//ں��޾��CLOSE
	Call gf_closeObject(w_rs)
End Function

Function f_GetGakkaMei()
'*******************************************************************************
' �@�@�@�\�F�w�Ȗ��̃Z�b�g
' �ԁ@�@�l�FTRUE:OK / FALSE:NG
' ���@�@���Fp_sGakkaGrp - �w�ȃO���[�v
' �@�@�@�@�@p_sNendo - �N�x
' �@�\�ڍׁF�w�ȃO���[�v�̎擾
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/07/27�@�c��
' �ρ@�@�X�F2001/08/28�@�J�e
'*******************************************************************************
    Dim w_sSQL
    Dim w_iRet
    Dim w_rs
    Dim w_grp,w_gakka,w_gakkamei,w_gakunen

    w_grp = 0:w_gakka=0:w_gakkamei="":w_gakunen=""

    f_GetGakkaMei = false
    
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M23_GROUP,"
    w_sSQL = w_sSQL & "M02_GAKKA_CD, "
    w_sSQL = w_sSQL & "M02_GAKKAMEI, "
    w_sSQL = w_sSQL & "M23_GAKUNEN "
    w_sSQL = w_sSQL & "FROM "
    w_sSQL = w_sSQL & "M02_GAKKA, M23_GAKKA_GRP "
    w_sSQL = w_sSQL & "Where "
'    w_sSQL = w_sSQL & "M23_NENDO = 2000 And "
    w_sSQL = w_sSQL & "M23_NENDO = " & m_iNendo & " And "
    w_sSQL = w_sSQL & "M02_NENDO = M23_NENDO And "
    w_sSQL = w_sSQL & "M02_GAKKA_CD = M23_GAKKA_CD "
    w_sSQL = w_sSQL & "order by M23_GROUP,M23_GAKUNEN,M02_GAKKA_CD"

    '== ں��޾�Ď擾 ==
    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
    If w_iRet <> 0 Then
        '== �擾����Ȃ������ꍇ ==
        Exit function
    End If

    If w_rs.EOF = true then 
	m_bErrFlg = True
	m_sErrMsg = "�w�ȃf�[�^������܂���"
	Exit Function
    End If 

   w_rs.MoveFirst
   Do Until w_rs.EOF
     w_grp = cint(w_rs("M23_GROUP"))
     
     '*** �w�ȃO���[�v����������A�w�ȃR�[�h���Ⴄ�Ƃ� ***
     if w_group = cint(w_rs("M23_GROUP")) and w_gakka <> cint(w_rs("M02_GAKKA_CD")) then 

	     '*** �w�N���傫���������w�� ***
		If w_gakunen > cint(w_rs("M23_GAKUNEN")) then 
			m_gakkamei(w_grp) = w_rs("M02_GAKKAMEI")
			m_Qgakkamei(w_grp) = w_gakkamei
		else 
			m_gakkamei(w_grp) = w_gakkamei
			m_Qgakkamei(w_grp) = w_rs("M02_GAKKAMEI")
		End If 

     '*** �w�ȃO���[�v���������A�w�Ȗ��̔z��Ƀ��R�[�h������B ***
     Elseif w_group <> cint(w_rs("M23_GROUP")) then
			m_gakkamei(w_grp) = w_rs("M02_GAKKAMEI")
			m_Qgakkamei(w_grp) = ""
     End If
     w_group = cint(w_rs("M23_GROUP"))
     w_gakkamei = w_rs("M02_GAKKAMEI")
     w_gakka = cint(w_rs("M02_GAKKA_CD"))
     w_gakunen = cint(w_rs("M23_GAKUNEN"))
	w_rs.MoveNext
   loop
   
     '*** �w�N�����v�̖��O ***
	m_gakkamei(0) = "�w�N�ʁ@���v"
	m_Qgakkamei(0) = ""

    f_GetGakkaMei = true
End Function


Function f_GetGakusei(p_cnt)
'*******************************************************************************
' �@�@�@�\�F�w�ȕʂ̊w���ꗗ���擾���W�v
' �ԁ@�@�l�FTRUE:OK / FALSE:NG
' ���@�@���Fp_Gakka - �w�ȃR�[�h
' �@�\�ڍׁF�w�Ȃɏ�������w���̈ꗗ���擾
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
    Dim w_sSQL
    Dim w_iRet
    Dim w_iCnt
    Dim w_Rs
    Dim w_iGak,w_iCls,w_iSeb,w_iNyu,w_iZai
	
    f_GetGakusei = False
    
    p_sGrp = m_sGakkaGrp(p_cnt)
	
    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "T13_GAKUNEN,SUBSTR(T13_CLASS,1,1) AS T13_CLASS, T13_ZAISEKI_KBN,T11_SEIBETU,T11_NYUGAKU_KBN,T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & "FROM T13_GAKU_NEN,T11_GAKUSEKI "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "T13_NENDO = "& m_iNENDO &" AND "
    w_sSQL = w_sSQL & "T13_GAKUSEI_NO = T11_GAkUSEI_NO AND "
'    w_sSQL = w_sSQL & "T11_NYUNENDO = ("&m_iNENDO&" - T13_GAKUNEN + 1) AND "
    w_sSQL = w_sSQL & "T13_ZAISEKI_KBN <= " & C_ZAI_TEIGAKU & " AND "
    w_sSQL = w_sSQL & "T13_GAKKA_CD IN "
    w_sSQL = w_sSQL & "    (select M23_GAKKA_CD from M23_GAKKA_GRP where M23_GROUP ='"& p_sGrp &"' and M23_NENDO = "& m_iNENDO &") "
    w_sSQL = w_sSQL & ""
    

    If gf_GetRecordset_OpenStatic(w_rs, w_sSQL) <> 0 Then
        '== �擾����Ȃ������ꍇ ==
        Exit Function
    End If
	
    w_rs.MoveFirst

    
    Do Until w_rs.EOF
	    '// �ϐ��ɑ���Bnull�̎��́A�O��������B
	    w_iGak = cint(gf_SetNull2Zero(w_rs("T13_GAKUNEN")))
	    w_iCls = cint(gf_SetNull2Zero(w_rs("T13_CLASS")))
	    w_iSeb = cint(gf_SetNull2Zero(w_rs("T11_SEIBETU")))
	    w_iNyu = cint(gf_SetNull2Zero(w_rs("T11_NYUGAKU_KBN")))
	    w_iZai = cint(gf_SetNull2Zero(w_rs("T13_ZAISEKI_KBN")))
		
		
'		if w_iGak >0 and w_iGak <= 6 and w_iCls > 0 and w_iCls <= m_iGrp_su then 
		if w_iGak >0 and w_iGak <= 6 and w_iCls > 0  then 
			'�w�N�̑S�̐��ɉ��Z
			m_Fld_M(w_iGak,p_cnt) = m_Fld_M(w_iGak,p_cnt) + 1
			
			'�w�N�i���q�j�̑S�̐��ɉ��Z
			If w_iSeb = C_SEIBETU_F then 
				m_Fld_F(w_iGak,p_cnt) = m_Fld_F(w_iGak,p_cnt) + 1
			End If 

			'���w���̑S�̐��ɉ��Z�@�R�N���ȏ�
	'		If w_iNyu = C_NYU_RYUGAKU and w_iGak > 2 Then 
	'			m_Fld_R(w_iGak,p_cnt) = m_Fld_R(w_iGak,p_cnt) + 1
	'		End If

			'�x�w���̑S�̐��ɉ��Z
			If w_iZai = C_ZAI_KYUGAKU Then 
				m_Fld_MK(w_iGak,p_cnt) = m_Fld_MK(w_iGak,p_cnt) + 1

				'�x�w���i���q�j�ɉ��Z
			    If w_iSeb = C_SEIBETU_F Then 
					m_Fld_FK(w_iGak,p_cnt) = m_Fld_FK(w_iGak,p_cnt) + 1
			    End If 
				
				'���w���i�x�w�j�ɉ��Z�@�R�N���ȏ�
	'		    If w_iNyu = C_NYU_RYUGAKU and w_iGak > 2 Then 
	'				m_Fld_RK(w_iGak,p_cnt) = m_Fld_RK(w_iGak,p_cnt) + 1
	'		    End If
			End If 
			
			'�Ώۊw�N�������N���X�̏ꍇ�A�N���X�ʂɏW�v�����B
			If cint(m_ikongoCls(w_iGak)) = C_CLASS_KONGO then
					m_Cls_M(w_iGak,w_iCls) = m_Cls_M(w_iGak,w_iCls) + 1
					'�����̏W�v
					If w_iSeb = C_SEIBETU_F then 
						m_Cls_F(w_iGak,w_iCls) = m_Cls_F(w_iGak,w_iCls) + 1
					End If
			End If 
		end if

		w_rs.MoveNext

    Loop

	'//ں��޾��CLOSE
	Call gf_closeObject(w_rs)

'���w���̏ꍇ�B

    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "T13_GAKUNEN,SUBSTR(T13_CLASS,1,1) AS T13_CLASS, T13_ZAISEKI_KBN,T11_SEIBETU,T11_NYUGAKU_KBN,T13_GAKUSEI_NO "
    w_sSQL = w_sSQL & "FROM T13_GAKU_NEN,T11_GAKUSEKI "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "T13_NENDO = "& m_iNENDO &" AND "
    w_sSQL = w_sSQL & "T13_GAKUSEI_NO = T11_GAkUSEI_NO AND "
'    w_sSQL = w_sSQL & "T11_NYUNENDO = ("&m_iNENDO&" - T13_GAKUNEN + 1) AND "
    w_sSQL = w_sSQL & "T13_ZAISEKI_KBN <= " & C_ZAI_TEIGAKU & " AND "
    w_sSQL = w_sSQL & "T11_NYUGAKU_KBN = " & C_NYU_RYUGAKU & " AND "
    w_sSQL = w_sSQL & "T13_GAKKA_CD IN "
    w_sSQL = w_sSQL & "    (select M23_GAKKA_CD from M23_GAKKA_GRP where M23_GROUP ='"& p_sGrp &"' and M23_NENDO = "& m_iNENDO &") "
    w_sSQL = w_sSQL & ""
    '== ں��޾�Ď擾 ==

'    response.write w_sSQL&"<BR>"
'    response.end


    w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
    If w_iRet <> 0 Then
        '== �擾����Ȃ������ꍇ ==
        Exit Function
    End If
'    w_iGak = cint(w_rs("T13_GAKUNEN"))
if w_rs.EOF = false then 
    w_rs.MoveFirst
    Do Until w_rs.EOF
    '// �ϐ��ɑ���Bnull�̎��́A�O��������B
    w_iGak = cint(gf_SetNull2Zero(w_rs("T13_GAKUNEN")))
    w_iCls = cint(gf_SetNull2Zero(w_rs("T13_CLASS")))
    w_iSeb = cint(gf_SetNull2Zero(w_rs("T11_SEIBETU")))
    w_iNyu = cint(gf_SetNull2Zero(w_rs("T11_NYUGAKU_KBN")))
    w_iZai = cint(gf_SetNull2Zero(w_rs("T13_ZAISEKI_KBN")))

'	if w_iGak >2 and w_iGak <= 6 and w_iCls > 0 and w_iCls <= m_iGrp_su then 
	if w_iGak >2 and w_iGak <= 6 and w_iCls > 0  then 
		'���w���̑S�̐��ɉ��Z�@�R�N���ȏ�
			m_Fld_R(w_iGak,p_cnt) = m_Fld_R(w_iGak,p_cnt) + 1

		'���w���i�x�w�j�ɉ��Z�@�R�N���ȏ�
		    If w_iZai = C_ZAI_KYUGAKU Then 
				m_Fld_RK(w_iGak,p_cnt) = m_Fld_RK(w_iGak,p_cnt) + 1
		    End If

	end if
	  w_rs.MoveNext
    Loop
end if
	'//ں��޾��CLOSE
	Call gf_closeObject(w_rs)

    f_GetGakusei = True
    Exit Function
    
End Function

Sub s_writeSum(p_grp)
'*******************************************************************************
' �@�@�@�\�F�W�v�l���e�[�u���ɏ����o���B
' �ԁ@�@�l�F
' ���@�@���F
' �@�\�ڍׁF
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
dim i,w_sCell
dim w_mFld_M,w_mFld_F,w_mFld_R
w_mFld_M = ""
w_mFld_F = ""
w_mFld_R = ""
w_sCell = "CELL2"
%>
<tr>
<td class="<%=w_sCell%>"><%=m_gakkamei(p_grp)%>
<% if m_Qgakkamei(p_grp) <> "" then %>
<br>(<%=m_Qgakkamei(p_grp)%>)
<% End If %>
</td>

<% for i = 1 to 6  '�w�N���ɃZ���̏�������
call gs_cellPtn(w_sCell)

'//*** �Z���̒l��ϐ��ɑ���B�i�x�w�҂�����ꍇ�́A�J�b�R������������j
'// �w����
if cint(m_Fld_MK(i,p_grp)) > 0 then w_mFld_M = "("&m_Fld_MK(i,p_grp)&")"
w_mFld_M = w_mFld_M & "<br>"&m_Fld_M(i,p_grp)

'// �w�����i���q�j
if m_Fld_FK(i,p_grp) > 0 then w_mFld_F = "("&m_Fld_FK(i,p_grp)&")"
w_mFld_F = w_mFld_F & "<br>"&m_Fld_F(i,p_grp)

'// ���w����
if m_Fld_RK(i,p_grp) > 0 then w_mFld_R = "("&m_Fld_RK(i,p_grp)&")"
w_mFld_R = w_mFld_R  & "<br>"&m_Fld_R(i,p_grp)

%>
<td class="<%=w_sCell%>" align="right" nowrap><%=w_mFld_M%></td>
<td class="<%=w_sCell%>" align="right" nowrap><%=w_mFld_F%></td>
<% If i >= 3 then %>
	<td class="<%=w_sCell%>" align="right" nowrap><%=w_mFld_R%></td>
<% End If %>

<%
w_mFld_M = ""
w_mFld_F = ""
w_mFld_R = ""


next

End Sub

Sub s_writeCell()
'*******************************************************************************
' �@�@�@�\�F�w�Ȗ��A�w�N���̏W�v���o���B
' �ԁ@�@�l�F
' ���@�@���F
' �@�\�ڍׁF
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
dim w_igrp,w_igak
		for w_igrp = 1 to m_iGrp_su
			for w_igak = 1 to 6
				if w_igak < 6 then
					'�w�Ȗ����v
					m_Fld_M(6,w_igrp)   = m_Fld_M(6,w_igrp)   + m_Fld_M(w_igak,w_igrp)
					m_Fld_MK(6,w_igrp) = m_Fld_MK(6,w_igrp) + m_Fld_MK(w_igak,w_igrp)
					m_Fld_F(6,w_igrp)    = m_Fld_F(6,w_igrp)   + m_Fld_F(w_igak,w_igrp)
					m_Fld_FK(6,w_igrp)  = m_Fld_FK(6,w_igrp) + m_Fld_FK(w_igak,w_igrp)
					m_Fld_R(6,w_igrp)   = m_Fld_R(6,w_igrp)    + m_Fld_R(w_igak,w_igrp)
					m_Fld_RK(6,w_igrp) = m_Fld_RK(6,w_igrp)  + m_Fld_RK(w_igak,w_igrp) 
				end if
					'�w�N�����v
					m_Fld_M(w_igak,0)  = m_Fld_M(w_igak,0)   + m_Fld_M(w_igak,w_igrp)
					m_Fld_MK(w_igak,0) = m_Fld_MK(w_igak,0) + m_Fld_MK(w_igak,w_igrp)
					m_Fld_F(w_igak,0)   = m_Fld_F(w_igak,0)    + m_Fld_F(w_igak,w_igrp)
					m_Fld_FK(w_igak,0) = m_Fld_FK(w_igak,0)  + m_Fld_FK(w_igak,w_igrp)
					m_Fld_R(w_igak,0)   = m_Fld_R(w_igak,0)    + m_Fld_R(w_igak,w_igrp)
					m_Fld_RK(w_igak,0) = m_Fld_RK(w_igak,0)  + m_Fld_RK(w_igak,w_igrp) 
			next
			call s_writeSum(w_igrp)
		next
			call s_writeSum(0)
End Sub

Sub s_kongoWrite()
'*******************************************************************************
' �@�@�@�\�F�����N���X�̕\���o���B
' �ԁ@�@�l�F
' ���@�@���F
' �@�\�ڍׁF
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
dim w_sCell
%>
	        <table class="hyo" border="1" width="">
	            <tr>

					
	                <th nowrap class="header">�����N���X</th>
	                <td nowrap class="detail" align="center"><%=gf_fmtWareki(date())%>����</td>

	            </tr>
	        </table>
<BR>
<table border=1 class="hyo" >

<%
for j = 0 to UBound(m_Cls_M,2) 
	If j = 0 Then '�w�b�_�̏�������
%>
		<tr>
		<th class="header">�N���X</th>
		<%
		for i = 1 to UBound(m_Cls_M)
			if cint(m_ikongoCls(i)) = C_CLASS_KONGO then 
			%>
			 <th class="header"><%=i%>�N</th>
			 <th class="header">���q</th>
			<%

			end if 
		next
		%>
		</tr>
		<%
	Else
		iNinzu = 0
		for i = 1 to UBound(m_Cls_M) 		
			 iNinzu = iNinzu + m_Cls_M(i,j)
		next
		
		If iNinzu > 0 Then
			call gs_cellPtn(w_sCell)

%>
			<tr>
			<td class="<%=w_sCell%>"><%=f_GetClassMei(1,j)%></th>
			<% for i = 1 to UBound(m_Cls_M) 
			call gs_cellPtn(w_sCell)

				if cint(m_ikongoCls(i)) = C_CLASS_KONGO then 
				%>
				 <td class="<%=w_sCell%>" align="right"><%=m_Cls_M(i,j)%></td>
				 <td class="<%=w_sCell%>" align="right"><%=m_Cls_F(i,j)%></td>
				<%

				end if 
			next
			%>
			</tr>
		<%
		End If
	End If
next
%>
</table>
<%
End Sub


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
    <title>�w�����ꗗ</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
<body>
<center>
    <%call gs_title("�w�����ꗗ","��@��")%>

<BR>

	        <table class="hyo" border="1" width="">
	            <tr>

					
	                <th nowrap class="header">�w�N�E�w�ȕ�</th>
	                <td nowrap class="detail" align="center"><%=gf_fmtWareki(date())%>����</td>

	            </tr>
	        </table>
<BR>
<table border=1 class="hyo">
<tr>
<th class="header" rowspan="2">�w��</th>
<th class="header" colspan="2"><span style="font-size:12px;">�P�N</span></th>
<th class="header" colspan="2"><span style="font-size:12px;">�Q�N</span></th>
<th class="header" colspan="3"><span style="font-size:12px;">�R�N</span></th>
<th class="header" colspan="3"><span style="font-size:12px;">�S�N</span></th>
<th class="header" colspan="3"><span style="font-size:12px;">�T�N</span></th>
<th class="header" colspan="3"><span style="font-size:12px;">�w�ȕʁ@���v</span></th>
</tr>
<tr>
<th class="header"><span style="font-size:12px;">�v</span></th>
<th class="header"><span style="font-size:12px;">���q</span></th>
<th class="header"><span style="font-size:12px;">�v</span></th>
<th class="header"><span style="font-size:12px;">���q</span></th>
<th class="header"><span style="font-size:12px;">�v</span></th>
<th class="header"><span style="font-size:12px;">���q</span></th>
<th class="header"><span style="font-size:12px;">���w</span></th>
<th class="header"><span style="font-size:12px;">�v</span></th>
<th class="header"><span style="font-size:12px;">���q</span></th>
<th class="header"><span style="font-size:12px;">���w</span></th>
<th class="header"><span style="font-size:12px;">�v</span></th>
<th class="header"><span style="font-size:12px;">���q</span></th>
<th class="header"><span style="font-size:12px;">���w</span></th>
<th class="header"><span style="font-size:12px;">�v</span></th>
<th class="header"><span style="font-size:12px;">���q</span></th>
<th class="header"><span style="font-size:12px;">���w</span></th>
</tr>
<% call s_writecell() %>
</table>
<table width="98%" border="0">
<TR><TD align="right">
<span class="CAUTION" style="text-align:right;">���i�@�j�����́A�x�w�Ґ��œ����ł��B</span><BR>
<span class="CAUTION" style="text-align:right;">���u���w�v�́A���w�����Ӗ����܂��B</span>
</td></tr>
</table>
<% '�����N���X�����݂���ꍇ�́A�����N���X�̕\���o���B%>
<BR>
<% if m_ikongoCls(0) = 1 then %>
<% call s_kongoWrite()
 end If %>
</center>
</body>
</head>
</html>
<%
End Sub
%>