<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������{�Ȗړo�^
' ��۸���ID : skn/skn0130/skn0130_db.asp
' �@      �\: �������{�Ȗڂ̓o�^�E�폜���s�Ȃ�
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H�R�[�h
'           txtSingakuCd        :�i�w�R�[�h
'           txtSinroName        :�i�H���́i�ꕔ�j
'           txtPageSinro        :�\���ϕ\���Ő��i�������g����󂯎������j
'           Sinro_syuseiCD      :�I�����ꂽ�i�H�R�[�h
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           txtSinroKBN     :�i�H�R�[�h�i�߂�Ƃ��j
'           txtSingakuCd        :�i�w�R�[�h�i�߂�Ƃ��j
'           txtSinroName        :�i�H���́i�߂�Ƃ��j
'           txtPageSinro        :�\���ϕ\���Ő��i�߂�Ƃ��j
' ��      ��:
'           ��DB�����̂�
'               �S��ʂ���킽���Ă����ް��̍X�V�E�폜���s�Ȃ�
'-------------------------------------------------------------------------
' ��      ��: 2001/07/24 �{�� ��
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public  m_sDBMode           'DB��Ӱ�ނ̐ݒ�
    Public  m_sMode             '���[�h�̐ݒ�
    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iSikenKbn         ':�����敪
    Public  m_iSikenCode        ':��������
    Public  m_sGakunen          ':�w�N
    Public  m_sClass            ':�׽No
    Public  m_sKamoku           ':�Ȗں���
'   Public  m_sJissiFLG         ':���{�׸�
    Public  m_iMain_FLG         ':���C������
    Public  m_iSeiseki_FLG      ':���ѓ��͋����t���O
    Public  m_iJISSI_FLG        ':���{�׸�
    Public  m_sJissiDate        ':���{���t
    Public  m_sJikan            ':���{����
    Public  m_sKyositu          ':���{����
    Public  m_sKokan1           ':���ѓ��͋����P
    Public  m_sKokan2           ':���ѓ��͋����Q
    Public  m_sKokan3           ':���ѓ��͋����R
    Public  m_sKokan4           ':���ѓ��͋����S
    Public  m_sKokan5           ':���ѓ��͋����T
    Public  m_sPageCD           ':�߰�ރi���o�[

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

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E��}�X�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

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

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '// ���Ұ�SET
        Call s_SetParam()

        '// DB�o�^
'        If m_sMode = "Delete" then
'            If Not f_Delete() then Exit do
'		End If

'        If m_sMode = "Update" then
            If Not f_Update() then Exit do
'		End If

'        If m_sDBMode = "T26" then
'            If Not f_Update() then Exit do
'        End If
'
'        If m_sDBMode = "T27" then
'            If Not f_Insert() then Exit do
'		End If

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
    call gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub


'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    Dim strErrMsg

    strErrMsg = ""

    m_sNendo     = Request("txtNendo")      '�N�x�̎擾
    m_sKyokanCD  = Session("KYOKAN_CD")     ':���[�U�[ID
    m_sPageCD    = Request("txtPageCD")     ': �߰�ރi���o�[
    m_sMode      = Request("txtMode")       'Ӱ�ނ̐ݒ�

    m_sDBMode    = Request("txtDBMode")         'DB��Ӱ�ނ̐ݒ�
    m_iKyokanCd = Session("KYOKAN_CD")         ':�����R�[�h
    m_iSyoriNen     = Session("NENDO")         ':�����N�x
    m_iSikenKbn  = Request("txtSikenKbn")           ':�����敪
    m_iSikenCode     = Request("txtSikenCode")      ':��������
    m_sGakunen   = Request("txtGakunen")            ':�w�N
    m_sClass     = Request("txtClass")          ':�׽No
    m_sKamoku    = Request("txtKamoku")         ':�Ȗں���
    m_iMain_FLG    = Request("txtMainF")         ':���C������
    m_iSeiseki_FLG    = Request("txtSeisekiF")   ':���ѓ��͋����t���O
    m_sJissiFLG      = Request("txtJissiFLG")       ':���{�׸�
    m_iJISSI_FLG = gf_SetNull2Zero(Request("chk1")) ':���{�׸�
    m_sJikan     = Request("txtJikan")          ':���{����
    m_sKyositu   = Request("txtKyositu")            ':���{����
    m_sKokan1    = Request("SKyokanCd1")            ':���ѓ��͋����P
    m_sKokan2    = Request("SKyokanCd2")            ':���ѓ��͋����Q
    m_sKokan3    = Request("SKyokanCd3")            ':���ѓ��͋����R
    m_sKokan4    = Request("SKyokanCd4")            ':���ѓ��͋����S
    m_sKokan5    = Request("SKyokanCd5")            ':���ѓ��͋����T

    If strErrmsg <> "" Then
        ' �G���[��\������t�@���N�V����
        Call err_page(strErrMsg)
        response.end
    End If
'   call s_viewForm(request.form)   '�f�o�b�O�p�@�����̓��e������
End Sub


'********************************************************************************
'*  [�@�\]  �X�V����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
function f_Update()
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    f_Update = False

	'//���{���Ȃ��ꍇ�A���ԥ������NULL������
	if gf_IsNull(m_sJikan) then m_sJikan = "Null"
	if gf_IsNull(m_sKyositu) then m_sKyositu = "Null"
	if m_sKyositu = C_CBO_NULL then m_sKyositu = "Null"

	w_Cls = split(m_sClass,"#")

'Response.Write "UPD<br>" & w_Cls(0) & "<br><br>"
'Response.end

	i = 0
	For i = 0 to UBound(w_Cls) 

		'�������Ԋ��ɊY���f�[�^������ꍇ�́A���Ԋ��f�[�^�̍X�V�A�����ꍇ�͐V�K�ǉ�
		If f_GetSJikanKensu(w_Cls(i)) > 0 Then
		
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " Update T26_SIKEN_JIKANWARI SET "
	
			If m_iMain_FLG = "1" And m_sJissiFLG <> 1 then '//���͊��Ԃ��݂�B
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_JISSI_FLG = " & m_iJISSI_FLG &","
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKEN_JIKAN = " & m_sJikan & ","
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_KYOSITU = " & m_sKyositu & ","
			End If 

			If m_iSeiseki_FLG = "1" then
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN1 = '" & m_sKokan1 & "',"
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN2 = '" & m_sKokan2 & "',"
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN3 = '" & m_sKokan3 & "',"
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN4 = '" & m_sKokan4 & "',"
			    w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN5 = '" & m_sKokan5 & "', "
			End If

			w_sSQL = w_sSQL & vbCrLf & " 	T26_UPD_DATE = TO_CHAR(SYSDATE,'YYYY/MM/DD'), "
			w_sSQL = w_sSQL & vbCrLf & " 	T26_UPD_USER = '" & Session("LOGIN_ID") & "' "
			w_sSQL = w_sSQL & vbCrLf & " WHERE  T26_NENDO = " & m_iSyoriNen
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_SIKEN_KBN = " & m_iSikenKbn
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_SIKEN_CD = '" & m_iSikenCode & "'"
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_GAKUNEN = " & m_sGakunen
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_CLASS = " & w_Cls(i)
			w_sSQL = w_sSQL & vbCrLf & " 	and T26_KAMOKU = '" & m_sKamoku & "'"
'Response.Write "UPD<br>" & w_sSQL & "<br><br>"
'esponse.end
		'�V�K�ǉ�
		Else
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " INSERT INTO T26_SIKEN_JIKANWARI ("
			w_sSQL = w_sSQL & vbCrLf & " 	T26_NENDO,"				'�N�x
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKEN_KBN,"         '�����敪
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKEN_CD,"          '�����R�[�h
			w_sSQL = w_sSQL & vbCrLf & " 	T26_GAKUNEN,"           '�w�N
			w_sSQL = w_sSQL & vbCrLf & " 	T26_CLASS,"             '�N���X�m�n
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KAMOKU,"            '�ȖڃR�[�h
			w_sSQL = w_sSQL & vbCrLf & " 	T26_JISSI_KYOKAN,"      '���{�����R�[�h/���ѓ��͋���
			w_sSQL = w_sSQL & vbCrLf & " 	T26_JISSI_FLG,"         '���{�t���O
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKENBI,"           '���{���t
			w_sSQL = w_sSQL & vbCrLf & " 	T26_MAIN_FLG,"          '���C�������t���O 
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_INP_FLG,"   '���ѓ��͋����t���O 
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN1,"   '���ѓ��͋����R�[�h1
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN2,"   '���ѓ��͋����R�[�h2
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN3,"   '���ѓ��͋����R�[�h3
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN4,"   '���ѓ��͋����R�[�h4
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SEISEKI_KYOKAN5,"   '���ѓ��͋����R�[�h5
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KANTOKU_KYOKAN,"    '�ē����R�[�h
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KYOSITU,"           '���{�����R�[�h
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SIKEN_JIKAN,"       '��������
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KAISI_JIKOKU,"      '�J�n����
			w_sSQL = w_sSQL & vbCrLf & " 	T26_SYURYO_JIKOKU,"     '�I������
			w_sSQL = w_sSQL & vbCrLf & " 	T26_KYOKAN_RENMEI,"     '�����A��
			w_sSQL = w_sSQL & vbCrLf & " 	T26_INS_DATE,"          '�o�^����
			w_sSQL = w_sSQL & vbCrLf & " 	T26_INS_USER,"          '�o�^��
			w_sSQL = w_sSQL & vbCrLf & " 	T26_UPD_DATE,"          '�X�V����
			w_sSQL = w_sSQL & vbCrLf & " 	T26_UPD_USER "          '�X�V��
			w_sSQL = w_sSQL & vbCrLf & ") VALUES ("
			w_sSQL = w_sSQL & vbCrLf & " " & m_iSyoriNen & ","		'�N�x
			w_sSQL = w_sSQL & vbCrLf & " " & m_iSikenKbn & ","      '�����敪
			w_sSQL = w_sSQL & vbCrLf & "'" & m_iSikenCode & "',"    '�����R�[�h
			w_sSQL = w_sSQL & vbCrLf & " " & m_sGakunen & ","       '�w�N
			w_sSQL = w_sSQL & vbCrLf & " " & w_Cls(i) & ","         '�N���X�m�n
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKamoku & "',"       '�ȖڃR�[�h
			w_sSQL = w_sSQL & vbCrLf & "'" & m_iKyokanCd & "',"     '���{�����R�[�h/���ѓ��͋���
			w_sSQL = w_sSQL & vbCrLf & " " & m_iJISSI_FLG & ","     '���{�t���O
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '���{���t
			'w_sSQL = w_sSQL & vbCrLf & "'" & m_iMain_FLG & "',"     '���C�������t���O 
			'w_sSQL = w_sSQL & vbCrLf & "'" & m_iSeiseki_FLG & "',"  '���ѓ��͋����t���O 
			w_sSQL = w_sSQL & vbCrLf & "'" & "1" & "',"				'���C�������t���O 
			w_sSQL = w_sSQL & vbCrLf & "'" & "1" & "',"				'���ѓ��͋����t���O 
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan1 & "',"       '���ѓ��͋����R�[�h1
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan2 & "',"       '���ѓ��͋����R�[�h2
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan3 & "',"       '���ѓ��͋����R�[�h3
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan4 & "',"       '���ѓ��͋����R�[�h4
			w_sSQL = w_sSQL & vbCrLf & "'" & m_sKokan5 & "',"       '���ѓ��͋����R�[�h5
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '�ē����R�[�h
			w_sSQL = w_sSQL & vbCrLf & "" & m_sKyositu & ","       '���{�����R�[�h
			If Trim(m_sJikan) = "" Or IsNull(m_sJikan) Then
				m_sJikan = 0
			End If
			w_sSQL = w_sSQL & vbCrLf & " " & m_sJikan & ","         '��������
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '�J�n����
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '�I������
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '�����A��
			w_sSQL = w_sSQL & vbCrLf & " " & "TO_CHAR(SYSDATE,'YYYY/MM/DD')"  & ","    '�o�^����
			w_sSQL = w_sSQL & vbCrLf & "'" & Session("LOGIN_ID") & "',"            	   '�o�^��
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & ","           '�X�V����
			w_sSQL = w_sSQL & vbCrLf & " " & "NULL" & " "           '�X�V��
			w_sSQL = w_sSQL & vbCrLf & ")"
'Response.Write "INS<br>" & w_sSQL & "<br><br>"
		End If		    
'Response.End

		w_iRet = gf_ExecuteSQL(w_sSQL)
		If w_iRet <> 0 Then
		    'ں��޾�Ă̎擾���s
		    m_bErrFlg = True
		    Exit Function
		End If
	Next

'response.end
    f_Update = True

End Function

'********************************************************************************
'*  [�@�\]  �������Ԋ��f�[�^�̊Y�����R�[�h�����擾����
'*  [����]  
'*  [�ߒl]  f_GetSJikanKensu�F���R�[�h��
'*  [����]  
'********************************************************************************
Function f_GetSJikanKensu(p_sClass)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetSJikanKensu = 0

	Do

		'//�N���X���̎擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  COUNT(*) as JikanData"
		w_sSql = w_sSql & vbCrLf & " FROM T26_SIKEN_JIKANWARI"
		w_sSql = w_sSql & vbCrLf & " WHERE  T26_NENDO = " & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & " and T26_SIKEN_KBN = " & m_iSikenKbn
		w_sSql = w_sSql & vbCrLf & " and T26_SIKEN_CD = '" & m_iSikenCode & "'"
		w_sSql = w_sSql & vbCrLf & " and T26_GAKUNEN = " & m_sGakunen
		w_sSql = w_sSql & vbCrLf & " and T26_CLASS = " & p_sClass
		w_sSql = w_sSql & vbCrLf & " and T26_KAMOKU = '" & m_sKamoku & "'"

'response.write w_sSql&vbCrLf&"<BR>ssssssssssssssss" & p_sClass
'response.end
		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)

		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'�������Ԋ��f�[�^�̊Y�����R�[�h�����擾
			f_GetSJikanKensu = cint(rs("JikanData"))
		End If

		Exit Do
	Loop

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �폜����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
function f_Delete
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��

    f_Delete = False

    w_sSQL = w_sSQL & vbCrLf & " delete "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & " T26_SIKEN_JIKANWARI "
    '���o�����̍쐬
    w_sSQL = w_sSQL & vbCrLf & " WHERE T26_NENDO = " & m_iSyoriNen
    w_sSQL = w_sSQL & vbCrLf & " and T26_SIKEN_KBN = " & m_iSikenKbn
    w_sSQL = w_sSQL & vbCrLf & " and T26_SIKEN_CD = '" & m_iSikenCode & "'"
    w_sSQL = w_sSQL & vbCrLf & " and T26_GAKUNEN = " & m_sGakunen
    w_sSQL = w_sSQL & vbCrLf & " and T26_CLASS = " & m_sClass
    w_sSQL = w_sSQL & vbCrLf & " and T26_KAMOKU = '" & m_sKamoku & "'"
        
'       response.write ("<BR>w_sSQL = " & w_sSQL)

    w_iRet = gf_ExecuteSQL(w_sSQL)

    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    End If

    f_Delete = True

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

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function gonext() {
		window.alert("<%= C_TOUROKU_OK_MSG %>");
		//document.frm.action = "./default.asp";
		document.frm.action = "./skn0130_main.asp";
		document.frm.submit();
    }
    //-->
    </SCRIPT>

    </head>

<body bgcolor="#ffffff" onLoad="setTimeout('gonext()',0000)">

<center>

<Form Name="frm" method="post">

<input type="hidden" name="txtMode"     value = "<%=m_sMode%>">
<input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
<input type="hidden" name="txtSikenKbn" value="<%= m_iSikenKbn %>">
<input type="hidden" name="txtSikenCode" value="<%= m_iSikenCode %>">

</From>
</center>

</body>

</html>


<%
    '---------- HTML END   ----------
End Sub

Sub Nyuryokuzumi()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

%>

    <html>
    <head>
    </head>

    <body>

    <center>
    <font size="2">���͂��ꂽ�A����R�[�h�͂��łɎg�p�ς݂ł�<br><br></font>
    <input type="button" onclick="javascript:history.back()" value="�߁@��">
    </center>
    </body>

    </html>


<%
    '---------- HTML END   ----------
End Sub
%>