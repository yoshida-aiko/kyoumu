<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������{�Ȗړo�^
' ��۸���ID : skn/skn0130/skn0130_regist.asp
' �@      �\: �������{�Ȗڂ̓o�^���s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/07/24 �{���@��
' ��      �X: 2001/12/07 ������ ���ѓ��͋������ǂ������f����UI��ς���̂���߂�B
' ��      �X: 2009/11/25 ��c     �׽�P�ʂœo�^����B
' ��      �X: 2019/06/24 ����     �N���X�̃��C���w�Ȃ̎�u������l�����Ȃ��ꍇ�́A�N���X���̗��Ɋw�Ȗ���\������
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n

    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iSikenKbn         ':�����敪
    Public  m_iSikenCode        ':��������
    Public  m_sGakunen          ':�w�N
    Public  m_sClass            ':�׽
    Public  m_sClassTmp         ':�׽��Ɨp
    Public  m_sKamoku           ':�Ȗں���
    Public  m_sKyokan_NAME      ':������
	Public  m_seisekiF

    Public m_sJikanWhere    '���Ԃ̏���
    Public m_sKyosituWhere  '�����R���{�̏���


    Public  m_Rs                ':�\���ް�

    Public m_bErrFlg        '�װ�׸�
    Public m_iNendo         '�N�x
    Public m_sKyokan_CD     '����CD
    Public m_iMax
    Public m_iDsp
    Public m_sPageCD
    Public m_sTitle         ''�V�K�o�^�E�C���̕\���p
    Public m_sDBMode        ''DB�ւ̍X�VӰ��
    Public m_sMode          ''��ʂ̕\����Ӱ��

    Public m_sGetTable    ''��ʂ̕\����Ӱ��

	Public m_sSeisekiDate
	Public m_chekdate

	Public  m_sClassName            '�N���X��(�N���X�̃��C���w�Ȃ̎�u������l�����Ȃ��ꍇ�́A�w�Ȗ�)		'//2019/06/24 Add Fujibayashi

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
    w_sMsgTitle="�������{�Ȗړo�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    m_bErrFlg = False

	m_chekdate = 0

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

        '// �l��ϐ��ɓ����
        Call s_SetParam()

        '// ��������\������
        if f_GetData_Kyokan(m_iKyokanCd,m_sKyokan_NAME) = False then
            exit do
        end if

        '// �ް���\������
        if f_GetData() = False then
            exit do
        end if

        '���ԂɊւ���WHRE���쐬����
        Call f_MakeJikanWhere()
        '���{�����Ɋւ���WHRE���쐬����
        Call f_MakeKyosituWhere()

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
    m_Rs.close
    set m_Rs = nothing
    Call gs_CloseDatabase()

End Sub



'********************************************************************************
'*  [�@�\]  �l��ϐ��ɓ����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub s_SetParam()
Dim w_clsTmp()
'''Session("SYORI_NENDO") = "999"

    On Error Resume Next
    Err.Clear

    m_iNendo     = Session("NENDO")
    m_sTitle = "�C��"
    m_sPageCD    = Request("txtPageCD")
    m_sMode = Request("txtMode")
    m_iKyokanCd = Session("KYOKAN_CD")              ':�����R�[�h
    m_iSyoriNen = Session("NENDO")                  ':�����N�x
    m_iSikenKbn = Request("txtSikenKbn")            ':�����敪
    m_iSikenCode = Request("txtSikenCd")            ':��������
    m_sGakunen = Request("txtGakunen")              ':�w�N
    m_sClass = Request("txtClass")                  ':�׽		'//2009.11.25 skn0130_main.asp �őI�����ꂽ�׽���n�����B
    m_sClassTmp = f_FstSplit(Request("txtClass"),"#")  ':�׽��Ɨp
    m_sKamoku = Request("txtKamoku")                ':�Ȗں���
    m_seisekiF = Request("txtSeisekiFlg")
	m_sGetTable = Request("txtGetTable")			'T26��T27�̂ǂ���̃e�[�u������f�[�^������Ă��邩

	m_sSeisekiDate = Request("txtKikan")			'���͊��ԏI����

	m_sClassName = Request("txtClassName")			'�N���X��(�N���X�̃��C���w�Ȃ̎�u������l�����Ȃ��ꍇ�́A�w�Ȗ�)		'//2019/06/24 Add Fujibayashi
    
End Sub

Function f_FstSplit(p_str,p_chr)
'********************************************************************************
'*  [�@�\]  �ŏ��̕������擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
	dim w_num
	f_FstSplit = p_str

	w_num = InStr(p_str,p_chr)
	If w_num <> 0 then f_FstSplit = left(p_str,w_num-1)

End Function

'********************************************************************************
'*  [�@�\]  �����̖��̂��擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
function f_GetData_Kyokan(p_iKyokanCd,p_sKyokan_NAME)
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    dim w_Rs

    On Error Resume Next
    Err.Clear

    f_GetData_Kyokan = False
    p_sKyokan_NAME = ""

    if trim(cstr(p_iKyokanCd)) = "" then
        exit function
    end if

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " M04.M04_NENDO "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKAN_CD "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_SEI "
    w_sSQL = w_sSQL & vbCrLf & " ,M04.M04_KYOKANMEI_MEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN M04 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    M04_NENDO = " & m_iNendo & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN_CD = '" & p_iKyokanCd & "'"

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    Else
        '�y�[�W���̎擾
        m_iMax = gf_PageCount(w_Rs,m_iDsp)
    End If

    p_sKyokan_NAME = w_Rs("M04_KYOKANMEI_SEI") & "  " & w_Rs("M04_KYOKANMEI_MEI")


    w_Rs.close

    f_GetData_Kyokan = True

end function

'********************************************************************************
'*  [�@�\]  �Ȗڂ̖��̂��擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
function f_GetKamoku(p_KamokuCd)
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    dim w_Rs

    On Error Resume Next
    Err.Clear

    f_GetKamoku = ""

    If Trim(cstr(p_KamokuCd)) = "" then
        Exit Function
    End If

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & "  M03_NENDO "
    w_sSQL = w_sSQL & vbCrLf & " ,M03_KAMOKU_CD "
    w_sSQL = w_sSQL & vbCrLf & " ,M03_KAMOKUMEI "
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    M03_KAMOKU "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    M03_NENDO = " & m_iNendo & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    M03_KAMOKU_CD = '" & p_KamokuCd & "'"

    w_iRet = gf_GetRecordset(w_Rs, w_sSQL)

    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    End If

    f_GetKamoku = gf_HTMLTableSTR(w_Rs("M03_KAMOKUMEI"))

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �X�V���̕\���ް����擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
function f_GetData()
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_Rs

    On Error Resume Next
    Err.Clear

    f_GetData = False

    w_sSQL = w_sSQL & vbCrLf & " SELECT "
    w_sSQL = w_sSQL & vbCrLf & " T26.T26_SIKENBI "          ''���{���t
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_JISSI_FLG "       ''���{�׸�
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SIKEN_JIKAN "         ''��������
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_KYOSITU "         ''���{����
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_MAIN_FLG "     		''���C�������t���O
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_INP_FLG "     ''���ѓ��͋����t���O
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN1 "         ''���ѓ��͋����P
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN2 "     ''���ѓ��͋����Q
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN3 "         ''���ѓ��͋����R
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN4 "     ''���ѓ��͋����S
    w_sSQL = w_sSQL & vbCrLf & " ,T26.T26_SEISEKI_KYOKAN5 "     ''���ѓ��͋����T
    w_sSQL = w_sSQL & vbCrLf & " ,M03.M03_KAMOKUMEI"                ''�Ȗږ�
    w_sSQL = w_sSQL & vbCrLf & " FROM "
    w_sSQL = w_sSQL & vbCrLf & "    T26_SIKEN_JIKANWARI T26 "
    w_sSQL = w_sSQL & vbCrLf & "    ,M03_KAMOKU M03 "
    w_sSQL = w_sSQL & vbCrLf & " WHERE "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_NENDO  = M03.M03_NENDO(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_KAMOKU = M03.M03_KAMOKU_CD(+) AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_NENDO = " & m_iSyoriNen & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_SIKEN_KBN = " & m_iSikenKbn & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_SIKEN_CD = '" & m_iSikenCode & "' AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_GAKUNEN = " & m_sGakunen & " AND "
    'w_sSQL = w_sSQL & vbCrLf & "    T26.T26_CLASS = " & m_sClassTmp & " AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_CLASS = " & m_sClass & " AND "		'//2009.11.25 iwata �׽�P�ʂœo�^����
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_KAMOKU = '" & m_sKamoku & "' AND "
    w_sSQL = w_sSQL & vbCrLf & "    T26.T26_JISSI_KYOKAN = '" & m_iKyokanCd & "'"

    w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
    If w_iRet <> 0 Then
        'ں��޾�Ă̎擾���s
        m_bErrFlg = True
        Exit Function
    Else
        '�y�[�W���̎擾
        m_iMax = gf_PageCount(m_Rs,m_iDsp)
    End If

    f_GetData = True

end function


'********************************************************************************
'*  [�@�\]  ���ԃR���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub f_MakeJikanWhere()

    m_sJikanWhere=""
    m_sJikanWhere = m_sJikanWhere & " M42_NENDO = " & m_iSyoriNen & ""

'response.write m_sJikanWhere

End Sub

'********************************************************************************
'*  [�@�\]  ���{�����R���{�Ɋւ���WHRE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub f_MakekyosituWhere()

    m_sKyosituWhere=""
    m_sKyosituWhere = " M06_NENDO = " & m_iSyoriNen & ""

'response.write m_sKyosituWhere

End Sub

'****************************************************
'[�@�\] �f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'       (���X�g�_�E���{�b�N�X�I��\���p)
'[����] pData1 : �f�[�^�P
'       pData2 : �f�[�^�Q
'[�ߒl] f_Selected : "SELECTED" OR ""
'
'****************************************************
Function f_Selected(pData1,pData2)

    On Error Resume Next
    Err.Clear

    f_Selected = ""

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected"
        Else
        End If
    End If

End Function

'****************************************************
'[�@�\] ׼޵����
'[����] pData1 : �f�[�^�P
'[�ߒl] f_Checked : "SELECTED" OR ""
'
'****************************************************
Function f_Checked(pData,p_Chk1,p_Chk2)

    On Error Resume Next
    Err.Clear

    p_Chk1=""
    p_Chk2=""

    if cstr(pData) = cstr(C_SIKEN_KBN_JISSI) then           ''���{
        p_Chk1=""
        p_Chk2="checked"
    elseif cstr(pData) = cstr(C_SIKEN_KBN_NOT_JISSI) then       ''���{���Ȃ�
        p_Chk1="checked"
        p_Chk2=""
    else
        p_Chk1=""
        p_Chk2=""
    end if

End Function

'********************************************************************************
'*  [�@�\]  �\������(����)���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Function f_GetSikenName()
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetSikenName = ""
    w_sSikenName = ""

    Do
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_NENDO=" & m_iNendo
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_iSikenKbn

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

        If rs.EOF = False Then
            w_sSikenName = rs("M01_SYOBUNRUIMEI")
        End If

        Exit Do
    Loop

	f_GetSikenName = w_sSikenName

    Call gf_closeObject(rs)

End Function

Function f_Nyuryokudate(p_sSikenDate,p_sGakunen)
'********************************************************************************
'*	[�@�\]	�w�N�ʎ������Ԃ��擾����B
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	Add 2001.12.26 ���c
'********************************************************************************
	dim w_date

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1


	w_date = gf_YYYY_MM_DD(date(),"/")
	w_Syuryo = "T24_SIKEN_SYURYO"
	w_kyokan = Session("KYOKAN_CD")

	if w_kyokan = NULL or w_kyokan = "" then w_kyokan = "@@@"

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  MIN(T24_SIKEN_NITTEI.T24_SIKEN_KAISI) as KAISI"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SIKEN_SYURYO) as SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,MAX(T24_SIKEN_NITTEI.T24_SEISEKI_SYURYO) as SEI_SYURYO"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN.M01_SYOBUNRUIMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & "  ,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD = T24_SIKEN_NITTEI.T24_SIKEN_KBN"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_NENDO = T24_SIKEN_NITTEI.T24_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_GAKUNEN =" & Cint(p_sGakunen)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_NENDO=" & Cint(m_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KBN=" & Cint(m_iSikenKbn)
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_KAISI <= '" & w_date & "' "
		'w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_SYURYO >= '" & w_date & "' "
		w_sSQL = w_sSQL & vbCrLf & "  Group By M01_SYOBUNRUIMEI"

'response.write w_sSQL & "<<<BR>"

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do
		End If

		If m_DRs.EOF Then
			Exit Do
		Else

			p_sSikenDate = m_DRs("SYURYO") '//okada 2001.12.25
'response.write " [ " & p_sSikenDate & " ] "
		End If
		f_Nyuryokudate = 0
		Exit Do
	Loop

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Dim w_Kyokan
Dim w_Chk1
Dim w_Chk2

    On Error Resume Next
    Err.Clear

	'// ��׳�ް�ɂ���ĸ׽�w���ς���
	if session("browser") = "IE" then
		w_Class = "class='num'"
	Else
		w_Class = ""
	End if

%>

<html>

<head>

<title>�g�p���ȏ��o�^</title>

    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
         flag = false;

         function Lock() {
            if(frm.chk1[0].checked){
                fm = document.frm;
                flag = !flag;
                fm.date.disabled = true;
                fm.txtTime.disabled = true;
                fm.room.disabled = true;
                }
            }
         function unLock() {
            if(frm.chk1[1].checked){
                fm = document.frm;
                fm.date.disabled = false;
                fm.txtTime.disabled = false;
                fm.room.disabled = false;
                }
            }
        //************************************************************
        //  [�@�\]  ���C���y�[�W�֖߂�
        //  [����]  �Ȃ�
        //  [�ߒl]  �Ȃ�
        //  [����]
        //
        //************************************************************
        function f_Back(){

            //document.frm.action="./default.asp";
            document.frm.action="./skn0130_main.asp";
            document.frm.target="";
            document.frm.submit();

        }

        //************************************************************
        //  [�@�\]  ���������Q�ƑI����ʃE�B���h�E�I�[�v��
        //  [����]
        //  [�ߒl]
        //  [����]
        //************************************************************
        function KyokanWin(p_iInt,p_sKNm) {

			var obj=eval("document.frm."+p_sKNm)

            URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+escape(obj.value)+"";
            //URL = "../../Common/com_select/SEL_KYOKAN/default.asp?txtI="+p_iInt+"&txtKNm="+p_sKNm+"";
            nWin=open(URL,"gakusei","location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no,width=530,height=650,top=0,left=0");
            nWin.focus();
            return true;
        }

	//************************************************************
	//  [�@�\] �N���A�{�^���������ꂽ�Ƃ�
	//  [����]
	//  [�ߒl]
	//  [����]
	//************************************************************
	function jf_Clear(pTextName,pHiddenName){
		eval("document.frm."+pTextName).value = "";
		eval("document.frm."+pHiddenName).value = "";
	}

	//************************************************************
	//  [�@�\] ���{����A���Ȃ��{�^���������ꂽ�Ƃ�
	//  [����]
	//  [�ߒl]
	//  [����]
	//************************************************************
	function jf_Action(pMode){

		if(pMode == "False"){


			// ���Ԃ�NUll�ɂ���
			document.frm.txtJikan.value = "";

			// ������NULL�ɂ���
			document.frm.txtKyositu.options[0].selected = true;

			//���Ԃ�readOnly�ɂ���
			document.frm.txtJikan.readOnly = true;

			//������disabled�ɂ���
			document.frm.txtKyositu.disabled = true;

		}else{
			//���Ԃ�readOnly���͂���
			document.frm.txtJikan.readOnly = false;

			//������disabled�ɂ���
			document.frm.txtKyositu.disabled = false;

		}

	}

	//************************************************************
	//  [�@�\]���͏������ԊO
	//  [����]
	//  [�ߒl]
	//  [����]
	//************************************************************
	function jf_Action2(){

			//���Ԃ�readOnly�ɂ���
			document.frm.txtJikan.readOnly = true;

			//���{���邵�Ȃ�
			//document.frm.chk1.disabled = false;

			//������disabled�ɂ���
			document.frm.txtKyositu.disabled = true;

	}


    //************************************************************
    //  [�@�\]  �g�p���ȏ��o�^
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_touroku(){

		// ���{����Ƃ�
		if(document.frm.chk1[1].checked == true && document.frm.chk1[1].disabled == false){

	        // ������NULL����������
	        if( f_Trim(document.frm.txtJikan.value) == "" ){
	            window.alert("�������Ԃ����͂���Ă��܂���");
	            document.frm.txtJikan.focus();
	            return;
	        }
	        // ���������p�p������������
	        //var str = new String(document.frm.txtJikan.value);
	        var str = document.frm.txtJikan.value;
	        if( isNaN(str) ){
	            window.alert("�������Ԃ����p�p�����ł͂���܂���");
	            document.frm.txtJikan.focus();
	            return ;
	        }

	        if( f_Trim(document.frm.txtJikan.value) < 0 ){
	            window.alert("�������Ԃ����������͂���Ă��܂���");
	            document.frm.txtJikan.focus();
	            return;
	        }

            if (f_chkNumber(f_Trim(document.frm.txtJikan.value))==1){
                alert("�������Ԃ����������͂���Ă��܂���")
	            document.frm.txtJikan.focus();
                return;
			}

	        // ������5���`�F�b�N������
	        var str = new String(document.frm.txtJikan.value);
			if( f_Trim(str) == 0 ){
	                window.alert("�������Ԃ�5���P�ʂœ��͂��Ă�������");
	                document.frm.txtJikan.focus();
	                return ;
			}else{
		        if( str.length < 2 ){
		            str = 0 + str;
		        }

		        if( f_Trim(str).substr(Number(str.length)-1,1) != 0 ){
		            if( f_Trim(str).substr(Number(str.length)-1,1) != 5 ){
		                window.alert("�������Ԃ�5���P�ʂœ��͂��Ă�������");
		                document.frm.txtJikan.focus();
		                return ;
		            }
				}
			}
		}
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		document.frm.chk1[0].disabled = false;
		document.frm.chk1[1].disabled = false;
        document.frm.action="./skn0130_db.asp";
        document.frm.target="";
        document.frm.txtDBMode.value = "Update";
        document.frm.submit();
    }

    //************************************************************
    //  [�@�\]  �����`�F�b�N
    //  [����]  p_num
    //  [�ߒl]  �����F0   ���s�F1
    //  [����]	�������ǂ������`�F�b�N(�}�C�i�X�l�A�����_�L�̏ꍇ�̓G���[��Ԃ�)
    //************************************************************
	function f_chkNumber(p_num){

		//���l�`�F�b�N
		if (isNaN(p_num)){
			return 1;
		}else{

			//�}�C�i�X���`�F�b�N
			var wStr = new String(p_num)
			if (wStr.match("-")!=null){
				return 1;
			};

			//�����_�`�F�b�N
			w_decimal = new Array();
			w_decimal = wStr.split(".")
			if(w_decimal.length>1){
				return 1;
			}

		};
		return 0;
	}


    //************************************************************
    //  [�@�\]  �g�p���ȏ��폜
    //  [����]  p_iPage :�\���Ő�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Delete(){

        document.frm.action="./skn0130_db.asp";
        document.frm.target="";
        document.frm.txtDBMode.value = "Delete";
        document.frm.submit();
    }

    //-->
    </script>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <style>
    .gray {background:gray;}
    .white {background:white;}
    </style>
</head>

<%
	'// ���{�׸�
	if gf_HTMLTableSTR(m_Rs("T26_JISSI_FLG")) = "2" then
		w_onLadFcnc = "onLoad=jf_Action('False')"
	End if

	'// ���͏������ԊO

	w_date = gf_YYYY_MM_DD(date(),"/")
'response.write m_sSeisekiDate & " " &  w_date  & " " & m_sGakunen
	'�w�N�̎������͏������Ԃ��擾�im_sGakunen�j
	Call f_Nyuryokudate(m_sSeisekiDate,m_sGakunen)

	if m_sSeisekiDate < w_date Then
'response.write m_sSeisekiDate & " " &  w_date
		w_onLadFcnc = "onLoad=jf_Action2()"
		m_chekdate = 1
	End if
%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%=w_onLadFcnc%>>
<%call gs_title("�������{�Ȗړo�^",m_sTitle)%>
<form name="frm" method="post">
<center>

<span class=CAUTION>
	�� �u���{����v�̏ꍇ�́A�u���ԁv���K�{���͂ƂȂ�܂�<BR>
	�� �u���{���Ȃ��v�ɂ���ƁA�u���ԁv�Ɓu���{�����v�̑I�����ł��܂���
</span>
<BR>
<br>

<table border="0" cellpadding="1" cellspacing="1" width="400">
    <tr>
        <td align="left">
            <table width="100%" border=1 CLASS="hyo">
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">����</th>
	                <TD CLASS="detail">�@<%=f_GetSikenName%></td>
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">�N���X</th>
<!--	                <TD CLASS="detail">�@<%=m_sGakunen & "�N�@" & gf_GetClassName(m_iNendo,m_sGakunen,m_sClass) %></td>				'2019/06/24 Del Fujibayashi-->
						<TD CLASS="detail">�@<%=m_sGakunen & "�N�@" & m_sClassName%></td>											<!--'2019/06/24 Add Fujibayashi-->
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">�Ȗږ���</th>
	                <TD CLASS="detail">�@<%=f_GetKamoku(m_sKamoku)%>�@</td>
	            </tr>
<!--
				<% if Not gf_IsNull(m_Rs("T26_SIKENBI")) then %>
		            <!tr>
						<TH nowrap CLASS="header" align="center" height="16">���{���t</th>
		                <TD CLASS="detail">�@<%= gf_fmtWareki(m_Rs("T26_SIKENBI")) %></td>
		            </tr>
				<% End if %>
-->
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">���@�@�{</th>
	                <% call f_Checked(gf_HTMLTableSTR(m_Rs("T26_JISSI_FLG")),w_Chk1,w_Chk2) %>
					<%
					if gf_SetNull2Zero(trim(m_Rs("T26_JISSI_FLG"))) = 0 then
						wClass = "JISSHIMI"
						w_Chk2 = "checked"
					Else
						wClass = "detail"
					End if
					%>
	                <TD CLASS="<%=wClass%>">
					<% if m_chekdate = 1 then %>
						<input type="radio" disabled name="chk1" value="2" <%=w_Chk1%> onClick="javascript:jf_Action('False')"><font>���{���Ȃ�
						<input type="radio" disabled name="chk1" value="1" <%=w_Chk2%> onClick="javascript:jf_Action('True')">���{����</font></td>
					<% else %>
						<input type="radio" name="chk1" value="2" <%=w_Chk1%> onClick="javascript:jf_Action('False')"><font>���{���Ȃ�
						<input type="radio" name="chk1" value="1" <%=w_Chk2%> onClick="javascript:jf_Action('True')">���{����</font></td>
					<% end if %>
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">���@�@��</th>
	                <TD CLASS="detail"><input type="text" name="txtJikan" size="4" value="<%=m_RS("T26_SIKEN_JIKAN")%>" <%=w_Class%>>&nbsp;��
					</td>
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16">���{����</th>
	                <TD CLASS="detail">
	    <%          '���ʊ֐�������{�����Ɋւ���R���{�{�b�N�X���o�͂���
	                call gf_ComboSet("txtKyositu",C_CBO_M06_KYOSITU,m_sKyosituWhere,"",True,gf_HTMLTableSTR(m_Rs("T26_KYOSITU")))
	    %>
	                </td>
	            </tr>
	            <tr>
					<TH nowrap CLASS="header" align="center" height="16" valign=top ROWSPAN="5">���ѓ��͋���</th>
	                <TD CLASS="detail">&nbsp;1�F<%=m_sKyokan_NAME%><br>
	                <input type=hidden name="SKyokanCd1" VALUE='<%=m_iKyokanCd%>'>
	                </td>
	            </tr>
	            <tr>
	                <TD CLASS="detail" nowrap>
	                <% call f_GetData_Kyokan(gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN2")),w_Kyokan)%>
	                &nbsp;2�F<input type=text class=text name="SKyokanNm2" VALUE='<%=w_Kyokan%>' readonly>
	                <input type=hidden name="SKyokanCd2" VALUE='<%=gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN2"))%>'>
	                <!--<input type=button class=button value="�I��" onclick="KyokanWin(2,'<%=w_Kyokan%>')">-->

	                <input type=button class=button value="�I��" onclick="KyokanWin(2,'SKyokanNm2')">
					<input type=button class=button value="�N���A" onclick="jf_Clear('SKyokanNm2','SKyokanCd2')">
	                </td>
	            </tr>
	            <tr>
	                <TD CLASS="detail" nowrap>
	                <% call f_GetData_Kyokan(gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN3")),w_Kyokan)%>
	                &nbsp;3�F<input type=text class=text name="SKyokanNm3" VALUE='<%=w_Kyokan%>' readonly>
	                <input type=hidden name="SKyokanCd3" VALUE='<%=gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN3"))%>'>
	                <!--<input type=button class=button value="�I��" onclick="KyokanWin(3,'<%=w_Kyokan%>')">-->
	                <input type=button class=button value="�I��" onclick="KyokanWin(3,'SKyokanNm3')">
					<input type=button class=button value="�N���A" onclick="jf_Clear('SKyokanNm3','SKyokanCd3')">
	                </td>
	            </tr>
	            <tr>
	                <TD CLASS="detail" nowrap>
	                <% call f_GetData_Kyokan(gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN4")),w_Kyokan)%>
	                &nbsp;4�F<input type=text class=text name="SKyokanNm4" VALUE='<%=w_Kyokan%>' readonly>
	                <input type=hidden name="SKyokanCd4" VALUE='<%=gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN4"))%>'>
	                <!--<input type=button class=button value="�I��" onclick="KyokanWin(4,'<%=w_Kyokan%>')">-->
	                <input type=button class=button value="�I��" onclick="KyokanWin(4,'SKyokanNm4')">
					<input type=button class=button value="�N���A" onclick="jf_Clear('SKyokanNm4','SKyokanCd4')">
	                </td>
	            </tr>
	            <tr>
	                <TD CLASS="detail" nowrap>
	                <% call f_GetData_Kyokan(gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN5")),w_Kyokan)%>
	                &nbsp;5�F<input type=text class=text name="SKyokanNm5" VALUE='<%=w_Kyokan%>' readonly>
	                <input type=hidden name="SKyokanCd5" VALUE='<%=gf_HTMLTableSTR(m_Rs("T26_SEISEKI_KYOKAN5"))%>'>
	                <!--<input type=button class=button value="�I��" onclick="KyokanWin(5,'<%=w_Kyokan%>')">-->
	                <input type=button class=button value="�I��" onclick="KyokanWin(5,'SKyokanNm5')">
					<input type=button class=button value="�N���A" onclick="jf_Clear('SKyokanNm5','SKyokanCd5')">
	                </td>
	            </tr>
            </TABLE>
        </td>
    </TR>
</TABLE>
<table width=40%>
    <tr>
        <td width="50%" align="left">
            <input type="button" class=button value="�@�o�@�^�@" OnClick="f_touroku()">
        </td>
        <td width="50%" align="right">
            <input type="Button" class=button value="�L�����Z��" OnClick="f_Back()">
        </td>
    </tr>
</table>

    <input type="hidden" name="txtDBMode" value = "<%=m_sGetTable%>">
    <input type="hidden" name="txtMode"   value = "<%=m_sMode%>">
    <input type="hidden" name="txtPageCD" value="<%= m_sPageCD %>">
    <input type="hidden" name="txtTitle"  value="<%= m_sTitle %>">

    <input type="hidden" name="txtSikenKbn"  value="<%= m_iSikenKbn %>">
    <input type="hidden" name="txtSikenCode" value="<%= m_iSikenCode %>">
    <input type="hidden" name="txtGakunen"   value="<%= m_sGakunen %>">
    <input type="hidden" name="txtClass"     value="<%= m_sClass %>">
    <input type="hidden" name="txtKamoku"    value="<%= m_sKamoku %>">
    <input type="hidden" name="txtMainF"    value="<%= m_Rs("T26_MAIN_FLG") %>">
    <input type="hidden" name="txtSeisekiF"    value="<%= m_Rs("T26_SEISEKI_INP_FLG") %>">
    <input type="hidden" name="txtJissiFLG"  value=<%=m_chekdate%>>
    <input type="hidden" name="txtClassName"     value="<%= m_sClassName %>">		<% '2019/06/24 Add Fujibayashi%>

</center>

</form>
</body>

</html>

<%
End Sub
%>

