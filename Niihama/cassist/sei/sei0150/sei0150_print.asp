<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0150/sei0150_print.asp
' �@      �\: ���ѓo�^�̈���p�E�B���h�E
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h   ��    SESSION���i�ۗ��j
'           :�N�x     ��    SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h   ��    SESSION���i�ۗ��j
'           :�N�x     ��    SESSION���i�ۗ��j
' ��      ��:
' (�p�^�[��)
' �E�ʏ���ƁA���ʊ���
' �E���l���́A��������(����)
' �E�]���s�\����(�F�{�d�g�̂�)
' �E�Ȗڋ敪(0:��ʉȖ�,1:���Ȗ�)
' �E�K�C�I���敪(1:�K�C,2:�I��)
' �E���x���ʋ敪(0:��ʉȖ�,1:���x���ʉȖ�)�𒲂ׂ�
'-------------------------------------------------------------------------
' ��      ��: 2011/07/13 �c��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
  '�G���[�n
    Dim m_bErrFlg       '//�װ�׸�

    Const C_ERR_GETDATA = "�f�[�^�̎擾�Ɏ��s���܂���"

    '�����I��p��Where����
    Dim m_iNendo        '//�N�x
    Dim m_sKyokanCd       '//�����R�[�h
    Dim m_sSikenKBN       '//�����敪
    Dim m_sGakuNo       '//�w�N
    Dim m_sClassNo        '//�w��
    Dim m_sKamokuCd       '//�ȖڃR�[�h
    Dim m_sSikenNm        '//������
    Dim m_rCnt          '//���R�[�h�J�E���g
    Dim m_sGakkaCd
    Dim m_iSyubetu        '//�o���l�W�v���@

    Dim m_iNKaishi
    Dim m_iNSyuryo
    Dim m_iKekkaKaishi
    Dim m_iKekkaSyuryo

    Dim m_iIdouEnd        '//�ٓ��Ώۊ��ԏI����

    Dim m_iKamoku_Kbn
    Dim m_iHissen_Kbn
  Dim m_ilevelFlg
  Dim m_Rs
  Dim m_SRs

    Dim m_iKongou

  Dim m_iSouJyugyou     '//�����Ǝ���
  DIm m_iJunJyugyou     '//�����Ǝ���

  Dim m_bKekkaNyuryokuFlg   '//���ۓ��͉\�׸�(True:���͉� / False:���͕s��)

  Dim m_iShikenInsertType

  Dim m_sSyubetu

  '2002/06/21
  Dim m_iKamokuKbn        '//�Ȗڋ敪(0:�ʏ���ƁA1:���ʉȖ�)
  Dim m_sKamokuBunrui       '//�Ȗڕ���(01:�ʏ���ƁA02:�F��ȖځA03:���ʉȖ�)

  Dim m_iSeisekiInpType
  Dim m_Date
  Dim m_bZenkiOnly
  Dim m_KekkaGaiDispFlg

  Dim m_MiHyokaFlg

  Dim m_bNiteiFlg

  Dim m_sGakkoNO       '�w�Z�ԍ�

  Dim m_lHaitoTani

'///////////////////////////���C������/////////////////////////////
  'Ҳ�ٰ�ݎ��s
  Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub Main()
  Dim w_iRet
  Dim w_sSQL
  Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

  'Message�p�̕ϐ��̏�����
  w_sWinTitle = "�L�����p�X�A�V�X�g"
  w_sMsgTitle = "���ѓo�^"
  w_sMsg = ""
  w_sRetURL = C_RetURL & C_ERR_RETURL
  w_sTarget = ""

  On Error Resume Next
  Err.Clear

  m_bErrFlg = false
  m_MiHyokaFlg = false

  Do
    '//�ް��ް��ڑ�
    If gf_OpenDatabase() <> 0 Then
      m_bErrFlg = True
      Exit Do
    End If

    '//�s���A�N�Z�X�`�F�b�N
    Call gf_userChk(session("PRJ_No"))

    '//���Ұ�SET
    Call s_SetParam()


    '2002.12.25 Ins
    '�w�Z�ԍ��̎擾
    if Not gf_GetGakkoNO(m_sGakkoNO) then Exit Do

'Response.Write "[1]"


    '//���ۊO��\�����邩�`�F�b�N
    if not gf_ChkDisp(C_KEKKAGAI_DISP,m_KekkaGaiDispFlg) then
      m_bErrFlg = True
      Exit Do
    End If
'Response.Write "[2]"

    '//���ѓ��͕��@�̎擾(0:�_��[C_SEISEKI_INP_TYPE_NUM]�A1:����[C_SEISEKI_INP_TYPE_STRING]�A2:���ہA�x��[C_SEISEKI_INP_TYPE_KEKKA])
    if not gf_GetKamokuSeisekiInp(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iSeisekiInpType) then
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[3]"

    '//�O���̂݊J�݂��ʔN�����ׂ�
    if not f_SikenInfo(m_bZenkiOnly) then
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[4]"

    '//���сA���ۓ��͊��ԃ`�F�b�N
    If not f_Nyuryokudate() Then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[5]"

    '//�o�����ۂ̎������擾
    '//�Ȗڋ敪(0:������,1:�ݐ�)
    If gf_GetKanriInfo(m_iNendo,m_iSyubetu) <> 0 Then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[6]"

    '//�F��O����擾
'   if not gf_GetNintei(m_iNendo,m_bNiteiFlg) then
    if not gf_GetGakunenNintei(m_iNendo,cint(m_sGakuNo),m_bNiteiFlg) then '2003.04.11 hirota
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[7]"

    If m_iKamokuKbn = C_JIK_JUGYO then  '�ʏ���Ƃ̏ꍇ
      '//�Ȗڏ����擾
      '//�Ȗڋ敪(0:��ʉȖ�,1:���Ȗ�)�A�y�сA�K�C�I���敪(1:�K�C,2:�I��)�𒲂ׂ�
      '//���x���ʋ敪(0:��ʉȖ�,1:���x���ʉȖ�)�𒲂ׂ�
      If not f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) Then m_bErrFlg = True : Exit Do
    end if

'Response.Write "[8]"
    If not f_GetKongoClass(cint(m_sGakuNo),cint(m_sClassNo),m_iKongou) Then m_bErrFlg = True : Exit Do

    '//���сA�w���f�[�^�擾
    If not f_GetStudent() Then m_bErrFlg = True : Exit Do

    If m_Rs.EOF Then
      Call gs_showWhitePage("�l���C�f�[�^�����݂��܂���B","���ѓo�^")
      Exit Do
    End If

'response.Write "[9]"
'response.end

    '//���ې��̎擾
    if not gf_GetSyukketuData2(m_SRs,m_sSikenKBN,m_sGakuNo,m_sClassNo,m_sKamokuCd,m_iNendo,m_iShikenInsertType,m_sSyubetu) then
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[13]"
'response.end
    '// �y�[�W��\��
    Call showPage()
    Exit Do
  Loop

  '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
  If m_bErrFlg = True Then
    w_sMsg = gf_GetErrMsg()

    if w_sMsg = "" then w_sMsg = C_ERR_GETDATA

'    Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
  End If

  '// �I������
  Call gf_closeObject(m_Rs)
  Call gf_closeObject(m_SRs)

  Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()

  Dim wArray

  wArray = split(request("arr"),",")
  m_iNendo   = wArray(0)
  m_sKyokanCd  = wArray(1)
  m_sGakuNo  = Cint(wArray(2))
  m_sClassNo   = Cint(wArray(3))
  m_sSikenKBN  = Cint(wArray(4))
  m_sGakkaCd   = wArray(5)
  m_sKamokuCd  = wArray(6)
  m_sSyubetu   = trim(wArray(7))
  m_iKamokuKbn = Cint(wArray(8))
  m_iSeisekiInpType = wArray(9)

  m_iShikenInsertType = 0

  if m_iKamokuKbn = Cint(C_JIK_JUGYO) then
    '�ʏ�Ȗ�
    m_sKamokuBunrui = C_KAMOKUBUNRUI_TUJYO
  else
    '���ʉȖ�
    m_sKamokuBunrui = C_KAMOKUBUNRUI_TOKUBETU
  end if

  m_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")

End Sub

'********************************************************************************
'*	[�@�\]	�������擾
'********************************************************************************
Function f_ShikenMei()
	Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_ShikenMei = ""

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD = " & cint(m_sSikenKBN)
	w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
	w_sSQL = w_sSQL & " AND M01_NENDO = " & cint(m_iNendo)

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

	If not w_Rs.EOF Then
		f_ShikenMei = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))
	End If

End Function

'********************************************************************************
'*  [�@�\]  ���C�e�[�u�����Ȗږ��̂��擾
'********************************************************************************
Function f_GetKamokuName(p_Gakunen,p_GakkaCd,p_KamokuCd)
	Dim w_sSQL
	Dim w_Rs
	Dim w_GakkaCd

	On Error Resume Next
	Err.Clear
                       
	f_GetKamokuName = ""

	w_sSQL = ""

	If Cstr(m_iKamokuKbn) = Cstr(C_TUKU_FLG_TUJO) Then '�ʏ���ƂƓ��ʊ����Ŏ����ς���B
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	T15_KAMOKUMEI AS KAMOKUMEI"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	T15_RISYU"
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 	T15_NYUNENDO=" & cint(m_iNendo) - cint(p_Gakunen) + 1
		w_sSQL = w_sSQL & " AND T15_GAKKA_CD='" & p_GakkaCd & "'"
		w_sSQL = w_sSQL & " AND T15_KAMOKU_CD='" & p_KamokuCd & "'"
	Else
		w_sSQL = w_sSQL & " SELECT "
		w_sSQL = w_sSQL & " 	M41_MEISYO AS KAMOKUMEI"
		w_sSQL = w_sSQL & " FROM "
		w_sSQL = w_sSQL & " 	M41_TOKUKATU"
		w_sSQL = w_sSQL & " WHERE "
		w_sSQL = w_sSQL & " 	M41_NENDO=" & cint(m_iNendo)
		w_sSQL = w_sSQL & " AND M41_TOKUKATU_CD='" & p_KamokuCd & "'"
	End If

	if gf_GetRecordset(w_Rs, w_sSQL) <> 0 then exit function

	If not w_Rs.EOF Then f_GetKamokuName = w_Rs("KAMOKUMEI")

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �O���J�݂��ǂ������ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Function f_SikenInfo(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs

    On Error Resume Next
    Err.Clear

    f_SikenInfo = false
  p_bZenkiOnly = false

  '//�����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   T15_KAMOKU_CD, "
  w_sSQL = w_sSQL & "   T15_HAITO" & m_sGakuNo
  w_sSQL = w_sSQL & "   ,T15_KAISETU" & m_sGakuNo
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T15_RISYU "
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   T15_NYUNENDO = " & Cint(m_iNendo)-cint(m_sGakuNo)+1
  w_sSQL = w_sSQL & " AND T15_GAKKA_CD = '" & m_sGakkaCd & "'"
  w_sSQL = w_sSQL & " AND T15_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'"
  'w_sSQL = w_sSQL & " AND T15_KAISETU" & m_sGakuNo & "=" & C_KAI_ZENKI

  if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

  'Response.Write w_ssql & "<BR>"
  'response.end

  '//�߂�l���
  If w_Rs.EOF = False Then

	if cint(w_Rs("T15_KAISETU" & m_sGakuNo)) = C_KAI_ZENKI then
  'Response.Write "C_KAI_ZENKI = " & C_KAI_ZENKI & "<BR>"
  'response.end
	    p_bZenkiOnly = True
	end if

	'�z���P�ʂ̎擾
	m_lHaitoTani = w_Rs("T15_HAITO" & m_sGakuNo)

  else
	Call f_SikenInfo_T16(p_bZenkiOnly)
  End If

  f_SikenInfo = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �O���J�݂��ǂ������ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Function f_SikenInfo_T16(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs

    On Error Resume Next
    Err.Clear

    f_SikenInfo_T16 = false
	p_bZenkiOnly = false

  '//�����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   T16_HAITOTANI "
  w_sSQL = w_sSQL & "  ,MAX(T16_KAISETU) AS T16_KAISETU " 	'//2009.10.16 ins
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T16_RISYU_KOJIN "
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   T16_NENDO = " & Cint(m_iNendo)
  w_sSQL = w_sSQL & " AND T16_GAKKA_CD = '" & m_sGakkaCd & "'"
  w_sSQL = w_sSQL & " AND T16_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'"
  'w_sSQL = w_sSQL & " AND T16_KAISETU = " & C_KAI_ZENKI

'2009.10.17 upd
'  w_sSQL = w_sSQL & " GROUP BY T16_HAITOTANI,T16_NENDO,T16_GAKKA_CD,T16_KAMOKU_CD,T16_KAISETU "
  w_sSQL = w_sSQL & " GROUP BY T16_HAITOTANI,T16_NENDO,T16_GAKKA_CD,T16_KAMOKU_CD "

  if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

  'Response.Write w_ssql  & "<BR>"
  'response.end

  '//�߂�l���
  If w_Rs.EOF = False Then

'2009.10.17 upd �^�ω����Ȃ��ƁA�����������������肳��Ȃ��B
'	if w_Rs("T16_KAISETU") = C_KAI_ZENKI then
	if cstr(w_Rs("T16_KAISETU")) = cstr(C_KAI_ZENKI) then

	    p_bZenkiOnly = True
	end if

	'�z���P�ʂ̎擾
	m_lHaitoTani = w_Rs("T16_HAITOTANI")

  End If

  f_SikenInfo_T16 = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  ���]���t���O�������Ă��邩���ׂ�
'********************************************************************************
function f_GetMihyoka(p_MiHyokaFlg)

  Dim w_sSQL,w_Rs
  Dim w_Table,w_FieldName,w_FromTable,w_KamokuCd

  On Error Resume Next
  Err.Clear

  f_GetMihyoka = false
  p_MiHyokaFlg = false

  if m_iKamokuKbn = C_JIK_JUGYO then
    w_Table = "T16"
    w_FromTable = "T16_RISYU_KOJIN"
    w_KamokuCd = "T16_KAMOKU_CD"
  else
    w_Table = "T34"
    w_FromTable = "T34_RISYU_TOKU"
    w_KamokuCd = "T34_TOKUKATU_CD"
  end if

  select case m_sSikenKBN
    case C_SIKEN_ZEN_TYU : w_FieldName = w_Table & "_DATAKBN_TYUKAN_Z"
    case C_SIKEN_ZEN_KIM : w_FieldName = w_Table & "_DATAKBN_KIMATU_Z"
    case C_SIKEN_KOU_TYU : w_FieldName = w_Table & "_DATAKBN_TYUKAN_K"
    case C_SIKEN_KOU_KIM : w_FieldName = w_Table & "_DATAKBN_KIMATU_K"
  end select

  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   " & w_FieldName & " as MIHYOKA "
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL &       w_FromTable
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   " & w_Table & "_NENDO = " & Cint(m_iNendo) & " and "
  w_sSQL = w_sSQL & "   " & w_KamokuCd & " = '" & m_sKamokuCd & "' and "
  w_sSQL = w_sSQL & "   " & w_Table & "_HAITOGAKUNEN = " & Cint(m_sGakuNo) & " and "
  w_sSQL = w_sSQL & "   " & w_Table & "_GAKKA_CD     = '" & m_sGakkaCd & "' and "
  w_sSQL = w_sSQL &     w_FieldName & "= 4 "


  If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function

  'Response.Write " 1"

  if not w_Rs.EOF then p_MiHyokaFlg = true

  f_GetMihyoka = true

end function


'********************************************************************************
'*  [�@�\]  �R���{�őI�����ꂽ�Ȗڂ̉Ȗڋ敪�y�сA�K�C�I���敪�𒲂ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Function f_GetKamokuInfo(p_iKamoku_Kbn,p_iHissen_Kbn,p_ilevelFlg)
  Dim w_sSQL
  Dim w_Rs

  On Error Resume Next
  Err.Clear

  f_GetKamokuInfo = false

  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   T15_RISYU.T15_KAMOKU_KBN"
  w_sSQL = w_sSQL & "   ,T15_RISYU.T15_HISSEN_KBN"
  w_sSQL = w_sSQL & "   ,T15_RISYU.T15_LEVEL_FLG"
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T15_RISYU"
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(m_sGakuNo) + 1
  w_sSQL = w_sSQL & " AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
  w_sSQL = w_sSQL & " AND T15_RISYU.T15_KAMOKU_CD='" & m_sKamokuCd & "' "

'response.write w_sSQL
'response.end

  If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

  '//�߂�l���
  If w_Rs.EOF = False Then
    p_iKamoku_Kbn = w_Rs("T15_KAMOKU_KBN")
    p_iHissen_Kbn = w_Rs("T15_HISSEN_KBN")
    p_ilevelFlg = w_Rs("T15_LEVEL_FLG")
  End If

  f_GetKamokuInfo = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �����N���X���ǂ����𒲂ׂ�
'*  [����]
'********************************************************************************
Function f_GetKongoClass(p_iGakunen,p_iClass,p_iKongo)
  Dim w_sSQL
  Dim w_Rs

  On Error Resume Next
  Err.Clear

  f_GetKongoClass = false

    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M05_SYUBETU "
    w_sSQL = w_sSQL & "FROM M05_CLASS "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_GAKUNEN = " & p_iGakunen & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_CLASSNO = " & p_iClass & " "

  If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

  'Response.Write "2"

  '//�߂�l���
  If w_Rs.EOF = False Then
    p_iKongo = Cint(w_Rs("M05_SYUBETU"))
  End If

  f_GetKongoClass = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �f�[�^�̎擾
'********************************************************************************
Function f_GetStudent()

  Dim w_sSQL
  Dim w_FieldName
  Dim w_Table
  Dim w_TableName
  Dim w_KamokuName

  On Error Resume Next
  Err.Clear

  f_GetStudent = false

  if m_iKamokuKbn = C_JIK_JUGYO then  '�ʏ���Ƃ̏ꍇ
    w_Table = "T16"
    w_TableName = "T16_RISYU_KOJIN"
    w_KamokuName = "T16_KAMOKU_CD"
  else
    w_Table = "T34"
    w_TableName = "T34_RISYU_TOKU"
    w_KamokuName = "T34_TOKUKATU_CD"
  end if

  '//�����A���l���͂ɂ��A����Ă���t�B�[���h��ς���
  if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then

    if m_bNiteiFlg and m_iKamokuKbn = C_JIK_JUGYO then
      w_FieldName = "HTEN"
    else
      w_FieldName = "SEI"
    end if

  else
    w_FieldName = "HYOKA"
  end if

  '//�������ʂ̒l���ꗗ��\��
  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI1, "
  w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI2, "
  w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI3, "
  w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI4, "

  Select Case m_sSikenKBN
    Case C_SIKEN_ZEN_TYU

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI,"
      w_sSQL = w_sSQL & w_Table & "_DATAKBN_TYUKAN_Z as DataKbn ,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z AS KEKA,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS KEKA_NASI,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_Z AS CHIKAI,"
      w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z as SOUJI,"
      w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z as JYUNJI, "

      if m_iKamokuKbn = C_JIK_JUGYO then
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI, "
      end if

    Case C_SIKEN_ZEN_KIM

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI,"
      w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_Z as DataKbn,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS KEKA_NASI,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_Z AS CHIKAI,"
      w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z as SOUJI, "
      w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z as JYUNJI, "

      if m_iKamokuKbn = C_JIK_JUGYO then
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI, "
      end if

    Case C_SIKEN_KOU_TYU

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K AS KEKA,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS KEKA_NASI,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_K AS CHIKAI,"
      w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K as SOUJI, "
      w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K as JYUNJI, "
      w_sSQL = w_sSQL & w_Table & "_DATAKBN_TYUKAN_K as DataKbn,"

      if m_iKamokuKbn = C_JIK_JUGYO then
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI, "
      end if

      '2002/12/25
      '���㍂��p��select�t�B�[���h�̒ǉ�
      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA_ZK,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_Z AS CHIKAI_ZK,"

      '2003/01/06 UPD �e�[�u���؂蕪��
      w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_Z AS KOUSINBI_ZK, "   '�O����
      w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_K AS KOUSINBI_TK, "   '�������

    Case C_SIKEN_KOU_KIM

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI_ZT,"
      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI_KT,"
      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI_KK,"

      w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z AS KEKA_ZT,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z AS KEKA_ZK,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K AS KEKA_KT,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_K AS KEKA,"

      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z AS KEKA_NASI_ZT,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z AS KEKA_NASI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K AS KEKA_NASI_KT,"
      w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_K AS KEKA_NASI,"

      w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_Z AS CHIKAI_ZT,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_Z AS CHIKAI_ZK,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_TYUKAN_K AS CHIKAI_KT,"
      w_sSQL = w_sSQL & w_Table & "_CHIKAI_KIMATU_K AS CHIKAI,"

      w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI,"

      w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K as SOUJI, "
      w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K as JYUNJI, "

      w_sSQL = w_sSQL & w_Table & "_SAITEI_JIKAN, "
      w_sSQL = w_sSQL & w_Table & "_KYUSAITEI_JIKAN, "

      w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_K as DataKbn,"
      w_sSQL = w_sSQL & w_Table & "_DATAKBN_KIMATU_Z as DataKbn_ZK,"

      if m_iKamokuKbn = C_JIK_JUGYO then
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_Z AS HYOKAYOTEI_ZT, "
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_Z AS HYOKAYOTEI_ZK, "
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_TYUKAN_K AS HYOKAYOTEI_KT, "
        w_sSQL = w_sSQL & " T16_HYOKAYOTEI_KIMATU_K AS HYOKAYOTEI, "

        w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_Z AS KOUSINBI_ZK, "
        w_sSQL = w_sSQL & " T16_KOUSINBI_KIMATU_K AS KOUSINBI_KK, "

        '2002/12/25
        '���㍂��p��select�t�B�[���h�̒ǉ�
        w_sSQL = w_sSQL & " T16_KOUSINBI_TYUKAN_K AS KOUSINBI_TK, "   '�������
      end if

  End Select

 '��������p�@INS 2005/06/13 ����
'�O���@����
  w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z as SOUJI1, "
  w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z as JYUNJI1, "
'�O���@����
  w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z as SOUJI2, "
  w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z as JYUNJI2, "
'����@����
  w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K as SOUJI3, "
  w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K as JYUNJI3, "
'����@����
  w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K as SOUJI4, "
  w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K as JYUNJI4, "

  w_sSQL = w_sSQL & " T13_GAKUSEI_NO AS GAKUSEI_NO,"
  w_sSQL = w_sSQL & " T13_GAKUSEKI_NO AS GAKUSEKI_NO,"
  w_sSQL = w_sSQL & " T11_SIMEI AS SIMEI, "

  w_sSQL = w_sSQL & " T13_SYUSEKI_NO1 AS SYUSEKI_NO1, "
  w_sSQL = w_sSQL & " T13_SYUSEKI_NO2 AS SYUSEKI_NO2, "

  if m_iKamokuKbn = C_JIK_JUGYO then
    w_sSQL = w_sSQL & "   T16_SELECT_FLG, "
    w_sSQL = w_sSQL & "   T16_LEVEL_KYOUKAN, "
    w_sSQL = w_sSQL & "   T16_OKIKAE_FLG, "
    If m_sGakkoNO = cstr(C_NCT_KURUME) OR m_sGakkoNO = cstr(C_NCT_NUMAZU) then
	    w_sSQL = w_sSQL & "   T16_MENJYO_FLG, "
	End If
  Else
	    w_sSQL = w_sSQL & "   0 AS T16_MENJYO_FLG, "
  end if

	'2003.8.25 �Ə��t���O�ǉ� �Ə��t���O��1�̏ꍇ�A������ɐ��т��R�s�[���Ȃ��ׁBITO �i�v���đΉ��j
	'2004.2.18 ���Òǉ��@�Ə��t���O�ǉ� �Ə��t���O��1�̏ꍇ�A������ɐ��т��R�s�[���Ȃ��ׁB�i���ÑΉ��j
    'If m_sGakkoNO = cstr(C_NCT_KURUME) then
    'If m_sGakkoNO = cstr(C_NCT_KURUME) OR m_sGakkoNO = cstr(C_NCT_NUMAZU) then
	'    w_sSQL = w_sSQL & "   T16_MENJYO_FLG, "
	'End If

  w_sSQL = w_sSQL & w_Table & "_HYOKA_FUKA_KBN as HYOKA_FUKA "

  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL &     w_TableName & ","
  w_sSQL = w_sSQL & "   T11_GAKUSEKI,"
  w_sSQL = w_sSQL & "   T13_GAKU_NEN "

  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL &       w_Table & "_NENDO = " & Cint(m_iNendo)
  w_sSQL = w_sSQL & " AND " & w_Table & "_GAKUSEI_NO = T11_GAKUSEI_NO "
  w_sSQL = w_sSQL & " AND " & w_Table & "_GAKUSEI_NO = T13_GAKUSEI_NO "
  w_sSQL = w_sSQL & " AND T13_GAKUNEN = " & cint(m_sGakuNo)

  w_sSQL = w_sSQL & " AND " & w_KamokuName & " = '" & m_sKamokuCd & "' "
  w_sSQL = w_sSQL & " AND " & w_Table & "_NENDO = T13_NENDO "

  '//�V���l����͖Ə��Ȗڂ͕\�����Ȃ� 2004/02/19
  if m_iKamokuKbn = C_JIK_JUGYO then
	  If m_sGakkoNO = cstr(C_NCT_NIIHAMA) then
		  w_sSQL = w_sSQL & " AND NVL(" & w_Table & "_MENJYO_FLG,0) <> 1 "
	  END IF
  end if

  if m_iKamokuKbn = C_JIK_JUGYO then
    '//�u�����̐��k�͂͂���(C_TIKAN_KAMOKU_MOTO = 1    '�u����)
    w_sSQL = w_sSQL & " AND T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
  end if

'****************************
  if m_iKongou <> C_CLASS_KONGO Then
    w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)

    'T13_GAKUSEKI_NO�Ń\�[�g����悤�ɏC���B
    'T13��T16�Ŋw�Дԍ����Ⴄ�ꍇ�����������ׂɔ��o�B
    '�l���C�����쐬�Ŋw�Дԍ����X�V���Ă��Ȃ��\������B2003.08.05 ITO
    w_sSQL = w_sSQL & " ORDER BY T13_GAKUSEKI_NO "

  else
    if m_sGakkoNO = C_NCT_KUMAMOTO Then
      w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
      w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "
    else
      if Cint(m_iKamoku_Kbn) <> C_KAMOKU_SENMON Then
				'INS STR 2010/05/20 iwata ���m �����N���X�A��ʁA�w��Ȗڂ́@�w�ȕ\���ɂ���
				If m_sGakkoNO = C_NCT_KOCHI Then
					If (m_sKamokuCd = 140046 ) OR (m_sKamokuCd = 140050 ) OR (m_sKamokuCd = 140047 ) OR (m_sKamokuCd = 140048 ) OR (m_sKamokuCd = 180011 ) OR (m_sKamokuCd = 180012 ) OR (m_sKamokuCd = 180013 ) OR (m_sKamokuCd = 180014 ) Then
						w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
						w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO1 "
					Else
						w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
						w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
					End If
				Else
				'INS END 2010/05/20 iwata
          w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
          w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
				End if
      else
		if (m_sGakkoNO = C_NCT_MAIZURU) Then
			'INS 2007/06/11
			If (m_sKamokuCd = 80001) OR (m_sKamokuCd = 80002) Then
				w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
				w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
			Else
			    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
				w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO1 "
			End If
			'INS END 2007/06/11
			'DEL 2007/06/11  w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
			'DEL 2007/06/11  w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "
		else
			If m_sGakkoNO = C_NCT_OKINAWA Then
				If m_sKamokuCd >= 900000 Then
					w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
					w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
				Else
				    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
					w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO1 "
				End If
			Else
				w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO1 "
			End If
		end if
      end if
    end if
  end if
'****************************

  'w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "

'Response.Write w_sSQL
'response.end

  If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit function

'Response.Write w_sSQL

  m_iSouJyugyou = gf_SetNull2String(m_Rs("SOUJI"))
  m_iJunJyugyou = gf_SetNull2String(m_Rs("JYUNJI"))

  '//ں��ރJ�E���g�擾
  m_rCnt = gf_GetRsCount(m_Rs)

  f_GetStudent = true

End Function



'********************************************************************************
'*  [�@�\]  �f�[�^�̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Function f_Nyuryokudate()

  Dim w_sSysDate
  Dim w_Rs

  On Error Resume Next
  Err.Clear

  f_Nyuryokudate = false

  m_bKekkaNyuryokuFlg = false   '���ۓ����׸�
  ''m_bSeiInpFlg = false

  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL & "   T24_SEISEKI_KAISI, "
  w_sSQL = w_sSQL & "   T24_SEISEKI_SYURYO, "
  w_sSQL = w_sSQL & "   T24_KEKKA_KAISI, "
  w_sSQL = w_sSQL & "   T24_KEKKA_SYURYO, "
  w_sSQL = w_sSQL & "   T24_IDOU_SYURYO, "
  w_sSQL = w_sSQL & "   M01_SYOBUNRUIMEI, "
  w_sSQL = w_sSQL & "   SYSDATE "
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T24_SIKEN_NITTEI, "
  w_sSQL = w_sSQL & "   M01_KUBUN"
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   M01_SYOBUNRUI_CD = T24_SIKEN_KBN"
  w_sSQL = w_sSQL & " AND M01_NENDO = T24_NENDO"
  w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD=" & cint(C_SIKEN)
  w_sSQL = w_sSQL & " AND T24_NENDO=" & Cint(m_iNendo)
  w_sSQL = w_sSQL & " AND T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
  w_sSQL = w_sSQL & " AND T24_SIKEN_CD='0'"
  w_sSQL = w_sSQL & " AND T24_GAKUNEN=" & Cint(m_sGakuNo)

  If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function

'   Response.Write "4" & w_sSQL

  If w_Rs.EOF Then
   Response.Write "  EOF "
    exit function
  Else
    m_sSikenNm = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))    '��������
    m_iNKaishi = gf_SetNull2String(w_Rs("T24_SEISEKI_KAISI"))   '���ѓ��͊J�n��
    m_iNSyuryo = gf_SetNull2String(w_Rs("T24_SEISEKI_SYURYO"))    '���ѓ��͏I����
    m_iKekkaKaishi = gf_SetNull2String(w_Rs("T24_KEKKA_KAISI"))   '���ۓ��͊J�n
    m_iKekkaSyuryo = gf_SetNull2String(w_Rs("T24_KEKKA_SYURYO"))  '���ۓ��͏I��

    m_iIdouEnd = gf_SetNull2String(w_Rs("T24_IDOU_SYURYO"))  '�ٓ��ΏۏI��

    w_sSysDate = gf_SetNull2String(w_Rs("SYSDATE"))         '�V�X�e�����t
  End If

  '���ۓ��͉\�׸�
  If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
    m_bKekkaNyuryokuFlg = True
  End If

  f_Nyuryokudate = true

End Function

'********************************************************************************
'*  [�@�\]  �f�[�^�̎擾
'********************************************************************************
Function f_Syukketu2New(p_gaku,p_kbn)
  Dim w_GAKUSEI_NO
  Dim w_SYUKKETU_KBN

  f_Syukketu2New = ""
  w_GAKUSEI_NO = ""
  w_SYUKKETU_KBN = ""
  w_SKAISU = ""

  If m_SRs.EOF Then
    Exit Function
  Else
    Do Until m_SRs.EOF
      w_GAKUSEI_NO = m_SRs("T21_GAKUSEKI_NO")
      w_SYUKKETU_KBN = m_SRs("T21_SYUKKETU_KBN")
      w_SKAISU = gf_SetNull2String(m_SRs("KAISU"))

      If Cstr(w_GAKUSEI_NO) = Cstr(p_gaku) AND cstr(w_SYUKKETU_KBN) = cstr(p_kbn) Then
        f_Syukketu2New = w_SKAISU
        Exit Do
      End If

      m_SRs.MoveNext
    Loop

    m_SRs.MoveFirst
  End If

End Function

'********************************************************************************
'*  [�@�\]  �m�茇�ې��A�x�������擾�B
'*  [����]  p_iNendo�@ �@�F�@�����N�x
'*          p_iSikenKBN�@�F�@�����敪
'*          p_sKamokuCD�@�F�@�ȖڃR�[�h
'*          p_sGakusei �@�F�@�T�N�Ԕԍ�
'*  [�ߒl]  p_iKekka   �@�F�@���ې�
'*          p_ichikoku �@�F�@�x����
'*          0�F����C��
'*  [����]  �����敪�ɓ����Ă���A���ې��A�x�������擾����B
'*      2002.03.20
'*      NULL��0�ɕϊ����Ȃ����߂ɁA�֐������W���[�����ō쐬�iCACommon.asp����R�s�[�j
'********************************************************************************
Function f_GetKekaChi(p_iNendo,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku)
  Dim w_sSQL
  Dim w_Rs
  Dim w_sKek,w_sChi
  Dim w_Table,w_TableName
  Dim w_Kamoku

  On Error Resume Next
  Err.Clear

  f_GetKekaChi = false

  p_iKekka = ""
  p_iChikoku = ""

  '���ʎ��ƁA���̑�(�ʏ�Ȃ�)�̐؂蕪��
  if trim(m_sSyubetu) = "TOKU" then
    w_Table = "T34"
    w_TableName = "T34_RISYU_TOKU"
    w_Kamoku = "T34_TOKUKATU_CD"
  else
    w_Table = "T16"
    w_TableName = "T16_RISYU_KOJIN"
    w_Kamoku = "T16_KAMOKU_CD"
  end if

  '/�����敪�ɂ���Ď���Ă���A�t�B�[���h��ς���B
  Select Case p_iSikenKBN
    Case C_SIKEN_ZEN_TYU
      w_sKek   = w_Table & "_KEKA_TYUKAN_Z"
      w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_Z"
      w_sChi   = w_Table & "_CHIKAI_TYUKAN_Z"
    Case C_SIKEN_ZEN_KIM
      w_sKek   = w_Table & "_KEKA_KIMATU_Z"
      w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_Z"
      w_sChi   = w_Table & "_CHIKAI_KIMATU_Z"
    Case C_SIKEN_KOU_TYU
      w_sKek   = w_Table & "_KEKA_TYUKAN_K"
      w_sKekG  = w_Table & "_KEKA_NASI_TYUKAN_K"
      w_sChi   = w_Table & "_CHIKAI_TYUKAN_K"
    Case C_SIKEN_KOU_KIM
      w_sKek   = w_Table & "_KEKA_KIMATU_K"
      w_sKekG  = w_Table & "_KEKA_NASI_KIMATU_K"
      w_sChi   = w_Table & "_CHIKAI_KIMATU_K"
  End Select

  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "
  w_sSQL = w_sSQL &   w_sKek   & " as KEKA, "
  w_sSQL = w_sSQL &   w_sKekG  & " as KEKA_NASI, "
  w_sSQL = w_sSQL &   w_sChi   & " as CHIKAI "
  w_sSQL = w_sSQL & " FROM "   & w_TableName
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "      " & w_Table & "_NENDO =" & p_iNendo
  w_sSQL = w_sSQL & "  AND " & w_Table & "_GAKUSEI_NO= '" & p_sGakusei & "'"
  w_sSQL = w_sSQL & "  AND " & w_Kamoku & "= '" & p_sKamokuCD & "'"

'response.write w_sSQL
'response.end


  If gf_GetRecordset(w_Rs, w_sSQL) <> 0 Then exit function

  ' Response.Write "5"

  '//�߂�l���
  If w_Rs.EOF = False Then
    p_iKekka = gf_SetNull2String(w_Rs("KEKA"))
    p_iChikoku = gf_SetNull2String(w_Rs("CHIKAI"))
  End If

  f_GetKekaChi = true

  Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\] �ٓ��`�F�b�N
'********************************************************************************
Sub s_IdouCheck(p_GakusekiNo,p_IdouKbn,p_IdouName,p_bNoChange)
  Dim w_IdoutypeName  '�ٓ��󋵖�

  w_IdoutypeName = ""
  p_IdouName = ""


  m_Date = m_iIdouEnd
'debug
'response.write "p_GakusekiNo = " & p_GakusekiNo & "<br>"
'response.write "m_Date = " & m_Date & "<br>"
'response.write "m_iNendo = " & m_iNendo & "<br>"

  p_IdouKbn = gf_Get_IdouChk(p_GakusekiNo,m_Date,m_iNendo,w_IdoutypeName)
'response.write "w_IdoutypeName = " & w_IdoutypeName & "<br>"
'response.write "p_IdouKbn = " & p_IdouKbn & "<br>"

  if Cstr(p_IdouKbn) <> "" and Cstr(p_IdouKbn) <> CStr(C_IDO_FUKUGAKU) AND _
    Cstr(p_IdouKbn) <> Cstr(C_IDO_TEI_KAIJO) AND _
    Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKA) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_KOKUHI) Then

    p_IdouName = "[" & w_IdoutypeName & "]"
    p_bNoChange = True
  end if

end Sub

'********************************************************************************
'*  [�@�\] ���т̃Z�b�g
'********************************************************************************
Sub s_SetGrades(p_sSeiseki,p_sHyoka,p_bNoChange)

  p_sSeiseki = gf_SetNull2String(m_Rs("SEI"))


    '�w�N�������̏ꍇ�̂�
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
      'If gf_SetNull2String(m_Rs("SEI")) = "" Then
        p_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))
      End If
    End If


  '//�ʏ���Ƃ̂Ƃ�
  if m_iKamokuKbn = C_JIK_JUGYO then

    p_bNoChange = False

	'2004.02.20 ITO
	'�v���Ă̏ꍇ�A�Ə��Ȗڂ̓_���͔�\��
	If m_sGakkoNO = cstr(C_NCT_KURUME) then

		if cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG"))) = "1" Then

			p_bNoChange = True

		End If

	End If

    '//�Ȗڂ��I���Ȗڂ̏ꍇ�́A���k���I�����Ă��邩�ǂ����𔻕ʂ���B�I�������Ȃ����k�͓��͕s�Ƃ���B
    if cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then

		if cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then p_bNoChange = True

    else
      if Cstr(m_iLevelFlg) = "1" then
        if isNull(m_Rs("T16_LEVEL_KYOUKAN")) = true then
          p_bNoChange = True
        else
          if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
            p_bNoChange = True
          End if
        End if
      End if
    end if

  end if

end Sub

'********************************************************************************
'*  [�@�\] ���ہA�x���̓��X�v�̎擾
'********************************************************************************
Sub s_SetKekkaTotal(p_sKekkasu,p_sChikaisu)
  Dim w_sData
  Dim w_iKekka_rui,w_iChikoku_rui

  '//����
  p_sKekkasu = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_KEKKA))

  '//1����
  w_sData = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_KEKKA_1))

  if p_sKekkasu = "" and w_sData = "" then
    p_sKekkasu = ""
  else
    p_sKekkasu = cint(gf_SetNull2Zero(p_sKekkasu)) + cint(gf_SetNull2Zero(w_sData))
  end if

  '//�x����
  p_sChikaisu = gf_SetNull2String(f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_TIKOKU))

  '//���ސ�
  w_sData = f_Syukketu2New(m_Rs("GAKUSEKI_NO"),C_KETU_SOTAI)

  if p_sChikaisu = "" and w_sData = "" then
    p_sChikaisu = ""
  else
    p_sChikaisu = cint(gf_SetNull2Zero(p_sChikaisu)) + cint(gf_SetNull2Zero(w_sData))
  end if

  '�u�o�����ۂ��ݐρv�Łu�O�����ԂłȂ��v�̏ꍇ

  if cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) and m_sSikenKBN <> C_SIKEN_ZEN_TYU then
    '�ȑO�̎����œo�^����Ă���f�[�^���擾
    call f_GetKekaChi(m_iNendo,m_iShikenInsertType,m_sKamokuCd,m_Rs("GAKUSEI_NO"),w_iKekka_rui,w_iChikoku_rui)

    '�ǂ����""�̎���""
    if p_sKekkasu = "" and w_iKekka_rui = "" then
      p_sKekkasu = ""
    else
      p_sKekkasu = cint(gf_SetNull2Zero(p_sKekkasu)) + cint(gf_SetNull2Zero(w_iKekka_rui))
    end if

    '�ǂ����""�̎���""
    if p_sChikaisu = "" and w_iChikoku_rui = "" then
      p_sChikaisu = ""
    else
      p_sChikaisu = cint(gf_SetNull2Zero(p_sChikaisu)) + cint(gf_SetNull2Zero(w_iChikoku_rui))
    end if
  end if

End Sub

'********************************************************************************
'*  [�@�\]  ���ہA�x�����̃Z�b�g
'********************************************************************************
Sub s_SetKekka(p_sKekka,p_sKekkaGai,p_sChikai)

  p_sKekka = gf_SetNull2String(m_Rs("KEKA"))
  p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI"))
  p_sChikai = gf_SetNull2String(m_Rs("CHIKAI"))


  '2002/12/25
  '���㍂��̏ꍇ�A�O���J�݉Ȗڂ͎����ɂ���ăR�s�[���̎�����ς���
If m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then

    '������Ԃ̎��A�O����������Z�b�g
    If m_sSikenKBN = C_SIKEN_KOU_TYU and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))  '�O������
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_TK"))  '�������

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))     '���ې�
        p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK")) '���ۑΏۊO
        p_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))    '�x����
      End If
    End If

    '��������̎��A������Ԃ���Z�b�g
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_TK"))  '�������
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))  '�������

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sKekka = gf_SetNull2String(m_Rs("KEKA_KT"))     '���ې�
        p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_KT")) '���ۑΏۊO
        p_sChikai = gf_SetNull2String(m_Rs("CHIKAI_KT"))    '�x����
      End If
    End If

 '������i�O����->������ɃZ�b�g�j
 Else

    '//�w�N�������̏ꍇ�̂�
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

      'If gf_SetNull2String(m_Rs("KEKA")) = "" Then
      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sKekka = gf_SetNull2String(m_Rs("KEKA_ZK"))     '���ې�
        p_sKekkaGai = gf_SetNull2String(m_Rs("KEKA_NASI_ZK")) '���ۑΏۊO
        p_sChikai = gf_SetNull2String(m_Rs("CHIKAI_ZK"))    '�x����
      End If
    End If
 End If

End Sub

'********************************************************************************
'*  [�@�\]  �e�[�u���T�C�Y�̃Z�b�g
'********************************************************************************
Sub s_SetTableWidth(p_TableWidth)

  p_TableWidth = 620

  '//���ۊO�\���t���O�I��
  if m_KekkaGaiDispFlg then
    p_TableWidth = p_TableWidth + 55
  end if

End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'********************************************************************************
Sub showPage()
  Dim w_sSeiseki
  Dim w_sHyoka

  Dim w_sChikai
  Dim w_sChikaisu

  Dim w_sKekka
  Dim w_sKekkaGai
  Dim w_sKekkasu

  Dim i

  Dim w_lSeiTotal '���э��v
  Dim w_lGakTotal '�w���l��

  Dim w_IdouKbn '�ٓ��^�C�v
  Dim w_IdouName

  Dim w_sInputClass
  Dim w_sInputClass1
  Dim w_sInputClass2

  Dim w_Disabled
  Dim w_Disabled2
  Dim w_TableWidth

  Dim w_Style
  Dim w_Style2

  '//�x���A����(�Ώ�)�A����(�ΏۊO)�̍��v
  Dim wChikokuSum,wKekkaSum,wKekkaGaiSum

  wChikokuSum = 0
  wKekkaSum = 0
  wKekkaGaiSum = 0

  w_lSeiTotal = 0
  w_lGakTotal = 0
  i = 1

  '//NN�Ή�
  If session("browser") = "IE" Then
    w_sInputClass  = "class='num'"
    w_sInputClass1 = "class='num'"
    w_sInputClass2 = "class='num'"
  Else
    w_sInputClass = ""
    w_sInputClass1 = ""
    w_sInputClass2 = ""
  End If

  w_Style = "border:solid 1px #000000;font-size:10pt;padding:0px;"
  w_Style2 = "border:solid 1px #000000;font-size:10pt;padding:1px;"
  '//�e�[�u���T�C�Y�̃Z�b�g
  Call s_SetTableWidth(w_TableWidth)

%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
  </head>
  <body LANGUAGE="javascript">
  <center>
  <form name="frm" method="post">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td align="center" nowrap>
				<table style="border-layout:fixed;border-collapse:collapse;border-style:solid;border-color:#000000;border-width:1px;padding:0px;margin:0px;" border="1" align="center" width="<%=w_TableWidth%>">
					<caption>���ѓo�^���</caption>
					<tr>
						<td style="<%=w_Style%>" align="center" nowrap >���ѓ��͊���</td>
						<td style="<%=w_Style%>" align="center"><%=f_ShikenMei()%></td>
						<td style="<%=w_Style%>" align="center" nowrap>�o�͓�</td>
						<td style="<%=w_Style%>" align="center"><%=FormatTime(now(),"YYYY/MM/DD HH24:II:SS")%></td>
						<td style="<%=w_Style%>" align="center" nowrap>���[�U�[��</td>
						<td style="<%=w_Style%>" align="center"><%=Session("USER_NM")%></td>
					</tr>
					<tr>
						<td style="<%=w_Style%>" align="center" nowrap>���{�Ȗ�</td>
						
						<%
						w_str = m_sGakuNo & "�N�@" & gf_GetClassName(m_iNendo,Cint(m_sGakuNo),m_sClassNo) & "�@" & f_GetKamokuName(Cint(m_sGakuNo),m_sGakkaCd,m_sKamokuCd)
						%>
						<td style="<%=w_Style%>" align="center" colspan="5"><%=w_str%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" valign="bottom" nowrap>
				<table style="border-layout:fixed;border-collapse:collapse;border-style:solid;border-color:#000000;border-width:1px;padding:0px;margin:0px;" border="1" align="center" width="<%=w_TableWidth%>">

					<tr>
						<td style="<%=w_Style%>" align="center" nowrap colspan="9">
							�����Ɛ�&nbsp;&nbsp;<%= m_iSouJyugyou %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����Ɛ�&nbsp;&nbsp;<%= m_iJunJyugyou %></td>
					</tr>
					<tr>
						<td style="<%=w_Style%>" align="center" rowspan="2" width="75" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></td>
						<td style="<%=w_Style%>" align="center" rowspan="2" width="150" nowrap>���@��</td>
						<td style="<%=w_Style%>" align="center" colspan="4" width="200" nowrap>���ї���</td>

						<td style="<%=w_Style%>" align="center" rowspan="2" width="65" nowrap>����</td>
						<td style="<%=w_Style%>" align="center" rowspan="2" width="65" nowrap>�x��</td>
						<td style="<%=w_Style%>" align="center" rowspan="2" width="65" nowrap>����</td>
					</tr>

					<tr>
						<td style="<%=w_Style%>" align="center" width="50" nowrap>�O��</td>
						<td style="<%=w_Style%>" align="center" width="50" nowrap>�O��</td>
						<td style="<%=w_Style%>" align="center" width="50" nowrap>�㒆</td>
						<td style="<%=w_Style%>" align="center" width="50" nowrap>�w��</td>

					</tr>
  <%

    m_Rs.MoveFirst

    Do Until m_Rs.EOF
      j = j + 1

      w_sSeiseki  = ""
      w_sHyoka    = ""
      w_sChikai   = ""
      w_sChikaisu = ""
      w_sKekka    = ""
      w_sKekkaGai = ""
      w_sKekkasu  = ""
      w_bNoChange = false

''      Call gs_cellPtn(w_cell)
w_cell="disp"
      '�X�^�C���V�[�g�ݒ�

      '//���ہA�x�����̃Z�b�g
      Call s_SetKekka(w_sKekka,w_sKekkaGai,w_sChikai)

      '//���уf�[�^�Z�b�g
      Call s_SetGrades(w_sSeiseki,w_sHyoka,w_bNoChange)

      '//�ٓ��`�F�b�N
      Call s_IdouCheck(m_Rs("GAKUSEKI_NO"),w_IdouKbn,w_IdouName,w_bNoChange)

      '//���ہA�x���̓��X�v�̎擾
      Call s_SetKekkaTotal(w_sKekkasu,w_sChikaisu)

    '// 2003/10/06 INSERT
	if (m_sGakkoNO <> cstr(C_NCT_NUMAZU)) and (m_sGakkoNO <> cstr(C_NCT_NIIHAMA)) Then
      '//������0��,���v��0���傫���ꍇ
      if cint(gf_SetNull2Zero(w_sKekka)) = 0 and cint(gf_SetNull2Zero(w_sKekkasu)) > 0 Then
        w_sKekka = cint(gf_SetNull2Zero(w_sKekkasu))
      end if
	end if

    '// 2003/10/06 INSERT
	if (m_sGakkoNO <> C_NCT_NUMAZU) and (m_sGakkoNO <> C_NCT_NIIHAMA) Then
      '//�x����0��,�x�v��0���傫���ꍇ
      if cint(gf_SetNull2Zero(w_sChikai)) = 0 AND cint(gf_SetNull2Zero(w_sChikaisu)) > 0 Then
        w_sChikai = cint(gf_SetNull2Zero(w_sChikaisu))
      end if
	end if

      %>

      <tr>
        <td style="<%=w_Style2%>" align="center" width="75"><%=m_Rs("GAKUSEKI_NO")%></td>

	    <td style="<%=w_Style2%>" align="left" width="150" nowrap><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>

        <% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>

          <td style="<%=w_Style2%>" align="center" width="50"  nowrap ><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI1")))%></td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap ><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI2")))%></td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap ><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI3")))%></td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap ><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI4")))%></td>
        <% else %>

          <td style="<%=w_Style2%>" align="center" width="50"  nowrap >-</td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap >-</td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap >-</td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap >-</td>
        <% end if %>

        <!--�I���Ȗڂ̎��ɖ��I���̏ꍇ�A���͕s�B�܂��A�x�w�Ȃ�-->
        <% If w_bNoChange = True Then %>

          <td style="<%=w_Style2%>" align="center" width="65" nowrap >-</td>
          <td style="<%=w_Style2%>" align="center" width="65" nowrap >-</td>
          <td style="<%=w_Style2%>" align="center" width="65" nowrap >-</td>

          <% if m_KekkaGaiDispFlg then %>

            <!--<td class="<%=w_cell%>" align="center" width="55" nowrap >-</td>-->
          <% end if %>

        <% Else %>

          <!-- ���� (���l���́A�������́A���тȂ����͂ɂ�菈���𕪂���) -->
		<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM  OR m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>

				<td style="<%=w_Style2%>" align="center" width="65" nowrap >

					<%=w_sSeiseki%>

				</td>

		<% else %>

			<td style="<%=w_Style2%>" align="center" width="65" nowrap >-</td>

		<% end if %>

		<!-- �x�� -->
          <td style="<%=w_Style2%>" align="center" width="65" nowrap ><%=w_sChikai%></td>

		<!-- ���� -->
          <td style="<%=w_Style2%>" align="center" width="65" nowrap ><%=w_sKekka%></td>

          <% if m_KekkaGaiDispFlg then %>
            <!--<td class="<%=w_cell%>" align="center" width="55" nowrap ><%=w_sKekkaGai%></td>-->
          <% end if %>

          <%
            if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
              '�\���݂̂̏ꍇ�̍��v�E���ϒl�����߂�
              If IsNull(w_sSeiseki) = False and IsNumeric(CStr(w_sSeiseki)) = True Then
                w_lSeiTotal = w_lSeiTotal + CLng(w_sSeiseki)
                w_lGakTotal = w_lGakTotal + 1
              End If
            end if
          %>
        <%End If%>

      </tr>

      <%
        wChikokuSum = wChikokuSum + cint(gf_SetNull2Zero(w_sChikai))
		If w_bNoChange = false Then
        	wKekkaSum = wKekkaSum + cint(gf_SetNull2Zero(w_sKekka))
			wKekkaGaiSum = wKekkaGaiSum + cint(gf_SetNull2Zero(w_sKekkaGai))
		'Else

		End If


        m_Rs.MoveNext
        i = i + 1
      Loop

      %>
<!--
<tr>
        <td style="<%=w_Style2%>" align="center" width="75">3M50M</td>
	    <td style="<%=w_Style2%>" align="left" width="150" nowrap>���݁[</td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap >100</td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap >100</td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap >100</td>
          <td style="<%=w_Style2%>" align="center" width="50"  nowrap >100</td>
          <td style="<%=w_Style2%>" align="center" width="65"  nowrap >100</td>
          <td style="<%=w_Style2%>" align="center" width="65"  nowrap >1</td>
          <td style="<%=w_Style2%>" align="center" width="65"  nowrap >2</td>
</tr>
-->
      <% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
        <tr>
          <td style="<%=w_Style%>" align="right" colspan="6" nowrap >���v</td>
          <td style="<%=w_Style%>" align="center" width="65" nowrap ><%=w_lSeiTotal%></td>

          <td style="<%=w_Style%>" align="center" width="65" nowrap ><%=wChikokuSum%></td>
          <td style="<%=w_Style%>" align="center" width="65" nowrap ><%=wKekkaSum%></td>

          <% if m_KekkaGaiDispFlg then %>
            <!--<td style="font-size:10pt;" align="center" nowrap><%=wKekkaGaiSum%></td>-->
          <% end if %>

        </tr>

        <tr>
          <td style="<%=w_Style%>" align="right" colspan="6" nowrap>���ϓ_</td>
		  <%
		  if w_lGakTotal > 0 then
		  	w_avg = gf_Round(w_lSeiTotal / w_lGakTotal,1)
		  else
		  	w_avg = 0
		  end if
		  %>
          <td style="<%=w_Style%>" align="center" width="65" nowrap ><%=w_avg%></td>
          <td style="<%=w_Style%>" align="center" colspan="2" nowrap>&nbsp;</td>
        </tr>
      <% else %>
        <tr>
          <td style="<%=w_Style%>" align="center" colspan="7" nowrap>���v</td>
          <td style="<%=w_Style%>" align="center" nowrap><%=wChikokuSum%></td>
          <td style="<%=w_Style%>" align="center" nowrap><%=wKekkaSum%></td>

          <% if m_KekkaGaiDispFlg then %>
            <!--<td style="font-size:10pt;" align="center" nowrap><%=wKekkaGaiSum%></td>-->
          <% end if %>

        </tr>
      <% end if %>

				</table>
			</td>
		</tr>
	</table>

  <input type="hidden" name="hidTotal" value="<%=w_lSeiTotal%>">
  <input type="hidden" name="hidGakTotal" value="<%=w_lGakTotal%>">


  </form>
  </center>
  </body>
  </html>
<%
End sub
%>