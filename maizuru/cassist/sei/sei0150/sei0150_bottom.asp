<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0150_bottom.asp
' �@      �\: ���y�[�W ���ѓo�^�̌������s��
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
' ��      ��: 2002/06/21 shin
' ��      �X: 2003/04/11 hirota �F��󋵂��w�N�ʂɂ݂�悤�ɕύX
' ��      �X: 2003/05/13 hirota �v���č���p�@���ѓ��͎��͎�u���Ԃ�K�{���͂Ƃ���
' ��      �X: 2005/12/16 �����@�@��������̏ꍇ�ő���Ǝ��Ԑ����擾����f_GetJyugyoJIkan()��ǉ�
' ��      �X: 2018/10/16 ���{ �����w���Ή�
' ��      �X: 2019/06/17 ���� 80001�F�H�w��b�A80002�F��񃊃e���V�[�̓��ʏ�������߂�
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

  Dim m_iSouJyugyou1     '//�����Ǝ���	INS 2005/06/13 �����@��������p
  DIm m_iJunJyugyou1     '//�����Ǝ���
  Dim m_iSouJyugyou2     '//�����Ǝ���
  DIm m_iJunJyugyou2     '//�����Ǝ���
  Dim m_iSouJyugyou3     '//�����Ǝ���
  DIm m_iJunJyugyou3     '//�����Ǝ���
  Dim m_iSouJyugyou4     '//�����Ǝ���
  DIm m_iJunJyugyou4     '//�����Ǝ���


  Dim m_bSeiInpFlg      '//���͊��ԃt���O
  Dim m_bKekkaNyuryokuFlg   '//���ۓ��͉\�׸�(True:���͉� / False:���͕s��)

  Dim m_iShikenInsertType

  Dim m_sSyubetu

  '2002/06/21
  Dim m_iKamokuKbn        '//�Ȗڋ敪(0:�ʏ���ƁA1:���ʉȖ�)
  Dim m_sKamokuBunrui       '//�Ȗڕ���(01:�ʏ���ƁA02:�F��ȖځA03:���ʉȖ�)

  Dim m_iSeisekiInpType
  Dim m_Date
  Dim m_bZenkiOnly
  Dim m_SchoolFlg,m_KekkaGaiDispFlg,m_HyokaDispFlg

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

    '//�]���s�\��\�����邩�`�F�b�N
    if not gf_ChkDisp(C_DATAKBN_DISP,m_SchoolFlg) then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[2]"

    '//�]���s�\�`�F�b�N�̏������K�v�Ȃ�
    '//���]���t���O�𒲂ׂ�
    if m_SchoolFlg then
      if not f_GetMihyoka(m_MiHyokaFlg) then
        m_bErrFlg = True
        Exit Do
      end if
    end if

'Response.Write "[3]"

    '//���ۊO��\�����邩�`�F�b�N
    if not gf_ChkDisp(C_KEKKAGAI_DISP,m_KekkaGaiDispFlg) then
      m_bErrFlg = True
      Exit Do
    End If


    '//�]���\���\�����邩�`�F�b�N
    if not gf_ChkDisp(C_HYOKAYOTEI_DISP,m_HyokaDispFlg) then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[5]"

    '//���ѓ��͕��@�̎擾(0:�_��[C_SEISEKI_INP_TYPE_NUM]�A1:����[C_SEISEKI_INP_TYPE_STRING]�A2:���ہA�x��[C_SEISEKI_INP_TYPE_KEKKA])
    if not gf_GetKamokuSeisekiInp(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iSeisekiInpType) then
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[6]"

    '//�O���̂݊J�݂��ʔN�����ׂ�
    if not f_SikenInfo(m_bZenkiOnly) then
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[7]"

    '//���сA���ۓ��͊��ԃ`�F�b�N
    If not f_Nyuryokudate() Then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[8]"

    '//�o�����ۂ̎������擾
    '//�Ȗڋ敪(0:������,1:�ݐ�)
    If gf_GetKanriInfo(m_iNendo,m_iSyubetu) <> 0 Then
      m_bErrFlg = True
      Exit Do
    End If

'Response.Write "[9]"

    '//�F��O����擾
'   if not gf_GetNintei(m_iNendo,m_bNiteiFlg) then
    if not gf_GetGakunenNintei(m_iNendo,cint(m_sGakuNo),m_bNiteiFlg) then '2003.04.11 hirota
      m_bErrFlg = True
      Exit Do
    end if

'Response.Write "[10]"

    If m_iKamokuKbn = C_JIK_JUGYO then  '�ʏ���Ƃ̏ꍇ
      '//�Ȗڏ����擾
      '//�Ȗڋ敪(0:��ʉȖ�,1:���Ȗ�)�A�y�сA�K�C�I���敪(1:�K�C,2:�I��)�𒲂ׂ�
      '//���x���ʋ敪(0:��ʉȖ�,1:���x���ʉȖ�)�𒲂ׂ�
      If not f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) Then m_bErrFlg = True : Exit Do
    end if

'Response.Write "[11]"
    If not f_GetKongoClass(cint(m_sGakuNo),cint(m_sClassNo),m_iKongou) Then m_bErrFlg = True : Exit Do

    '//���сA�w���f�[�^�擾
    If not f_GetStudent() Then m_bErrFlg = True : Exit Do

    If m_Rs.EOF Then
      Call gs_showWhitePage("�l���C�f�[�^�����݂��܂���B","���ѓo�^")
      Exit Do
    End If

	IF m_sGakkoNO = cstr(C_NCT_FUKUSHIMA) THEN
		'//��������̏ꍇ,�ő�̎��Ǝ��Ԑ����擾 INS 2005/12/16����
		IF NOT f_GetJyugyoJIkan()Then m_bErrFlg = True : Exit Do
	    If m_Rs.EOF Then
	      Call gs_showWhitePage("�l���C�f�[�^�����݂��܂���B","���ѓo�^")
	      Exit Do
	    End If

	END IF

'Response.Write "[12]"
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

  m_iNendo   = request("txtNendo")
  m_sKyokanCd  = request("txtKyokanCd")

  m_sSikenKBN  = cint(request("sltShikenKbn"))
  m_sGakuNo  = cint(request("txtGakuNo"))
  m_sClassNo   = cint(request("txtClassNo"))
  m_sKamokuCd  = request("txtKamokuCd")
  m_sGakkaCd   = request("txtGakkaCd")
  m_sSyubetu   = trim(Request("SYUBETU"))
  m_iShikenInsertType = 0

  m_iKamokuKbn = cint(Request("hidKamokuKbn"))

  if m_iKamokuKbn = C_JIK_JUGYO then
    '�ʏ�Ȗ�
    m_sKamokuBunrui = C_KAMOKUBUNRUI_TUJYO
  else
    '���ʉȖ�
    m_sKamokuBunrui = C_KAMOKUBUNRUI_TOKUBETU
  end if

  m_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")

End Sub

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
  w_sSQL = w_sSQL & " FROM "
  w_sSQL = w_sSQL & "   T16_RISYU_KOJIN "
  w_sSQL = w_sSQL & " 	,T13_GAKU_NEN "		'2018.10.16 Add Kiyomoto �����w���Ή�
  w_sSQL = w_sSQL & " WHERE "
  w_sSQL = w_sSQL & "   T16_NENDO = " & Cint(m_iNendo)
'  w_sSQL = w_sSQL & " AND T16_GAKKA_CD = '" & m_sGakkaCd & "'"		'2018.10.16 Del Kiyomoto �����w���Ή�
  w_sSQL = w_sSQL & " AND T16_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'"
  'w_sSQL = w_sSQL & " AND T16_KAISETU = " & C_KAI_ZENKI
  '2018.10.16 Add Kiyomoto �����w���Ή� -->
  w_sSQL = w_sSQL & " AND T13_NENDO = T16_NENDO "
  w_sSQL = w_sSQL & " AND T13_GAKUNEN = T16_HAITOGAKUNEN "
  w_sSQL = w_sSQL & " AND T13_GAKUSEI_NO = T16_GAKUSEI_NO "
  w_sSQL = w_sSQL & " AND T13_CLASS = " & Cint(m_sClassNo)
  '2018.10.16 Add Kiyomoto �����w���Ή� <--
  w_sSQL = w_sSQL & " GROUP BY T16_HAITOTANI,T16_NENDO,T16_GAKKA_CD,T16_KAMOKU_CD,T16_KAISETU "

  if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

  'Response.Write w_ssql
  'response.end

  '//�߂�l���
  If w_Rs.EOF = False Then

	if w_Rs("T16_KAISETU") = C_KAI_ZENKI then
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
        w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
        w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
      else
		if (m_sGakkoNO = C_NCT_MAIZURU) Then
			'INS 2007/06/11
			'--2019/06/17 DELETE FUJIBAYASHI(80001�A80002�̓��ʏ�������߂�)
			'If (m_sKamokuCd = 80001) OR (m_sKamokuCd = 80002) Then
			'	w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
			'	w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO2 "
			'Else
			'--2019/06/17 DELETE END
			    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
				w_sSQL = w_sSQL & " ORDER BY T13_SYUSEKI_NO1 "
			'End If		'--2019/06/17 DELETE FUJIBAYASHI(80001�A80002�̓��ʏ�������߂�)
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

  '��������p�@INS 2005/06/13 ����	
  m_iSouJyugyou1 = gf_SetNull2String(m_Rs("SOUJI1"))
  m_iJunJyugyou1 = gf_SetNull2String(m_Rs("JYUNJI1"))
  m_iSouJyugyou2 = gf_SetNull2String(m_Rs("SOUJI2"))
  m_iJunJyugyou2 = gf_SetNull2String(m_Rs("JYUNJI2"))
  m_iSouJyugyou3 = gf_SetNull2String(m_Rs("SOUJI3"))
  m_iJunJyugyou3 = gf_SetNull2String(m_Rs("JYUNJI3"))
  m_iSouJyugyou4 = gf_SetNull2String(m_Rs("SOUJI4"))
  m_iJunJyugyou4 = gf_SetNull2String(m_Rs("JYUNJI4"))


  '//ں��ރJ�E���g�擾
  m_rCnt = gf_GetRsCount(m_Rs)

  f_GetStudent = true

End Function


'********************************************************************************
'*  [�@�\]  ���Ǝ��Ԑ��f�[�^�̎擾�i��������̏ꍇ�j
'*			�쐬 2005/12/16 ����
'*			MAX(���Ǝ��Ԑ�)�Ŏ擾����
'********************************************************************************
Function f_GetJyugyoJIkan()

  Dim w_sSQL
  Dim w_FieldName
  Dim w_Table
  Dim w_TableName
  Dim w_KamokuName
  Dim m_RsMax

  On Error Resume Next
  Err.Clear

  f_GetJyugyoJIkan = false

  if m_iKamokuKbn = C_JIK_JUGYO then  '�ʏ���Ƃ̏ꍇ
    w_Table = "T16"
    w_TableName = "T16_RISYU_KOJIN"
    w_KamokuName = "T16_KAMOKU_CD"
  else
    w_Table = "T34"
    w_TableName = "T34_RISYU_TOKU"
    w_KamokuName = "T34_TOKUKATU_CD"
  end if


  '//�������ʂ̒l���ꗗ��\��
  w_sSQL = ""
  w_sSQL = w_sSQL & " SELECT "

  Select Case m_sSikenKBN
    Case C_SIKEN_ZEN_TYU

      w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_TYUKAN_Z) as SOUJI,"
      w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_TYUKAN_Z) as JYUNJI, "


    Case C_SIKEN_ZEN_KIM

      w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_KIMATU_Z) as SOUJI, "
      w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_KIMATU_Z) as JYUNJI, "

    Case C_SIKEN_KOU_TYU

      w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_TYUKAN_K) as SOUJI, "
      w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_TYUKAN_K) as JYUNJI, "

    Case C_SIKEN_KOU_KIM


      w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_KIMATU_K) as SOUJI, "
      w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_KIMATU_K) as JYUNJI, "

  End Select

 '��������p�@INS 2005/06/13 ����
'�O���@����
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_TYUKAN_Z) as SOUJI1, "
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_TYUKAN_Z) as JYUNJI1, "
'�O���@����
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_KIMATU_Z) as SOUJI2, "
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_KIMATU_Z) as JYUNJI2, "
'����@����
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_TYUKAN_K) as SOUJI3, "
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_TYUKAN_K) as JYUNJI3, "
'����@����
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_SOJIKAN_KIMATU_K) as SOUJI4, "
  w_sSQL = w_sSQL & " MAX(" & w_Table & "_JUNJIKAN_KIMATU_K) as JYUNJI4 "

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

  else
    if m_sGakkoNO = C_NCT_KUMAMOTO Then
      w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
    else
      if Cint(m_iKamoku_Kbn) <> C_KAMOKU_SENMON Then
        w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
      else
		if (m_sGakkoNO = C_NCT_MAIZURU) Then
			'INS 2007/06/11
			'--2019/06/17 DELETE FUJIBAYASHI(80001�A80002�̓��ʏ�������߂�)
			'If (m_sKamokuCd = 80001) OR (m_sKamokuCd = 80002) Then
			'	w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
			'Else
			'--2019/06/17 DELETE END
			    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
			'End If		'--2019/06/17 DELETE FUJIBAYASHI(80001�A80002�̓��ʏ�������߂�)
			'INS END 2007/06/11
			'DEL 2007/06/11 w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
		else
			If m_sGakkoNO = C_NCT_OKINAWA Then
				If m_sKamokuCd >= 900000 Then
					w_sSQL = w_sSQL & " AND T13_CLASS = " & cint(m_sClassNo)
				Else
				    w_sSQL = w_sSQL & " AND T13_GAKKA_CD = '" & m_sGakkaCd & "' "
				End If
			Else

			End If
		end if
      end if
    end if
  end if
'****************************


  If gf_GetRecordset(m_RsMax,w_sSQL) <> 0 Then Exit function


  m_iSouJyugyou = gf_SetNull2String(m_RsMax("SOUJI"))
  m_iJunJyugyou = gf_SetNull2String(m_RsMax("JYUNJI"))

  '��������p�@INS 2005/06/13 ����	
  m_iSouJyugyou1 = gf_SetNull2String(m_RsMax("SOUJI1"))
  m_iJunJyugyou1 = gf_SetNull2String(m_RsMax("JYUNJI1"))
  m_iSouJyugyou2 = gf_SetNull2String(m_RsMax("SOUJI2"))
  m_iJunJyugyou2 = gf_SetNull2String(m_RsMax("JYUNJI2"))
  m_iSouJyugyou3 = gf_SetNull2String(m_RsMax("SOUJI3"))
  m_iJunJyugyou3 = gf_SetNull2String(m_RsMax("JYUNJI3"))
  m_iSouJyugyou4 = gf_SetNull2String(m_RsMax("SOUJI4"))
  m_iJunJyugyou4 = gf_SetNull2String(m_RsMax("JYUNJI4"))

   m_RsMax.close

f_GetJyugyoJIkan = true


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
  m_bSeiInpFlg = false

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
'   Response.Write "  EOF "
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

  '���͊��ԓ��Ȃ琳��
  If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
    m_bSeiInpFlg = true
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
'response.write "m_Date = " & m_Date & "<br>"

  p_IdouKbn = gf_Get_IdouChk(p_GakusekiNo,m_Date,m_iNendo,w_IdoutypeName)

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


  '2002/12/25
  '���㍂��̏ꍇ�A�O���J�݉Ȗڂ͎����ɂ���ăR�s�[���̎�����ς���
    If m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then

    '������Ԃ̎��A�O����������Z�b�g
    If m_sSikenKBN = C_SIKEN_KOU_TYU and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))  '�O������
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_TK"))  '�������

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))
      End If
    End If

    '��������̎��A������Ԃ���Z�b�g
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_TK"))  '�������
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))  '�������

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
        p_sSeiseki = gf_SetNull2String(m_Rs("SEI_KT"))
      End If
    End If

  '������i�O����->������ɃZ�b�g�j
  Else
    '�w�N�������̏ꍇ�̂�
    If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
      w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
      w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

      if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
      'If gf_SetNull2String(m_Rs("SEI")) = "" Then
        p_sSeiseki = gf_SetNull2String(m_Rs("SEI_ZK"))
      End If
    End If

  End If



  '//�ʏ���Ƃ̂Ƃ�
  if m_iKamokuKbn = C_JIK_JUGYO then

    if m_HyokaDispFlg then
      p_sHyoka = gf_SetNull2String(m_Rs("HYOKAYOTEI"))
      if p_sHyoka = "" then p_sHyoka = "�E"
    end if

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
'*  [�@�\]  �]���s�\����(�F�{�d�g�̂�)
'********************************************************************************
Sub s_SetHyoka(p_IdouKbn,p_DataKbn,p_Checked,p_Disabled2,p_Disabled)

  '//�]���s�\�f�[�^�ݒ�
  p_DataKbn = 0
  p_Checked = ""
  p_Disabled2 = ""

  p_DataKbn = cint(gf_SetNull2Zero(m_Rs("DataKbn")))

  If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
    w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
    w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))

    if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then

      p_DataKbn = cint(gf_SetNull2Zero(m_Rs("DataKbn_ZK")))

    end if
  end if

  if p_Disabled <> "" then p_Disabled2 = "disabled"

  if p_DataKbn = cint(C_HYOKA_FUNO) then
    p_Checked = "checked"
    p_Disabled2 = "disabled"

  elseif p_DataKbn = cint(C_MIHYOKA) then
    p_Disabled2 = "disabled"

  end if

  if not m_bSeiInpFlg Then p_Disabled2 = ""

  select case Cstr(p_IdouKbn)
    case Cstr(C_IDO_KYU_BYOKI),Cstr(C_IDO_KYU_HOKA)
      p_DataKbn = C_KYUGAKU

    case Cstr(C_IDO_TAI_2NEN),Cstr(C_IDO_TAI_HOKA),Cstr(C_IDO_TAI_SYURYO)
      p_DataKbn = C_TAIGAKU
  end select

End Sub

'********************************************************************************
'*  [�@�\]  �e�[�u���T�C�Y�̃Z�b�g
'********************************************************************************
Sub s_SetTableWidth(p_TableWidth)

  p_TableWidth = 610

  '//�]���s�\����������(�F�{�d�g�̂�)
  if m_SchoolFlg then
    p_TableWidth = 660
  end if

  '//�]���\��\���t���O�I���A�܂��́A�ʏ���Ƃ̂Ƃ�
  if m_HyokaDispFlg and Cstr(m_iKamokuKbn) = Cstr(C_TUKU_FLG_TUJO) then
    p_TableWidth = p_TableWidth + 50
  end if

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

  Dim w_Padding
  Dim w_Padding2

  Dim w_Disabled
  Dim w_Disabled2
  Dim w_TableWidth

  '//�x���A����(�Ώ�)�A����(�ΏۊO)�̍��v
  Dim wChikokuSum,wKekkaSum,wKekkaGaiSum

  wChikokuSum = 0
  wKekkaSum = 0
  wKekkaGaiSum = 0

  w_Padding = "style='padding:2px 0px;'"
  w_Padding2 = "style='padding:2px 0px;font-size:10px;'"

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

  '//�e�[�u���T�C�Y�̃Z�b�g
  Call s_SetTableWidth(w_TableWidth)

  if m_SchoolFlg then
    if m_MiHyokaFlg or (not m_bSeiInpFlg) then
      w_Disabled = "disabled"
    end if
  end if
%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--
  //************************************************************
  //  [�@�\]  �y�[�W���[�h������
  //************************************************************
  function window_onload() {
    //�X�N���[����������
    parent.init();
    //���l���͂̂Ƃ��̂�
    <% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
      //���э��v�l�̎擾
      f_GetTotalAvg();
    <% end if %>

    //�����ԂƏ����Ԃ�hidden�ɃZ�b�g
    document.frm.hidSouJyugyou.value = "<%= m_iSouJyugyou %>";
    document.frm.hidJunJyugyou.value = "<%= m_iJunJyugyou %>";


  // INS 2005/06/13 �����@��������p 
  document.frm.hidSouJyugyou_TZ.value = "<%= m_iSouJyugyou1 %>";
  document.frm.hidJunJyugyou_TZ.value = "<%= m_iJunJyugyou1 %>";
  document.frm.hidSouJyugyou_KZ.value = "<%= m_iSouJyugyou2 %>";
  document.frm.hidJunJyugyou_KZ.value = "<%= m_iJunJyugyou2 %>";
  document.frm.hidSouJyugyou_TK.value = "<%= m_iSouJyugyou3 %>";
  document.frm.hidJunJyugyou_TK.value = "<%= m_iJunJyugyou3 %>";
  document.frm.hidSouJyugyou_KK.value = "<%= m_iSouJyugyou4 %>";
  document.frm.hidJunJyugyou_KK.value = "<%= m_iJunJyugyou4 %>";

    document.frm.target = "topFrame";
    document.frm.action = "sei0150_middle.asp";
    document.frm.submit();
  }

  //************************************************************
  //  [�@�\]  �]���{�^���������ꂽ�Ƃ�
  //************************************************************
  function f_change(p_iS){
    w_sButton = eval("document.frm.button"+p_iS);
    w_sHyouka = eval("document.frm.Hyoka"+p_iS);

    <%If m_sSikenKBN = C_SIKEN_ZEN_TYU Then%>
      if(w_sButton.value == "�E") {
        w_sButton.value = "��";
        w_sHyouka.value = "��";
        return true;
      }
      if(w_sButton.value == "��") {
        w_sButton.value = "�E";
        w_sHyouka.value = "";
        return true;
      }

    <%Else%>

      if(w_sButton.value == "�E") {
        w_sButton.value = "��";
        w_sHyouka.value = "��";
        return true;
      }
      if(w_sButton.value == "��") {
        w_sButton.value = "��";
        w_sHyouka.value = "��";
        return true;
      }
      if(w_sButton.value == "��") {
        w_sButton.value = "�E";
        w_sHyouka.value = "";
        return true;
      }
    <%End If%>
  }

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //************************************************************
    function f_Touroku(){
    if(!f_InpCheck()){
      alert("���͒l���s���ł�");
      return false;
    }

    //�����Ԃ̂ݓ��̓`�F�b�N��ǉ� 2003.08.04 ITO
    if(parent.topFrame.document.frm.txtJunJyugyou.value == ""){
      parent.topFrame.document.frm.txtJunJyugyou.focus();
      alert("�����Ǝ��Ԑ������͂���Ă��܂���");
      return false;
    }

    // �v���č���̏ꍇ
    <% If (m_sGakkoNO = C_NCT_KURUME) AND (m_bSeiInpFlg) AND (m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM) AND (m_bKekkaNyuryokuFlg) Then %>
      if(!jf_CheckInpVal()){
        alert("��u���������͂���Ă��܂���");
        return false;
      }

	  if(parent.topFrame.document.frm.txtSouJyugyou.value < (<%=m_lHaitoTani%> * 30)){
			    parent.topFrame.document.frm.txtSouJyugyou.focus();
		        alert("�����Ǝ��Ԑ�������܂���B�P�ʐ��~30���Ԉȏ�œ��͂��Ă��������B");
		        return false;
	  }

	  if(parent.topFrame.document.frm.txtJunJyugyou.value < (<%=m_lHaitoTani%> * 24)){
	    parent.topFrame.document.frm.txtJunJyugyou.focus();
        alert("�����Ǝ��Ԑ�������܂���B�P�ʐ��~24���Ԉȏ�œ��͂��Ă��������B");
        return false;
	  }
    <% End If %>

    if(!confirm("<%=C_TOUROKU_KAKUNIN%>")) { return false;}
    document.frm.hidSouJyugyou.value = parent.topFrame.document.frm.txtSouJyugyou.value;
    document.frm.hidJunJyugyou.value = parent.topFrame.document.frm.txtJunJyugyou.value;

    //�w�b�_���󔒕\��
    parent.topFrame.document.location.href="white.asp";

    //�o�^����
    <% if m_iKamokuKbn = C_JIK_JUGYO then %>
      document.frm.hidUpdMode.value = "TUJO";
      document.frm.action="sei0150_upd.asp";
    <% Else %>
      document.frm.hidUpdMode.value = "TOKU";
      document.frm.action="sei0150_upd_toku.asp";
    <% End if %>
    document.frm.target="main";
    document.frm.submit();
  }

  //************************************************************
  //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
  //************************************************************
  function f_Cancel(){
    parent.document.location.href="default.asp";
  }

  //************************************************************
  //  [�@�\]  ���т̍��v�ƕ��ς����߂�
  //  [����]  �Ȃ�
  //  [�ߒl]  �Ȃ�
  //  [����]  ���ѓ��͊��ԊO�A���ԓ��ɂ���Čv�Z�̎d����ς���
  //  [���l]
  //************************************************************
  function f_GetTotalAvg(){
    var i;
    var total;
    var avg;
    var cnt;

    total = 0;
    cnt = 0;
    avg = 0;

    <% If m_bSeiInpFlg Then %>
      //�w�����ł̃��[�v
      for(i=0;i<<%=m_rCnt%>;i++) {
        //���݂��邩�ǂ���
        textbox = eval("document.frm.Seiseki" + (i+1));
        if(textbox){
          //�����̓`�F�b�N
          if (textbox.value != "") {
            //�����łȂ��͖̂�������
            if(!isNaN(textbox.value)){
              total = total + parseInt(textbox.value);
              cnt = cnt + 1;
            }
          }
        }
      }

    <% Else %>
      total = document.frm.hidTotal.value;
      cnt   = document.frm.hidGakTotal.value;
    <% End If%>

    document.frm.txtTotal.value = total;

    //�l�̌ܓ�
    if (cnt!=0){
      avg = total/cnt;
      avg = avg * 10;
      avg = Math.round(avg);
      avg = avg / 10;
    }

    document.frm.txtAvg.value=avg;
  }

    //************************************************************
    //  [�@�\]  ���l�^�`�F�b�N
    //************************************************************
  function f_CheckNum(pFromName){
    var wFromName,w_len;

    wFromName = eval(pFromName);

    if(isNaN(wFromName.value)){
      wFromName.focus();
      wFromName.select();
      return false;
    }else{
      //���`�F�b�N
      if(wFromName.name.indexOf("Seiseki") != -1){
        if(wFromName.value > 100){
          wFromName.focus();
          wFromName.select();
          return false;
        }
      }

      //�x���́A2���܂�
      if(wFromName.name.indexOf("Chikai") != -1){
        w_len = 2;
      }else{
        w_len = 3;
      }

      if(wFromName.value.length > w_len){
        wFromName.focus();
        wFromName.select();
        return false;
      }

      //�}�C�i�X���`�F�b�N
      var wStr = new String(wFromName.value)
      if (wStr.match("-")!=null){
        wFromName.focus();
        wFromName.select();
        return false;
      }

      if(wFromName.name.indexOf("txtAvg") == -1){
        //�����_�`�F�b�N
        w_decimal = new Array();
        w_decimal = wStr.split(".")

        if(w_decimal.length>1){
          wFromName.focus();
          wFromName.select();
          return false;
        }
      }
    }

    return true;
  }

    //************************************************************
    //  [�@�\]  �v���č��ꎞ�̐��тƎ�u�����̓��̓`�F�b�N
  //  [�ڍ�]�@���� & ��u���ԋ��ɓ��͉̎�
  //  �@�@�@�@���т����͂���Ă���ꍇ�͎�u���Ԃ�K�{���͂Ƃ���
  //  [�쐬]  2003.05.13 hirota
    //************************************************************
  function jf_CheckInpVal(){
    //�w�����ł̃��[�v
    for(i=0;i<<%=m_rCnt%>;i++) {
      //���݂��邩�ǂ���
      textbox1 = eval("document.frm.Seiseki" + (i+1));  // ����
      textbox2 = eval("document.frm.Kekka" + (i+1));    // ����
      if(textbox2){
        //�����̓`�F�b�N
        if (textbox2.value == "") {
          if(textbox1){
            if(textbox1.value != ""){
              textbox2.focus();
              return false;
            }
          }
        }
      }
    }
    return true;
  }

    //************************************************************
    //  [�@�\]  �召�`�F�b�N
    //************************************************************
  function f_CheckDaisyou(){
    wObj1 = eval("parent.topFrame.document.frm.txtSouJyugyou");
    wObj2 = eval("parent.topFrame.document.frm.txtJunJyugyou");

    if(wObj1.value != "" && wObj2.value != ""){
      if(wObj1.value < wObj2.value){
        wObj1.focus();
        return false;
      }
    }
    return true;
  }

  //************************************************
  //Enter �L�[�ŉ��̓��̓t�H�[���ɓ����悤�ɂȂ�
  //�����Fp_inpNm �Ώۓ��̓t�H�[����
  //    �Fp_frm �Ώۃt�H�[��
  //�@�@�Fi   ���݂̔ԍ�
  //�ߒl�F�Ȃ�
  //���̓t�H�[�������Axxxx1,xxxx2,xxxx3,�c,xxxxn
  //�̖��O�̂Ƃ��ɗ��p�ł��܂��B
  //************************************************
  function f_MoveCur(p_inpNm,p_frm,i){
    if (event.keyCode == 13){   //�����ꂽ�L�[��Enter(13)�̎��ɓ����B
      i++;

      //���͉\�̃e�L�X�g�{�b�N�X��T���B����������t�H�[�J�X���ڂ��ď����𔲂���B
          for (w_li = 1; w_li <= 99; w_li++) {

        if (i > <%=m_rCnt%>) i = 1; //i���ő�l�𒴂���ƁA�͂��߂ɖ߂�B
        inpForm = eval("p_frm."+p_inpNm+i);

        //���͉\�̈�Ȃ�t�H�[�J�X���ڂ��B
        if (typeof(inpForm) != "undefined") {
          inpForm.focus();      //�t�H�[�J�X���ڂ��B
          inpForm.select();     //�ڂ����e�L�X�g�{�b�N�X����I����Ԃɂ���B
          break;
        //���͕t���Ȃ玟�̍��ڂ�
        } else{
          i++
        }
          }
    }else{
      return false;
    }
    return true;
  }

  //************************************************
  //  �������͎��̐��я���
  //
  //************************************************
  function f_SetSeiseki(w_num){
    var ob = new Array();

    ob[0] = eval("parent.topFrame.document.frm.sltHyoka");
    ob[1] = eval("document.frm.Seiseki" + w_num);
    ob[2] = eval("document.frm.hidSeiseki" + w_num);
    ob[3] = eval("document.frm.hidHyokaFukaKbn" + w_num);

    if(ob[0].value.length == 0){
      ob[1].value = "";
      ob[2].value = "";
      ob[3].value = 0;
    }else{
      var vl = ob[0].value.split('#@#');

      ob[1].value = vl[0];
      ob[2].value = vl[0];
      ob[3].value = vl[1];
    }
  }

  //************************************************
  //  ���̓`�F�b�N
  //************************************************
  function f_InpCheck(){
    var w_length;
    var ob;

    //�����ԁE�����ԓ��̓`�F�b�N
    if(!f_CheckNum("parent.topFrame.document.frm.txtSouJyugyou")){ return false; }
    if(!f_CheckNum("parent.topFrame.document.frm.txtJunJyugyou")){ return false; }
    // �召�`�F�b�N�s�v 2003.02.20
    // if(!f_CheckDaisyou()){ return false; }

    w_length = document.frm.elements.length;

    for(i=0;i<w_length;i++){
      ob = eval("document.frm.elements[" + i + "]")

      //if(ob.type=="text" && ob.name != "txtAvg"  && ob.name != "txtTotal"){
      if(ob.type=="text" && ob.name != "txtAvg"  && ob.name != "txtTotal"  && ob.name != "ChikaiSum"  && ob.name != "KekkaSum"){
        ob = eval("document.frm." + ob.name);

        if(!f_CheckNum(ob)){return false;}
      }
    }
    return true;
  }

  //************************************************
  //�]���s�\���N���b�N���ꂽ�Ƃ��̏���
  //************************************************
  function f_InpDisabled(p_num){

    <% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
      var ob = new Array();

      ob[0] = eval("document.frm.chkHyokaFuno" + p_num);
      ob[1] = eval("document.frm.Seiseki" + p_num);

      if(ob[0].checked){
        ob[1].value = "";
        ob[1].disabled = true;

      }else{
        ob[1].disabled = false;
      }
    <% end if %>

    //���l���͂̂Ƃ��̂�
    <% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
      f_GetTotalAvg();
    <% end if %>
  }

  //************************************************************
  //  [�@�\]  �x���A���ہA���ۊO�̍��v�̌v�Z
  //************************************************************
  function f_CalcSum(p_Name){
    var w_num;
    var total = 0;
    var cnt = 0;

    var ob_kei = eval("document.frm." + p_Name + "Sum")

    //�w�����ł̃��[�v
    for(w_num=0;w_num<<%=m_rCnt%>;w_num++) {
      //���݂��邩�ǂ���
      textbox = eval("document.frm." + p_Name + (w_num+1));
      if(textbox){
        //�����̓`�F�b�N
        if (textbox.value != "") {
          //�����łȂ��͖̂�������
          if(!isNaN(textbox.value)){
            total = total + parseInt(textbox.value);
          }
        }
        cnt = cnt + 1;
      }
    }

    ob_kei.value = total;
  }

  //-->
  </SCRIPT>
  </head>
  <body LANGUAGE="javascript" onload="window_onload();">
  <center>
  <form name="frm" method="post">

  <table width="<%=w_TableWidth%>">
  <tr>
  <td>

  <table class="hyo" align="center" width="<%=w_TableWidth%>" border="1">
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

      Call gs_cellPtn(w_cell)

      '�X�^�C���V�[�g�ݒ�
      if not m_bSeiInpFlg Then
        w_sInputClass1 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
        w_Disabled = "disabled"
      End if

      if Not m_bKekkaNyuryokuFlg Then
        w_sInputClass2 = "class='" & w_cell & "' style='text-align:right;' readonly tabindex='-1'"
      End if

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

      '//�]���s�\����(�F�{�d�g�̂�)
      if m_SchoolFlg  then
        Call s_SetHyoka(w_IdouKbn,w_DataKbn,w_Checked,w_Disabled2,w_Disabled)
      end if

      %>

      <tr>
        <td class="<%=w_cell%>" align="center" width="65"  nowrap <%=w_Padding%>><%=m_Rs("GAKUSEKI_NO")%></td>
        <input type="hidden" name="txtGseiNo<%=i%>"   value="<%=m_Rs("GAKUSEI_NO")%>">
        <input type="hidden" name="hidNoChange<%=i%>" value="<%=w_bNoChange%>">

      	<% If m_sGakkoNO = cstr(C_NCT_KURUME) then %>

      		<!-- �Ə��t���O���P�̏ꍇ  -->
      		<% If cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG"))) = "1" Then %>

		        <td class="<%=w_cell%>" align="left" width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%>[�C����]</td>

			<% Else %>

		        <td class="<%=w_cell%>" align="left" width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>

			<% End If %>

		<% Else %>

		    <td class="<%=w_cell%>" align="left" width="150" nowrap <%=w_Padding%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>

		<% End If %>

        <% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>

          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI1")))%></td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI2")))%></td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI3")))%></td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>><%=gf_IIF(w_bNoChange,"-",gf_HTMLTableSTR(m_Rs("SEI4")))%></td>
        <% else %>

          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
          <td class="<%=w_cell%>" align="center" width="30"  nowrap <%=w_Padding2%>>-</td>
        <% end if %>

        <!--�I���Ȗڂ̎��ɖ��I���̏ꍇ�A���͕s�B�܂��A�x�w�Ȃ�-->
        <% If w_bNoChange = True Then %>

          <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>

          <% if m_HyokaDispFlg and m_iKamokuKbn = C_JIK_JUGYO then %>

            <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
          <% end if %>


          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>

          <% if m_KekkaGaiDispFlg then %>

            <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>
          <% end if %>

          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>-</td>

          <% if m_SchoolFlg then %>

            <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>
            <input type="hidden" name="chkHyokaFuno<%=i%>" value="<%=w_DataKbn%>">
          <% end if %>

        <% Else %>

          <!-- ���� (���l���́A�������́A���тȂ����͂ɂ�菈���𕪂���) -->
          <!--
          20040218 �C�� shiki
          ���Â̏ꍇ�A�Ə��t���O���P�̏ꍇ�́A���т̃e�L�X�gBOX��b��t���ĕ\���̂�
          ���]���`�F�b�N�{�b�N�X���N���b�N����Ă��������Ȃ��悤��
          hidMenjiFlg = 1 �𗧂Ă�
          -->
		<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>

			<!-- ����  -->
          	<% If m_sGakkoNO = cstr(C_NCT_NUMAZU) then %>

          		<!-- �Ə��t���O���P�̏ꍇ  -->
          		<% If cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG"))) = "1" Then %>
					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><font size="2">b</font>&nbsp;<input type="text" <%=w_sInputClass1%>  name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=3 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);" onChange="f_GetTotalAvg();" style="border-color:#FFFFFF #FFFFFF #FFFFFF #FFFFFF; border-style: solid; text-align:left; vertical-align:middle; " readonly></td>

					<!-- ���]���`�F�b�N�{�b�N�X���N���b�N���ꂽ���p�Ƀt���O�𗧂Ă� -->
					<input type="hidden" name="hidMenjiFlg<%=i%>" value="1">

				<% Else %>
					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass1%>  name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=3 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);" onChange="f_GetTotalAvg();"></td>

				<% End If %>

			<!-- ���ÈȊO  -->
			<% Else %>

					<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass1%>  name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=3 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>);" onChange="f_GetTotalAvg();"></td>

			<% End If %>

			<!-- END -->

		<% elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>

				<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>

				<% if not m_bSeiInpFlg Then %>
					<%=w_sSeiseki%>
				<% else %>
					<input type="button" class="<%=w_cell%>" style="text-align:center;" name="Seiseki<%=i%>" value="<%=w_sSeiseki%>" size=2 onClick="f_SetSeiseki(<%=i%>);" <%=w_Disabled%>>
				<% end if %>

				</td>
				<input type="hidden" name="hidSeiseki<%=i%>" value="<%=w_sSeiseki%>">
				<input type="hidden" name="hidHyokaFukaKbn<%=i%>" value="<%=m_Rs("HYOKA_FUKA")%>">

		<% else %>

			<td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>-</td>

		<% end if %>

          <!-- �]���\�� -->
          <% If m_HyokaDispFlg and m_iKamokuKbn = C_JIK_JUGYO then %>
            <% if m_bSeiInpFlg and (m_sSikenKBN = C_SIKEN_ZEN_TYU or m_sSikenKBN = C_SIKEN_KOU_TYU) then %>
              <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>>
                <input type="button" name="button<%=i%>" value="<%=w_sHyoka%>" size="2" onClick="f_change(<%=i%>);" class="<%=w_cell%>" style="text-align:center">
                <input type="hidden" name="Hyoka<%=i%>"  value="<%=trim(w_sHyoka)%>">
              </td>
            <% else %>
              <td class="<%=w_cell%>" align="center" width="50" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sHyoka)%></td>
            <% end if %>
          <% end if %>

		<!-- �x�� -->
          <td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Chikai<%=i%> value="<%=w_sChikai%>" size=2 maxlength=2 onKeyDown="f_MoveCur('Chikai',this.form,<%=i%>)" onChange="f_CalcSum('Chikai');" readonly = true></td>
		<!-- ���� -->
          <td class="<%=w_cell%>" align="right"  width="55" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sChikaisu)%></td>
          <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=Kekka<%=i%> value="<%=w_sKekka%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kekka',this.form,<%=i%>)" onChange="f_CalcSum('Kekka');"></td>

          <% if m_KekkaGaiDispFlg then %>
            <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>><input type="text" <%=w_sInputClass2%>  name=KekkaGai<%=i%> value="<%=w_sKekkaGai%>" size=2 maxlength=3 onKeyDown="f_MoveCur('KekkaGai',this.form,<%=i%>)" onChange="f_CalcSum('KekkaGai');"></td>
          <% end if %>

          <td class="<%=w_cell%>" align="right"  width="55" nowrap <%=w_Padding%>><%=gf_HTMLTableSTR(w_sKekkasu)%></td>

		<!--���x ���Â̂�-->
		<%if m_sGakkoNO = cstr(C_NCT_NUMAZU) THEN %>
         <td class="<%=w_cell%>" align="center" width="55" nowrap <%=w_Padding%>>
			<input type="text" <%=w_sInputClass2%>  name=Kokyu<%=i%> value="<%=m_Rs("KEKA_NASI")%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Kokyu',this.form,<%=i%>);">
         </td>
		<% end if %>

		<!-- �]���s�\���� -->
		<!--
          20040218 �C�� shiki
          ���Â̏ꍇ�A�Ə��t���O���P�̏ꍇ�́A�]���s�\�̃`�F�b�NBOX��DISABLED
		-->
		<% if m_SchoolFlg then %>
			<td class="<%=w_cell%>" width="50" align="center" nowrap <%=w_Padding%>>
			<% if w_DataKbn = C_HYOKA_FUNO or w_DataKbn = C_MIHYOKA or w_DataKbn = 0 then %>

				<!-- ����  -->
				<% If m_sGakkoNO = cstr(C_NCT_NUMAZU) then %>
					<!-- �Ə��t���O���P�̏ꍇ  -->
					<% If cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG"))) = "1" Then %>
						<input type="checkbox" name="chkHyokaFuno<%=i%>" <%=w_Disabled%> value="3"  <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);" disabled>
					<% Else %>
						<input type="checkbox" name="chkHyokaFuno<%=i%>" <%=w_Disabled%> value="3"  <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);">
					<% End If %>
				<% Else %>

					<input type="checkbox" name="chkHyokaFuno<%=i%>" <%=w_Disabled%> value="3"  <%=w_Checked%> onClick="f_InpDisabled(<%=i%>);">

				<% End If %>

			<% else %>
				<input type="hidden" name="chkHyokaFuno<%=i%>" value="<%=w_DataKbn%>">
			<% end if %>

			</td>
 		<% end if %>
		<!-- END -->

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

		<%
		'2003.8.25 ITO
		'2004.02.18 ���Â��Ə��Ȗڂ���������悤�ɂȂ����ׁA���L����ɏ��Â��܂߂�
		'�v���Ă̏ꍇ�A�Ə��Ȗڂ͑O�������т��w�N���ɃR�s�[���Ȃ��ׁA�Ə��t���O��ݒ�
		If m_sGakkoNO = cstr(C_NCT_KURUME) OR m_sGakkoNO = cstr(C_NCT_NUMAZU) then
		'If m_sGakkoNO = cstr(C_NCT_KURUME) then
		%>
			<input type="hidden" name="hidMenjyo<%=i%>" value="<%=cint(gf_SetNull2Zero(m_Rs("T16_MENJYO_FLG")))%>">
		<%

		'���̑��̊w�Z�́A�S�Ēʏ�̉Ȗڂŏ�������ׂ�0��ݒ�
		Else
		%>
			<input type="hidden" name="hidMenjyo<%=i%>" value="0">
		<%
		End If
		%>

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

      <% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
        <tr>
          <td class="header" align="right" colspan="7" nowrap>
            <FONT COLOR="#FFFFFF"><B>���v</B></FONT>
            <input type="text" name="txtTotal" size="5" <%=w_sInputClass%> readonly>
          </td>

          <td class="header" align="center" nowrap><input type="text" name="ChikaiSum" size="4" <%=w_sInputClass%> readonly value="<%=wChikokuSum%>"></td>
          <td class="header" align="center" nowrap>&nbsp;</td>
          <td class="header" align="center" nowrap><input type="text" name="KekkaSum" size="4" <%=w_sInputClass%> readonly value="<%=wKekkaSum%>"></td>

          <% if m_KekkaGaiDispFlg then %>
            <td class="header" align="center" nowrap><input type="text" name="KekkaGaiSum" size="4" <%=w_sInputClass%> readonly value="<%=wKekkaGaiSum%>"></td>
          <% end if %>

          <td class="header" align="center" colspan="2" nowrap>&nbsp;</td>
        </tr>

        <tr>
          <td class="header" align="right" colspan="7" nowrap>
            <FONT COLOR="#FFFFFF"><B>�@���ϓ_</B></FONT>
            <input type="text" name="txtAvg" size="5" <%=w_sInputClass%> readonly>
          </td>
          <td class="header" align="center" colspan="6" nowrap>&nbsp;</td>
        </tr>
      <% else %>
        <tr>
          <td class="header" align="center" colspan="7" nowrap><FONT COLOR="#FFFFFF"><B>���v</B></FONT></td>
          <td class="header" align="center" nowrap><input type="text" name="ChikaiSum" size="4" <%=w_sInputClass%> readonly value="<%=wChikokuSum%>"></td>
          <td class="header" align="center" nowrap>&nbsp;</td>
          <td class="header" align="center" nowrap><input type="text" name="KekkaSum" size="4" <%=w_sInputClass%> readonly value="<%=wKekkaSum%>"></td>

          <% if m_KekkaGaiDispFlg then %>
            <td class="header" align="center" nowrap><input type="text" name="KekkaGaiSum" size="4" <%=w_sInputClass%> readonly value="<%=wKekkaGaiSum%>"></td>
          <% end if %>

          <td class="header" align="center" colspan="2" nowrap>&nbsp;</td>
        </tr>
      <% end if %>

    </table>

    </td>
    </tr>

    <tr>
    <td align="center">
    <table>
      <tr>
        <td align="center" align="center" colspan="13">
          <%If m_bSeiInpFlg or m_bKekkaNyuryokuFlg Then%>
            <input type="button" class="button" value="�@�o�@�^�@" onClick="f_Touroku();">
          <%End If%>
            <input type="button" class="button" value="�L�����Z��" onClick="f_Cancel();">
        </td>
      </tr>
    </table>
    </td>
    </tr>
  </table>

  <input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
  <input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
  <input type="hidden" name="KamokuCd"     value="<%=m_sKamokuCd%>">
  <input type="hidden" name="i_Max"        value="<%=i%>">
  <input type="hidden" name="sltShikenKbn" value="<%=m_sSikenKBN%>">
  <input type="hidden" name="txtGakuNo"    value="<%=m_sGakuNo%>">
  <input type="hidden" name="txtGakkaCd"   value="<%=m_sGakkaCd%>">
  <input type="hidden" name="txtClassNo"   value="<%=m_sClassNo%>">
  <input type="hidden" name="txtKamokuCd"  value="<%=m_sKamokuCd%>">
  <input type="hidden" name="PasteType"    value="">

  <input type="hidden" name="hidSouJyugyou">
  <input type="hidden" name="hidJunJyugyou">
  <input type="hidden" name="hidUpdMode">

 <!-- INS 2005/06/13 �����@��������p -->
  <input type="hidden" name="hidSouJyugyou_TZ">
  <input type="hidden" name="hidJunJyugyou_TZ">
  <input type="hidden" name="hidSouJyugyou_KZ">
  <input type="hidden" name="hidJunJyugyou_KZ">
  <input type="hidden" name="hidSouJyugyou_TK">
  <input type="hidden" name="hidJunJyugyou_TK">
  <input type="hidden" name="hidSouJyugyou_KK">
  <input type="hidden" name="hidJunJyugyou_KK">



  <input type="hidden" name="hidKamokuKbn" value="<%=m_iKamokuKbn%>">
  <input type="hidden" name="hidKamokuBunrui" value="<%=m_sKamokuBunrui%>">
  <input type="hidden" name="hidSeisekiInpType" value="<%=m_iSeisekiInpType%>">

  <input type="hidden" name="hidKikan" value="<%=m_bSeiInpFlg%>">
  <input type="hidden" name="hidKekkaNyuryokuFlg" value="<%=m_bKekkaNyuryokuFlg%>">

  <input type="hidden" name="hidTotal" value="<%=w_lSeiTotal%>">
  <input type="hidden" name="hidGakTotal" value="<%=w_lGakTotal%>">
  <input type="hidden" name="txtUpdDate" value="<%=request("txtUpdDate")%>">

  <input type="hidden" name="hidZenkiOnly" value="<%=m_bZenkiOnly%>">

  <!--<input type="text" name="hidMihyoka" value ="<%=w_DataKbn%>">-->
  <input type="hidden" name="hidMihyoka" value ="<%=w_DataKbn%>">
  <input type="hidden" name="hidSchoolFlg" value ="<%=m_SchoolFlg%>">
  <input type="hidden" name="hidKekkaGaiDispFlg" value ="<%=m_KekkaGaiDispFlg%>">
  <input type="hidden" name="hidHyokaDispFlg" value ="<%=m_HyokaDispFlg%>">

  <input type="hidden" name="hidTableWidth" value ="<%=w_TableWidth%>">


  <input type="hidden" name="hidFromSei"   value ="<%=m_iNKaishi%>">
  <input type="hidden" name="hidToSei"     value ="<%=m_iNSyuryo%>">
  <input type="hidden" name="hidFromKekka" value ="<%=m_iKekkaKaishi%>">
  <input type="hidden" name="hidToKekka"   value ="<%=m_iKekkaSyuryo%>">

<%'2003.8.24 �w�Z�ԍ���sei0150_upd.asp�ɓn���ׂɒǉ��B ITO%>
  <input type="hidden" name="hidGakkoNo"   value ="<%=m_sGakkoNO%>">


  </form>
  </center>
  </body>
  </html>
<%
End sub
%>