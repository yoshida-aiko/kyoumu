<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����������ꗗ
' ��۸���ID : web/web0360/web0360_main.asp
' �@      �\: ������\��
'-------------------------------------------------------------------------
' ��      ��:   txtClubCd		:����CD
'               KYOKAN_CD       '//����CD
'
' ��      �n:	txtMode			:�������[�h
'               GAKUSEI_NO		:�w��NO
'
' ��      ��:
'           �������\��
'               �󔒃y�[�W��\��
'           ���\���{�^���������ꂽ�ꍇ
'               �E�I�����ꂽ�������������ꗗ�\������
'               �E���O�C���҂��ږ�̏ꍇ�́A�o�^�A�폜���\�ƂȂ�
'               �E�ږ�ȊO�̎��́A�Q�Ƃ݂̂Ƃ���
'-------------------------------------------------------------------------
' ��      ��: 2001/08/22 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	Public m_iSyoriNen			'//�N�x
	Public m_iKyokanCd			'//��������
	Public m_sClubCd			'//�N���uCD

    Public m_bKomon				'//�ږ₩�ǂ����𔻕ʂ����׸�
    Public m_bUpdate_OK			'//�X�V�����׸�
    Public m_sKomonKyokanStr	'//�ږ⋳��CD

    'ں��ރZ�b�g
    Public m_Rs					'//�����ꗗں��޾�āi�����ҁj
    Public m_Rs2				'//�����ꗗں��޾�āi�ޕ��ҁj
    Public m_iRsCnt				'//ں��ރJ�E���g
    Public m_bGetMember			'//�����擾�׸�

	Dim	gTaibuFlg				'// �ޕ����׸�
	Dim	gTaibubi				'// �ޕ���
	Dim gFieldName				'// 

    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

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

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�����������ꗗ"
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
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '//�l�̏�����
        Call s_ClearParam()

        '//�ϐ��Z�b�g
        Call s_SetParam()

'//�f�o�b�O
'Call s_DebugPrint()

		'//�������A�ږ⋳�������擾
		w_iRet = f_GetClubInfo()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//�����̎擾
        w_iRet = f_GetMember()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'//�ږ⋳���ȊO��USER�͎Q�Ƃ݂̂Ƃ���
        Call s_SetViewInfo()

        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop



    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Rs)

    '// �I������
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen       = ""
    m_iKyokanCd       = ""
	m_sClubCd           = ""
	m_sKomonKyokanStr = ""
	m_bKomon          = False

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen  = Session("NENDO")
    m_iKyokanCd  = Session("KYOKAN_CD")
	m_sClubCd    = Request("txtClubCd")
	Session("HyoujiNendo") = m_iSyoriNen

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen  = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd  = " & m_iKyokanCd & "<br>"
    response.write "m_sClubCd    = " & m_sClubCd   & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  �����ȊO��USER�͎Q�Ƃ݂̂Ƃ���
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetViewInfo()

	m_bUpdate_OK = False

	'//�ږ�̋����͓o�^�E�폜���\
	If m_bKomon = True Then
		m_bUpdate_OK = True
	End If

End Sub

'********************************************************************************
'*  [�@�\]  �N���u���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �������A�ږ⋳���������擾
'********************************************************************************
Function f_GetClubInfo()

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubInfo = 1

	Do

		'//���������擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD1, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD2, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD3, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD4, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_KOMON_CD5, "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUJYOKYO_KBN"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD='" &  m_sClubCd & "'"

'response.write w_sSQL & "<br>"
'response.end
		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_GetClubInfo = 99
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//������
			m_sClubName = rs("M17_BUKATUDOMEI")

			'//�@���O�C���҂��ږ⋳�����ǂ����𔻒f���A
			'//�A�ږ⋳��CD���J���}��؂�ŕۑ�����
			For i = 1 To 5

				'//�@���O�C���҂��ږ⋳�����ǂ����𔻒f
				If trim(gf_SetNull2String(rs("M17_KOMON_CD" & i))) = trim(m_iKyokanCd) Then
					m_bKomon = True
				End If

				'//�A�ږ⋳��CD���J���}��؂�ŕۑ�����
				If trim(gf_SetNull2String(rs("M17_KOMON_CD" & i))) <> "" Then
					If m_sKomonKyokanStr = "" Then
						m_sKomonKyokanStr = rs("M17_KOMON_CD" & i)
					Else
						m_sKomonKyokanStr = m_sKomonKyokanStr & "," & rs("M17_KOMON_CD" & i)
					End If
				End If

			Next

		End If

		'//����I��
		f_GetClubInfo = 0
		Exit Do
	Loop

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �����ꗗ���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �����ҁA�ޕ��҂̏��Ԃɕ��ׂ邽�߂ɁA�����ăf�[�^���擾
'********************************************************************************
Function f_GetMember()

	Dim w_iRet
	Dim w_sSQL
	Dim rs
	Dim w_lCnt1
	Dim w_lCnt2

	On Error Resume Next
	Err.Clear

	f_GetMember = 1

	Do

		'//���������擾�i�����Ҏ擾�j
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_NYUBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_NYUBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI.T11_SIMEI"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO"
'		w_sSql = w_sSql & vbCrLf & "  AND T13_GAKU_NEN.T13_NENDO = T11_GAKUSEKI.T11_NYUNENDO + T13_GAKU_NEN.T13_GAKUNEN - 1"
		w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen

		w_sSql = w_sSql & vbCrLf & "  AND (  (T13_GAKU_NEN.T13_CLUB_1='" & m_sClubCd & "' AND T13_GAKU_NEN.T13_CLUB_1_FLG = 1)"
		w_sSql = w_sSql & vbCrLf & "      Or (T13_GAKU_NEN.T13_CLUB_2='" & m_sClubCd & "' AND T13_GAKU_NEN.T13_CLUB_2_FLG = 1)"
		w_sSql = w_sSql & vbCrLf & "      )"

		'w_sSql = w_sSql & vbCrLf & "  AND ((T13_GAKU_NEN.T13_CLUB_1=" & m_sClubCd & " AND T13_GAKU_NEN.T13_CLUB_1_FLG in (1,2))"
		'w_sSql = w_sSql & vbCrLf & "      Or (T13_GAKU_NEN.T13_CLUB_2=" & m_sClubCd & " AND T13_GAKU_NEN.T13_CLUB_2_FLG in (1,2))"
		'w_sSql = w_sSql & vbCrLf & "      )"
		w_sSql = w_sSql & vbCrLf & " ORDER BY "
		w_sSql = w_sSql & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write "111111111111<br>" & w_sSQL & "<br>"
		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If w_iRet <> 0 Then
'response.end
			'ں��޾�Ă̎擾���s
			f_GetMember = 99
			Exit Do
		End If

		'//���������擾�i�ޕ��Ҏ擾�j
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_NYUBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_1_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2 "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_NYUBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_TAIBI "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLUB_2_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI.T11_SIMEI"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN "
		w_sSql = w_sSql & vbCrLf & "  ,T11_GAKUSEKI "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO"
'		w_sSql = w_sSql & vbCrLf & "  AND T13_GAKU_NEN.T13_NENDO = T11_GAKUSEKI.T11_NYUNENDO + T13_GAKU_NEN.T13_GAKUNEN - 1"
		w_sSql = w_sSql & vbCrLf & "  AND  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen

		w_sSql = w_sSql & vbCrLf & "  AND ((T13_GAKU_NEN.T13_CLUB_1='" & m_sClubCd & "' AND T13_GAKU_NEN.T13_CLUB_1_FLG = 2)"
		w_sSql = w_sSql & vbCrLf & "      Or (T13_GAKU_NEN.T13_CLUB_2='" & m_sClubCd & "' AND T13_GAKU_NEN.T13_CLUB_2_FLG = 2)"
		w_sSql = w_sSql & vbCrLf & "      )"

		'w_sSql = w_sSql & vbCrLf & "  AND ((T13_GAKU_NEN.T13_CLUB_1=" & m_sClubCd & " AND T13_GAKU_NEN.T13_CLUB_1_FLG in (1,2))"
		'w_sSql = w_sSql & vbCrLf & "      Or (T13_GAKU_NEN.T13_CLUB_2=" & m_sClubCd & " AND T13_GAKU_NEN.T13_CLUB_2_FLG in (1,2))"
		'w_sSql = w_sSql & vbCrLf & "      )"
		w_sSql = w_sSql & vbCrLf & " ORDER BY "
		w_sSql = w_sSql & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASS "
		w_sSql = w_sSql & vbCrLf & "  ,T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write "222222222222<br>" & w_sSQL & "<br>"
		w_iRet = gf_GetRecordset(m_Rs2, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
'response.end
			f_GetMember = 99
			Exit Do
		End If

        m_iRsCnt = 0

		w_lCnt1 = 0
		w_lCnt2 = 0

		'�����Ґ�
        If m_Rs.EOF = False Then
            w_lCnt1 = gf_GetRsCount(m_Rs)
        End If

		'�ޕ��Ґ�
        If m_Rs2.EOF = False Then
            w_lCnt2 = gf_GetRsCount(m_Rs2)
        End If

		'//ں��ރJ�E���g�擾
        '//�������擾
        m_iRsCnt = w_lCnt1 + w_lCnt2
		
		If m_iRsCnt > 0 Then
			m_bGetMember = True
		Else
			m_bGetMember = False
		End If

        'If m_Rs.EOF = False Then
        '    m_iRsCnt = gf_GetRsCount(m_Rs)
		'	m_bGetMember = True
		'Else
		'	m_bGetMember = False
        'End If

		'//����I��
		f_GetMember = 0
		Exit Do
	Loop

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function


'********************************************************************************
'*  [�@�\]  �N���X�����擾
'*  [����]  p_iGakuNen:�w�N,p_iClassNo:�N���XNO
'*  [�ߒl]  f_GetClassName:�N���X����
'*  [����]  
'********************************************************************************
Function f_GetClassName(p_iGakuNen,p_iClassNo)
    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

	f_GetClassName = ""
	w_sClassName = ""

    Do
        '�N���X�}�X�^���f�[�^���擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_CLASSRYAKU"
        w_sSql = w_sSql & vbCrLf & "  ,M05_CLASS.M05_GAKKA_CD"
        w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN= " & p_iGakuNen
        w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_CLASSNO= "   & p_iClassNo

'response.write w_sSQL & "<br>"

		'//�f�[�^�擾
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            Exit Do
        End If

        If rs.EOF = False Then
            w_sClassName = rs("M05_CLASSRYAKU")
            'w_sGakkaCd = rs("M05_GAKKA_CD")
        End If

        Exit Do
    Loop

	'//�߂�l���
	f_GetClassName = w_sClassName

	'//ں���CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �ږ⋳����\��(HTML�����o��)
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ShowKomon()
	Dim i
	Dim w_Ary
	Dim w_sRowspan

	Do

		'//�ږ⋳�����ݒ肳��Ă��Ȃ��ꍇ
		If LenB(m_sKomonKyokanStr)=0 Then%>
			<table class="hyo" border="1">
				<tr>
					<th nowrap class="header" width="40"  align="center" >�ږ�</th>
					<td nowrap class="detail" width="120"  align="center">�\</td>
				</tr>
			</table>
			<%
		Else

			'//�ږ⋳��CD(CSV�`��)�擾
			w_Ary = split(m_sKomonKyokanStr,",")
			iMax = UBound(w_Ary)

			If iMax >= 3 Then
				w_sRowspan="rowspan=2"
			Else
				w_sRowspan="rowspan=1"
			End If

			'//�w�b�_�����o��
			%>
			<table class="hyo" border="1">
				<tr>
					<th nowrap class="header" width="50"  align="center" <%=w_sRowspan%> >�ږ�</th>
			<%For i = 0 To iMax%>

				<%If i = 3 Then%>
					</tr><tr>
				<%End If%>
					<td nowrap class="detail" width="120"  align="center"><%=gf_GetKyokanNm(m_iSyoriNen,w_Ary(i))%></td>
			<%Next

			'//���s����̏ꍇ�A�̂���̋󔒍s��\��
			If i-1 >= 3 Then
				For j = 1 To 6-i%>
					<td nowrap class="detail" width="100"  align="center"><br></td>
				<%
				Next
			End If

			%>
				</tr>
			</table>
		<%
		End If

		Exit Do
	Loop

End Sub


'********************************************************************************
'*  [�@�\]  �ޕ��҂��擾
'********************************************************************************
Sub s_Taibu(s_NyuTai_Flg)

	Dim wTaibuFlg

	'// �ޕ��҃t���O
	'�����҂̏ꍇ
	if s_NyuTai_Flg = 1 Then		
		wC1  = m_Rs("T13_CLUB_1")
		wC2  = m_Rs("T13_CLUB_2")
		wC1F = m_Rs("T13_CLUB_1_FLG")
		wC2F = m_Rs("T13_CLUB_2_FLG")
	'�ޕ��҂̏ꍇ
	else
		wC1  = m_Rs2("T13_CLUB_1")
		wC2  = m_Rs2("T13_CLUB_2")
		wC1F = m_Rs2("T13_CLUB_1_FLG")
		wC2F = m_Rs2("T13_CLUB_2_FLG")
	end if

	wTaibuFlg1 = False
	wTaibuFlg2 = False
	gTaibuFlg  = False
	gTaibubi   = ""

	if Not gf_IsNull(wC1) Then
'response.write "1111111  " & CStr(wC1) & " = " & CStr(m_sClubCd) & " = " & m_Rs("T13_GAKUSEI_NO") & "<br>"
		if (CStr(wC1) = CStr(m_sClubCd)) then

			if (Cint(wC1F) = 2) then
				wTaibuFlg1 = True
				
				'�����҂̏ꍇ
				if s_NyuTai_Flg = 1 Then
					gTaibubi = m_Rs("T13_CLUB_1_TAIBI")
				'�ޕ��҂̏ꍇ
				else
					gTaibubi = m_Rs2("T13_CLUB_1_TAIBI")
				end if
				gFieldName = 1		'// �N���u1���Ă���
			End if

			If (Cint(wC1F) = 1) then
				gFieldName = 1		'// �N���u1���Ă���
			End if

		End If

	End if

	if Not gf_IsNull(wC2) then
'response.write "2222222  " & CStr(wC2) & " = " & CStr(m_sClubCd) & " = " & m_Rs("T13_GAKUSEI_NO") & "<br>"
		if (CStr(wC2) = CStr(m_sClubCd)) then

			if (Cint(wC2F) = 2) then
				wTaibuFlg2 = True
				'�����҂̏ꍇ
				if s_NyuTai_Flg = 1 Then
					gTaibubi = m_Rs("T13_CLUB_2_TAIBI")
				'�ޕ��҂̏ꍇ
				else
					gTaibubi = m_Rs2("T13_CLUB_2_TAIBI")
				end if

				gFieldName = 2		'// �N���u2���Ă���
			End if

			If (Cint(wC2F) = 1) then
				gFieldName = 2		'// �N���u2���Ă���
			End if

		End if
	End if

	if wTaibuFlg1 OR wTaibuFlg2 then
		gTaibuFlg = True
	End if


End SUb



'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	dim w_NyuBi '������

%>

    <html>
    <head>
    <title>�����������ꗗ</title>
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
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�,�o�^��ʂ�\��
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){

        //���X�g����submit
		//��t���[��
		parent.topFrame.location.href="./web0360_insTop.asp?txtClubCd=<%=m_sClubCd%>"

		//���t���[��
		parent.main.location.href="./default3.asp?txtClubCd=<%=m_sClubCd%>"

        return;

    }

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Cancel(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default2.asp"
    }

	//************************************************************
	//  [�@�\]  �ޕ��{�^���������ꂽ�Ƃ�
	//  [����]  �Ȃ�
	//  [�ߒl]  �Ȃ�
	//  [����]
	//************************************************************
	function f_Taibu(){

		var i
		var w_bCheck = 1

		//�`�F�b�N�������擾
		var iMax = document.frm.chkMax.value
		if (iMax==0){
			//alert("No Avairable")
			return 1;
		}

		// ���͒l������
		iRet = f_CheckData();
		if( iRet != 0 ){
			return;
		}

		//�폜���ݒ肳��Ă���ꍇ�̃��b�Z�[�W�\��
		if(iMax==1){
			if(obj2.value == "<%=C_DELETE0%>"){
	            window.alert("�f�[�^�̍폜���ݒ肳��Ă��܂��B" + "\n" + "���s����Ɠ��ޕ��̗������폜����܂��B");
			}
		}else{
			for (i = 0; i < iMax; i++) {
				if(obj2[i].value == "<%=C_DELETE0%>"){
		            window.alert("�f�[�^�̍폜���ݒ肳��Ă��܂��B" + "\n" + "���s����Ɠ��ޕ��̗������폜����܂��B");
					break;
				}
			}
		}

		if (!confirm("�X�V���Ă���낵���ł����H")) {
			document.frm.hidTaibubi.value = "";
			return ;
		}

		//���X�g����submit
		document.frm.txtMode.value = "DELETE";
		document.frm.target = "main";
		document.frm.action = "./web0360_edt.asp"
		document.frm.submit();
		return;
	}

    //************************************************************
    //  [�@�\]  �`�F�b�N�����`�F�b�N����Ă��邩
    //  [����]  �Ȃ�
    //  [�ߒl]  0:����OK�A1:�����װ
    //************************************************************
    function f_CheckData(p_bChk) {

		obj  = eval(document.frm.txtTaibubi);
		obj2 = eval(document.frm.txtNyububiC);
		objTaibu = eval(document.frm.hidTaibuFlg);

		//�`�F�b�N�������擾
		var iMax = document.frm.chkMax.value
		if (iMax==0){
			//alert("No Avairable")
			return 1;
		}

		if(iMax==1){

			// �������`�F�b�N
			if(obj2.value == ""){
				obj2.value = "<%= gf_YYYY_MM_DD(date(),"/") %>";
			}else{
				if(obj2.value != "<%=C_DELETE0%>"){
					if( chk_dateSplit(obj2.value) == 1 ){
					    obj2.focus();
					    return 1;
					}
				}
			}

			// �܂��������̐l���Ώ�
			if(objTaibu == "False") {
				// �폜�w�肪����Ă��Ȃ��ꍇ
				if(obj2.value != "<%=C_DELETE0%>"){
					// �ޕ��������͂���Ă�����t���O�Ƀ`�F�b�N�����邩�`�F�b�N����
					if(obj.value != ""){
						if(document.frm.GAKUSEI_NO.checked==false){
							alert("�ޕ��������͂���Ă��܂��B�ޕ��o�^����ꍇ�́A�ޕ����Ƀ`�F�b�N��t���Ă��������B")
							document.frm.GAKUSEI_NO.focus();
							return 1;
						}
					}
				}
			}

			if(document.frm.GAKUSEI_NO.checked==false){
//				alert("�ޕ��o�^���鐶�k���I������Ă��܂���")
//				return 1;
			}else{
				if(obj.value == ""){
					obj.value = "<%= gf_YYYY_MM_DD(date(),"/") %>";
				}else{
					// ���t�`�F�b�N
					if( chk_dateSplit(obj.value) == 1 ){
					    obj.focus();
					    return 1;
					}
				}

				// ���t�召�`�F�b�N
		        if( DateParse(obj.value,obj2.value) >= 1){
		            window.alert("�J�n���ƏI�����𐳂������͂��Ă�������");
		            obj.focus();
		            return 1;
		        }
				document.frm.hidTaibubi.value = obj.value;
			}

		}else{

			var i
			var w_bCheck = 1
			for (i = 0; i < iMax; i++) {

				// �������`�F�b�N
				if(obj2[i].value == ""){
					obj2[i].value = "<%= gf_YYYY_MM_DD(date(),"/") %>";
				}else{
					if(obj2[i].value != "<%=C_DELETE0%>"){
						if( chk_dateSplit(obj2[i].value) == 1 ){
						    obj2[i].focus();
						    return 1;
						}
					}
				}

				// �܂��������̐l���Ώ�
				if(objTaibu[i].value == "False") {
					// �폜�w�肪����Ă��Ȃ��ꍇ
					if(obj2[i].value != "<%=C_DELETE0%>"){
						// �ޕ��������͂���Ă�����t���O�Ƀ`�F�b�N�����邩�`�F�b�N����
						if(obj[i].value != ""){
							if(document.frm.GAKUSEI_NO[i].checked==false){
								alert("�ޕ��������͂���Ă��܂��B�ޕ��o�^����ꍇ�́A�ޕ����Ƀ`�F�b�N��t���Ă��������B")
								document.frm.GAKUSEI_NO[i].focus();
								return 1;
							}
						}
					}
				}

				if(document.frm.GAKUSEI_NO[i].checked==true){
					w_bCheck = 0

					if(obj[i].value == ""){
						obj[i].value = "<%= gf_YYYY_MM_DD(date(),"/") %>";
					}else{
						// ���t�`�F�b�N
						if( chk_dateSplit(obj[i].value) == 1 ){
						    obj[i].focus();
						    return 1;
						}
					}

					// ���t�召�`�F�b�N
			        if( DateParse(obj[i].value,obj2[i].value) >= 1){
			            window.alert("�J�n���ƏI�����𐳂������͂��Ă�������");
			            obj[i].focus();
			            return 1;
			        }

					if(document.frm.hidTaibubi.value != ""){
						document.frm.hidTaibubi.value = document.frm.hidTaibubi.value + ",";
					}
					document.frm.hidTaibubi.value = document.frm.hidTaibubi.value + obj[i].value;

				}
			};

			if(w_bCheck == 1){
//				alert("�ޕ��o�^���鐶�k���I������Ă��܂���")
//				return 1;
			};
		};
        return 0;
    }
    
    //************************************************************
    //  [�@�\]  �ڍ׃{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_detail(pGAKUSEI_NO){

			url = "/cassist/gak/gak0300/kojin.asp?hidGAKUSEI_NO=" + pGAKUSEI_NO;
			w   = 700;
			h   = 630;

			wn  = "SubWindow";
			opt = "directoris=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0,resizable=no";
			if (w > 0)
				opt = opt + ",width=" + w;
			if (h > 0)
				opt = opt + ",height=" + h;
			newWin = window.open(url, wn, opt);

//		document.frm.hidGAKUSEI_NO.value = pGAKUSEI_NO;
//		document.forms[0].submit();
    }

    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">
    <center>
    <form name="frm" method="post">
	<br>

	<%
	Do 

		'=====================
		'//�ږ⋳���\��
		'=====================
		Call s_ShowKomon()
		%>

		<br>

		<%
		'=====================
		'//�o�^�A�폜�{�^��
		'=====================
		'//�ږ�̋����̏ꍇ
		If m_bKomon = True Then

			'//���������Ȃ��ꍇ
			If m_bGetMember = false Then%>
				<span class="msg">�������o�^����Ă��܂���</span>
				<br><br>
				<br>
			<%End If%>

			<table><tr><td>
				<span class="msg">
				�������҂�o�^����ۂ́u�����o�^�v�{�^�����N���b�N���Ă��������B<br>
				<%If m_bGetMember Then%>
				���ޕ��҂�o�^����ۂ͑ޕ�����w���̑ޕ������`�F�b�N(������)���A<BR>
				�@&nbsp;�ޕ�������͂̏�u�X�V�v�{�^�����N���b�N���Ă��������B�i�󔒂̏ꍇ�͏�����������܂��B�j<BR>
				���폜�i�����f�[�^��������j����ꍇ�́A���������Ɂu<%=C_DELETE0%>�v�i10���j����͂���<BR>
				�@&nbsp;�u�X�@�V�v�{�^�����N���b�N���Ă��������B<BR>
				�����t���C������ꍇ�́A���A�ޕ�������ύX���āu�X�@�V�v�{�^�����N���b�N���Ă��������B<BR>
				<%End If%>
				</span>
			</td></tr></table>

	        <table>
				<tr>
				<td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="�����o�^"></td>
				<%If m_bGetMember = True Then%><td ><input class="button" type="button" onclick="javascript:f_Taibu();" value="�@�X�@�V�@"></td><%End If%>
				</tr>
	        </table>

		<%End If%>

		<%
		'=====================
		'//���X�g���\��
		'=====================

		'//���������Ȃ��ꍇ
		If m_bGetMember = false Then
			If m_bKomon = false Then%>
				<br><br>
				<span class="msg">�������o�^����Ă��܂���</span>
			<%End If
			Exit Do
		End If
		%>


		<table>
			<tr><td valign="top" align="right">

			<table><tr><td>
				<span class="msg">
				<%If m_bKomon = True Then%>
					���͗�F�i2001/01/01 ���� <%=C_DELETE0%>�j<BR>
					���t�������͂̏ꍇ�A�����I�Ɍ��݂̓��t������܂��B<BR>
				<%End If%>
				</span>
			</td></tr></table>

			<table class=hyo border="1" bgcolor="#FFFFFF">
				<!--�w�b�_-->
				<tr>
					<%If m_bKomon = True Then%><th nowrap class="header" width="45"  align="center">�ޕ�</th><%End If%>
					<th nowrap class="header" width="40"  align="center">�N���X</th>
					<th nowrap class="header" width="40"  align="center"><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th>
					<th nowrap class="header" width="150" align="center">����</th>
					<th nowrap class="header" align="center">������</th>
					<% If m_bKomon = True Then%>
						<th nowrap class="header" align="center">�ޕ���</th>
					<% End if %>
				</tr>
		<%
		'//���s�J�E���g
		'w_iCnt = INT(m_iRsCnt/2 + 0.9)
'--- �����ҕ\�� ------------------------------------------------------------------------------------------------------------------------------
		Do Until m_Rs.EOF

			'//���ټ�Ă̸׽���Z�b�g
			Call gs_cellPtn(w_Class)
			i = i + 1
			
			'// ��������ϐ��ɑ��
			If m_sClubCd = m_Rs("T13_CLUB_1") then 
				w_NyuBi = m_Rs("T13_CLUB_1_NYUBI")
			Else
				w_NyuBi = m_Rs("T13_CLUB_2_NYUBI")
			End If

			'// �ޕ��҂��擾
			Call s_Taibu(1)
			%>
				<tr>
					<%If m_bKomon = True Then%>
						<td nowrap class="<%=w_Class%>" width="45"  align="center">
							<% If gTaibuFlg Then %>
								�ޕ���<input type="hidden" name="GAKUSEI_NO" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<% Else %>
								<input type="checkbox" name="GAKUSEI_NO" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<% End if %>
							<input type="hidden" name="hidTaibuFlg"  value="<%=gTaibuFlg%>">
							<input type="hidden" name="hidGakuseiNo" value="<%=m_Rs("T13_GAKUSEI_NO")%>">
							<input type="hidden" name="hidFieldName" value="<%=gFieldName%>">
						</td>
					<%End If%>
					<td nowrap class="<%=w_Class%>" align="center"><%=m_Rs("T13_GAKUNEN")%>-<%=f_GetClassName(m_Rs("T13_GAKUNEN"),m_Rs("T13_CLASS"))%><br></td>
					<td nowrap class="<%=w_Class%>" align="left"  ><%=m_Rs("T13_GAKUSEKI_NO")%><br></td>
					<td nowrap class="<%=w_Class%>" align="left"  ><a href="#" onClick="f_detail('<%=m_Rs("T13_GAKUSEI_NO")%>')"><%=m_Rs("T11_SIMEI")%></a><br></td>
					<%If m_bKomon = True Then%>
						<td nowrap class="<%=w_Class%>" align="center"  >
						<input type="text" style="width:80px;" name="txtNyububiC" maxlength="10" value="<%=w_NyuBi%>" id="id_Txt1<%=i-1%>">&nbsp;
						<input type="button" class="button" onclick="fcalender('id_Txt2<%=i-1%>')" value="�I��">
					<%Else%>
						<td nowrap class="<%=w_Class%>" align="center"  >
						<%=w_NyuBi%>&nbsp;
					<%End If%>
					</td>

					<% If m_bKomon = True Then %>
						<td nowrap class="<%=w_Class%>" align="center"><input type="text" style="width:80px;" name="txtTaibubi"  maxlength="10" id="id_Txt2<%=i-1%>" value="<%=gTaibubi%>">&nbsp;<input type="button" class="button" onclick="fcalender('id_Txt2<%=i-1%>')" value="�I��"></td>
					<% End if %>
				</tr>

				<%m_Rs.MoveNext%>
		<%Loop
'---------------------------------------------------------------------------------------------------------------------------------
		%>
		<%
'--- �ޕ��ҕ\�� ------------------------------------------------------------------------------------------------------------------------------
		'�ږ⋳���̏ꍇ�̂ݑޕ��҂�\��
		If m_bKomon = True Then
			Do Until m_Rs2.EOF

				'//���ټ�Ă̸׽���Z�b�g
				Call gs_cellPtn(w_Class)
				i = i + 1
				'// ��������ϐ��ɑ��
				If m_sClubCd = m_Rs2("T13_CLUB_1") then 
					w_NyuBi = m_Rs2("T13_CLUB_1_NYUBI")
				Else
					w_NyuBi = m_Rs2("T13_CLUB_2_NYUBI")
				End If

				'// �ޕ��҂��擾
				Call s_Taibu(2)
				%>
					<tr>
						<%If m_bKomon = True Then%>
							<td nowrap class="<%=w_Class%>" width="45"  align="center">
								<% If gTaibuFlg Then %>
									�ޕ���<input type="hidden" name="GAKUSEI_NO" value="<%=m_Rs2("T13_GAKUSEI_NO")%>">
								<% Else %>
									<input type="checkbox" name="GAKUSEI_NO" value="<%=m_Rs2("T13_GAKUSEI_NO")%>">
								<% End if %>
								<input type="hidden" name="hidTaibuFlg"  value="<%=gTaibuFlg%>">
								<input type="hidden" name="hidGakuseiNo" value="<%=m_Rs2("T13_GAKUSEI_NO")%>">
								<input type="hidden" name="hidFieldName" value="<%=gFieldName%>">
							</td>
						<%End If%>
						<td nowrap class="<%=w_Class%>" align="center"><%=m_Rs2("T13_GAKUNEN")%>-<%=f_GetClassName(m_Rs2("T13_GAKUNEN"),m_Rs2("T13_CLASS"))%><br></td>
						<td nowrap class="<%=w_Class%>" align="left"  ><%=m_Rs2("T13_GAKUSEKI_NO")%><br></td>
						<td nowrap class="<%=w_Class%>" align="left"  ><a href="#" onClick="f_detail('<%=m_Rs2("T13_GAKUSEI_NO")%>')"><%=m_Rs2("T11_SIMEI")%></a><br></td>
					<%If m_bKomon = True Then%>
						<td nowrap class="<%=w_Class%>" align="center"  >
						<input type="text" style="width:80px;" name="txtNyububiC" maxlength="10" value="<%=w_NyuBi%>" id="id_Txt1<%=i-1%>">&nbsp;
						<input type="button" class="button" onclick="fcalender('id_Txt2<%=i-1%>')" value="�I��">
					<%Else%>
						<td nowrap class="<%=w_Class%>" align="center"  >
						<%=w_NyuBi%>&nbsp;
					<%End If%>
					</td>
						<% If m_bKomon = True Then %>
							<td nowrap class="<%=w_Class%>" align="center"><input type="text" style="width:80px;" name="txtTaibubi"  maxlength="10" id="id_Txt2<%=i-1%>" value="<%=gTaibubi%>">&nbsp;<input type="button" class="button" onclick="fcalender('id_Txt2<%=i-1%>')" value="�I��"></td>
						<% End if %>
					</tr>

					<%m_Rs2.MoveNext%>
			<%Loop
		End If
'---------------------------------------------------------------------------------------------------------------------------------
		%>

				</table>

				<table><tr><td>
					<span class="msg">
					<%If m_bKomon = True Then%>
						���͗�F�i2001/01/01 ���� <%=C_DELETE0%>�j<BR>
						���t�������͂̏ꍇ�A�����I�Ɍ��݂̓��t������܂��B<BR>
					<%End If%>
					</span>
				</td></tr></table>

			</td></tr>
		</table>
		<br>

		<%
		'//�ږ�̋����̏ꍇ
		If m_bKomon = True Then%>
			<table>
				<tr>
					<td ><input class="button" type="button" onclick="javascript:f_Touroku();" value="�����o�^"></td>
					<td ><input class="button" type="button" onclick="javascript:f_Taibu();" value="�@�X�@�V�@"></td>
				</tr>
			</table>
		<%End If%>

		<%Set m_Rs  = Nothing%>
		<%Set m_Rs2 = Nothing%>
		<%Exit Do%>
	<%Loop%>

	<!--�l�n��-->
    <INPUT TYPE="HIDDEN" NAME="txtMode"   value = "">
	<input type="hidden" name="txtClubCd" value="<%=m_sClubCd%>">
	<input type="hidden" name="chkMax" value="<%=i%>">
	<input type="hidden" name="hidTaibubi">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>

