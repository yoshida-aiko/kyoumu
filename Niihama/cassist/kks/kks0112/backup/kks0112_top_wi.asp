<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0112/kks0112_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:txtGakunen     :�w�N
'            txtClass       :�w��
'            txtTuki        :��
' ��      ��:
'           �������\��
'               �O����̃R���{�{�b�N�X�͓�����\��
'               ���̃R���{�{�b�N�X�͓�����\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��s���ꗗ��\��������
'           ���o�^�{�^���N���b�N��
'               ���͂��ꂽ����o�^����
'-------------------------------------------------------------------------
' ��      ��: 2001/07/03 �ɓ����q
' ��      �X: 2001/12/20 �J�e�ǖ�@���w����։ȖڑΉ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 0

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '��������
    Public m_iKyokanCd          '�N�x
    Public m_sGakki             '//�w��
    Public m_sGakunen           '//�w�N
    Public m_sClassNo           '//�N���XNO
    Public m_sClassMei          '//�N���X��
    Public m_sTuki_Zenki_Start  '//�O���J�n��
    Public m_sTuki_Kouki_Start  '//����J�n��
    Public m_sTuki_Kouki_End    '//����I����
    Public m_Rs_Month           '//��
    Public m_Rs_Sbj             '//����
    Public m_Rs_Daigae          '//��֎���
    Public m_Rs_RyuDaigae       '//���w���̑�֎���
	Public m_JigenCount			'//������
    Public m_Rs_Jigen			'//���R�[�h�Z�b�g
    
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
    w_sMsgTitle="���Əo������"
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
		
        '//�w�����擾
        m_sGakki = Request("GAKKI")
		
        If trim(m_sGakki) <> "" Then
            '//�O���E��������擾
            m_sTuki_Zenki_Start = Request("Tuki_Zenki_Start")
            m_sTuki_Kouki_Start = Request("Tuki_Kouki_Start")
            m_sTuki_Kouki_End   = Request("Tuki_Kouki_End")
        Else
            '//�O���E��������擾
            w_iRet = f_GetGakkiInfo()
            If w_iRet <> 0 Then
                m_bErrFlg = True
                Exit Do
            End If
        End If
		
        '//���O�C�������̎󎝋��Ȃ��擾(�N�x�A����CD�A�w�����)
        w_iRet = f_GetSubject()
        If w_iRet <> 0 Then
           m_bErrFlg = True
            Exit Do
        End If
		
        '//��֎��Ƃ��擾
        w_iRet = f_GetDaigae()
        If w_iRet <> 0 Then
           m_bErrFlg = True
            Exit Do
        End If
		
        '//���w���̑�֎��Ƃ��擾
        w_iRet = f_GetRyuDaigae()
        If w_iRet <> 0 Then
           m_bErrFlg = True
           Exit Do
        End If
		
		'//�������̎擾
		if not f_Get_JigenData() then
			m_bErrFlg = True
			Exit Do
		end if
		
        '// �y�[�W��\��
        Call showPage()
        Exit Do
    Loop
	
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        'Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gf_closeObject(m_Rs_Month)
    Call gf_closeObject(m_Rs_Sbj)
    Call gf_closeObject(m_Rs_Daigae)

    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_sGakki    = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub

'********************************************************************************
'*  [�@�\]  �O���E��������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetGakkiInfo()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetGakkiInfo = 1

    Do
        '�Ǘ��}�X�^����w�������擾
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NO, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_KANRI, "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_BIKO"
        w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO=" & cInt(m_iSyoriNen) & " AND "
        'w_sSQL = w_sSQL & vbCrLf & "   (M00_KANRI.M00_NO=10 Or M00_KANRI.M00_NO=11 Or M00_KANRI.M00_NO=12) "   '//[M00_NO]10:�O���J�n 11:����J�n
        w_sSQL = w_sSQL & vbCrLf & "   (M00_KANRI.M00_NO=" & C_K_ZEN_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_SYURYO & ") "  '//[M00_NO]10:�O���J�n 11:����J�n

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetGakkiInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            Do Until rs.EOF

                If cInt(rs("M00_NO")) = C_K_ZEN_KAISI Then
                    m_sTuki_Zenki_Start = rs("M00_KANRI")
                ElseIf cInt(rs("M00_NO")) = C_K_KOU_KAISI Then
                    m_sTuki_Kouki_Start = rs("M00_KANRI")
                ElseIf cInt(rs("M00_NO")) = C_K_KOU_SYURYO Then
                    m_sTuki_Kouki_End = rs("M00_KANRI")
                End If
                rs.MoveNext
            Loop

            '//���݂̑O���������
            If gf_YYYY_MM_DD(date(),"/") < m_sTuki_Kouki_Start Then
                m_sGakki = "ZENKI"
            Else
                m_sGakki = "KOUKI"
            End If

        End If

        '//����I��
        f_GetGakkiInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �R���{�����擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetTuki(p_sGakki)
    Dim w_iRet
    Dim w_sSQL


    On Error Resume Next
    Err.Clear


    If p_sGakki ="ZENKI" Then

        '//�w���J�n��
        w_iStartTuki = Month(m_sTuki_Zenki_Start)

        '//�w���I����
        If day(m_sTuki_Kouki_Start) <> 1 Then
            w_iEndTuki = Month(m_sTuki_Kouki_Start)
        Else
            w_iEndTuki = Month(m_sTuki_Kouki_Start) - 1
        End If

        w_iCnt = w_iEndTuki-w_iStartTuki

        For i = 0 To w_iCnt
            w_iMonth = w_iStartTuki + i
            %>
            <option value="<%=w_iMonth%>"  <%=f_Selected(cint(w_iMonth),cint(Month(date())))%>><%=w_iMonth%>
            <%
        Next

    Else
        '//�w���J�n��
        w_iStartTuki = Month(m_sTuki_Kouki_Start)

        '//�w���I����
        w_iEndTuki = Month(m_sTuki_Kouki_End)

        w_iCnt = (12+w_iEndTuki) - w_iStartTuki

        For i = 0 To w_iCnt
            w_iMonth = w_iStartTuki + i
            If w_iMonth > 12 Then
                w_iMonth = w_iMonth - 12
            End If
            %>
            <option value="<%=w_iMonth%>"  <%=f_Selected(cint(w_iMonth),cint(Month(date())))%>><%=w_iMonth%>
            <%
        Next

    End If

End Sub

'********************************************************************************
'*  [�@�\]  ���O�C�������̎󎝋��Ȃ��擾(�N�x�A����CD�A�w�����)
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetSubject()

    Dim w_iRet
    Dim w_sSQL
    Dim w_sGakkiKbn '//�w���敪

    On Error Resume Next
    Err.Clear

    f_GetSubject = 1

    Do

        '//�O����敪���擾
        If m_sGakki = "ZENKI" Then
            w_sGakkiKbn = cstr(C_GAKKI_ZENKI)   '//1:�O��
        Else
            w_sGakkiKbn = cstr(C_GAKKI_KOUKI)   '//2:���
        End If

        '//�󎝎��Ƃ��擾
		'//�ʏ���ƂƓ��ʊ�����UNION�łȂ��ŁA���o����
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT DISTINCT "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_GAKUNEN, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_CLASS, "
        w_sSQL = w_sSQL & "  M05_CLASS.M05_CLASSMEI, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KAMOKU, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KYOKAN, "
        w_sSQL = w_sSQL & "  M03_KAMOKU.M03_KAMOKUMEI , "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_TUKU_FLG"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "  T20_JIKANWARI ,M05_CLASS,M03_KAMOKU"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_NENDO = M05_CLASS.M05_NENDO AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KAMOKU = M03_KAMOKU.M03_KAMOKU_CD AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_NENDO = M03_KAMOKU.M03_NENDO AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_GAKKI_KBN='" & w_sGakkiKbn & "' AND "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KYOKAN='" & m_iKyokanCd & "' AND "
        'w_sSQL = w_sSQL & "  (T20_JIKANWARI.T20_TUKU_FLG='0' Or T20_JIKANWARI.T20_TUKU_FLG Is Null)"
        '//C_TUKU_FLG_TUJO = "1"(0:�ʏ����,1:���ʊ���(HR��))
        w_sSQL = w_sSQL & "  (T20_JIKANWARI.T20_TUKU_FLG='" & C_TUKU_FLG_TUJO & "' Or T20_JIKANWARI.T20_TUKU_FLG Is Null)"
        w_sSQL = w_sSQL & " UNION ALL "
        w_sSQL = w_sSQL & " SELECT  DISTINCT "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_GAKUNEN, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_CLASS, "
        w_sSQL = w_sSQL & "  M05_CLASS.M05_CLASSMEI, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KAMOKU, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_KYOKAN, "
        w_sSQL = w_sSQL & "  M41_TOKUKATU.M41_MEISYO, "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_TUKU_FLG "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "  T20_JIKANWARI ,M05_CLASS,M41_TOKUKATU"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "  T20_JIKANWARI.T20_CLASS = M05_CLASS.M05_CLASSNO "
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_GAKUNEN = M05_CLASS.M05_GAKUNEN"
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_NENDO = M05_CLASS.M05_NENDO "
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_KAMOKU = M41_TOKUKATU.M41_TOKUKATU_CD "
		w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_NENDO = M41_TOKUKATU.M41_NENDO "
    	w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_NENDO=" & cInt(m_iSyoriNen) & " "
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_GAKKI_KBN='" & w_sGakkiKbn & "' "
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_KYOKAN='" & m_iKyokanCd & "' "
        'w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_TUKU_FLG='1'"   '//0:�ʏ����,1:���ʊ���(HR��)
        w_sSQL = w_sSQL & "  AND T20_JIKANWARI.T20_TUKU_FLG='" & C_TUKU_FLG_TOKU & "'"
		'//���Ƌ敪(C_JUGYO_KBN_JUHYO = 0�F���ƂƂ݂Ȃ�, C_JUGYO_KBN_NOT_JUGYO = 1:���ƂƂ݂Ȃ��Ȃ�)
        w_sSQL = w_sSQL & "  AND M41_TOKUKATU.M41_JUGYO_KBN=" & C_JUGYO_KBN_JUHYO
        w_sSQL = w_sSQL & " ORDER BY T20_GAKUNEN,T20_CLASS"

        iRet = gf_GetRecordset(m_Rs_Sbj, w_sSQL)

'response.write w_sSQL & "<br>"

        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetSubject = 99
            Exit Do
        End If

        '//����I��
        f_GetSubject = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  ���w����֎��Ԋ������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetRyuDaigae()

    Dim w_iRet
    Dim w_sSQL
    Dim w_sGakkiKbn '//�w���敪

    On Error Resume Next
    Err.Clear

    f_GetRyuDaigae = 1

    Do

        '//�O����敪���擾
        If m_sGakki = "ZENKI" Then
            w_sGakkiKbn = cstr(C_GAKKI_ZENKI)   '//1:�O��
        Else
            w_sGakkiKbn = cstr(C_GAKKI_KOUKI)   '//2:���
        End If

        '//�󎝎��Ƃ��擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & " T27.T27_GAKUNEN, "
        w_sSQL = w_sSQL & " T27.T27_CLASS, "
        w_sSQL = w_sSQL & " M05.M05_CLASSMEI, "
        w_sSQL = w_sSQL & " T27.T27_KAMOKU_CD, "
        w_sSQL = w_sSQL & " T27.T27_KYOKAN_CD, "
        w_sSQL = w_sSQL & " T16.T16_KAMOKUMEI, "
        w_sSQL = w_sSQL & " T16.T16_OKIKAE_FLG "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & " T16_RISYU_KOJIN T16,T27_TANTO_KYOKAN T27 ,M05_CLASS M05"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & " T27.T27_KYOKAN_CD 	= '"	&	m_iKyokanCd			&	"' AND "
        w_sSQL = w_sSQL & " T16.T16_OKIKAE_FLG 	= "		&	C_TIKAN_KAMOKU_SAKI	&	"  AND "
        w_sSQL = w_sSQL & " T27.T27_NENDO		= "		& 	cInt(m_iSyoriNen) 	& 	"  AND "
'        w_sSQL = w_sSQL & " T27.T27_GAKKI_KBN	= '" 	& 	w_sGakkiKbn 		& 	"' AND "
        w_sSQL = w_sSQL & " T16.T16_KAMOKU_CD 	= T27.T27_KAMOKU_CD AND "
        w_sSQL = w_sSQL & " T16.T16_NENDO 		= T27.T27_NENDO AND"
        w_sSQL = w_sSQL & " T27.T27_NENDO 		= M05.M05_NENDO AND "
        w_sSQL = w_sSQL & " T27.T27_GAKUNEN		= M05.M05_GAKUNEN AND "
        w_sSQL = w_sSQL & " T27.T27_CLASS 		= M05.M05_CLASSNO "
'response.write w_ssql
        iRet = gf_GetRecordset(m_Rs_RyuDaigae, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetRyuDaigae = 99
            Exit Do
        End If

        '//����I��
        f_GetRyuDaigae = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  ��֎��Ԋ������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDaigae()

    Dim w_iRet
    Dim w_sSQL
    Dim w_sGakkiKbn '//�w���敪

    On Error Resume Next
    Err.Clear

    f_GetDaigae = 1

    Do

        '//�O����敪���擾
        If m_sGakki = "ZENKI" Then
            w_sGakkiKbn = cstr(C_GAKKI_ZENKI)   '//1:�O��
        Else
            w_sGakkiKbn = cstr(C_GAKKI_KOUKI)   '//2:���
        End If

        '//�󎝎��Ƃ��擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_NENDO, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_YOUBI_CD, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_JIGEN, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_KAMOKU, "
        w_sSQL = w_sSQL & "  M03_KAMOKU.M03_KAMOKUMEI, "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_KYOKAN"
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN ,"
        w_sSQL = w_sSQL & "  M03_KAMOKU"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_KAMOKU = M03_KAMOKU.M03_KAMOKU_CD(+) AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_NENDO  = M03_KAMOKU.M03_NENDO(+) AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_NENDO  = " & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_KYOKAN ='" & m_iKyokanCd & "' AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_GAKUNEN Is Null AND "
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_CLASS Is Null AND"
        w_sSQL = w_sSQL & "  T23_DAIGAE_JIKAN.T23_GAKKI_KBN='" & w_sGakkiKbn & "'"
'response.write w_ssql
        iRet = gf_GetRecordset(m_Rs_Daigae, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetDaigae = 99
            Exit Do
        End If

        '//����I��
        f_GetDaigae = 0
        Exit Do
    Loop

End Function

'****************************************************
'[�@�\] �f�[�^1�ƃf�[�^2���������� "SELECTED" ��Ԃ�
'       (���X�g�_�E���{�b�N�X�I��\���p)
'[����] pData1 : �f�[�^�P
'       pData2 : �f�[�^�Q
'[�ߒl] f_Selected : "SELECTED" OR ""
'                   
'****************************************************
Function f_Selected(pData1,pData2)

    f_Selected = ""

    If IsNull(pData1) = False And IsNull(pData2) = False Then
        If trim(cStr(pData1)) = trim(cstr(pData2)) Then
            f_Selected = "selected" 
        Else
        End If
    End If

End Function

'********************************************************************************
'*	[�@�\]	�����N�x�̎������̎擾
'*	[����]	
'*	[�ߒl]	true:���� false:���s
'*	[����]	
'********************************************************************************
function f_Get_JigenData()
	Dim w_sSQL
	Dim w_iRet
	
	On Error Resume Next
	Err.Clear
	
	f_Get_JigenData = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & "  MAX(M07_JIKAN) "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "  MM07_JIGEN "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & "  M07_NENDO = " & m_iSyoriNen
	
	w_iRet = gf_GetRecordset(m_Rs_Jigen,w_sSQL)
	
	If w_iRet <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		exit function
	End If
	
	'//�f�[�^�Ȃ�
	if m_Rs_Jigen.EOF then
		f_Get_JigenData = true
		exit function
	end if
	
	m_JigenCount = cInt(m_Rs_Jigen(0))
	
	f_Get_JigenData = true
	
end function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	Dim w_num
%>
    <html>
    <head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>���Əo������</title>
	
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
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
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

        if (document.frm.KYOUKA.value==""){
            alert("���ƃf�[�^������܂���B")
            return ;
        };

        var vl = document.frm.KYOUKA.value.split('#@#');

//        if (vl[0]=='KBTU'){
//            //�ʎ���(��ʁA�ۖں��ނ��擾)
//            document.frm.SYUBETU.value=vl[0];
//            document.frm.KAMOKU_CD.value=vl[1];
//
//            document.frm.GAKUNEN.value=vl[2];
//            document.frm.KAMOKU_NAME.value=vl[3];
//
//            'document.frm.KAMOKU_NAME.value=vl[2];
//
//        }else{
            //�ʏ�E���ʎ���(��ʁA�ۖں��ށA�w�N�A�׽NO���擾)
            document.frm.SYUBETU.value=vl[0];
            document.frm.KAMOKU_CD.value=vl[1];
            document.frm.GAKUNEN.value=vl[2];
            document.frm.CLASSNO.value=vl[3];

            document.frm.CLASS_NAME.value=vl[4];
            document.frm.KAMOKU_NAME.value=vl[5];
//        }

        //document.frm.action = "./kks0110_main.asp";
        document.frm.action="./WaitAction.asp";
        document.frm.target="main";
        document.frm.submit();

    }

    //************************************************************
    //  [�@�\]  �w����ύX������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_ChangeGakki(){

        //�{��ʂ�submit
        document.frm.target = "topFrame";
        document.frm.action = "./kks0112_top.asp"
        document.frm.submit();
        return;
    }
	
	//************************************************************
	//	[�@�\]	�J�����_�[�\��
	//	[����]	�e�L�X�g�̒��g(���t)
	//	[�ߒl]	
	//	[����]
	//************************************************************
	function _g_popCalender(w_txtname){
		var wWidth = 250 ;
		var wHeigh = 230 ;
		var wStr = "calender.asp?txtName="  + w_txtname;
		var PositionX,PositionY;
		
		var wCondition = "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,titlebar=0,width=" + wWidth + ",height=" + wHeigh;
		
		PositionX = window.screen.availWidth  / 2 - wWidth / 2;
		PositionY = window.screen.availHeight / 2 - wHeigh / 2;
		
		wCondition = wCondition + ",left=" + PositionX + ",top=" + PositionY;
		
		var popup = window.open(wStr,"calender",wCondition);
		
		if(!popup.opener)popup.opener=self;
		
		if(navigator.appName=="Netscape"&&navigator.appVersion.charAt(0)=="2"){
	 		popup.location.href = location.href.substring(0,location.href.lastIndexOf("/")+1)+"../../common/com_select/SEL_CAL/default.asp";
	 	}
	}
	
    //-->
    </SCRIPT>
	
    </head>
    <body LANGUAGE="javascript" onload="return window_onload();">
    <% call gs_title("���Əo������","��@��") %>
    <form name="frm" method="post">
    <center>
    <table border="0">
	    <tr>
		    <td align="right" class="search" nowrap>

			    <table border="0">
			        <tr>
				        <td nowrap>�Ȗ�</td>
						<td colspan="7" nowrap>
				            <%
				            '//���ƃf�[�^���Ȃ��ꍇ
				            If m_Rs_Sbj.EOF And m_Rs_RyuDaigae.EOF Then
				            %>
				            <select name="KYOUKA" style="width:200px;" DISABLED>
				                <option value="">���ƃf�[�^������܂���
				            <%Else%>
				            <select name="KYOUKA" style="width:200px;">
				            <%
				            '//========================
				            '//���Ǝ��Ԋ��f�[�^��\��
				            '//========================
				                If m_Rs_Sbj.EOF = False Then
				                    Do Until m_Rs_Sbj.EOF 
				                    If m_Rs_Sbj("T20_TUKU_FLG")="1" Then
				                        '//���ʊ����̏ꍇ
				                        w_Kamoku = m_Rs_Sbj("M03_KAMOKUMEI")
				                        w_Kamoku_CD = m_Rs_Sbj("T20_KAMOKU")
				                        w_Syubetu = "TOKU"  '//���ʊ���
				                    Else
				                        w_Kamoku = m_Rs_Sbj("M03_KAMOKUMEI")
				                        w_Kamoku_CD = m_Rs_Sbj("T20_KAMOKU")
				                        w_Syubetu = "TUJO"  '//�ʏ����
				                    End If
				            %>
				                <!--<option value="<%=CStr(w_Syubetu & "#@#" & w_Kamoku_CD & "#@#" & m_Rs_Sbj("T20_GAKUNEN") & "#@#" & m_Rs_Sbj("T20_CLASS"))%>"><%=m_Rs_Sbj("T20_GAKUNEN") & "�N&nbsp;&nbsp;" & m_Rs_Sbj("M05_CLASSMEI") & "&nbsp;&nbsp;&nbsp;" & w_Kamoku %>-->
				                <option value="<%=CStr(w_Syubetu & "#@#" & w_Kamoku_CD & "#@#" & m_Rs_Sbj("T20_GAKUNEN") & "#@#" & m_Rs_Sbj("T20_CLASS")) & "#@#" &  m_Rs_Sbj("M05_CLASSMEI") & "#@#" & w_Kamoku%>"><%=m_Rs_Sbj("T20_GAKUNEN") & "�N&nbsp;&nbsp;" & m_Rs_Sbj("M05_CLASSMEI") & "&nbsp;&nbsp;&nbsp;" & w_Kamoku %>
				            <%
				                    m_Rs_Sbj.MoveNext
				                    Loop
				                End If

				                '//===========================
				                '//���w���̑�֎��Ԋ��f�[�^��ǉ��\��
				                '//===========================
				                If m_Rs_RyuDaigae.EOF = False Then
				                    Do Until m_Rs_RyuDaigae.EOF 
				                    w_Syubetu = "KBTU"  '//�ʎ���
				            %>
							<option value="<%=CStr(w_Syubetu & "#@#" & m_Rs_RyuDaigae("T27_KAMOKU_CD") & "#@#" & m_Rs_RyuDaigae("T27_GAKUNEN") & "#@#" & m_Rs_RyuDaigae("T27_CLASS")) & "#@#" &  m_Rs_RyuDaigae("M05_CLASSMEI") & "#@#" & m_Rs_RyuDaigae("T16_KAMOKUMEI")%>">���w�����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=m_Rs_RyuDaigae("T16_KAMOKUMEI")%>
				            <%
				                    m_Rs_RyuDaigae.MoveNext
				                    Loop
				                End If
				            End If

				                '//===========================
				                '//��֎��Ԋ��f�[�^��ǉ��\��
				                '//===========================
'				                If m_Rs_Daigae.EOF = False Then
'				                    Do Until m_Rs_Daigae.EOF 
'				                    w_Syubetu = "KBTU"  '//�ʎ���
'				            %>

				            <!--option Value="<=CStr(w_Syubetu & "#@#" & w_Kamoku_CD) & "#@#" & m_Rs_Daigae("M03_KAMOKUMEI")>">�ʎ���&nbsp;&nbsp;&nbsp;<=m_Rs_Daigae("M03_KAMOKUMEI")-->






				            <%
				                    'm_Rs_Daigae.MoveNext
				                    'Loop
'				                End If
'				            End If
				            %>
				            </select>
				        </td>
				    </tr>
				    
				    <tr><td colspan="8"><hr></td></tr>
				    
				    <tr>
				    	<th class="header" align="center">���ۓ���</th>
				    	<td>���t�@
				    		<input type="text" name="txtDate" value="<%=gf_YYYY_MM_DD(date(),"/")%>">
				    		<input type="button" class="button" onClick="fcalender('txtDate')" value="�I��">
				    	</td>
				    	
				    	<td>����
				    		<select name="sltJigen">
				    		
				    		<% for w_num=1 to m_JigenCount %>
				    			<option value="<%=w_num%>"><%=w_num%>
				    		<% next %>
				    	</td>
				    	
				    	<td align="right" nowrap>
							<input class="button" type="button" onclick="javascript:f_Search();" value="�@���@�́@">
						</td>
				    </tr>
				    
				    <tr>
				    	<th class="header" colspan="" align="center">���ۈꗗ�Q��</th>
				    	
				    	<td>�w�茎
				    		<select name="sltMonth">
				    			<option value="4">4
				    			<option value="5">5
				    			<option value="6">6
				    			<option value="7">7
				    			<option value="8">8
				    			<option value="9">9
				    			<option value="10">10
				    			<option value="11">11
				    			<option value="12">12
				    			<option value="1">1
				    			<option value="2">2
				    			<option value="3">3
				    		</select>
				    	</td>
				    	
				    	<td colspan="2" align="right" nowrap>
							<input class="button" type="button" onclick="javascript:f_Search();" value="�@�\�@���@">
						</td>
				    </tr>
				    
				    
			    </table>

		    </td>
	    </tr>
    </table>

    <!--�l�n���p-->
    <input type="hidden" name="Tuki_Zenki_Start" value="<%=m_sTuki_Zenki_Start%>">
    <input type="hidden" name="Tuki_Kouki_Start" value="<%=m_sTuki_Kouki_Start%>">
    <input type="hidden" name="Tuki_Kouki_End"   value="<%=m_sTuki_Kouki_End%>">
    <INPUT TYPE=HIDDEN NAME="NENDO"     value = "<%=m_iSyoriNen%>">
    <INPUT TYPE=HIDDEN NAME="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
    <INPUT TYPE=HIDDEN NAME="GAKUNEN"   value = "">
    <INPUT TYPE=HIDDEN NAME="CLASSNO"   value = "">
    <INPUT TYPE=HIDDEN NAME="KAMOKU_CD" value = "">
    <INPUT TYPE=HIDDEN NAME="SYUBETU"   value = "">

    <INPUT TYPE=HIDDEN NAME="KAMOKU_NAME"   value = "">
    <INPUT TYPE=HIDDEN NAME="CLASS_NAME" value = "">

    <input TYPE="HIDDEN" NAME="txtURL" VALUE="kks0112_bottom.asp">
    <input TYPE="HIDDEN" NAME="txtMsg" VALUE="���΂炭���҂���������">

    </form>
    </center>
    </body>
    </html>
<%
End Sub
%>