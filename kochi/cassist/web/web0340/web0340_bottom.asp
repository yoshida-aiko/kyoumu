<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �l���C�I���Ȗڌ���
' ��۸���ID : web/web0340/web0340_main.asp
' �@      �\: ���y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/25 �O�c
' ��      �X: 2001/08/28 �ɓ����q �w�b�_���؂藣���Ή�
' ��      �X: 2015/08/19 ���{ 1�N�Ԕԍ��̕���50��70�ɕύX
' ��      �X: 2015/08/27 ���� �Ȗڂ̃f�[�^�擾���@�ύX(T15_RISYU��T16_RISYU_KOJIN)
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
    Const DebugFlg = 6
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public  m_iMax          ':�ő�y�[�W
    Public  m_iDsp          '// �ꗗ�\���s��
    Public  m_sPageCD       ':�\���ϕ\���Ő��i�������g����󂯎������j
    Public  m_Krs           '�Ȗڗp���R�[�h�Z�b�g
    Public  m_Grs           '�w���p���R�[�h�Z�b�g
    Public  m_KSrs          '�Ȗڐ��̃��R�[�h�Z�b�g
'    Public  m_rs            '���R�[�h�Z�b�g
    Dim     m_iNendo        '//�N�x
    Dim     m_sKyokanCd     '//�����R�[�h
    Dim     m_sGakunen      '//�w�N
    Dim     m_sClass        '//�N���X
    Dim     m_sKBN          '//�敪
    Dim     m_sGRP          '//�O���[�v�敪
    Dim     m_KrCnt         '//�Ȗڂ̃��R�[�h�J�E���g
    Dim     m_KSrCnt        '//�Ȗڐ��̃��R�[�h�J�E���g
    Dim     m_GrCnt         '//�w���̃��R�[�h�J�E���g
    Dim     m_cell          '�z�F�̐ݒ�
    Dim     m_iSTani        
	Dim		m_sRisyuJotai	'���C��ԃt���O add 2001/10/25
    Dim     i               
    Dim     j               
    Dim     k               

    '�G���[�n
    Public  m_bErrFlg       '�װ�׸�
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
    w_sMsgTitle="�A�������o�^"
    w_sMsg=""
    w_sRetURL=C_RetURL & C_ERR_RETURL
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
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
        '// ���Ұ�SET
        Call s_SetParam()

		'���C��ԋ敪���擾(���C�����肵�Ă邩�ǂ����j
		'C_K_RIS_MAE = 0        '�m�菈���O
		'C_K_RIS_ATO = 1        '�m�菈����
		if f_GetKanriM(m_iNendo,C_K_RIS_JOUTAI,m_sRisyuJotai) <> 0 then 
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
	        m_bErrFlg = True
	        Call w_sMsg("�Ǘ��}�X�^�̗��C��ԋ敪������܂���B")
	        Exit Do
		end if

'-----------------------------------------------------
'm_sRisyuJotai = "1" 'test�p
'-----------------------------------------------------

        '//�Ȗڂ̏��擾
        w_iRet = f_KamokuData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

		If m_Krs.EOF Then
			Call showPage_NoData()
	        Exit Do
		End If

        '//�w���̏��擾
        w_iRet = f_GakuseiData()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

        '//�敪,�I����ʂ̑����P�ʎ擾
        w_iRet = f_Tani()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            Exit Do
        End If

        '// �y�[�W��\��
        Call showPage()

        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(m_Krs)
    '//ں��޾��CLOSE
    Call gf_closeObject(m_Grs)
    '// �I������
    Call gs_CloseDatabase()
    
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iNendo    = request("txtNendo")
    m_sKyokanCd = request("txtKyokanCd")
    m_sGakunen  = request("txtGakunen")
    m_sClass    = request("txtClass")
    m_sKBN      = Cint(request("txtKBN"))
    m_sGRP      = Cint(request("txtGRP"))
    m_iDsp      = C_PAGE_LINE

End Sub

Function f_KamokuData()
'******************************************************************
'�@�@�@�\�F�Ȗڂ̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_iNyuNendo

    On Error Resume Next
    Err.Clear
    f_KamokuData = 1

    Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

        '//�Ȗڂ̃f�[�^�擾
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT DISTINCT "
        m_sSQL = m_sSQL & vbCrLf & "     T16_KAMOKUMEI,T16_KAMOKU_CD,T16_HAITOTANI"
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T16_RISYU_KOJIN "
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T16_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_HISSEN_KBN = " & C_HISSEN_SEN & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_HAITOTANI <> " & C_T15_HAITO & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_GRP = " & m_sGRP & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T16_KAMOKU_KBN = " & m_sKBN & " "
        m_sSQL = m_sSQL & vbCrLf & " AND EXISTS ( SELECT 'X' "
        m_sSQL = m_sSQL & vbCrLf & "              FROM  "
        m_sSQL = m_sSQL & vbCrLf & "                    T11_GAKUSEKI,T13_GAKU_NEN "
        m_sSQL = m_sSQL & vbCrLf & "              WHERE  "
        m_sSQL = m_sSQL & vbCrLf & "                    T13_NENDO = T16_NENDO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_GAKUSEI_NO = T16_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T13_GAKUSEI_NO = T11_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & "              AND   T11_NYUNENDO = " & w_iNyuNendo & " "
        m_sSQL = m_sSQL & vbCrLf & "             ) "

'response.write m_sSQL & "<BR>"

        Set m_Krs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Krs, m_sSQL,m_iDsp)

'response.write "w_iRet = " & w_iRet & "<BR>"
'response.write m_Krs.EOF & "<BR>"

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_KrCnt=gf_GetRsCount(m_Krs)

    f_KamokuData = 0

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
'    If m_bErrFlg = True Then
'        w_sMsg = gf_GetErrMsg()
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
'		response.end
'    End If

End Function

Function f_GakuseiData()
'******************************************************************
'�@�@�@�\�F�w���̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_GakuseiData = 1

    Do
        '//�w���̃f�[�^�擾
        m_sSQL = ""
        m_sSQL = m_sSQL & vbCrLf & " SELECT "
        m_sSQL = m_sSQL & vbCrLf & "     T13_GAKUSEKI_NO,T11_SIMEI,T13_GAKUSEI_NO "
        m_sSQL = m_sSQL & vbCrLf & " FROM "
        m_sSQL = m_sSQL & vbCrLf & "     T13_GAKU_NEN,T11_GAKUSEKI "
        m_sSQL = m_sSQL & vbCrLf & " WHERE "
        m_sSQL = m_sSQL & vbCrLf & "     T13_GAKUSEI_NO = T11_GAKUSEI_NO(+) "
        m_sSQL = m_sSQL & vbCrLf & " AND T13_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T13_GAKUNEN = " & m_sGakunen & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & vbCrLf & " AND T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO & " "
        m_sSQL = m_sSQL & vbCrLf & " ORDER BY T13_GAKUSEKI_NO "

'response.write m_sSQL & "<BR>"

        Set m_Grs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Grs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If
    m_GrCnt=gf_GetRsCount(m_Grs)

    f_GakuseiData = 0

    Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
'    If m_bErrFlg = True Then
'        w_sMsg = gf_GetErrMsg()
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
'		response.end
'    End If

End Function

Function f_Tani()
'******************************************************************
'�@�@�@�\�F�敪,�I����ʂ̑����P�ʎ擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Dim w_iNyuNendo,w_rs

    On Error Resume Next
    Err.Clear
    f_Tani = 1

    Do

        w_iNyuNendo = Cint(m_iNendo) - Cint(m_sGakunen) + 1

        '//�敪,�I����ʂ̑����P�ʎ擾
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT "
        m_sSQL = m_sSQL & "     T18_GAKUNEN_SU"&Cint(m_sGakunen)&" "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T18_SELECTSYUBETU "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T18_NYUNENDO = " & w_iNyuNendo & " "
        m_sSQL = m_sSQL & " AND T18_GRP = " & m_sGRP & " "
        m_sSQL = m_sSQL & " AND T18_GAKKA_CD = " & m_sClass & " "
        m_sSQL = m_sSQL & " AND T18_KAMOKUSYU_CD = " & m_sKBN & " "
        m_sSQL = m_sSQL & " AND T18_GAKUNEN_SEL = " & C_T18_SEL_GAKU & " "

'response.write m_sSQL & "<BR>"
        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(w_rs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

        m_iSTani = w_rs("T18_GAKUNEN_SU"&Cint(m_sGakunen)&"")

	    f_Tani = 0

    	Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
'    If m_bErrFlg = True Then
'        w_sMsg = gf_GetErrMsg()
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
'		response.end
'    End If
    Call gf_closeObject(w_rs)

End Function

Function f_KibouData()
'******************************************************************
'�@�@�@�\�F��]���̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
	Dim w_rs,w_sSQL
	
    On Error Resume Next
    Err.Clear
    
    f_KibouData = 1

    Do
        '//��]���̃f�[�^�擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     T16_SELECT_FLG,T16_KIBOU_FLG "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     T16_RISYU_KOJIN "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "     T16_NENDO = " & m_iNendo & " "
        w_sSQL = w_sSQL & " AND T16_GAKUSEI_NO = '" & m_Grs("T13_GAKUSEI_NO") & "' "
        w_sSQL = w_sSQL & " AND T16_KAMOKU_CD = '" & m_Krs("T16_KAMOKU_CD") & "' "

'        Set w_rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordset_OpenStatic(w_rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

	If w_rs.EOF = false then
		If cint(m_sRisyuJotai) = C_K_RIS_MAE then 

			'�m�菈���O-----------------------------------------------------
			'���肵�Ă���ꍇ
	        If Cint(gf_SetNull2Zero(w_rs("T16_SELECT_FLG"))) = C_SENTAKU_YES Then%>
		        <td class=<%=m_cell%>  width="88">
		        <input type=button class=<%=m_cell%> name=button<%=k%>_<%=j%> value="��" onclick="javascript:f_Chenge(<%=k%>,<%=j%>)" style="text-align:center">
		        <input type=hidden name=MAE<%=k%>_<%=j%> value="��">
		        <input type=hidden name=ATO<%=k%>_<%=j%> value="��">
		        <input type=hidden name=KibouFLG<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
		        </td>
	        <%Else
				'��]���Ă���ꍇ
				If Cint(gf_SetNull2Zero(w_rs("T16_KIBOU_FLG"))) = 0 Then%>
			        <td class=<%=m_cell%>   width="88">
			        <input type=button class=<%=m_cell%> name=button<%=k%>_<%=j%> value="" onclick="javascript:f_Chenge(<%=k%>,<%=j%>)" style="text-align:center">
			        <input type=hidden name=MAE<%=k%>_<%=j%> value="">
			        <input type=hidden name=ATO<%=k%>_<%=j%> value="">
			        <input type=hidden name=KibouFLG<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
			        </td>

				<%Else
					'�����Ȃ��ꍇ
				%>
			        <td class=<%=m_cell%>   width="88">
			        <input type=button class=<%=m_cell%> name=button<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>' onclick="javascript:f_Chenge(<%=k%>,<%=j%>)" style="text-align:center">
			        <input type=hidden name=MAE<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
			        <input type=hidden name=ATO<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
			        <input type=hidden name=KibouFLG<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
			        </td>
				<%End If
	        End If
		Else
			'�m�菈����-----------------------------------------------------
	        If Cint(gf_SetNull2Zero(w_rs("T16_SELECT_FLG"))) = C_SENTAKU_YES Then%>
		        <td class=<%=m_cell%>  width="88" align="center">��</td>
	        <%Else
					'�����Ȃ��ꍇ
			%>
			        <td class=<%=m_cell%> width="88" align="center">�@</td>
			<%
			End If
		End If
	Else 

	  If cint(m_sRisyuJotai) = C_K_RIS_MAE then 
		'�m�菈���O-----------------------------------------------------
%>
        <td class=<%=m_cell%>   width="88">
        <input type=button class=<%=m_cell%> name=button<%=k%>_<%=j%> value="" onclick="javascript:f_Chenge(<%=k%>,<%=j%>)" style="text-align:center">
        <input type=hidden name=MAE<%=k%>_<%=j%> value="">
        <input type=hidden name=ATO<%=k%>_<%=j%> value="">
        <input type=hidden name=KibouFLG<%=k%>_<%=j%> value='<%=Cint(w_rs("T16_KIBOU_FLG"))%>'>
        </td>
<%
	  Else 
		'�m�菈����-----------------------------------------------------
%>
        <td class=<%=m_cell%> width="88" align="center">�@</td>
<%
	  End If
	End If
	    f_KibouData = 0
	    Exit Do

    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(w_rs)
    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If

End Function

Function f_KamokusuData()
'******************************************************************
'�@�@�@�\�F�Ȗڐ��̃f�[�^�擾
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    On Error Resume Next
    Err.Clear
    f_KamokusuData = 1

    m_KSrCnt=""

    Do
        '//�Ȗڐ��̃f�[�^�擾
        m_sSQL = ""
        m_sSQL = m_sSQL & " SELECT T16_KAMOKU_CD "
        m_sSQL = m_sSQL & " FROM "
        m_sSQL = m_sSQL & "     T16_RISYU_KOJIN ,T13_GAKU_NEN "
        m_sSQL = m_sSQL & " WHERE "
        m_sSQL = m_sSQL & "     T16_NENDO = " & m_iNendo & " "
        m_sSQL = m_sSQL & " AND T16_NENDO = T13_NENDO "
        m_sSQL = m_sSQL & " AND T16_GAKUSEI_NO = T13_GAKUSEI_NO "
        m_sSQL = m_sSQL & " AND T16_HAITOGAKUNEN = T13_GAKUNEN "
        m_sSQL = m_sSQL & " AND T13_CLASS = " & m_sClass & " "
        m_sSQL = m_sSQL & " AND T16_SELECT_FLG = " & C_SENTAKU_YES & " "
        m_sSQL = m_sSQL & " AND T16_KAMOKU_CD = '" & m_Krs("T16_KAMOKU_CD") & "' "
        m_sSQL = m_sSQL & " AND T16_HAITOGAKUNEN = " & m_sGakunen & " "

        Set m_KSrs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_KSrs, m_sSQL,m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            Exit Do 
        End If

	    m_KSrCnt=gf_GetRsCount(m_KSrs)

        If m_KSrs.EOF Then
            m_KSrCnt = "0"%>
	        <td class=disph><%=m_Krs("T16_KAMOKUMEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=m_KSrCnt%>" class="CELL2" name=Kamoku<%=i%> readonly></td>
        <%Else%>
	        <td class=disph><%=m_Krs("T16_KAMOKUMEI")%></td>
	        <td class=disp width=24><input type=text size=4 value="<%=m_KSrCnt%>" class="CELL2" name=Kamoku<%=i%> readonly></td>
        <%End If

	    f_KamokusuData = 0

	    Exit Do

    Loop


    '//ں��޾��CLOSE
    Call gf_closeObject(m_KSrs)

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		response.end
    End If

End Function

'********************************************************************************
'*  [�@�\]  �Ǘ��}�X�^���f�[�^���擾
'*  [����]  p_iNendo	�N�x
'*  �@�@�@  p_iNo		�����ԍ�
'*  [�ߒl]  p_iKanri	�Ǘ��f�[�^
'*  [����]  �Ǘ��}�X�^���f�[�^���擾����B
'********************************************************************************
Function f_GetKanriM(p_iNendo,p_iNo,p_sKanri)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    f_GetKanriM = 0
    p_sKanri = ""

    Do 

		'//�Ǘ��}�X�^��藚�C��ԋ敪���擾
		'//���C��ԋ敪(C_K_RIS_JOUTAI = 28)
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_NENDO=" & cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M00_KANRI.M00_NO=" & C_K_RIS_JOUTAI	'���C��ԋ敪(=28)

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            f_GetKanriM = iRet
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			'//Public Const C_K_RIS_MAE = 0    '����O
			'//Public Const C_K_RIS_ATO = 1    '�����
			p_sKanri = w_Rs("M00_KANRI")
		End If

        f_GetKanriM = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

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
	<SCRIPT language="javascript">
	<!--
    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		parent.location.href = "white.asp?txtMsg=�l���C�I���Ȗڂ̃f�[�^������܂���B"
        return;
    }
	//-->
	</SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>
    </center>
	<form name="frm" method="post">

	<input type="hidden" name="txtMsg" value="�l���C�I���Ȗڂ̃f�[�^������܂���B">

	</form>
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
Dim w_iKhalf
Dim w_iGhalf
Dim n

    On Error Resume Next
    Err.Clear

i = 0
k = 0
n = 0
%>
<HTML>


<link rel=stylesheet href="../../common/style.css" type=text/css>
    <title>�l���C�I���Ȗڌ���</title>

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

        //�w�b�_��submit
        document.frm.target = "middle";
        document.frm.action = "web0340_middle.asp"
        document.frm.submit();
        return;

    }

    //************************************************************
    //  [�@�\]  �{�^����VALUE�̕ύX
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Chenge(p_iS,p_iK){

        w_sBNm = eval("document.frm.button"+p_iS+"_"+p_iK);
        w_sMAE = eval("document.frm.MAE"+p_iS+"_"+p_iK);
        w_sATO = eval("document.frm.ATO"+p_iS+"_"+p_iK);

        //w_sKNm = eval("document.frm.Kamoku"+p_iK);
        w_sKNm = eval("parent.middle.document.frm.Kamoku"+p_iK);
        w_sKFLG = eval("document.frm.KibouFLG"+p_iS+"_"+p_iK);

        if(w_sBNm.value == "��"){
			if (w_sMAE.value == "��"){
					if (w_sKFLG.value == "0"){
			            w_sBNm.value = "";
			            w_sATO.value = "";
			            w_sKNm.value--;
					}else{
			            w_sBNm.value = w_sKFLG.value;
			            w_sATO.value = w_sKFLG.value;
			            w_sKNm.value--;
					}
			}else{
	            w_sBNm.value = w_sMAE.value;
	            w_sATO.value = w_sMAE.value;
	            w_sKNm.value--;
			}
        }else{
			if (w_sBNm.value == ""){
				if (w_sMAE.value == "��"){
		            w_sBNm.value = "��";
		            w_sATO.value = "��";
		            w_sKNm.value++;
				}else{
					if (w_sMAE.value == ""){
			            w_sBNm.value = "��";
			            w_sATO.value = "��";
			            w_sKNm.value++;
					}else{
			            w_sBNm.value = w_sMAE.value;
			            w_sATO.value = w_sMAE.value;
			            w_sKNm.value--;
					}
				}
			}else{
				if (w_sMAE.value == "��"){
		            w_sBNm.value = "��";
		            w_sATO.value = "��";
		            w_sKNm.value++;
				}else{
					if (w_sMAE.value == ""){
			            w_sBNm.value = "��";
			            w_sATO.value = "��";
			            w_sKNm.value++;
					}else{
			            w_sBNm.value = "��";
			            w_sATO.value = "��";
			            w_sKNm.value++;
					}
				}
			}
        }
        return;
    }
    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Cansel(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default2.asp"

    
    }
    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){

        var i;
        var j;
        i = 1;

<%  If m_sKBN = C_KAMOKU_IPPAN AND m_sGRP <> C_SENTAKU_JIYU Then%>

        do{
            j = 1;
            w_sTTNI = 0;
			w_sFLG = true

            do{

                w_sATO = eval("document.frm.ATO"+i+"_"+j);
                w_sTsu = eval("document.frm.Tanisuu"+j);
                w_sTsuG = eval("document.frm.txtSTani");

                if(w_sATO.value =="��"){
                    w_sTTNI = w_sTTNI + Number(w_sTsu.value);
                }
                if(w_sTTNI >= w_sTsuG.value){
                    break;
                }

            j++; }  while(j<=document.frm.n_Max.value);

            if(w_sTTNI < w_sTsuG.value){
                if (!confirm("�Œ�擾�P�ʂɒB���Ă��Ȃ��l�����܂����o�^���܂����H")) {
                   return ;
                }
                document.frm.action="web0340_upd.asp";
                document.frm.target="main";
                document.frm.submit();
				w_sFLG = false
                break;
            }
        i++; }  while(i<=document.frm.k_Max.value);

		if(w_sFLG == true){
			if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
			   return ;
			}
			document.frm.action="web0340_upd.asp";
			document.frm.target="main";
			document.frm.submit();
		}
<%Else%>
        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }
        document.frm.action="web0340_upd.asp";
        document.frm.target="main";
        document.frm.submit();
<%End If%>
    
    }
    //-->
    </SCRIPT>

	<center>

	<body onload="return window_onload()">
	<FORM NAME="frm" method="post">

	    <%
		'//�B���t�B�[���h�ɉȖ�CD�Ɗe�Ȗڂ̒P�ʐ����i�[(�o�^���Ɏg�p����)
        m_Krs.MoveFirst
        Do Until m_Krs.EOF
	        n = n + 1
		    %>
	        <input type=hidden name=kamokuCd<%=n%> value="<%=m_Krs("T16_KAMOKU_CD")%>">
	        <input type=hidden name=Tanisuu<%=n%> value="<%=m_Krs("T16_HAITOTANI")%>">
		    <%
	        m_Krs.MoveNext
        Loop%>
	<table class=hyo border=1>

	    <%
	        m_Grs.MoveFirst
	        Do Until m_Grs.EOF
	            Call gs_cellPtn(m_cell)
		        k = k + 1
		        j = 0
			    %>
			    <tr>
			        <td class=<%=m_cell%> width="70"><%=m_Grs("T13_GAKUSEKI_NO")%>
			        <input type=hidden name=gakuNo<%=k%> value="<%=m_Grs("T13_GAKUSEI_NO")%>"></td>
			        <td class=<%=m_cell%>  width="120"><%=m_Grs("T11_SIMEI")%>
			        <input type=hidden name=gakuNm<%=k%> value="<%=m_Grs("T11_SIMEI")%>"></td>
			    <%
		        m_Krs.MoveFirst
		        Do Until m_Krs.EOF
			        j = j + 1
			        Call f_KibouData() 
			        m_Krs.MoveNext
		        Loop

		        m_Grs.MoveNext
	        Loop%>
	    </tr>
	</table>
	<% If cint(m_sRisyuJotai) = C_K_RIS_MAE then %>
	<table>
	    <tr>
	        <td align=center><input type=button class=button value="�@�o�@�^�@" onclick="javascript:f_Touroku()"></td>
	        <td align=center><input type=button class=button value="�L�����Z��" onclick="javascript:f_Cansel()"></td>
	    </tr>
	</table>
	<% End If %>

	<input type="hidden" name="n_Max"       value="<%=n%>">
	<input type="hidden" name="k_Max"       value="<%=k%>">
	<input type="hidden" name="txtNendo"    value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtSTani"    value="<%=m_iSTani%>">

	<input type="hidden" name="txtGakunen"  value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClass"    value="<%=m_sClass%>">
	<input type="hidden" name="txtKBN"      value="<%=m_sKBN%>">
	<input type="hidden" name="txtGRP"      value="<%=m_sGRP%>">
	<input type="hidden" name="txtRisyu"      value="<%=m_sRisyuJotai%>">


	</FORM>
	</center>
	</BODY>
	</HTML>
<%
End Sub
%>