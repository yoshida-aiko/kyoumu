<%
'/************************************************************************
' �V�X�e����: �L�����p�X�A�V�X�g�V�X�e��
' ��  ��  ��: ���ʏ����|�L�����p�X�A�V�X�g
' ��۸���ID : CACommon.asp
' �@      �\: ���̃t�@�C���ɂ̓L�����p�X�A�V�X�g�ŗL�̊֐��A��`�����Ă��������B
'-------------------------------------------------------------------------
' ��      ��: 2001.03.15 ���u �m��
' ��      �X: 2001.07.12 �J�e �ǖ�
' �@      �@: 2001.07.18 ���{ ����  '//�ő厞�����\���p�֐��ǉ�
' �@      �@: 2001.07.22 �J�e �ǖ�  '//�����֌W�֐��ǉ�
' �@      �@: 2001.12.01 �c�� ��K�@'//���ȏ��o�^��Full��Normal�̈Ⴂ�����������̂��C��
' �@      �@: 2002.04.26 shin	  �@'//�ٓ����̎擾�֐�(gf_Set_Idou)�̏C��
'*************************************************************************/


'//////////////////////////////////////////////////////////////////////////////////////////
'
'	�֐��ꗗ
'
'//////////////////////////////////////////////////////////////////////////////////////////
'���݂ɃZ���ɐF������				gs_cellPtn(p_sCell)
'null��0�ɕϊ��|����ver.			gf_nInt(p_str)
'null��""�ɕϊ��|������ver.			gf_nStr(p_str)
'null���w�蕶���ɕϊ�				gf_Null(p_str,p_henkan)
'�l�̌ܓ�							gf_Round(p_num, p_keta)
'�^�C�g�����o���T�u���[�`��			gs_title(p_title,p_subtitle)
'�y�[�W�֌W�̕\���p�T�u���[�`��		gs_pageBar(p_Rs,p_sPageCD,p_iDsp,p_pageBar)
'�o���f�[�^�̎擾					gf_GetSyukketuData(p_oRecordset, p_sSikenKbn, p_sGakunen, _p_sTantoKyokan, p_sClass, p_sKamokuCD, _p_sKaisibi, p_sSyuryobi, p_s1NenBango)
'�o�����擾����J�n���ƏI�����̎擾	gf_GetKaisiSyuryo(p_sSikenKbn, p_sGakunen, p_sKaisibi, p_sSyuryobi)
'������\���i�f�o�b�O�p�j			gs_viewForm(p_form)
'�ő厞�������擾					gf_GetJigenMax(p_iJMax)
'�w�Ѝ��ڃ��x���ʕ\��				gf_empItem(p_ItemNo)
'���j���[���ڃ��x���ʕ\��			gf_empMenu(p_iMenuID)
'�s���A�N�Z�X�`�F�b�N				gf_userChk(p_PRJ_No)
'�O���E��������擾				gf_GetGakkiInfo(p_sGakki,p_sZenki_Start,p_sKouki_Start,p_sKouki_End)
'�j�����ނ���j�����̂�Ԃ�			gf_GetYoubi(p_CD)
'LOGIN�����l���S�C���ǂ����̔��f	gf_Tannin(p_Nendo,p_Kyokan)				'8/11 �O�c �ǉ�
'���݂̓��t�Ɉ�ԋ߂������敪���擾	gf_Get_SikenKbn(p_iSiken_Kbn,p_kikan,p_gakunen)	'8/17 �J�e �ǉ�
'�������̂��擾����					gf_GetKyokanNm(p_iNendo,p_sCD)
'�w�Ȗ��̂��擾����					gf_GetGakkaNm(p_iNendo,p_sCD)
'�N���X���̂��擾����				gf_GetClassName(p_iNendo,p_iGakuNen,p_ClassNo)
'1�N�Ԕԍ�,5�N�Ԕԍ��̖��̂��擾    gf_GetGakuNomei(p_iNendo,p_iGakuKBN)
'USER���̂��擾����					gf_GetUserNm(p_iNendo,p_sID)
'���ʋ����\��̌������擾           gf_GetKengen_web0300(p_sKengen)
'�g�p���ȏ��o�^�����̌������擾     gf_GetKengen_web0320(p_sKengen)
'�l���C�I���Ȗڌ���̌������擾   gf_GetKengen_web0340(p_sKengen)
'���x���ʉȖڌ���̌������擾  		gf_GetKengen_web0390(p_sKengen)
'�m�茇�ې��A�x�������擾           gf_GetKekaChi(p_iNendo,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku,p_iKekkaGai)
'�Ǘ��}�X�^���o�����ێ�ʂ��擾	gf_GetKanriInfo(p_iNendo,p_iSyubetu)
'�o���̓��͂��ł��Ȃ��Ȃ�����擾	gf_Get_SyuketuEnd(p_iGakunen,p_sEndDay)
'�X�֔ԍ�����̏Z������				gf_ComvZip(p_sKenMei,p_sSityosonCD,p_sSityosonMei,p_sZipCD,p_sTyoikiMei,p_iNendo)�@Add 2001.12.5.���c
'�ٓ��󋵃`�F�b�N�֐�				gf_Get_IdouChk(p_Gakusei_No,p_Date,p_iNendo) Add 2001.12.18 ���c
'�ٓ����̎擾�֐�					gf_Set_Idou(p_sGakusekiCd,p_iNendo,ByRef p_SSSS)

'�o�����擾����J�n���ƏI�����A
'�������т̓o�^�����������敪�̎擾 gf_GetStartEnd(p_Mode,p_SyoriNen,p_Syubetu,p_sSikenKbn,p_sGakunen,p_ClassNo,p_Kamoku,p_sKaisibi,p_sSyuryobi,p_ShikenInsertKbn)

'�������ѓo�^�̍X�V���擾			gf_GetUpdateDate(p_Nendo,p_Syubetu,p_KamokuCd,p_sGakunen,p_ClassNo,p_ShikenKbn,p_UpdateDate)

'�������{�I�������擾����			gf_GetShikenDate(p_iNendo,p_sGakunen,p_ShikenKbn,p_UpdateDate,p_Type)

'�Ȗږ����擾						gf_GetKamokuMei(p_SyoriNen,p_KamokuCd,p_KamokuKbn)


'�o���f�[�^�̎擾2					gf_GetSyukketuData2(p_oRecordset,p_sSikenKbn,p_sGakunen,p_sClass,p_sKamokuCD,p_Nendo,p_ShikenInsertType,p_Syubetu)

'�Ȗڕ]���擾						gf_GetKamokuTensuHyoka(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_iTensu,p_uData)
'�Ȗڕ]�����X�g�擾					gf_GetKamokuHyokaData(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_lDataCount,p_uData)
'���ѓ��͕��@�擾					gf_GetKamokuSeisekiInp(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_iSeiseki)
'�Ȗڑ����R�[�h�擾					gf_GetKamokuZokusei(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_iZokuseiCD)
'�Ȗڕ]���擾						gf_GetTensuHyoka(p_iNendo,p_iHyokaNo,p_iTensu,p_uData)
'�Ȗڕ]���擾						gf_GetHyokaData(p_iNendo,p_iHyokaNo,p_lDataCount,p_uData)
'�]���`��No���擾����				gf_iGetHyokaNo(p_iKamokuZokusei_CD,p_iNendo)
'�Ȗڑ����R�[�h�擾(M03_KAMOKU)		f_GetZokuseiCDTujyo(p_iNendo,p_sKamokuCD,p_iZokuseiCD)
'�Ȗڑ����R�[�h�擾(M110_NINTEI_H)	f_GetZokuseiCDNintei(p_sBunruiCD,p_iZokuseiCD)
'�Ȗڑ����R�[�h�擾(M41_TOKUKATU)	f_GetZokuseiCDToku(p_iNendo,p_sKamokuCD,p_iZokuseiCD)
'�]���`��No���擾����				gf_SeisekiInp(p_iKamokuZokusei_CD,p_iNendo,p_iSeiseki)
'�����̃��b�Z�[�W���o�͂���HTML		gs_showWhitePage(p_Msg,p_Title)
'�w�Z�ԍ��o�^�`�F�b�N				gf_ChkDisp(p_Type,p_ChkFlg)

'�F��m��O��𒲂ׂ�				gf_GetNintei(p_iNendo,p_bNiteiFlg) add 2002/09/26 shin
'�w�Z�ԍ����擾����					gf_GetGakkoNO(p_iGakkoNO)

'�ٓ����̎擾�֐��i�s���o���Łj     gf_Set_IdouGyozi(p_sGakusekiCd,p_iNendo,p_Data,ByRef p_SSSS)

'�F��m��O��𒲂ׂ�
'(�w�N�ʔF��ɑΉ�)					gf_GetGakunenNintei(p_iNendo,p_iGakunen,p_bNiteiFlg) 2003.04.11 hirota

'�F��R�[�h���擾					gf_GetNinteiCD(p_iNendo, Byref p_sNinteiCD) 2003.04.11 hirota


'** �\����` **

'** �ϐ��錾 ** 

'** �O����ۼ��ެ��` **

'////////////////////////////////////////////////////////////////////////
'// ���݂ɃZ���ɐF������
'//
'// ���@���F
'// �߂�l�F�Z���̃N���X��
'// 
'////////////////////////////////////////////////////////////////////////
sub gs_cellPtn(p_sCell) 

    if p_sCell = "" then p_sCell = C_CELL2

    if p_sCell = C_CELL1 then 
        p_sCell = C_CELL2
    else 
        p_sCell = C_CELL1
    end if

End sub

'////////////////////////////////////////////////////////////////////////
'// null��0�ɕϊ��|����ver.
'//
'// ���@���Fnull�`�F�b�N�������
'// �߂�l�Fnull��0�ɒu����������
'// 
'////////////////////////////////////////////////////////////////////////
function gf_nInt(p_nstr)
    gf_nInt=gf_Null(p_nstr,"0")
end function

'////////////////////////////////////////////////////////////////////////
'// null��""�ɕϊ��|������ver.
'//
'// ���@���Fnull�`�F�b�N�������
'// �߂�l�Fnull��""�ɒu����������
'// 
'////////////////////////////////////////////////////////////////////////
function gf_nStr(p_str)
    gf_nStr=gf_Null(p_str,"")
end function

'////////////////////////////////////////////////////////////////////////
'// null���w�蕶���ɕϊ�
'//
'// ���@���Fnull�`�F�b�N������́C�u������������
'// �߂�l�Fnull���w��̕��ɒu����������
'// 
'////////////////////////////////////////////////////////////////////////
Function  gf_Null(p_str,p_henkan) 
    if isnull(p_str) then 
        gf_Null = p_henkan
    else
        gf_Null=p_str
    end if
end function

'////////////////////////////////////////////////////////////////////////
'// �l�̌ܓ�
'//
'// ���@���F�l�̌ܓ��������l
'// �@�@�@�F�l�̌ܓ��Ώی�
'// �߂�l�F�l�̌ܓ������l
'// 
'////////////////////////////////////////////////////////////////////////
Function gf_Round(p_num, p_keta)
    Dim k
    Dim x
    If p_keta >= 0 Then
        k = CLng(10 ^ p_keta)
        x = Int(p_num * k + 0.5) / k
        gf_Round = x
    Else
        k = CLng(10 ^ (-p_keta))
        x = Int(p_num / k + 0.5) * k
        gf_Round = x
    End If
End Function

'////////////////////////////////////////////////////////////////////////
'// �^�C�g�����o���T�u���[�`��
'//
'// ���@���F�^�C�g���ƃT�u�^�C�g��
'// �߂�l�F�Ȃ�
'// 
'////////////////////////////////////////////////////////////////////////
sub gs_title(p_title,p_subtitle)

%>
    <table cellspacing="0" cellpadding="0" border="0" width="98%">
    <tr>
    <td height="27" width="100%" align="left"
    >

        <DIV class=title><%=p_title%></DIV>

    </td
    >
    </tr
    >

    <tr
    ><td height="4" width="5%" background="<%=C_IMAGE_DIR%>table_sita.gif"
    ><img src="<%=C_IMAGE_DIR%>sp.gif"
    ></td
    ></tr
    >

    <tr
    ><td height="10" class=title_Sub width="5%" align="right" valign="top"
    >

        <table class=title_Sub cellspacing="0" cellpadding="0" bgcolor=#393976 height="10" border="0"
        ><tr
        ><td align="center" valign="middle"
        ><DIV class=title_Sub
	><img src="<%=C_IMAGE_DIR%>sp.gif" width=8
        ><font color="#ffffff"
	><%=p_subtitle%></font
	><img src="<%=C_IMAGE_DIR%>sp.gif" width=8
        ></DIV
        ></td
        ></tr
        ></table
        >
    </td
    ></tr
    ></table>
<%

end sub


'********************************************************************************
'*  [�@�\]  �y�[�W�֌W�̕\���p�T�u���[�`��
'*  [����]  p_Rs            �F�ꗗ��\�����郌�R�[�h�Z�b�g
'*  �@�@�@ p_sPageCD        �F�y�[�W�ԍ�
'* �@�@�@  p_iDsp           �F1�y�[�W�̍ő�\�������B
'*  �@�@�@ p_pageBar        �F
'*  [�ߒl]  p_pageBar       �F�ł����y�[�W�o�[HTML
'*  [����]  
'********************************************************************************
sub gs_pageBar(p_Rs,p_sPageCD,p_iDsp,p_pageBar)
    Dim w_bNxt              '// NEXT�\���L��
    Dim w_bBfr              '// BEFORE�\���L��
    Dim w_iNxt              '// NEXT�\���Ő�
    Dim w_iBfr              '// BEFORE�\���Ő�
    Dim w_iCnt              '// �ް��\������
    Dim w_iMax              '// �ް��\������
    Dim i,w_iSt,w_iEd

    Dim w_iRecordCnt        '//���R�[�h�Z�b�g�J�E���g

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

    '////////////////////////////////////////
    '      �y�[�W�֌W�̐ݒ�
    '////////////////////////////////////////
    '���R�[�h�����擾
    w_iRecordCnt = gf_GetRsCount(p_Rs)
    w_iMax = gf_PageCount(p_Rs,p_iDsp)

    'EOF�̂Ƃ��̐ݒ�
    If  p_sPageCD >= w_iMax Then
        p_sPageCD = w_iMax
    End If

    '�O�y�[�W�̐ݒ�
    If INT(p_sPageCD)=1 Then
        w_bBfr=False
        w_iBfr=0
    Else
        w_bBfr=True
        w_iBfr=p_sPageCD-1
    End If

    '��y�[�W�̐ݒ�
    If p_sPageCD=w_iMax Then
        w_bNxt=False
        w_iNxt=p_sPageCD
    Else
        w_bNxt=True
        w_iNxt=p_sPageCD+1
    End If
    
	'�y�[�W�̃��X�g�̎n��(w_iSt)�ƏI���(w_iEd)����
	'��{�I�ɑI������Ă���y�[�W(p_sPageCD)���^���ɗ���悤�ɂ���B
    w_iEd = p_sPageCD + 5
    w_iSt = p_sPageCD - 4
    
	'�y�[�W�̃��X�g��10�Ȃ����A�I���y�[�W�����X�g�̐^���ɂ��Ȃ��Ƃ��B
    If p_sPageCD < 5 Then w_iEd = 10
    If w_iEd > w_iMax then w_iEd = w_iMax:w_iSt = w_iMax - 9
    If w_iSt < 1 or w_iMax < 10 then w_iSt = 1
    
    '��Βl�y�[�W�̐ݒ�
    call gs_AbsolutePage(p_Rs,p_sPageCD,p_iDsp)
    
'////////////////////////////////////////
'      �y�[�W�֌W�̐ݒ�(�����܂�)
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
'   for i = 1 to w_iMax
        If i = p_sPageCD then 
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
end sub

'*******************************************************************************
' �@�@�@�\�F�o���f�[�^�̎擾
' �ԁ@�@�l�F�擾����
' �@�@�@�@�@(True)����, (False)���s
' ���@�@���Fp_oRecordset - ���R�[�h�Z�b�g
' �@�@�@�@�@p_sSikenKbn - �����敪
' �@�@�@�@�@p_sGakunen - �w�N
' �@�@�@�@�@p_sTantoKyokan - �����b�c
' �@�@�@�@�@p_sClass - �N���XNo
' �@�@�@�@�@p_sKamokuCD - �ȖڃR�[�h
' �@�@�@�@�@p_sKaisibi - �J�n��
' �@�@�@�@�@p_sSyuryobi - �I����
' �@�@�@�@�@p_s1NenBango - �P�N�Ԕԍ�
' �@�\�ڍׁF�w�肳�ꂽ�����̏o���̃f�[�^���擾����
' ���@�@�l�F�Ȃ�
'*******************************************************************************
Function gf_GetSyukketuData(p_oRecordset,p_sSikenKbn,p_sGakunen,p_sTantoKyokan,p_sClass,p_sKamokuCD,p_sKaisibi,p_sSyuryobi,p_s1NenBango)
	
	Dim w_sSql			'SQL
	
	On Error Resume Next
	
	'== ������ ==
	gf_GetSyukketuData = False
	
	'== �o�����擾����J�n���ƏI�������擾���� ==
	'//(�����Ԃ̊���)
	If gf_GetKaisiSyuryo(p_sSikenKbn, p_sGakunen, p_sKaisibi, p_sSyuryobi) <> True Then
		Exit Function
	End If
	
	'== �o�����擾���� ==
	w_sSql = ""
	w_sSql = w_sSql & vbCrLf & "SELECT "
	w_sSql = w_sSql & vbCrLf & "	Count(T21_GAKUSEKI_NO) as KAISU,"
	w_sSql = w_sSql & vbCrLf & "	T21_CLASS,"
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & "	T21_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "FROM "
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU "
	w_sSql = w_sSql & vbCrLf & "Where "
	w_sSql = w_sSql & vbCrLf & "	T21_NENDO = " & session("NENDO") & " "		'�N�x
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_GAKUNEN = " & p_sGakunen & " "			'�w�N
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_KAMOKU = '" & p_sKamokuCD & "' " 		'�Ȗ�
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_KYOKAN = '" & p_sTantoKyokan & "' "		'����
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_HIDUKE >= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sKaisibi & "' "						'�J�n��
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_HIDUKE <= "
	w_sSql = w_sSql & vbCrLf & "	'" & p_sSyuryobi & "' "						'�I����
	w_sSql = w_sSql & vbCrLf & "	And "
	w_sSql = w_sSql & vbCrLf & "	T21_SYUKKETU_KBN IN ('" & C_KETU_KEKKA & "','" & C_KETU_TIKOKU & "','"& C_KETU_SOTAI &"','" & C_KETU_KEKKA_1 & "')"
	
	'== �P�N�Ԕԍ����w�肳��Ă���ꍇ ==
	If p_s1NenBango <> "" Then
		w_sSql = w_sSql & vbCrLf & "And "
		w_sSql = w_sSql & vbCrLf & "T21_GAKUSEKI_NO = " & p_s1NenBango & " "			'�N���X
	End If
	
	w_sSql = w_sSql & vbCrLf & "Group By "
	w_sSql = w_sSql & vbCrLf & " T21_CLASS,"
	w_sSql = w_sSql & vbCrLf & " T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & vbCrLf & " T21_GAKUSEKI_NO "
	w_sSql = w_sSql & vbCrLf & "Order By "
	w_sSql = w_sSql & vbCrLf & " T21_CLASS, "
	w_sSql = w_sSql & vbCrLf & " T21_GAKUSEKI_NO "
	
	'== �f�[�^�̎擾 ==
	Set p_oRecordset = Server.CreateObject("ADODB.Recordset")
	
	'== ���s�����Ƃ� ==
	If gf_GetRecordset(p_oRecordset, w_sSql) <> 0 Then
		p_oRecordset.Close
		Set p_oRecordset = Nothing
		Exit Function
	End If
	
	gf_GetSyukketuData = True
	
End Function

Function gf_GetKaisiSyuryo(p_sSikenKbn, p_sGakunen, p_sKaisibi, p_sSyuryobi)
'*******************************************************************************
' �@�@�@�\�F�o�����擾����J�n���ƏI�����̎擾
' �ԁ@�@�l�F�擾����
' �@�@�@�@�@(True)����, (False)���s
' ���@�@���Fp_sSikenKbn - �����敪
' �@�@�@�@�@p_sGakunen - �w�N
' �@�@�@�@�@p_sKaisibi - �J�n��
' �@�@�@�@�@p_sSyuryobi - �I����
' �@�\�ڍׁF�o�����擾����J�n���ƏI�����̎擾
' ���@�@�l�F�Ȃ�
'*******************************************************************************
	Dim w_bRtn 						'�߂�l
	Dim w_sSql
	Dim w_iNendo
	
	Dim w_oRecordset				'���R�[�h�Z�b�g
	
	w_iNendo = session("NENDO")

	On Error Resume Next
	
	'== ������ ==
	gf_GetKaisiSyuryo = False
	w_bRtn = False

	'== �����ɂ���Ď擾����f�[�^��ύX���� ==
	Select Case p_sSikenKbn
		Case C_SIKEN_ZEN_TYU		'�O������
			'== SQL�쐬 ==
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "M00_KANRI "
			w_sSql = w_sSql & vbCrLf & "From M00_KANRI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "M00_NENDO = " & w_iNendo & " "	'�N�x
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "M00_NO = " & C_K_ZEN_KAISI & " " 				'�O���J�n��
			w_sSql = w_sSql & vbCrLf & "Union "
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_KAISI "
			w_sSql = w_sSql & vbCrLf & "FROM T24_SIKEN_NITTEI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "T24_NENDO = " & w_iNendo & " "	'�N�x
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN = " & C_SIKEN_ZEN_TYU & " "		'�����敪
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_GAKUNEN = " & p_sGakunen & " "				'�w�N
			w_sSql = w_sSql & vbCrLf & "Order By "
			w_sSql = w_sSql & vbCrLf & "1"
		Case C_SIKEN_ZEN_KIM		'�O������

			'== SQL�쐬 ==
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_KAISI, "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_SYURYO "
			w_sSql = w_sSql & vbCrLf & "FROM T24_SIKEN_NITTEI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "T24_NENDO = " & w_iNendo & " "	'�N�x
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN IN (" & C_SIKEN_ZEN_TYU & " "		'�����敪
			w_sSql = w_sSql & vbCrLf & ", " & C_SIKEN_ZEN_KIM & " "		'�����敪
			w_sSql = w_sSql & vbCrLf & ") "
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_GAKUNEN = " & p_sGakunen & " "				'�w�N
			w_sSql = w_sSql & vbCrLf & "Order By "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN"

		Case C_SIKEN_KOU_TYU		'�������
			'== SQL�쐬 ==
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "M00_KANRI "
			w_sSql = w_sSql & vbCrLf & "From M00_KANRI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "M00_NENDO = " & w_iNendo & " "	'�N�x
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "M00_NO = " & C_K_KOU_KAISI & " " 				'����J�n��
			w_sSql = w_sSql & vbCrLf & "Union "
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_KAISI "
			w_sSql = w_sSql & vbCrLf & "FROM T24_SIKEN_NITTEI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "T24_NENDO = " & w_iNendo & " "	'�N�x
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN = " & C_SIKEN_KOU_TYU & " "		'�����敪
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_GAKUNEN = " & p_sGakunen & " "				'�w�N
			w_sSql = w_sSql & vbCrLf & "Order By "
			w_sSql = w_sSql & vbCrLf & "1"
			
		Case C_SIKEN_KOU_KIM		'�������
			'== SQL�쐬 ==
			w_sSql = ""
			w_sSql = w_sSql & vbCrLf & "SELECT "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_KAISI, "
			w_sSql = w_sSql & vbCrLf & "T24_JISSI_SYURYO "
			w_sSql = w_sSql & vbCrLf & "FROM T24_SIKEN_NITTEI "
			w_sSql = w_sSql & vbCrLf & "Where "
			w_sSql = w_sSql & vbCrLf & "T24_NENDO = " & w_iNendo & " "	'�N�x
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & " T24_SIKEN_KBN IN (" & C_SIKEN_KOU_TYU & " "		'�����敪
			w_sSql = w_sSql & vbCrLf & ", " & C_SIKEN_KOU_KIM & " "		'�����敪
			w_sSql = w_sSql & vbCrLf & ")"
			w_sSql = w_sSql & vbCrLf & "And "
			w_sSql = w_sSql & vbCrLf & "T24_GAKUNEN = " & p_sGakunen & " "				'�w�N
			w_sSql = w_sSql & vbCrLf & "Order By "
			w_sSql = w_sSql & vbCrLf & "T24_SIKEN_KBN"
	End Select

	'== �f�[�^�̎擾 ==
	Set w_oRecordset = Server.CreateObject("ADODB.Recordset")
	
	'== ���s�����Ƃ� ==.
	    If gf_GetRecordset(w_oRecordset, w_sSql) <> 0 Then
		w_oRecordset.Close
		Set w_oRecordset = Nothing
		
		Exit Function
	End If


	'== �Q�����Ȃ������ꍇ ==
	'If gf_GetRsCount(w_oRecordset) < 2 Then
	'	w_oRecordset.Close
	'	Set w_oRecordset = Nothing
	'	
	'	Exit Function
	'End If
	
	'== �J�n���ƏI�����̐ݒ� ==
	w_oRecordset.MoveFirst
		Select Case p_sSikenKbn
		Case C_SIKEN_ZEN_TYU, C_SIKEN_KOU_TYU		'�O�����ԁA�������
			'== �J�n�� ==
			p_sKaisibi = w_oRecordset("M00_KANRI")
			
			w_oRecordset.MoveNext
			
			'== �I���� ==
			p_sSyuryobi = FormatDateTime(DateAdd("d", -1, w_oRecordset("M00_KANRI")))
		Case C_SIKEN_ZEN_KIM, C_SIKEN_KOU_KIM		'�O�������A�������
			'== �J�n�� ==
			p_sKaisibi = FormatDateTime(DateAdd("d", 1, w_oRecordset("T24_JISSI_SYURYO")))
			w_oRecordset.MoveNext
			'== �I���� ==
			p_sSyuryobi = FormatDateTime(DateAdd("d", -1, w_oRecordset("T24_JISSI_KAISI")))
			
	End Select
	
	'== ���� ==
	w_oRecordset.Close
	Set w_oRecordset = Nothing
	
	gf_GetKaisiSyuryo = True
	
End Function

'////////////////////////////////////////////////////////////////////////
'// �y�[�W�ɓn���ꂽ������\���i�f�o�b�O�p�j
'//
'// ���@���Frequest.form
'// �߂�l�F�Ȃ�
'// �ڍׁ@�F�������������l<br>�̌`�őS�Ă̈�����\������B
'// ���l�@�Fmethod��post�̏ꍇ�ɂ̂ݗL���ł��Bget�̏ꍇ�̓v���p�e�B�����Ă��������B
'// 
'////////////////////////////////////////////////////////////////////////
sub gs_viewForm(p_form)
for each name In p_form
    response.write name&"="&p_form(name)&"<br>"
next

end sub

'/// �֐����ύX�̂��m�点�B7/20 �܂�
sub s_viewForm(p_form)
    response.write "�֐������ς��܂����B<br>"
    response.write "call gs_viewForm(request.form)<br>"
    response.write "���g���Ă��������B�J�e"
end sub

'********************************************************************************
'*  [�@�\]  �ő厞�������擾
'*  [����]  
'*  [�ߒl]  p_iJMax:�ő厞����
'*  [����]  
'********************************************************************************
Function gf_GetJigenMax(p_iJMax)

    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    p_iJMax = ""

    Do
        
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT MAX(""T20_JIGEN"") AS MAXJIGEN"
        w_sSQL = w_sSQL & vbCrLf & " FROM T20_JIKANWARI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE " 
        w_sSQL = w_sSQL & vbCrLf & " T20_NENDO = " & SESSION("NENDO")

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            'm_bErrFlg = True
            Exit Do
        End If
        
        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            'm_bErrFlg = True
            Exit Do
        End If
        
        '// �擾�����l���i�[
        p_iJMax = CInt(w_Rs("MAXJIGEN"))
        '// ����I��
        Exit Do

    Loop

    gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �w�Ѝ��ڃ��x���ʕ\��
'*  [����]  p_ItemNo�F���ڂ�NO
'*  [�ߒl]  true/false
'*  [����]  �����ʂ̍��ڕ\���ۂ��o���܂��B
'********************************************************************************
Function gf_empItem(p_ItemNo)
	Dim w_sLevel
	Dim w_iRet,w_Rs,w_sSql
	
 	gf_empItem = false

'===============================(�f�o�b�O�p)
' 	gf_empItem = True
'===============================

	w_sLevel = "T50_LEVEL" & Trim(Session("LEVEL"))
'	w_sLevel = "T50_LEVEL1"
'response.write Session("LEVEL")
    Do
	w_sSql = ""
	w_sSql = w_sSql & "Select " & w_sLevel & " "
	w_sSql = w_sSql & "From T50_KOMOKU_LEVEL "
	w_sSql = w_sSql & "Where T50_NO = " & p_ItemNo & " "

	w_iRet = gf_GetRecordset(w_Rs, w_sSql)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            'm_bErrFlg = True
            Exit Do
        End If

        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            'm_bErrFlg = True
            Exit Do
        End If

        '// �\������������ꍇ��true��Ԃ��B
        If CInt(gf_SetNull2Zero(w_Rs(w_sLevel))) = 1 Then
			gf_empItem = true
			m_HyoujiFlg = 1			'<-- �\���׸�	08/01�ǉ�(��Ŷ�)
		End if

        '// ����I��
        Exit Do

    Loop
    gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  ���j���[���ڃ��x���ʕ\��
'*  [����]  p_iMenuID�F���ڂ�NO
'*  [�ߒl]  true/false
'*  [����]  �����ʂ̍��ڕ\���ۂ��o���܂��B
'********************************************************************************
Function gf_empMenu(p_iMenuID)
	Dim w_sLevel
	Dim w_iRet,w_Rs,w_sSq
	Dim w_Where
	
	gf_empMenu = false

	'// Session("LEVEL")��NULL�Ȃ�A�ʂ���
	if gf_IsNull(Trim(Session("LEVEL"))) then Exit Function

	'// Session("LEVEL")��"0"�Ȃ�A�ʂ���
	if Cint(Session("LEVEL")) = Cint(0) then Exit Function

	w_sLevel = "T51_LEVEL" & Trim(Session("LEVEL"))

	'// WHERE���쐬
	Select Case p_iMenuID
		Case "WEB0300" : w_Where = "T51_ID in ('WEB0300','WEB0301','WEB0302')"
		Case "WEB0320" : w_Where = "T51_ID in ('WEB0320','WEB0321')"
		Case "WEB0340" : w_Where = "T51_ID in ('WEB0340','WEB0341','WEB0342')"
		Case "WEB0390" : w_Where = "T51_ID in ('WEB0390','WEB0391','WEB0392')"
		Case "SEI0200" : w_Where = "T51_ID in ('SEI0200','SEI0210','SEI0221','SEI0222','SEI0223','SEI0224','SEI0230')"
		Case "SEI0300" : w_Where = "T51_ID in ('SEI0300','SEI0301','SEI0302')"
		Case Else :		 w_Where = "T51_ID =  '" & p_iMenuID & "'"
	End Select

    Do
		w_sSql = ""
		w_sSql = w_sSql & "Select " & w_sLevel & " "
		w_sSql = w_sSql & "From T51_SYORI_LEVEL "
		w_sSql = w_sSql & "Where "
		w_sSql = w_sSql & 		w_Where
		w_sSql = w_sSql & " ORDER BY  T51_ID "
		w_iRet = gf_GetRecordset(w_Rs, w_sSql)

		If w_iRet <> 0 Then
		    'ں��޾�Ă̎擾���s
		    'm_bErrFlg = True
		    Exit Do
		End If

		If w_Rs.EOF = true Then
		    'ں��޾�Ă̎擾���s
		    'm_bErrFlg = True
		    Exit Do
		End If

		w_flg = false
		w_Rs.movefirst
		Do Until w_Rs.EOF
			If trim(gf_SetNull2String(w_Rs(w_sLevel))) = "1" then 
				w_flg = true
				exit do
			end if
		w_Rs.movenext
		Loop

		If w_flg <> true Then
		    '�Ώ�ں��ނȂ�
		    'm_bErrFlg = True
		    Exit Do
		End If

		'// �\������������ꍇ��true��Ԃ��B
	'	If Cint(gf_SetNull2Zero(w_Rs(w_sLevel))) = 1 Then gf_empMenu = true

		gf_empMenu = true

		'// ����I��
		Exit Do

    Loop

End Function

'********************************************************************************
'*  [�@�\]  ���j���[���ڃ��x���ʕ\��
'*  [����]  p_iMenuID�F���ڂ�NO
'*  [�ߒl]  true/false
'*  [����]  �����ʂ̍��ڕ\���ۂ��o���܂��B
'********************************************************************************
Function gf_empPasChg()
	Dim w_iRet,w_Rs,w_sSql,i	
	Dim w_Where
	
	gf_empPasChg = false

    Do
		w_sSql = ""
		w_sSql = w_sSql & "Select * "
		w_sSql = w_sSql & "From T51_SYORI_LEVEL "
		w_sSql = w_sSql & "Where "
		w_sSql = w_sSql & "T51_ID = 'WEB0400'"
		w_iRet = gf_GetRecordset(w_Rs, w_sSql)

		If w_iRet <> 0 Then
		    'ں��޾�Ă̎擾���s
		    'm_bErrFlg = True
		    Exit Do
		End If

		w_Rs.MoveFirst
		If w_Rs.EOF = true Then
		    'ں��޾�Ă̎擾���s
		    'm_bErrFlg = True
		    Exit Do
		End If

		w_flg = false

		For i = 3 to 12
 
			If cint(gf_nInt(w_Rs(i))) = cint(1) then 
				w_flg = true
				exit do
			end if
		Next

		For i = 18 to 56

			If cint(gf_nInt(w_Rs(i))) = cint(1) then 

				w_flg = true
				exit do
			end if

		Next

		If w_flg <> true Then
		    '�Ώ�ں��ނȂ�
		    'm_bErrFlg = True
		    Exit Do
		End If
		'// �\������������ꍇ��true��Ԃ��B
	'	If Cint(gf_SetNull2Zero(w_Rs(w_sLevel))) = 1 Then gf_empMenu = true

		gf_empPasChg = true

		'// ����I��
		Exit Do

    Loop

End Function

'********************************************************************************
'*  [�@�\]  �s���A�N�Z�X�`�F�b�N
'*  [����]  p_PRJ_No = ���������̃L�[	C_LEVEL_NOCHK�́A�������������Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �f�[�^�x�[�X�ɐڑ���Ɏg�p
'********************************************************************************
Function gf_userChk(p_PRJ_No)

	On Error Resume Next
	Err.Clear

	gf_userChk = False
	m_bErrFlg = False

	Do

		'// ���O�C���`�F�b�N
		if gf_IsNull(Session("LOGIN_ID")) then
			m_bErrFlg = True
		    w_sWinTitle="�L�����p�X�A�V�X�g"
		    w_sMsgTitle="���O�C���G���["
		    w_sRetURL = C_RetURL & "default.asp"
            m_sErrMsg = "�Z�b�V�������^�C���A�E�g����܂���\n�ēx���O�C�����Ȃ����Ă�������"
			w_sTarget = "_top"
			Exit do
		End if

		'// p_PRJ_No��C_LEVEL_NOCHK�́A�������������Ȃ�
		if p_PRJ_No = C_LEVEL_NOCHK then Exit Do

		'// �����`�F�b�N
		if Not gf_empMenu(p_PRJ_No) then
			m_bErrFlg = True
		    w_sWinTitle="�L�����p�X�A�V�X�g"
		    w_sMsgTitle="�����G���["
		    w_sRetURL = C_RetURL & "login/default.asp"
            m_sErrMsg = "����������܂���"
			w_sTarget = "_top"
			Exit do
		End if

		Exit do
	Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
		'// �����I��
'		response.end
    End If

	gf_userChk = True

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
Function gf_GetGakkiInfo(p_sGakki,p_sZenki_Start,p_sKouki_Start,p_sKouki_End)

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    gf_GetGakkiInfo = 1

	p_sZenki_Start = ""
	p_sKouki_Start = ""
	p_sKouki_End = ""
	p_sGakki = ""

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
        w_sSQL = w_sSQL & vbCrLf & "   M00_KANRI.M00_NENDO=" & SESSION("NENDO") & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   (M00_KANRI.M00_NO=" & C_K_ZEN_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_KAISI & " Or M00_KANRI.M00_NO=" & C_K_KOU_SYURYO & ") "  '//[M00_NO]10:�O���J�n 11:����J�n

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrMsg = Err.description
            gf_GetGakkiInfo = 99
            Exit Do
        End If

        If rs.EOF = False Then
            Do Until rs.EOF

                If cInt(rs("M00_NO")) = C_K_ZEN_KAISI Then
                    p_sZenki_Start = rs("M00_KANRI")
                ElseIf cInt(rs("M00_NO")) = C_K_KOU_KAISI Then
                    p_sKouki_Start = rs("M00_KANRI")
                ElseIf cInt(rs("M00_NO")) = C_K_KOU_SYURYO Then
                    p_sKouki_End = rs("M00_KANRI")
                End If
                rs.MoveNext
            Loop

            '//���݂̑O���������
            If gf_YYYY_MM_DD(date(),"/") < p_sKouki_Start Then
                p_sGakki = C_GAKKI_ZENKI
            Else
                p_sGakki = C_GAKKI_KOUKI
            End If

        End If

        '//����I��
        gf_GetGakkiInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �j�����擾
'*  [����]  p_CD(�j������)
'*  [�ߒl]  gf_GetYoubi
'*  [����]  �j���̗��̂�Ԃ�
'********************************************************************************
Function gf_GetYoubi(p_CD)
Dim w_sYoubi

    On Error Resume Next
    Err.Clear

    w_sYoubi = ""
	w_sYoubi= WeekdayName(cInt(p_CD), True)

	'//�߂�l���
    gf_GetYoubi = w_sYoubi

    Err.Clear

End Function

'********************************************************************************
'*  [�@�\]  �S�C�̊m�F
'*  [����]  p_iNendo	�����N�x
'*		   p_iKyokan�@�����R�[�h
'*		   p_iBefore	�L���N�x�i�����N�x���܂߁A���N�O�܂ł����̂ڂ��Ē��ׂ�̂��j
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function gf_Tannin(p_iNendo,p_iKyokanCd,p_iBefore)

    Dim w_iRet
    Dim w_sSQL
    Dim rs
    Dim w_Cnt

    On Error Resume Next
    Err.Clear

    gf_Tannin = 1

    Do
        '�N���X�}�X�^����S�C�����擾
		w_sSQL = ""
		w_sSQL = w_sSQL & "	SELECT "
		w_sSQL = w_sSQL & "		M05_TANNIN "
		w_sSQL = w_sSQL & "	FROM "
		w_sSQL = w_sSQL & "		M05_CLASS "
		w_sSQL = w_sSQL & "	WHERE "
		w_sSQL = w_sSQL & "		M05_TANNIN = '"& p_iKyokanCd & "' "
		w_sSQL = w_sSQL & "	AND M05_NENDO <= " & p_iNendo & " "
		w_sSQL = w_sSQL & "	AND M05_NENDO > " & p_iNendo - p_iBefore& " "
		Set rs = Server.CreateObject("ADODB.Recordset")
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			m_bErrFlg = True
			Exit Do 
		End If
		w_Cnt=cint(gf_GetRsCount(rs))
'		If w_Cnt = 0 Then
'			Exit Do
'		End If
		If rs.EOF then Exit Do	'���R�[�h�Z�b�g�����Ȃ������Ƃ��B

        '//����I��
        gf_Tannin = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  ���݂̓��t�Ɉ�ԋ߂������敪���擾
'*  [����]  p_kikan			�F�ΏۂƂȂ���ԁi���̒萔�Q�Ɓj
'* 		   p_gakunen		�F�ΏۂƂȂ�w�N�i0�̎��́A�S�w�N�j
'*  [�ߒl]  �Ȃ�
'*  [����]  �����\���͌��݂̓��t�Ɉ�ԋ߂�������m��B
'* C_SIKEN_KIKAN�F�������ԁ@C_JISSI_KIKAN�F���{���ԁ@C_SEISEKI_KIKAN�F���ѓo�^����
'********************************************************************************
Function gf_Get_SikenKbn(p_iSiken_Kbn,p_kikan,p_gakunen)
    Dim w_iRet,w_kikanFld
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    gf_Get_SikenKbn = 1
    p_iSiken_Kbn = 0
    w_kikanFld = ""
    
    Select Case p_kikan
    	case C_SIKEN_KIKAN
		w_kikanFld = "T24_SIKEN_SYURYO"
    	case C_JISSI_KIKAN
		w_kikanFld = "T24_JISSI_SYURYO"
    	case C_SEISEKI_KIKAN
		w_kikanFld = "T24_SEISEKI_SYURYO"
    End Select
    
    Do
        '���݂̓��t�Ɉ�ԋ߂������敪���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "    MIN(T24_SIKEN_KBN) as SIKEN_KBN"
        w_sSQL = w_sSQL & " FROM T24_SIKEN_NITTEI"
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "       T24_NENDO = " & session("NENDO")
if p_gakunen > 0 then 
        w_sSQL = w_sSQL & "   AND T24_GAKUNEN = " & p_gakunen
end if
        w_sSQL = w_sSQL & "   AND " & w_kikanFld & " >= '" & gf_YYYY_MM_DD(date(),"/") & "'"
'        w_sSQL = w_sSQL & " ORDER BY " & w_kikanFld &" ASC"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            gf_Get_SikenKbn = 99
            Exit Do
        End If

        'If rs.EOF = False And ISNULL(rs("SIKEN_KBN")) = False Then
        If ISNULL(rs("SIKEN_KBN")) = False Then
            p_iSiken_Kbn = cint(rs("SIKEN_KBN"))
		Else
            p_iSiken_Kbn = C_SIKEN_ZEN_TYU
        End If

'response.write w_sSQL & "<br>"
'response.write p_iSiken_Kbn & "<br>"

        '//����I��
        gf_Get_SikenKbn = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �����̎������擾(�\���p)
'*  [����]  �Ȃ�
'*  [�ߒl]  f_GetKyokanNm:��������
'*  [����]  
'********************************************************************************
Function gf_GetKyokanNm(p_iNendo,p_sCD)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    gf_GetKyokanNm = ""
	w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKANMEI_SEI,M04_KYOKANMEI_MEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M04_KYOKAN "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M04_KYOKAN_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M04_NENDO = " & p_iNendo & " "

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M04_KYOKANMEI_SEI") & "�@" & rs("M04_KYOKANMEI_MEI")
        End If

        Exit Do
    Loop

	'//�߂�l���
    gf_GetKyokanNm = w_sName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

'********************************************************************************
'*  [�@�\]  �w�Ȗ����擾(�\���p)
'*  [����]  �Ȃ�
'*  [�ߒl]  gf_GetGakkaNm:�w�Ȗ�
'*  [����]  
'********************************************************************************
Function gf_GetGakkaNm(p_iNendo,p_sCD)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    gf_GetGakkaNm = ""
	w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKAMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKA "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M02_GAKKA_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M02_NENDO = " & p_iNendo & " "

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M02_GAKKAMEI")
        End If

        Exit Do
    Loop

	'//�߂�l���
    gf_GetGakkaNm = w_sName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

'********************************************************************************
'*  [�@�\]  �N���X�����擾����
'*  [����]  p_iNendo  �F�����N�x
'*          p_iGakuNen�F�w�N
'*          p_ClassNo �F�N���XNO
'*  [�ߒl]  gf_GetClassName�F�N���X��
'*  [����]  
'********************************************************************************
Function gf_GetClassName(p_iNendo,p_iGakuNen,p_ClassNo)
	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	gf_GetClassName = ""
	w_sClassName = ""

	Do

		'//�N���X���̎擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_CLASSMEI"
		w_sSql = w_sSql & vbCrLf & " FROM M05_CLASS"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M05_CLASS.M05_NENDO=" & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN=" & p_iGakuNen
		w_sSql = w_sSql & vbCrLf & "  AND M05_CLASS.M05_CLASSNO=" & p_ClassNo

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//�N���X��
			w_sClassName = rs("M05_CLASSMEI")
		End If

		Exit Do
	Loop

	'//�߂�l���
	gf_GetClassName = w_sClassName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �P�N�Ԕԍ��A�T�N�Ԕԍ��̖��̂��擾����
'*  [����]  p_iNendo  �F�����N�x
'*          p_iGakuKBN�F�P�N�Ԕԍ�or�T�N�Ԕԍ�
'*  [�ߒl]  gf_GetGakuNomei�F����
'*  [����]  
'********************************************************************************
Function gf_GetGakuNomei(p_iNendo,p_iGakuKBN)
	Dim w_iRet
	Dim w_sSQL
	Dim w_gaku_rs

	On Error Resume Next
	Err.Clear

	gf_GetGakuNomei = ""
	w_sGakuNomei = ""

	Do

		'//�N���X���̎擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M00_KANRI "
		w_sSql = w_sSql & vbCrLf & " FROM M00_KANRI "
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M00_NENDO=" & p_iNendo
		w_sSql = w_sSql & vbCrLf & "  AND M00_NO=" & p_iGakuKBN
		w_sSql = w_sSql & vbCrLf & "  AND M00_SYUBETU= 0 "

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(w_gaku_rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If w_gaku_rs.EOF = False Then
			'//�N���X��
			w_sGakuNomei = w_gaku_rs("M00_KANRI")
		End If

		Exit Do
	Loop

	'//�߂�l���
	gf_GetGakuNomei = w_sGakuNomei

	'//ں��޾��CLOSE
	Call gf_closeObject(w_gaku_rs)

End Function

'********************************************************************************
'*  [�@�\]  USER�}�X�^���USER�����擾(�\���p)
'*  [����]  �Ȃ�
'*  [�ߒl]  gf_GetUserNm:USER��
'*  [����]  
'********************************************************************************
Function gf_GetUserNm(p_iNendo,p_sID)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    gf_GetUserNm = ""
	w_sName = ""

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M10_USER.M10_USER_NAME"
		w_sSQL = w_sSQL & vbCrLf & " FROM M10_USER"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M10_USER.M10_NENDO=" & p_iNendo
 		w_sSQL = w_sSQL & vbCrLf & " AND M10_USER.M10_USER_ID='" & p_sID & "'"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M10_USER_NAME")
        End If

        Exit Do
    Loop

	'//�߂�l���
    gf_GetUserNm = w_sName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function

'********************************************************************************
'*  [�@�\]  ���ʋ����\��̌������擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sKengen
'*  [����]  
'********************************************************************************
Function gf_GetKengen_web0300(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0300 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('WEB0300','WEB0301','WEB0302') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            gf_GetKengen_web0300 = 99
            Exit Do
        End If

		if wLevRs.Eof then
            msMsg = "�������擾�ł��܂���ł���"
            gf_GetKengen_web0300 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "WEB0300" : p_sKengen = C_ACCESS_FULL			'//�A�N�Z�X����FULL�A�N�Z�X��
			Case "WEB0301" : p_sKengen = C_ACCESS_NORMAL        '//�A�N�Z�X�������
			Case "WEB0302" : p_sKengen = C_ACCESS_VIEW          '//�A�N�Z�X�����Q�Ƃ̂�
		End Select

'		p_sKengen = C_ACCESS_FULL   'C_ACCESS_FULL   = "FULL"		
		'p_sKengen = C_ACCESS_NORMAL 'C_ACCESS_NORMAL = "NORMAL"	
		'p_sKengen = C_ACCESS_VIEW   'C_ACCESS_VIEW   = "VIEW"		

		'== ���� ==
	    Call gf_closeObject(wLevRs)

        gf_GetKengen_web0300 = 0
        Exit Do
    Loop

End Function


'********************************************************************************
'*  [�@�\]  �g�p���ȏ��o�^�����̌������擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sKengen
'*  [����]  
'********************************************************************************
Function gf_GetKengen_web0320(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0320 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('WEB0320','WEB0321') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)

        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            gf_GetKengen_web0320 = 99
            Exit Do
        End If

		if wLevRs.Eof then
            msMsg = "�������擾�ł��܂���ł���"
            gf_GetKengen_web0320 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "WEB0320" : p_sKengen = C_WEB0320_ACCESS_FULL  		'//�A�N�Z�X����FULL�A�N�Z�X��
			Case "WEB0321" : p_sKengen = C_WEB0320_ACCESS_NORMAL       '//�A�N�Z�X�������
		End Select

		'p_sKengen =  C_WEB0320_ACCESS_FULL     'C_ACCESS_FULL   = "FULL"		'//�A�N�Z�X����FULL�A�N�Z�X��
		'p_sKengen =   C_WEB0320_ACCESS_NORMAL   'C_ACCESS_NORMAL = "NORMAL"		'//�A�N�Z�X�������

		'== ���� ==
	    Call gf_closeObject(wLevRs)

        gf_GetKengen_web0320 = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �l���C�I���Ȗڌ��菈���̌������擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sKengen
'*  [����]  
'********************************************************************************
Function gf_GetKengen_web0340(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0340 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('WEB0340','WEB0341','WEB0342') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            gf_GetKengen_web0340 = 99
            Exit Do
        End If

		if wLevRs.Eof then
            msMsg = "�������擾�ł��܂���ł���"
            gf_GetKengen_web0340 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "WEB0340" : p_sKengen = C_WEB0340_ACCESS_FULL  		'//�A�N�Z�X����FULL�A�N�Z�X��
			Case "WEB0341" : p_sKengen = C_WEB0340_ACCESS_SENMON        '//�A�N�Z�X�����S�������̂�
			Case "WEB0342" : p_sKengen = C_WEB0340_ACCESS_TANNIN        '//�A�N�Z�X�����S�C�̂�
		End Select

		'p_sKengen =  C_WEB0340_ACCESS_FULL     'C_ACCESS_FULL   = "FULL"		'//�A�N�Z�X����FULL�A�N�Z�X��
'		p_sKengen =   C_WEB0340_ACCESS_SENMON   'C_ACCESS_SENMON = "SENMON"		'//�A�N�Z�X�����S�������̂�
		'p_sKengen =  C_WEB0340_ACCESS_TANNIN   'C_ACCESS_TANNIN = "TANNIN"		'//�A�N�Z�X�����S�C�̂�

		'== ���� ==
	    Call gf_closeObject(wLevRs)

        gf_GetKengen_web0340 = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  ���x���ʉȖڌ��菈���̌������擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sKengen
'*  [����]  
'********************************************************************************
Function gf_GetKengen_web0390(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0390 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('WEB0390','WEB0391','WEB0392') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            gf_GetKengen_web0390 = 99
            Exit Do
        End If

		if wLevRs.Eof then
            msMsg = "�������擾�ł��܂���ł���"
            gf_GetKengen_web0390 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "WEB0390" : p_sKengen = C_WEB0340_ACCESS_FULL  		'//�A�N�Z�X����FULL�A�N�Z�X��
			Case "WEB0391" : p_sKengen = C_WEB0340_ACCESS_SENMON        '//�A�N�Z�X�����S�������̂�
			Case "WEB0392" : p_sKengen = C_WEB0340_ACCESS_TANNIN        '//�A�N�Z�X�����S�C�̂�
		End Select

		'p_sKengen =  C_WEB0340_ACCESS_FULL     'C_ACCESS_FULL   = "FULL"		'//�A�N�Z�X����FULL�A�N�Z�X��
'		p_sKengen =   C_WEB0340_ACCESS_SENMON   'C_ACCESS_SENMON = "SENMON"		'//�A�N�Z�X�����S�������̂�
		'p_sKengen =  C_WEB0340_ACCESS_TANNIN   'C_ACCESS_TANNIN = "TANNIN"		'//�A�N�Z�X�����S�C�̂�

		'== ���� ==
	    Call gf_closeObject(wLevRs)

        gf_GetKengen_web0390 = 0
        Exit Do
    Loop

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
'********************************************************************************
Function gf_GetKekaChi(p_iNendo,p_Syubetu,p_iSikenKBN,p_sKamokuCD,p_sGakusei,p_iKekka,p_iChikoku,p_iKekkaGai)

    Dim w_sSQL
    Dim w_KekaChiRs
    Dim w_sKek,p_sChi
    Dim w_Table,w_TableName
    Dim w_Kamoku
    
    On Error Resume Next
    Err.Clear
    
    gf_GetKekaChi = 1
    
    p_iKekka = 0
    p_iChikoku = 0
	
	'���ʎ��ƁA���̑�(�ʏ�Ȃ�)�̐؂蕪��
	if trim(p_Syubetu) = "TOKU" then
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_Kamoku = "T34_TOKUKATU_CD"
	else
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_Kamoku = "T16_KAMOKU_CD"
	end if
	
	'/ �����敪�ɂ���Ď���Ă���A�t�B�[���h��ς���B
	Select Case p_iSikenKBN
		Case C_SIKEN_ZEN_TYU
			w_sKek = w_Table & "_KEKA_TYUKAN_Z"
			w_sKekG= w_Table & "_KEKA_NASI_TYUKAN_Z"
			p_sChi = w_Table & "_CHIKAI_TYUKAN_Z "
		Case C_SIKEN_ZEN_KIM
			w_sKek = w_Table & "_KEKA_KIMATU_Z"
			w_sKekG= w_Table & "_KEKA_NASI_KIMATU_Z"
			p_sChi = w_Table & "_CHIKAI_KIMATU_Z "
		Case C_SIKEN_KOU_TYU
			w_sKek = w_Table & "_KEKA_TYUKAN_K"
			w_sKekG= w_Table & "_KEKA_NASI_TYUKAN_K"
			p_sChi = w_Table & "_CHIKAI_TYUKAN_K"
		Case C_SIKEN_KOU_KIM
			w_sKek = w_Table & "_KEKA_KIMATU_K"
			w_sKekG= w_Table & "_KEKA_NASI_KIMATU_K"
			p_sChi = w_Table & "_CHIKAI_KIMATU_K"
	End Select
	
	w_sSQL = ""
	w_sSQL = w_sSQL &  " SELECT "
	w_sSQL = w_sSQL & " " & w_sKek & " as KEKA, "
	w_sSQL = w_sSQL & " " & w_sKekG & " as KEKA_NASI, "
	w_sSQL = w_sSQL & " " & p_sChi & " as CHIKAI "
	w_sSQL = w_sSQL & " FROM " & w_TableName
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "      " & w_Table & "_NENDO =" & p_iNendo
	w_sSQL = w_sSQL & "  AND " & w_Table & "_GAKUSEI_NO= '" & p_sGakusei & "'"
	w_sSQL = w_sSQL & "  AND " & w_Kamoku & "= '" & p_sKamokuCD & "'"
	
	'response.write "w_sSQL =" & w_sSQL & "<BR>"
	
    If gf_GetRecordset(w_KekaChiRs, w_sSQL) <> 0 Then
        'ں��޾�Ă̎擾���s
        msMsg = Err.description
    End If
	
	'//�߂�l���
	If w_KekaChiRs.EOF = False Then
		p_iKekka = w_KekaChiRs("KEKA")
		p_iKekkaGai = w_KekaChiRs("KEKA_NASI")
		p_iChikoku = w_KekaChiRs("CHIKAI")
	End If
	
    gf_GetKekaChi = 0
    
    Call gf_closeObject(w_KekaChiRs)

End Function

'********************************************************************************
'*  [�@�\]  �Ǘ��}�X�^���o�����ۂ̎������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sSyubetu = C_K_KEKKA_RUISEKI_SIKEN : ������(=0)
'*  [�ߒl]  p_sSyubetu = C_K_KEKKA_RUISEKI_KEI   �F�ݐ�(=1)
'*  [����]  
'********************************************************************************
Function gf_GetKanriInfo(p_iNendo,p_iSyubetu)
    Dim w_sSQL
    Dim w_Rs
    Dim w_iRet

    On Error Resume Next
    Err.Clear
    
    gf_GetKanriInfo = 1

    Do 

		'//�Ǘ��}�X�^��茇�ۗݐϏ��敪���擾
		'//���ۗݐϏ��敪(C_K_KEKKA_RUISEKI = 32)
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_SYUBETU"
		w_sSQL = w_sSQL & vbCrLf & " FROM M00_KANRI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M00_KANRI.M00_NENDO=" & cint(p_iNendo)
		w_sSQL = w_sSQL & vbCrLf & "  AND M00_KANRI.M00_NO=" & C_K_KEKKA_RUISEKI	'���ۗݐϏ��敪(=32)

'response.write w_sSQL  & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            gf_GetKanriInfo = 99
            Exit Do
        End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			'//Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '������
			'//Public Const C_K_KEKKA_RUISEKI_KEI = 1      '�ݐ�
			p_iSyubetu = w_Rs("M00_SYUBETU")

		End If

        gf_GetKanriInfo = 0
        Exit Do
    Loop

    Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �o���̓��͂��ł��Ȃ��Ȃ�����擾
'*  [����]  p_gakunen		�F�ΏۂƂȂ�w�N
'*  [�ߒl]  p_sEndDay		�F�o���̓��͂��ł��Ȃ��Ȃ��
'*  [����]  
'********************************************************************************
Function gf_Get_SyuketuEnd(p_iGakunen,p_sEndDay)
    Dim w_iRet,w_sSQL,rs
	Dim w_sGakki,w_sZenki_Start,w_sKouki_Start,w_sKouki_End
    Dim w_sDate

    On Error Resume Next
    Err.Clear

    gf_Get_SyuketuEnd = 1

	w_sDate = gf_YYYY_MM_DD(date(),"/")
	'�w�����̎擾
	call gf_GetGakkiInfo(p_sGakki,p_sZenki_Start,p_sKouki_Start,p_sKouki_End)

	'�����l����i�O���J�n���j
    p_sEndDay = p_sZenki_Start
    
    Do
        '���݂̓��t�Ɉ�ԋ߂������敪���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "    T24_JISSI_SYURYO "
'        w_sSQL = w_sSQL & "    T24_SEISEKI_SYURYO "	'--2001/12/18 add �����I����������
        w_sSQL = w_sSQL & " FROM T24_SIKEN_NITTEI "
        w_sSQL = w_sSQL & " WHERE "
        w_sSQL = w_sSQL & "       T24_NENDO = " & session("NENDO")
        w_sSQL = w_sSQL & "   AND T24_GAKUNEN = " & p_iGakunen
        w_sSQL = w_sSQL & " ORDER BY T24_JISSI_SYURYO DESC"
'        w_sSQL = w_sSQL & " ORDER BY T24_SEISEKI_SYURYO DESC" '--2001/12/18 add �����I����������

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            gf_Get_SikenKbn = 99
            Exit Do
        End If

		rs.MoveFirst
		Do Until rs.EOF

			'���ѓ��͊��ԏI�������߂��Ă�����A
			'�o�����͂ł��Ȃ��Ȃ�������̐��ѓ��͊��ԏI�����ɐݒ�
			If rs("T24_JISSI_SYURYO") < w_sDate then 
'			If rs("T24_SEISEKI_SYURYO") < w_sDate then ' --2001/12/18 add �����I����������
				p_sEndDay = rs("T24_JISSI_SYURYO")
'				p_sEndDay = rs("T24_SEISEKI_SYURYO") ' --2001/12/18 add �����I����������
				Exit Do
			End If
			rs.MoveNext

		Loop
		
        '//����I��
        gf_Get_SyuketuEnd = 0
        Exit Do
    Loop

    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �X�֔ԍ�����Z�����擾����
'*  [����]  
'*			p_sKenMei - ����
'*			p_sSityosonCD - �s����CD
'*			p_sSityosonMei - �s������
'*			p_sZipCD - �X�֔ԍ�
'*			p_sTyoikiMei - ���於
'*  [�ߒl]  �擾����
'*  [�ߒl]  True(OK),False(Cancel)
'*  [����]  
'********************************************************************************
Function gf_ComvZip(ByRef p_sZipCD,ByRef p_sKenMei,ByRef p_sSityosonCD,ByRef p_sSityosonMei,ByRef p_sTyoikiMei,ByRef p_iNendo)

    Dim w_bRtn
    Dim w_sSQL
    Dim w_oRecord
    Dim w_oclsSch
    
    On Error Resume Next
    Err.Clear
    
    '== ������ ==
    gf_ComvZip = 1
    
    p_sKenMei = ""
    p_sSityosonCD = ""
    p_sSityosonMei = ""
    p_sTyoikiMei = ""

Do
    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "M12_SITYOSON_CD,"
    w_sSQL = w_sSQL & "M12_SITYOSONMEI,"
    w_sSQL = w_sSQL & "M12_TYOIKIMEI "
    w_sSQL = w_sSQL & ", M16_KENMEI "           '2001/07/23 Mod
    w_sSQL = w_sSQL & "FROM M12_SITYOSON "
    w_sSQL = w_sSQL & ", M16_KEN "              '2001/07/23 Mod
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & " M12_YUBIN_BANGO = '" & p_sZipCD & "' "
    w_sSQL = w_sSQL & " And "
    w_sSQL = w_sSQL & " M16_NENDO = " & cint(p_iNendo)
    w_sSQL = w_sSQL & " And "
    w_sSQL = w_sSQL & " M16_KEN_CD = M12_KEN_CD "
    w_sSQL = w_sSQL & " Order By "
    w_sSQL = w_sSQL & " M12_YUBIN_BANGO, "
    w_sSQL = w_sSQL & " M12_RENBAN"

'response.write w_sSQL & "<br>"

    '== ں��޾�Ď擾 ==
    w_bRtn = gf_GetRecordset(w_oRecord, w_sSQL)

    If w_bRtn <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            gf_Get_SikenKbn = 99
            Exit Do
    End If

'    If w_oRecord.EOF = False Then

        p_sKenMei = w_oRecord("M16_KENMEI")
        p_sSityosonCD = w_oRecord("M12_SITYOSON_CD")
        p_sSityosonMei = w_oRecord("M12_SITYOSONMEI")
        p_sTyoikiMei = w_oRecord("M12_TYOIKIMEI")

'	End If

        '//����I��
        gf_ComvZip = 0
    	Exit Do
    Loop

    Call gf_closeObject(w_oRecord)
    
End Function

'********************************************************************************
'*	[�@�\]	�ٓ�����̏ꍇ�ړ��󋵂̎擾
'*	[����]	p_Gakusei_No:�w��NO
'*			p_Date		:���Ǝ��{��
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	2001.12.19 �ŁF���c
'********************************************************************************
Function gf_Get_IdouChk(p_Gakuseki_No,p_Date,p_iNendo,ByRef p_sKubunName)

	Dim w_sSQL
	Dim w_Rs
	Dim w_IdoFlg
	Dim w_sKubunName

	On Error Resume Next
	Err.Clear

	w_IdoFlg = False

	Do

		'// ���׃f�[�^
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_8, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_8"
		w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cint(p_iNendo) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO='" & p_Gakuseki_No & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM>0"

'response.write w_sSQL

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			Exit Do
		End If

		If w_Rs.EOF = 0 Then
			i = 1
			'//8�c�ő�ړ���
			Do Until Cint(i) > cint(8)    '//C_IDO_MAX_CNT = 8�c�ő�ړ���
				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) = "" Then
					Exit Do
				End If
'Response.Write "[" & gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) & " > " & p_Date & "]"
				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) > p_Date  Then
					'//1���ڂ̈ٓ����Ώۓ��t��薢���̏ꍇ�̏���
					If i = 1 then
						i = 0
					End if
					
					Exit Do
				End If
				i = i + 1
			Loop

'response.write "�w���m�n" & p_Gakuseki_No & " : i = " & i
			w_sKubunName = ""

			If i = 1 then
				'//�ŏ��̈ړ��������Ɠ���薢���̏ꍇ�A���Ɠ��Ɉړ���Ԃł͂Ȃ�
				'w_IdoFlg = False
				'w_sKubunName = ""

				w_sKubunName = gf_SetNull2String(w_Rs("T13_IDOU_KBN_" & i))

				w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i)),p_iNendo,p_sKubunName)
			Elseif i = 0 then '//1���ڂ̈ٓ����Ώۓ��t��薢���̏ꍇ

				w_bRet = False
				w_sKubunName = ""

			Else

   				w_sKubunName = gf_SetNull2String(w_Rs("T13_IDOU_KBN_" & i-1))

				 w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i-1)),p_iNendo,p_sKubunName)

			End If
'response.write "����:" & w_sKubunName & "�ٓ����R�F" & p_sKubunName
		End If

		Exit Do
	Loop

	gf_Get_IdouChk = w_sKubunName

	Call gf_closeObject(w_Rs)

	Err.Clear

End Function

'********************************************************************************
'*	[�@�\]	�ٓ����̎擾�֐�
'*	[����]	p_iGakusei_No:�w��NO
'*			p_iNendo		:�����N�x
'*	[�ߒl]	0:���擾���� 1:���s  p_SSSS : �ٓ�����
'*	[����]	2001.12.22 �ŁF���c
'*	[�C��]	2002.04.26 shin ���w�A��w�����̏ꍇ�́A�߂�l�P�ɐݒ�
'********************************************************************************
Function gf_Set_Idou(p_sGakusekiCd,p_iNendo,ByRef p_SSSS)

		gf_Set_Idou = 1

		Dim w_Date
		Dim w_SSSR
		
		w_Date = gf_YYYY_MM_DD(p_iNendo & "/" & month(date()) & "/" & day(date()),"/")
 		'//C_IDO_FUKUGAKU=3:���w�AC_IDO_TEI_KAIJO=5:��w����
		'p_SSSS = ""
		w_SSSR = ""

		p_SSSS = gf_Get_IdouChk(p_sGakusekiCd,w_Date,p_iNendo,w_SSSR)

'response.write w_Date
'response.write w_SSSR
'response.write p_SSSS

		IF CStr(p_SSSS) <> "" Then

			IF Cstr(p_SSSS) <> CStr(C_IDO_FUKUGAKU) AND Cstr(p_SSSS) <> Cstr(C_IDO_TEI_KAIJO) Then

					p_SSSS = w_SSSR

					gf_Set_Idou =0
			Else

				w_SSSR = ""
				p_SSSS = ""
			
				gf_Set_Idou = 1

			End if

		End if

'response.write p_SSSS

End Function

'********************************************************************************
'*  [�@�\]  ���C�f�[�^����X�V�����擾����B
'*  [����]  
'*			p_iNendo - �����N�x
'*			p_iGakunen - �w�N
'*			p_sGakkaCd - �w�ȃR�[�h
'*			p_sKamokuCd - �ȖڃR�[�h
'*			p_iCourseCd - �R�[�X�R�[�h
'*  [�ߒl]  �X�V���t
'*  [����]  
'********************************************************************************
Function gf_GetT16UpdDate(p_iNendo,p_iGakunen,p_sGakkaCd,p_sKamokuCd,p_iCourseCd)

    Dim w_bRtn
    Dim w_sSQL
    Dim w_oRecord
    
    On Error Resume Next
    Err.Clear
    
    '== ������ ==
    gf_GetT16UpdDate = ""

Do
    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & " T16_UPD_DATE "
    w_sSQL = w_sSQL & " FROM T16_RISYU_KOJIN "
    w_sSQL = w_sSQL & "WHERE "
    w_sSQL = w_sSQL & "     T16_NENDO        =  " & p_iNendo
    w_sSQL = w_sSQL & " And T16_HAITOGAKUNEN =  " & p_iGakunen
    w_sSQL = w_sSQL & " And T16_GAKKA_CD     = '" & p_sGakkaCd & "'"
    w_sSQL = w_sSQL & " And T16_KAMOKU_CD    = '" & p_sKamokuCd & "'"

'response.write w_sSQL & "<br>"

    '== ں��޾�Ď擾 ==
    w_bRtn = gf_GetRecordset(w_oRecord, w_sSQL)

    If w_bRtn <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            Exit Do
    End If

    gf_GetT16UpdDate = gf_SetNull2String(w_oRecord("T16_UPD_DATE"))

    Exit Do
Loop

    Call gf_closeObject(w_oRecord)
    
End Function

'*******************************************************************************
' �@�@�@�\�F�o�����擾����J�n���ƏI�����܂��A�������т̓o�^�����������敪�̎擾
' �ԁ@�@�l�F(True)����, (False)���s
' ���@�@���Fp_sSikenKbn  - �����敪
'			p_sGakunen   - �w�N
'			p_SyoriNen   - �����N�x�g�p
'			p_GakusekiNo - �w�Дԍ�
'			p_Kamoku     - �ȖڃR�[�h
'			p_Mode       - �������[�h(kks�����Əo���Aother�����ѓo�^)
'			p_Syubetu    - �Ȗڎ��(TUJO�F�ʏ����)
'			(�߂�l)p_ShikenInsertKbn - �����敪 [���т��Ƃ�Ƃ��Ɏg�p]
'			(�߂�l)p_sKaisibi   - �J�n��
'			(�߂�l)p_sSyuryobi  - �I����
'			
' �@�\�ڍׁF�o�����擾����J�n���ƏI�����̎擾
' ���@�@�l�Fgf_GetKaisiSyuryo�̃J�X�^�}�C�Y��
' �@�@�@�@�F���Ă��邪���t�̂Ƃ�����Ⴄ���ߕʊ֐���
'
' �@�@�@�@�F2002/06/13�@shin
'*******************************************************************************
Function gf_GetStartEnd(p_Mode,p_SyoriNen,p_Syubetu,p_sSikenKbn,p_sGakunen,p_ClassNo,p_Kamoku,p_sKaisibi,p_sSyuryobi,p_ShikenInsertKbn)
	
	Dim w_iNendo		'�N�x
	Dim w_UpdateDate	'�X�V��
	Dim w_sGakki		'�w��
	Dim w_sZenki_Start	'�O���J�n��
	Dim w_sKouki_Start	'����J�n��
	Dim w_sKouki_End	'����I����
	Dim w_iSyubetu
	Dim w_num
	
	On Error Resume Next
	Err.clear
	
	gf_GetStartEnd = False
	
	w_iNendo = p_SyoriNen

	'//�O���E��������擾
	if gf_GetGakkiInfo(w_sGakki,w_sZenki_Start,w_sKouki_Start,w_sKouki_End) <> 0 then : Exit function
	
	'//�I�������擾
	if p_Mode = "kks" then
		p_sSyuryobi = gf_YYYY_MM_DD(w_iNendo & "/" & month(date()) & "/" & day(date()),"/")
	else
		if not gf_GetShikenDate(w_iNendo,p_sGakunen,p_sSikenKbn,p_sSyuryobi,"END") then : exit function
	end if
	
	'//�����敪���A�O�����ԂȂ�A�����̎��ѓo�^���Ă��Ȃ����߁A�O���J�n����݌v�擾�J�n���ɂ���
	if cint(p_sSikenKbn) = 1 then
		p_sKaisibi  = gf_YYYY_MM_DD(w_sZenki_Start,"/")
		gf_GetStartEnd = True
		exit function
	end if
	
	'//���Əo���o�^(kks0100���Ă΂ꂽ�Ƃ�)
	if p_Mode = "kks" then
		'//�݌v�J�n�����擾���邽�߁A�����̎��т�o�^���������敪���擾����
		for w_num = cint(p_sSikenKbn)-1 to 1 Step -1
			

			'//�Ȗڂ̎��ѓo�^���Ă��邩���ׂ邽�߁A�X�V�����擾���� ==
			if not gf_GetUpdateDate(w_iNendo,p_Syubetu,p_Kamoku,p_sGakunen,p_ClassNo,w_num,w_UpdateDate) then : exit function
			
			if gf_SetNull2String(w_UpdateDate) <> "" then
				
				if not gf_GetShikenDate(w_iNendo,p_sGakunen,w_num,p_sKaisibi,"START") then : exit function
				
				p_ShikenInsertKbn = w_num
				
				'//�������{�I�����̎��̓�����݌v���J�n���邽�߁{�P����
				'p_sKaisibi = gf_YYYY_MM_DD(DateAdd("d",1,p_sKaisibi),"/")
				p_sKaisibi = gf_YYYY_MM_DD(p_sKaisibi,"/")
				
				gf_GetStartEnd = True
				exit function
			end if
		next
		
		'�����ɂ���̂́A�����o�^���Ă��Ȃ��Ƃ� �Ȃ̂ŁA�O���J�n�����Z�b�g
		p_sKaisibi  = gf_YYYY_MM_DD(w_sZenki_Start,"/")
		
	else
		'//�o�����ۂ̎������擾 �Ȗڋ敪(0:������,1:�ݐ�)
		If gf_GetKanriInfo(p_SyoriNen,w_iSyubetu) <> 0 Then : exit function
		
		if cint(w_iSyubetu) = C_K_KEKKA_RUISEKI_SIKEN then

			'�J�n��
			if not gf_GetShikenDate(w_iNendo,p_sGakunen,p_sSikenKbn-1,p_sKaisibi,"START") then : exit function
			
			'p_sKaisibi = gf_YYYY_MM_DD(DateAdd("d",1,p_sKaisibi),"/")
			p_sKaisibi = gf_YYYY_MM_DD(p_sKaisibi,"/")

		else
			'//�݌v�J�n�����擾���邽�߁A�����̎��т�o�^���������敪���擾����
			for w_num = cint(p_sSikenKbn)-1 to 1 Step -1
				
				'== �Ȗڂ̎��ѓo�^���Ă��邩���ׂ邽�߁A�X�V�����擾���� ==
				if not gf_GetUpdateDate(w_iNendo,p_Syubetu,p_Kamoku,p_sGakunen,p_ClassNo,w_num,w_UpdateDate) then : exit function
				
				if gf_SetNull2String(w_UpdateDate) <> "" then
					
					'//�������{�J�n�����擾����

					if not gf_GetShikenDate(w_iNendo,p_sGakunen,w_num+1,p_sKaisibi,"START") then : exit function
					
					p_ShikenInsertKbn = w_num
					
					'//�������{�I�����̎��̓�����݌v���J�n���邽�߁{�P����
					'p_sKaisibi = gf_YYYY_MM_DD(DateAdd("d",1,p_sKaisibi),"/")
					p_sKaisibi = gf_YYYY_MM_DD(p_sKaisibi,"/")
					
					gf_GetStartEnd = True
					exit function
				end if
			next
			
			'�����ɂ���̂́A�����o�^���Ă��Ȃ��Ƃ� �Ȃ̂ŁA�O���J�n�����Z�b�g
			p_sKaisibi  = gf_YYYY_MM_DD(w_sZenki_Start,"/")
			
		end if
	end if
	
	gf_GetStartEnd = True
	
End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڂ̎��ѓo�^���Ă��邩���ׂ邽�߁A�X�V�����擾����
' 
' �ԁ@�@�l�F
' �@�@�@�@�@(True)����, (False)���s
' ���@�@���F_sSikenKbn - �����敪
' �@�@�@�@�@p_Nendo - �N�x
' �@�@�@�@�@p_KamokuCd - �ȖڃR�[�h
'			p_GakusekiNo - �w��NO
'			(�߂�l)p_UpdateDate - �Ȗڎ��ѓo�^�̍X�V��
' �@�\�ڍׁF
' ���@�@�l�Fgf_GetStartEnd�Ŏg�p
'			Add 2002/06/13 shin
'*******************************************************************************
function gf_GetUpdateDate(p_Nendo,p_Syubetu,p_KamokuCd,p_sGakunen,p_ClassNo,p_ShikenKbn,p_UpdateDate)
	
	Dim w_Sql,w_Rs
	Dim w_ShikenType
	Dim w_Table
	Dim w_TableName
	Dim w_KamokuName
	
	On Error Resume Next
	Err.Clear
	
	gf_GetUpdateDate = false
	
	if trim(p_Syubetu) = "TOKU" then
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_KamokuName = "T34_TOKUKATU_CD"
	else
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	end if
	
	select case cint(p_ShikenKbn)
		case C_SIKEN_ZEN_TYU '�O�����Ԏ���
			w_ShikenType = w_Table & "_KOUSINBI_TYUKAN_Z"
			
		case C_SIKEN_ZEN_KIM '�O����������
			w_ShikenType = w_Table & "_KOUSINBI_KIMATU_Z"
			
		case C_SIKEN_KOU_TYU '������Ԏ���
			w_ShikenType = w_Table & "_KOUSINBI_TYUKAN_K"
			
		case C_SIKEN_KOU_KIM '�����������
			w_ShikenType = w_Table & "_KOUSINBI_KIMATU_K"
			
		case else
			exit function
			
	end select
	
	w_Sql = ""
	w_Sql = w_Sql & " select "
	w_Sql = w_Sql & " 		Max(" & w_ShikenType & ") "
	w_Sql = w_Sql & " from "
	w_Sql = w_Sql & " 		" & w_TableName
	w_Sql = w_Sql & " 		,T13_GAKU_NEN "
	w_Sql = w_Sql & " where "
	w_Sql = w_Sql & " 		" & w_Table & "_NENDO = " & p_Nendo
	w_Sql = w_Sql & " and	" & w_KamokuName & "= '"   & p_KamokuCd   & "' "
	w_Sql = w_Sql & " and	" & w_Table & "_GAKUSEI_NO = T13_GAKUSEI_NO "
	w_Sql = w_Sql & " and	T13_CLASS = " & p_ClassNo
	w_Sql = w_Sql & " and	T13_GAKUNEN = " & p_sGakunen
	w_Sql = w_Sql & " and	" & w_ShikenType & " is not NULL "
	
	if gf_GetRecordset(w_Rs,w_Sql) <> 0 then exit function
	
	p_UpdateDate = w_Rs(0)
	
	gf_GetUpdateDate = true
	
end function

'*******************************************************************************
' �@�@�@�\�F�������{�I�������擾����
' �ԁ@�@�l�F
' 			(True)����, (False)���s
' ���@�@���Fp_sSikenKbn - �����敪
' 			p_sGakunen - �w�N
' 			p_iNendo - �N�x
'			(�߂�l)p_UpdateDate - �I����
' �@�\�ڍׁF
' ���@�@�l�Fgf_GetStartEnd�Ŏg�p
'			Add 2002/06/13 shin
'*******************************************************************************
function gf_GetShikenDate(p_iNendo,p_sGakunen,p_ShikenKbn,p_UpdateDate,p_Type)
	Dim w_sSql,w_Rs
	
	On Error Resume Next
	Err.Clear
	
	gf_GetShikenDate = false
	
	w_sSql = ""
	w_sSql = w_sSql & " select "
	
	if p_Type = "END" then
		w_sSql = w_sSql & "		T24_IDOU_SYURYO "
	else
		w_sSql = w_sSql & "		T24_IDOU_KAISI "
	end if
	
	w_sSql = w_sSql & " from "
	w_sSql = w_sSql & "		T24_SIKEN_NITTEI "
	w_sSql = w_sSql & " Where "
	w_sSql = w_sSql & "		T24_NENDO = " & p_iNendo
	w_sSql = w_sSql & " And "
	w_sSql = w_sSql & "		T24_SIKEN_KBN = " & p_ShikenKbn
	w_sSql = w_sSql & " And "
	w_sSql = w_sSql & "		T24_GAKUNEN = " & p_sGakunen
	
	If gf_GetRecordset(w_Rs,w_sSql) <> 0 Then
		'ں��޾�Ă̎擾���s
		Exit function
	End If
	
	if w_Rs.EOF then : exit function
	
	p_UpdateDate = w_Rs(0)
	
	gf_GetShikenDate = true
	
end function

'*******************************************************************************
' �@�@�@�\�F�o���f�[�^�̎擾
' �ԁ@�@�l�F�擾����
' �@�@�@�@�@(True)����, (False)���s
' ���@�@���Fp_oRecordset - ���R�[�h�Z�b�g
' �@�@�@�@�@p_sSikenKbn - �����敪
' �@�@�@�@�@p_sGakunen - �w�N
' �@�@�@�@�@p_sClass - �N���XNo
' �@�@�@�@�@p_sKamokuCD - �ȖڃR�[�h
' �@�@�@�@�@p_sKaisibi - �J�n��
' �@�@�@�@�@p_sSyuryobi - �I����
' �@�@�@�@�@p_s1NenBango - �P�N�Ԕԍ�
' �@�\�ڍׁF�w�肳�ꂽ�����̏o���̃f�[�^���擾����
' ���@�@�l�F�Ȃ�
'*******************************************************************************
Function gf_GetSyukketuData2(p_oRecordset,p_sSikenKbn,p_sGakunen,p_sClass,p_sKamokuCD,p_Nendo,p_ShikenInsertType,p_Syubetu)
	
	Dim w_sSql
	Dim w_sKaisibi,w_sSyuryobi
	
	On Error Resume Next
	
	'== ������ ==
	gf_GetSyukketuData2 = false
	
	'== �o�����擾����J�n���ƏI�������擾���� ==
	'//(�����Ԃ̊���)

	if not gf_GetStartEnd("other",p_Nendo,p_Syubetu,p_sSikenKbn,p_sGakunen,p_sClass,p_sKamokuCD,w_sKaisibi,w_sSyuryobi,p_ShikenInsertType) then
		Exit Function
	End If
	
	'== �o�����擾���� ==
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		Sum(T21_JIKANSU) as KAISU,"
	w_sSql = w_sSql & "		T21_CLASS,"
	w_sSql = w_sSql & "		T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & "		T21_GAKUSEKI_NO "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T21_SYUKKETU "
	w_sSql = w_sSql & " Where "
	w_sSql = w_sSql & "		T21_NENDO = " & p_Nendo & " "			'�N�x
	w_sSql = w_sSql & "	And	T21_GAKUNEN = " & p_sGakunen & " "		'�w�N
	w_sSql = w_sSql & "	And T21_KAMOKU = '" & p_sKamokuCD & "' " 	'�Ȗ�
	w_sSql = w_sSql & "	And T21_HIDUKE >= '" & w_sKaisibi & "' "	'�J�n��
	w_sSql = w_sSql & "	And T21_HIDUKE <= '" & w_sSyuryobi & "' "	'�I����
	w_sSql = w_sSql & "	And T21_SYUKKETU_KBN IN (" & C_KETU_KEKKA & "," & C_KETU_TIKOKU & ","& C_KETU_SOTAI &"," & C_KETU_KEKKA_1 & ")"
	w_sSql = w_sSql & " Group By "
	w_sSql = w_sSql & "		T21_CLASS,"
	w_sSql = w_sSql & " 	T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & " 	T21_GAKUSEKI_NO "
	w_sSql = w_sSql & " Order By "
	w_sSql = w_sSql & " 	T21_CLASS, "
	w_sSql = w_sSql & " 	T21_GAKUSEKI_NO "
	

	If gf_GetRecordset(p_oRecordset,w_sSql) <> 0 Then : exit function
	

	gf_GetSyukketuData2 = True
	
End Function

'********************************************************************************
'*  [�@�\]  �w�Z�ԍ����o�^����Ă��邩�`�F�b�N����
'*  [����]  p_ChkFlg(out),p_Type(in)��[C_KEKKAGAI_DISP,C_HYOKAYOTEI_DISP,C_DATAKBN_DISP]
'*  [�ߒl]  
'*          gf_ChkDisp(True������I���AFalse���G���[)
'*  [����]  
'*  		�w�Z���Ƃɏ������Ⴄ�ۂɎg�p
'*  		p_ChkFlg��True�Ȃ珈��������
'*  		
'********************************************************************************
function gf_ChkDisp(p_Type,p_ChkFlg)
	Dim w_sSQL
	Dim w_Rs
	Const C_DISP = 1
	
	On Error Resume Next
	Err.Clear
	
	gf_ChkDisp = false
	p_ChkFlg = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "		M00_KANRI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M00_KANRI "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "      M00_NENDO = " & p_Type
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function
	
	if w_Rs.EOF then
		gf_ChkDisp = true
		exit function
	end if
	
	if cint(w_Rs(0)) = C_DISP then p_ChkFlg = true
	
	Call gf_closeObject(w_Rs)
	
	gf_ChkDisp = true
	
end function
'********************************************************************************
'*  [�@�\]  �Ȗږ����擾
'*  [����]  
'*  [�ߒl]  
'*  [����]  
'********************************************************************************
function gf_GetKamokuMei(p_SyoriNen,p_KamokuCd,p_KamokuKbn)
	Dim w_sSQL,w_Rs
    
	gf_GetKamokuMei = ""
	
	On Error Resume Next
    Err.Clear
	
	'�ʏ����
	if p_KamokuKbn = C_JIK_JUGYO then
		w_sSQL = ""
		w_sSQL = w_sSQL & "select "
		w_sSQL = w_sSQL & "		M03_KAMOKUMEI "
		w_sSQL = w_sSQL & "from"
		w_sSQL = w_sSQL & "		M03_KAMOKU "
		w_sSQL = w_sSQL & "where "
		w_sSQL = w_sSQL & "		M03_NENDO =" & cint(p_SyoriNen)
		w_sSQL = w_sSQL & "	and	M03_KAMOKU_CD = '" & p_KamokuCd & "'"
	'���ʊ���
	else
		w_sSQL = ""
		w_sSQL = w_sSQL & "select "
		w_sSQL = w_sSQL & "		M41_MEISYO "
		w_sSQL = w_sSQL & "from"
		w_sSQL = w_sSQL & "		M41_TOKUKATU "
		w_sSQL = w_sSQL & "where "
		w_sSQL = w_sSQL & "		M41_NENDO =" & cint(p_SyoriNen)
		w_sSQL = w_sSQL & "	and	M41_TOKUKATU_CD = '" & p_KamokuCd & "'"
	end if
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function
	
	gf_GetKamokuMei = w_Rs(0)
	
end function

'*******************************************************************************
' �@�@�@�\�F�o���f�[�^�̎擾
' �ԁ@�@�l�F�擾����
' �@�@�@�@�@(True)����, (False)���s
' ���@�@���Fp_oRecordset - ���R�[�h�Z�b�g
' �@�@�@�@�@p_sSikenKbn - �����敪
' �@�@�@�@�@p_sGakunen - �w�N
' �@�@�@�@�@p_sClass - �N���XNo
' �@�@�@�@�@p_sKamokuCD - �ȖڃR�[�h
' �@�@�@�@�@p_sKaisibi - �J�n��
' �@�@�@�@�@p_sSyuryobi - �I����
' �@�@�@�@�@p_s1NenBango - �P�N�Ԕԍ�
' �@�\�ڍׁF�w�肳�ꂽ�����̏o���̃f�[�^���擾����
' ���@�@�l�F�Ȃ�
'*******************************************************************************
Function gf_GetSyukketuData2(p_oRecordset,p_sSikenKbn,p_sGakunen,p_sClass,p_sKamokuCD,p_Nendo,p_ShikenInsertType,p_Syubetu)
	Dim w_sSql
	Dim w_sKaisibi,w_sSyuryobi
	
	On Error Resume Next
	
	'== ������ ==
	gf_GetSyukketuData2 = false

	'== �o�����擾����J�n���ƏI�������擾���� ==
	'//(�����Ԃ̊���)
	if not gf_GetStartEnd("other",p_Nendo,p_Syubetu,p_sSikenKbn,p_sGakunen,p_sClass,p_sKamokuCD,w_sKaisibi,w_sSyuryobi,p_ShikenInsertType) then
		Exit Function
	End If
	
	'== �o�����擾���� ==
	w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & "		Sum(T21_JIKANSU) as KAISU,"
	w_sSql = w_sSql & "		T21_CLASS,"
	w_sSql = w_sSql & "		T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & "		T21_GAKUSEKI_NO "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & "		T21_SYUKKETU "
	w_sSql = w_sSql & " Where "
	w_sSql = w_sSql & "		T21_NENDO = " & p_Nendo & " "			'�N�x
	w_sSql = w_sSql & "	And	T21_GAKUNEN = " & p_sGakunen & " "		'�w�N
	w_sSql = w_sSql & "	And T21_KAMOKU = '" & p_sKamokuCD & "' " 	'�Ȗ�
	w_sSql = w_sSql & "	And T21_HIDUKE >= '" & w_sKaisibi & "' "	'�J�n��
	w_sSql = w_sSql & "	And T21_HIDUKE <= '" & w_sSyuryobi & "' "	'�I����
	w_sSql = w_sSql & "	And T21_SYUKKETU_KBN IN (" & C_KETU_KEKKA & "," & C_KETU_TIKOKU & ","& C_KETU_SOTAI &"," & C_KETU_KEKKA_1 & ")"
	w_sSql = w_sSql & " Group By "
	w_sSql = w_sSql & "		T21_CLASS,"
	w_sSql = w_sSql & " 	T21_SYUKKETU_KBN,"
	w_sSql = w_sSql & " 	T21_GAKUSEKI_NO "
	w_sSql = w_sSql & " Order By "
	w_sSql = w_sSql & " 	T21_CLASS, "
	w_sSql = w_sSql & " 	T21_GAKUSEKI_NO "
	
	If gf_GetRecordset(p_oRecordset,w_sSql) <> 0 Then : exit function
	
	gf_GetSyukketuData2 = True
	
End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڕ]���擾
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_sKamokuCD - �ȖڃR�[�h(IN)
'               �i�F��Ȗڂ̏ꍇ�͕��ރR�[�h���w�肷��j
'           p_sKamokuBunrui - �Ȗڕ��ރR�[�h(IN)
'               C_KAMOKUBUNRUI_TUJYO = �ʏ�Ȗ�
'               C_KAMOKUBUNRUI_NINTEI = �F��Ȗ�
'               C_KAMOKUBUNRUI_TOKUBETU = ���ʉȖ�
'           p_iTensu - �_��(IN)
'           p_uData - �]���f�[�^(OUT)
'
' �@�\�ڍׁF�w��̉ȖڃR�[�h�Ɠ_������p_uData�ɕ]���A�]��A���_�Ȗڂ�ݒ肷��
' ���@�@�l�F�_���]���̉ȖڑΏ�
'           call��
'           ret = gf_GetKamokuTensuHyoka(m_iNendo, w_KamokuCD, C_KAMOKUBUNRUI_TUJYO, w_iTensu, w_udata)
'
'           2002.06.19 ����
'*******************************************************************************
Function gf_GetKamokuTensuHyoka(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_iTensu,p_uData)
    Dim w_iZokuseiCD         '�Ȗڑ���
    Dim w_iHyokaNo
    
    On Error Resume Next
    
    gf_GetKamokuTensuHyoka = False
    
    '�Ȗڑ����擾
    If Not gf_GetKamokuZokusei(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,w_iZokuseiCD) Then
        Exit Function
    End If
    
    '�Ȗڑ�������]��NO�擾
    w_iHyokaNo = gf_iGetHyokaNo(w_iZokuseiCD, p_iNendo)
    
    '�]��NO����]���f�[�^�擾
    If Not gf_GetTensuHyoka(p_iNendo,w_iHyokaNo,p_iTensu,p_uData) Then
        Exit Function
    End If
	
    gf_GetKamokuTensuHyoka = True
             
End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڕ]�����X�g�擾
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_sKamokuCD - �ȖڃR�[�h(IN)
'               �i�F��Ȗڂ̏ꍇ�͕��ރR�[�h���w�肷��j
'           p_sKamokuBunrui - �Ȗڕ��ރR�[�h(IN)
'               C_KAMOKUBUNRUI_TUJYO = �ʏ�Ȗ�
'               C_KAMOKUBUNRUI_NINTEI = �F��Ȗ�
'               C_KAMOKUBUNRUI_TOKUBETU = ���ʉȖ�
'           p_lDataCount -  �]���f�[�^����(OUT)
'           p_uData - �]���f�[�^(OUT)
'
' �@�\�ڍׁF�w��̉ȖڃR�[�h�Ɠ_������p_uData()�ɕ]���A�]��A���_�Ȗڂ�ݒ肷��
'           p_uData()�͓��I�z���call���Ő錾���邱�ƁB�i�錾�͊֐����ōs���j
'           p_uData()�̌�����p_lDataCount�ɃZ�b�g�����B�܂��A�z��C���f�b�N�X��
'           1 �` p_lDataCount�܂ł��L���B
' ���@�@�l�F�_���]���̉ȖڑΏ�
'           call��
'           ret = gf_GetKamokuHyokaData(m_iNendo, w_KamokuCD, C_KAMOKUBUNRUI_TUJYO, w_lConut, w_udata())
'
'           2002.06.19 ����
'*******************************************************************************
Function gf_GetKamokuHyokaData(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_lDataCount,p_uData)
    Dim w_iZokuseiCD         '�Ȗڑ���
    Dim w_iHyokaNo
    
    On Error Resume Next
    
    gf_GetKamokuHyokaData = False
    
    '�Ȗڑ����擾
    If Not gf_GetKamokuZokusei(p_iNendo, p_sKamokuCD, p_sKamokuBunrui, w_iZokuseiCD) Then
        Exit Function
    End If

    '�Ȗڑ�������]��NO�擾
    w_iHyokaNo = gf_iGetHyokaNo(w_iZokuseiCD, p_iNendo)
    
    '�]��NO����]���f�[�^�擾
    If Not gf_GetHyokaData(p_iNendo, w_iHyokaNo, p_lDataCount, p_uData) Then
        Exit Function
    End If
	
    gf_GetKamokuHyokaData = True
             
End Function

'*******************************************************************************
' �@�@�@�\�F���ѓ��͕��@�擾
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_sKamokuCD - �ȖڃR�[�h(IN)
'               �i�F��Ȗڂ̏ꍇ�͕��ރR�[�h���w�肷��j
'           p_sKamokuBunrui - �Ȗڕ��ރR�[�h(IN)
'               C_KAMOKUBUNRUI_TUJYO = �ʏ�Ȗ�
'               C_KAMOKUBUNRUI_NINTEI = �F��Ȗ�
'               C_KAMOKUBUNRUI_TOKUBETU = ���ʉȖ�
'           p_iSeiseki - ���ѓ��͕��@(OUT)
'
' �@�\�ڍׁF�ȖڃR�[�h���琬�ѓ��͕��@���擾
' ���@�@�l�F2002.06.19 ����
'*******************************************************************************
Function gf_GetKamokuSeisekiInp(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_iSeiseki)
    Dim w_iZokuseiCD         '�Ȗڑ���
    Dim w_iHyokaNo
    
    On Error Resume Next
    
    gf_GetKamokuSeisekiInp = False

    '�Ȗڑ����擾
    If Not gf_GetKamokuZokusei(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,w_iZokuseiCD) Then
        Exit Function
    End If

    '�Ȗڑ������琬�ѓ��͕��@�擾
    If Not gf_SeisekiInp(w_iZokuseiCD,p_iNendo,p_iSeiseki) Then
        Exit Function
    End If

    gf_GetKamokuSeisekiInp = True
	
End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڑ����R�[�h�擾
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_sKamokuCD - �ȖڃR�[�h(IN)
'               �i�F��Ȗڂ̏ꍇ�͕��ރR�[�h���w�肷��j
'           p_sKamokuBunrui - �Ȗڕ��ރR�[�h(IN)
'               C_KAMOKUBUNRUI_TUJYO = �ʏ�Ȗ�
'               C_KAMOKUBUNRUI_NINTEI = �F��Ȗ�
'               C_KAMOKUBUNRUI_TOKUBETU = ���ʉȖ�
'           p_iZokuseiCD - �����R�[�h(OUT)
'
' �@�\�ڍׁF�w��̉ȖڃR�[�h�̉Ȗڑ������擾����
'           �Ȗڕ��ނɂ�葮���擾�̃}�X�^��؂蕪����
' ���@�@�l�F2002.06.19 ����
'*******************************************************************************
Public Function gf_GetKamokuZokusei(p_iNendo,p_sKamokuCD,p_sKamokuBunrui,p_iZokuseiCD)
    
    gf_GetKamokuZokusei = False
    
    '�Ȗڑ���
    Select Case p_sKamokuBunrui
    '�ʏ�Ȗ�
    Case C_KAMOKUBUNRUI_TUJYO
        '�Ȗ�M���瑮���R�[�h�擾
        If Not f_GetZokuseiCDTujyo(p_iNendo, p_sKamokuCD, p_iZokuseiCD) Then
            Exit Function
        End If
        
    '�F��Ȗ�
    Case C_KAMOKUBUNRUI_NINTEI
        '�F��Ȗ�M���瑮���R�[�h�擾
        If Not f_GetZokuseiCDNintei(p_sKamokuCD, p_iZokuseiCD) Then
            Exit Function
        End If
    
    '���ʊ���
    Case C_KAMOKUBUNRUI_TOKUBETU
        '���ʊ���M���瑮���R�[�h�擾
        If Not f_GetZokuseiCDToku(p_iNendo, p_sKamokuCD, p_iZokuseiCD) Then
            Exit Function
        End If
    
    End Select
    
    gf_GetKamokuZokusei = True
    
End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڕ]���擾
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_iHyokaNo - �]��NO(IN)
'           p_iTensu - �_��(IN)
'           p_uData - �]���f�[�^(OUT)
'
' �@�\�ڍׁF�_������]��NO��p_uData�ɕ]���A�]��A���_�Ȗڂ�ݒ肷��
' ���@�@�l�F�]��NO�����łɕ������Ă���ꍇ�ɂ͒���call����
'           �]��NO��������Ȃ��Ƃ��́Agf_GetKamokuTensuHyoka��call
'           2002.06.19 ����
'*******************************************************************************
Function gf_GetTensuHyoka(p_iNendo,p_iHyokaNo,p_iTensu,p_uData)
    Dim w_oRecord
    Dim w_sSql
    
    ReDim p_uData(3)
    
    On Error Resume Next
    
    gf_GetTensuHyoka = False
	
    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_MEI, "
    w_sSql = w_sSql & " 	M08_HYOTEI, "
    w_sSql = w_sSql & " M08_HYOKA_SYOBUNRUI_RYAKU "
    
    w_sSql = w_sSql & " FROM M08_HYOKAKEISIKI "
    
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & " 	M08_HYOUKA_NO = " & p_iHyokaNo						'�]��NO
    w_sSql = w_sSql & " AND M08_MIN <= " & p_iTensu								'�_��
    w_sSql = w_sSql & " AND M08_MAX >= " & p_iTensu
    w_sSql = w_sSql & " AND M08_NENDO = " & p_iNendo							'�N�x
    w_sSql = w_sSql & " AND M08_HYOKA_TAISYO_KBN = " & C_HYOKA_TAISHO_IPPAN		'��ʊw��
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '�Ȗ�M�Ȃ����G���[
    If w_oRecord.EOF Then
        Exit Function
    End If
    
    '�f�[�^�Z�b�g
    p_uData(0) = w_oRecord("M08_HYOKA_SYOBUNRUI_MEI")		'�]��
    p_uData(1) = w_oRecord("M08_HYOTEI")					'�]��
    p_uData(2) = w_oRecord("M08_HYOKA_SYOBUNRUI_RYAKU")		
    
    Call gf_closeObject(w_oRecord)
    
    gf_GetTensuHyoka = True
	
End Function


'*******************************************************************************
' �@�@�@�\�F�Ȗڕ]���擾
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_iHyokaNo - �]��NO(IN)
'           p_lDataCount - ����(OUT)
'           p_uData - �]���f�[�^(OUT)
'
' �@�\�ڍׁF�_������]��NO��p_uData�ɕ]���A�]��A���_�Ȗڂ�ݒ肷��
' ���@�@�l�F�]��NO�����łɕ������Ă���ꍇ�ɂ͒���call����
'           �]��NO��������Ȃ��Ƃ��́Agf_GetKamokuTensuHyoka��call
'           2002.06.19 ����
'*******************************************************************************
Function gf_GetHyokaData(p_iNendo,p_iHyokaNo,p_lDataCount,p_uData)
    Dim w_oRecord
    Dim w_sSql
    Dim w_lIdx
    
    On Error Resume Next
    
    gf_GetHyokaData = False
    
    p_lDataCount = 0
    
    w_sSql = ""
    w_sSql = w_sSql & " SELECT"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_MEI,"
    w_sSql = w_sSql & " 	M08_HYOTEI,"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_RYAKU"
    
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & " 	M08_HYOKAKEISIKI"
    
    w_sSql = w_sSql & " WHERE"
    w_sSql = w_sSql & " 	M08_HYOUKA_NO = " & p_iHyokaNo      '�]��NO
    w_sSql = w_sSql & " AND M08_NENDO = " & p_iNendo        '�N�x
    w_sSql = w_sSql & " AND M08_HYOKA_TAISYO_KBN = " & C_HYOKA_TAISHO_IPPAN     '��ʊw��
    
    w_sSql = w_sSql & " ORDER BY"
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_CD"
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '�Ȗ�M�Ȃ����G���[
    If w_oRecord.EOF Then
        Exit Function
    End If
    
    p_lDataCount = gf_GetRsCount(w_oRecord)
    
    '�z��f�[�^�錾
    ReDim p_uData(p_lDataCount,3)
    w_lIdx = 0
    
    Do Until w_oRecord.EOF
        
        '�f�[�^�Z�b�g
        p_uData(w_lIdx,0) = w_oRecord("M08_HYOKA_SYOBUNRUI_MEI")	'�]��
        p_uData(w_lIdx,1) = w_oRecord("M08_HYOTEI")					'�]��
        p_uData(w_lIdx,2) = w_oRecord("M08_HYOKA_SYOBUNRUI_RYAKU")	
        
        w_lIdx = w_lIdx + 1
        w_oRecord.MoveNext
    Loop
    
    Call gf_closeObject(w_oRecord)
    
    gf_GetHyokaData = True

End Function

'*******************************************************************************
' �@�@�@�\�F�]���`��No���擾����
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iKamokuZokusei_CD - �Ȗڑ���CD , p_iNendo - �Ώ۔N�x
' �@�\�ڍׁFgf_GetHyokaNo�@-�@�]���`��No
' ���@�@�l�F2002.06.12�@���c
'*******************************************************************************
Function gf_iGetHyokaNo(p_iKamokuZokusei_CD,p_iNendo)
    Dim w_oRecord
    Dim w_sSql
    
    w_sSql = ""
    w_sSql = w_sSql & " Select "
    w_sSql = w_sSql & " 	M100_HYOUKA_NO "
    w_sSql = w_sSql & " From "
    w_sSql = w_sSql & " 	M100_KAMOKU_ZOKUSEI "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & " 	M100_NENDO =" & p_iNendo
    w_sSql = w_sSql & " AND M100_ZOKUSEI_CD =" & p_iKamokuZokusei_CD
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    gf_iGetHyokaNo = CInt(w_oRecord("M100_HYOUKA_NO"))
    
    Call gf_closeObject(w_oRecord)
    
End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڑ����R�[�h�擾(M03_KAMOKU)
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_sKamokuCD - �ȖڃR�[�h(IN)
'           p_iZokuseiCD - �����R�[�h(OUT)
'
' �@�\�ڍׁFM03_KAMOKU����Ȗڑ������擾����
' ���@�@�l�F2002.06.19 ����
'*******************************************************************************
Function f_GetZokuseiCDTujyo(p_iNendo,p_sKamokuCD,p_iZokuseiCD)
    Dim w_oRecord
    Dim w_sSql
    
    On Error Resume Next
    
    f_GetZokuseiCDTujyo = False
    
    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & " 	M03_ZOKUSEI_CD"
    w_sSql = w_sSql & " FROM"
    w_sSql = w_sSql & " 	M03_KAMOKU"
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & " 	M03_NENDO =" & p_iNendo
    w_sSql = w_sSql & " AND M03_KAMOKU_CD = '" & Trim(p_sKamokuCD) & "'"
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '�Ȗ�M�Ȃ����G���[
    If w_oRecord.EOF Then exit function
    
    p_iZokuseiCD = CInt(w_oRecord("M03_ZOKUSEI_CD"))
    
    Call gf_closeObject(w_oRecord)
    
    f_GetZokuseiCDTujyo = True
    
End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڑ����R�[�h�擾(M110_NINTEI_H)
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_sBunruiCD - ���ރR�[�h(IN)
'           p_iZokuseiCD - �����R�[�h(OUT)
'
' �@�\�ڍׁFM110_NINTEI_H����Ȗڑ������擾����
' ���@�@�l�F2002.06.19 ����
'*******************************************************************************
Function f_GetZokuseiCDNintei(p_sBunruiCD,p_iZokuseiCD)
	
    Dim w_oRecord
    Dim w_sSql
    
    On Error Resume Next
    
    f_GetZokuseiCDNintei = False
    
    w_sSql = ""
    w_sSql = w_sSql & vbCrLf & "SELECT "
    w_sSql = w_sSql & vbCrLf & " M110_ZOKUSEI_CD"
    w_sSql = w_sSql & vbCrLf & " FROM"
    w_sSql = w_sSql & vbCrLf & " M110_NINTEI_H"
    w_sSql = w_sSql & vbCrLf & " WHERE "
    w_sSql = w_sSql & vbCrLf & " M110_BUNRUI_CD = '" & Trim(p_sBunruiCD) & "'"
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '�Ȗ�M�Ȃ����G���[
    If w_oRecord.EOF Then exit function
    
    p_iZokuseiCD = CInt(w_oRecord("M110_ZOKUSEI_CD"))
    
    Call gf_closeObject(w_oRecord)
    
    f_GetZokuseiCDNintei = True
	
End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڑ����R�[�h�擾(M41_TOKUKATU)
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iNendo - �N�x(IN)
'        �@ p_sKamokuCD - �ȖڃR�[�h(IN)
'           p_iZokuseiCD - �����R�[�h(OUT)
'
' �@�\�ڍׁFM41_TOKUKATU����Ȗڑ������擾����
' ���@�@�l�F2002.06.19 ����
'*******************************************************************************
Function f_GetZokuseiCDToku(p_iNendo,p_sKamokuCD,p_iZokuseiCD)
	
    Dim w_oRecord
    Dim w_sSql
    
    On Error Resume Next
    
    f_GetZokuseiCDToku = False
    
    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & " 	M41_ZOKUSEI_CD"
    w_sSql = w_sSql & " FROM"
    w_sSql = w_sSql & " 	M41_TOKUKATU"
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & " 	M41_NENDO =" & p_iNendo
    w_sSql = w_sSql & " AND M41_TOKUKATU_CD = '" & Trim(p_sKamokuCD) & "'"
    
	If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '�Ȗ�M�Ȃ����G���[
    If w_oRecord.EOF Then exit function
    
    p_iZokuseiCD = CInt(w_oRecord("M41_ZOKUSEI_CD"))
    
    Call gf_closeObject(w_oRecord)
    
    f_GetZokuseiCDToku = True
    
End Function

'*******************************************************************************
' �@�@�@�\�F�]���`��No���擾����
' �ԁ@�@�l�FTrue/False
' ���@�@���Fp_iKamokuZokusei_CD - �Ȗڑ���CD
'           p_iNendo - �Ώ۔N�x
'           p_iSeiseki - ���ѓ��͕��@
' �@�\�ڍׁF
' ���@�@�l�F2002.06.19 ����
'*******************************************************************************
Function gf_SeisekiInp(p_iKamokuZokusei_CD,p_iNendo,p_iSeiseki)
	Dim w_oRecord
    Dim w_sSql
    
    gf_SeisekiInp = False
    
    On Error Resume Next
    
    w_sSql = ""
    w_sSql = w_sSql & " Select "
    w_sSql = w_sSql & " 	M100_SEISEKI_INP "
    w_sSql = w_sSql & " From "
    w_sSql = w_sSql & " 	M100_KAMOKU_ZOKUSEI "
    w_sSql = w_sSql & " WHERE M100_NENDO =" & p_iNendo
    w_sSql = w_sSql & " 	AND M100_ZOKUSEI_CD =" & p_iKamokuZokusei_CD

    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function

    '�f�[�^�Ȃ��G���[
    If w_oRecord.EOF Then exit function
	
	p_iSeiseki = cInt(w_oRecord("M100_SEISEKI_INP"))
    
    Call gf_closeObject(w_oRecord)
    
    gf_SeisekiInp = True
    
End Function

'********************************************************************************
'*	[�@�\]	�����̃��b�Z�[�W���o�͂���HTML
'*	[����]	p_Msg���G���[���b�Z�[�W
'*			p_Title���^�C�g��
'*	[�ߒl]	
'*	[����]	�G���[���Ɏg�p
'********************************************************************************
Sub gs_showWhitePage(p_Msg,p_Title)
%>
	<html>
	<head>
		<title><%=Server.HTMLEncode(p_Title)%></title>
		<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>
	
	<body LANGUAGE="javascript">
	<form name="frm">
	
	<center>
	<br><br><br>
		<span class="msg"><%=Server.HTMLEncode(p_Msg)%></span>
	</center>
	
	<input type="hidden" name="txtMsg" value="<%=Server.HTMLEncode(p_Msg)%>">
	</form>
	</body>
	</html>
<%
End Sub

'********************************************************************************
'*  [�@�\]  �w�Z�ԍ����o�^����Ă��邩�`�F�b�N����
'*  [����]  p_ChkFlg(out),p_Type(in)��[C_KEKKAGAI_DISP,C_HYOKAYOTEI_DISP,C_DATAKBN_DISP]
'*  [�ߒl]  
'*          gf_ChkDisp(True������I���AFalse���G���[)
'*  [����]  
'*  		�w�Z���Ƃɏ������Ⴄ�ۂɎg�p
'*  		p_ChkFlg��True�Ȃ珈��������
'*  		
'********************************************************************************
function gf_ChkDisp(p_Type,p_ChkFlg)
	Dim w_sSQL
	Dim w_Rs
	
	On Error Resume Next
	Err.Clear
	
	gf_ChkDisp = false
	p_ChkFlg = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "		M00_KANRI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M00_KANRI "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "      M00_NENDO = " & p_Type
	
	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function
	
	if w_Rs.EOF then
		gf_ChkDisp = true
		exit function
	end if
	
	if cint(w_Rs(0)) = C_DISP then p_ChkFlg = true
	
	Call gf_closeObject(w_Rs)
	
	gf_ChkDisp = true
	
end function

'********************************************************************************
'*  [�@�\]  �F��m��O��𒲂ׂ�
'*  [����]  p_iNendo;�N�x[IN]�Ap_bNiteiFlg:true(�F��O)�Afalse(�F���)[OUT]
'*  [�ߒl]  true:����,false:���s
'********************************************************************************
Function gf_GetNintei(p_iNendo,p_bNiteiFlg)
	
	Dim w_sSQL,w_Rs
	
	On Error Resume Next
	Err.Clear
	
	gf_GetNintei = false
	p_bNiteiFlg = false
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " select "
	w_sSQL = w_sSQL & " 	M00_SYUBETU "
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & " 	M00_KANRI "
	w_sSQL = w_sSQL & " where "
	w_sSQL = w_sSQL & " 	M00_NENDO =  " & p_iNendo & " and "
	w_sSQL = w_sSQL & " 	M00_NO = " & C_K_HANTEI_JOUTAI
	
	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function
	
	if w_Rs.EOF then exit function
	
	if cint(w_Rs(0)) = C_K_HANTEI_ATO then
		p_bNiteiFlg = true
	end if
	
	gf_GetNintei = true
	
	Call gf_closeObject(w_Rs)
	
End Function

'********************************************************************************
'*  [�@�\]  �w�Z�ԍ��擾
'*  [����]  
'*  [�ߒl]  p_iGakkoNO:�w�Z�ԍ� ���R�[�h���Ȃ��ꍇ��""��Ԃ�
'*  [����]  true:����,false:���s
'********************************************************************************
function gf_GetGakkoNO(p_iGakkoNO)
	Dim w_sSQL
	Dim w_Rs
	
	On Error Resume Next
	Err.Clear
	
	gf_GetGakkoNO = False
	
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "		M00_KANRI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M00_KANRI "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "     M00_NENDO = " & C_GAKKO_NO
	w_sSQL = w_sSQL & "     AND M00_NO = " & C_DISP_NO

	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

	if w_Rs.EOF then 
		p_iGakkoNO = ""
		gf_GetGakkoNO = True
		exit function
	end if

	p_iGakkoNO = w_Rs("M00_KANRI")
	gf_GetGakkoNO = True
	
end function

'********************************************************************************
'*	[�@�\]	�ٓ����̎擾�֐��i�s���o���Łj
'*	[����]	p_iGakusei_No:�w��NO
'*			p_iNendo		:�����N�x
'*          p_Data          :�Ώۓ��t
'*	[�ߒl]	0:���擾���� 1:���s  p_SSSS : �ٓ�����
'*	[����]	2003.03.14 �ŁF���c
'********************************************************************************
Function gf_Set_IdouGyozi(p_sGakusekiCd,p_iNendo,p_Data,ByRef p_SSSS)

		gf_Set_IdouGyozi = 1

		Dim w_Date
		Dim w_SSSR
		
		w_Date = p_Data 'gf_YYYY_MM_DD(p_iNendo & "/" & month(date()) & "/" & day(date()),"/")
 		'//C_IDO_FUKUGAKU=3:���w�AC_IDO_TEI_KAIJO=5:��w����
		'p_SSSS = ""
		w_SSSR = ""

		p_SSSS = gf_Get_IdouChk(p_sGakusekiCd,w_Date,p_iNendo,w_SSSR)

'response.write w_Date
'response.write w_SSSR
'response.write p_SSSS

		IF CStr(p_SSSS) <> "" Then

			IF Cstr(p_SSSS) <> CStr(C_IDO_FUKUGAKU) AND Cstr(p_SSSS) <> Cstr(C_IDO_TEI_KAIJO) Then

					p_SSSS = w_SSSR

					gf_Set_IdouGyozi =0
			Else

				w_SSSR = ""
				p_SSSS = ""
			
				gf_Set_IdouGyozi = 1

			End if

		End if

'response.write p_SSSS

End Function

'********************************************************************************
'*  [�@�\]  �F��m��O��𒲂ׂ�
'*  [����]  p_iNendo;�N�x[IN]�Ap_bNiteiFlg:true(�F��O)�Afalse(�F���)[OUT]
'*  [�ߒl]  true:����,false:���s
'********************************************************************************
Function gf_GetGakunenNintei(p_iNendo,p_iGakunen,p_bNiteiFlg)

	Dim w_sNinteiCD
	Dim w_iNinFLG

	On Error Resume Next
	Err.Clear

	gf_GetGakunenNintei = false
	p_bNiteiFlg = false

	'�F��R�[�h���擾�iM00_KANRI�j
	if Not gf_GetNinteiCD(p_iNendo,w_sNinteiCD) then
		exit function
	end if

	'�G���[�`�F�b�N
	if w_sNinteiCD = "" then
		exit function
	end if

	'�G���[�`�F�b�N�@�o�C�g���`�F�b�N�i= 5�o�C�g�j
'	if f_LenB(w_sNinteiCD) = 5 then
'		exit function
'	end if

	'�w�N�̔F��FLG���擾����
    w_iNinFLG = Mid(w_sNinteiCD, p_iGakunen, 1)

	if Not IsNumeric(w_iNinFLG) then
		exit function
	end if

	'�F����True��Ԃ�
	if cint(w_iNinFLG) = C_K_HANTEI_ATO then
		p_bNiteiFlg = true
	end if

	gf_GetGakunenNintei = true

End Function

'********************************************************************************
'*  [�@�\]  �F��R�[�h���擾
'*  [����]  p_iNendo;�N�x[IN]
'*  [�ߒl]  true:����,false:���s
'********************************************************************************
Function gf_GetNinteiCD(p_iNendo, Byref p_sNinteiCD) 
	Dim w_sSQL,w_Rs

	gf_GetNinteiCD = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M00_KANRI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M00_KANRI "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M00_NENDO = " & p_iNendo & " AND "
	w_sSQL = w_sSQL & " 	M00_NO    = " & C_K_HANTEI_JOUTAI

	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

	if w_Rs.EOF then exit function

	p_sNinteiCD = w_Rs(0)

	Call gf_closeObject(w_Rs)

	gf_GetNinteiCD = true

End Function

%>