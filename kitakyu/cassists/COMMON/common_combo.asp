<%
'/******************************************************************************
' �V�X�e�����F�L�����p�X�A�V�X�g
' ���@���@���F���ʏ���
' ��۸���ID�F-
' �@�@�@�@�\�F�R���{�{�b�N�X���ʉ��֐�
'-------------------------------------------------------------------------------
' ��@�@�@���F2001.02.26 ���u�@�m��
' �ρ@�@�@�X�F2001.03.02 �c���@��K     �Ή��e�[�u���ǉ�
' �@�@�@�@�@�@2001.03.05 �c���@��K     �Ή��e�[�u���ǉ�
' �@�@�@�@�@�@2001.03.06 �c���@��K     �w�N�p�֐��ǉ�
' �@�@�@�@�@�@2001.03.26 ���c�@���@     �������ǉ�
' �@�@�@�@�@�@2001.05.14 ���c�@���@     �ψ���ʒǉ�
'             2001.07.05 ���{�@�����@   �������̕\���ǉ�
'             2001.07.12 �≺�@�K��Y   �Ή��e�[�u���ǉ�
'******************************************************************************/

'ComboBox�ɃZ�b�g���鎞�̃����O�X���Œ蒷�ɂ���
Private Const mC_MST_NAMELEN = 50       ''�}�X�^�̖��̃f�[�^�����O�X
Private Const mC_MAX_IDNUMBER = 999999  ''�ő��ID�ԍ��i���O�e�[�u���j

'ComboBox�pnull��p�l
Public Const C_CBO_NULL = "@@@"       ''�}�X�^�̖��̃f�[�^�����O�X

'�e�[�u��ID�p�̌Œ�l
Public Const C_CBO_M01_KUBUN = 1
Public Const C_CBO_M01_KUBUN_R = 101
Public Const C_CBO_M02_GAKKA = 2
Public Const C_CBO_M02_GAKKA_R = 102
Public Const C_CBO_M05_CLASS = 5
Public Const C_CBO_M05_CLASS_R = 105
Public Const C_CBO_M05_CLASS_N = 115
Public Const C_CBO_M05_CLASS_G = 125
Public Const C_CBO_M11_KYOKANSHITU = 11
Public Const C_CBO_M20_COURSE = 20
Public Const C_CBO_M12_SITYOSON = 12
Public Const C_CBO_M16_KEN = 16
Public Const C_CBO_M17_BUKATUDO = 17
Public Const C_CBO_M34_IIN = 34
Public Const C_CBO_M29_KYOKAN_YOTEI = 29
'Public Const C_CBO_M40_CALENDER = 40
'Public Const C_CBO_M40_CALENDER_H = 401
Public Const C_CBO_M27_SIKEN = 27
Public Const C_CBO_M04_KYOKAN = 04
Public Const C_CBO_T15_RISYU = 15
Public Const C_CBO_T15_RISYU_GRP = 150
Public Const C_CBO_T11_GAKUSEKI_N = 111
Public Const C_CBO_T16_R_KOJIN_S = 116
Public Const C_CBO_M06_KYOSITU = 6
Public Const C_CBO_T18_SEL_SYUBETU = 18
Public Const C_CBO_M42_JIKAN = 42
Public Const C_CBO_T26_SIKEN_JIK_KAMOKU = 126

Public Const C_CBO_T32_GYOJI_M = 32

Public Function gf_ComboSet(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' �@    �\:ComboBox�Z�b�g
' ��    �l:OK=True/NG=False
' ��    ��:p_oCombo - ComboBox
'          p_sTableName - �e�[�u����
'          p_sWhere - Where����(WHERE��͗v��Ȃ�)
'          p_sSelectOption - <SELECT>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
'          p_bWhite - �擪�ɋ󔒂����邩
'          p_sSelectCD - �W���I�����������R�[�h(""�Ȃ�I���Ȃ�)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��    �l:(���g�p��j
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'��L�̏o��HTML�i�󔒍s��VALUE�́f�������f�ƂȂ�܂��B
'<select name='cboTest' ><Option Value='@@@'>�@�@�@ <Option Value='1'>�x�w(�a�C�E����) <Option Value='2'>�x�w(�o�ϓI��) 
'<Option Value='3'>���w <Option Value='4'>��w </select>
'*************************************************************************************
    Dim w_sId           'ID�t�B�[���h��
    Dim w_sName         '���̃t�B�[���h��
    Dim w_sTableName    '���̃e�[�u����
    Dim w_rst

    gf_ComboSet = False
    do 
    ''�}�X�^����SELECT����t�B�[���h�����擾
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If
	
    ''�}�X�^SELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If

'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '�󔒂�Option�̑��
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">�@�@�@�@�@ "& Chr(13)
    End If

    ''EOF�łȂ���΁A�f�[�^���Z�b�g
    If Not w_rst.EOF Then
        Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If
    Response.write("</select>" & chr(13))

    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
   

    gf_ComboSet = True
    Exit Do
    Loop
End Function

Public Function gf_CalComboSet(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD,p_sDisp)
'*************************************************************************************
' �@    �\:ComboBox�Z�b�g(���tVer)
' ��    �l:OK=True/NG=False
' ��    ��:p_oCombo - ComboBox
'          p_sTableName - �e�[�u����
'          p_sWhere - Where����(WHERE��͗v��Ȃ�)
'          p_sSelectOption - <SELECT>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
'          p_bWhite - �擪�ɋ󔒂����邩
'          p_sSelectCD - �W���I�����������R�[�h(""�Ȃ�I���Ȃ�)
'          p_sDISP - �\�����@(""�Ȃ�yyyy/mm/dd�ŕ\�L)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��    �l:(���g�p��j
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'��L�̏o��HTML�i�󔒍s��VALUE�́f�������f�ƂȂ�܂��B
'<select name='cboTest' ><Option Value='@@@'>�@�@�@ <Option Value='1'>�x�w(�a�C�E����) <Option Value='2'>�x�w(�o�ϓI��) 
'<Option Value='3'>���w <Option Value='4'>��w </select>
'*************************************************************************************
    Dim w_sId           'ID�t�B�[���h��
    Dim w_sName         '���̃t�B�[���h��
    Dim w_sTableName    '���̃e�[�u����
    Dim w_rst

    gf_CalComboSet = False
    if p_sDisp = "" then p_sDisp = "yyyy/mm/dd"
    do 
    ''�}�X�^����SELECT����t�B�[���h�����擾
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If
        
    ''�}�X�^SELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '�󔒂�Option�̑��
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">�@�@�@ "& Chr(13)
    End If

    ''EOF�łȂ���΁A�f�[�^���Z�b�g
    If Not w_rst.EOF Then
        Call s_CalMstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD,p_sDisp)
    End If
    Response.write("</select>" & chr(13))
    
    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
    
    
    gf_CalComboSet = True
    Exit Do
    Loop
End Function

Public Function gf_PluComboSet(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' �@    �\:ComboBox�Z�b�g
' ��    �l:OK=True/NG=False
' ��    ��:p_oCombo - ComboBox
'          p_sTableName - �e�[�u����
'          p_sWhere - Where����(WHERE��͗v��Ȃ�)
'          p_sSelectOption - <SELECT>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
'          p_bWhite - �擪�ɋ󔒂����邩
'          p_sSelectCD - �W���I�����������R�[�h(""�Ȃ�I���Ȃ�)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��    �l:(���g�p��j
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'��L�̏o��HTML�i�󔒍s��VALUE�́f�������f�ƂȂ�܂��B
'<select name='cboTest' ><Option Value='@@@'>�@�@�@ <Option Value='1'>�x�w(�a�C�E����) <Option Value='2'>�x�w(�o�ϓI��) 
'<Option Value='3'>���w <Option Value='4'>��w </select>
'*************************************************************************************
    Dim w_sId           'ID�t�B�[���h��
    Dim w_sName         '���̃t�B�[���h��
    Dim w_sTableName    '���̃e�[�u����
    Dim w_rst

    gf_PluComboSet = False
    do 
    ''�}�X�^����SELECT����t�B�[���h�����擾
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        response.write "ERR1"
        Exit Do
    End If

    ''�}�X�^SELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        response.write "ERR2"
        Exit Do
    End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '�󔒂�Option�̑��
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">�@�@�@ "& Chr(13)
    End If

    ''EOF�łȂ���΁A�f�[�^���Z�b�g
    If Not w_rst.EOF Then
        Call s_PluMstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If
    Response.write("</select>" & chr(13))

    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
   

    gf_PluComboSet = True
    Exit Do
    Loop
End Function


Public Function gf_ComboSetHead(p_sCombo, p_sSelectOption)
'*************************************************************************************
' �@    �\:ComboBox�Z�b�g(�w�b�_�[)
' ��    �l:OK=True/NG=False
' ��    ��:p_oCombo - ComboBox
'          p_sSelectOption - <SELECT>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��    �l:
'*************************************************************************************
    Dim w_sId           'ID�t�B�[���h��
    Dim w_sName         '���̃t�B�[���h��
    Dim w_sTableName    '���̃e�[�u����
    Dim w_rst
    
    gf_ComboSetHead = False
    
    do 
        Response.write(chr(13)&"<select name='" & p_sCombo & "' " & p_sSelectOption & ">" & Chr(13))
        gf_ComboSetHead = True
    Exit Do
    Loop
End Function



Public Function gf_ComboSetData(p_iTableID, p_sWhere ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' �@    �\:ComboBox�Z�b�g(�f�[�^)
' ��    �l:OK=True/NG=False
' ��    ��:p_oCombo - ComboBox
'          p_sTableName - �e�[�u����
'          p_sWhere - Where����(WHERE��͗v��Ȃ�)
'          p_sSelectOption - <SELECT>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
'          p_bWhite - �擪�ɋ󔒂����邩
'          p_sSelectCD - �W���I�����������R�[�h(""�Ȃ�I���Ȃ�)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��    �l:(���g�p��j
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'��L�̏o��HTML�i�󔒍s��VALUE�́f�������f�ƂȂ�܂��B
'<select name='cboTest' ><Option Value='@@@'>�@�@�@ <Option Value='1'>�x�w(�a�C�E����) <Option Value='2'>�x�w(�o�ϓI��) 
'<Option Value='3'>���w <Option Value='4'>��w </select>
'*************************************************************************************
    Dim w_sId           'ID�t�B�[���h��
    Dim w_sName         '���̃t�B�[���h��
    Dim w_sTableName    '���̃e�[�u����
    Dim w_rst
    
    gf_ComboSetData = False
    
    do 
    ''�}�X�^����SELECT����t�B�[���h�����擾
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If
        
    ''�}�X�^SELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------

    '�󔒂�Option�̑��
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">�@�@�@" & Chr(13)
    End If

    ''EOF�łȂ���΁A�f�[�^���Z�b�g
    If Not w_rst.EOF Then
        Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If
    
    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
    
    
    gf_ComboSetData = True
    Exit Do
    Loop
End Function

Public Function gf_ComboSetFoot()
'*************************************************************************************
' �@    �\:ComboBox�Z�b�g(�t�b�_�[)
' ��    �l:OK=True/NG=False
' ��    ��:
' �@�\�ڍ�:
' ��    �l:
'*************************************************************************************
    
    gf_ComboSetFoot = False
    
    do 
        Response.write("</select>"&Chr(13))
        gf_ComboSetFoot = True
    Exit Do
    Loop
End Function










Private Function f_MstFieldName(p_iTableID, p_sID , p_sName , p_sTableName)
'*************************************************************************************
' �@    �\: �}�X�^�t�B�[���h���Z�b�g
' ��    �l: OK=True/NG=False
' ��    ��:  p_iTableID
'            p_sId - ID�t�B�[���h��(O)
'            p_sName - ���̃t�B�[���h��(O)
'            p_sTableName - �e�[�u����(I)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����iCombo�Z�b�g�p�j
' ��    �l:
'*************************************************************************************

    f_MstFieldName = True
    
    ''�e�[�u������ID�Ɩ��̂̃t�B�[���h�����Z�b�g
    Select Case p_iTableID
    Case C_CBO_M01_KUBUN            '�敪�}�X�^(���͕̂��ʂ�)
        p_sTableName = "M01_Kubun"
        p_sID = "M01_SYOBUNRUI_CD"
        p_sName = "M01_SYOBUNRUIMEI"
    Case C_CBO_M01_KUBUN_R          '�敪�}�X�^(���̂͗���)
        p_sTableName = "M01_Kubun"
        p_sID = "M01_SYOBUNRUI_CD"
        p_sName = "M01_SYOBUNRUIMEI_R"
    Case C_CBO_M02_GAKKA            '�w�ȃ}�X�^
        p_sTableName = "M02_GAKKA"
        p_sID = "M02_GAKKA_CD"
        p_sName = "M02_GAKKAMEI"
    Case C_CBO_M02_GAKKA_R          '�w�ȃ}�X�^����
        p_sTableName = "M02_GAKKA"
        p_sID = "M02_GAKKA_CD"
        p_sName = "M02_GAKKARYAKSYO"
    Case C_CBO_M05_CLASS                '�N���X�}�X�^
        p_sTableName = "M05_CLASS"
        p_sID = "M05_CLASSNO"
        p_sName = "M05_CLASSMEI"
    Case C_CBO_M05_CLASS_R          '�N���X�}�X�^(����)
        p_sTableName = "M05_CLASS"
        p_sID = "M05_CLASSNO"
        p_sName = "M05_CLASSRYAKU"
    Case C_CBO_M05_CLASS_N          '�N���X�}�X�^(�����N�x)
        p_sTableName = "M05_CLASS"
        p_sID = "M05_NENDO"
        p_sName = "M05_NENDO"
    Case C_CBO_M05_CLASS_G          '�N���X�}�X�^(�w�N)
        p_sTableName = "M05_CLASS"
        p_sID = "M05_GAKUNEN"
        p_sName = "M05_GAKUNEN"
    Case C_CBO_M20_COURSE           '�R�[�X�}�X�^
        p_sTableName = "M20_COURSE"
        p_sID = "M20_COURSE_CD"
        p_sName = "M20_COURSEMEI"
    Case C_CBO_M11_KYOKANSHITU         '�������}�X�^
        p_sTableName = "M11_KYOKANSITU"
        p_sID = "M11_KYOKANSITU_CD"
        p_sName = "M11_KYOKANSITUMEI"
    Case C_CBO_M12_SITYOSON         '�s�����}�X�^
        p_sTableName = "M12_SITYOSON"
        p_sID = "M12_SITYOSON_CD"
        p_sName = "M12_SITYOSONMEI"
    Case C_CBO_M16_KEN              '���}�X�^
        p_sTableName = "M16_KEN"
        p_sID = "M16_KEN_CD"
        p_sName = "M16_KENMEI"
    Case C_CBO_M17_BUKATUDO              '�������}�X�^
        p_sTableName = "M17_BUKATUDO"
        p_sID = "M17_BUKATUDO_CD"
        p_sName = "M17_BUKATUDOMEI"
    Case C_CBO_M34_IIN                  '�ψ��}�X�^
        p_sTableName = "M34_IIN"
        p_sID = "M34_DAIBUN_CD"
        p_sName = "M34_IIN_NAME"
    Case C_CBO_M29_KYOKAN_YOTEI                  '�����\��}�X�^
        p_sTableName = "M29_KYOKAN_YOTEI"
        p_sID = "M29_YOTEI_CD"
        p_sName = "M29_YOTEIMEI"
'    Case C_CBO_M40_CALENDER                  '�J�����_�}�X�^
'        p_sTableName = "M40_CALENDER"
'        p_sID = "M40_DATE"
'        p_sName = "M40_DATE"
'    Case C_CBO_M40_CALENDER_H                '�J�����_�}�X�^(����\��)
'        p_sTableName = "M40_CALENDER"
'        p_sID = "M40_DATE"
'        p_sName = "M40_HI"
    Case C_CBO_T32_GYOJI_M    '�s������(�J�����_�[�p�Awhere���"GROUP BY T32_HIDUKE"������K�v������܂�)
        p_sTableName = "T32_GYOJI_M"
        p_sID = "T32_HIDUKE"
        p_sName = "T32_HIDUKE"
    Case C_CBO_M27_SIKEN                  '�����}�X�^
        p_sTableName = "M27_SIKEN"
        p_sID = "M27_SIKEN_CD"
        p_sName = "M27_SIKENMEI"
'    Case C_CBO_M04_KYOKAN                  '�����}�X�^
'        p_sTableName = "M04_KYOKAN"
'        p_sID = "M04_KYOKAN_CD"
'        p_sName = "M04_KYOKANMEI_SEI"
    Case C_CBO_M04_KYOKAN                  '�����}�X�^�i�������̕\���j
        p_sTableName = "M04_KYOKAN"
        p_sID = "M04_KYOKAN_CD"
        p_sName = "M04_KYOKANMEI_SEI"
        p_sName = p_sName & " , M04_KYOKANMEI_MEI"
        'p_sName = p_sName & " , M04_KYOKANMEI_KANA_SEI"
        'p_sName = p_sName & " , M04_KYOKANMEI_KANA_MEI"
    Case C_CBO_T15_RISYU                  '���C�e�[�u��
        p_sTableName = "T15_RISYU"
        p_sID = "T15_KAMOKU_CD"
        p_sName = "T15_KAMOKUMEI"
    Case C_CBO_T15_RISYU_GRP                  '���C�e�[�u��
        p_sTableName = "T15_RISYU"
        p_sID = "T15_KAMOKU_CD"
        p_sName = "T15_KAMOKUMEI"
    Case C_CBO_T11_GAKUSEKI_N                  '�w�����i�������̕\���j
        p_sTableName = "T11_GAKUSEKI T11"
        p_sTableName = p_sTableName & " , T13_GAKU_NEN T13"
        p_sID = "T11_GAKUSEI_NO"
        p_sName = "T13_GAKUSEKI_NO"
        p_sName = p_sName & " , T11_SIMEI"
    Case C_CBO_T16_R_KOJIN_S
        p_sTableName = "T16_RISYU_KOJIN"
        p_sID = "T16_GRP"
        p_sName = "T16_SYUBETU_MEI"
    Case C_CBO_M06_KYOSITU
        p_sTableName = "M06_KYOSITU"
        p_sID = "M06_KYOSITU_CD"
        p_sName = "M06_KYOSITUMEI"
    Case C_CBO_T18_SEL_SYUBETU
        p_sTableName = "T18_SELECTSYUBETU"
        p_sTableName = p_sTableName & " , T15_RISYU "
        p_sID = "T18_GRP"
        p_sName = "T18_SYUBETU_MEI"
    Case C_CBO_M42_JIKAN
        p_sTableName = "M42_SIKEN_PATTERN"
        p_sID = "M42_SIKEN_JIKAN"
        p_sName = "M42_SIKEN_JIKAN"
    Case C_CBO_T26_SIKEN_JIK_KAMOKU                  '�������Ԋ��Ȗږ��i�������̕\���j
        p_sTableName = "T26_SIKEN_JIKANWARI"
        p_sTableName = p_sTableName & " , M03_KAMOKU "
        p_sID = "T26_KAMOKU"
        p_sName = "M03_KAMOKUMEI"
'�ȉ�����Ȃ����̂̓e�[�u�����ɒǉ�����
'    Case C_CBO_M01_KUBUN            '�敪�}�X�^(���͕̂��ʂ�)
'        p_sTableName = "M01_Kubun"
'        p_sID = "M01_SYOBUNRUI_CD"
'        p_sName = "M01_SYOBUNRUIMEI"

    Case Else
        'TODO20010226T.TakaokaCall gs_criticalMsg(4101)        '"�}�X�^�w�肪����܂���"
        f_MstFieldName = False
    End Select
    
    

End Function


Private Function f_MstSelect(p_rst, p_sID , _
                             p_sName , p_sTableName , _
                             p_sWhere )
'*************************************************************************************
' �@    �\:�}�X�^SELECT
' ��    �l:OK=True/NG=False
' ��    ��:p_rst - Recordset
'            p_sId - ID�t�B�[���h��
'            p_sName - ���̃t�B�[���h��
'            p_sTableName - �e�[�u����
'            p_sWhere - WHERE����
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����iCombo�Z�b�g�p�j
' ��    �l:
'*************************************************************************************

    Dim w_sSQL
    
    f_MstSelect = False
    
    ''SQL�쐬
    w_sSQL = "SELECT "
    if p_sID <> p_sName Then
     w_sSQL = w_sSQL & vbCrLf & p_sID & ", "
    end if

    w_sSQL = w_sSQL & vbCrLf & p_sName & " "
    w_sSQL = w_sSQL & vbCrLf & "FROM " & p_sTableName
    If p_sWhere <> "" Then
        w_sSQL = w_sSQL & vbCrLf & " WHERE " & p_sWhere
    End If

'Response.write w_sSql

'------------20010810 tani
'�����ł����A�u�w�Дԍ��@�����v�̃R���{�ɂȂ�Ƃ��́A
'�w�Дԍ��Ń\�[�g���܂��B
    If p_sName = "T13_GAKUSEKI_NO , T11_SIMEI" Then
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T13_GAKUSEKI_NO"
    Else
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY " & p_sID
    End If
'--------------
'�߂��Ƃ��́A��̕��������āA���̃R�����g���͂����B
'        w_sSQL = w_sSQL & vbCrLf & " ORDER BY " & p_sID



'----------------------------------------------------------
'response.write w_sSQL & "<BR>"	 '(�f�o�b�O)
'response.write "<< �f�o�b�O���E�E�Emochi >>"
'----------------------------------------------------------

    ''SELECT
    If gf_GetRecordset(p_rst, w_sSQL) <> 0 Then
        Exit Function

    End If

    f_MstSelect = True
    Exit Function

End Function


Private Sub s_MstDataSet(p_sCombo, p_rst, p_sID, p_sName,p_sSelectCD)
'*************************************************************************************
' �@    �\:�}�X�^�f�[�^�Z�b�g
' ��    �l:
' ��    ��:p_oCombo - ComboBox
'            p_rst - RecordSet
'            p_sId - ID�t�B�[���h��
'            p_sName - ���̃t�B�[���h��
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����iCombo�Z�b�g�p�j
' ��    �l:
'*************************************************************************************
    Dim w_sNameData
    Dim w_sData
    
    ''���R�[�h�Z�b�g=EOF�܂Ń��[�v
    Do Until p_rst.EOF
        '''Combo�ɃZ�b�g����f�[�^�ҏW
        'w_sNameData = gf_GetFieldValue(p_rst, p_sName)
        'w_sData = w_sNameData & Space(mC_MST_NAMELEN - len(w_sNameData))
        w_sData = gf_GetFieldValue(p_rst, p_sID)

        response.write(" <Option Value='" & w_sData & "'")
		if Not gf_IsNull(p_sSelectCD) then
	        If CStr(p_sSelectCD) = CStr(w_sData) Then
	            response.write " Selected "
	        End If
		End if
        response.Write(">" & gf_GetFieldValue(p_rst, p_sName) & Chr(13))
        p_rst.MoveNext      '''�����R�[�h��
        
    Loop

End Sub

Private Sub s_calMstDataSet(p_sCombo, p_rst, p_sID, p_sName,p_sSelectCD,p_sDisp)
'*************************************************************************************
' �@    �\:�J�����_�}�X�^�f�[�^�Z�b�g
' ��    �l:
' ��    ��:p_oCombo - ComboBox
'            p_rst - RecordSet
'            p_sId - ID�t�B�[���h��
'            p_sName - ���̃t�B�[���h��
'            p_sDisp - �\�L���@
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����iCombo�Z�b�g�p�j
' ��    �l:
'*************************************************************************************
    Dim w_sNameData
    Dim w_sData
    Dim w_sDisp
    Dim w_DE
    
    w_DE = array("","sun","mon","tue","wed","thu","fri","sat")
    ''���R�[�h�Z�b�g=EOF�܂Ń��[�v
    Do Until p_rst.EOF
        '''Combo�ɃZ�b�g����f�[�^�ҏW
        'w_sNameData = gf_GetFieldValue(p_rst, p_sName)
        'w_sData = w_sNameData & Space(mC_MST_NAMELEN - len(w_sNameData))
        w_sData = gf_GetFieldValue(p_rst, p_sID)
        
        response.write(" <Option Value='" & w_sData & "'")
        If CStr(p_sSelectCD) = CStr(w_sData) Then
            response.write " Selected "
        End If
    w_sDisp = p_sDisp
    w_sNameData = gf_GetFieldValue(p_rst, p_sName)
    w_sDisp = Replace(w_sDisp,"yyyy",year(w_sNameData))
    w_sDisp = Replace(w_sDisp,"yy",right(year(w_sNameData),2))
    w_sDisp = Replace(w_sDisp,"mm",Month(w_sNameData))
    w_sDisp = Replace(w_sDisp,"dd",day(w_sNameData))
    w_sDisp = Replace(w_sDisp,"DJ",WeekDayName(Weekday(w_sNameData),true))
    w_sDisp = Replace(w_sDisp,"DE",w_DE(Weekday(w_sNameData)))
        response.Write(">" & w_sDisp & Chr(13))
        p_rst.MoveNext      '''�����R�[�h��
        
    Loop

End Sub

Private Sub s_PluMstDataSet(p_sCombo, p_rst, p_sID, p_sName,p_sSelectCD)
'*************************************************************************************
' �@    �\:�}�X�^�f�[�^�Z�b�g
' ��    �l:
' ��    ��:p_oCombo - ComboBox
'            p_rst - RecordSet
'            p_sId - ID�t�B�[���h��
'            p_sName - ���̃t�B�[���h��
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƕ������̂�SELECT����iCombo�Z�b�g�p�j
' ��    �l:
'*************************************************************************************
    Dim w_sNameData
    Dim w_sData
    Dim w_sField
    
    ''���R�[�h�Z�b�g=EOF�܂Ń��[�v
    Do Until p_rst.EOF
        '''Combo�ɃZ�b�g����f�[�^�ҏW
        'w_sNameData = gf_GetFieldValue(p_rst, p_sName)
        'w_sData = w_sNameData & Space(mC_MST_NAMELEN - len(w_sNameData))
        w_sData = gf_GetFieldValue(p_rst, p_sID)
        
        w_sNameData = Split(p_sName,",")
        response.write(" <Option Value='" & w_sData & "'")
        If CStr(p_sSelectCD) = CStr(w_sData) Then
            response.write " Selected "
        End If
        'response.Write(">" & gf_GetFieldValue(p_rst, p_sName) & Chr(13))
        response.Write(">")
        
        For Each w_sField In w_sNameData
            response.Write(gf_GetFieldValue(p_rst, Trim(w_sField)))
            response.Write("�@")
        Next
        
        response.Write(Chr(13))
        p_rst.MoveNext      '''�����R�[�h��
        
    Loop

End Sub




Public Function gf_ComboSet_99(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' �@    �\:ComboBox�Z�b�g
' ��    �l:OK=True/NG=False
' ��    ��:p_oCombo - ComboBox
'          p_sTableName - �e�[�u����
'          p_sWhere - Where����(WHERE��͗v��Ȃ�)
'          p_sSelectOption - <SELECT>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
'          p_bWhite - �擪�ɋ󔒂����邩
'          p_sSelectCD - �W���I�����������R�[�h(""�Ȃ�I���Ȃ�)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��    �l:(���g�p��j
'Call gf_ComboSet_99("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'��L�̏o��HTML�i�󔒍s��VALUE�́f�������f�ƂȂ�܂��B
'<select name='cboTest' ><Option Value='@@@'>�@�@�@ <Option Value='1'>�x�w(�a�C�E����) <Option Value='2'>�x�w(�o�ϓI��) 
'<Option Value='3'>���w <Option Value='4'>��w </select>
'*************************************************************************************
    Dim w_sId           'ID�t�B�[���h��
    Dim w_sName         '���̃t�B�[���h��
    Dim w_sTableName    '���̃e�[�u����
    Dim w_rst

    gf_ComboSet_99 = False
    do 
    ''�}�X�^����SELECT����t�B�[���h�����擾
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If

    ''�}�X�^SELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If
    
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '�󔒂�Option�̑��
    If p_bWhite Then
        response.Write " <Option Value="& 99 &">�@�@�@ "& Chr(13)
    End If

    ''EOF�łȂ���΁A�f�[�^���Z�b�g
    If Not w_rst.EOF Then
        Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If
    Response.write("</select>" & chr(13))

    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
   

    gf_ComboSet_99 = True
    Exit Do
    Loop
End Function



Public Function gf_cboNull(p_sStr)
'*************************************************************************************
' �@    �\:ComboBox����null��p�l�̕ϊ�
' ��    �l:�ϊ���̕�����
' ��    ��:�ϊ�������������
' �@�\�ڍ�:
' ��    �l:
'*************************************************************************************
    
    gf_cboNull = p_sStr
    
    if p_sStr = C_CBO_NULL then gf_cboNull = ""

End Function

Public Function gf_ComboSet_Gakka(p_sCombo, p_iTableID, p_sWhere , p_sSelectOption ,p_bWhite ,p_sSelectCD)
'*************************************************************************************
' �@    �\:ComboBox�Z�b�g
' ��    �l:OK=True/NG=False
' ��    ��:p_oCombo - ComboBox
'          p_sTableName - �e�[�u����
'          p_sWhere - Where����(WHERE��͗v��Ȃ�)
'          p_sSelectOption - <SELECT>�^�O�ɂ���I�v�V����( onchange = 'a_change()' )�Ȃ�
'          p_bWhite - �擪�ɋ󔒂����邩
'          p_sSelectCD - �W���I�����������R�[�h(""�Ȃ�I���Ȃ�)
' �@�\�ڍ�:�w�肳�ꂽ�e�[�u������A���ނƖ��̂�SELECT����ComboBox�ɃZ�b�g����
' ��    �l:(���g�p��j
'Call gf_ComboSet("cboTest",C_CBO_M01_KUBUN," M01_DAIBUNRUI_CD=9 ","",True,"@@@")
'��L�̏o��HTML�i�󔒍s��VALUE�́f�������f�ƂȂ�܂��B
'<select name='cboTest' ><Option Value='@@@'>�@�@�@ <Option Value='1'>�x�w(�a�C�E����) <Option Value='2'>�x�w(�o�ϓI��) 
'<Option Value='3'>���w <Option Value='4'>��w </select>
'*************************************************************************************
    Dim w_sId           'ID�t�B�[���h��
    Dim w_sName         '���̃t�B�[���h��
    Dim w_sTableName    '���̃e�[�u����
    Dim w_rst

    gf_ComboSet_Gakka = False
    do 
    ''�}�X�^����SELECT����t�B�[���h�����擾
    If f_MstFieldName(p_iTableID, w_sId, w_sName, w_sTableName) = False Then
        Exit Do
    End If

    ''�}�X�^SELECT
    If f_MstSelect(w_rst, w_sId, w_sName, w_sTableName, p_sWhere) = False Then
        Exit Do
    End If
'-------------2001/08/10 tani
If w_rst.EOF then p_sSelectOption = " DISABLED " & p_sSelectOption
'--------------
    Response.write(chr(13) & "<select name='" & p_sCombo & "' " & p_sSelectOption & ">") & Chr(13)

    '�󔒂�Option�̑��
    If p_bWhite Then
        response.Write " <Option Value="&C_CBO_NULL&">�@�@�@�@�@ "& Chr(13)
    End If

    ''EOF�łȂ���΁A�f�[�^���Z�b�g
    If Not w_rst.EOF Then
        Call s_MstDataSet(p_sCombo, w_rst, w_sId, w_sName,p_sSelectCD)
    End If

    response.write(" <Option Value='" & C_CLASS_ALL & "'")
    If CStr(p_sSelectCD) = CStr(C_CLASS_ALL) Then
        response.write " Selected "
    End If
    response.Write(">" & "�S�w��" & Chr(13))

    Response.write("</select>" & chr(13))

    If Not w_rst Is Nothing Then
        w_rst.Close
        Set w_rst = Nothing
    End If
   
    gf_ComboSet_Gakka = True
    Exit Do
    Loop
End Function

%>