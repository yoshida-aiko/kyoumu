<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �s���o������
' ��۸���ID : kks/kks0140/kks0140_bottom.asp
' �@      �\: ���y�[�W �s���o�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��: NENDO     '//�N�x
'             KYOKAN_CD '//����CD
'             GAKUNEN   '//�w�N
'             CLASSNO   '//�N���XNO
'             GYOJI_CD  '//�s��CD
'             GYOJI_MEI '//�s����
'             KAISI_BI  '//�J�n��
'             SYURYO_BI '//�I����
'             SOJIKANSU '//�����Ԑ�
' ��      ��:
' ��      �n: NENDO     '//�N�x
'             KYOKAN_CD '//����CD
'             GAKUNEN   '//�w�N
'             CLASSNO   '//�N���XNO
'             GYOJI_CD  '//�s��CD
'             GYOJI_MEI '//�s����
'             KAISI_BI  '//�J�n��
'             SYURYO_BI '//�I����
'             SOJIKANSU '//�����Ԑ�
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��s���o�����͂�\��
'           ���\���{�^���N���b�N��
'               �w�肵�������ɂ��Ȃ����w�Z��\��������
'           ���o�^�{�^���N���b�N��
'               ���͏���o�^����
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
    Const DebugPrint = 0
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    '�擾�����f�[�^�����ϐ�
    Public m_iSyoriNen      '//�����N�x
    Public m_iKyokanCd      '//����CD
    Public m_sGakunen       '//�w�N
    Public m_sClassNo       '//�׽NO
    Public m_sTuki          '//��
    Public m_sGyoji_Cd      '//�s��CD
    Public m_sGyoji_Mei     '//�s����
    Public m_sKaisi_Bi      '//�J�n��
    Public m_sSyuryo_Bi     '//�I����
    Public m_sSoJikan       '//�����Ԑ�
	Public m_sEndDay		'//���͂ł��Ȃ��Ȃ��

    '//ں��ރZ�b�g
    Public m_Rs_M           '//recordset���׏��
    Public m_Rs_G           '//recordset�s���o�����
    Public m_iRsCnt         '//�w�b�_ں��ސ�

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
    w_sMsgTitle="�s���o������"
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

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))


        '//�ϐ�������
        Call s_ClearParam()

        '// ���Ұ�SET
        Call s_SetParam()

'//�f�o�b�O
'Call s_DebugPrint()

        '// ���k���X�g���擾
        w_iRet = f_Get_DetailData()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

        '// �s���o�����׏��擾
        w_iRet = f_Get_AbsInfo()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		If m_Rs_M.EOF = True Then
	        '// �y�[�W��\��
	        Call showWhitePage("���k��񂪂���܂���")
		Else
	        '// �y�[�W��\��
	        Call showPage()
		End If

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gf_closeObject(m_Rs_M)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen   = ""
    m_iKyokanCd  = ""
    m_sGakunen = ""
    m_sClassNo = ""
    m_sTuki = ""

    m_sGyoji_Cd  = ""
    m_sGyoji_Mei = ""
    m_sKaisi_Bi  = ""
    m_sSyuryo_Bi = ""
    m_sSoJikan   = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen = trim(Request("NENDO"))
    m_iKyokanCd = trim(Request("KYOKAN_CD"))
    m_sGakunen  = trim(Request("GAKUNEN"))
    m_sClassNo  = trim(Request("CLASSNO"))
    m_sTuki     = trim(Request("TUKI"))

    m_sGyoji_Cd  = trim(Request("GYOJI_CD"))
    m_sGyoji_Mei = trim(Request("GYOJI_MEI"))
    m_sKaisi_Bi  = trim(Request("KAISI_BI"))
    m_sSyuryo_Bi = trim(Request("SYURYO_BI"))
    m_sSoJikan   = trim(Request("SOJIKANSU"))

	call gf_Get_SyuketuEnd(cint(m_sGakunen),m_sEndDay)

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()

    response.write "<font color=#000000>m_iSyoriNen = " & m_iSyoriNen  & "</font><br>"
    response.write "<font color=#000000>m_iKyokanCd = " & m_iKyokanCd  & "</font><br>"
    response.write "<font color=#000000>m_sGakunen  = " & m_sGakunen   & "</font><br>"
    response.write "<font color=#000000>m_sClassNo  = " & m_sClassNo   & "</font><br>"
    response.write "<font color=#000000>m_sTuki     = " & m_sTuki      & "</font><br>"

    response.write "<font color=#000000>m_sGyoji_Cd = " & m_sGyoji_Cd  & "</font><br>"
    response.write "<font color=#000000>m_sGyoji_Mei= " & m_sGyoji_Mei & "</font><br>"
    response.write "<font color=#000000>m_sKaisi_Bi = " & m_sKaisi_Bi  & "</font><br>"
    response.write "<font color=#000000>m_sSyuryo_Bi= " & m_sSyuryo_Bi & "</font><br>"
    response.write "<font color=#000000>m_sSoJikan  = " & m_sSoJikan   & "</font><br>"

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�ꗗ���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Get_DetailData()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_Get_DetailData = 1

    Do 

        '// ���׃f�[�^
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO, "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN," 
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO,"
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "   T11_GAKUSEKI.T11_SIMEI "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN,T11_GAKUSEKI "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN=" & cInt(m_sGakunen) & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS=" & cInt(m_sClassNo)
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY T13_GAKU_NEN.T13_GAKUSEKI_NO"

'response.write "<font color=#000000>" & w_sSQL & "</font><BR>"
        iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_DetailData = 99
            Exit Do
        End If

        '//�������擾
        m_iRsCnt = 0
        If m_Rs_M.EOF = False Then
            m_iRsCnt = gf_GetRsCount(m_Rs_M)
        End If

        f_Get_DetailData = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  �s���o���f�[�^���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Get_AbsInfo()

    Dim w_sSQL

    On Error Resume Next
    Err.Clear
    
    f_Get_AbsInfo = 1

    Do 

        '// �o���f�[�^
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD, "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN," 
        w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS, "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUSEKI_NO, "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_KEKKA "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_SYUKKETU "
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "   T22_NENDO=" & cInt(m_iSyoriNen)   & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GAKUNEN=" & cInt(m_sGakunen)  & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T22_CLASS=" & cInt(m_sClassNo)    & " AND "
        w_sSQL = w_sSQL & vbCrLf & "   T22_GYOJI_CD='" & m_sGyoji_Cd & "'"

'response.write "<font color=#000000><br>" & w_sSQL & "<br>"
        iRet = gf_GetRecordset(m_Rs_G, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_AbsInfo = 99
            Exit Do
        End If

        f_Get_AbsInfo = 0
        Exit Do
    Loop

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear

%>
    <html>
    <head>
    <title>�s���p�o������</title>
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

		//�X�N���[����������
		parent.init();

        //�w�b�_����\��submit
        document.frm.target = "topFrame";
        document.frm.action = "kks0140_middle.asp"
        document.frm.submit();
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
        parent.document.location.href="default.asp"

    }

    //************************************************************
    //  [�@�\]  ���̓`�F�b�N(onBlur��)
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_CheckData(p_ObjName,p_Total){

        var objName="document.frm."+p_ObjName
        var Kekka = eval(objName);

        if (f_Trim(Kekka.value)!=""){

            //if (isNaN(f_Trim(Kekka.value))){
            if (f_chkNumber(f_Trim(Kekka.value))==1){
                alert("���͒l���s���ł�")
                Kekka.focus();
                return;
            }else{
                var vKekka = new Number(Kekka.value)
                var vTotal = new Number(p_Total)
                if(vKekka > vTotal){
                    alert("�����Ԃ𒴂������Ԃ����͂���Ă��܂��B")
                    Kekka.focus();
                    return;
                };
            };
        };

    };

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
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Touroku(){

        //���k��
        if (document.frm.iMax.value <= 0){
            //alert("�f�[�^������܂���B")
            return;
        };

		//���̓`�F�b�N(NN�Ή�)
		iRet = f_CheckData_All();
		if( iRet != 0 ){
			return;
		}

        if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
           return ;
        }

		//�w�b�_���󔒕\��
		parent.topFrame.document.location.href="white.htm"

        //���X�g����submit
        document.frm.target = "main";
        document.frm.action = "./kks0140_edt.asp"
        document.frm.submit();
        return;
    }

    //************************************************************
    //  [�@�\]  ���͒l������(�o�^�{�^��������)
    //  [����]  �Ȃ�
    //  [�ߒl]  0:����OK�A1:�����װ
    //  [����]  ���͒l��NULL�����A�p���������A�����������s��
    //          ���n�ް��p���ް������H����K�v������ꍇ�ɂ͉��H���s��
    //************************************************************
    function f_CheckData_All() {

		if(document.frm.iMax.value==1){
			var wKekka = new String("SU_" & document.frm.GAKUSEKI_NO.value)

			iRet = f_CheckData_NN(wKekka,<%=m_sSoJikan%>);
			if( iRet != 0 ){
				return 1;
			}

		}else{

			var i
			var w_bCheck = 0
			for (i = 0; i < document.frm.iMax.value; i++) {
				var wKekka = new String("SU_" + document.frm.GAKUSEKI_NO[i].value)
				iRet = f_CheckData_NN(wKekka,<%=m_sSoJikan%>);
				if( iRet != 0 ){
					w_bCheck = 1;
					break;
				};
			};

			if (w_bCheck == 1){
				return 1;
			};
		};
		return 0;
	};

    //************************************************************
    //  [�@�\]  ���̓`�F�b�N(NN�p)
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_CheckData_NN(p_ObjName,p_Total){

        var objName="document.frm."+p_ObjName
        var Kekka = eval(objName);
        
		if (typeof(Kekka) != "undefined"){

			if (f_Trim(Kekka.value)!=""){

			    //if (isNaN(f_Trim(Kekka.value))){
			    if (f_chkNumber(f_Trim(Kekka.value))==1){
			        alert("���͒l���s���ł�")
			        Kekka.focus();
			        return 1;
			    }else{
			        var vKekka = new Number(Kekka.value)
			        var vTotal = new Number(p_Total)
			        if(vKekka > vTotal){
			            alert("�����Ԃ𒴂������Ԃ����͂���Ă��܂��B")
			            Kekka.focus();
			        	return 1;
			        };
			    };
			};
		};
        return 0;
    };


    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <center>

    <form name="frm" method="post" onClick="return false;">
    <%Do%>

        <%If m_Rs_M.EOF = True Then%>
            <br><br>
            <span class="msg">���k��񂪂���܂���</span>
			<%Exit Do%>
		<%End If%>

        <%If Trim(Request("GYOJI_CD")) = "" Then%>
        <%Else%>

            <!--���׃��X�g��-->

            <table >
                <tr><td valign="top" >
            <table class="hyo"  border="1" >

            <%i = 0%>
            <%If m_Rs_M.EOF = True Then%>
            <%Else%>

                <%
				'//���s�J�E���g
	            w_iCnt = INT(m_iRsCnt/2 + 0.9)

				Dim w_IdouCnt
				Dim w_sIdouMei
				w_IdouCnt = 1

                Do Until m_Rs_M.EOF
                    i = i + 1

                    '//�w��NO���擾
                    w_iGakusekiNo = m_Rs_M("T13_GAKUSEKI_NO")

					'//�ٓ�������ꍇ�擾����
					w_IdouCnt = gf_Set_Idou(Cstr(w_iGakusekiNo),m_iSyoriNen,w_sIdouMei)

                    '//���ټ�Ă̸׽���Z�b�g
                    Call gs_cellPtn(w_Class) 
                %>
                    <tr>
                        <td nowrap class=<%=w_Class%> width="80"  align="center"><%=w_iGakusekiNo%><input type="hidden" name="GAKUSEKI_NO" value="<%=w_iGakusekiNo%>"><br></td>
                        
						<td nowrap class=<%=w_Class%> width="150" align="left"><%=m_Rs_M("T11_SIMEI")%><br></td>
                        <%
                        If m_Rs_G.EOF = False Then
                            m_Rs_G.MoveFirst
                            Do Until m_Rs_G.EOF
                                w_iKekka = ""

                                If cStr(trim(m_Rs_G("T22_GAKUSEKI_NO"))) = cStr(trim(w_iGakusekiNo)) Then
     
	                               w_iKekka = gf_SetNull2Zero(m_Rs_G("T22_GYOJI_KEKKA"))

                                    Exit Do
                                End If
                                m_Rs_G.MoveNext
                            Loop
                            m_Rs_G.MoveFirst
                            If cInt(w_iKekka) = 0 Then
                                w_iKekka = ""
                            End If

                        End If
                        %>
                       	<% IF w_IdouCnt = 1 Then %>
						 	<td class="<%=w_Class%>" width="80" align="center">
					   	<% Else %>
							<td class="NOCHANGE" width="80" align="center" >
						<% End IF %>
						<%
						'//NN�Ή�
						If session("browser") = "IE" Then
							w_sInputClass = "class='num'"
						Else
							w_sInputClass = ""
						End If

						%>
<% IF w_IdouCnt = 1 Then %>
					<% 'If m_sEndDay < m_sSyuryo_Bi then %>

							<input <%=w_sInputClass%> type="text" name="SU_<%=w_iGakusekiNo%>" value="<%=w_iKekka%>"  size="5" maxlength="2"><br></td>

					<% 'Else %>

							<%'=w_iKekka%><!-- <br></td> -->

					<% 'End If %>
<% Else %>

<%=w_sIdouMei%><br></td>

<%End if%>
                    </tr>
                    <%
					'//2��ڕ\��
					If i = w_iCnt Then
					%>
                        </table>
                    </td>
					<td width="10"><br></td>
					<td valign="top" >
                        <table class="hyo"  border="1" >

                        <%'//���ټ�Ă̸׽���N���A
						w_Class = ""
						%>

                    <%End If%>

                    <%m_Rs_M.MoveNext%>
                <%Loop%>
            <%End If%>

                    </table>
                </td></tr>
            </table>
            <br>
	<% 'If m_sEndDay < m_sSyuryo_Bi then %>
            <table>
                <td ><input class=button type="button" onclick="javascript:f_Touroku();" value="�@�o�@�^�@"></td>
                <td ><input class=button type="button" onclick="javascript:f_Cancel();" value="�L�����Z��"></td>
            </table>
	<% 'Else %>
            <!--table>
                <td ><input class=button type="button" onclick="javascript:f_Cancel();" value=" �߁@�� "></td>
            </table-->
	<% 'End If %>

        <%
        End If

        Exit Do

    Loop%>

    <!--�l�n���p-->
    <input type="hidden" name="NENDO"     value="<%=m_iSyoriNen%>">
    <input type="hidden" name="KYOKAN_CD" value="<%=m_iKyokanCd%>">
    <input type="hidden" name="GAKUNEN"   value="<%=m_sGakunen%>">
    <input type="hidden" name="CLASSNO"   value="<%=m_sClassNo%>">
    <input type="hidden" name="TUKI"      value="<%=m_sTuki%>">
    <input type="hidden" name="iMax"      value="<%=i%>">

    <INPUT TYPE=HIDDEN NAME="GYOJI_CD"  value = "<%=m_sGyoji_Cd%>">
    <INPUT TYPE=HIDDEN NAME="GYOJI_MEI" value = "<%=m_sGyoji_Mei%>">
    <INPUT TYPE=HIDDEN NAME="KAISI_BI"  value = "<%=m_sKaisi_Bi%>">
    <INPUT TYPE=HIDDEN NAME="SYURYO_BI" value = "<%=m_sSyuryo_Bi%>">
    <INPUT TYPE=HIDDEN NAME="SOJIKANSU" value = "<%=m_sSoJikan%>">
    <INPUT TYPE=HIDDEN NAME="ENDDAY" value = "<%=m_sEndDay%>">

    </form>
    </center>
    </body>
    </html>
<%
End Sub

'********************************************************************************
'*  [�@�\]  ��HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
    <html>
    <head>
    <title>�s���p�o������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
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
    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">
	<center>
	<br><br><br>
		<span class="msg"><%=p_Msg%></span>
	</center>

    </body>
    </html>
<%
End Sub
%>