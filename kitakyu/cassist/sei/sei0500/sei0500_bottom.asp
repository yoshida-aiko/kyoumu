<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���͎������ѓo�^
' ��۸���ID : sei/sei0500/sei0500_bottom.asp
' �@      �\: ���y�[�W ���͎����̐��т���͂���
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:

'-------------------------------------------------------------------------
' ��      ��: 2001/09/06 ���`�i�K
' ��      �X: 2016/05/18 Nishimura �ٓ�(�x�w��)�̏ꍇ�X�V�ł��Ȃ���Q�Ή�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�

	'�����I��p��Where����
    Public m_iNendo			'�N�x
    Public m_sKyokanCd		'�����R�[�h
    Public m_sGakuNo		'�w�N
    Public m_sClassNo		'�w��
    Public m_sKamokuCd		'�ȖڃR�[�h
    Public m_sSiKenCd		'�����R�[�h

    Public m_SeitoRs		'ں��޾�ĵ�޼ު��(���k)
    Public m_rCnt			'ں��޶ݳ�(���k)
    Public m_lKikan			'0�F���ѓ��͊��ԓ��A1�F���ѓ��͊��ԊO

	Public	m_iMax			'�ő�y�[�W
	Public  m_Half			'�ő�y�[�W�̔���
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

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���͎������ѓo�^"
	w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_lKikan = 0
	
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

		'//���ԃf�[�^�̎擾
        w_iRet = f_Nyuryokudate()
		If w_iRet = 1 Then
			m_lKikan = 1
			'// �y�[�W��\��
			'Call No_showPage()
			'Exit Do
		End If
		
		w_iRet = 0
		
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If

		'//�N���X�ʐ��k�f�[�^�擾
        w_iRet = f_GetClassData()
		If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
		If m_SeitoRs.EOF Then
			Call ShowPage_No()
			Exit Do
		End If

	   '// �y�[�W��\��
	   Call showPage()
	   Exit Do
	Loop

	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If


    '// �I������
    Call gf_closeObject(m_SeitoRs)
    Call gs_CloseDatabase()

End Sub

Sub s_SetParam()
'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	m_iNendo	= request("txtNendo")
	m_sKyokanCd	= request("txtKyokanCd")
	m_sSiKenCd	= Cint(request("txtShikenCd"))
	m_sGakuNo	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")

End Sub

Function f_GetClassData()
'********************************************************************************
'*	[�@�\]	���k���擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
	Dim w_iNyuNendo

	On Error Resume Next
	Err.Clear
	f_GetClassData = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_TOKUTEN,  "
		w_sSQL = w_sSQL & vbCrLf & " 	T11.T11_SIMEI,  "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_GAKUSEKI_NO "
		w_sSQL = w_sSQL & vbCrLf & " FROM  "
		w_sSQL = w_sSQL & vbCrLf & " 	T11_GAKUSEKI T11, "
		w_sSQL = w_sSQL & vbCrLf & " 	T13_GAKU_NEN T13, "
		w_sSQL = w_sSQL & vbCrLf & " 	T33_SIKEN_SEISEKI T33  "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "	T33.T33_GAKUSEKI_NO  = T13.T13_GAKUSEKI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "	T13.T13_GAKUSEI_NO   = T11.T11_GAKUSEI_NO  AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_NENDO        = T13.T13_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_NENDO        =  " & m_iNendo          & " AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_SIKEN_CD     =  " & m_sSiKenCd        & " AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_SIKEN_KAMOKU = '" & m_sKamokuCd       & "' AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_GAKUNEN      =  " & m_sGakuNo         &" AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T33.T33_CLASS        =  " & m_sClassNo
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY T33_GAKUSEKI_NO "

        iRet = gf_GetRecordset(m_SeitoRs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
			m_bErrFlg = True
            msMsg = Err.description
            f_GetClassData = 99
            Exit Do
        End If

		'//ں��ރJ�E���g�擾
		m_rCnt = gf_GetRsCount(m_SeitoRs)

		f_GetClassData = 0
		Exit Do
	Loop

End Function

Function f_Nyuryokudate()
'********************************************************************************
'*	[�@�\]	���͊��Ԏ擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	On Error Resume Next
	Err.Clear
	f_Nyuryokudate = 1

	Do

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & "  SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SEISEKI_KAISI, "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SEISEKI_SYURYO "
		w_sSQL = w_sSQL & vbCrLf & "  FROM  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27_SIKEN M27 "
		w_sSQL = w_sSQL & vbCrLf & "  WHERE  "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KBN    =  " & C_SIKEN_JITURYOKU & " AND "	'�����敪(���͎����̂�)
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_NENDO        =  " & m_iNendo	& "  AND "		'�����N�x
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_CD     =  " & m_sSiKenCd  & "  AND "
		w_sSQL = w_sSQL & vbCrLf & "  	M27.M27_SIKEN_KAMOKU = '" & m_sKamokuCd & "' AND "
		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_SEISEKI_KAISI  <= '" & gf_YYYY_MM_DD(date(),"/") & "' AND"
		w_sSQL = w_sSQL & vbCrLf & " 	M27.M27_SEISEKI_SYURYO >= '" & gf_YYYY_MM_DD(date(),"/") & "' "

		w_iRet = gf_GetRecordset(m_DRs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_Nyuryokudate = 99
			m_bErrFlg = True
			Exit Do 
		End If

		If m_DRs.EOF Then
			Exit Do
		End If

	    Call gf_closeObject(m_DRs)

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
	On Error Resume Next
	Err.Clear

	'// ���k���̔���
	m_Half = gf_Round(m_rCnt / 2, 0)

%>
	<html>
	<head>
	<link rel=stylesheet href="../../common/style.css" type=text/css>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT language="javascript">
	<!--
	    //************************************************************
	    //  [�@�\]  �y�[�W���[�h������
	    //  [����]
	    //  [�ߒl]
	    //  [����]
	    //************************************************************
	    function window_onload(){

			//�X�N���[����������
			parent.init();

			//���э��v�l�̎擾
			f_GetTotalAvg();

	        //submit
	        document.frm.target = "topFrame";
	        document.frm.action = "sei0500_middle.asp?<%=Request.Form.Item%>"
	        document.frm.submit();

	    }
		//************************************************************
		//	[�@�\]	�L�����Z���{�^���������ꂽ�Ƃ�
		//	[����]	�Ȃ�
		//	[�ߒl]	�Ȃ�
		//	[����]
		//************************************************************
		function f_Cansel(){

	        //�����y�[�W��\��
	        parent.document.location.href="default.asp"
		
		}
	    //************************************************************
	    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
	    //  [����]  �Ȃ�
	    //  [�ߒl]  �Ȃ�
	    //  [����]
	    //************************************************************
	    function f_Touroku(){

			// ��������
			for(i=1; i < <%= m_rCnt %>; i++){
//Ins_s 2016/05/18 Nishimura
//�ٓ��̏ꍇ�̓`�F�b�N���Ȃ�
				objIdo = eval("document.frm.hidIdoCnt"+i)

				if (objIdo.value == 1){
//Ins_e 2016/05/18 Nishimura
					obj = eval("document.frm.Seiseki"+i)
					if (obj.value.match(/[^0-9]/) ){
			            alert("���͒l���s���ł�");
						obj.focus();
						return;
					}
				}
			}

			//�o�^����
	        document.frm.action = "sei0500_upd.asp?<%=Request.Form.Item%>";
	        document.frm.target = "main";
	        document.frm.submit();

	    }

		//************************************************
		//  [�@�\]  Enter �L�[�ŉ��̓��̓t�H�[���ɓ����悤�ɂȂ�
		//  [����]  p_inpNm	�Ώۓ��̓t�H�[����
		//          p_frm	�Ώۃt�H�[��
		//          i		���݂̔ԍ�
		//  [����]  
		//************************************************
		function f_MoveCur(p_inpNm,p_frm,i){
			if (event.keyCode == 13){		//�����ꂽ�L�[��Enter(13)�̎��ɓ����B
				i++;
				if (i > <%=m_rCnt%>) i = 1; //i���ő�l�𒴂���ƁA�͂��߂ɖ߂�B
				inpForm = eval("p_frm."+p_inpNm+i);
				inpForm.focus();			//�t�H�[�J�X���ڂ��B
				inpForm.select();			//�ڂ����e�L�X�g�{�b�N�X����I����Ԃɂ���B
			}else{
				return false;
			}
			return true;
		}

	//-->
	</SCRIPT>
	</head>

    <body onLoad="window_onload();">
	<form name="frm" method="post" onClick="return false;">
	<center>

	<table >
		<tr>
			<td valign="top">

				<table class="hyo" border="1" align="center" width="280">
				<%

					Dim w_IdouCnt
					Dim w_sIdouMei

					i = 1
					Do until m_SeitoRs.Eof or i > m_Half

						'**�ٓ������@Add 2001.12.22 oakda*********************
						w_IdouCnt = gf_Set_Idou(Cstr(m_SeitoRs("T33_GAKUSEKI_NO")),m_iNendo,w_sIdouMei)

						if w_sIdouMei <> "" then
							w_sIdouMei = "[" & w_sIdouMei & "]"
						End if
						'*****************************************************

						Call gs_cellPtn(w_cell)
						%>
							<tr>
								<td class="<%=w_cell%>" width="50" align="center"><%=m_SeitoRs("T33_GAKUSEKI_NO")%></td>
								<td class="<%=w_cell%>" width="200"><%=m_SeitoRs("T11_SIMEI")%><%=w_sIdouMei%></td>
								<input type="hidden" align="center" name="hidIdoCnt<%=i%>" value="<%= w_IdouCnt %>">

<% IF w_IdouCnt = 1 Then %>
							<%If m_lKikan = 1 Then%>
								<td class="<%=w_cell%>" width="30"align="right">
								<input type="hidden" class='num' align="center"  name="Seiseki<%=i%>" value="<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)">
								<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>
								</td>
							<% Else %>
								<td class="<%=w_cell%>" width="30"><input type="text" class='num' align="center"  name="Seiseki<%=i%>" value="<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)"></td>
							<% End IF %>
								<input type="hidden" align="center" name="hidGakusekiNo<%=i%>" value="<%= m_SeitoRs("T33_GAKUSEKI_NO") %>">
<% Else %>
								<td class="<%=w_cell%>" align="center" width="30">-</td>
<% End IF %>
							</tr>
						<%
						i = i + 1
						m_SeitoRs.MoveNext
					Loop
				%>
				</table>

			</td>
			<td valign="top">

				<table class="hyo" border="1" align="center" width="280">
				<%
					Do until m_SeitoRs.Eof
						
						'**�ٓ������@Add 2001.12.22 oakda*********************
						w_IdouCnt = gf_Set_Idou(Cstr(m_SeitoRs("T33_GAKUSEKI_NO")),m_iNendo,w_sIdouMei)

						if w_sIdouMei <> "" then
							w_sIdouMei = "[" & w_sIdouMei & "]"
						End if
						'*****************************************************

						Call gs_cellPtn(w_cell)

						%>
							<tr>
								<td class="<%=w_cell%>" width="50" align="center"><%=m_SeitoRs("T33_GAKUSEKI_NO")%></td>
								<td class="<%=w_cell%>" width="200"><%=m_SeitoRs("T11_SIMEI")%><%=w_sIdouMei%></td>
								<input type="hidden" align="center" name="hidIdoCnt<%=i%>" value="<%= w_IdouCnt %>">

<% IF w_IdouCnt = 1 Then %>
							<%If m_lKikan = 1 Then%>
								<td class="<%=w_cell%>" width="30" align="right">
								<input type="hidden" class='num' align="center"  name="Seiseki<%=i%>" value="<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)">
								<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>
								</td>
							<% Else %>
								<td class="<%=w_cell%>" width="30"><input type="text" class='num' align="center"  name="Seiseki<%=i%>" value="<%=gf_SetNull2String(m_SeitoRs("T33_TOKUTEN"))%>" size=2 maxlength=3 onKeyDown="f_MoveCur('Seiseki',this.form,<%=i%>)"></td>
							<% End IF %>

								<input type="hidden" align="center" name="hidGakusekiNo<%=i%>" value="<%= m_SeitoRs("T33_GAKUSEKI_NO") %>">
<% Else %>
								<td class="<%=w_cell%>" align="center" width="30">-</td>
<% End IF %>
							</tr>
						<%
						i = i + 1
						m_SeitoRs.MoveNext
					Loop%>




				</table>

			</td>
		</tr>
		<tr>
			<td colspan=2>
				<table class="hyo" border="1" align="center" width="100%">
					<tr>
						<td class="header" nowrap align="right">
							<FONT COLOR="#FFFFFF"><B>���э��v</B></FONT>
							<input type="text" name="txtTotal" size="5" <%=w_sInputClass%> readonly>
						</td>
					</tr>
					<tr>
						<td class="header" nowrap align="right">
							<FONT COLOR="#FFFFFF"><B>���ϓ_</B></FONT>
							<input type="text" name="txtAvg" size="5" <%=w_sInputClass%> readonly>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<table width="50%">
		<tr>
			<td align="center" nowrap>
				<input type="button" class="button" value="�@�o�@�^�@" onclick="javascript:f_Touroku()">�@
				<input type="button" class="button" value="�L�����Z��" onclick="javascript:f_Cansel()">
			</td>
		</tr>
	</table>

	<input type="hidden" name="hidRecCnt" value="<%= m_rCnt %>">
	<input type="hidden" name="i_Max"       value="<%=i%>">
	<input type="hidden" name="PasteType" value="">
	<input type="hidden" name="i_Maxherf" value="<%=m_Half%>">
	</FORM>
	</center>
	</body>
	<SCRIPT>
		//************************************************************
		//	[�@�\]	���т��ύX���ꂽ�Ƃ�
		//	[����]	�Ȃ�
		//	[�ߒl]	�Ȃ�
		//	[����]	���т̍��v�ƕ��ς����߂�
		//	[���l]	�w���̑�����������͍̂Ō�ł��邽�߁A���̈ʒu�ɏ����B
		//************************************************************
		function f_GetTotalAvg(){
			var i;
			var total;
			var avg;
			var cnt;

			total = 0;
			cnt = 0;
			avg = 0;

	<%If m_iKikan <> "NO" Then	'���͊��Ԓ�%>

			//�w�����ł̃��[�v
			for(i=0;i<<%=i%>;i++) {

				//���݂��邩�ǂ���
				textbox = eval("document.frm.Seiseki" + (i+1));
				if (textbox) {
					//�����̓`�F�b�N
					if (textbox.value != "") {
						//�����łȂ��͖̂�������
						if (!isNaN(textbox.value)) {
							total = total + parseInt(textbox.value);
						}
					}
					cnt = cnt + 1;
				}
			}

	<% Else	'���͊��Ԓ��ł͂Ȃ�%>
		total = <%=w_lSeiTotal%>;
		cnt   = <%=w_lGakTotal%>;
	<% End If%>

			document.frm.txtTotal.value=total;

			//�l�̌ܓ�
			if (cnt!=0){
				avg = total/cnt;
				avg = avg * 10;
				avg = Math.round(avg);
				avg = avg / 10;
			}
			
			document.frm.txtAvg.value=avg;
		}
	</SCRIPT>

	</html>
<%
End sub

Sub No_showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

    <body>
	<form name="frm" method="post">
	<center>
	<br><br><br>
		<span class="msg">���ѓ��͊��ԊO�ł��B</span>
	</center>

	<input type="hidden" name="txtMsg" value="���ѓ��͊��ԊO�ł��B">

	</form>
	</body>
	</html>

<%
End Sub
Sub showPage_No()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

    <body>
	<form name="frm" method="post">
	</head>

	<body>
	<br><br><br>
	<center>
		<span class="msg">�f�[�^�����݂��܂���B</span>
	</center>

	<input type="hidden" name="txtMsg" value="�f�[�^�����݂��܂���B">

	</form>
	</body>
	</html>

<%
End Sub
%>