<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0150_middle.asp
' �@      �\: ���y�[�W ���ѓo�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:
'           :
' ��      ��:
' ��      �n:
'           :
' ��      ��:
'           �������\��
'				
'			���\���{�^���N���b�N��
'				
'-------------------------------------------------------------------------
' ��      ��: 2002/06/21 shin
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Dim  m_bErrFlg           '�װ�׸�
	
	'//�����I��p��Where����
	Dim m_iNendo		'�N�x
	Dim m_sSikenKBN		'�����敪
	Dim m_iGakunen		'�w�N
	Dim m_sClassNo		'�w��
	Dim m_sKamokuCd		'�ȖڃR�[�h
	Dim m_sGakkaCd
	
	Dim m_FromSei
	Dim m_ToSei
	Dim m_FromKekka
	Dim m_ToKekka
	
	Dim m_bSeiInpFlg		'���͊��ԃt���O
	Dim m_bKekkaNyuryokuFlg	'���ۓ��͉\�׸�(True:���͉� / False:���͕s��)
	
	Dim m_UpdateDate
	
	'2002/06/21
	Dim m_iKamokuKbn
	Dim m_sKamokuBunrui
	Dim m_iSeisekiInpType
	
	Dim m_iDataCount
	Dim m_AryHyokaData()
	
	Dim m_iCount
	Dim m_sMiHyoka
	Dim m_Checked
	Dim m_Disabled
	Dim m_SchoolFlg
	Dim m_HyokaDispFlg
	Dim m_KekkaGaiDispFlg
	
	Dim m_TableWidth

	Dim m_sGakkoNO	'�w�Z�ԍ�
	
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
	Dim w_sSQL
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���ѓo�^"
	w_sMsg=""
	w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget=""
	
	On Error Resume Next
	Err.Clear
	
	m_bErrFlg = False
	
	Do
		'//�ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If
		
		'//���Ұ�SET
		Call s_SetParam()
		
		'//�s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'���ѓ��͕��@���������͂̂Ƃ��A�Ȗڕ]���f�[�^�擾
		if m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_STRING) then
			if not gf_GetKamokuHyokaData(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iDataCount,m_AryHyokaData) then 
				m_bErrFlg = True
				Exit Do
			end if
		end if
		
		'�w�Z�ԍ����擾
		if Not gf_GetGakkoNO(m_sGakkoNO) then
	        m_bErrFlg = True
		end if

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
	Call gs_CloseDatabase()
	
End Sub

'********************************************************************************
'*	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()
	
	m_iNendo	= request("txtNendo")
	m_sSikenKBN	= Cint(request("sltShikenKbn"))
	m_iGakunen	= Cint(request("txtGakuNo"))
	m_sClassNo	= Cint(request("txtClassNo"))
	m_sKamokuCd	= request("txtKamokuCd")
	m_sGakkaCd	= request("txtGakkaCd")
	
	m_bSeiInpFlg	= cbool(request("hidKikan"))
	m_bKekkaNyuryokuFlg	= request("hidKekkaNyuryokuFlg")
	
	m_iKamokuKbn	 	= request("hidKamokuKbn")
	m_sKamokuBunrui 	= request("hidKamokuBunrui")
	m_iSeisekiInpType 	= cint(request("hidSeisekiInpType"))
	
	m_UpdateDate = request("txtUpdDate")
	
	m_iCount = cint(request("i_Max"))
	m_sMiHyoka = request("hidMihyoka")
	m_SchoolFlg = cbool(request("hidSchoolFlg"))
	m_HyokaDispFlg = cbool(request("hidHyokaDispFlg"))
	m_KekkaGaiDispFlg = cbool(request("hidKekkaGaiDispFlg"))
	
	m_TableWidth = cint(request("hidTableWidth"))
	
	m_FromSei = gf_SetNull2String(request("hidFromSei"))
	m_ToSei = gf_SetNull2String(request("hidToSei"))
	m_FromKekka = gf_SetNull2String(request("hidFromKekka"))
	m_ToKekka = gf_SetNull2String(request("hidToKekka"))
	
	m_Checked  = ""
	m_Disabled = ""
	
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
	
	If m_iKamokuKbn = C_TUKU_FLG_TUJO Then '�ʏ���ƂƓ��ʊ����Ŏ����ς���B
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
'*  [�@�\]  ���ѓo�^�������̏ꍇ�]���R���{���쐬
'********************************************************************************
Sub s_SetHyokaCombo()
	Dim w_Str,w_lIdx
	
	w_Str = ""
	w_Str = w_Str & "<select name='sltHyoka'>"
	
	for w_lIdx = 0 to m_iDataCount-1
		
		w_Str = w_Str & "<option value='" & m_AryHyokaData(w_lIdx,0)
		'w_Str = w_Str & "#@#" & m_AryHyokaData(w_lIdx,1)
		w_Str = w_Str & "#@#" & m_AryHyokaData(w_lIdx,2)
		w_Str = w_Str & "'>" & m_AryHyokaData(w_lIdx,0)
		
	next
	
	w_Str = w_Str & "<option value=''>�N���A"
	w_Str = w_Str & "</select>"
	
	response.write w_Str
	
End Sub

'********************************************************************************
'*  [�@�\]  ���]���̐ݒ�
'********************************************************************************
Sub setHyokaType()
	
	'�Ȗڂ����]��
	if cint(gf_SetNull2Zero(m_sMiHyoka)) = cint(C_MIHYOKA) then
		m_Checked = "checked"
	end if
	
	'���͊��ԊO
	if not m_bSeiInpFlg then
		m_Disabled = "disabled"
	end if
	
End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'********************************************************************************
Sub showPage()
	Dim w_sInputClass
		
	'//NN�Ή�
	If session("browser") = "IE" Then
		w_sInputClass = "class='num'"
	Else
		w_sInputClass = ""
	End If

%>

<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT language="javascript">
<!--
	//************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //************************************************************
    function window_onload(){
		//�X�N���[����������
		parent.init();
	}
	
	//************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //************************************************************
    function f_Touroku(){
        parent.main.f_Touroku();
    }
	
	//************************************************************
	//	[�@�\]	�L�����Z���{�^���������ꂽ�Ƃ�
	//************************************************************
	function f_Cancel(){
		//�����y�[�W��\��
        parent.document.location.href="default.asp";
	}
	
	//************************************************************
	//	[�@�\]	�y�[�X�g�{�^���������ꂽ�Ƃ�
	//************************************************************
	function f_Paste(pType){
		parent.main.document.frm.PasteType.value=pType;
		
		//submit�ŉ�ʂ��J���ƃE�B���h�E�̃X�e�[�^�X���ݒ�ł��Ȃ����ߤ
		//��U��y�[�W���J���Ă���A�V�E�B���h�E�ɑ΂���submit����B
		nWin=open("","Paste","location=no,menubar=no,resizable=yes,scrollbars=no,scrolling=no,status=no,toolbar=no,width=300,height=600,top=0,left=0");
		parent.main.document.frm.target="Paste";
		parent.main.document.frm.action="sei0150_paste.asp";
		parent.main.document.frm.submit();
	}
	
	//************************************************************
	//	[�@�\]	���]�����`�F�b�N���ꂽ�Ƃ�
	//************************************************************
	function setHyoka(){
		var w_num,w_type;
		var ob = new Array();
		
		if(document.frm.chkMiHyoka.checked){
			parent.main.document.frm.hidMihyoka.value=<%=C_MIHYOKA%>;
			w_type = true;
		}else{
			parent.main.document.frm.hidMihyoka.value="";
			w_type = false;
		}
		
		for(w_num=1;w_num<<%=m_iCount%>;w_num++){
			ob[0] = eval("parent.main.document.frm.chkHyokaFuno" + w_num);
			
			<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
				ob[1] = eval("parent.main.document.frm.Seiseki" + w_num);
			<% end if %>
			
			<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
				ob[2] = eval("parent.main.document.frm.hidSeiseki" + w_num);
			<% end if %>
			
			if(typeof(ob[0]) != "undefined" && ob[0].type == "checkbox"){
				if(w_type){
					ob[0].checked = false;
					<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
						ob[1].value = "";
					<% end if %>
					
					<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
						ob[2].value = "";
					<% end if %>
				}
				
				ob[0].disabled = w_type;
				
				<% if m_iSeisekiInpType <> C_SEISEKI_INP_TYPE_KEKKA then %>
					ob[1].disabled = w_type;
				<% end if %>
			}
		}
		<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
			parent.main.f_GetTotalAvg();
		<% end if %>
	}
	
	//-->
	</SCRIPT>
	</head>
	<body onload="window_onload();">
	<table border="0" cellpadding="0" cellspacing="0" height="245" width="100%">
		<tr>
			<td>
				<%
				If m_bSeiInpFlg or m_bKekkaNyuryokuFlg Then
					call gs_title(" ���ѓo�^ "," �o�@�^ ")
				Else
					call gs_title(" ���ѓo�^ "," �\�@�� ")
				End If
				%>
			</td>
		</tr>
		<tr>
			<td align="center" nowrap>
			<form name="frm" method="post">
				<table border=1 class=hyo width=670>
					<tr>
						<th class="header3" colspan="6" nowrap align="center">
						���ѓ��͊��ԁ@<%=f_ShikenMei()%>�@�@�@�X�V���F<%=m_UpdateDate%>
						</th>
					</tr>
					<tr>
						<th class=header3 width="96"  align="center">���ѓ��͊���</th><td class=detail width="239"  align="center" colspan="2"><%=m_FromSei%> �` <%=m_ToSei%></td>
						<th class=header3 width="96"  align="center">
							<%if m_sGakkoNO = C_NCT_KURUME then%>
								<font size=1>��u���ԓ��͊���</font>
							<%else%>
								���ۓ��͊���
							<%end if%>
						</th><td class=detail width="239"  align="center" colspan="2"><%=m_FromKekka%> �` <%=m_ToKekka%></td>
					</tr>
					<tr>
						<th class=header3 width="96"  align="center">���{�Ȗ�</th>
						<%
							w_str = m_iGakunen & "�N�@" & gf_GetClassName(m_iNendo,m_iGakunen,m_sClassNo) & "�@" & f_GetKamokuName(m_iGakunen,m_sGakkaCd,m_sKamokuCd)
						%>
						<td class=detail colspan="5" align="center"><%=w_str%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<span class=msg2>
				���u���X�v�v�́A���Əo�����̓��j���[�����X���͂��ꂽ��L�����܂ł̏o���󋵂ł��B<br>
				<%if m_KekkaGaiDispFlg then%>���u�ΏۊO�v�́A�����Ȃǂ̗݌v����͂��Ă��������B<br><%end if%>
				<%if m_bSeiInpFlg or m_bKekkaNyuryokuFlg then %>���w�b�_�̕����F���u<FONT COLOR="#99CCFF">����</FONT>�v�̂悤�ɂȂ��Ă��镔�����N���b�N����ƁAExcel�\��t���p�̉�ʂ��J���܂��B<br><%end if%>
				<!--<%if m_bSeiInpFlg and m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_STRING) then %>�� ���͂������u�]���v��I����A�Y������w���̐��ї����N���b�N���ĉ������B<BR><% end if %>-->
				<%
				if m_HyokaDispFlg and m_iKamokuKbn = C_TUKU_FLG_TUJO then
					Select Case m_sSikenKBN
						Case C_SIKEN_ZEN_TYU
							%>�� �]�������N���b�N����ƁA�]���̓��͂��ł��܂��B�i�����E�̏��ŕ\������܂��j<br><%
						Case C_SIKEN_KOU_TYU
							%>�� �]�������N���b�N����ƁA�]���̓��͂��ł��܂��B�i���������E�̏��ŕ\������܂��j<br><%
						Case Else
							response.write "<BR>"
					End Select
				end if
				%>
				</span>
				
				<% if m_bSeiInpFlg or m_bKekkaNyuryokuFlg Then %>
					<input type="button" class="button" value="�@�o�@�^�@" onclick="f_Touroku();">�@
				<% end if %>
				<input type="button" class="button" value="�L�����Z��" onclick="f_Cancel();">
				
			</td>
		</tr>
		<tr>
			<td align="center" valign="bottom" nowrap>
				<table class="hyo" border="1" align="center" width="<%=m_TableWidth%>">
					<tr><th class="header3" colspan="14" nowrap align="center">
							�����Ɛ�&nbsp;<%If m_bSeiInpFlg or m_bKekkaNyuryokuFlg Then%><input type="text" <%=w_sInputClass%> maxlength="3" style="width:30px" name="txtSouJyugyou" value="<%= Request("hidSouJyugyou") %>"><% Else %><%= Request("hidSouJyugyou") %><% End if%>�@
							�����Ɛ�&nbsp;<%If m_bSeiInpFlg or m_bKekkaNyuryokuFlg Then%><input type="text" <%=w_sInputClass%> maxlength="3" style="width:30px" name="txtJunJyugyou" value="<%= Request("hidJunJyugyou") %>"><% Else %><%= Request("hidJunJyugyou") %><% End if%>�@
							&nbsp;&nbsp;
							<%
							if m_bSeiInpFlg then
								
								'���ѓ��͕��@���������͂̂Ƃ��A�]���R���{�\��
								if m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_STRING) then
									Call s_SetHyokaCombo()
								end if
							end if
							%>
							
							<%
							if m_SchoolFlg then
								Call setHyokaType()
							%>
							&nbsp;&nbsp;
							<input type="checkbox" name="chkMiHyoka" value="4" <%=m_Disabled%> onClick="setHyoka();" <%=m_Checked%>>���]���@
							<% end if %>
								
						</th></tr>                                                                                                                                                 
					<tr>
						<th class="header3" rowspan="2" width="65" nowrap><%=gf_GetGakuNomei(m_iNendo,C_K_KOJIN_1NEN)%></th>
						<th class="header3" rowspan="2" width="150" nowrap>���@��</th>
						<th class="header3" colspan="4" width="120" nowrap>���ї���</th>
						
						<% if m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_NUM) and m_bSeiInpFlg then %>
							<th class="header3" rowspan="2" width="50" nowrap onClick="f_Paste('Seiseki')"><FONT COLOR="#99CCFF">����</FONT></th>
						<% else %>
							<th class="header3" rowspan="2" width="50" nowrap>����</th>
						<% end if %>
						
						<% If m_HyokaDispFlg and cstr(m_iKamokuKbn) = cstr(C_JIK_JUGYO) then %>
							<th class="header3" rowspan="2" width="50" nowrap>�]��</th>
						<% end if %>
						
						<th class="header3" colspan="2" width="110" nowrap>�x��</th>
						<th class="header3" colspan="<%=gf_IIF(m_KekkaGaiDispFlg,3,2)%>" width="<%=gf_IIF(m_KekkaGaiDispFlg,165,110)%>" nowrap">
							<%if m_sGakkoNO = C_NCT_KURUME then%>
								<font size=1>��u����</font>
							<%else%>
								����
							<%end if%>
						</th>
						<% if m_SchoolFlg then %>
							<th class="header3" rowspan="2" width="50"  nowrap>�]��<br>�s�\</th>
						<% end if %>
					</tr>
					
					<tr>
						<th class="header2" width="30" nowrap><span style="font-size:10px;">�O��</span></th>
						<th class="header2" width="30" nowrap><span style="font-size:10px;">�O��</span></th>
						<th class="header2" width="30" nowrap><span style="font-size:10px;">�㒆</span></th>
						<th class="header2" width="30" nowrap><span style="font-size:10px;">�w��</span></th>
						
						<% if m_bKekkaNyuryokuFlg then %>
							<th class="header2" width="55" nowrap onClick="f_Paste('Chikai')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">����</FONT></span></th>
						<% else %>
							<th class="header2" width="55" nowrap ><span style="font-size:10px;">����</span></th>
						<% end if %>
						
						<th class="header2" width="55" nowrap><span style="font-size:10px;">���X�v</span></th>
						
						<% if m_bKekkaNyuryokuFlg then %>
							<th class="header2" width="55" nowrap onClick="f_Paste('Kekka')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">�Ώ�</FONT></span></th>
							
							<% if m_KekkaGaiDispFlg then %>
								<th class="header2" width="55" nowrap onClick="f_Paste('KekkaGai')"><span style="font-size:10px;"><FONT COLOR="#99CCFF">�ΏۊO</FONT></span></th>
							<% end if %>
							
						<% else %>
							<th class="header2" width="55" nowrap><span style="font-size:10px;">�Ώ�</span></th>
							
							<% if m_KekkaGaiDispFlg then %>
								<th class="header2" width="55" nowrap><span style="font-size:10px;">�ΏۊO</span></th>
							<% end if %>
							
						<% end if %>
						
						<th class="header2" width="55" nowrap><span style="font-size:10px;">���X�v</span></th>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	
	<input type="hidden" name="hidSeisekiInpType" value="<%=m_iSeisekiInpType%>">
	<input type="hidden" name="hidKekkaGaiDispFlg" value="<%=m_KekkaGaiDispFlg%>">
	<input type="hidden" name="hidKekkaNyuryokuFlg" value="<%=m_bKekkaNyuryokuFlg%>">
	
	</body>
	</html>
<%
End sub
%>