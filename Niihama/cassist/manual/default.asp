<%@ Language=VBScript %>

<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ����}�j���A��
' ��۸���ID : mst/manual/default.asp
' �@      �\: �����N��̃y�[�W�̕ύX���s��
'-------------------------------------------------------------------------
' ��      ��:�Ȃ�
' ��      ��:�Ȃ�
' ��      �n:m_sLinkNo	:�I�����ꂽ�����N��i���o�[
' ��      ��:
'           �������\��
'			�C�ӂ̃y�[�W�����C���t���[���ɕ\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/07/26 �≺�@�K��Y
' ��      �X: 2001/09/25 �≺�@�K��Y
'*************************************************************************/

	Public	m_shtmlName

%>
<!--#include file="../Common/com_All.asp"-->
<html>

<head>
<title>���������V�X�e���}�j���A���FCampus Assist manual</title>
<link rel=stylesheet href="../common/style.css" type=text/css>
<script language="javascript">
<!--
    //************************************************************
    //  [�@�\]  �C�ӂ̃y�[�W��\��
    //  [����]  txtpage :�\����
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_PageClick(p_No){

        document.frm.action = "./main.asp";
        document.frm.target = "_top";
	document.frm.txtLinkNo.value = p_No;
        document.frm.submit();

    }
//-->
</script>
</head>

<body marginheight=0 marginwidth=0 bgcolor="#ffffff" topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0">
<div align="center">

<form name="frm" action="./main.asp" target="" Method="POST">

<table cellspacing="0" cellpadding="0" width=100% height=100% border="0">
	<tr>
		<td width="100%" height="12" align="center" background="img/ue.gif">
		<img src="img/sp.gif">
		</td>
	</tr>
	<tr>
		<td align="center" valign="middle">

<table class=manual cellspacing="0" cellpadding="0" width=660 height=100% border="0">
	<tr>
		<td class=manual colspan="2" width="660" height="35%" align="center">
		<img src="img/topimage.jpg">
		</td>
	</tr>
	<tr>
		<td class=manual align="right" valign="top" height="65%" width="504">

			<img src="../image/sp.gif" height="15"><br>

			<table border="0" width=100% cellspacing="0" cellpadding="0">

			<tr>
			<td valign="top">
				<table width=120 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>1�E���O�C��</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(1)">���O�C��</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif"></td>
					</tr>
					<tr>
						<td class=manual><b>2�E���C�����j���[</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(2)">�ٓ��󋵈ꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(3)">�w�����ꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><b>3�E�o������</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(4)">���Əo������</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(5)">�s���o������</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(6)">�����o������</a></td>
					</tr>
				</table>
			</td>
			<td valign="top">

				<table width=174 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>4�E�����E����</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(7)">�������{�Ȗړo�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(8)">�����ēƏ��\���o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(9)">���ѓo�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(10)">���͎������ѓo�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(11)">���і������o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(35)">���ȓ����o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(12)">�������Ԋ��i�N���X�ʁj</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(13)">�������ԋ����\��ꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(14)">���шꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(15)">�l�ʐ��шꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(16)">���N�Y���҈ꗗ</a></td>
					</tr>
				</table>
			</td>
			<td valign="top">
				<table width=174 cellspacing="1" cellpadding="1" border="0">
					<tr>
					<tr>
						<td class=manual><b>5�E�X�P�W���[��</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(17)">�s�������ꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(18)">�N���X�ʎ��Ǝ��Ԉꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(19)">�����ʎ��Ǝ��Ԉꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(20)">���Ԋ������A��</a></td>
					</tr>
						<td class=manual><b>6�E���̑�����</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(21)">�i�H����o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(22)">�g�p���ȏ��o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(23)">�w���v�^�������o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(24)">�������������o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(25)">�e��ψ��o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(26)">�l���C�I���Ȗڌ���</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(27)">�����������ꗗ</a></td>
					</tr>
				</table>
			</td>
			<td valign="top">
				<table width=140 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>7�E��񌟍�</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(28)">�w����񌟍�</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(29)">���w�Z��񌟍�</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(30)">�����w�Z��񌟍�</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(31)">�i�H���񌟍�</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(32)">�󂫎��ԏ�񌟍�</a></td>
					</tr>
					<tr>
						<td class=manual><b>8�E�x���@�\</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(33)">���ʋ����\��</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(34)">�A���f����</a></td>
					</tr>
				</table>
			</td>
			</tr>
			</table>
		</td>
		<td>
			<img src="../image/sp.gif" width="10">
		</td>
	</tr>
</table>
		</td>
	</tr>
	<tr>
		<td width="100%" height="12" align="center">
			<input type="button" name="close" value="����" onClick="window.close();">
		</td>
	</tr>
	<tr>
		<td width="100%" height="12" align="right">
			<span class="msg"><font size="1">������}�j���A���Ɏg�p���Ă����ʂ́A�J�����̂��̂ɂ��A���i�łƈꕔ�قȂ�ꍇ������܂��B</font></span>
		</td>
	</tr>
	<tr>
		<td width="100%" height="12" align="center" background="img/sita.gif">
		<img src="img/sp.gif">
		</td>
	</tr>
</table>

<input type="hidden" name="txtLinkNo" value="">
</form>

</body>

</html>