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
' ��      �X: 
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

<table class=manual cellspacing="0" cellpadding="0" width=504 height=100% border="0">
	<tr>
		<td class=manual colspan="2" width="504" height="35%" align="center">
		<img src="../image/title.gif" width="504" height="214">
		</td>
	</tr>
	<tr>
		<td class=manual align="right" valign="top" height="65%" width="504">

			<img src="../image/sp.gif" height="15"><br>

			<table border="0" width=100% cellspacing="0" cellpadding="0">
			<tr>
			<td class=manual colspan="3">
				<table width=100% bgcolor="#3A449E" cellspacing="0" cellpadding="0" border="0">
					<tr>
						<td align="center"><font size="3" color="#ffffff"><b>����}�j���A��</b></font></td>
					</tr>
				</table>
			<img src="../image/sp.gif" height="5"><br>

			</td>
			</tr>

			<tr>
			<td valign="top">
				<table width=156 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>1�E�V�X�e���T�v</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(1)">�V�X�e���T�v</a>
						</td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif"></td>
					</tr>
					<tr>
						<td class=manual><b>2�E�V�X�e����{����</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(2)">�L�[�{�[�h����</a>
						</td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(3)">�}�E�X����</a>
						</td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(4)">��ʑ���</a>
						</td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif"></td>
					</tr>
					<tr>
						<td class=manual><b>3�E���O�C��</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(5)">���O�C��</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif"></td>
					</tr>
					<tr>
						<td class=manual><b>4�E���C�����j���[</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(6)">���C�����j���[</a></td>
					</tr>
				</table>
			</td>
			<td valign="top">

				<table width=174 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>5�E�o������</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(7)">���Əo������</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(8)">�s���o������</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(9)">�����o������</a></td>
					</tr>
					<tr>
						<td class=manual><b>6�E�����E����</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(10)">�������{�Ȗړo�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(11)">�����ēƏ��\���o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(12)">���ѓo�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(13)">�������Ԋ��i�N���X�ʁj</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(14)">�������ԋ����\��ꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(15)">���шꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(16)">�l�ʐ��шꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(17)">���N�Y���҈ꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><b>7�E�X�P�W���[��</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(18)">�s�������ꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(19)">�N���X�ʎ��Ǝ��Ԉꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(20)">�����ʎ��Ǝ��Ԉꗗ</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(21)">���Ԋ������A��</a></td>
					</tr>
				</table>
			</td>
			<td valign="top">
				<table width=174 cellspacing="1" cellpadding="1" border="0">
					<tr>
						<td class=manual><b>8�E���̑�����</b></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(22)">�i�H����o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(23)">�g�p���ȏ��o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(24)">�w���v�^�������o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(25)">�e��ψ��o�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(26)">�l���C�I���Ȗړo�^</a></td>
					</tr>
					<tr>
						<td class=manual><img src="../image/sp.gif" width="5" height="1">
						<a href="javascript:f_PageClick(27)">�����������o�^</a></td>
					</tr>
					<tr>
						<td class=manual><b>9�E��񌟍�</b></td>
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
						<a href="javascript:f_PageClick(32)">�󂫎���</a></td>
					</tr>
					<tr>
						<td class=manual><b>10�E�x���@�\</b></td>
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

<input type="hidden" name="txtLinkNo" value="">
</form>

</body>

</html>