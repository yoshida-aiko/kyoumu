<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ʋ����\��
' ��۸���ID : web/web0300/calender.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:	SESSION("KYOKAN_CD"):����CD
'            	SESSION("NENDO")	:�N�x
'				TUKI				:��
'				cboKyositu			:����CD
'
' ��      �n:	hidDay     :���ɂ�
'				hidYear    :�N
'				hidMonth   :��
'				hidKyositu :����CD
' ��      ��:
'           �������\��
'               �I�����ꂽ���̃J�����_�[��\��
'           �����t�N���b�N��
'               ���̃t���[���ɑI�����ꂽ���t�̋�������\��
'-------------------------------------------------------------------------
' ��      ��: 2001/08/06 �ɓ����q
' ��      �X:
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_iSyoriNen          '��������
    Public m_iKyokanCd          '�N�x
    Public m_iTuki              '//��
	Public m_iKyosituCd			'//����CD
	Public m_sKyosituName		'//��������
	Public m_SDate
	Public m_EDate
	Public m_sDay    	'//��
	Public m_sKyokanNm

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
    w_sMsgTitle="�����o������"
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
'Call s_DebugPrint

		'//�������擾
		w_iRet = f_GetKyousituName()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//���t���擾
        w_iRet = f_GetDate()
        If w_iRet <> 0 Then
            m_bErrFlg = True
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
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen  = ""
    m_iKyokanCd  = ""
    m_iTuki      = ""
	m_sKyokanNm  = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSyoriNen  = Session("NENDO")    
    'm_iKyokanCd  = Session("KYOKAN_CD")
    m_iKyokanCd  = Request("SKyokanCd1")

    m_iTuki      = Request("TUKI")
	m_iKyosituCd = Request("cboKyositu")
	m_sDay       = Request("hidDay")
	m_sKyokanNm  =Request("SKyokanNm1")

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen  = " & m_iSyoriNen  & "<br>"
    response.write "m_iKyokanCd  = " & m_iKyokanCd  & "<br>"
    response.write "m_iTuki      = " & m_iTuki      & "<br>"
    response.write "m_iKyosituCd = " & m_iKyosituCd & "<br>"
    response.write "m_sDay       = " & m_sDay       & "<br>"
    response.write "m_sKyokanNm  = " & m_sKyokanNm  & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  ���t�f�[�^���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDate()

	Dim w_iRet
	Dim w_sSQL
	Dim rs
	Dim w_sSDate
	Dim w_sEDate

	On Error Resume Next
	Err.Clear

	f_GetDate = 1

	Do

		'//�s�����׃e�[�u�����J�����_�[�f�[�^���擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T32.T32_HIDUKE, "
		w_sSQL = w_sSQL & vbCrLf & "  T32.T32_YOUBI_CD"
		w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M T32"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      T32.T32_NENDO=" & cInt(m_iSyoriNen)
        w_sSQL = w_sSQL & vbCrLf & "  AND SUBSTR(T32.T32_HIDUKE,6,2)='" & gf_fmtZero(m_iTuki,2) & "'"
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "  T32.T32_HIDUKE,T32.T32_YOUBI_CD"

'response.write w_sSQL & "<BR>"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDate = 99
			Exit Do
		End If

		If rs.EOF = False then
			rs.MoveFirst
			m_SDate = rs("T32_HIDUKE")
			rs.MoveLast
			m_EDate = rs("T32_HIDUKE")
		End If

		'//����I��
		f_GetDate = 0
		Exit Do

	Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKyousituName()

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetKyousituName = 1

    Do
		'//�������擾
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU.M06_KYOSITUMEI"
		w_sSql = w_sSql & vbCrLf & " FROM M06_KYOSITU"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M06_KYOSITU.M06_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND M06_KYOSITU.M06_KYOSITU_CD=" & m_iKyosituCd

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKyousituName = 99
            Exit Do
        End If

		If rs.EOF = False Then
			m_sKyosituName = rs("M06_KYOSITUMEI")
		End If

        '//����I��
        f_GetKyousituName = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �Y�����t�ɗ\�肪�����Ă��邩�ǂ����ɂ��ATD��COLOR��Ԃ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_KyosituYoteiInfo(p_iJigenCnt,p_sDay,p_sColor)

    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_iYoyakCnt

    On Error Resume Next
    Err.Clear

    f_KyosituYoteiInfo = 1
	p_sColor = ""

    Do
		'//�������擾
		'w_sSql = ""
		'w_sSql = w_sSql & vbCrLf & " SELECT "
		'w_sSql = w_sSql & vbCrLf & "  COUNT(*) AS CNT"
		'w_sSql = w_sSql & vbCrLf & " FROM "
		'w_sSql = w_sSql & vbCrLf & "  T58_KYOSITU_YOYAKU T58"
		'w_sSql = w_sSql & vbCrLf & " WHERE "
		'w_sSql = w_sSql & vbCrLf & "  T58.T58_NENDO=" & m_iSyoriNen
		'w_sSql = w_sSql & vbCrLf & "  AND T58.T58_HIDUKE='" & gf_YYYY_MM_DD(p_sDay,"/") & "' "
		'w_sSql = w_sSql & vbCrLf & "  AND T58.T58_KYOSITU=" & m_iKyosituCd

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T58.T58_YOUBI_CD,"
		w_sSql = w_sSql & vbCrLf & "  T58.T58_JIGEN"
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T58_KYOSITU_YOYAKU T58"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T58.T58_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND T58.T58_HIDUKE='" & gf_YYYY_MM_DD(p_sDay,"/") & "' "
		w_sSql = w_sSql & vbCrLf & "  AND T58.T58_KYOSITU=" & m_iKyosituCd
		w_sSql = w_sSql & vbCrLf & "  GROUP BY"
		w_sSql = w_sSql & vbCrLf & "  T58.T58_YOUBI_CD,"
		w_sSql = w_sSql & vbCrLf & "  T58.T58_JIGEN"

'response.write w_sSQL & "<br>"
        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_KyosituYoteiInfo = 99
            Exit Do
        End If

		'w_iYoyakCnt = rs("CNT")
		w_iYoyakCnt = 0
		Do Until rs.EOF

			w_iYoyakCnt = w_iYoyakCnt +1	
		
			RS.MoveNext
		Loop
		'w_iYoyakCnt = rs.RecordCount		

'response.write w_iYoyakCnt

		If cint(w_iYoyakCnt) = 0 Then
			'//�\�񂪓����Ă��Ȃ�
			p_sColor = ""
		Else

			If cint(w_iYoyakCnt) >= cint(p_iJigenCnt) Then
				'//�S�Ă̎����ɗ\�񂪓����Ă���
				p_sColor = "FILLFULL"
			Else
				'//�ꕔ�̎����ɗ\�񂪓����Ă���
				p_sColor = "FILLPART"
			End If
		End If

        '//����I��
        f_KyosituYoteiInfo = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �J�����_�[���쐬
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_MakeCalendar()

	Dim myTable()
	Dim w_sTdColor

	On Error Resume Next
	Err.Clear

	f_MakeCalendar = 1

	Do

		myDate    = m_SDate

		myWeekTbl = split("��,��,��,��,��,��,�y",",")
		myMonthTbl= split("31,28,31,30,31,30,31,31,30,31,30,31",",")
		myYear = year(m_SDate)

		'//�[�N����
		If (myYear Mod 4=0 And myYear Mod 100<>0 ) or (myYear Mod 400=0) Then
			myMonthTbl(1) = 29	'//2��
		End If

		myMonth = m_iTuki
		myWeek = cint(Weekday(myYear & "/" & myMonth & "/01"))-1	'//�������̗j�����擾
		myTblLine = int(((myWeek+myMonthTbl(myMonth-1))/7)+0.9)		'//�s�����擾

		ReDim myTable(7*myTblLine)

		'//������
		For i=0 To 7*myTblLine-1
			myTable(i)="�@"
		Next

		'//���t���i�[
		For i = 0 to myMonthTbl(myMonth-1)-1
			myTable(i+myWeek)=i+1
		Next

		'// ***********************
		'// **  �J�����_�[�̕\��
		'// ***********************

		'//�����̍ő�l���擾(�����̗\��󋵂��擾���邽��)
		w_iRet = f_GetJigen(w_iJigenCnt)
		If w_iRet <> 0 Then
			Exit Do
		End If

		'response.write("<table border='1' class='hyo' width='98%'  >")
		response.write("<table border='1' class='hyo' width='80%'  >")
		response.write("<tr>")

		'=============
		'�w�b�_��
		'=============
		'//�j����\��
		For i = 0 to 6
		   response.write("<th align='center' class='header'>")
		   response.write(myWeekTbl(i))
		   response.write("</th>")
		Next
		response.write("</tr>")

		'=============
		'���ו�
		'=============
		'//���ɂ���\��
		For i = 0 to myTblLine-1

			'//���ټ�Ă̸׽���Z�b�g
			Call gs_cellPtn(w_Class)
		   response.write("<tr>")

		   For j=0 To 7-1
		    myDat = myTable(j+(i*7))

			'//TD�F���
			w_sTdClassColor=w_Class

			If myDat <> "�@" Then
				'=============================================================
				'//�Y�����t�ɗ\�肪�����Ă��邩�ǂ����ɂ��ATD��COLOR��Ԃ�
				w_sDay = myYear & "/" & myMonth & "/" & myDat

				w_sColor = ""
				w_iRet = f_KyosituYoteiInfo(w_iJigenCnt,w_sDay,w_sColor)
				If w_iRet <> 0 Then
					Exit Do
				End If

				If w_sColor <> "" Then
					w_sTdClassColor=w_sColor
				End If
				'=============================================================
			End If

		    response.write("<td align='center' class='" + w_sTdClassColor + "' > ")
			If myDat="�@" Then
			    response.write("�@")
			Else
				If m_sDay<> "" Then

					If cint(m_sDay) = myDat Then
						'response.write("<span class='select_date'>" & myDat & "</span>")
						response.write("<b>" & myDat & "</b>")
					Else
						response.write("<A HREF='javascript:f_ListClick(" & myDat & ")'>" & myDat & "</A>")
					End If
				Else
					response.write("<A HREF='javascript:f_ListClick(" & myDat & ")'>" & myDat & "</A>")
				End If

			End If

		    response.write("</td>")
			Next
		   response.write("</tr>")
		Next

		response.write("</table>")

		'//����I��
		f_MakeCalendar = 0
		Exit Do

	Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �������̍ő�l�ƍŏ��l���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetJigen(p_iCnt)

    Dim w_iRet
    Dim w_sSQL
    Dim rs

    On Error Resume Next
    Err.Clear

    f_GetJigen = 1

    Do

		'w_sSql = ""
		'w_sSql = w_sSql & vbCrLf & " SELECT "
		'w_sSql = w_sSql & vbCrLf & "  MAX(T20_JIKANWARI.T20_JIGEN) AS MAX "
		'w_sSql = w_sSql & vbCrLf & " FROM T20_JIKANWARI"
		'w_sSql = w_sSql & vbCrLf & " WHERE "
		'w_sSql = w_sSql & vbCrLf & "      T20_JIKANWARI.T20_NENDO=" & m_iSyoriNen
		'w_sSql = w_sSql & vbCrLf & "  AND T20_JIKANWARI.T20_GAKKI_KBN=" & Session("GAKKI")

		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  MAX(m07_JIGEN.m07_JIKAN) AS MAX "
		w_sSql = w_sSql & vbCrLf & " FROM m07_JIGEN"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "      m07_JIGEN.M07_NENDO=" & m_iSyoriNen
		'w_sSql = w_sSql & vbCrLf & "  AND T20_JIKANWARI.T20_GAKKI_KBN=" & Session("GAKKI")

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetJigen = 99
            Exit Do
        End If

		If ISNULL(rs("MAX")) Then
			p_iCnt = 0
		Else
			p_iCnt = rs("MAX")
		End If

        '//����I��
        f_GetJigen = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
    Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
%>
    <html>
    <head>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <title>���ʋ����\��</title>

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
    //  [�@�\] �J�����_�[���t�N���b�N��
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_ListClick(p_date){

		var wArg

		//�J�����_�[�y�[�W���ĕ\��
		wArg = ""
		wArg = wArg + "?TUKI=<%=m_iTuki%>"
		wArg = wArg + "&cboKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&hidDay="+p_date
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=Server.URLEncode(m_iKyokanCd)%>"

		parent.middle.location.href="./calendar.asp"+wArg
		//parent.middle.location.href="./calendar.asp?TUKI=<%=m_iTuki%>&cboKyositu=<%=m_iKyosituCd%>&hidDay="+p_date

		//���X�g�y�[�W���ĕ\��
		wArg = ""
		wArg = wArg + "?hidDay="+p_date
		wArg = wArg + "&hidYear=<%=year(m_SDate)%>"
		wArg = wArg + "&hidMonth=<%=month(m_SDate)%>"
		wArg = wArg + "&hidKyositu=<%=m_iKyosituCd%>"
		wArg = wArg + "&SKyokanNm1=<%=Server.URLEncode(request("SKyokanNm1"))%>"
		wArg = wArg + "&SKyokanCd1=<%=Server.URLEncode(m_iKyokanCd)%>"

		parent.bottom.location.href="./web0300_lst.asp"+wArg

	}

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <br>
    <form name="frm" method="post">
<%
'//�f�o�b�O
'Call s_DebugPrint()
%>

	<center>

    <table class="hyo" border="1" width="80%">

        <tr>
            <th class="header" width="20%" align="center" nowrap><font size="2">���p��</font></th>
            <td class="detail" width="80%" align="left"   nowrap colspan="2"><font size="2"><%=m_sKyokanNm%></font></td>
        </tr>
        <tr>
            <th class="header" width="20%" align="center" nowrap><font size="2">����</font></th>
            <td class="detail" width="40%" align="left"   nowrap><font size="2"><%=m_sKyosituName%></font></td>
            <td class="detail" width="40%" align="center" nowrap><font size="2"><%=year(m_SDate)%>�N�@<%=Month(m_SDate)%>��</font></td>
        </tr>
    </table>
	<br>

	<%
	'//�J�����_�[�\��
	Call f_MakeCalendar()
	%>
	<table width="80%" border=0><tr>
	<td align="right" nowrap>
		<span class="msg" ><font size="2">���ԕ\���F�S�Ė��܂��Ă��܂��B<br>���\���F�ꕔ���܂��Ă��܂��B</font></span>
	</td>
	</tr></table>

	<!--�l�n�p-->
	<input type="hidden" name="TUKI"       value="<%=m_iTuki%>">
	<input type="hidden" name="cboKyositu" value="<%=m_iKyosituCd%>">
	<input type="hidden" name="SKyokanNm1" value="<%=Server.HTMLEncode(request("SKyokanNm1"))%>">
	<input type="hidden" name="SKyokanCd1" value="<%=m_iKyokanCd%>">

	<input type="hidden" name="hidDay"     value="">
	<input type="hidden" name="hidYear"    value="<%=year(m_SDate) %>">
	<input type="hidden" name="hidMonth"   value="<%=month(m_SDate)%>">
	<input type="hidden" name="hidKyositu" value="<%=m_iKyosituCd%>">

	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>
