
<%
'----------------------------------------
'�t�@�C���̃f�[�^��\��
'----------------------------------------

'--- �I�u�W�F�N�g�쐬 ---
Set w_oFile = Server.CreateObject("Scripting.FileSystemObject")

'--- �t�@�C�����J���i�ǂݎ���p�A�t�@�C�������݂��Ȃ��Ƃ��͐V�K�쐬�j ---
Set w_oCSV = w_oFile.OpenTextFile(Server.MapPath("Sample.csv"),1,True)

'--- �t�@�C���̃f�[�^��\�� ---
Do Until w_oCSV.AtEndofStream
    Response.Write "<p>" & w_oCSV.ReadLine
Loop

'--- �t�@�C������� ---
w_oCSV.Close

'--- �I�u�W�F�N�g��� ---
Set w_oCSV = Nothing
Set w_oFile = Nothing

%>
