
<!--#INCLUDE FILE="include02.asp"-->

<%

' SQL�����Z�b�V�����ϐ��Ɋi�[ �iGO �� EndCSV.asp�j
    Session.Contents("g_CSV")=w_sSQL
    
' �Y������Ј������邩�ǂ����̔���
	if g_rRs.EOF=true then
		w_sFLG="3"
	end if

' �����̎w�肪�����ꍇ�i���ׂďo�͂��邩�ǂ����̃��b�Z�[�W�j
	if w_sStartCD = "" AND w_sEndCD = "" AND w_sName = "" AND w_checkDel <> 1 then
		w_sFLG="2"
	end if
	
' �m�F���b�Z�[�W
	Session.Contents("SELECT")="CSV"
	Response.Redirect "CorEKAKUNIN.asp?FLG=" & w_sFLG

    w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing

%>

