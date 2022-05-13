
w_sSql = ""
w_sSql = w_sSql & " SELECT "
w_sSql = w_sSql & "		T120_MISYU_GAKUNEN, "
w_sSql = w_sSql & "		T120_KAMOKU_CD, "
w_sSql = w_sSql & "		T120_KAMOKUMEI "
w_sSql = w_sSql & " FROM "
w_sSql = w_sSql & "		T120_SAISIKEN, "
w_sSql = w_sSql & "		M08_HYOKAKEISIKI "
w_sSql = w_sSql & "	WHERE "

w_sSql = w_sSql & "		    T120_KYOUKAN_CD = '" & m_sKyokanCd & "'"
w_sSql = w_sSql & "		AND ( T120_SYUTOKU_NENDO Is Null or T120_SYUTOKU_NENDO = " & Session("NENDO") & " ) "
w_sSql = w_sSql & "		AND T120_HYOKA_FUKA_KBN = 1 "

'TableÇÃåãçá
w_sSql = w_sSql & "		AND M08_NENDO = T120_NENDO"

'ì_êîÇÃçiÇËçûÇ›
w_sSql = w_sSql & " 	AND M08_HYOUKA_NO = 2 "
w_sSql = w_sSql & "		AND M08_HYOKA_TAISYO_KBN = 0 "
w_sSql = w_sSql & "		AND M08_HYOKA_SYOBUNRUI_CD = 4 "
w_sSql = w_sSql & " 	AND T120_SEISEKI <= M08_MAX "
w_sSql = w_sSql & " 	AND T120_SEISEKI >= M08_MIN "

w_sSql = w_sSql & " GROUP BY "
w_sSql = w_sSql & "		T120_MISYU_GAKUNEN, "
w_sSql = w_sSql & "		T120_KAMOKU_CD, "
w_sSql = w_sSql & "		T120_KAMOKUMEI "
