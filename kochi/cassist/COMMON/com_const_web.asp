<%
'/************************************************************************
' システム名    :   教務事務システム
' 処  理  名    :   定数定義
' ﾌﾟﾛｸﾞﾗﾑID     :   COM_const
' 機      能    :   使用定数の定義
'-------------------------------------------------------------------------
' 作      成: 2001/07/13 谷脇
' 変      更: 2001/07/18 根本
' 変      更: 2001/07/19 伊藤　置換科目ﾌﾗｸﾞを追加
' 変      更: 2001/07/31 根本　資本金桁数を追加
'*************************************************************************/

'******************************************
'システム関連
'******************************************
Public const C_M00_NENDO = 9999
Public Const C_School_CD = 1            '高専名
Public Const C_LEVEL_NOCHK = "XXXXXXX"  '権限ﾁｪｯｸをしない

'******************************************
'表示関連
'******************************************
Public const C_PAGE_LINE = 10      '検索リストの表示件数


'//デバッグ用
Public const C_RetURL = "/cassist/"
'Public const C_RetURL = "/catest/"

Public const C_IMAGE_DIR = "/cassist/image/"   '
Public const C_CELL1 = "CELL2"  'リストを交互に出すためのセル設定１。
Public const C_CELL2 = "CELL1"  'リストを交互に出すためのセル設定２。

Public const C_MAIN_FRAME = "fTopMain"
Public const C_LOGIN_FLG = 1	'// ログイン画面からきたしるし

Public const C_TABLE_WIDTH = "90%"  'テーブル幅

Public Const C_SYORYAKU_KETA = 4    '//表示時に省略する桁数（資本金）
Public Const C_ERR_RETURL = "login/default.asp"		'//Err時に戻るURL

'******************************************
'表示関連（授業時間一覧）'//2001/07/18根本追加
'******************************************
Public Const C_YOUBI_MIN = 2              '曜日コード（始）
Public Const C_YOUBI_MAX = 6              '曜日コード（終）

'******************************************
'学籍関連
'******************************************
Public const C_GAKU_KETA = 9    '学生番号の桁数。(最大１０桁)

'******************************************
'試験関連
'******************************************
Public const C_SIKEN_KIKAN = 0		'試験準備期間
Public const C_JISSI_KIKAN = 1		'試験実施期間
Public const C_SEISEKI_KIKAN = 2	'成績入力期間
Public const C_SIKEN_CODE_NULL = "0" '試験コードヌル値
'****************************************
'履修関連
'****************************************
'置換科目ﾌﾗｸﾞ
Public Const C_TIKAN_TUJO = 0     '置換なし(通常)

'メイン教官フラグ
Public Const C_MAIN_KYOKAN_NO = 0
Public Const C_MAIN_KYOKAN_YES = 1

'成績入力教官フラグ
Public Const C_SEISEKI_INP_FLG_NO = 0
Public Const C_SEISEKI_INP_FLG_YES = 1

'****************************************
'ツール関連
'****************************************
'連絡登録
Public Const C_KAKU_MI = 0      '未確認
Public Const C_KAKU_SUMI = 1    '確認済

'****************************************
'学籍関連（項目別権限）
'2001.12.3  大幅変更 - 岡田
'2002.05.10 一部追加 - 持永
'****************************************
'基本情報
Public Const C_T13_GAKUSEI_NO		=   4   '学生番号               
Public Const C_T11_SIMEI            =   6   '氏名（漢字）           
Public Const C_T11_SIMEI_ROMA       =   7   '氏名（ローマ字）       
Public Const C_T11_SIMEI_KD         =   8   '氏名（カナ）濁点あり   
Public Const C_T11_SIMEI_KYU        =   11  '旧氏名（漢字）      	
Public Const C_T11_SIMEI_ROMA_KYU   =   12  '旧氏名（ローマ字） 　　
Public Const C_T11_SIMEI_KD_KYU     =   13  '旧氏名（カナ）濁点あり 
Public Const C_T11_KAIMEI_DATE		=	141	'最終改姓名日			
Public Const C_T11_HON_ZIP          =   20  '本籍郵便番号           
Public Const C_T11_HON_JUSYO        =   22  '本籍住所               
Public Const C_T11_GEN_ZIP          =   25  '現住所郵便番号         
Public Const C_T11_GEN_JUSYO        =   26  '現住所                 
Public Const C_T11_GEN_TEL          =   28  '現住所電話番号         
Public Const C_T11_KIN_TEL			=	174 '緊急連絡先

'個人情報
Public Const C_T11_SEIBETU          =   16  '性別区分               
Public Const C_T11_SEINENBI         =   17  '生年月日               
Public Const C_T11_KETUEKI          =   18  '血液型区分             
Public Const C_T11_RH               =   19  'ＲＨ区分               
Public Const C_T11_HOG_SIMEI        =   32  '保護者氏名（漢字）     
Public Const C_T11_HOG_SIMEI_K      =   33  '保護者氏名（カナ）     
Public Const C_T11_HOG_ZOKU         =   34  '保護者続柄区分         
Public Const C_T11_HOG_ZIP          =   35  '保護者住所郵便番号     
Public Const C_T11_HOG_JUSYO        =   37  '保護者住所             
Public Const C_T11_HOG_TEL          =   38  '保護者電話番号         
Public Const C_T11_HOS_SIMEI        =   48  '保証人氏名(漢字)       
Public Const C_T11_HOS_SIMEI_K      =   49  '保証人氏名(カナ)       
Public Const C_T11_HOS_ZOKU         =   50  '保証人続柄区分         
Public Const C_T11_HOS_ZIP          =   51  '保証人住所郵便番号     
Public Const C_T11_HOS_JUSYO        =   53  '保証人住所             
Public Const C_T11_HOS_TEL          =   54  '保証人電話番号         
Public Const C_T11_SYUSSINKO		=	175 '出身校
Public Const C_T11_SYUSSINKOKU      =   61  '出身国              	
Public Const C_T11_RYUGAKU_KBN      =   62  '留学区分             　
Public Const C_T11_KAZOKU_1         =   68  '家族名称１				
Public Const C_T11_KAZOKU_ZOKU_1    =   69  '家族続柄区分１			
Public Const C_T11_KAZOKU_2         =   70  '家族名称２				
Public Const C_T11_KAZOKU_ZOKU_2    =   71  '家族続柄区分２　　　　 
Public Const C_T11_KAZOKU_3         =   72  '家族名称３				
Public Const C_T11_KAZOKU_ZOKU_3    =   73  '家族続柄区分３			
Public Const C_T11_KAZOKU_4         =   74  '家族名称４				
Public Const C_T11_KAZOKU_ZOKU_4    =   75  '家族続柄区分４			
Public Const C_T11_KAZOKU_5         =   76  '家族名称５				
Public Const C_T11_KAZOKU_ZOKU_5    =   77  '家族続柄区分５			
Public Const C_T11_KAZOKU_6         =   78  '家族名称６				
Public Const C_T11_KAZOKU_ZOKU_6    =   79  '家族続柄区分６			
Public Const C_T11_KAZOKU_7         =   80  '家族名称７				
Public Const C_T11_KAZOKU_ZOKU_7    =   81  '家族続柄区分７			
Public Const C_T11_KAZOKU_8         =   82  '家族名称８				
Public Const C_T11_KAZOKU_ZOKU_8    =   83  '家族続柄区分８			
Public Const C_T11_HOG_SEINEIBI		=	172 '保護者生年月日
Public Const C_T11_HOS_SEINEIBI		=	173 '保証人生年月日
Public Const C_T11_KAZOKU_SEINEIBI_1=	176 '家族生年月日１
Public Const C_T11_KAZOKU_SEINEIBI_2=	177 '家族生年月日２
Public Const C_T11_KAZOKU_SEINEIBI_3=	178 '家族生年月日３
Public Const C_T11_KAZOKU_SEINEIBI_4=	179 '家族生年月日４
Public Const C_T11_KAZOKU_SEINEIBI_5=	180 '家族生年月日５
Public Const C_T11_KAZOKU_SEINEIBI_6=	181 '家族生年月日６
Public Const C_T11_KAZOKU_SEINEIBI_7=	182 '家族生年月日７
Public Const C_T11_KAZOKU_SEINEIBI_8=	183 '家族生年月日８

'入学情報
Public Const C_T11_NYUNENDO         =   1   '入学年度    			
Public Const C_T11_TYUGAKKO_CD      =   137 '中学校名               
'Public Const C_T11_SYUSSINKO        =   172 '出身校                 
Public Const C_T13_TYUSOTUGYOBI     =   140 '中学校卒業日           
Public Const C_T11_NYUGAKU_KBN      =   29  '入学区分               
Public Const C_T11_NYU_GAKKA        =   30  '入学学科               
Public Const C_T11_NYUGAKUBI        =   31  '入学年月日             
Public Const C_T11_JUKEN_NO         =   39  '受験番号               
Public Const C_T11_TYU_CLUB         =   57  '中学校クラブ活動       
Public Const C_T11_TYU_CLUB_SYOSAI  =   58  '中学校クラブ活動詳細   
Public Const C_T11_NYU_SEISEKI      =   63  '入試成績               

'学年情報
Public Const C_T13_GAKUSEKI_NO      =   85  '学籍番号               
Public Const C_T13_ZAISEKI_KBN      =   86  '在籍区分               
Public Const C_T13_GAKKA_CD         =   87  '学科コード             
Public Const C_T13_COURCE_CD        =   88  '所属コース          　 
Public Const C_T13_GAKUNEN          =   89  '学年                   
Public Const C_T13_CLASS            =   90  'クラス                 
Public Const C_T13_SYUSEKI_NO1      =   91  '出席番号１             
Public Const C_T13_SYUSEKI_NO2      =   92  '出席番号２             
Public Const C_T13_RYOSEI_KBN       =   93  '入寮区分               
Public Const C_T13_RYUNEN_FLG       =   94  '留年区分               
Public Const C_T13_CLUB_1           =   122 'クラブ活動1            
Public Const C_T13_CLUB_1_NYUBI		=	142	'クラブ活動1入部日		
Public Const C_T13_CLUB_2           =   123 'クラブ活動2            
Public Const C_T13_CLUB_2_NYUBI		=	143 'クラブ活動2入部日
Public Const C_T13_NENSYOKEN        =   129 '年毎所見 				
Public Const C_T13_TOKUKATU         =   184 '特別活動
Public Const C_T13_TOKUKATU_DET     =   131 '特別活動詳細           
Public Const C_T13_SINTYO           =   132 '身長                   
Public Const C_T13_TAIJYU           =   133 '体重                   
Public Const C_T13_SEKIJI_TYUKAN_Z	=	144	'前期中間席次
Public Const C_T13_SEKIJI_KIMATU_Z 	=	145	'前期期末席次
Public Const C_T13_SEKIJI_TYUKAN_K 	=	146	'後期中間席次
Public Const C_T13_SEKIJI			=	120	'学年末席次
Public Const C_T13_NINZU_TYUKAN_Z	=	148	'前期中間クラス人数
Public Const C_T13_NINZU_KIMATU_Z  	=	149	'前期期末クラス人数
Public Const C_T13_NINZU_TYUKAN_K  	=	150	'後期中間クラス人数
Public Const C_T13_CLASSNINZU		=	151	'学年末クラス人数
Public Const C_T13_HEIKIN_TYUKAN_Z 	=	152	'前期中間平均点
Public Const C_T13_HEIKIN_KIMATU_Z 	=	153	'前期期末平均点
Public Const C_T13_HEIKIN_TYUKAN_K 	=	154	'後期中間平均点
Public Const C_T13_HEIKIN_KIMATU_K 	=	155	'学年末平均点
Public Const C_T13_SUMJYUGYO		=	156	'総授業日数
Public Const C_T13_SUMSYUSSEKI		=	126	'出席日数
Public Const C_T13_SUMRYUGAKU		=	128	'留学中の授業日数    
Public Const C_T13_KESSEKI_TYUKAN_Z	=	159	'前期中間欠席日数
Public Const C_T13_KESSEKI_KIMATU_Z	=	160	'前期期末欠席日数
Public Const C_T13_KESSEKI_TYUKAN_K	=	161	'後期中間欠席日数
Public Const C_T13_SUMKESSEKI		=	125	'学年末欠席日数
Public Const C_T13_KIBIKI_TYUKAN_Z	=	163	'前期中間忌引日数
Public Const C_T13_KIBIKI_KIMATU_Z	=	164	'前期期末忌引日数
Public Const C_T13_KIBIKI_TYUKAN_K	=	165	'後期中間忌引日数
Public Const C_T13_SUMKIBTEI		=	127	'出席停止･忌引き日数        
Public Const C_T_CLASSIIN			=	167	'クラス役員
Public Const C_T_TANNIN				=	168	'担任名
Public Const C_T_JYUGYORYOMENJYO	=	169	'授業料免除
Public Const C_T_SYOGAKUKIN			=	170	'奨学金

'その他予備情報
Public Const C_T_JIYUSENTAKU		=	46 '自由選択項目

'総合所見
'Public Const C_T11_SOGOSYOKEN       =   46  '総合所見               
Public Const C_T13_IDOU_NUM			=	171	'異動回数
Public Const C_T13_IDOU_KBN         =   96  '異動区分1              
Public Const C_T13_IDOU_BI          =   97  '異動年月日1            
Public Const C_T13_IDOU_BIK         =   98  '異動備考1              
Public Const C_T13_IDOU_KBN2        =   99  '異動区分2              
Public Const C_T13_IDOU_BI2         =   100 '異動年月日2            
Public Const C_T13_IDOU_BIK2        =   101 '異動備考2              
Public Const C_T13_IDOU_KBN3        =   102 '異動区分3              
Public Const C_T13_IDOU_BI3         =   103 '異動年月日3            
Public Const C_T13_IDOU_BIK3        =   104 '異動備考3              
Public Const C_T13_IDOU_KBN4        =   105 '異動区分4              
Public Const C_T13_IDOU_BI4         =   106 '異動年月日4            
Public Const C_T13_IDOU_BIK4        =   107 '異動備考4              
Public Const C_T13_IDOU_KBN5        =   108 '異動区分5              
Public Const C_T13_IDOU_BI5         =   109 '異動年月日5            
Public Const C_T13_IDOU_BIK5        =   110 '異動備考5              
Public Const C_T13_IDOU_KBN6        =   111 '異動区分6              
Public Const C_T13_IDOU_BI6         =   112 '異動年月日6            
Public Const C_T13_IDOU_BIK6        =   113 '異動備考6              
Public Const C_T13_IDOU_KBN7        =   114 '異動区分7              
Public Const C_T13_IDOU_BI7         =   115 '異動年月日7            
Public Const C_T13_IDOU_BIK7        =   116 '異動備考7              
Public Const C_T13_IDOU_KBN8        =   117 '異動区分8              
Public Const C_T13_IDOU_BI8         =   118 '異動年月日8            
Public Const C_T13_IDOU_BIK8        =   119 '異動備考8              
Public Const C_T13_IDOU_ENDBI       =   185 '異動終了日             


'未使用区分-----------------------------------------------------------
Public Const C_T11_SINRO            =   40  '卒業後進路      
Public Const C_T11_SOTUKEN_DAI      =   42  '卒研論題            
Public Const C_T11_SOTU_KYOKAN_CD1  =   43  '卒研教官１（教官コード）
Public Const C_T11_SOTU_KYOKAN_CD2  =   44  '卒研教官２（教官コード）
Public Const C_T11_SOTU_KYOKAN_CD3  =   45  '卒研教官３（教官コード）
Public Const C_T11_KODOSYOKEN       =   55  '行動所見               
Public Const C_T11_SYUMITOKUGI      =   56  '趣味･特技･資格取得     
Public Const C_T11_RYO_KIBO         =   59  '入寮希望区分            
'Public Const C_T11_NYU_GAKUNEN		=	64	'入学学年
Public Const C_T11_TENNYUNEND       =   65  '転入入学年度            
Public Const C_T13_NENDO            =   84  '処理年度           
'Public Const C_T13_SUMRYUGAKU       =   128 '留学中の授業日数    
'Public Const C_T13_TOKUKATU         =   130 '特別活動            
Public Const C_T13_ZAISEKI_END_KBN  =   134 '在籍区分(年度終わり)    
Public Const C_T13_NENBIKO          =   135 '年毎備考               
Public Const C_T11_TYOSA_BIK        =   136 '調査書備考             
'Public Const C_T13_SEKIJI           =   120 '学年末席次
'Public Const C_T13_SUMSYUSSEKI      =   126 '出席日数
'Public Const C_T13_SUMRYUGAKU       =   128 '留学中の授業日数    
'Public Const C_T13_SUMKESSEKI       =   125 '学年末欠席日数
'Public Const C_T13_SUMKIBTEI        =   127 '出席停止･忌引き日数
'------ 持永追加 ------
Public Const C_T11_KOJIN_BIK        =   138 '個人備考          
Public Const C_T11_SIMEI_GAIJI      =   139 '氏名外字          

'------ 金澤追加 '02/6/7 ------
Public Const C_M01_DAIBUNRUI150     =   150 '特別欠席          


'---------------------------------------------------------------------

'------ 前田追加 ------
'委員小分類コード
Public Const C_M34_SYOBUN_CD        =   0 
'選択科目グループ
Public Const C_T18_GRP              =   0
'単位数
Public Const C_T15_HAITO            =   0
'修得単位数
Public Const C_T18_SEL_GAKU         =   0   '学年決定
Public Const C_T18_SEL_TANI         =   0   '単位数決定
'----------------------

'****************************************
'自由選択項目関連
'****************************************
'** 定数定義 **
'分類（許可）
'Private Const C_BUNRUI_KYOKA = 1
'Private Const C_BUNRUIMEI_KYOKA = "自由項目"

'自由項目使用フラグ
Public Const C_JIYU_USE_YES = 1    '//使用する
Public Const C_JIYU_USE_NO = 0     '//使用しない

'自由項目タイプ
Public Const C_TYPECD_CHECK = 1  '//チェック
Public Const C_TYPECD_ZEN = 2    '//全角(漢字含む)
Public Const C_TYPECD_HAN = 3    '//半角(英数字)
Public Const C_TYPECD_NUM = 4    '//数値

'****************************************
'出欠入力関連
'****************************************
Public Const C_JIMU_FLG_NOTJIMU = "0"   '//事務ﾌﾗｸﾞ(事務以外で入力)
Public Const C_JIMU_FLG_JIMU = "1"      '//事務ﾌﾗｸﾞ(事務で入力)
Public Const C_TUKU_FLG_TUJO = "0"  '//時間割テーブル特別活動ﾌﾗｸﾞ(0:通常授業)
Public Const C_TUKU_FLG_TOKU = "1"  '//時間割テーブル特別活動ﾌﾗｸﾞ(1:特別活動(HR等))

'****************************************
'メッセージ関連
'****************************************
Public Const C_TOUROKU_KAKUNIN = "登録してもよろしいですか？"				'// 登録確認メッセージ
Public Const C_SAKUJYO_KAKUNIN = "削除してもよろしいですか？"				'// 削除確認メッセージ
Public Const C_TOUROKU_OK_MSG  = "登録が終了しました"						'// 登録終了メッセージ
Public Const C_SAKUJYO_OK_MSG  = "削除しました" 							'// 削除終了メッセージ
Public Const C_BRANK_VIEW_MSG  = "項目を選んで表示ボタンを押してください"   '// 空白ページメッセージ
Public Const C_UPDATE_OK_MSG   = "更新が終了しました"						'// 更新完了メッセージ

'****************************************
'日付関連
'****************************************
Public  Const C_NENDO_KAISITUKI = 4             '年度開始月

'****************************************
'アクセス権限関連
'****************************************
'//特別教室関連
Public  Const C_ACCESS_FULL   = "0"		'//アクセス権限FULLアクセス可
Public  Const C_ACCESS_NORMAL = "1"		'//アクセス権限一般
Public  Const C_ACCESS_VIEW   = "2"		'//アクセス権限参照のみ

'//使用教科書登録関連
Public  Const C_WEB0320_ACCESS_FULL   = "0"		'//アクセス権限FULLアクセス可
Public  Const C_WEB0320_ACCESS_NORMAL = "1"		'//アクセス権限専門教官

'//個人履修選択科目決定関連
Public  Const C_WEB0340_ACCESS_FULL   = "0"		'//アクセス権限FULLアクセス可
Public  Const C_WEB0340_ACCESS_SENMON = "1"		'//アクセス権限専門教官
Public  Const C_WEB0340_ACCESS_TANNIN = "2"		'//アクセス権限担任

'//レベル別科目決定関連
Public  Const C_WEB0390_ACCESS_FULL   = "0"		'//アクセス権限FULLアクセス可
Public  Const C_WEB0390_ACCESS_SENMON = "1"		'//アクセス権限種別メイン教官のみ可
Public  Const C_WEB0390_ACCESS_TANNIN = "2"		'//アクセス権限担任

'//成績一覧
Public  Const C_SEI0200_ACCESS_FULL   = "0"		'//アクセス権限FULLアクセス可
Public  Const C_SEI0200_ACCESS_TANNIN = "1"		'//アクセス権限担任
Public  Const C_SEI0200_ACCESS_GAKKA = "2"		'//アクセス権限学科

'処理別権限ID
'成績、欠課、遅刻一覧	'//伊藤　追加　2001/12/02
Public  Const C_ID_SEI0200 = "SEI0200"		'//FULL権限
Public  Const C_ID_SEI0210 = "SEI0210"		'//学科別
Public  Const C_ID_SEI0221 = "SEI0221"		'//1年生
Public  Const C_ID_SEI0222 = "SEI0222"		'//2年生
Public  Const C_ID_SEI0223 = "SEI0223"		'//3年生
Public  Const C_ID_SEI0224 = "SEI0224"		'//4年生
Public  Const C_ID_SEI0225 = "SEI0225"		'//5年生
Public  Const C_ID_SEI0230 = "SEI0230"		'//担任
%>
