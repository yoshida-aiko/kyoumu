<%
'/************************************************************************
' �V�X�e����    :   ���������V�X�e��
' ��  ��  ��    :   �萔��`
' ��۸���ID     :   COM_const
' �@      �\    :   �g�p�萔�̒�`
'-------------------------------------------------------------------------
' ��      ��: 2001/07/13 �J�e
' ��      �X: 2001/07/18 ���{
' ��      �X: 2001/07/19 �ɓ��@�u���Ȗ��׸ނ�ǉ�
' ��      �X: 2001/07/31 ���{�@���{��������ǉ�
'*************************************************************************/

'******************************************
'�V�X�e���֘A
'******************************************
Public const C_M00_NENDO = 9999
Public Const C_School_CD = 1            '���ꖼ
Public Const C_LEVEL_NOCHK = "XXXXXXX"  '�������������Ȃ�

'******************************************
'�\���֘A
'******************************************
Public const C_PAGE_LINE = 10      '�������X�g�̕\������


'//�f�o�b�O�p
Public const C_RetURL = "/cassist/"
'Public const C_RetURL = "/catest/"

Public const C_IMAGE_DIR = "/cassist/image/"   '
Public const C_CELL1 = "CELL2"  '���X�g�����݂ɏo�����߂̃Z���ݒ�P�B
Public const C_CELL2 = "CELL1"  '���X�g�����݂ɏo�����߂̃Z���ݒ�Q�B

Public const C_MAIN_FRAME = "fTopMain"
Public const C_LOGIN_FLG = 1	'// ���O�C����ʂ��炫�����邵

Public const C_TABLE_WIDTH = "90%"  '�e�[�u����

Public Const C_SYORYAKU_KETA = 4    '//�\�����ɏȗ����錅���i���{���j
Public Const C_ERR_RETURL = "login/default.asp"		'//Err���ɖ߂�URL

'******************************************
'�\���֘A�i���Ǝ��Ԉꗗ�j'//2001/07/18���{�ǉ�
'******************************************
Public Const C_YOUBI_MIN = 2              '�j���R�[�h�i�n�j
Public Const C_YOUBI_MAX = 6              '�j���R�[�h�i�I�j

'******************************************
'�w�Њ֘A
'******************************************
Public const C_GAKU_KETA = 9    '�w���ԍ��̌����B(�ő�P�O��)

'******************************************
'�����֘A
'******************************************
Public const C_SIKEN_KIKAN = 0		'������������
Public const C_JISSI_KIKAN = 1		'�������{����
Public const C_SEISEKI_KIKAN = 2	'���ѓ��͊���
Public const C_SIKEN_CODE_NULL = "0" '�����R�[�h�k���l
'****************************************
'���C�֘A
'****************************************
'�u���Ȗ��׸�
Public Const C_TIKAN_TUJO = 0     '�u���Ȃ�(�ʏ�)

'���C�������t���O
Public Const C_MAIN_KYOKAN_NO = 0
Public Const C_MAIN_KYOKAN_YES = 1

'���ѓ��͋����t���O
Public Const C_SEISEKI_INP_FLG_NO = 0
Public Const C_SEISEKI_INP_FLG_YES = 1

'****************************************
'�c�[���֘A
'****************************************
'�A���o�^
Public Const C_KAKU_MI = 0      '���m�F
Public Const C_KAKU_SUMI = 1    '�m�F��

'****************************************
'�w�Њ֘A�i���ڕʌ����j
'2001.12.3  �啝�ύX - ���c
'2002.05.10 �ꕔ�ǉ� - ���i
'****************************************
'��{���
Public Const C_T13_GAKUSEI_NO		=   4   '�w���ԍ�               
Public Const C_T11_SIMEI            =   6   '�����i�����j           
Public Const C_T11_SIMEI_ROMA       =   7   '�����i���[�}���j       
Public Const C_T11_SIMEI_KD         =   8   '�����i�J�i�j���_����   
Public Const C_T11_SIMEI_KYU        =   11  '�������i�����j      	
Public Const C_T11_SIMEI_ROMA_KYU   =   12  '�������i���[�}���j �@�@
Public Const C_T11_SIMEI_KD_KYU     =   13  '�������i�J�i�j���_���� 
Public Const C_T11_KAIMEI_DATE		=	141	'�ŏI��������			
Public Const C_T11_HON_ZIP          =   20  '�{�ЗX�֔ԍ�           
Public Const C_T11_HON_JUSYO        =   22  '�{�ЏZ��               
Public Const C_T11_GEN_ZIP          =   25  '���Z���X�֔ԍ�         
Public Const C_T11_GEN_JUSYO        =   26  '���Z��                 
Public Const C_T11_GEN_TEL          =   28  '���Z���d�b�ԍ�         
Public Const C_T11_KIN_TEL			=	174 '�ً}�A����

'�l���
Public Const C_T11_SEIBETU          =   16  '���ʋ敪               
Public Const C_T11_SEINENBI         =   17  '���N����               
Public Const C_T11_KETUEKI          =   18  '���t�^�敪             
Public Const C_T11_RH               =   19  '�q�g�敪               
Public Const C_T11_HOG_SIMEI        =   32  '�ی�Ҏ����i�����j     
Public Const C_T11_HOG_SIMEI_K      =   33  '�ی�Ҏ����i�J�i�j     
Public Const C_T11_HOG_ZOKU         =   34  '�ی�ґ����敪         
Public Const C_T11_HOG_ZIP          =   35  '�ی�ҏZ���X�֔ԍ�     
Public Const C_T11_HOG_JUSYO        =   37  '�ی�ҏZ��             
Public Const C_T11_HOG_TEL          =   38  '�ی�ғd�b�ԍ�         
Public Const C_T11_HOS_SIMEI        =   48  '�ۏؐl����(����)       
Public Const C_T11_HOS_SIMEI_K      =   49  '�ۏؐl����(�J�i)       
Public Const C_T11_HOS_ZOKU         =   50  '�ۏؐl�����敪         
Public Const C_T11_HOS_ZIP          =   51  '�ۏؐl�Z���X�֔ԍ�     
Public Const C_T11_HOS_JUSYO        =   53  '�ۏؐl�Z��             
Public Const C_T11_HOS_TEL          =   54  '�ۏؐl�d�b�ԍ�         
Public Const C_T11_SYUSSINKO		=	175 '�o�g�Z
Public Const C_T11_SYUSSINKOKU      =   61  '�o�g��              	
Public Const C_T11_RYUGAKU_KBN      =   62  '���w�敪             �@
Public Const C_T11_KAZOKU_1         =   68  '�Ƒ����̂P				
Public Const C_T11_KAZOKU_ZOKU_1    =   69  '�Ƒ������敪�P			
Public Const C_T11_KAZOKU_2         =   70  '�Ƒ����̂Q				
Public Const C_T11_KAZOKU_ZOKU_2    =   71  '�Ƒ������敪�Q�@�@�@�@ 
Public Const C_T11_KAZOKU_3         =   72  '�Ƒ����̂R				
Public Const C_T11_KAZOKU_ZOKU_3    =   73  '�Ƒ������敪�R			
Public Const C_T11_KAZOKU_4         =   74  '�Ƒ����̂S				
Public Const C_T11_KAZOKU_ZOKU_4    =   75  '�Ƒ������敪�S			
Public Const C_T11_KAZOKU_5         =   76  '�Ƒ����̂T				
Public Const C_T11_KAZOKU_ZOKU_5    =   77  '�Ƒ������敪�T			
Public Const C_T11_KAZOKU_6         =   78  '�Ƒ����̂U				
Public Const C_T11_KAZOKU_ZOKU_6    =   79  '�Ƒ������敪�U			
Public Const C_T11_KAZOKU_7         =   80  '�Ƒ����̂V				
Public Const C_T11_KAZOKU_ZOKU_7    =   81  '�Ƒ������敪�V			
Public Const C_T11_KAZOKU_8         =   82  '�Ƒ����̂W				
Public Const C_T11_KAZOKU_ZOKU_8    =   83  '�Ƒ������敪�W			
Public Const C_T11_HOG_SEINEIBI		=	172 '�ی�Ґ��N����
Public Const C_T11_HOS_SEINEIBI		=	173 '�ۏؐl���N����
Public Const C_T11_KAZOKU_SEINEIBI_1=	176 '�Ƒ����N�����P
Public Const C_T11_KAZOKU_SEINEIBI_2=	177 '�Ƒ����N�����Q
Public Const C_T11_KAZOKU_SEINEIBI_3=	178 '�Ƒ����N�����R
Public Const C_T11_KAZOKU_SEINEIBI_4=	179 '�Ƒ����N�����S
Public Const C_T11_KAZOKU_SEINEIBI_5=	180 '�Ƒ����N�����T
Public Const C_T11_KAZOKU_SEINEIBI_6=	181 '�Ƒ����N�����U
Public Const C_T11_KAZOKU_SEINEIBI_7=	182 '�Ƒ����N�����V
Public Const C_T11_KAZOKU_SEINEIBI_8=	183 '�Ƒ����N�����W

'���w���
Public Const C_T11_NYUNENDO         =   1   '���w�N�x    			
Public Const C_T11_TYUGAKKO_CD      =   137 '���w�Z��               
'Public Const C_T11_SYUSSINKO        =   172 '�o�g�Z                 
Public Const C_T13_TYUSOTUGYOBI     =   140 '���w�Z���Ɠ�           
Public Const C_T11_NYUGAKU_KBN      =   29  '���w�敪               
Public Const C_T11_NYU_GAKKA        =   30  '���w�w��               
Public Const C_T11_NYUGAKUBI        =   31  '���w�N����             
Public Const C_T11_JUKEN_NO         =   39  '�󌱔ԍ�               
Public Const C_T11_TYU_CLUB         =   57  '���w�Z�N���u����       
Public Const C_T11_TYU_CLUB_SYOSAI  =   58  '���w�Z�N���u�����ڍ�   
Public Const C_T11_NYU_SEISEKI      =   63  '��������               

'�w�N���
Public Const C_T13_GAKUSEKI_NO      =   85  '�w�Дԍ�               
Public Const C_T13_ZAISEKI_KBN      =   86  '�ݐЋ敪               
Public Const C_T13_GAKKA_CD         =   87  '�w�ȃR�[�h             
Public Const C_T13_COURCE_CD        =   88  '�����R�[�X          �@ 
Public Const C_T13_GAKUNEN          =   89  '�w�N                   
Public Const C_T13_CLASS            =   90  '�N���X                 
Public Const C_T13_SYUSEKI_NO1      =   91  '�o�Ȕԍ��P             
Public Const C_T13_SYUSEKI_NO2      =   92  '�o�Ȕԍ��Q             
Public Const C_T13_RYOSEI_KBN       =   93  '�����敪               
Public Const C_T13_RYUNEN_FLG       =   94  '���N�敪               
Public Const C_T13_CLUB_1           =   122 '�N���u����1            
Public Const C_T13_CLUB_1_NYUBI		=	142	'�N���u����1������		
Public Const C_T13_CLUB_2           =   123 '�N���u����2            
Public Const C_T13_CLUB_2_NYUBI		=	143 '�N���u����2������
Public Const C_T13_NENSYOKEN        =   129 '�N������ 				
Public Const C_T13_TOKUKATU         =   184 '���ʊ���
Public Const C_T13_TOKUKATU_DET     =   131 '���ʊ����ڍ�           
Public Const C_T13_SINTYO           =   132 '�g��                   
Public Const C_T13_TAIJYU           =   133 '�̏d                   
Public Const C_T13_SEKIJI_TYUKAN_Z	=	144	'�O�����ԐȎ�
Public Const C_T13_SEKIJI_KIMATU_Z 	=	145	'�O�������Ȏ�
Public Const C_T13_SEKIJI_TYUKAN_K 	=	146	'������ԐȎ�
Public Const C_T13_SEKIJI			=	120	'�w�N���Ȏ�
Public Const C_T13_NINZU_TYUKAN_Z	=	148	'�O�����ԃN���X�l��
Public Const C_T13_NINZU_KIMATU_Z  	=	149	'�O�������N���X�l��
Public Const C_T13_NINZU_TYUKAN_K  	=	150	'������ԃN���X�l��
Public Const C_T13_CLASSNINZU		=	151	'�w�N���N���X�l��
Public Const C_T13_HEIKIN_TYUKAN_Z 	=	152	'�O�����ԕ��ϓ_
Public Const C_T13_HEIKIN_KIMATU_Z 	=	153	'�O���������ϓ_
Public Const C_T13_HEIKIN_TYUKAN_K 	=	154	'������ԕ��ϓ_
Public Const C_T13_HEIKIN_KIMATU_K 	=	155	'�w�N�����ϓ_
Public Const C_T13_SUMJYUGYO		=	156	'�����Ɠ���
Public Const C_T13_SUMSYUSSEKI		=	126	'�o�ȓ���
Public Const C_T13_SUMRYUGAKU		=	128	'���w���̎��Ɠ���    
Public Const C_T13_KESSEKI_TYUKAN_Z	=	159	'�O�����Ԍ��ȓ���
Public Const C_T13_KESSEKI_KIMATU_Z	=	160	'�O���������ȓ���
Public Const C_T13_KESSEKI_TYUKAN_K	=	161	'������Ԍ��ȓ���
Public Const C_T13_SUMKESSEKI		=	125	'�w�N�����ȓ���
Public Const C_T13_KIBIKI_TYUKAN_Z	=	163	'�O�����Ԋ�������
Public Const C_T13_KIBIKI_KIMATU_Z	=	164	'�O��������������
Public Const C_T13_KIBIKI_TYUKAN_K	=	165	'������Ԋ�������
Public Const C_T13_SUMKIBTEI		=	127	'�o�Ȓ�~�����������        
Public Const C_T_CLASSIIN			=	167	'�N���X����
Public Const C_T_TANNIN				=	168	'�S�C��
Public Const C_T_JYUGYORYOMENJYO	=	169	'���Ɨ��Ə�
Public Const C_T_SYOGAKUKIN			=	170	'���w��

'���̑��\�����
Public Const C_T_JIYUSENTAKU		=	46 '���R�I������

'��������
'Public Const C_T11_SOGOSYOKEN       =   46  '��������               
Public Const C_T13_IDOU_NUM			=	171	'�ٓ���
Public Const C_T13_IDOU_KBN         =   96  '�ٓ��敪1              
Public Const C_T13_IDOU_BI          =   97  '�ٓ��N����1            
Public Const C_T13_IDOU_BIK         =   98  '�ٓ����l1              
Public Const C_T13_IDOU_KBN2        =   99  '�ٓ��敪2              
Public Const C_T13_IDOU_BI2         =   100 '�ٓ��N����2            
Public Const C_T13_IDOU_BIK2        =   101 '�ٓ����l2              
Public Const C_T13_IDOU_KBN3        =   102 '�ٓ��敪3              
Public Const C_T13_IDOU_BI3         =   103 '�ٓ��N����3            
Public Const C_T13_IDOU_BIK3        =   104 '�ٓ����l3              
Public Const C_T13_IDOU_KBN4        =   105 '�ٓ��敪4              
Public Const C_T13_IDOU_BI4         =   106 '�ٓ��N����4            
Public Const C_T13_IDOU_BIK4        =   107 '�ٓ����l4              
Public Const C_T13_IDOU_KBN5        =   108 '�ٓ��敪5              
Public Const C_T13_IDOU_BI5         =   109 '�ٓ��N����5            
Public Const C_T13_IDOU_BIK5        =   110 '�ٓ����l5              
Public Const C_T13_IDOU_KBN6        =   111 '�ٓ��敪6              
Public Const C_T13_IDOU_BI6         =   112 '�ٓ��N����6            
Public Const C_T13_IDOU_BIK6        =   113 '�ٓ����l6              
Public Const C_T13_IDOU_KBN7        =   114 '�ٓ��敪7              
Public Const C_T13_IDOU_BI7         =   115 '�ٓ��N����7            
Public Const C_T13_IDOU_BIK7        =   116 '�ٓ����l7              
Public Const C_T13_IDOU_KBN8        =   117 '�ٓ��敪8              
Public Const C_T13_IDOU_BI8         =   118 '�ٓ��N����8            
Public Const C_T13_IDOU_BIK8        =   119 '�ٓ����l8              
Public Const C_T13_IDOU_ENDBI       =   185 '�ٓ��I����             


'���g�p�敪-----------------------------------------------------------
Public Const C_T11_SINRO            =   40  '���ƌ�i�H      
Public Const C_T11_SOTUKEN_DAI      =   42  '�����_��            
Public Const C_T11_SOTU_KYOKAN_CD1  =   43  '���������P�i�����R�[�h�j
Public Const C_T11_SOTU_KYOKAN_CD2  =   44  '���������Q�i�����R�[�h�j
Public Const C_T11_SOTU_KYOKAN_CD3  =   45  '���������R�i�����R�[�h�j
Public Const C_T11_KODOSYOKEN       =   55  '�s������               
Public Const C_T11_SYUMITOKUGI      =   56  '�����Z����i�擾     
Public Const C_T11_RYO_KIBO         =   59  '������]�敪            
'Public Const C_T11_NYU_GAKUNEN		=	64	'���w�w�N
Public Const C_T11_TENNYUNEND       =   65  '�]�����w�N�x            
Public Const C_T13_NENDO            =   84  '�����N�x           
'Public Const C_T13_SUMRYUGAKU       =   128 '���w���̎��Ɠ���    
'Public Const C_T13_TOKUKATU         =   130 '���ʊ���            
Public Const C_T13_ZAISEKI_END_KBN  =   134 '�ݐЋ敪(�N�x�I���)    
Public Const C_T13_NENBIKO          =   135 '�N�����l               
Public Const C_T11_TYOSA_BIK        =   136 '���������l             
'Public Const C_T13_SEKIJI           =   120 '�w�N���Ȏ�
'Public Const C_T13_SUMSYUSSEKI      =   126 '�o�ȓ���
'Public Const C_T13_SUMRYUGAKU       =   128 '���w���̎��Ɠ���    
'Public Const C_T13_SUMKESSEKI       =   125 '�w�N�����ȓ���
'Public Const C_T13_SUMKIBTEI        =   127 '�o�Ȓ�~�����������
'------ ���i�ǉ� ------
Public Const C_T11_KOJIN_BIK        =   138 '�l���l          
Public Const C_T11_SIMEI_GAIJI      =   139 '�����O��          

'------ ���V�ǉ� '02/6/7 ------
Public Const C_M01_DAIBUNRUI150     =   150 '���ʌ���          


'---------------------------------------------------------------------

'------ �O�c�ǉ� ------
'�ψ������ރR�[�h
Public Const C_M34_SYOBUN_CD        =   0 
'�I���ȖڃO���[�v
Public Const C_T18_GRP              =   0
'�P�ʐ�
Public Const C_T15_HAITO            =   0
'�C���P�ʐ�
Public Const C_T18_SEL_GAKU         =   0   '�w�N����
Public Const C_T18_SEL_TANI         =   0   '�P�ʐ�����
'----------------------

'****************************************
'���R�I�����ڊ֘A
'****************************************
'** �萔��` **
'���ށi���j
'Private Const C_BUNRUI_KYOKA = 1
'Private Const C_BUNRUIMEI_KYOKA = "���R����"

'���R���ڎg�p�t���O
Public Const C_JIYU_USE_YES = 1    '//�g�p����
Public Const C_JIYU_USE_NO = 0     '//�g�p���Ȃ�

'���R���ڃ^�C�v
Public Const C_TYPECD_CHECK = 1  '//�`�F�b�N
Public Const C_TYPECD_ZEN = 2    '//�S�p(�����܂�)
Public Const C_TYPECD_HAN = 3    '//���p(�p����)
Public Const C_TYPECD_NUM = 4    '//���l

'****************************************
'�o�����͊֘A
'****************************************
Public Const C_JIMU_FLG_NOTJIMU = "0"   '//�����׸�(�����ȊO�œ���)
Public Const C_JIMU_FLG_JIMU = "1"      '//�����׸�(�����œ���)
Public Const C_TUKU_FLG_TUJO = "0"  '//���Ԋ��e�[�u�����ʊ����׸�(0:�ʏ����)
Public Const C_TUKU_FLG_TOKU = "1"  '//���Ԋ��e�[�u�����ʊ����׸�(1:���ʊ���(HR��))

'****************************************
'���b�Z�[�W�֘A
'****************************************
Public Const C_TOUROKU_KAKUNIN = "�o�^���Ă���낵���ł����H"				'// �o�^�m�F���b�Z�[�W
Public Const C_SAKUJYO_KAKUNIN = "�폜���Ă���낵���ł����H"				'// �폜�m�F���b�Z�[�W
Public Const C_TOUROKU_OK_MSG  = "�o�^���I�����܂���"						'// �o�^�I�����b�Z�[�W
Public Const C_SAKUJYO_OK_MSG  = "�폜���܂���" 							'// �폜�I�����b�Z�[�W
Public Const C_BRANK_VIEW_MSG  = "���ڂ�I��ŕ\���{�^���������Ă�������"   '// �󔒃y�[�W���b�Z�[�W
Public Const C_UPDATE_OK_MSG   = "�X�V���I�����܂���"						'// �X�V�������b�Z�[�W

'****************************************
'���t�֘A
'****************************************
Public  Const C_NENDO_KAISITUKI = 4             '�N�x�J�n��

'****************************************
'�A�N�Z�X�����֘A
'****************************************
'//���ʋ����֘A
Public  Const C_ACCESS_FULL   = "0"		'//�A�N�Z�X����FULL�A�N�Z�X��
Public  Const C_ACCESS_NORMAL = "1"		'//�A�N�Z�X�������
Public  Const C_ACCESS_VIEW   = "2"		'//�A�N�Z�X�����Q�Ƃ̂�

'//�g�p���ȏ��o�^�֘A
Public  Const C_WEB0320_ACCESS_FULL   = "0"		'//�A�N�Z�X����FULL�A�N�Z�X��
Public  Const C_WEB0320_ACCESS_NORMAL = "1"		'//�A�N�Z�X������勳��

'//�l���C�I���Ȗڌ���֘A
Public  Const C_WEB0340_ACCESS_FULL   = "0"		'//�A�N�Z�X����FULL�A�N�Z�X��
Public  Const C_WEB0340_ACCESS_SENMON = "1"		'//�A�N�Z�X������勳��
Public  Const C_WEB0340_ACCESS_TANNIN = "2"		'//�A�N�Z�X�����S�C

'//���x���ʉȖڌ���֘A
Public  Const C_WEB0390_ACCESS_FULL   = "0"		'//�A�N�Z�X����FULL�A�N�Z�X��
Public  Const C_WEB0390_ACCESS_SENMON = "1"		'//�A�N�Z�X������ʃ��C�������̂݉�
Public  Const C_WEB0390_ACCESS_TANNIN = "2"		'//�A�N�Z�X�����S�C

'//���шꗗ
Public  Const C_SEI0200_ACCESS_FULL   = "0"		'//�A�N�Z�X����FULL�A�N�Z�X��
Public  Const C_SEI0200_ACCESS_TANNIN = "1"		'//�A�N�Z�X�����S�C
Public  Const C_SEI0200_ACCESS_GAKKA = "2"		'//�A�N�Z�X�����w��

'�����ʌ���ID
'���сA���ہA�x���ꗗ	'//�ɓ��@�ǉ��@2001/12/02
Public  Const C_ID_SEI0200 = "SEI0200"		'//FULL����
Public  Const C_ID_SEI0210 = "SEI0210"		'//�w�ȕ�
Public  Const C_ID_SEI0221 = "SEI0221"		'//1�N��
Public  Const C_ID_SEI0222 = "SEI0222"		'//2�N��
Public  Const C_ID_SEI0223 = "SEI0223"		'//3�N��
Public  Const C_ID_SEI0224 = "SEI0224"		'//4�N��
Public  Const C_ID_SEI0225 = "SEI0225"		'//5�N��
Public  Const C_ID_SEI0230 = "SEI0230"		'//�S�C
%>
