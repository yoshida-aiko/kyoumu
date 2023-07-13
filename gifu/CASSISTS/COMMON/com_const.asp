<%
'/******************************************************************************
' システム名：キャンパスアシスト
' 処　理　名：共通定数
' ﾌﾟﾛｸﾞﾗﾑID：-
' 機　　　能：共通定数
'-------------------------------------------------------------------------------
' 作　　　成：2001.06. -　佐野　大悟
' 変　　　更：2001.07.14　佐野　大悟    異動回数の上限を追加
' 　　　　　　2001.07.16　田部　雅幸    留学区分(87)にその他私費留学を追加
' 　　　　　　2001.07.18　岩田　由佳    管理マスタ-学籍状態区分(5)の変更
' 　　　　　　2001.07.18　松本　優子    業種区分(108)を追加
' 　　　　　　2001.07.23　田部　雅幸    重み付け対象区分(89)を入試処理対象区分に変更
' 　　　　　　2001.07.23　田部　雅幸    重み付け小分類区分(90)を入試処理対象小分類区分に変更
' 　　　　　　2001.07.23　松尾  敦子    置換科目フラグ、希望フラグ、評価不可区分追加
' 　　　　　　2001.07.23　甲斐　章伍    学校状況区分、中学校区分を追加
' 　　　　　　2001.07.26　川口　恭司    部活動状況区分を追加
' 　　　　　　2001.07.26　高田　智恵美  評価対象区分を追加
' 　　　　　　2001.07.30　田部　雅幸    確約辞退フラグ(83)に未定を追加
' 　　　　　　2001.07.30　田部　雅幸    入学確定区分(120)を追加
' 　　　　　　2001.07.30　梅月　　　    管理マスタ・試験時間割管理区分(14)を追加
' 　　　　　　2001.07.30　篠原　康明    学力試験応募区分(121)を追加
' 　　　　　　2001.07.30　篠原　康明    使用フラグ(122)を追加
' 　　　　　　2001.07.23　田部　雅幸    重み付け小分類区分(90)に面接評価計を追加
' 　　　　　　2001.07.31　篠原　康明    手入力区分(123)を追加
' 　　　　　　2001.08.01　永翁　裕美    ツールバー定数を追加
' 　　　　　　2001.08.01　篠原　康明    区分マスタ以外(19.20.21)を追加
' 　　　　　　2001.08.01　竹田　俊昭    ＣＳＶ書込み定数を追加
' 　　　　　　2001.08.07　岩田　由佳    転入処理区分を追加
' 　　　　　　2001.08.09　篠原　康明    追加合格候補区分(124) 追加合格区分(125)を追加
' 　　　　　　2001.08.09　篠原　康明    過年度生フラグ(126)を追加
' 　　　　　　2001.08.09　篠原　康明    入学受付フラグ(127)を追加
' 　　　　　　2001.08.09　田部　雅幸    重み付け小分類区分(90)に小計を追加
' 　　　　　　2001.08.11　田部　雅幸    入試で使用するＣＳＶ用識別子を追加
' 　　　　　　2001.08.14　伊藤　晃      評価予定区分(128)を追加
' 　　　　　　2001.08.16　田部　雅幸    端数処理区分(129)を追加
' 　　　　　　2001.08.16　田部　雅幸    重み付け小分類区分(90)を変更
' 　　　　　　2001.08.16　田部　雅幸    管理マスタ・基本中学校成績段階数(24)を追加
' 　　　　　　2001.08.16　田部　雅幸    管理マスタ・推薦志望可能学科数(25)を追加
' 　　　　　　2001.08.16　田部　雅幸    管理マスタ・学力志望可能学科数(26)を追加
' 　　　　　　2001.08.16　田部　雅幸    管理マスタ・編入志望可能学科数(27)を追加
' 　　　　　　2001.08.16　篠原　康明    中学校成績段階数区分(130)を追加
' 　　　　　　2001,08.16　篠原　康明    中学校成績入力段階数フラグ(131)を追加
' 　　　　　　2001.08.18　松尾　敦子    平均点科目区分(132)追加
' 　　　　　　　　　　　　　　　　　    管理マスタ・履修状態区分(28)
' 　　　　　　　　　　　　　　　　　    管理マスタ・前期授業時間割状態区分(29)
' 　　　　　　　　　　　　　　　　　    管理マスタ・後期授業時間割状態区分(30)
' 　　　　　　　　　　　　　　　　　    管理マスタ・欠課・欠席設定条件区分(31)
' 　　　　　　　　　　　　　　　　　    管理マスタ・欠課累積情報区分追加(32)
' 　　　　　　2001.08.24　多賀　成文    ComboBox設定用　試験マスタ　試験コード_試験名称を追加
' 　　　　　　　　　　　　　　　　　    ComboBox設定用　試験マスタ　試験科目コード_試験科目名称を追加
' 　　　　　　2001.08.25　多賀　成文    ComboBox設定用　教官マスタ　教官コード_姓&名を追加
' 　　　　　　2001.08.25　田部　雅幸    入試区分(135)を追加
' 　　　　　　　　　　　　　　　　　    学力移行済みフラグ(136)を追加
' 　　　　　　　　　　　　　　　　　    入試出席フラグ(137)を追加
' 　　　　　　　　　　　　　　　　　    入試取り消しフラグ(138)を追加
' 　　　　　　2001.08.29　梅月　亮     役職を追加
' 　　　　　　2001.08.30　佐野　大悟   判定使用フラグを追加
'            2001.08.31　松尾　敦子　　特別活動評価区分を追加
' 　　　　　　2001.08.25　田部　雅幸    入試健康診断区分(140)を追加
'            2001.09.03  佐野　大悟   欠課・欠席設定(M15_KEKKA_SETTEI)のコードを追加
'            2001.09.04  佐野　大悟   前期終了日・年度開始終了日の廃止
' 　　　　　　2001.09.05　田部　雅幸    入試卒業フラグ(144)を追加
' 　　　　　　2001.09.06　田部　雅幸    入学受付フラグ(127)に入学辞退を追加
' 　　　　　　2001.09.08　田部　雅幸    入試処理対象小分類区分(90)の今まであった項目を廃止
'                                   　入試処理対象小分類区分(90)に推薦区分１・２を追加
'                                   　入試処理対象小分類区分(90)に５段階に対する処理を追加
'                                   　入試処理対象小分類区分(90)に10段階に対する処理を追加
'                                   　推薦入試処理対象区分(145)を追加
'                                   　学力入試処理対象区分(146)を追加
'                                   　中学校成績処理対象区分(148)を追加
'                                   　入試処理対象区分(89)にその他を追加
' 　　　　　　2001.09.12　田部　雅幸    推薦入試処理対象区分(145)に健康診断を追加
'                                   　学力入試処理対象区分(146)に健康診断を追加
'            2001.09.18　山本　富志江　評価形式マスタ　略称　可・不可を追加
' 　　　　　　2001.09.20　田部　雅幸    入試評価形式区分(149)を追加
' 　　　　　　2001.09.20　田部　雅幸    九州の高専の固有番号を追加
' 　　　　　　2001.09.22　田部　雅幸    受験地区分(142)を追加
' 　　　　　　2001.09.25　岩田　由佳    判定状態区分(38)を追加
'            2001.09.27　佐野　大悟    換算フラグ(する/しない)を追加
' 　　　　　　2001.10.08　田部　雅幸    編入学処理対象区分(147)を追加
' 　　　　　　2001.10.16　岩田  由佳    入学回生区分(39)を追加
' 　　　　　　2001.10.19　岩田  由佳    卒業証書発行番号(最大値）(40)を追加
' 　　　　　　2001.10.25　佐野　大悟    カウント区分(105)に追加
' 　　　　　　2001.10.30　上村　久美子  欠席区分(19)に一欠課(8)を追加
'******************************************************************************/

Public Const C_IDO_MAX = 8              '異動回数の上限

'***********************
'コード定義(大分類コード)
'***********************
Public Const C_SEIBETU = 1              '性別
Public Const C_GENGOU = 2               '元号
Public Const C_NYUGAKU = 3              '入学区分
Public Const C_ZAISEKI = 5              '在籍区分
Public Const C_TUGAKU = 8               '通学区分
Public Const C_IDO = 9                  '在籍異動区分
Public Const C_ZOKUGARA = 10            '続柄
Public Const C_SOTUGYO = 11             '卒業区分
Public Const C_NAISIN = 12              '内申区分
Public Const C_SIKEN = 13               '試験区分
Public Const C_GAKKI = 14               '学期区分
Public Const C_HISSEN = 15              '必修・選択区分
Public Const C_JUGYO_KEITAI = 16        '授業形態区分
Public Const C_HIJOKIN = 17             '常勤・非常勤区分
Public Const C_KYOKAN = 18              '教官区分
Public Const C_KESSEKI = 19             '欠席区分
Public Const C_KES_TODOKEDE = 20        '欠席届出区分
'Public Const C_YOUBI = 21               '曜日コード        'VB定数を使用してください。
Public Const C_SOFU = 22                '送付区分
Public Const C_NYURYO = 23              '入寮区分
Public Const C_RYOKIBO = 24             '入寮希望区分
Public Const C_MENJO = 25               '免除区分
Public Const C_KYOKA_KEIRETU = 26       '教科系列区分
Public Const C_RISYU = 27               '履修学科区分 - Add 2001.07.11 岡田
Public Const C_TANI = 28                '単位取得区分
Public Const C_GYOJI = 33               '行事区分
Public Const C_NISSU = 34               '出席日数区分
Public Const C_KITUEN = 35              '喫煙許可区分
Public Const C_KURUMA = 36              '車通学許可区分
Public Const C_NYU_KAKU = 37            '入学確約区分
Public Const C_JURYO = 38               '受領区分
Public Const C_JIKANWARI = 39           '時間割編成区分
Public Const C_CLASS_TYPE = 40          'クラス分け編成区分
Public Const C_SINKYU = 41              '進級制度区分
Public Const C_SUISEN = 42              '推薦区分
Public Const C_KENGEN = 45              '権限コード
Public Const C_KAISETUKI = 51           '開設期間コード
Public Const C_KAISETU = 52             '開設区分
Public Const C_KENNAIGAI = 55           '県内外区分
Public Const C_TEISYUTU = 61            '提出済み区分
Public Const C_BU_SYOZOKU = 62          '部活動所属区分
Public Const C_RONIN = 63               '現役浪人区分
Public Const C_NYUTAI = 64              '入学対象区分
Public Const C_BU = 65                  '部区分
Public Const C_TUJO = 66                '通常科目区分
Public Const C_OMOMI = 67               '重み付け方式区分
Public Const C_ENZAN = 68               '演算区分
Public Const C_TORIKOMI = 69            '取り込みデータ区分
Public Const C_KEKKA_HANTEI = 70        '欠課判定方式区分
Public Const C_USER = 71                'ユーザー区分
Public Const C_BLOOD = 72               '血液型
Public Const C_RH = 73                  '血液型RH区分
Public Const C_KAMOKU = 74              '科目区分
Public Const C_SINRO = 75               '進路先区分
Public Const C_SINGAKU = 76             '進学区分
'Public Const C_RISYU = 77               '履修区分　-　2001.07.11 岡田 (選択種別がある為不要) DEL
Public Const C_NYU_JOKYO = 78           '入力状況選択
Public Const C_KENSIN = 79              '検診診断区分
Public Const C_GENGO = 80               '言語区分
Public Const C_RYUGAKU = 81             '留学区分
Public Const C_NYUSI_INSATU = 82        '入試関連帳票印刷フラグ
Public Const C_KAKUYAKU_JITAIFLG = 83      '確約辞退フラグ
Public Const C_NYUSI_GOHI = 84          '入学試験合否区分
Public Const C_SISEN_NYUGAKU = 85       '推薦入学不合格フラグ
Public Const C_HOSYONIN = 86            '保証人保護者同一フラグ
Public Const C_RYUNEN = 87              '留年区分
Public Const C_NYURYOKU = 88            '入力済フラグ区分
'Public Const C_OMOMI_TAISYO = 89        '重み付け対象区分
Public Const C_NYUSI_TAISYO = 89        '入試処理対象区分
'Public Const C_OMOMI_SYOBUN = 90        '重み付け小分類区分
Public Const C_NYUSI_SYOBUN = 90        '入試処理対象小分類区分
Public Const C_IIN = 91                 '委員区分
Public Const C_NYUTAIRYO = 92           '入退寮区分
Public Const C_HIROSA_TANI = 93         '広さの単位区分
Public Const C_HEYA_JOKYO = 94          '部屋状況区分
Public Const C_NYURYO_RIYU = 95         '入寮理由区分
Public Const C_SYUSSINKO = 96           '出身校区分
Public Const C_HANTEI_TAISYO = 97       '判定対象区分
Public Const C_SEN_SYUBETU = 98         '選択科目種別区分
Public Const C_LEVEL_BETU = 99          'レベル別科目区分
Public Const C_SENTAKU_KAHI = 100       '選択可否区分
Public Const C_JUGYO_KBN = 103          '授業区分
Public Const C_SIKEN_KBN = 104          '試験実施区分
Public Const C_COUNT_KBN = 105          'カウント区分
Public Const C_SYUTAI_KBN = 106         '修了退学区分
Public Const C_TAIGKU_KBN = 107         '退学区分
Public Const C_GYOSYU_KBN = 108         '業種区分
Public Const C_SYORI_KBN = 109          '処理区分
Public Const C_LEVEL_KBN = 110          'レベル区分
Public Const C_TIKAN_KAMOKU = 112       '置換科目フラグ
Public Const C_KIBOU_FLG = 113          '希望フラグ
Public Const C_HYOKA_FUKA = 114         '評価不可区分
Public Const C_GAKKO_JYOKYO = 115       '学校状況区分
Public Const C_TYUGAKKO_KBN = 116       '中学校区分
Public Const C_KYUGAKU_KBN = 117        '休学区分
Public Const C_BUKATUDO_JYOKYO = 118    '中学校区分
Public Const C_HYOKA_TAISYO = 119       '評価対象区分
Public Const C_NYUSI_KAKUTEI = 120      '入試確定区分
Public Const C_NYUSI_GAKUOUBO_KBN = 121 '学力入試応募希望区分
Public Const C_NYUSI_SIYOU_KBN = 122    '使用フラグ
Public Const C_TENYURYOKU_KBN = 123     '手入力区分
Public Const C_TUIKA_KOHO_KBN = 124     '追加合格候補区分
Public Const C_TUIKA_GOKAKU_KBN = 125   '追加合格区分
Public Const C_KANENDO_KBN = 126        '過年度生フラグ
Public Const C_UKETUKE_KBN = 127        '入学受付フラグ
Public Const C_HYOKAYOTEI_KBN = 128     '評価予定区分
Public Const C_HASU_SYORI_KBN = 129     '端数処理区分
Public Const C_TYUSEI_DAN_KBN = 130     '中学校成績段階数区分
Public Const C_TYUSEI_NYU_KBN = 131     '中学校成績入力段階数フラグ
Public Const C_HEIKIN_KAMOKU_KBN = 132  '平均点科目区分
Public Const C_KEKKA_JYOHOU_KBN = 133   '欠課欠席情報区分
Public Const C_SUUCHI_KIJYUN_KBN = 134  '数値基準区分
Public Const C_NYUSI_KBN = 135          '入試区分
Public Const C_NYUSI_IKO_FLG = 136      '学力移行済みフラグ
Public Const C_NYUSI_SUSSEKI_FLG = 137  '入試出席フラグ
Public Const C_NYUSI_TORIKESI_FLG = 138 '入試取り消しフラグ
Public Const C_HANTEI_FLG = 139         '判定使用フラグ
Public Const C_NYUSI_KENKO_SINDAN_KBN = 140     '入試健康診断区分
Public Const C_HYOKATOKU_KBN = 141      '特別活動評価区分
Public Const C_JUKENTI_KBN = 142        '受験地区分
Public Const C_SYOMEISYO_KBN = 143      '証明書区分
Public Const C_NYUSI_SOTUGYO_FLG = 144      '入試卒業フラグ
Public Const C_NYUSI_SUISEN_SYORI_KBN = 145     '推薦入試処理対象区分
Public Const C_NYUSI_GAKURYOKU_SYORI_KBN = 146  '学力入試処理対象区分
Public Const C_NYUSI_HENNYU_SYORI_KBN = 147     '編入学処理対象区分
Public Const C_NYUSI_TYUGAKKO_SYORI_KBN = 148   '中学校成績処理対象区分
Public Const C_NYUSI_HYOKA_KBN = 149            '入試評価形式区分


'***********************
'コード定義(小分類コード)
'***********************

'性別(C_SEIBETU=1)
Public Const C_SEIBETU_M = 1            '男性
Public Const C_SEIBETU_F = 2            '女性

'元号(C_GENGOU=2)
Public Const C_GEN_MEIJI = 1           '明治
Public Const C_GEN_TAISYOU = 2         '大正
Public Const C_GEN_SYOWA = 3           '昭和
Public Const C_GEN_HEISEI = 4          '平成

'入学区分(C_NYUGAKU=3)
Public Const C_NYU_SUISEN_G = 1         '推薦(学力)
Public Const C_NYU_SUISEN_C = 2         '推薦(クラブ)
Public Const C_NYU_GAKU = 3             '学力選抜
Public Const C_NYU_HENNYU = 4           '編入学
Public Const C_NYU_TENNYU = 5           '転入学
Public Const C_NYU_RYUGAKU = 6          '留学

'在籍区分(C_ZAISEKI=5)-------在籍区分（年度終わり）と併用してください
Public Const C_ZAI_ZAIGAKU = 0            '在学中
Public Const C_ZAI_KYUGAKU = 1            '休学中
Public Const C_ZAI_TEIGAKU = 2            '停学
Public Const C_ZAI_SOTUGYO = 3            '卒業
Public Const C_ZAI_TAIGAKU = 4            '退学
Public Const C_ZAI_SYU_TAIGAKU = 5        '終了退学
Public Const C_ZAI_TENSYUTU = 6           '転出
Public Const C_ZAI_JOSEKI = 7             '除籍
'Public Const C_ZAI_RYUGAKU = 8            '留学中　　　休学中で代用してください　2001.07.11　DEL

'通学区分(C_TUGAKU=8)
Public Const C_TUG_JITAKU = 0           '自宅通学
Public Const C_TUG_RYOSEI = 1           '寮生
Public Const C_TUG_SONOTA = 2           'その他

'在籍異動区分(C_IDO=9)
Public Const C_IDO_KYU_BYOKI = 1    '休学(病気・怪我)
Public Const C_IDO_KYU_HOKA = 2     '休学(経済的他)
Public Const C_IDO_FUKUGAKU = 3     '復学
Public Const C_IDO_TEIGAKU = 4      '停学
Public Const C_IDO_TEI_KAIJO = 5    '停学解除
Public Const C_IDO_TAI_2NEN = 6     '退学(二年連続)
Public Const C_IDO_TAI_HOKA = 7     '退学(その他)
Public Const C_IDO_TAI_SYURYO = 8   '修了退学
Public Const C_IDO_JO_MINO = 9      '除籍(授業料未納)
Public Const C_IDO_JO_SIBO = 10      '除籍(死亡・行方不明)

Public Const C_IDO_TENKO = 11       '転出
Public Const C_IDO_TENKA = 12       '転科
'Public Const C_IDO_TORIKESI = 13   '入学取消     　    除籍を利用してください　2001.07.09　DEL
Public Const C_IDO_KOKUHI = 14      '国費への移行
'Public Const C_IDO_RYUGAKU = 15     '留学            　休学で代用してください  2001.07.11　DEL

'続柄(C_ZOKUGARA=10)
Public Const C_ZOKU_TITI = 1    '父
Public Const C_ZOKU_HAHA = 2    '母
Public Const C_ZOKU_OJI = 3     '叔父・叔母
Public Const C_ZOKU_SOFUBO = 4  '祖父母
Public Const C_ZOKU_KYODAI = 5  '兄弟
Public Const C_ZOKU_TIJIN = 6   '知人
Public Const C_ZOKU_SONOTA = 9  'その他

'卒業区分(C_SOTUGYO=11)
'Public Const C_SOTU_SOTU = 1        '卒業
'Public Const C_SOTU_JOSEKI = 2      '除籍
'Public Const C_SOTU_TAIGAKU = 3     '退学
'Public Const C_SOTU_SYURYO = 4      '修了退学
'Public Const C_SOTU_TENKO = 5       '転校
Public Const C_SOTU_SOTU = 3        '卒業
Public Const C_SOTU_TAIGAKU = 4     '退学
Public Const C_SOTU_SYURYO = 5      '修了退学
Public Const C_SOTU_TENKO = 6       '転校
Public Const C_SOTU_JOSEKI = 7      '除籍

'内申区分(C_NAISIN=12)
Public Const C_NAISIN_5TEN = 1      '５点法
Public Const C_NAISIN_10TEN = 2     '10点法

'試験区分(C_SIKEN=13)
Public Const C_SIKEN_ZEN_TYU = 1    '前期中間試験
Public Const C_SIKEN_ZEN_KIM = 2    '前期期末試験
Public Const C_SIKEN_KOU_TYU = 3    '後期中間試験
Public Const C_SIKEN_KOU_KIM = 4    '後期期末試験
Public Const C_SIKEN_JITURYOKU = 5  '実力試験
Public Const C_SIKEN_TUISI = 6      '追試験

'学期区分(C_GAKKI=14)
Public Const C_GAKKI_ZENKI = 1  '前期
Public Const C_GAKKI_KOUKI = 2  '後期

'必修・選択区分(C_HISSEN=15)
Public Const C_HISSEN_HIS = 1   '必修
Public Const C_HISSEN_SEN = 2   '選択

'授業形態区分(C_JUGYO_KEITAI=16)
Public Const C_JUG_KOUGI = 1        '講義
Public Const C_JUG_ENSYU = 2        '演習
Public Const C_JUG_JIKKEN = 3       '実験・演習
Public Const C_JUG_JITUGI = 4       '実技
Public Const C_JUG_KAGAI = 5        '課外授業
Public Const C_JUG_SONOTA = 6       'その他

'常勤・非常勤区分(C_HIJOKIN=17)
Public Const C_HIJOKIN_JOKIN = 1    '常勤
Public Const C_HIJOKIN_HIJOKIN = 2  '非常勤

'教官区分(C_KYOKAN=18)
Public Const C_KYOKAN_KOTYO = 1     '校長
Public Const C_KYOKAN_KYOJU = 2     '教授
Public Const C_KYOKAN_JOKYOJU = 3   '助教授
Public Const C_KYOKAN_KOSI = 4      '講師
Public Const C_KYOKAN_JOSYU = 5     '助手
Public Const C_KYOKAN_SONOTA = 9    'その他
'非常勤削除 2001.08.08 matsuo
'非常勤は常勤・非常勤区分を見てください


'欠席区分(C_KESSEKI=19)
Public Const C_KETU_SYUSSEKI = 0    '出席
Public Const C_KETU_KEKKA = 1       '欠課(時限マスタの単位数を使用)
Public Const C_KETU_TIKOKU = 2      '遅刻
Public Const C_KETU_SOTAI = 3       '早退
Public Const C_KETU_KIBIKI = 4      '忌引
Public Const C_KETU_TEIGAKU = 5     '停学
Public Const C_KETU_TOKUBETU = 6    '特別欠席
Public Const C_KETU_KOKETU = 7      '公欠
Public Const C_KETU_KEKKA_1 = 8     '１欠課(時限マスタに関係なく、１欠課)

'欠席届出区分(C_KES_TODOKEDE=20)
Public Const C_KES_NASI = 0     '届出なし
Public Const C_KES_BYOKI = 1    '病気欠席
Public Const C_KES_SONOTA = 2   'その他の理由
Public Const C_KES_KIBIKI = 3   '忌引等
Public Const C_KES_TEISI = 4    '出席停止

'曜日コード(C_YOBIKODO=21)              'VB定数を使ってください。
'Public Const C_YOBI_NITI = 0    '日
'Public Const C_YOBI_GETU = 1    '月
'Public Const C_YOBI_KA = 2      '火
'Public Const C_YOBI_SUI = 3     '水
'Public Const C_YOBI_MOKU = 4    '木
'Public Const C_YOBI_KIN = 5     '金
'Public Const C_YOBI_DO = 6      '土

'送付区分(C_SOFU=22)
Public Const C_SOFU_MI = 0      '未発行
Public Const C_SOFU_SUMI = 1    '発行済

'入寮区分(C_NYURYO=23)
Public Const C_NYURYO_NO = 0      '寮外生
Public Const C_NYURYO_YES = 1     '入寮中

'入寮希望区分(C_RYOKIBO=24)
Public Const C_RYOKIBO_NO = 0     '希望しない
Public Const C_RYOKIBO_YES = 1    '希望する

'免除区分(C_MENJO = 25)
Public Const C_MENJ_JUGYORYOMENJO = 1           '授業料免除（経済的理由）
Public Const C_MENJ_KYUGAKU = 2                 '免除（休学）
Public Const C_MENJ_SIBO = 3                    '免除（死亡・行方不明）
Public Const C_MENJ_TOKU = 4                    '免除（特別な理由）
Public Const C_MENJ_MINOTAIGAKU = 5             '免除（未納による退学）
Public Const C_MENJ_TYOSYUYUYOKYOKATAIGAKU = 6  '免除（徴収猶予許可者の退学）
Public Const C_MENJ_TORIKESI = 7                '免除取消
Public Const C_MENJ_JIKOKANSEI = 8              '時効完成
Public Const C_MENJ_SAIKENSIBO = 9              '債権金額の減（死亡・行方不明）
Public Const C_MENJ_SAIKENTAIGAKU = 10          '債権金額の減（退学）
Public Const C_MENJ_SAIKENJOSEKI = 11           '債権金額の減（除籍）
Public Const C_MENJ_SAIKENKIKANHEN = 12         '債権金額の減（期間変更）
Public Const C_MENJ_SAIKENTENSYUTU = 13         '債権金額の減（転出）
Public Const C_MENJ_SAIKENTYOKOKAMOKU = 14      '債権金額の減（聴講科目の減少）
Public Const C_MENJ_SAIKENNYUTORIKESI = 15      '債権金額の減（入学取消）
Public Const C_MENJ_SAIKENKOKUHI = 16           '債権金額の減（国費への以降）
Public Const C_MENJ_SAIKENTEISEI = 17           '債権金額の減（訂正）
Public Const C_MENJ_SAIKENKIKANHENSAI = 18         '債権金額の増（期間変更）
Public Const C_MENJ_SAIKENTYOKOTUIKA = 19       '債権金額の増（聴講科目の追加）
Public Const C_MENJ_SAIKENTESEI = 20           '債権金額の増（訂正）
Public Const C_MENJ_SAIKENKANRI = 21            '転学部（債権管理簿整理用）
Public Const C_MENJ_SONOTA = 22                 'その他

'教科系列区分(C_KYOKA_KEIRETU=26)
Public Const C_KEI_NASI = 0     'なし
Public Const C_KEI_KOKUGO = 1   '国語系
Public Const C_KEI_SYAKAI = 2   '社会系
Public Const C_KEI_SUGAKU = 3   '数学系
Public Const C_KEI_RIKA = 4     '理科系
Public Const C_KEI_TAIIKU = 5   '体育系
Public Const C_KEI_GEIJUTU = 6  '芸術系
Public Const C_KEI_GAIKOKU = 7  '外国語系
Public Const C_KEI_SONOTA = 9   'その他

'履修学科区分(C_RISYU=27)
Public Const C_RIS_KYOTU = 0    '共通

'単位取得区分(C_TANI=28)
Public Const C_TANI_SYUTOKU = 0     '取得
Public Const C_TANI_MI = 1          '未取得
Public Const C_TANI_MISENTAKU = 2   '未選択

'行事区分(C_GYOJI=33)
Public Const C_GYO_JUGYO = 0        '授業
Public Const C_GYO_GYOJI = 1        '行事

'出席日数区分(C_NISSU=34)
Public Const C_NIS_NASI = 0         'なし
Public Const C_NIS_1KAI = 1         '１年間皆勤賞
Public Const C_NIS_3KAI = 2         '３年間皆勤賞
Public Const C_NIS_5KAI = 3         '５年間皆勤賞
Public Const C_NIS_GAKUYU = 4       '学生成績優秀賞
Public Const C_NIS_TANYU = 5        '単年度成績優秀賞
Public Const C_NIS_3YU = 6          '３年間成績優秀賞
Public Const C_NIS_5YU = 7          '５年間成績優秀賞
Public Const C_NIS_KOJO = 8         '成績向上賞

'喫煙許可区分(C_KITUEN=35)
Public Const C_KITUEN_FUKYOKA = 0   '不許可
Public Const C_KITUEN_KYOKA = 1     '許可

'車通学許可区分(C_KURUMA=36)
Public Const C_CAR_FUKYOKA = 0      '不許可
Public Const C_CAR_KYOKA = 1        '許可

'入学確約区分(C_NYU_KAKU=37)
Public Const C_NYU_KAKUYAKU = 0   '入学
Public Const C_NYU_JITAI = 1      '辞退

'受領区分(C_JURYO=38)
Public Const C_JURYO_NO = 0       '未受領
Public Const C_JURYO_YES = 1      '受領

'時間割編成区分(C_JIKANWARI=39)
Public Const C_JIK_TUJO = 0     '通常
Public Const C_JIK_HEN = 1      '変則式

'クラス分け編成区分(C_CLASS_TYPE=40)
'Public Const C_CLASS_GAKKA = 1    '学科別クラス    UPDATE 2001/07/12
'Public Const C_CLASS_KONGO = 2    '混合クラス      UPDATE 2001/07/12
Public Const C_CLASS_GAKKA = 0    '学科別クラス
Public Const C_CLASS_KONGO = 1    '混合クラス

'進級制度区分(C_SINKYU=41)
Public Const C_SIN_GAKU = 1     '学年制
Public Const C_SIN_TANI = 2     '単位制
Public Const C_SIN_HOKA = 9     'その他

'推薦区分(C_SUISEN=42)
Public Const C_SUI_GAKU = 1     '学力
Public Const C_SUI_SPORTS = 2   'スポーツ

'権限コード(C_KENGEN=45)
Public Const C_KENGEN_GAKU = 0     '学生
Public Const C_KENGEN_KYOKAN = 1   '教官
Public Const C_KENGEN_RYOMU = 2    '寮務
Public Const C_KENGEN_JIMU = 3     '事務
Public Const C_KENGEN_KANRI = 4    '管理者

'開設期間コード(C_KAISETUKI=51)
Public Const C_KAI_TUNEN = 0    '通年
Public Const C_KAI_ZENKI = 1    '前期
Public Const C_KAI_KOUKI = 2    '後期
Public Const C_KAI_NASI = 3     '開設しない

'開設区分(C_KAISETU=52)
Public Const C_KAISETU_NO = 0    '未開設
Public Const C_KAISETU_YES = 1   '開設

'県内外区分(C_KENNAIGAI=55)
Public Const C_KEN_NAI = 0      '県内
Public Const C_KEN_GAI = 1      '県外

'提出済み区分(C_TEISYUTU=61)
Public Const C_TEISYUTU_MI = 0       '未提出
Public Const C_TEISYUTU_SUMI = 1     '提出

'部活動所属区分(C_BU_SYOZOKU=62)
Public Const C_BU_SYOZOKU_NO = 0        '未所属
Public Const C_BU_SYOZOKU_YES = 1       '所属

'現役浪人区分(C_RONIN=63)
Public Const C_RONIN_GEN = 0      '現役
Public Const C_RONIN_RON = 1      '浪人

'入学対象区分(C_NYUTAI=64)
Public Const C_NYU_HON = 0      '本科
Public Const C_NYU_SENKO = 1    '専攻科
Public Const C_NYU_HOKA = 2     '他大学

'部区分(C_BU=65)
Public Const C_BU_BU = 0      '部
Public Const C_BU_HOKA = 1    'その他

'通常科目区分(C_TUJO=66)
Public Const C_TUJO_TU = 0      '通常科目
Public Const C_TUJO_GAI = 1     '授業外科目
Public Const C_TUJO_TOKU = 2    '特別活動

'重み付け方式区分(C_OMOMI=67)
Public Const C_OMO_NASI = 0                 '調整なし
Public Const C_OMO_AVG = 1                  '×１００÷平均点
Public Const C_OMO_Multiply_Divided = 2     '×Ａ÷Ｂ
Public Const C_OMO_Multiply_Only = 3        '×Ａ

'演算区分(C_ENZAN=68)
Public Const C_EN_KA = 1        '加算
Public Const C_EN_GEN = 2       '減算
Public Const C_EN_JOU = 3       '乗算
Public Const C_EN_JO = 4        '除算

'取り込みデータ区分(C_TORIKOMI=69)
Public Const C_TORI_KUBUN = 1       '区分データ
Public Const C_TORI_SI = 2          '市町村データ
Public Const C_TORI_KEN = 3         '県データ
Public Const C_TORI_TYU = 4         '中学校データ
Public Const C_TORI_KA = 5          '科目データ
Public Const C_TORI_KYOKAN = 6      '教官データ
Public Const C_TORI_GENGO = 7       '元号データ
Public Const C_TORI_KYOSITU = 8     '教室データ
Public Const C_TORI_KYOKANSITU = 9  '教官室データ
Public Const C_TORI_USER = 10       'ユーザーデータ
Public Const C_TORI_BU = 11         '部活動データ

'欠課判定方式区分(C_KEKKA_HANTEI=70)
Public Const C_KEKKA_HANTEI_KOTEI = 0        '固定
Public Const C_KEKKA_HANTEI_KNTN = 1         '休学無・単位無
Public Const C_KEKKA_HANTEI_KNTA = 2         '休学無・単位有
Public Const C_KEKKA_HANTEI_KATN = 3         '休学有・単位無
Public Const C_KEKKA_HANTEI_KATA = 4         '休学有・単位有

'ユーザー区分(C_USER=71)
Public Const C_USER_GAKU = 0    '学生
Public Const C_USER_KYOKAN = 1  '教官
Public Const C_USER_RYOMU = 2   '寮務       '教務の間違いでは？
Public Const C_USER_JIMU = 3    '事務
Public Const C_USER_ADMIN = 4   '管理者

'血液型(C_BLOOD=72)
Public Const C_BLOOD_A = 1   'A型
Public Const C_BLOOD_B = 2   'B型
Public Const C_BLOOD_O = 3   'O型
Public Const C_BLOOD_AB = 4  'AB型
Public Const C_BLOOD_X = 5   'その他

'血液型RH区分(C_RH=73)
Public Const C_RH_PLUS = 0    '+
Public Const C_RH_MINUS = 1   '-

'科目区分(C_KAMOKU=74)
Public Const C_KAMOKU_IPPAN = 0     '一般科目
Public Const C_KAMOKU_SENMON = 1    '専門科目
Public Const C_KAMOKU_TOKUBETU = 2  '特別科目

'進路先区分(C_SINRO=75)
Public Const C_SINRO_SINGAKU = 1      '進学
Public Const C_SINRO_SYUSYOKU = 2     '就職
Public Const C_SINRO_SONOTA = 3       'その他

'進学区分(C_SINGAKU=76)
Public Const C_SIN_KOKURITU = 1     '国立大学
Public Const C_SIN_SIRITU = 2       '私立大学
Public Const C_SIN_SENKO = 3        '高専専攻科
Public Const C_SIN_SENMON = 4       '専門学校

'履修区分(C_RISYU = 77) -　2001.07.11 岡田 (選択種別がある為不要)

'入力状況選択(C_NYU_JOKYO=78)
Public Const C_NYUJOKYO_NYU_MI = 1      '未入力
Public Const C_NYUJOKYO_IPP_MI = 2      '一般未入力
Public Const C_NYUJOKYO_HA_MI = 3       '歯未入力
Public Const C_NYUJOKYO_IPP_SUMI = 4    '一般済
Public Const C_NYUJOKYO_HA_SUMI = 5     '歯済
Public Const C_NYUJOKYO_NYU_SUMI = 6    '入力済

'検診診断区分(C_KENSIN=79)
Public Const C_KEN_NASI = 0     '異常なし
Public Const C_KEN_KANSATU = 1  '定期的観察
Public Const C_KEN_SINDAN = 2   '専門医による診断要

'言語区分(C_GENGO=80)
Public Const C_GENGO_JPN = 1     '日本語
Public Const C_GENGO_ENG = 1     '英語

'留学区分(C_RYUGAKU=81)
Public Const C_RYUGAKU_KOKUHI = 1       '国費留学
Public Const C_RYUGAKU_SIHI = 2         '私費留学
Public Const C_RYUGAKU_SONOTA = 3       'その他私費留学 2001/07/16 Add 田部

'入試関連帳票印刷フラグ(C_NYUSI_INSATU=82)
Public Const C_NYUSI_INSATU_MI = 0       '印刷していない
Public Const C_NYUSI_INSATU_SUMI = 1     '印刷済み

'確約辞退フラグ(C_KAKUYAKU_JITAIFLG=83)
Public Const C_KAKUYAKU_MITEI = 0       '未定 2001/07/30 Add
Public Const C_KAKUYAKU_KAKUYAKU = 1    '確約
Public Const C_KAKUYAKU_JITAI = 2       '辞退

'入学試験合否区分(C_NYUSI_GOHI=84)
Public Const C_GOHI_GOKAKU = 1     '合格
Public Const C_GOHI_FUGOKAKU = 2   '不合格

'推薦入学不合格フラグ(C_SISEN_NYUGAKU=85)
Public Const C_SUISEN_GOKAKU = 0        '推薦入学を受けていない
Public Const C_SUISEN_FUGOKAKU = 1      '推薦入学に不合格

'保証人保護者同一フラグ     (C_HOSYONIN=86)
Public Const C_HOSYONIN_DOUITU = 0          '同一人物
Public Const C_HOSYONIN_BETUJIN = 1         '別人

'留年区分(C_RYUNEN = 87)
Public Const C_RYUNEN_ON = 1          '留年した
Public Const C_RYUNEN_OFF = 0         '留年なし

'入力済フラグ区分(C_NYURYOKU=88)
Public Const C_NYURYOKU_MI = 0       '未入力
Public Const C_NYURYOKU_SUMI = 1     '入力済

'重み付け対象区分(C_OMOMI_TAISYO=89)変更
'Public Const C_OMOMI_TAISYO_SUISEN = 1       '推薦入学
'Public Const C_OMOMI_TAISYO_GAKURYOKU = 2    '学力試験
'Public Const C_OMOMI_TAISYO_HENNYU = 3       '編入試験
'入試処理対象区分(C_NYUSI_TAISYO=89)
Public Const C_NYUSI_TAISYO_TYUGAKU = 0         '中学校成績
Public Const C_NYUSI_TAISYO_SUISEN = 1          '推薦入学
Public Const C_NYUSI_TAISYO_GAKURYOKU = 2       '学力試験
Public Const C_NYUSI_TAISYO_HENNYU = 3          '編入試験
Public Const C_NYUSI_TAISYO_HENNYU_SUISEN = 4   '編入学推薦
Public Const C_NYUSI_TAISYO_SONOTA = 9          'その他

'重み付け小分類区分       (C_OMOMI_SYOBUN=90)
'/******* 確定したら追加してください ************/
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU1  　　　　　　　　'中学校科目１
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU2                  '中学校科目２
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU3                  '中学校科目３
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU4                  '中学校科目4
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU5                  '中学校科目5
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU6                  '中学校科目6
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU7                  '中学校科目7
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU8                  '中学校科目8
'Public Const C_OMOMIZUKESYO_TYUGAKUKAMOKU9                  '中学校科目9
'Public Const C_OMOMIZUKESYO_1_1                             '１年１学期
'Public Const C_OMOMIZUKESYO_1_2                             '１年２学期
'Public Const C_OMOMIZUKESYO_1_3                             '１年３学期
'Public Const C_OMOMIZUKESYO_1                               '１年
'Public Const C_OMOMIZUKESYO_2_1                             '２年1学期
'Public Const C_OMOMIZUKESYO_2_2                             '２年２学期
'Public Const C_OMOMIZUKESYO_2_3                             '２年３学期
'Public Const C_OMOMIZUKESYO_2                               '２年
'Public Const C_OMOMIZUKESYO_3_1                             '３年1学期
'Public Const C_OMOMIZUKESYO_3_2                             '３年２学期
'Public Const C_OMOMIZUKESYO_3_3                             '３年３学期
'Public Const C_OMOMIZUKESYO_3                               '３年
'Public Const C_OMOMIZUKESYO_DANKAITI                        '段階値合計
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接学業評点１
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接学業評点２
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接学業評点３
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接学業評点４
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接学業評点５
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接態度評点１
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接態度評点２
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接態度評点３
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接態度評点４
'Public Const C_OMOMIZUKESYO_MENSETUGAKUGYOU                 '面接態度評点５
'Public Const C_OMOMIZUKESYO_MENSETU                         '面接合計
'Public Const C_OMOMIZUKESYO_SAKUBUNNAIYOU                   '作文内容
'Public Const C_OMOMIZUKESYO_SAKUBUNHYOUKI                   '作文表記
'Public Const C_OMOMIZUKESYO_SAKUBUNNGOUKEI                  '作文合計
'Public Const C_OMOMIZUKESYO_SUPOTU                          'スポーツ能力等評点
'Public Const C_OMOMIZUKESYO_KAMOKU1                         '科目１
'Public Const C_OMOMIZUKESYO_KAMOKU2                         '科目２
'Public Const C_OMOMIZUKESYO_KAMOKU3                         '科目３
'Public Const C_OMOMIZUKESYO_KAMOKU4                         '科目４
'Public Const C_OMOMIZUKESYO_KAMOKU5                         '科目５
'Public Const C_OMOMIZUKESYO_5KAMOKUGOUKEI                   '５科目合計
'Public Const C_OMOMIZUKESYO_MENSETU                         '面接
'Public Const C_OMOMIZUKESYO_KAMOKU2-1                       '科目１
'Public Const C_OMOMIZUKESYO_KAMOKU2-2                       '科目２
'Public Const C_OMOMIZUKESYO_KAMOKU2-3                       '科目３
'Public Const C_OMOMIZUKESYO_KAMOKU2-4                       '科目４
'Public Const C_OMOMIZUKESYO_KAMOKU2-5                       '科目５
'Public Const C_OMOMIZUKESYO_KAMOKUGOUKEI                    '科目合計
'Public Const C_OMOMIZUKESYO_SENMONKAMOKU1                   '専門科目１
'Public Const C_OMOMIZUKESYO_SENMONKAMOKU2                   '専門科目２
'Public Const C_OMOMIZUKESYO_SENMONKAMOKU3                   '専門科目3
'Public Const C_OMOMIZUKESYO_SENMONKAMOKU4                   '専門科目4
'Public Const C_OMOMIZUKESYO_SENMONKAMOKU5                   '専門科目5
'Public Const C_OMOMIZUKESYO_SENMONKAMOKUGOUKEI              '専門科目合計
'Public Const C_OMOMIZUKESYO_MENSETUKAN1                     '面接官１
'Public Const C_OMOMIZUKESYO_MENSETUKAN2                     '面接官２
'Public Const C_OMOMIZUKESYO_MENSETUKAN3                     '面接官３
'Public Const C_OMOMIZUKESYO_MENSETUKAN4                     '面接官４
'Public Const C_OMOMIZUKESYO_MENSETUKAN5                     '面接官５
'入試処理対象小分類区分       (C_NYUSI_SYOBUN=90)
Public Const C_NYUSI_SYO_SUISEN_KBN_1 = 1               '推薦区分１
Public Const C_NYUSI_SYO_SUISEN_KBN_2 = 2               '推薦区分２
Public Const C_NYUSI_SYO_5DANKAI = 11                   '５段階に対する処理
Public Const C_NYUSI_SYO_10DANKAI = 12                  '10段階に対する処理
'2001/09/08 Comment Start ->
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU1 = 1                 '中学校科目1
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU2 = 2                 '中学校科目2
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU3 = 3                 '中学校科目3
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU4 = 4                 '中学校科目4
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU5 = 5                 '中学校科目5
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU6 = 6                 '中学校科目6
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU7 = 7                 '中学校科目7
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU8 = 8                 '中学校科目8
'Public Const C_NYUSI_SYO_TYUGAKUKAMOKU9 = 9                 '中学校科目9
''2001/08/16 Add ->
'Public Const C_NYUSI_SYO_TYUGAKU_CONV = 10                  '異種中学校成績変換値(通常の段階値と異なる場合に使用：５段階なのに１０段階だった場合)
'
'Public Const C_NYUSI_SYO_DANKAITI_1_1 = 11                  '段階値１年１学期
'Public Const C_NYUSI_SYO_DANKAITI_1_2 = 12                  '段階値１年２学期
'Public Const C_NYUSI_SYO_DANKAITI_1_3 = 13                  '段階値１年３学期
'Public Const C_NYUSI_SYO_DANKAITI_1 = 14                    '段階値１年
'Public Const C_NYUSI_SYO_DANKAITI_2_1 = 15                  '段階値２年１学期
'Public Const C_NYUSI_SYO_DANKAITI_2_2 = 16                  '段階値２年２学期
'Public Const C_NYUSI_SYO_DANKAITI_2_3 = 17                  '段階値２年３学期
'Public Const C_NYUSI_SYO_DANKAITI_2 = 18                    '段階値２年
'Public Const C_NYUSI_SYO_DANKAITI_3_1 = 19                  '段階値３年１学期
'Public Const C_NYUSI_SYO_DANKAITI_3_2 = 20                  '段階値３年２学期
'Public Const C_NYUSI_SYO_DANKAITI_3_3 = 21                  '段階値３年３学期
'Public Const C_NYUSI_SYO_DANKAITI_3 = 22                    '段階値３年
''Public Const C_NYUSI_SYO_DANKAITI_GOKEI = 23                '段階値合計
'Public Const C_NYUSI_SYO_DANKAITI_GOKEI_SOTEN = 23          '段階値合計素点 2001/08/16 Add
'Public Const C_NYUSI_SYO_DANKAITI_GOKEI_TYOSEITEN = 24      '段階値合計調整点 2001/08/16 Add
'
'Public Const C_NYUSI_SYO_MENSETU_GAKUGYOU_1 = 31            '面接学業評点１     (推薦↓)
'Public Const C_NYUSI_SYO_MENSETU_GAKUGYOU_2 = 32            '面接学業評点２
'Public Const C_NYUSI_SYO_MENSETU_GAKUGYOU_3 = 33            '面接学業評点３
'Public Const C_NYUSI_SYO_MENSETU_GAKUGYOU_4 = 34            '面接学業評点４
'Public Const C_NYUSI_SYO_MENSETU_GAKUGYOU_5 = 35            '面接学業評点５
'Public Const C_NYUSI_SYO_MENSETU_TAIDO_1 = 36               '面接態度評点１
'Public Const C_NYUSI_SYO_MENSETU_TAIDO_2 = 37               '面接態度評点２
'Public Const C_NYUSI_SYO_MENSETU_TAIDO_3 = 38               '面接態度評点３
'Public Const C_NYUSI_SYO_MENSETU_TAIDO_4 = 39               '面接態度評点４
'Public Const C_NYUSI_SYO_MENSETU_TAIDO_5 = 40               '面接態度評点５
'Public Const C_NYUSI_SYO_MENSETU_GOKEI = 41                 '面接合計
'
'Public Const C_NYUSI_SYO_SAKUBUN_NAIYOU = 42                '作文内容
'Public Const C_NYUSI_SYO_SAKUBUN_HYOUKI = 43                '作文表記
'Public Const C_NYUSI_SYO_SAKUBUNN_GOKEI = 44                '作文合計
'
'Public Const C_NYUSI_SYO_SPORTS = 45                        'スポーツ能力等評点
'Public Const C_NYUSI_SYO_SUISENSYO = 46                     '推薦書
'Public Const C_NYUSI_SYO_SYOKEI = 47                        '小計
'
'Public Const C_NYUSI_SYO_KAMOKU1_GAK = 51                   '科目１     (学力↓)
'Public Const C_NYUSI_SYO_KAMOKU2_GAK = 52                   '科目２
'Public Const C_NYUSI_SYO_KAMOKU3_GAK = 53                   '科目３
'Public Const C_NYUSI_SYO_KAMOKU4_GAK = 54                   '科目４
'Public Const C_NYUSI_SYO_KAMOKU5_GAK = 55                   '科目５
'Public Const C_NYUSI_SYO_KAMOKU_GOKEI_GAK = 56              '５科目合計
'
'Public Const C_NYUSI_SYO_MENSETU_GAK = 57                   '面接
'Public Const C_NYUSI_SYO_NAISINTEN = 58                     '内申点
'
'Public Const C_NYUSI_SYO_KAMOKU1_HEN = 61                   '科目１     (編入↓)
'Public Const C_NYUSI_SYO_KAMOKU2_HEN = 62                   '科目２
'Public Const C_NYUSI_SYO_KAMOKU3_HEN = 63                   '科目３
'Public Const C_NYUSI_SYO_KAMOKU4_HEN = 64                   '科目４
'Public Const C_NYUSI_SYO_KAMOKU5_HEN = 65                   '科目５
'Public Const C_NYUSI_SYO_KAMOKU_GOKEI_HEN = 66              '５科目合計
'
'Public Const C_NYUSI_SYO_SENMON_KAMOKU1 = 67                '専門科目１
'Public Const C_NYUSI_SYO_SENMON_KAMOKU2 = 68                '専門科目２
'Public Const C_NYUSI_SYO_SENMON_KAMOKU3 = 69                '専門科目3
'Public Const C_NYUSI_SYO_SENMON_KAMOKU4 = 70                '専門科目4
'Public Const C_NYUSI_SYO_SENMON_KAMOKU5 = 71                '専門科目5
'Public Const C_NYUSI_SYO_SENMON_KAMOKU_GOKEI = 72           '専門科目合計
'
'Public Const C_NYUSI_SYO_MENSETUKAN1 = 73                   '面接官１
'Public Const C_NYUSI_SYO_MENSETUKAN2 = 74                   '面接官２
'Public Const C_NYUSI_SYO_MENSETUKAN3 = 75                   '面接官３
'Public Const C_NYUSI_SYO_MENSETUKAN4 = 76                   '面接官４
'Public Const C_NYUSI_SYO_MENSETUKAN5 = 77                   '面接官５
'Public Const C_NYUSI_SYO_MENSETU_HYOKAKEI = 78              '面接評価計
'
'Public Const C_NYUSI_SYO_TYOSASYO = 81                      '調査書
'Public Const C_NYUSI_SYO_SONOTA = 82                        'その他 2001/08/16 Add
''2001/08/16 Mod ->
''Public Const C_NYUSI_SYO_KENKO_SINDAN = 82                  '健康診断
''Public Const C_NYUSI_SYO_SOU_GOKEI = 83                     '合計
''Public Const C_NYUSI_SYO_BIKO = 84                          '備考
'Public Const C_NYUSI_SYO_KENKO_SINDAN = 83                  '健康診断
'Public Const C_NYUSI_SYO_SOU_GOKEI = 84                     '合計
'Public Const C_NYUSI_SYO_BIKO = 85                          '備考
''2001/08/16 Mod <-
'2001/09/08 Comment End <-


'委員区分(C_IIN=91)
Public Const C_IIN_GAKKO = 1    '学校毎委員
Public Const C_IIN_CLASS = 2    'クラス毎委員

'入退寮区分(C_NYUTAIRYO=92)
Public Const C_NYUTAIRYO_NYURYO = 1     '入寮者
Public Const C_NYUTAIRYO_TAIRYO = 2     '退寮者

'広さの単位区分(C_HIROSA_TANI=93)
Public Const C_HIROSA_HEIHO = 1        '㎡
Public Const C_HIROSA_JYO = 2          '畳

'部屋状況区分(C_HEYA_JOKYO=94)
Public Const C_HEYA_SIYOUFUKA = 1          '使用不可
Public Const C_HEYA_MANSITU = 2            '満室
Public Const C_HEYA_KUSITU = 3             '空室
Public Const C_HEYA_ITIBUKUSITU = 4        '一部空室

'入寮理由区分(C_NYURYO_RIYU=95)
Public Const C_NYURYORIYU_TUGAKU = 1   '通学不可能
Public Const C_NYURYORIYU_KOTU = 2     '交通不便
Public Const C_NYURYORIYU_KENKYU = 3   '研究のため
Public Const C_NYURYORIYU_SONOTA = 4   'その他

'出身校区分(C_SYUSSINKO=96)
Public Const C_SYUSSIN_KOKO = 1     '高校
Public Const C_SYUSSIN_KOSEN = 2    '他高専
Public Const C_SYUSSIN_GAIKOKU = 3  '外国
Public Const C_SYUSSIN_SONOTA = 4   'その他

'判定対象区分(C_HANTEI_TAISYO=97)
Public Const C_HANTEI_TAISYO_YES = 0     '対象
Public Const C_HANTEI_TAISYO_NO = 1  '対象外

'選択科目種別区分(C_SEN_SYUBETU=98)
Public Const C_SENTAKU_TUJO = 0    '通常選択
Public Const C_SENTAKU_JIYU = 1    '自由選択
Public Const C_SENTAKU_NINI = 2    '任意選択

'レベル別科目区分(C_LEVEL_BETU=99)
Public Const C_LEVEL_NO = 0                     'レベル別科目でない
Public Const C_LEVEL_YES = 1                    'レベル別科目である

'選択可否区分(C_SENTAKU_KAHI=100)
Public Const C_SENTAKU_NO = 0                   '選択しない
Public Const C_SENTAKU_YES = 1                  '選択する

'休日フラグ(C_KYUJITU_FLG=102)
Public Const C_HEIJITU = 0                      '平日
Public Const C_DONITI = 1                       '土日
Public Const C_SYUKUJITU = 2                    '祝日

'授業区分(C_JUGYO_KBN=103)
Public Const C_JUGYO_KBN_JUHYO = 0          '授業とみなす
Public Const C_JUGYO_KBN_NOT_JUGYO = 1      '授業とみなさない

'試験実施区分(C_SIKEN_KBN=104)
Public Const C_SIKEN_KBN_MINYU = 0          '未入力
Public Const C_SIKEN_KBN_JISSI = 1          '実施
Public Const C_SIKEN_KBN_NOT_JISSI = 2      '実施しない
Public Const C_SIKEN_KBN_NOT_JUGYO = 3      '授業悲実施

'カウント区分(C_COUNT_KBN=105)
Public Const C_COUNT_KBN_GYOJI = 0          '行事
Public Const C_COUNT_KBN_JUGYO = 1          '授業
Public Const C_COUNT_KBN_ETC = 2            'その他
Public Const C_COUNT_KBN_BOTH = 3           '両方　2001/10/25

'修了退学区分(C_SYUTAI_KBN = 106)
Public Const C_SYUTAI_KBN_FUKANO = 0         '不可能
Public Const C_SYUTAI_KBN_KANO = 1           '可能

'退学区分(C_TAIGKU_KBN = 107)
Public Const C_TAIGAKU_KBN_SINAI = 0         'しない
Public Const C_TAIGAKU_KBN_SURU = 1          'する

'処理区分（C_SYORI_KBN = 109）         '処理区分

'レベル区分（C_LEVEL_KBN = 110)        'レベル区分
Public Const C_LV_1 = 1
Public Const C_LV_2 = 2
Public Const C_LV_3 = 3
Public Const C_LV_4 = 4
Public Const C_LV_5 = 5

'置換科目フラグ(C_TIKAN_KAMOKU = 112)
Public Const C_TIKAN_KAMOKU_NASI = 0    '置換なし
Public Const C_TIKAN_KAMOKU_MOTO = 1    '置換元
Public Const C_TIKAN_KAMOKU_SAKI = 2    '置換先

'希望フラグ(C_KIBOU_FLG = 113)
Public Const C_KIBOU_FLG_NASI = 0       'なし
Public Const C_KIBOU_FLG_1 = 1          '第一希望
Public Const C_KIBOU_FLG_2 = 2          '第二希望
Public Const C_KIBOU_FLG_3 = 3          '第三希望
Public Const C_KIBOU_FLG_KETTEI = 9     '決定

'評価不可区分(C_HYOKA_FUKA = 114)
Public Const C_HYOKA_FUKA_NASI = 0      'なし
Public Const C_HYOKA_FUKA_SESEKI = 1    '成績
Public Const C_HYOKA_FUKA_KEKKA = 2     '欠課
Public Const C_HYOKA_FUKA_BOTH = 3      '成績、欠課

'休学区分(C_KYUGAKU_KBN = 117)
Public Const C_KYUGAKU_KBN_ZAI = 0      '在学中
Public Const C_KYUGAKU_KBN_KYU = 1      '休学中


'評価対象区分(C_HYOKA_TAISYO = 119)
Public Const C_HYOKA_TAISHO_IPPAN = 0   '一般学科
Public Const C_HYOKA_TAISHO_SENKOU = 1  '専攻科
Public Const C_HYOKA_TAISHO_HOKA = 2    '他大学


'入試確定区分(C_NYUSI_KAKUTEI = 120)
Public Const C_NYUSI_KAKUTEI_MITEI = 0      '未定
Public Const C_NYUSI_KAKUTEI_KAKUTEI = 1    '確定

'学力試験応募区分(C_GAKUNYUSI_OUBO KBN = 121)
Public Const C_NYUSI_GAKUOUBO_KBN_NO = 0    '希望しない
Public Const C_NYUSI_GAKUOUBO_KBN_YES = 1   '希望する

'使用フラグ(C_NYUSI_SIYOU_KBN = 122)
Public Const C_NYUSI_SIYOU_KBN_NO = 0       '使用しない
Public Const C_NYUSI_SIYOU_KBN_YES = 1      '使用する

'手入力区分(C_TENYURYOKU_KBN = 123)
Public Const C_TENYURYOKU_KBN_NO = 0        '使用しない
Public Const C_TENYURYOKU_KBN_YES = 1       '使用する

'追加合格候補区分(C_TUIKA_KOHO_KBN=124)
Public Const C_TUIKA_KOHO_KBN_NO = 0        '無関係
Public Const C_TUIKA_KOHO_KBN_YES = 1       '追加合格候補

'追加合格区分(C_TUIKA_GOKAKU_KBN=125)
Public Const C_TUIKA_GOKAKU_KBN_NO = 0      '無関係
Public Const C_TUIKA_GOKAKU_KBN_YES = 1     '合格

'過年度生フラグ(C_KANENDO_KBN=126)
Public Const C_KANENDO_KBN_NO = 0           '過年度生でない
Public Const C_KANENDO_KBN_YES = 1          '過年度生

'入学受付フラグ(C_UKETUKE_KBN=127)
Public Const C_UKETUKE_KBN_OK = 0           '入学
Public Const C_UKETUKE_KBN_CANCEL = 1       '入学辞退

'評価予定区分(C_HYOKAYOTEI_KBN=128)
Public Const C_HYOKAYOTEI_KBN_1MARU = 1     '○
Public Const C_HYOKAYOTEI_KBN_2MARU = 2     '◎

'端数処理区分(C_HASU_SYORI_KBN = 129)
Public Const C_HASU_SYORI_KIRISUTE = 0      '切り捨て
Public Const C_HASU_SYORI_KIRIAGE = 1       '切り上げ
Public Const C_HASU_SYORI_SISYAGONYU = 2    '四捨五入

'中学校成績段階数区分(C_TYUGAKU_DAN_KBN =130)
Public Const C_TYUSEI_DAN_KBN_5DAN = 1     '5段階
Public Const C_TYUSEI_DAN_KBN_10DAN = 2    '10段階

'中学校成績入力段階数区分(public const C_TYUGAKU_NYU_KBN=131)
Public Const C_TYUSEI_NYU_KBN_ONAZI = 0    '同じ
Public Const C_TYUSEI_NYU_KBN_KOTONARU = 1 '異なる

'平均点科目区分(C_HEIKIN_KAMOKU_KBN = 132 )
Public Const C_HEIKIN_KAMOKU_KBN_OFF = 0        '平均に含まない
Public Const C_HEIKIN_KAMOKU_KBN_ON = 1         '平均に含む


'欠課欠席情報区分(C_KEKKA_JYOHOU_KBN = 133)
Public Const C_KEKKA_JYOHOU_KBN_JIKAN = 1       '最低授業時間数
Public Const C_KEKKA_JYOHOU_KBN_KAMOKU = 2      '欠課科目数条件
Public Const C_KEKKA_JYOHOU_KBN_HEIKIN = 3      '平均点条件
Public Const C_KEKKA_JYOHOU_KBN_KEKKA = 4      '欠課換算条件
Public Const C_KEKKA_JYOHOU_KBN_SYUSSEKI = 5       '出席時数条件
Public Const C_KEKKA_JYOHOU_KBN_KESSEKI = 6       '欠席換算条件

'数値基準区分(C_SUUCHI_KIJYUN_KBN = 134)
Public Const C_SUUCHI_KIJYUN_KBN_NO = 0     '含まない
Public Const C_SUUCHI_KIJYUN_KBN_INC = 1    '含む

'入試区分(C_NYUSI_KBN = 135)
Public Const C_NYUSI_KBN_SUISEN = 1         '推薦入学
Public Const C_NYUSI_KBN_GAKURYOKU = 2      '学力試験
Public Const C_NYUSI_KBN_HENNYU = 3         '編入試験

'学力移行済みフラグ(C_NYUSI_IKO_FLG = 136)
Public Const C_NYUSI_IKO_NO = 0             '移行していない
Public Const C_NYUSI_IKO_YES = 1            '移行した

'入試出席フラグ(C_NYUSI_SUSSEKI_FLG = 137)
Public Const C_NYUSI_SUSSEKI_YES = 0        '出席
Public Const C_NYUSI_SUSSEKI_NO = 1         '欠席

'入試取り消しフラグ(C_NYUSI_TORIKESI_FLG = 138)
Public Const C_NYUSI_TORIKESI_NO = 0        '無関係
Public Const C_NYUSI_TORIKESI_YES = 1       '入学取り消し

'判定使用フラグ(C_HANTEI_FLG = 139)
Public Const C_HANTEI_NO = 0        '判定に使用しない
Public Const C_HANTEI_YES = 1       '判定に使用する

'入試健康診断区分(C_NYUSI_KENKO_SINDAN_KBN = 140)
Public Const C_NYUSI_KENKO_HUYO = 0         '検査不要
Public Const C_NYUSI_KENKO_SAIKENSA = 1     '要再検
Public Const C_NYUSI_KENKO_SEIMITU = 2      '要精密

'特別活動評価区分(C_HYOKATOKU_KBN = 141)
Public Const C_HYOKATOKU_GOUKAKU = 0    '合格
Public Const C_HYOKATOKU_FUGOKAKU = 1   '不合格

'証明書区分(C_SYOMEISYO_KBN=143）
Public Const C_SYO_ZAIGAKU = 1              '在学証明書
Public Const C_SYO_SOTUGYO = 2              '卒業証明書
Public Const C_SYO_MIKOMI = 3               '卒業見込証明書
Public Const C_SYO_GAKUGYO = 4              '学業成績証明書
Public Const C_SYO_TANNISYU = 5             '単位修得証明書
Public Const C_SYO_SYURYO = 6               '修了証明書
Public Const C_SYO_SYURYO_M = 7             '修了証明書(毛筆）
Public Const C_SYO_TYOSASYO = 8             '調査書

'C_NYUSI_SOTUGYO_FLG = 144      '入試卒業フラグ
Public Const C_NYUSI_SOTUGYO_NO = 0         '卒業見込み
Public Const C_NYUSI_SOTUGYO_YES = 1        '卒業

'推薦入試処理対象区分(C_NYUSI_SUISEN_SYORI_KBN = 145)
Public Const C_NYUSI_SUISEN_SEISEKI_1 = 1           '推薦成績１
Public Const C_NYUSI_SUISEN_SEISEKI_2 = 2           '推薦成績２
Public Const C_NYUSI_SUISEN_SEISEKI_3 = 3           '推薦成績３
Public Const C_NYUSI_SUISEN_SEISEKI_4 = 4           '推薦成績４
Public Const C_NYUSI_SUISEN_SEISEKI_5 = 5           '推薦成績５
Public Const C_NYUSI_SUISEN_SEISEKI_6 = 6           '推薦成績６
Public Const C_NYUSI_SUISEN_SEISEKI_7 = 7           '推薦成績７
Public Const C_NYUSI_SUISEN_SEISEKI_8 = 8           '推薦成績８
Public Const C_NYUSI_SUISEN_SEISEKI_9 = 9           '推薦成績９
Public Const C_NYUSI_SUISEN_SEISEKI_10 = 10         '推薦成績１０
Public Const C_NYUSI_SUISEN_SEISEKI_11 = 11         '推薦成績１１
Public Const C_NYUSI_SUISEN_SEISEKI_12 = 12         '推薦成績１２
Public Const C_NYUSI_SUISEN_SEISEKI_13 = 13         '推薦成績１３
Public Const C_NYUSI_SUISEN_SEISEKI_14 = 14         '推薦成績１４
Public Const C_NYUSI_SUISEN_SEISEKI_15 = 15         '推薦成績１５
Public Const C_NYUSI_SUISEN_SEISEKI_16 = 16         '推薦成績１６
Public Const C_NYUSI_SUISEN_SEISEKI_17 = 17         '推薦成績１７
Public Const C_NYUSI_SUISEN_SEISEKI_18 = 18         '推薦成績１８
Public Const C_NYUSI_SUISEN_SEISEKI_19 = 19         '推薦成績１９
Public Const C_NYUSI_SUISEN_SEISEKI_20 = 20         '推薦成績２０
Public Const C_NYUSI_SUISEN_DANKAITI_1 = 21         '段階値合計１
Public Const C_NYUSI_SUISEN_DANKAITI_2 = 22         '段階値合計２
Public Const C_NYUSI_SUISEN_GOKEI_1 = 23            '推薦成績合計１
Public Const C_NYUSI_SUISEN_GOKEI_2 = 24            '推薦成績合計２
Public Const C_NYUSI_SUISEN_BIKO = 25               '備考
Public Const C_NYUSI_SUISEN_KENKO_SINDAN = 26       '健康診断

'学力入試処理対象区分(C_NYUSI_GAKURYOKU_SYORI_KBN = 146)
Public Const C_NYUSI_GAKURYOKU_KAMOKU_1 = 1             '科目１
Public Const C_NYUSI_GAKURYOKU_KAMOKU_2 = 2             '科目２
Public Const C_NYUSI_GAKURYOKU_KAMOKU_3 = 3             '科目３
Public Const C_NYUSI_GAKURYOKU_KAMOKU_4 = 4             '科目４
Public Const C_NYUSI_GAKURYOKU_KAMOKU_5 = 5             '科目５
Public Const C_NYUSI_GAKURYOKU_KAMOKU_GOKEI_1 = 6       '科目合計１
Public Const C_NYUSI_GAKURYOKU_KAMOKU_GOKEI_2 = 7       '科目合計２
Public Const C_NYUSI_GAKURYOKU_SEISEKI_1 = 8            '学力成績１
Public Const C_NYUSI_GAKURYOKU_SEISEKI_2 = 9            '学力成績２
Public Const C_NYUSI_GAKURYOKU_SEISEKI_3 = 10           '学力成績３
Public Const C_NYUSI_GAKURYOKU_SEISEKI_4 = 11           '学力成績４
Public Const C_NYUSI_GAKURYOKU_SEISEKI_5 = 12           '学力成績５
Public Const C_NYUSI_GAKURYOKU_SEISEKI_6 = 13           '学力成績６
Public Const C_NYUSI_GAKURYOKU_SEISEKI_7 = 14           '学力成績７
Public Const C_NYUSI_GAKURYOKU_SEISEKI_8 = 15           '学力成績８
Public Const C_NYUSI_GAKURYOKU_SEISEKI_9 = 16           '学力成績９
Public Const C_NYUSI_GAKURYOKU_SEISEKI_10 = 17          '学力成績１０
Public Const C_NYUSI_GAKURYOKU_SEISEKI_11 = 18          '学力成績１１
Public Const C_NYUSI_GAKURYOKU_SEISEKI_12 = 19          '学力成績１２
Public Const C_NYUSI_GAKURYOKU_SEISEKI_13 = 20          '学力成績１３
Public Const C_NYUSI_GAKURYOKU_SEISEKI_14 = 21          '学力成績１４
Public Const C_NYUSI_GAKURYOKU_SEISEKI_15 = 22          '学力成績１５
Public Const C_NYUSI_GAKURYOKU_DANKAITI_1 = 23          '段階値合計１
Public Const C_NYUSI_GAKURYOKU_DANKAITI_2 = 24          '段階値合計２
Public Const C_NYUSI_GAKURYOKU_GOKEI_1 = 25             '学力成績合計１
Public Const C_NYUSI_GAKURYOKU_GOKEI_2 = 26             '学力成績合計２
Public Const C_NYUSI_GAKURYOKU_SOU_GOKEI_1 = 27         '学力成績総合計１
Public Const C_NYUSI_GAKURYOKU_SOU_GOKEI_2 = 28         '学力成績総合計２
Public Const C_NYUSI_GAKURYOKU_BIKO = 29                '備考
Public Const C_NYUSI_GAKURYOKU_KENKO_SINDAN = 30        '健康診断

'編入学処理対象区分(C_NYUSI_HENNYU_SYORI_KBN = 147)
Public Const C_NYUSI_HENNYU_SEISEKI_1 = 1           '編入成績１
Public Const C_NYUSI_HENNYU_SEISEKI_2 = 2           '編入成績２
Public Const C_NYUSI_HENNYU_SEISEKI_3 = 3           '編入成績３
Public Const C_NYUSI_HENNYU_SEISEKI_4 = 4           '編入成績４
Public Const C_NYUSI_HENNYU_SEISEKI_5 = 5           '編入成績５
Public Const C_NYUSI_HENNYU_SEISEKI_6 = 6           '編入成績６
Public Const C_NYUSI_HENNYU_SEISEKI_7 = 7           '編入成績７
Public Const C_NYUSI_HENNYU_SEISEKI_8 = 8           '編入成績８
Public Const C_NYUSI_HENNYU_SEISEKI_9 = 9           '編入成績９
Public Const C_NYUSI_HENNYU_SEISEKI_10 = 10         '編入成績１０
Public Const C_NYUSI_HENNYU_SEISEKI_11 = 11         '編入成績１１
Public Const C_NYUSI_HENNYU_SEISEKI_12 = 12         '編入成績１２
Public Const C_NYUSI_HENNYU_SEISEKI_13 = 13         '編入成績１３
Public Const C_NYUSI_HENNYU_SEISEKI_14 = 14         '編入成績１４
Public Const C_NYUSI_HENNYU_SEISEKI_15 = 15         '編入成績１５
Public Const C_NYUSI_HENNYU_SEISEKI_16 = 16         '編入成績１６
Public Const C_NYUSI_HENNYU_SEISEKI_17 = 17         '編入成績１７
Public Const C_NYUSI_HENNYU_SEISEKI_18 = 18         '編入成績１８
Public Const C_NYUSI_HENNYU_SEISEKI_19 = 19         '編入成績１９
Public Const C_NYUSI_HENNYU_SEISEKI_20 = 20         '編入成績２０
Public Const C_NYUSI_HENNYU_GOKEI_1 = 21            '編入成績合計１
Public Const C_NYUSI_HENNYU_GOKEI_2 = 22            '編入成績合計２
Public Const C_NYUSI_HENNYU_SOU_GOKEI_1 = 23        '編入成績総合計１
Public Const C_NYUSI_HENNYU_SOU_GOKEI_2 = 24        '編入成績総合計２
Public Const C_NYUSI_HENNYU_BIKO = 25               '備考
Public Const C_NYUSI_HENNYU_KENKO_SINDAN = 26       '健康診断

'中学校成績処理対象区分(C_NYUSI_TYUGAKKO_SYORI_KBN = 148)
Public Const C_NYUSI_TYUGAKKO_KAMOKU_1 = 1      '中学校科目１
Public Const C_NYUSI_TYUGAKKO_KAMOKU_2 = 2      '中学校科目２
Public Const C_NYUSI_TYUGAKKO_KAMOKU_3 = 3      '中学校科目３
Public Const C_NYUSI_TYUGAKKO_KAMOKU_4 = 4      '中学校科目4
Public Const C_NYUSI_TYUGAKKO_KAMOKU_5 = 5      '中学校科目5
Public Const C_NYUSI_TYUGAKKO_KAMOKU_6 = 6      '中学校科目6
Public Const C_NYUSI_TYUGAKKO_KAMOKU_7 = 7      '中学校科目7
Public Const C_NYUSI_TYUGAKKO_KAMOKU_8 = 8      '中学校科目8
Public Const C_NYUSI_TYUGAKKO_KAMOKU_9 = 9      '中学校科目9
Public Const C_NYUSI_TYUGAKKO_GOKEI = 10        '合計



'*************************
'コード定義(区分マスタ以外)
'*************************

Public Const C_K_JOTAI = 1          '状態区分
Public Const C_K_KOJIN_5NEN = 2     '学籍５年間個人番号呼称
Public Const C_K_KOJIN_1NEN = 3     '学籍１年間個人番号呼称
Public Const C_K_GAK_NENDO = 4      '学籍年度フラグ
Public Const C_K_GAK_JOTAI = 5      '学籍状態区分
Public Const C_K_GAK_SINNYU = 6     '学籍新入生状態区分
Public Const C_K_GAK_NENJI = 7      '学籍年次処理区分
Public Const C_K_ZEN_KAISI = 10     '前期開始日
Public Const C_K_KOU_KAISI = 11     '後期開始日
Public Const C_K_KOU_SYURYO = 12    '後期終了日
Public Const C_K_WAREKI_NENDO = 13  '和暦年度
Public Const C_K_SIKEN_JIGEN = 14   '試験時限
Public Const C_K_SINKYU_HANTEI = 15 '進級判定状態区分
Public Const C_K_RISYU_TANI = 16    '履修単位入力区分
Public Const C_NYUSI_TAISYO_SUISEN_CSV = 19     '推薦CSV番号
Public Const C_NYUSI_TAISYO_GAKURYOKU_CSV = 20  '学力CSV番号
Public Const C_NYUSI_TAISYO_HENNYU_CSV = 21     '編入CSV番号

Public Const C_K_JIKANWARI = 23     '授業時間割確定区分

Public Const C_K_DANKAISU = 24                  '基本中学校成績段階数
Public Const C_K_KIBO_GAKKASU_SUISEN = 25       '推薦志望可能学科数
Public Const C_K_KIBO_GAKKASU_GAKURYOKU = 26    '学力志望可能学科数
Public Const C_K_KIBO_GAKKASU_HENNYU = 27       '編入志望可能学科数

Public Const C_K_RIS_JOUTAI = 28                '履修状態区分
'Public Const C_K_JIK_ZENKI = 29                 '前期授業時間割状態区分    2001/10/04佐野廃止
'Public Const C_K_JIK_KOUKI = 30                 '後期授業時間割状態区分

Public Const C_K_KEKKA_KESSEKI = 31             '欠課・欠席設定条件区分
Public Const C_K_KEKKA_RUISEKI = 32             '欠課累積情報区分
Public Const C_K_TANI_SAITEI_JIKAN = 33         '一単位最低授業時間数

'2001/09/04廃止
'Public Const C_K_ZEN_SYURYO = 34     '前期終了日
'Public Const C_K_NEN_KAISI = 35      '年度開始日
'Public Const C_K_NEN_SYURYO = 36     '年度終了日

Public Const C_K_RISKOJIN_JOUTAI = 37           '個人履修状態区分
Public Const C_K_HANTEI_JOUTAI = 38             '判定状態区分

Public Const C_K_NYUGAKU_KAISEI = 39             '入学回生区分  2001/10/16 追加
Public Const C_K_SOTUSYOSYO_MAXNO = 40           '卒業証書発行番号(最大値） 2001/10/19 追加

'***********************
'管理マスタ（種別情報）
'***********************
'年度
'管理マスタの年度取得条件
Public Const C_JYO_NENDO = 9999
Public Const C_JYO_NO = 0
Public Const C_JYO_SYUBETU = 0

'状態区分(C_KANRI_JOTAI=1)
Public Const C_K_JOTAI_SYOKI = 0        '初期状態
Public Const C_K_JOTAI_NYUSI = 0        '入試設定確定済
Public Const C_K_JOTAI_GAKU = 0         '学籍設定確定済
Public Const C_K_JOTAI_SINKYU = 0       '進級判定確定済
Public Const C_K_JOTAI_SOTUGYO = 0      '卒業判定確定済

'学籍年度フラグ(C_K_GAK_NENDO=4)
Public Const C_K_GAK_NENDO_NO = 0      '学籍処理年でない
Public Const C_K_GAK_NENDO_YES = 1     '学籍処理年です

'学籍状態区分(C_K_GAK_JOTAI=5)
Public Const C_K_GAK_JOTAI_BEFORE = 0    '取込前
Public Const C_K_GAK_JOTAI_AFTER = 1     '取込後
Public Const C_K_GAK_JOTAI_KAKUTEI = 2   '確定済

'学籍新入生状態区分(C_K_GAK_SINNYU=6)
Public Const C_K_GAK_SINNYU_BEFORE = 0   '学籍取込前
Public Const C_K_GAK_SINNYU_AFTER = 1    '学籍取込後
Public Const C_K_GAK_SINNYU_JIDO = 2     '学生番号自動設定後

'学籍年次処理区分(C_K_GAK_NENJI=7)
Public Const C_K_GAK_NENJI_BEFORE = 0   '年次処理前
Public Const C_K_GAK_NENJI_AFTER = 1   '年次処理後

'試験時間割管理区分(C_K_SIKEN_JIKAN = 14)   2001/07/30
Public Const C_K_SIKEN_JIKAN_JIKAN = 0     '試験を時限で管理せず
Public Const C_K_SIKEN_JIKAN_JIGEN = 1     '試験を時限で管理する

'履修単位入力区分(C_K_RISYU_TANI = 16)
Public Const C_K_RISYU_TANI_INTEGER = 0     '整数入力
Public Const C_K_RISYU_TANI_DECIMAL = 1     '小数入力


'授業時間割確定区分(C_K_JIKANWARI = 23 )
Public Const C_K_JIKANWARI_NO = 1       '未確定
Public Const C_K_JIKANWARI_ZEN = 2      '前期確定
Public Const C_K_JIKANWARI_KOU = 3      '後期確定
Public Const C_K_JIKANWARI_ALL = 4      '全確定     2001/10/04佐野追加

'基本中学校成績段階数(Const C_K_DANKAISU = 24)
Public Const C_K_DANKAISU_5 = 1
Public Const C_K_DANKAISU_10 = 2

Public Const C_K_DANKAISU_NO5 = 5
Public Const C_K_DANKAISU_NO10 = 10

'履修状態区分(C_K_RIS_JOUTAI = 28)
Public Const C_K_RIS_MAE = 0        '確定処理前
Public Const C_K_RIS_ATO = 1        '確定処理後

''前期授業時間割状態区分(C_K_JIK_ZENKI = 29)            2001/10/04佐野廃止
'Public Const C_K_JIK_ZENKI_MAE = 0        '確定処理前
'Public Const C_K_JIK_ZENKI_ATO = 1        '確定処理後
''後期授業時間割状態区分(C_K_JIK_KOUKI = 30)
'Public Const C_K_JIK_KOUKI_MAE = 0        '確定処理前
'Public Const C_K_JIK_KOUKI_ATO = 1        '確定処理後

'欠課・欠席設定条件区分(C_K_KEKKA_KESSEKI = 31)
Public Const C_K_KEKKA_KESSEKI_KOTEI = 1        '固定式
Public Const C_K_KEKKA_KESSEKI_HEN = 2          '変則式（単位なし/休学なし）
Public Const C_K_KEKKA_KESSEKI_HEN_KYU = 3      '変則式（単位なし/休学あり）
Public Const C_K_KEKKA_KESSEKI_HEN_TANI = 4     '変則式（単位あり/休学なし）
Public Const C_K_KEKKA_KESSEKI_HEN_TANIKYU = 5  '変則式（単位あり/休学あり）

'欠課累積情報区分(C_K_KEKKA_RUISEKI = 32)
Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '試験毎
Public Const C_K_KEKKA_RUISEKI_KEI = 1      '累積

'個人履修状態区分(C_K_RISKOJIN_JOUTAI = 37)
Public Const C_K_RISKOJIN_MAE = 0        '確定処理前
Public Const C_K_RISKOJIN_ATO = 1        '確定処理後

'判定状態区分(C_K_HANTEI_JOUTAI = 38)
Public Const C_K_HANTEI_MAE = 0         '確定処理前
Public Const C_K_HANTEI_ATO = 1         '確定処理後

'******************************************
'履修処理
'******************************************

'個人履修処理区分
'(個人履修処理のクラスパラメータとして使用)
Public Const C_RISKOJIN_RYUNEN = 6 '0   '個人履修留年処理
Public Const C_RISKOJIN_TENNYU = 5 '1   '個人履修転入処理

'学科（履修登録、個人履修で使用）
Public Const C_GAKKA_ALL = "00"    '全学科共通

'******************************************
'判定処理
'******************************************
'判定処理区分（HAN0110クラスの第３パラメータ）
Public Const C_HAN_SINKYU = 3       '進級判定
Public Const C_HAN_SOTUGYO = 4      '卒業判定

'欠課・欠席設定(M15_KEKKA_SETTEI)のコード
Public Const C_K_KEKKA_SAITEI = 1   '留年判定 (最低授業時間数)
Public Const C_K_KEKKA_KAMOKU = 2   '留年判定 (欠課科目数)
Public Const C_K_KEKKA_HEIKIN = 3   '学年修了条件 (平均点)
Public Const C_K_KEKKA_TIKOKU = 4   '遅刻→欠課換算数
Public Const C_K_KEKKA_KEKKA = 5    '欠課→欠席換算数

'休学を考慮するか(M15_KEKKA_KBN)
Public Const C_K_KEKKA_NASI = 0     'なし
Public Const C_K_KEKKA_TUJO = 1     '通常
Public Const C_K_KEKKA_KYUGAKU = 2  '休学

'単位を考慮しない場合のM15_TANIの値
Public Const C_K_KEKKA_TANI_NASI = 0

'区分
Public Const C_K_KEKKA_SINAI = 0         'しない
Public Const C_K_KEKKA_SURU = 1          'する

'換算フラグ
Public Const C_KANSAN_YES = 1        '換算する
Public Const C_KANSAN_NO = 0         '換算しない

'↑欠課・欠席設定条件区分(C_K_KEKKA_KESSEKI = 31)も参照

'******************************************
'スケジュール関連
'******************************************
Public Const C_GAKUNEN_ALL = 0      '全学年
Public Const C_CLASS_ALL = 99       '全クラス

Public Const C_HEAD_LONGVACATION = 1       '長期休暇
Public Const C_HEAD_HOLIDAY = 2            '休日

'授業時間割（Ｔ20）　特別活動フラグ
Public Const C_JIK_JUGYO = 0               '通常授業
Public Const C_JIK_TOKUBETU = 1            '特別活動


'******************************************
'== 役職区分 ==
'******************************************
Public Const C_YAKUSYOKU_KOTYO = "001"             '/* 校長
Public Const C_YAKUSYOKU_SYUJI = "003"             '/* 主事
Public Const C_YAKUSYOKU_GAKKOUI = "032"           '/* 学校医

'******************************************
'== 評価形式マスタ(M08) 評価形式小分類略称 ==
'******************************************
Public Const C_HYOKAKEISIKI_KA = "0"
Public Const C_HYOKAKEISIKI_FUKA = "1"

'******************************************
'== 高専の固有番号(九州) ==
'== NCT : 国立工業高等専門学校の略 ==
'******************************************
Public Const C_NCT_KURUME 	 = "46"          '久留米高専
Public Const C_NCT_ARIKAKE 	 = "47"         '有明高専
Public Const C_NCT_KITAKYU 	 = "48"         '北九州高専
Public Const C_NCT_SASEBO 	 = "49"          '佐世保高専
Public Const C_NCT_KUMAMOTO  = "50"        '熊本電波高専
Public Const C_NCT_YATSUSIRO = "51"       '八代高専
Public Const C_NCT_MIYAZAKI  = "53"        '都城高専
Public Const C_NCT_KAGOSHIMA = "54"       '鹿児島高専
'2005.09/13 Add_S 西村
Public Const C_NCT_GIFU 	 = "23"          '岐阜高専

public Const C_DELETE0 = "0000000000"


'科目分類コード（C_KAMOKUBUNRUI_KBN = 152)  '2002/06/21
Public Const C_KAMOKUBUNRUI_TUJYO = "01"        '通常科目
Public Const C_KAMOKUBUNRUI_NINTEI = "02"       '認定科目
Public Const C_KAMOKUBUNRUI_TOKUBETU = "03"     '特別科目

'成績入力方法
Public Const C_SEISEKI_INP_TYPE_NUM = 0			'点数、欠課、遅刻
Public Const C_SEISEKI_INP_TYPE_STRING = 1		'文字、欠課、遅刻
Public Const C_SEISEKI_INP_TYPE_KEKKA = 2		'欠課、遅刻

'******************************************
'== データ区分
'== 2002/06/24
'******************************************
Public Const C_TAIGAKU		= 1         '退学
Public Const C_KYUGAKU		= 2         '休学
Public Const C_HYOKA_FUNO	= 3         '評価不能
Public Const C_MIHYOKA		= 4         '未評価

'******************************************
'== 管理マスタコード
'== 2002/07/03
'******************************************
Public Const C_GAKKO_NO        = 9995		'学校番号
Public Const C_KEKKAGAI_DISP   = 9996		'欠課対象外表示コード
Public Const C_HYOKAYOTEI_DISP = 9997      '評価予定表示コード
Public Const C_DATAKBN_DISP    = 9998      'データ区分表示コード

Public Const C_DISP_NO = 0      '表示しない
Public Const C_DISP    = 1      '表示する

%>