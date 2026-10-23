# -*- coding: utf-8 -*-
"""从微信读书**全库**整理「≤N 元 且 评分 ≥M」的书目，输出 lowprice_seed.json。

为什么是「采样」而不是「穷举」：
  微信读书没有「按价格浏览全库」的接口，搜索结果也不支持按价格排序。
  唯一能用的入口是 /web/search/name_books（JSON、支持 maxIdx 翻页），
  它的返回里**直接带 price 与 newRating**，所以不用逐本查价。
  代价是：能不能搜到某本低价书，取决于关键词是否覆盖到它。
  因此这是**采样清单**——覆盖度由 GROUPS 决定，随时可以加词/加深页数重跑。

用法：
  python fetch_lowprice.py                     # 默认 ≤2 元、评分 ≥70，写入 <技能>/data/lowprice_seed.json
  python fetch_lowprice.py --cap 5 --min-rating 75
  python fetch_lowprice.py --pages-scale 2     # 所有分组的页数翻倍（覆盖更深、更慢）
  python fetch_lowprice.py --keywords 论语,菜谱 --pages 2 --out ./seed.json

分组与页数：
  GROUPS 里每组带自己的页数——热门类目（古籍、文学、经管、心理、杂志专刊）翻 5 页，
  长尾类目（工具书、泛搜索）翻 2 页。这样把请求预算花在命中率高的地方。
"""
import argparse
import json
import os
import re
import sys
import time
from urllib.parse import quote

from playwright.sync_api import sync_playwright

import wconfig

BASE = 'https://weread.qq.com'
PAGE = 40          # 接口固定每页 40 条
HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(HERE, '..', 'data', 'lowprice_seed.json')

# (分组名, 每组翻几页, 关键词)
# 页数是「最多翻几页」，接口 hasMore=0 会提前停，所以设大一点不会浪费太多。
GROUPS = [
    # ===== 第 1/2 轮实测的高产出类目（已剔除零贡献词）=====
    ('古籍经典', 5, [
        '论语', '道德经', '老子', '庄子', '孟子', '楚辞', '唐诗', '宋词', '元曲',
        '史记', '资治通鉴', '红楼梦', '西游记', '水浒传', '三国演义', '菜根谭',
        '围炉夜话', '浮生六记', '古文观止', '声律启蒙', '弟子规', '三字经',
        '百家姓', '千字文', '孙子兵法', '三十六计', '战国策', '左传',
        '幼学琼林', '曾国藩家书', '古诗源', '纳兰词',
    ]),
    ('杂志专刊', 5, [
        '专刊', '增刊', '双周刊', '文摘', '意林', '青年文摘',
        '南风窗', '三联生活周刊', '中国新闻周刊', '南方人物周刊',
        '哈佛商业评论', '中国国家地理', '瑞丽', '合订本',
        '视野', '雪球',
    ]),
    ('文学', 5, [
        '散文', '诗歌', '小说', '杂文', '随笔', '传记', '民间故事',
        '寓言', '童话', '短篇小说', '经典阅读', '名著', '诗词',
    ]),
    ('心理成长', 5, [
        '情商', '沟通', '习惯', '认知', '幸福', '人际关系', '性格', '专注', '心理',
    ]),
    ('经管', 5, [
        '理财', '投资', '股票', '基金', '创业', '管理', '营销', '销售',
        '经济学', '财务', '商业模式',
    ]),
    ('人文社科', 4, [
        '佛学', '易经', '法律', '教育', '文化', '国学', '民俗', '心经', '六祖坛经',
        '民法典', '劳动法', '合同法', '婚姻法',
    ]),
    ('生活', 4, [
        '食谱', '养生', '瑜伽', '减肥', '宠物', '围棋', '象棋', '烘焙',
        '按摩', '伤寒论', '食疗', '药膳', '偏方', '家常菜', '汤谱',
    ]),
    ('外国名著', 4, [
        # 书名：本身足够独特，搜书名最精准（实测每个书名都能稳定命中 1–4 本）
        '傲慢与偏见', '简爱', '老人与海', '钢铁是怎样炼成的', '呼啸山庄',
        '战争与和平', '红与黑', '复活', '猎人笔记', '双城记', '三个火枪手',
        '基督山伯爵', '茶花女', '雾都孤儿', '鲁滨逊漂流记', '爱的教育', '昆虫记',
        # 作者名：**只保留多产作家**。实测「一个作者名下多本作品」时用作者名才划算
        # （凡尔纳 +4、柯南道尔 +3、小仲马 +3、屠格涅夫 +1）；
        # 其余作者名（莎士比亚/阿加莎/海明威/笛福/托尔斯泰/司汤达…）命中与书名重复或为 0，已剔除。
        '凡尔纳', '柯南道尔', '小仲马', '屠格涅夫',
    ]),
    ('泛搜索', 2, [
        '全集', '白话', '一分钟', '速成', '大全', '必读',
    ]),

    # ===== 第 3 轮新增：从已命中书名里挖出的「公约数」=====
    # 思路：对 291 本命中书做 n-gram + 贪心集合覆盖，取能覆盖最多书的书名片段。
    # 例：'雪球' 一词覆盖 49 本（雪球专刊 / 雪球·私募专刊…），比 '专刊' 覆盖面更大；
    #     '中华人民共和国' 覆盖 9 本（各单行法），'鉴赏' 覆盖 9 本诗词鉴赏类。
    ('书名公约数', 5, [
        '雪球', '鉴赏', '中华人民共和国', '视野',
        '女人', '古代', '昆虫', '商业', '智慧',
        '卡耐基', '卡内基', '巴菲特', '芒格',
    ]),

    # ===== 第 3 轮新增：出版社 =====
    # 数据来源：291 本书的 publisher 分布高度集中，头部几家贡献 20+ 本，
    # 说明这些社（多为古籍/教材/公版书）还有大量 ≤2 元的书没被类目词覆盖到。
    # 出版社：12 家里只有 3 家有产出（合计 5 本），其余 9 家 4 页内 0 本 —— 精简到 3 个
    ('出版社', 4, [
        '辽海出版社', '企业管理出版社', '华侨出版社',
    ]),

    # ===== 第 3 轮新增：历史人物 / 管理投资大师 =====
    ('人物', 3, [
        # 中国：帝王将相 / 文人 / 近现代（只保留实测有产出的）
        '曹操', '诸葛亮', '曾国藩', '王阳明', '朱元璋', '李鸿章', '孙中山',
        '胡雪岩', '张之洞', '梁启超', '鲁迅', '胡适',
        '苏东坡', '苏轼', '陶渊明', '辛弃疾', '王维', '李清照',
        '韩非', '老舍', '季羡林', '张爱玲', '丰子恺', '钱学森', '马云',
        # 外国：思想家 / 科学家 / 作家
        '弗洛伊德', '爱因斯坦', '霍金', '富兰克林', '梭罗',
        '契诃夫', '莫泊桑', '马克吐温', '杰克伦敦', '川端康成',
        '陀思妥耶夫斯基',
        # 政要 / 艺术家
        '林肯', '罗斯福', '丘吉尔', '拿破仑', '毕加索',
        # 商业 / 投资 / 管理
        '洛克菲勒', '巴菲特', '芒格', '彼得林奇', '索罗斯',
        '格雷厄姆', '稻盛和夫', '戴尔卡耐基', '比尔盖茨',
    ]),

    # ===== 思想 / 心理 / 经济 =====
    # 实测有效：阿德勒(2)、叔本华(2)、韦伯(2)、亚当斯密(1)、熊彼特(1)、加缪(1)、贝索斯(1)
    # 零贡献但保留做兜底：马克思、凯恩斯、罗素、马斯克（2 元档这类书极少，见 SKILL.md 的说明）
    # 实测有效 5 个（合计 6 本）；马克思/凯恩斯/罗素/马斯克/荣格 在 2 元档全为 0，已剔除
    ('思想心理', 3, [
        '阿德勒', '叔本华', '加缪', '熊彼特', '韦伯',
    ]),

    # ===== 公版作家 / 通俗小说 =====
    # 公版（无版权限制）作家的作品在微信读书大量低价在售。实测有效：
    # 爱默生 6、爱伦坡 4、吉卜林 2，柯林斯/梅里美/张恨水/周瘦鹃/还珠楼主/卧龙生/温瑞安/黄易 各 1。
    # 武侠三大家（金庸/古龙/梁羽生）与都德/霍桑/史蒂文森/康拉德/芥川龙之介/森鸥外/宫泽贤治/
    # 小泉八云/樋口一叶/平江不肖生/包天笑/诸葛青云/萧逸 实测 4 页内全为 0（武侠类定价高，2 元档没有），已剔除。
    ('公版作家', 4, [
        '爱默生', '爱伦坡', '吉卜林', '柯林斯', '梅里美',
        '张恨水', '周瘦鹃', '还珠楼主', '卧龙生', '温瑞安', '黄易',
    ]),

    # ===== AI / 科技 =====
    # 坦白说：≤2 元这一档 AI 书极少（实测 30 个 AI 词只捞到 1 本 + 英伟达 2 本），
    # 因为这类书多出版于近两年、定价普遍 40–80 元。这组只留 8 个核心词做兜底，
    # 不占太多预算；想把 AI 类看全，应该把 --cap 调到 10 元以上再跑。
    ('AI科技', 2, [
        'AI', '人工智能', '大模型', '英伟达',
    ]),
]

# 平台分类纠偏：微信读书自己的 category 有明显不合理（道德经→历史、论语→教育学习、
# 红楼梦→历史）。这里只处理**已验证明显不对**的组合，宁少不错：
#   (书名正则, 允许被改写的平台一级分类白名单, 新的一级分类)
# 白名单的意义：只有当平台把它归到这几类时才纠正，避免把正常分类也改坏。
CLASS_FIX = [
    (r'论语|孟子|荀子|大学|中庸|道德经|老子|庄子|诗经|楚辞|唐诗|宋词|元曲|古文观止|'
     r'菜根谭|世说新语|史记|资治通鉴|孙子兵法|三十六计|山海经|搜神记|战国策|左传|'
     r'弟子规|三字经|百家姓|千字文|声律启蒙|围炉夜话|浮生六记|纳兰|人间词话',
     ('历史', '教育学习', '社会文化', '哲学宗教', '文学', '国学'), '国学古籍'),
    (r'红楼梦|西游记|水浒传?|三国演义|聊斋|儒林外史|镜花缘|封神演义|东周列国志',
     ('历史', '教育学习', '童书', '文学'), '古典文学'),
    (r'^中华人民共和国|中华人民共和国.+法$', ('社会文化', '教育学习', '政治军事'), '法律法规'),
    (r'专刊|特刊|周刊|月刊|文摘|杂志|期刊|增刊|意林|半月谈|南风窗|故事会|知音',
     ('经济理财', '文学', '社会文化', '个人成长', '期刊杂志'), '期刊杂志'),
]


def fix_type(title, cat):
    """返回 (纠偏后的一级分类, 是否发生过纠偏)。"""
    p = (cat or '').split('-')[0].strip() if cat else ''
    t = title or ''
    for pat, allow, new in CLASS_FIX:
        if re.search(pat, t) and (not p or p in allow):
            return new, (p != new)
    return p, False


# 平台对一部分低价书不返回 category（实测 474 本里 328 本为 null），
# 补抓 /web/book/info 也拿不到。这里按书名/作者/出版社推出**近似类目**，
# 目标词表只用微信读书自己的一级分类名，推断结果在报告里标 ≈。
# 顺序即优先级：先命中先生效，所以把强特征（期号、法条、童话、佛经）放前面。
TYPE_RULES = [
    (r'专刊|特刊|周刊|月刊|文摘|杂志|期刊|增刊|合订本|画报|第\s*\d+\s*期', '期刊杂志'),
    (r'中华人民共和国|法典|条例|法律|法规|宪法|刑法|民法|合同法|劳动法|公司法|'
     r'婚姻法|仲裁|司法解释|律师', '法律法规'),
    (r'少年儿童|绘本|童话|寓言|儿童|宝宝|亲子|王国|勇士|历险|注音|小学生|一年级', '童书'),
    (r'(?<!哈)佛(?:教|学|经|法|典)|禅|菩提|般若|心经|金刚经|坛经|道教|易经|周易|宗教|哲学', '哲学宗教'),
    (r'论语|孟子|荀子|大学中庸|道德经|老子|庄子|诗经|楚辞|史记|资治通鉴|孙子|三十六计|'
     r'山海经|搜神记|战国策|左传|古文观止|千字文|弟子规|三字经|百家姓|声律启蒙|'
     r'家书|家训|诫子书|碑|答问|经史|集注|注疏|国学|古籍|文白对照|四书五经|名句', '国学古籍'),
    (r'红楼梦|西游记|水浒|三国演义|聊斋|儒林外史|镜花缘|封神演义|东周列国', '古典文学'),
    (r'经济|理财|投资|股票|基金|保险|会计|财务|管理|营销|商业|创业|金融|货币|'
     r'贸易|税务|审计|职场|领导力|供应链', '经济理财'),
    (r'成长|心理|情商|自律|习惯|思维|正能量|心灵|幸福|焦虑|沟通|口才|性格|自我|'
     r'励志|成功|效率|专注', '个人成长'),
    (r'医|药|病|养生|健康|中医|针灸|护理|临床|食疗|瑜伽|健身|减肥', '医学健康'),
    (r'科学|技术|物理|化学|生物|数学|工程|计算机|程序|算法|数据|宇宙|自然|'
     r'昆虫|植物|动物|天文|地理', '科学技术'),
    (r'教材|教辅|考试|真题|辅导|英语|单词|语法|作文|教程|讲义|习题|字典|词典', '教育学习'),
    (r'传[（(：:]|自传|传记|回忆录|书信|日记|生平|年谱|洛克菲勒|曾国藩|巴菲特', '人物传记'),
    (r'历史|中国史|世界史|通史|战争史|王朝|帝国', '历史'),
    (r'菜谱|食谱|烹饪|家常菜|家居|收纳|旅行|美容|穿搭|宠物|园艺|手工|咖啡|茶艺|'
     r'围棋|象棋|太极|生活', '生活'),
    (r'文化|社会|民俗|人类学|社会学|传播|新闻|伦理|考古|方略|政略', '社会文化'),
    (r'小说|散文|诗歌|诗选|文集|随笔|杂文|故事|作品集|读本|文学|爱情|野草|朝花夕拾', '文学'),
]

# 出版社只在域名特征明显时才用，避免乱猜
PUB_RULES = [
    (r'立信会计|财经|经济', '经济理财'),
    (r'少年儿童', '童书'),
    (r'卫生|医学|医药', '医学健康'),
    (r'法律', '法律法规'),
    (r'古籍|中华书局|岳麓书社|汇聚文源|辞书', '国学古籍'),
    (r'科学技术|科学普及|电子工业|机械工业', '科学技术'),
    (r'教育|人民教育', '教育学习'),
]


# 第一层规则没命中时的补充层：人物传记/赏析/生活/教育这类题名特征，
# 以及「作者即体裁」的情况（外国名著靠作者名比靠书名可靠）。
TYPE_RULES2 = [
    (r'传$|大传|全传|自述|评传|对话|访谈', '人物传记'),
    (r'心学|知行合一|精神分析|禅|修行', '哲学宗教'),
    (r'弗洛伊德|爱因斯坦|尼采|王阳明|张爱玲|苏东坡|曾国藩', '人物传记'),
    (r'帝王将相|王朝兴衰', '历史'),
    (r'鉴赏|赏析|诗词|诗选', '文学'),
    (r'卡耐基|卡内基', '个人成长'),
    (r'生意经|赚钱|百万富翁|好战略|经营|致富|创业', '经济理财'),
    (r'烘焙|厨|食谱|菜|养生|艾灸|推拿|按摩|经络', '医学健康'),
    (r'简史|时间简史|科普|寂静的春天', '科学技术'),
    (r'语文|教育漫话|课本|教材全解|教育家|敏感期|育儿|讲座', '教育学习'),
    (r'英文原版|英文版|双语|短篇', '文学'),
    (r'静心|修养|优雅|魅力|不抱怨|会说话|人脉|读心|潜能|心态|嫁给|心情|情绪', '个人成长'),
    (r'段永平|巴菲特|芒格|索罗斯|格雷厄姆', '经济理财'),
    (r'智囊|白话译本|奇书|鬼谷子|冰鉴|挺经|嘉言|曾文正', '国学古籍'),
    (r'书法|艺术|美丽中国|敦煌|民俗|文物', '社会文化'),
    (r'钢铁是怎样炼成的|鲁滨逊|柳林风声|长腿叔叔|爱的教育|名人名言|我这一辈子', '文学'),
]

AUTHOR_RULES = [
    (r'契诃夫|莫泊桑|凡尔纳|小仲马|大仲马|狄更斯|司汤达|托尔斯泰|川端康成|柯南|爱伦|'
     r'梅里美|勃朗特|奥斯汀|梭罗|卡夫卡|马克吐温|欧亨利|杰克伦敦|雨果|伏尔泰|格雷厄姆', '文学'),
    (r'鲁迅|老舍|胡适|张爱玲|沈从文|汪曾祺|丰子恺|冰心|梁启超|李叔同|弘一|闻一多', '文学'),
    (r'卡耐基|卡内基|希尔|阿德勒|富兰克林|科利尔', '个人成长'),
    (r'稻盛和夫|段永平|巴菲特|芒格|索罗斯', '经济理财'),
    (r'弗洛伊德|阿德勒|尼采|叔本华|荣格', '哲学宗教'),
    (r'释|法师|居士|上人', '哲学宗教'),
]


# 极少数书平台无分类、规则也咬不住：书名与作者里找不到任何可判定的词。
# 这类按 bookId 人工指定，优先于一切推断。销售/职场在微信读书里挂在「经济理财」下，
# 所以一级分类记经济理财，二级（销售/职场）留在说明里。
MANUAL_TYPE = {
    '189564': ('经济理财', '销售/职场题材（书名与作者无可判定词）'),
}


def infer_type1(title, author='', publisher='', kws=None, book_id=''):
    """平台没给分类时推一个近似一级分类；推不出返回 ('', None)。

    只用**书名**做正文匹配：搜索关键词是采样用的词，不是这本书的内容标签，
    掺进来会带偏（实测「心理学故事」会被关键词「哲学」拉到哲学宗教，
    「李鸿章传」也会因为关键词接在后面让 `传$` 失效）。
    """
    if book_id and book_id in MANUAL_TYPE:
        return MANUAL_TYPE[book_id][0], 'manual'
    head = title or ''
    for pat, cat in TYPE_RULES:
        if re.search(pat, head):
            return cat, 'title'
    for pat, cat in TYPE_RULES2:
        if re.search(pat, head):
            return cat, 'title'
    for pat, cat in AUTHOR_RULES:
        if author and re.search(pat, author):
            return cat, 'author'
    for pat, cat in PUB_RULES:
        if publisher and re.search(pat, publisher):
            return cat, 'publisher'
    return '', None


def all_kw_pages(scale=1.0):
    """展开成 [(关键词, 页数)]，scale 可整体放大页数。"""
    out = []
    for _name, pages, words in GROUPS:
        n = max(1, int(round(pages * scale)))
        for w in words:
            out.append((w, n))
    return out


def scan(ctx, cap, min_rating, items, delay):
    """按 (关键词, 页数) 逐个扫，只留「有价、未下架、评分达标、≤cap」的书。"""
    found, scanned, req = {}, 0, 0
    total = len(items)
    for i, (kw, pages) in enumerate(items, 1):
        got_here = 0
        for p in range(pages):
            url = '%s/web/search/name_books?keyword=%s&maxIdx=%d&count=%d' % (
                BASE, quote(kw), p * PAGE, PAGE)
            req += 1
            try:
                j = ctx.request.get(url, timeout=25000).json()
            except Exception as e:
                print('  [fail] %-10s p%d %s' % (kw, p + 1, str(e)[:60]))
                break
            books = j.get('books') or []
            scanned += len(books)
            for it in books:
                bi = it.get('bookInfo') or {}
                bid = bi.get('bookId')
                price = bi.get('price')
                nr = bi.get('newRating')
                if not bid or not isinstance(price, (int, float)):
                    continue
                if price <= 0 or price > cap + 1e-9:      # 0 元是免费书、-1 是异常值
                    continue
                if bi.get('soldout') in (1, True):        # 已下架买不到
                    continue
                if not isinstance(nr, (int, float)) or nr / 10.0 < min_rating:
                    continue
                rec = found.get(bid)
                if rec:
                    if kw not in rec['kw']:
                        rec['kw'].append(kw)
                else:
                    found[bid] = {
                        'bookId': bid,
                        'title': bi.get('title') or '',
                        'author': bi.get('author') or '',
                        'publisher': bi.get('publisher') or '',
                        'price': round(float(price), 2),
                        'rating': round(nr / 10.0, 1),
                        'rcnt': bi.get('newRatingCount') or 0,
                        'deepLink': bi.get('deepLink') or '',
                        'cover': bi.get('cover') or '',
                        'kw': [kw],
                    }
                    got_here += 1
            if not j.get('hasMore'):
                break
            time.sleep(delay)
        if i % 10 == 0 or got_here >= 3:
            print('  %3d/%d  %-10s(%dp) +%d　累计 %d 本（已扫 %d 条 / %d 请求）'
                  % (i, total, kw, pages, got_here, len(found), scanned, req))
    return found, scanned, req


def enrich(ctx, found, delay):
    """补分类：搜索结果里 category 为 null，用 /web/book/info 逐本补，并做分类纠偏。"""
    todo = list(found.values())
    print('补分类：%d 本' % len(todo))
    fixed_n = inferred_n = 0
    for i, b in enumerate(todo, 1):
        try:
            j = ctx.request.get(BASE + '/web/book/info?bookId=' + b['bookId'], timeout=20000).json()
        except Exception:
            continue
        b['category'] = j.get('category') or ''
        t1, changed = fix_type(b['title'], b['category'])
        b['platformType1'] = (b['category'].split('-')[0].strip() if b.get('category') else '')
        if t1:
            b['_type_src'] = 'fixed' if changed else 'platform'
        else:
            t1, why = infer_type1(b['title'], b.get('author') or '',
                                  b.get('publisher') or '', b.get('kw'),
                                  book_id=b.get('bookId'))
            b['_type_src'] = ('manual' if why == 'manual' else 'inferred') if t1 else ''
            b['_type_infer_by'] = why or ''
            if t1:
                inferred_n += 1
        b['_type1'] = t1
        b['_type_fixed'] = changed
        if changed:
            fixed_n += 1
        if j.get('centPrice') is not None:
            b['price'] = round(j['centPrice'] / 100.0, 2)
        if j.get('newRating'):
            b['rating'] = round(j['newRating'] / 10.0, 1)
        if j.get('newRatingCount') is not None:
            b['rcnt'] = j['newRatingCount']
        if j.get('deepLink'):
            b['deepLink'] = j['deepLink']
        if i % 25 == 0:
            print('  %d/%d（已纠偏 %d，推断补类 %d）' % (i, len(todo), fixed_n, inferred_n))
        time.sleep(delay)
    print('  分类纠偏：%d 本按规则改写；平台无分类、按书名/作者/出版社推断：%d 本'
          % (fixed_n, inferred_n))
    return found


def reclassify(path):
    """只重算种子分类，不重抓：补空分类 + 重跑纠偏，并给出推断覆盖率。"""
    d = json.load(open(path, encoding='utf-8'))
    books = d['books']
    was_empty = 0
    inferred = {}
    for b in books:
        b.setdefault('author', '')
        b.setdefault('publisher', '')
        b.setdefault('kw', [])
        b.setdefault('category', '')
        if not (b.get('_type1') or '').strip():
            was_empty += 1
        t1, changed = fix_type(b['title'], b.get('category') or '')
        if t1:
            b['_type1'] = t1
            b['_type_src'] = 'fixed' if changed else 'platform'
            b['_type_fixed'] = changed
        else:
            t1, why = infer_type1(b['title'], b.get('author') or '',
                                  b.get('publisher') or '', b.get('kw'),
                                  book_id=b.get('bookId'))
            b['_type1'] = t1
            b['_type_src'] = ('manual' if why == 'manual' else 'inferred') if t1 else ''
            b['_type_infer_by'] = why or ''
            b['_type_fixed'] = False
            if t1:
                inferred[why] = inferred.get(why, 0) + 1
    meta = d.setdefault('meta', {})
    meta['typeInferred'] = sum(inferred.values())
    meta['typeInferredBy'] = inferred
    meta['typeEmptyBefore'] = was_empty
    meta['reclassifiedAt'] = time.strftime('%Y-%m-%d %H:%M')
    json.dump(d, open(path, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    left = [b for b in books if not (b.get('_type1') or '').strip()]
    print('种子 %d 本：原有空分类 %d 本' % (len(books), was_empty))
    print('  推断补类 %d 本（书名/作者命中 %d、出版社命中 %d）；仍无法归类 %d 本'
          % (sum(inferred.values()), inferred.get('title', 0),
             inferred.get('publisher', 0), len(left)))
    print('  分类分布：')
    import collections as _c
    for k, v in _c.Counter(b['_type1'] or '(空)' for b in books).most_common():
        print('    %-8s %d' % (k, v))
    if left:
        print('  未归类示例：')
        for b in left[:10]:
            print('    %-30s pub=%s' % (b['title'][:30], (b.get('publisher') or '-')[:16]))
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cap', type=float, default=2.0, help='价格上限（元），默认 2')
    ap.add_argument('--min-rating', type=float, default=70.0, help='微信评分下限，默认 70')
    ap.add_argument('--pages-scale', type=float, default=1.0,
                    help='页数统一放大倍数（默认 1，即按 GROUPS 各自页数）')
    ap.add_argument('--pages', type=int, default=0, help='统一指定每个词翻几页（覆盖分组设置）')
    ap.add_argument('--delay', type=float, default=0.35, help='请求间隔秒，默认 0.35')
    ap.add_argument('--enrich', action='store_true', help='逐本补分类 + 纠偏（默认开启）')
    ap.add_argument('--no-enrich', action='store_true', help='跳过补分类')
    ap.add_argument('--keywords', default='', help='自定义词表，逗号分隔（默认用 GROUPS）')
    ap.add_argument('--out', default=DEFAULT_OUT, help='输出路径')
    ap.add_argument('--reclassify', action='store_true',
                    help='不抓取，只对 --out 指向的种子重算分类（补空分类 + 重跑纠偏）')
    a = ap.parse_args()

    if a.reclassify:
        reclassify(a.out)
        return

    if a.keywords:
        words = [x.strip() for x in a.keywords.split(',') if x.strip()]
        items = [(w, a.pages or 2) for w in words]
    else:
        items = all_kw_pages(a.pages_scale)
        if a.pages:
            items = [(w, a.pages) for w, _ in items]

    page_budget = sum(p for _, p in items)
    print('关键词 %d 个 · 计划请求 %d 次（每页 40 条，理论上限 %d 条）'
          % (len(items), page_budget, page_budget * PAGE))
    print('目标：≤%.2f 元 且 微信评分 ≥%.1f' % (a.cap, a.min_rating))
    print('分组页数：' + '，'.join('%s %d 页' % (n, p) for n, p, _ in GROUPS))

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            wconfig.PROFILE_DIR, **wconfig.launch_kwargs(headless=True))
        if not any(c['name'] == 'wr_vid' for c in ctx.cookies()):
            print('LOGIN_LOST：登录态失效，先跑 python scripts/login.py')
            ctx.close()
            return 2

        t0 = time.time()
        found, scanned, req = scan(ctx, a.cap, a.min_rating, items, a.delay)
        t_scan = time.time() - t0
        print('\n扫描阶段完成：%d 请求 / %d 条 / 命中 %d 本 / 用时 %.0f 秒'
              % (req, scanned, len(found), t_scan))

        if found and not a.no_enrich:
            t1 = time.time()
            found = enrich(ctx, found, min(a.delay, 0.15))
            print('补分类用时 %.0f 秒' % (time.time() - t1))
        ctx.close()

    books = sorted(found.values(), key=lambda b: (-b['rating'], b['price'], b['title']))
    for b in books:
        b.setdefault('category', '')
        b.setdefault('_type1', '')
        b.setdefault('_type_fixed', False)
        b.setdefault('platformType1', '')
    elapsed = round(time.time() - t0, 1)
    out = {
        'meta': {
            'generatedAt': time.strftime('%Y-%m-%d %H:%M'),
            'cap': a.cap,
            'minRating': a.min_rating,
            'keywordCount': len(items),
            # 存下本次实际用过的词表快照，analyze_keywords.py 靠它判断「零贡献词」，
            # 否则词表改过后会把新加的词误判成白费请求。
            'keywordsUsed': [w for w, _ in items],
            'plannedPages': page_budget,
            'requests': req,
            'scanned': scanned,
            'matched': len(books),
            'pagesByGroup': {n: p for n, p, _ in GROUPS},
            'typeFixed': sum(1 for b in books if b.get('_type_fixed')),
            'enriched': bool(found and not a.no_enrich),
            'source': 'WeRead /web/search/name_books 关键词采样（非全库穷举）',
            'elapsedSec': elapsed,
            'elapsedMin': round(elapsed / 60.0, 1),
        },
        'books': books,
    }
    d = os.path.dirname(os.path.abspath(a.out))
    if d:
        os.makedirs(d, exist_ok=True)
    with open(a.out, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    print('\n=== 完成 ===')
    print('扫描 %d 条 → 命中 %d 本（≤%.2f 元且评分 ≥%.1f）'
          % (scanned, len(books), a.cap, a.min_rating))
    print('总请求 %d 次 · 总用时 %.1f 分钟（%.0f 秒）' % (req, elapsed / 60.0, elapsed))
    print('已写入', os.path.abspath(a.out))
    if books:
        print('\n评分最高的 12 本：')
        for b in books[:12]:
            print('  %5.2f 元  %4.1f 分  %-34s %s'
                  % (b['price'], b['rating'], (b['title'] or '')[:34],
                     ('[%s]' % b['_type1']) if b.get('_type1') else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main())
