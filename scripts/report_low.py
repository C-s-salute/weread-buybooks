# -*- coding: utf-8 -*-
"""零钱模式的两张 sheet：零钱购书方案 / 零钱备选。

这里只负责排版。数据、阈值与文案口径由 build_excel 通过 configure() 注入 ——
两张表用到的模块级名字都在这儿列清，避免两个模块在 import 期互相依赖。
"""
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# 由 build_excel.configure() 用同名变量覆盖
COIN = LOW_CAP = LOW_TRIGGER = LOW_MAX_BOOKS = 0
LOW_POOL = LOW_SEED = LOW_POOL_BAG = LOW_AFFORD = LOW_SEED_META = None
LOW_CHEAPEST = LOW_PLANS = LOW_LATER = SEED_INFERRED = BY_ID = None
BORDER = HEAD_FILL = HEAD_FONT = TITLE_FONT = CENTER = MONEY = None


def configure(**kw):
    """注入 build_excel 侧的数据与样式常量。"""
    globals().update(kw)


def build_lowplan_sheet(wb):
    """零钱购书方案：余额 ≤ LOW_TRIGGER 时，直接给出「这点钱买哪几本」。

    候选池 = 全库 ≤LOW_CAP 元且评分达标的书（data/lowprice_seed.json），
    与用户书架无关 —— 因为余额一两元时，书架里往往根本没有可买的书。
    """
    ws = wb.create_sheet('零钱购书方案')
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([5, 34, 18, 11, 9, 9, 10, 22, 11], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    NC = 9
    r = 1

    def _band(text, color='1F3864', size=11, h=26):
        nonlocal r
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(bold=True, size=size, color=color, name='微软雅黑')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        ws.row_dimensions[r].height = h
        r += 1

    def _note(text, color='595959', h=26, size=9):
        nonlocal r
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(size=size, color=color, name='微软雅黑')
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        ws.row_dimensions[r].height = h
        r += 1

    def _head(cols):
        nonlocal r
        for i, h in enumerate(cols, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.fill = HEAD_FILL
            c.font = HEAD_FONT
            c.alignment = CENTER
            c.border = BORDER
        ws.row_dimensions[r].height = 24
        r += 1

    c = ws.cell(row=r, column=1,
                value='零钱购书方案：余额 %.2f 元怎么花（全库 ≤%.2f 元清单里挑）'
                      % (COIN, LOW_CAP))
    c.font = TITLE_FONT
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    ws.row_dimensions[r].height = 30
    r += 1
    _note('触发条件：可用余额 ≤ %.2f 元（阈值可用环境变量 WEREAD_LOW_TRIGGER 调整）。'
          '候选池来自微信读书全库「≤%.2f 元且微信评分 ≥70」的 %d 本书，'
          '**不是你书架里的书** —— 余额这么少时，书架里通常没有可买的。'
          '价值函数用微信评分（全库书抓不到阅读人数等数据，无法算四层价值分）。'
          % (LOW_TRIGGER, LOW_CAP, len(LOW_SEED)), h=44)

    if not LOW_PLANS:
        _note('没有可用的全库清单（data/lowprice_seed.json 不存在或为空），'
              '先跑 python scripts/fetch_lowprice.py 生成。', 'C00000', 30)
        ws.freeze_panes = 'A4'
        return ws

    # 一句话结论
    best = LOW_PLANS[0]
    if best['plan']:
        _note('结论：按「%s %s」，这 %.2f 元能买 **%d 本**、花 %.2f 元、剩 %.2f 元，'
              '评分合计 %.1f。三个方案选一个执行即可，不要叠加。'
              % (best['no'], best['name'], COIN, best['stat']['n'], best['stat']['spent'],
                 best['stat']['left'], best['stat']['tv']), 'C00000', 30, 10)
    else:
        _note('余额 %.2f 元低于清单里最便宜的一本（%.2f 元），当前一本都买不了。'
              % (COIN, min((x['_price'] for x in LOW_POOL_BAG), default=0)), 'C00000', 26, 10)

    GREEN = PatternFill('solid', fgColor='E2F0D9')
    for pl in LOW_PLANS:
        r += 1
        st = pl['stat']
        _band('%s %s —— %d 本 / 花 %.2f 元 / 剩 %.2f 元 / 评分合计 %.1f（均 %.1f）'
              % (pl['no'], pl['name'], st['n'], st['spent'], st['left'], st['tv'], st['av']))
        _note(pl['desc'], h=22)
        if not pl['plan']:
            _note('（本方案在 %.2f 元预算内无解）' % COIN, '7F7F7F', 22)
            continue
        _head(['#', '书名', '作者', '一级分类', '标价(元)', '微信评分',
               '评分人数', '出版社', '购买'])
        for i, b in enumerate(pl['plan'], 1):
            vals = [i, b['title'], b['author'], b['_type1'], b['_price'],
                    b['_rating'], b['_rcnt'], b['publisher'], '去微信读书']
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.border = BORDER
                c.font = Font(size=10, name='微软雅黑', bold=(j == 6))
                c.alignment = Alignment(horizontal='left' if j in (2, 3, 8) else 'center',
                                        vertical='center', wrap_text=(j in (2, 3, 8)))
                if j == 5:
                    c.number_format = MONEY
                if j == 6:
                    c.fill = GREEN
                if j == 9 and b.get('deepLink'):
                    c.hyperlink = b['deepLink']
                    c.font = Font(size=10, name='微软雅黑', color='0563C1', underline='single')
            ws.row_dimensions[r].height = 22
            r += 1
        _note('小计：%d 本 / %.2f 元 / 剩余 %.2f 元（评分合计 %.1f，均 %.1f）'
              % (st['n'], st['spent'], st['left'], st['tv'], st['av']), h=22)

    # 再攒 20 书币
    r += 1
    if LOW_LATER:
        ls = LOW_LATER['stat']
        _hd = ('④ 如果完成本周阅读任务（+20 书币）—— 能多买 %d 本' % LOW_LATER['addN']
               if LOW_LATER.get('addN') else
               '④ 如果完成本周阅读任务（+20 书币）—— 本数不变，但能换成更好的书')
        _band(_hd)
        _note('按余额 %.2f 元重算「精选 %d 本」：%d 本 / 花 %.2f 元 / 剩 %.2f 元 / 评分合计 %.1f'
              '（比现在 +%.1f 分）。书币来源与领取规则以 App 内活动页为准。'
              % (COIN + 20, LOW_MAX_BOOKS, ls['n'], ls['spent'], ls['left'], ls['tv'],
                 LOW_LATER.get('gain') or 0), h=26)
        for i, b in enumerate(LOW_LATER['plan'], 1):
            c = ws.cell(row=r, column=1, value='%d. %s（%.2f 元，%s 分）'
                                              % (i, b['title'], b['_price'], b['_rating']))
            c.font = Font(size=10, name='微软雅黑')
            c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
            ws.row_dimensions[r].height = 20
            r += 1
    else:
        _note('④ 再攒 20 书币：按当前清单重算，本数和评分都不会提升（%d 元内已到上限），'
              '不必为凑单而等。想要更多可选书，可把 WEREAD_LOW_MAXBOOKS 调大。' % (COIN + 20), h=26)

    r += 1
    _note('※ 三个方案互斥，选一个执行。价格与评分取自搜索接口，与书籍详情页一致（已抽验）；'
          '价格会随平台调整变化，清单是 %s 的快照。'
          % (LOW_SEED_META.get('generatedAt') or '上次抓取'))
    ws.freeze_panes = 'A4'
    return ws


def build_lowprice_sheet(wb):
    """零钱备选：≤ LOW_CAP 元的高价值书，分两个来源。

    区 A「你的书架/书单」：能直接下单，最相关；
    区 B「全库采样」：与书架无关，书架里一本低价书都没有时兜底。
    区 B 来自 data/lowprice_seed.json（fetch_lowprice.py 生成）。
    """
    ws = wb.create_sheet('零钱备选(≤%g元)' % LOW_CAP)
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([5, 34, 18, 11, 9, 9, 10, 10, 24, 12, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    NC = 11
    r = 1

    def _band(text, color='1F3864', size=11):
        nonlocal r
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(bold=True, size=size, color=color, name='微软雅黑')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        ws.row_dimensions[r].height = 26
        r += 1

    def _note(text, color='595959', h=26):
        nonlocal r
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(size=9, color=color, name='微软雅黑')
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        ws.row_dimensions[r].height = h
        r += 1

    def _head(cols):
        nonlocal r
        for i, h in enumerate(cols, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.fill = HEAD_FILL
            c.font = HEAD_FONT
            c.alignment = CENTER
            c.border = BORDER
        ws.row_dimensions[r].height = 24
        r += 1

    GREEN = PatternFill('solid', fgColor='E2F0D9')
    GREY = '7F7F7F'

    c = ws.cell(row=r, column=1,
                value='零钱备选：标价 ≤ %.2f 元的高价值书（余额不足时"还能买什么"）' % LOW_CAP)
    c.font = TITLE_FONT
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    ws.row_dimensions[r].height = 30
    r += 1
    _note('分两个来源：① 你自己的书架/书单里 ≤%.2f 元的书（能直接下单，最相关）；'
          '② 微信读书全库的 ≤%.2f 元高分书（与你书架无关，由 fetch_lowprice.py 关键词采样生成，'
          '不是全库穷举）。当前余额 %.2f 元。' % (LOW_CAP, LOW_CAP, COIN), h=34)

    # ---------- 区 A：书架/书单 ----------
    r += 1
    _band('① 你的书架 / 书单：标价 ≤ %.2f 元、未下架、有评分的 %d 本' % (LOW_CAP, len(LOW_POOL)))
    if LOW_POOL:
        _head(['#', '书名', '作者', '一级分类', '标价(元)', '价值评分', '微信评分',
               '评分人数', '当前余额', '阅读进度', '下架概率'])
        for i, b in enumerate(LOW_POOL, 1):
            afford = b['_price'] <= COIN + 1e-9
            vals = [i, b.get('title') or '', b.get('author') or '', b.get('_type1') or '',
                    b['_price'], round(b['_vscore'], 1), b.get('_rating') or 0,
                    b.get('_rcnt') or 0,
                    ('现在能买' if afford else '还差 %.2f 元' % (b['_price'] - COIN)),
                    '%d%%' % (b.get('_prog') or 0), round(b.get('_risk') or 0, 1)]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.border = BORDER
                c.font = Font(size=10, name='微软雅黑', bold=(j == 6),
                              color='1F3864' if afford else GREY)
                c.alignment = Alignment(horizontal='left' if j in (2, 3) else 'center',
                                        vertical='center', wrap_text=(j in (2, 3)))
                if j == 5:
                    c.number_format = MONEY
                if afford:
                    c.fill = GREEN
            ws.row_dimensions[r].height = 22
            r += 1
        if LOW_AFFORD:
            _note('你现在的余额 %.2f 元能买其中 %d 本。' % (COIN, len(LOW_AFFORD)))
        else:
            _note('余额 %.2f 元连最便宜的一本（%.2f 元）都还差 %.2f 元。'
                  % (COIN, LOW_CHEAPEST['_price'], LOW_CHEAPEST['_price'] - COIN), 'C00000')
    else:
        _note('你的书架和书单里没有满足条件的书（≤%.2f 元、未下架、有评分）—— '
              '这不是数据出错，是这类书本来就少。请看下面第 ② 区。' % LOW_CAP, 'C00000')

    # ---------- 区 B：全库采样 ----------
    r += 2
    _mr = LOW_SEED_META.get('minRating')
    _mr = ('%.0f' % _mr) if isinstance(_mr, (int, float)) else '70'
    _band('② 微信读书全库：≤ %.2f 元、微信评分 ≥ %s 的 %d 本（与你书架无关，可当备选）'
          % (LOW_CAP, _mr, len(LOW_SEED)))
    if LOW_SEED:
        _head(['#', '书名', '作者', '一级分类', '标价(元)', '微信评分', '评分人数',
               '出版社', '与你书架', '命中关键词', '购买'])
        for i, b in enumerate(LOW_SEED, 1):
            owned = BY_ID.get(b['bookId'])
            if owned and owned['_cat'] == '永久可读':
                rel, col = '已拥有', GREY
            elif owned:
                rel, col = '在书架未买断', 'C00000'
            else:
                rel, col = '未加书架', '1F3864'
            vals = [i, b.get('title') or '', b.get('author') or '', type_disp(b),
                    float(b.get('price') or 0), b.get('rating') or 0, b.get('rcnt') or 0,
                    b.get('publisher') or '', rel, '、'.join(b.get('kw') or [])[:24], '去微信读书']
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.border = BORDER
                c.font = Font(size=10, name='微软雅黑', bold=(j == 6),
                              color=col if j == 9 else '000000')
                c.alignment = Alignment(horizontal='left' if j in (2, 3, 8, 10) else 'center',
                                        vertical='center', wrap_text=(j in (2, 3, 8)))
                if j == 5:
                    c.number_format = MONEY
                if j == 11 and b.get('deepLink'):
                    c.hyperlink = b['deepLink']
                    c.font = Font(size=10, name='微软雅黑', color='0563C1', underline='single')
            ws.row_dimensions[r].height = 22
            r += 1
        _pg = LOW_SEED_META.get('pagesByGroup') or {}
        _pg_txt = '、'.join('%s %s 页' % (k, v) for k, v in _pg.items()) if _pg else '-'
        _rounds = LOW_SEED_META.get('rounds') or 1
        _round_txt = ('（分 %d 轮抓取后合并去重）' % _rounds) if _rounds > 1 else ''
        _note('区 ② 的口径：全库关键词采样 —— %d 个关键词、本轮 %s 次请求 / %s 分钟（含补分类）、'
              '命中 %d 本%s；累计 %s 请求、扫描 %s 条。生成于 %s。分组页数：%s。'
              '分类已按规则纠正 %d 本（平台原分类明显不当的，例如《道德经》被归到「历史」、《论语》被归到「教育学习」），'
              '其余沿用微信读书自己的分类。'
              '其中 %d 本平台没有返回分类，按书名/作者/出版社归入最接近的一级分类，分类名前的 ≈ 就是这个意思。'
              '注意：微信读书没有「按价格浏览全库」的接口，所以这是采样不是穷举；'
              '要扩大覆盖，改 fetch_lowprice.py 顶部的 GROUPS 后重跑。'
              % (LOW_SEED_META.get('keywordCount') or 0,
                 LOW_SEED_META.get('lastRunRequests') or LOW_SEED_META.get('requests') or '-',
                 LOW_SEED_META.get('lastRunMinutes') or LOW_SEED_META.get('elapsedMin') or '-',
                 LOW_SEED_META.get('matched') or len(LOW_SEED),
                 _round_txt,
                 LOW_SEED_META.get('requests') or '-',
                 LOW_SEED_META.get('scanned') or '-',
                 LOW_SEED_META.get('generatedAt') or '未知',
                 _pg_txt,
                 LOW_SEED_META.get('typeFixed') or 0,
                 SEED_INFERRED), h=58)
    else:
        _note('没有全库清单文件（data/lowprice_seed.json）。'
              '跑一次 python scripts/fetch_lowprice.py 就会生成，之后本区自动填充。', 'C00000')

    r += 1
    _note('※ 两区都按评分降序；区 ① 的「当前余额」列直接告诉你买不买得起，绿色行 = 现在能买。'
          '区 ② 的价格与评分取自搜索接口，已抽验与书籍详情页一致。')
    ws.freeze_panes = 'A4'
    return ws
