# -*- coding: utf-8 -*-
"""生成「购书方案」HTML 报告（自包含单文件，无外链依赖）

数据全部 import 自 build_excel.py，保证与 Excel / Markdown 口径完全一致：
  · 余额 COIN、候选池 COIN_POOL_X、四方案 RATIO_PLAN / MAIN_PLAN / VAL_PICK / RISK_PLAN
  · 方案③ 的分类口径 TYPE1_STD（= 「分类基准与类型评分」同一套 15 个一级分类）
输出：微信读书_购书方案.html
"""
import os
import sys
import html

try:
    import wconfig   # noqa: F401  统一工作目录（设置了 WEREAD_WORK 时切 cwd）
except ImportError:
    pass

# 同目录的 build_excel 直接取，不依赖 cwd 或 PYTHONPATH
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_excel as M

COIN = M.COIN
OUT = '微信读书_购书方案.html'


def E(s):
    return html.escape('' if s is None else str(s))


def bl(b, wrap=False):
    """书名 → 微信读书详情页（新标签打开）；无 deepLink 时退化为纯文本。"""
    t = E(b.get('title'))
    u = (b.get('deepLink') or '').strip()
    label = '《%s》' % t if wrap else t
    if not u:
        return '<span class="bt">%s</span>' % label
    return ('<a class="bt" href="%s" target="_blank" rel="noopener noreferrer" '
            'title="在微信读书打开这本书">%s</a>' % (E(u), label))


def m2(x):
    return '%.2f' % (x or 0)


def m1(x):
    return '%.1f' % (x or 0)


def pct1(x):
    """下架风险本身即「概率 ×100」，一律按百分比展示。"""
    return '%.1f%%' % (x or 0)


def per_of(b):
    return (b['_vscore'] / b['_price']) if (b['_vscore'] and b['_price']) else None


def risk_cls(r):
    return 'r' if (r or 0) >= 30 else ('y' if (r or 0) >= 20 else 'g')


CSS = r"""
:root{
  --ink:#231F19; --ink2:#5B5348; --ink3:#948B7D;
  --line:#E7DFD1; --line2:#F1EBE0;
  --paper:#FAF6EE; --card:#ffffff;
  --blue:#1E4638; --blue2:#2E6B52; --blueSoft:#ECF3ED;
  --red:#B03A2B; --redSoft:#FAECE8;
  --green:#3E6B4A; --gold:#A67C3D; --goldSoft:#FAF0DC;
  --pink:#FF4593; --pinkSoft:#FFF0F6;
  --grass:#40C833; --grassSoft:#F0FBEC;
  --cyan:#39CFFF; --cyanSoft:#EEFAFF;
  --goldDeep:#B8860B;
  --goldBright:#E3B017; --goldSoftB:#FEF9E7;
  --silver:#A9B4C0; --silverInk:#64717F; --silverSoft:#F5F7FA;
  --blueTag:#1B6FC4; --blueTagSoft:#EDF4FD;
  --kai:"汉仪行楷简","HYXingKaiJ","HYXingKaiS","华文行楷","STXingkai","STKaiti","KaiTi","楷体",cursive;
  --shadow:0 1px 2px rgba(30,70,56,.05),0 8px 24px -14px rgba(30,70,56,.16);
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{
  margin:0;background:var(--paper);color:var(--ink);font-size:15px;line-height:1.65;
  font-family:"PingFang SC","Microsoft YaHei","Noto Sans SC",system-ui,-apple-system,sans-serif;
  -webkit-font-smoothing:antialiased;
}
.wrap{max-width:1180px;margin:0 auto;padding:0 28px}

/* ---------- hero ---------- */
.hero{
  padding:62px 0 38px;border-bottom:1px solid var(--line);
  background:radial-gradient(115% 120% at 88% -10%,#EDF4EC 0%,#FAF6EE 58%);
}
.kicker{font-size:12.5px;letter-spacing:.2em;color:var(--blue2);font-weight:600;margin:0}
.hero h1{
  font-size:clamp(34px,4.6vw,52px);line-height:1.12;margin:16px 0 0;
  letter-spacing:-.025em;font-weight:700;
}
.hero h1 em{font-style:normal;color:var(--blue2)}
.lede{max-width:640px;color:var(--ink2);font-size:16px;margin:20px 0 0}
.lede b{color:var(--ink)}

.hero-grid{display:grid;grid-template-columns:1fr 292px;gap:46px;align-items:start}
.balcard{
  background:linear-gradient(165deg,#FFFDF7 0%,#FCF4DF 58%,#F7EAC5 100%);
  border:1px solid #EBD9AE;border-radius:16px;padding:20px 22px;
  box-shadow:0 1px 2px rgba(166,124,61,.1),0 12px 28px -18px rgba(166,124,61,.45);
}
.balcard .k{font-size:11px;letter-spacing:.16em;color:#A8813A;font-weight:700}
.balcard .big{
  font-size:42px;line-height:1.1;letter-spacing:-.04em;font-variant-numeric:tabular-nums;
  margin:6px 0 0;font-weight:700;color:var(--goldDeep);text-shadow:0 1px 0 rgba(255,255,255,.8);
}
.balcard .big span{font-size:15px;font-weight:500;color:var(--ink3);margin-left:4px;letter-spacing:0}
.balcard ul{list-style:none;margin:16px 0 0;padding:14px 0 0;border-top:1px dashed #E4D3A8}
.balcard li{display:flex;justify-content:space-between;align-items:baseline;font-size:12.5px;padding:4px 0;color:var(--ink3)}
.balcard li b{color:var(--ink);font-variant-numeric:tabular-nums;font-size:13px}
.balcard li.warn b{color:var(--red)}

.alert{
  margin:26px 0 0;background:var(--redSoft);border:1px solid #F2D2CB;border-left:4px solid var(--red);
  border-radius:12px;padding:14px 18px;font-size:13.5px;color:#7A3A2E;line-height:1.7;
}
.alert b{color:var(--red);font-weight:700}
.alert .t{display:block;font-size:11px;letter-spacing:.14em;font-weight:700;color:var(--red);margin-bottom:3px}

.funnel{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:36px 0 0}
.fstep{
  border:1px solid #E0EDE4;border-radius:12px;padding:14px 14px 12px;position:relative;
  background:#F2F9F4;transition:box-shadow .18s;
}
.fstep:not(:last-child):after{
  content:"›";position:absolute;right:-10px;top:50%;transform:translateY(-50%);
  font-size:17px;color:#8DB29A;line-height:1;z-index:2;
}
.fstep b{display:block;font-size:23px;letter-spacing:-.02em;font-variant-numeric:tabular-nums;line-height:1.2;color:#4A7A5D}
.fstep span{display:block;font-size:11.5px;color:#44634F;margin-top:4px;line-height:1.4}

/* 漏斗五级：绿色由浅到深，逐级收窄；第 3、4 级是「减法」，用淡红边框标出 */
.fstep:nth-child(1){background:#F2F9F4;border-color:#E0EDE4}
.fstep:nth-child(1) b{color:#4A7A5D}
.fstep:nth-child(2){background:#E4F1E8;border-color:#CDE3D3}
.fstep:nth-child(2) b{color:#406F53}
.fstep:nth-child(3){background:#D6EBDB;border-color:#EBBBAE}
.fstep:nth-child(3) b{color:#CC0000}
.fstep:nth-child(4){background:#C6E2CD;border-color:#DE9F8D}
.fstep:nth-child(4) b{color:#CC0000}
.fstep:nth-child(3) span,.fstep:nth-child(4) span{color:#3C5A48}

/* 收口：微信绿 */
.fstep.final{
  background:linear-gradient(132deg,#0DC267 0%,#05AE55 48%,#039A49 100%);
  border-color:#04A44F;color:#fff;
  box-shadow:0 10px 24px -16px rgba(6,190,96,.55);
}
.fstep.final b,.fstep.final span{color:#fff;text-shadow:0 1px 3px rgba(0,64,32,.42)}

.flags{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:24px}
.flag{background:var(--card);border:1px solid var(--line);border-left:3px solid var(--line);border-radius:10px;padding:13px 15px;font-size:13px;color:var(--ink2)}
.flag b{display:block;color:var(--ink);font-size:13px;margin-bottom:3px;font-weight:600}
.flag.r{border-left-color:var(--red)}
.flag.b{border-left-color:var(--blue2)}
.flag.g{border-left-color:var(--green)}

/* ---------- nav ---------- */
#nav{position:sticky;top:0;z-index:50;background:rgba(250,246,238,.9);backdrop-filter:blur(12px);border-bottom:1px solid var(--line)}
.navwrap{display:flex;align-items:center;gap:20px;height:55px;overflow-x:auto;scrollbar-width:none}
.navwrap::-webkit-scrollbar{display:none}
.brand{font-size:13px;font-weight:700;letter-spacing:.02em;color:var(--blue);white-space:nowrap;padding-right:6px;border-right:1px solid var(--line);margin-right:2px}
#nav a{font-size:13px;color:var(--ink2);text-decoration:none;white-space:nowrap;padding:3px 0;border-bottom:2px solid transparent;transition:.15s}
#nav a:hover{color:var(--blue)}
#nav a.on{color:var(--blue);border-color:var(--blue2);font-weight:600}
.navbal{margin-left:auto;font-size:12.5px;color:var(--ink3);white-space:nowrap}
.navbal b{color:var(--goldDeep);font-size:14px;font-variant-numeric:tabular-nums}

/* ---------- section ---------- */
section{padding:58px 0 12px;border-bottom:1px solid var(--line2);scroll-margin-top:64px}
.sec-h{display:flex;align-items:baseline;gap:12px;margin:0 0 6px}
.sec-h .no{font-size:11.5px;letter-spacing:.16em;color:var(--ink3);font-weight:700}
.sec-h h2{font-size:26px;margin:0;letter-spacing:-.01em;font-weight:700}
.sec-sub{color:var(--ink2);font-size:14px;margin:0 0 26px;max-width:820px}
.sec-sub b{color:var(--ink)}

/* ---------- plan cards ---------- */
.pcards{display:grid;grid-template-columns:repeat(4,1fr);gap:15px}
.pcard{
  background:var(--card);border:1px solid var(--line);border-top:3px solid var(--line);
  border-radius:15px;padding:20px 18px;box-shadow:var(--shadow);cursor:pointer;
  transition:transform .18s,box-shadow .18s,border-color .18s;position:relative;display:flex;flex-direction:column;
}
.pcard:hover{transform:translateY(-3px);box-shadow:0 12px 30px -14px rgba(30,70,56,.28);border-color:#D2E2D6}
.pcard{border-top-width:4px}
.pcard.t1{
  border-top-color:var(--goldBright);
  background:linear-gradient(132deg,#FDF7E1 0%,#F5DC9C 50%,#E3B94F 100%);
}
.pcard.t2{
  border-top-color:var(--grass);
  background:linear-gradient(132deg,#EEFAE8 0%,#C0EBAF 50%,#8BD97B 100%);
}
.pcard.t3{
  border-top-color:transparent;
  background:linear-gradient(132deg,#FFFFFF 0%,#E6EAEF 32%,#C6CDD5 60%,#EDF1F4 82%,#B7BFC7 100%);
}
.pcard.t3:before{
  content:"";position:absolute;left:-1px;right:-1px;top:-4px;height:4px;border-radius:15px 15px 0 0;
  background:linear-gradient(90deg,#8E9BA8 0%,#E8EDF2 26%,#AEB8C4 52%,#F7FAFC 76%,#97A4B1 100%);
}
.pcard.t4{
  border-top-color:var(--red);
  background:linear-gradient(132deg,#FDEFEA 0%,#F4D1C8 50%,#E4A795 100%);
}
.pcard.t1:hover{border-color:#C79A19;box-shadow:0 16px 34px -16px rgba(190,145,20,.5)}
.pcard.t2:hover{border-color:#35A82A;box-shadow:0 16px 34px -16px rgba(53,168,42,.45)}
.pcard.t3:hover{border-color:#99A2AB;box-shadow:0 16px 34px -16px rgba(118,128,138,.5)}
.pcard.t4:hover{border-color:#962D1F;box-shadow:0 16px 34px -16px rgba(150,45,31,.42)}
.pcard .pn{font-size:11.5px;font-weight:700;letter-spacing:.12em;color:var(--ink3)}
.pcard.t1 .pn,.pcard.t1 .kv div:first-child b{color:#5A4204}
.pcard.t2 .pn,.pcard.t2 .kv div:first-child b{color:#0F4A16}
.pcard.t3 .pn,.pcard.t3 .kv div:first-child b{color:#33404D}
.pcard.t4 .pn,.pcard.t4 .kv div:first-child b{color:#732015}
.pcard h3{font-size:18px;margin:5px 0 5px;letter-spacing:-.01em;color:#211C15}
.pcard .goal{font-size:12.5px;color:#3B352B;margin:0 0 16px;min-height:58px;line-height:1.55;flex:1}
.pcard .kv{border-top-color:rgba(35,31,25,.16)}
.pcard .kv span,.pcard .pfoot{color:#4C453A}
.pcard .kv b small{color:#4C453A}
.recbadge{position:absolute;top:-10px;right:14px;background:#8A6414;color:#fff;font-size:10.5px;letter-spacing:.06em;padding:2px 9px;border-radius:20px;font-weight:600}
.kv{display:grid;grid-template-columns:1fr 1fr;gap:11px 8px;padding-top:14px;border-top:1px dashed var(--line);margin-top:auto}
.kv b{display:block;font-size:20px;letter-spacing:-.03em;font-variant-numeric:tabular-nums;line-height:1.25}
.kv b small{font-size:11px;font-weight:500;color:var(--ink3);letter-spacing:0;margin-left:2px}
.kv span{font-size:11px;color:var(--ink3)}
.pfoot{margin-top:13px;font-size:11.5px;color:var(--ink3);text-align:right}
.guide{
  margin-top:18px;background:var(--blueSoft);border:1px solid #D9E8DC;border-radius:12px;
  padding:13px 18px;font-size:13px;color:var(--blue);display:flex;gap:26px;flex-wrap:wrap;align-items:center;
}
.guide .k{font-size:11.5px;letter-spacing:.12em;color:#7C8F82;font-weight:700}
.guide b{font-weight:700}

/* ---------- book link ---------- */
.bt{color:inherit;text-decoration:none;border-bottom:1px solid rgba(35,31,25,.12);
  transition:color .15s,border-color .15s}
a.bt:hover{color:var(--blue2);border-bottom-color:var(--blue2)}
.cand .ci a.bt,.alert a.bt{border-bottom-color:rgba(35,31,25,.09)}
.pickcard a.bt{border-bottom-color:rgba(31,41,51,.22)}
.pickcard a.bt:hover{color:#0B141C;border-bottom-color:#0B141C}

/* ---------- table ---------- */
.tbl{width:100%;border-collapse:separate;border-spacing:0;font-size:13.5px;background:var(--card);border-radius:14px;overflow:hidden;box-shadow:var(--shadow);border:1px solid var(--line)}
.tbl th{background:var(--blue);color:#fff;font-weight:600;font-size:12px;padding:11px 9px;text-align:center;letter-spacing:.03em;white-space:nowrap}
.tbl td{padding:9px;border-bottom:1px solid var(--line2);text-align:center;font-variant-numeric:tabular-nums;vertical-align:middle}
.tbl tr:last-child td{border-bottom:0}
.tbl tbody tr:hover{background:#F7FAF7}
.tbl .idx{color:var(--ink3);font-size:12px;width:34px}
.tbl .bk{text-align:left;max-width:326px}
.tbl .bk b{display:block;font-family:var(--kai);font-weight:400;font-size:15.5px;line-height:1.42;letter-spacing:.01em}
.tbl .bk small{color:var(--ink3);font-size:11.5px}
.tbl td.money{font-weight:600}
.tbl td.hi{color:var(--red);font-weight:700}
.tbl td.gray{color:var(--ink3)}
.tfoot td{background:#F5F9F5;font-weight:700;font-size:13px;border-top:2px solid #D9E7DC}
.chip{display:inline-block;font-size:11px;padding:1px 7px;border-radius:6px;background:var(--blueSoft);color:var(--blue2);white-space:nowrap}
.chip.pay{background:var(--redSoft);color:var(--red)}
.chip.exp{background:var(--goldSoft);color:var(--gold)}
.risk{font-weight:600;font-variant-numeric:tabular-nums}
.tbl td.v{color:#9A6B15;font-weight:700;font-size:14px}
.tbl tr.ok{background:#EAF7EC}
.tbl tr.ok:hover{background:#DFF3E2}
.tbl td.now{color:#2C9921;font-weight:600}
.lowplans{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
.lowplan{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px;box-shadow:var(--shadow);display:flex;flex-direction:column}
.lowplan h3{font-size:17px;margin:0 0 6px;color:var(--ink);font-weight:600}
.lowplan h3 i{font-style:normal;color:var(--gold);margin-right:6px;font-size:18px}
.lowplan .lpstat{font-size:12.5px;color:var(--ink2);margin-bottom:8px}
.lowplan .lpstat b{color:var(--red);font-size:16px}
.lowplan p{font-size:12.5px;color:var(--ink3);line-height:1.65;margin:0 0 10px}
.lowplan ul{list-style:none;padding:8px 0 0;margin:0;border-top:1px dashed var(--line2)}
.lowplan li{display:flex;justify-content:space-between;gap:8px;font-size:12.5px;padding:4px 0;border-bottom:1px dashed var(--line2)}
.lowplan li:last-child{border-bottom:0}
.lowplan li span{color:var(--ink3);white-space:nowrap;font-variant-numeric:tabular-nums}
.lowplan li.more{color:var(--ink3);justify-content:flex-start}
.lowplan.zero{border-color:#E3C9C2;background:#FDF6F4}
.lplater{margin-top:14px;font-size:12.5px;color:var(--ink2);background:var(--goldSoft);border:1px solid #EAD9A8;border-radius:10px;padding:10px 14px;line-height:1.7}
@media(max-width:760px){.lowplans{grid-template-columns:1fr}}

.risk.r{color:var(--red)}
.risk.y{color:var(--gold)}
.risk.g{color:var(--ink3)}

/* mini bars */
.vs{display:flex;align-items:center;gap:6px;justify-content:center}
.vs i{display:block;width:46px;height:5px;background:var(--line);border-radius:9px;overflow:hidden;position:relative}
.vs i s{position:absolute;left:0;top:0;height:100%;background:linear-gradient(90deg,#E3C270 0%,#C1932A 100%);border-radius:9px}
.vs b{font-size:12.5px;font-weight:700;min-width:28px;text-align:right;color:#9A6B15}

/* ---------- plan 3 ---------- */
.pickcard{
  background:linear-gradient(132deg,#FFFFFF 0%,#E7EBEF 30%,#C7CED6 58%,#EEF2F5 80%,#B5BDC5 100%);
  color:#1F2933;border:1px solid #C3C9D0;border-radius:20px;padding:30px 32px;display:grid;
  grid-template-columns:1.45fr 1fr;gap:34px;
  box-shadow:0 18px 40px -22px rgba(74,90,106,.45),inset 0 1px 0 rgba(255,255,255,.9);
}
.pickcard .lab{font-size:11.5px;letter-spacing:.16em;color:#7A5A0F;font-weight:700}
.pickcard h3{
  font-family:var(--kai);font-weight:400;font-size:clamp(27px,3.15vw,38px);
  margin:8px 0 4px;line-height:1.28;letter-spacing:.01em;color:#16212D;
}
.pickcard .by{font-size:13px;color:#3A4550;margin:0}
.pickcard .tags{display:flex;gap:8px;flex-wrap:wrap;margin:16px 0 0}
.pickcard .tags span{font-size:11.5px;background:rgba(255,255,255,.72);border:1px solid rgba(31,41,51,.16);padding:2px 9px;border-radius:20px;color:#33404D}
.pickcard .tags span.tagv{color:#7A5A0F;font-weight:700;background:rgba(255,255,255,.92);border-color:rgba(166,124,61,.45)}
.pickcard .fillbar{margin-top:20px}
.pickcard .fillbar .fr{display:flex;justify-content:space-between;font-size:12px;color:#3A4550;font-weight:600}
.pickcard .fillbar .ft{height:9px;background:rgba(31,41,51,.14);border-radius:99px;overflow:hidden;margin-top:7px}
.pickcard .fillbar .ft i{display:block;height:100%;background:linear-gradient(90deg,#66737F 0%,#A7B4C2 55%,#EEF3F8 100%);border-radius:99px}
.pickcard .nums{display:grid;grid-template-columns:1fr 1fr;gap:18px 10px;align-content:center}
.pickcard .nums b{display:block;font-size:27px;letter-spacing:-.03em;font-variant-numeric:tabular-nums;line-height:1.15;color:#16212D}
.pickcard .nums b small{font-size:12px;font-weight:500;color:#3A4550;letter-spacing:0;margin-left:3px}
.pickcard .nums span{font-size:11.5px;color:#3A4550}

.alt{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:16px}
.alt div{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;font-size:12.5px;color:var(--ink2)}
.alt div h5{margin:0 0 4px;font-size:12px;color:var(--ink3);font-weight:600;letter-spacing:.04em}
.alt div b{color:var(--ink);font-size:16.6px;font-family:var(--kai);font-weight:400;letter-spacing:.01em}
.alt div em{font-style:normal;color:var(--red);font-weight:600}

/* ---------- type grid ---------- */
.tgh{display:flex;align-items:center;gap:10px;margin:38px 0 6px}
.tgh h3{font-size:19px;margin:0;letter-spacing:-.01em}
.tgh .cnt{font-size:12px;color:var(--ink3);background:var(--blueSoft);color:var(--blue2);padding:2px 9px;border-radius:20px;font-weight:600}
.tgnote{font-size:13px;color:var(--ink2);margin:0 0 20px;max-width:880px}
.tgnote b{color:var(--red)}
.tgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:13px}
.tcard{background:var(--card);border:1px solid var(--line);border-radius:14px;overflow:hidden;box-shadow:var(--shadow);display:flex;flex-direction:column}
.tcard.hl{border-color:#D8B476;box-shadow:0 0 0 1px #DFC48B inset,0 10px 26px -16px rgba(166,124,61,.4)}
.thead{padding:11px 14px 10px;background:#F4F8F4;border-bottom:1px solid var(--line)}
.thead .tr{font-size:10.5px;color:var(--blue2);font-weight:700;letter-spacing:.09em}
.thead h4{margin:1px 0 2px;font-size:15px;letter-spacing:-.01em}
.thead .tm{font-size:11px;color:var(--ink3)}
.cand{display:grid;grid-template-columns:16px 1fr auto;gap:9px;padding:9px 13px;border-bottom:1px solid var(--line2);align-items:center}
.cand:last-child{border-bottom:0}
.cand .cr{font-size:11px;color:var(--ink3);font-variant-numeric:tabular-nums}
.cand .ci b{display:block;font-family:var(--kai);font-weight:400;font-size:16.8px;line-height:1.42;color:var(--ink);letter-spacing:.01em}
.cand .ci small{color:var(--ink3);font-size:11px}
.cand .cn{text-align:right;font-variant-numeric:tabular-nums;line-height:1.3}
.cand .cn b{font-size:13.5px;color:#9A6B15}
.cand .cn small{display:block;font-size:10px;color:var(--ink3)}
.cand .cn small.rt{color:#4590C9;font-size:11px;font-weight:600}
.cand.pick{background:#FFFBF1}
.cand.pick .ci b:after{content:"已选";font-family:"PingFang SC","Microsoft YaHei",sans-serif;font-style:normal;font-size:9.5px;background:var(--gold);color:#fff;padding:0 5px;border-radius:4px;margin-left:6px;vertical-align:1px;letter-spacing:.04em}
.empty{padding:15px 14px;font-size:12px;color:var(--ink3);line-height:1.6;flex:1}
.empty b{color:var(--ink2)}

/* ---------- boards ---------- */
.board{margin-top:34px}
.board-h{display:flex;align-items:baseline;gap:12px;margin-bottom:6px;flex-wrap:wrap}
.board-h h3{font-size:19px;margin:0;letter-spacing:-.01em}
.board-h .tag{font-size:11.5px;padding:2px 9px;border-radius:20px;font-weight:600;letter-spacing:.04em}
.board-h .tag.now{background:var(--grassSoft);color:#2C9921}
.board-h .tag.next{background:var(--blueTagSoft);color:var(--blueTag)}
.board-h h3{position:relative;padding-left:12px}
.board-h h3:before{content:"";position:absolute;left:0;top:.28em;width:4px;height:1.05em;border-radius:2px;background:var(--line)}
.board.now .board-h h3:before{background:#2C9921}
.board.next .board-h h3:before{background:var(--blueTag)}
.board.now .tbl th{background:#2C9921}
.board.next .tbl th{background:var(--blueTag)}
.board-h.risk h3:before{background:var(--red)}
.tbl.rk th,#s-p4 .tbl th{background:var(--red)}
#s-p4 .tfoot td{background:#FAECE8;border-top-color:#F2D2CB}
.board-h .tip{font-size:12.5px;color:var(--ink3);margin-left:auto}
.board-note{font-size:12.5px;color:var(--ink3);margin:0 0 12px}
.tbl td .rt{font-weight:600;font-variant-numeric:tabular-nums}
.plus{color:var(--gold);font-weight:700}

.riskbox{display:grid;grid-template-columns:repeat(3,1fr);gap:13px;margin-top:14px}
.rs{background:var(--card);border:1px solid var(--line);border-top:3px solid var(--red);border-radius:13px;padding:16px 18px}
.rs b{display:block;font-size:26px;letter-spacing:-.03em;font-variant-numeric:tabular-nums;line-height:1.2;color:var(--red)}
.rs b small{font-size:12px;font-weight:500;color:var(--ink3);margin-left:4px;letter-spacing:0}
.rs span{display:block;font-size:12px;color:var(--ink3);margin-top:3px}
.rs.b,.rs.g{border-top-color:var(--red)}

/* ---------- reward box ---------- */
.rewardbox{margin-top:18px;background:var(--card);border:1px solid var(--line);border-left:3px solid var(--gold);border-radius:13px;padding:18px 20px}
.rhead{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:14px}
.rhead h4{margin:0;font-size:15px;font-weight:700}
.rhead .pill{font-size:11.5px;background:var(--goldSoft);color:var(--gold);font-weight:600;padding:2px 10px;border-radius:20px}
.rgrid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
.rgrid h6{margin:0 0 4px;font-size:12.5px;color:var(--blue2);font-weight:700}
.rgrid p{margin:0;font-size:12.5px;color:var(--ink2);line-height:1.65}
.rgrid p b{color:var(--ink);font-weight:600}
.rfoot{margin:14px 0 0;padding-top:12px;border-top:1px dashed var(--line);font-size:12.5px;color:var(--ink2);line-height:1.7}
.rfoot b{color:var(--red);font-weight:700}

/* ---------- score rule ---------- */
.steps{display:grid;grid-template-columns:repeat(4,1fr);gap:13px;margin-top:8px}
.step{background:var(--card);border:1px solid var(--line);border-radius:13px;padding:16px;position:relative}
.step .sn{font-size:10.5px;font-weight:700;letter-spacing:.14em;color:var(--blue2)}
.step h5{margin:5px 0 5px;font-size:14.5px;font-weight:600}
.step p{margin:0;font-size:12px;color:var(--ink2);line-height:1.6}
.step code{background:#F3F8F4;color:var(--blue);padding:1px 5px;border-radius:5px;font-size:11.5px}
.result{
  margin-top:14px;background:#F4F8F4;border:1px dashed #D0E3D6;border-radius:12px;padding:14px 18px;
  font-size:13px;color:var(--blue);display:flex;gap:22px;align-items:center;flex-wrap:wrap;
}
.result b{font-weight:700}
.result .formula{margin:0;background:transparent;padding:0}

/* ---------- notes ---------- */
.nlist{display:grid;grid-template-columns:1fr 1fr;gap:13px}
.note{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:15px 17px;font-size:13px;color:var(--ink2)}
.note b{display:block;color:var(--ink);font-size:13.5px;margin-bottom:4px}
.note .em{font-weight:600;color:var(--ink)}
.note.r{border-left:3px solid var(--red)}
.note.b{border-left:3px solid var(--blue2)}
.formula{background:#F4F8F4;border-radius:8px;padding:8px 11px;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12px;color:var(--blue);margin-top:7px;line-height:1.7}

footer{padding:36px 0 52px;font-size:12.5px;color:var(--ink3)}
footer b{color:var(--ink2)}
.fwrap{display:flex;justify-content:space-between;gap:20px;flex-wrap:wrap}

#top{
  position:fixed;right:24px;bottom:24px;width:40px;height:40px;border-radius:50%;
  border:1px solid var(--line);background:#fff;color:var(--blue2);cursor:pointer;
  box-shadow:var(--shadow);display:none;place-items:center;font-size:15px;z-index:60;
}
#top.show{display:grid}

@media (max-width:1080px){
  .pcards{grid-template-columns:1fr 1fr}
  .tgrid{grid-template-columns:1fr 1fr}
  .funnel{grid-template-columns:repeat(3,1fr)}
  .fstep:nth-child(3):after{display:none}
  .hero-grid{grid-template-columns:1fr;gap:24px}
  .steps,.rgrid{grid-template-columns:1fr 1fr}
  .riskbox{grid-template-columns:1fr 1fr}
}
@media (max-width:760px){
  .wrap{padding:0 18px}
  .pcards,.tgrid,.flags,.nlist,.alt,.steps,.rgrid,.riskbox{grid-template-columns:1fr}
  .pickcard{grid-template-columns:1fr;padding:24px}
  .funnel{grid-template-columns:1fr;gap:8px}
  .fstep{display:flex;align-items:center;gap:14px;padding:10px 14px}
  .fstep b{font-size:20px;min-width:58px}
  .fstep span{font-size:11.5px;margin-top:0}
  .fstep:after{display:none}
  .tbl{font-size:12px}
  .tbl .bk{max-width:none}
  .scroll{overflow-x:auto;margin:0 -18px;padding:0 18px}
  .tbl{min-width:720px}
}
.swipe{display:none}
@media (max-width:760px){
  .swipe{display:block;font-size:11.5px;color:var(--ink3);margin:0 0 6px;text-align:right}
}
@media print{
  #nav,#top{display:none}
  .bt{border-bottom:0}
  body{background:#fff}
  section{page-break-inside:avoid}
  .pcard:hover{transform:none}
}
"""


# ------------------------------------------------------------------ 组装
def fmt_cnt(n):
    n = int(n or 0)
    if n >= 10000:
        return '%.1f 万人' % (n / 10000.0)
    return '{:,} 人'.format(n)


def wallet_parts():
    a = (M.WALLET or {}).get('android') or {}
    bal = round(float(a.get('balance') or COIN), 2)
    gift = round(float(a.get('giftBalance') or 0), 2)
    paid = round(bal - gift, 2)
    return bal, gift, paid


def sold_stats():
    shelf = M.BOOKS
    sold = [b for b in shelf if b['_sold']]
    need_sold = sum(1 for b in M.NEED if b['_sold'])
    return len(shelf), len(sold), need_sold


def expiry():
    """权益/赠币福利到期时间（取 wallet 的 welfare.showExpiredTime，回退会员卡 expiredTime）"""
    import datetime
    W = M.WALLET or {}
    ts = ((W.get('android') or {}).get('welfare') or {}).get('showExpiredTime') \
        or (W.get('memberCard') or {}).get('expiredTime')
    if not ts:
        return '未取到', '未取到'
    d = datetime.datetime.fromtimestamp(int(ts))
    full = d.strftime('%Y-%m-%d %H:%M')
    short = ('今天 %s' % d.strftime('%H:%M')) if d.date() == datetime.date.today() else d.strftime('%m-%d %H:%M')
    return full, short


def hero():
    need = len(M.NEED)
    ok = len(M.COIN_OK)
    sold = M.COIN_OK_SOLDOUT
    dropped = len(M.COIN_POOL) - len(M.COIN_POOL_X)
    pool = len(M.COIN_POOL_X)
    bal, gift, paid = wallet_parts()
    n_shelf, n_sold, n_need_sold = sold_stats()
    sold_pct = 100.0 * n_sold / n_shelf if n_shelf else 0
    exp_full, exp_short = expiry()
    top_risk = sorted(M.COIN_POOL_X, key=lambda b: -b['_risk'])[:2]
    top_risk_max = top_risk[0]['_risk'] if top_risk else 0
    ex = '、'.join(bl(b, wrap=True) for b in top_risk)
    if top_risk:
        alert_txt = ('有 <b>%d 本</b>付费书在你买之前就已经下架了。可买池里风险最高的是 %s，'
                     '下架概率 <b>%s</b>（池均值 %s）。'
                     '<b>想买就别等，先买风险高的那几本。</b>'
                     % (n_need_sold, ex, pct1(top_risk_max), pct1(M.RISK_AVG)))
    else:
        alert_txt = ('有 <b>%d 本</b>付费书在你买之前就已经下架了。'
                     '当前余额 %s 元买不到任何一本，可买池为空，'
                     '所以暂时没有「该先买哪本」的排序 —— <b>书币补上后重跑本脚本即可。</b>'
                     % (n_need_sold, m2(COIN)))
    return f"""
<header class="hero"><div class="wrap">
  <div class="hero-grid">
    <div>
      <p class="kicker">微信读书 · 书币余额与购书方案</p>
      <h1>{m2(bal)} 元书币，<br><em>该怎么花</em></h1>
      <p class="lede">账户余额 <b>{m2(bal)} 元</b>，其中赠币 {m2(gift)} 元、充值书币 {m2(paid)} 元
      （{'全部为赠币' if paid <= 0 else '两种币混用'}），权益到期时间 <b>{exp_full}</b>。
      书架 {n_shelf} 本里已经有 <b>{n_sold} 本下架（{sold_pct:.1f}%）</b>，其中 {n_need_sold} 本你还没买断就已经买不到了。
      下架不可逆，今天能买的书，下周可能只剩「已下架」三个字。</p>
    </div>
    <div class="balcard">
      <span class="k">可用余额</span>
      <div class="big">{m2(bal)}<span>元</span></div>
      <ul>
        <li><span>赠币（阅读/活动所得）</span><b>{m2(gift)}</b></li>
        <li><span>充值书币</span><b>{m2(paid)}</b></li>
        <li class="warn"><span>权益到期</span><b>{exp_short}</b></li>
        <li><span>待付费书目</span><b>{need} 本</b></li>
      </ul>
    </div>
  </div>

  <div class="alert">
    <span class="t">下架预警</span>
    {alert_txt}
  </div>

  <div class="funnel">
    <div class="fstep"><b>{need}</b><span>待付费书目<br>（书架 {n_shelf} 本）</span></div>
    <div class="fstep"><b>{ok}</b><span>单本标价 ≤ {m2(COIN)} 元<br>任选其一都付得起</span></div>
    <div class="fstep out"><b>{('−%d' % sold) if sold else '0'}</b><span>剔除已下架<br>买不到也读不了</span></div>
    <div class="fstep out"><b>{('−%d' % dropped) if dropped else '0'}</b><span>剔除已读完未买断<br>未获永久授权</span></div>
    <div class="fstep final"><b>{pool}</b><span>真实可买池<br>四方案共同候选</span></div>  </div>
</div></header>
"""


def nav():
    return """
<nav id="nav"><div class="wrap navwrap">
  <span class="brand">购书方案</span>
  <a href="#s-plans">四方案对比</a>
  <a href="#s-p1">① 性价比最高</a>
  <a href="#s-p2">② 本数最多</a>
  <a href="#s-p3">③ 绝对价值最高</a>
  <a href="#s-p4">④ 抢下架</a>
  <a href="#s-boards">好书榜</a>
  <a href="#s-rule">评分规则</a>
  <a href="#s-notes">口径说明</a>
  <span class="navbal">余额 <b>%.2f</b> 元</span>
</div></nav>
""" % COIN


def plan_cards():
    rows = [
        ('t1', '01', '性价比最高', True,
         '按「价值评分 ÷ 标价」从高到低装，每一元买到的价值最多。',
         M.PS_RATIO, 'p1',
         [('本数', '%d' % M.PS_RATIO['n'], '本'),
          ('花费', m2(M.PS_RATIO['spent']), '元'),
          ('剩余', m2(M.PS_RATIO['left']), '元'),
          ('性价比', m1(M.PS_RATIO['per']), '分/元')]),
        ('t2', '02', '本数最多', False,
         '0/1 背包本数优先，同本数下价值分合计最高——书架厚度最大化。',
         M.PS_CNT, 'p2',
         [('本数', '%d' % M.PS_CNT['n'], '本'),
          ('花费', m2(M.PS_CNT['spent']), '元'),
          ('剩余', m2(M.PS_CNT['left']), '元'),
          ('性价比', m1(M.PS_CNT['per']), '分/元')]),
        ('t3', '03', '绝对价值最高', False,
         '只买 1 本，一次把积分尽量花光：从 15 个一级分类各取 V 前 2 名，再挑最值的那一本。',
         M.PS_VAL, 'p3',
         [('本数', '1', '本'),
          ('花费', m2(M.PS_VAL['spent']), '元'),
          ('剩余', m2(M.PS_VAL['left']), '元'),
          ('花满度', '%.1f' % (100 * M.PS_VAL['fill']), '%')]),
        ('t4', '04', '抢下架', False,
         '只看价值分 ≥ 70 的书，最大化「下架风险 × 价值分」——先抢可能没了的好书。',
         M.PS_RISK, 'p4',
         [('本数', '%d' % M.PS_RISK['n'], '本'),
          ('花费', m2(M.PS_RISK['spent']), '元'),
          ('剩余', m2(M.PS_RISK['left']), '元'),
          ('性价比', m1(M.PS_RISK['per']), '分/元')]),
    ]
    out = ['<div class="pcards">']
    for cls, no, name, rec, goal, st, anchor, kvs in rows:
        badge = '<span class="recbadge">推荐</span>' if rec else ''
        cells = ''.join('<div><b>%s<small>%s</small></b><span>%s</span></div>' % (v, u, lab)
                        for lab, v, u in kvs)
        out.append(f"""<article class="pcard {cls}" data-go="{anchor}">
  {badge}<span class="pn">方案 {no}</span>
  <h3>{name}</h3>
  <p class="goal">{goal}</p>
  <div class="kv">{cells}</div>
  <div class="pfoot">查看逐本清单 ↓</div>
</article>""")
    out.append('</div>')
    return '\n'.join(out)


def plan_table(books, st, coin, accent_note=''):
    head = ['#', '书名 / 作者', '一级分类', '权益', '价值评分', '微信评分', '标价', '累计', '剩余', '性价比', '下架风险', '进度']
    th = ''.join('<th%s>%s</th>' % (' class="bk"' if h == '书名 / 作者' else '', h) for h in head)
    body = []
    spent = 0.0
    for i, b in enumerate(books, 1):
        spent = round(spent + b['_price'], 2)
        left = round(coin - spent, 2)
        p = per_of(b)
        cat = b['_cat']
        chip = 'pay' if cat == '付费会员卡可读' else ('exp' if cat == '体验卡可读' else '')
        short = {'付费会员卡可读': '会员卡', '体验卡可读': '体验卡'}.get(cat, cat)
        vs = b['_vscore'] or 0
        prog = '—' if b['_prog'] is None else '%d%%' % b['_prog']
        body.append(f"""<tr>
  <td class="idx">{i}</td>
  <td class="bk"><b>{bl(b, wrap=True)}</b><small>{E(b.get('author'))}</small></td>
  <td><span class="chip">{E(b['_type1'])}</span></td>
  <td><span class="chip {chip}">{E(short)}</span></td>
  <td><div class="vs"><i><s style="width:{min(100, vs):.0f}%"></s></i><b>{m1(vs)}</b></div></td>
  <td class="rating">{m1(b['_rating'])}</td>
  <td class="money">{m2(b['_price'])}</td>
  <td class="gray">{m2(spent)}</td>
  <td>{m2(left)}</td>
  <td class="hi">{m1(p) if p else '—'}</td>
  <td><span class="risk {risk_cls(b['_risk'])}">{pct1(b['_risk'])}</span></td>
  <td class="gray">{prog}</td>
</tr>""")
    foot = f"""<tr class="tfoot">
  <td colspan="2">合计 {st['n']} 本</td><td colspan="2"></td>
  <td class="v">{m1(st['tv'])}</td><td></td><td>{m2(st['spent'])}</td><td>{m2(st['spent'])}</td>
  <td>{m2(st['left'])}</td><td class="hi">{m1(st['per'])}</td>
  <td>{pct1(st['risk'])}</td><td></td>
</tr>"""
    note = '<p class="tnote">%s</p>' % accent_note if accent_note else ''
    return f'<div class="scroll"><table class="tbl"><thead><tr>{th}</tr></thead><tbody>{"".join(body)}{foot}</tbody></table></div>{note}'


def sec_plan(no, sid, title, sub, table_html):
    return f"""<section id="{sid}">
  <div class="wrap">
    <div class="sec-h"><span class="no">{no}</span><h2>{title}</h2></div>
    <p class="sec-sub">{sub}</p>
    {table_html}
    <div style="height:46px"></div>
  </div>
</section>"""


def sec_lowmode():
    """零钱模式：余额 ≤ LOW_TRIGGER 时，用全库 2 元清单给出的 3 个具体买法。

    为什么要单独一块：余额一两元时，基于书架的方案常常一本都排不出来，
    用户真正需要的是"这点钱买哪几本"，而不是"买不了"。
    """
    if not M.LOW_MODE:
        return ''
    plans = M.LOW_PLANS
    cap = M.LOW_CAP
    if not plans:
        return f"""<section id="s-lowmode">
  <div class="wrap">
    <div class="sec-h"><span class="no">零钱模式</span><h2>余额 {m2(COIN)} 元</h2></div>
    <p class="sec-sub">已进入零钱模式（余额 ≤ {m2(M.LOW_TRIGGER)} 元），但还没有全库清单。
    跑一次 <code>python scripts/fetch_lowprice.py</code> 生成后再看。</p>
  </div>
</section>"""

    best = plans[0]
    if best['plan']:
        lead = ('这 %s 元按「%s %s」能买 <b>%d 本</b>、花 %s 元、剩 %s 元，评分合计 %.1f。'
                '三个方案选一个执行，不要叠加。'
                % (m2(COIN), best['no'], best['name'], best['stat']['n'],
                   m2(best['stat']['spent']), m2(best['stat']['left']), best['stat']['tv']))
    else:
        cheap = min((x['_price'] for x in M.LOW_POOL_BAG), default=0)
        lead = ('余额 %s 元低于清单里最便宜的一本（%s 元），当前一本都买不了 —— '
                '做完本周阅读任务拿到 20 书币后再看。' % (m2(COIN), m2(cheap)))

    cards = []
    for pl in plans:
        st = pl['stat']
        if pl['plan']:
            items = ''.join(
                f'<li><b>{bl(x, wrap=True)}</b><span>{m2(x["_price"])} 元 · '
                f'{m1(x["_rating"])} 分</span></li>' for x in pl['plan'][:8])
            more = ('<li class="more">… 共 %d 本</li>' % len(pl['plan'])) if len(pl['plan']) > 8 else ''
        else:
            items = '<li class="more">本方案在当前预算内无解</li>'
            more = ''
        cards.append(f"""<div class="lowplan{' zero' if not pl['plan'] else ''}">
  <h3><i>{pl['no']}</i>{pl['name']}</h3>
  <div class="lpstat"><b>{st['n']}</b> 本 · 花 {m2(st['spent'])} 元 · 剩 {m2(st['left'])} 元</div>
  <p>{pl['desc']}</p>
  <ul>{items}{more}</ul>
</div>""")

    later_html = ''
    if M.LOW_LATER:
        ls = M.LOW_LATER['stat']
        _add = M.LOW_LATER.get('addN') or 0
        _head = ('能买 <b>%d 本</b>（多 %d 本）' % (ls['n'], _add)) if _add else \
                ('本数不变，但能换成更好的书（评分合计 +%.1f 分）' % (M.LOW_LATER.get('gain') or 0))
        later_html = f"""<p class="lplater">④ 如果完成本周阅读任务（+20 书币）：按 {m2(COIN + 20)} 元重算，
    {_head} —— {ls['n']} 本 / 花 {m2(ls['spent'])} 元 / 剩 {m2(ls['left'])} 元、
    评分合计 {ls['tv']:.1f}。书币来源与领取规则以 App 内活动页为准。</p>"""
    else:
        later_html = ('<p class="lplater">④ 再攒 20 书币重算，本数和评分都不会提升 —— '
                      '这个价位已到「精选 %d 本」的上限，不必为凑单而等。'
                      '想要更多可选书可调大 WEREAD_LOW_MAXBOOKS。</p>' % M.LOW_MAX_BOOKS)

    return f"""<section id="s-lowmode">
  <div class="wrap">
    <div class="sec-h"><span class="no">零钱模式</span>
      <h2>余额 {m2(COIN)} 元 · 三个具体买法</h2></div>
    <p class="sec-sub">{lead}<br>
    余额 ≤ {m2(M.LOW_TRIGGER)} 元，基于书架的四套方案大概率一本都排不出来，
    所以这里直接用<b>微信读书全库 ≤ {m2(cap)} 元、评分 ≥ 70 的 {len(M.LOW_SEED)} 本书</b>当候选池
    —— 它们不在你书架里，但可以直接买。价值函数用微信评分（全库书缺阅读人数等数据，算不了四层价值分）。</p>
    <div class="lowplans">{''.join(cards)}</div>
    {later_html}
    <div style="height:26px"></div>
  </div>
</section>"""


def sec_plans():
    return f"""<section id="s-plans" style="border-bottom:0;padding-bottom:0">
  <div class="wrap">
    <div class="sec-h"><span class="no">总览</span><h2>四套方案，选一套</h2></div>
    <p class="sec-sub">四个方案的目标互相冲突——「每一元最值」「本数最多」「单本最值」「最怕下架」不可能同时最优。
    <b>它们的候选池完全相同（{len(M.COIN_POOL_X)} 本，均已剔除下架与已读完未买断）</b>，差别只在优化目标。点卡片跳到逐本清单。</p>
    {plan_cards()}
    <div class="guide">
      <span class="k">怎么选</span>
      <span><b>没偏好 → ①</b>（性价比最高）</span>
      <span>想把书架堆厚 → ②</span>
      <span>只想认真读完一本 → ③</span>
      <span>好书怕没了 → ④</span>
    </div>
    <div style="height:52px"></div>
  </div>
</section>"""


def sec_p3():
    pv = M.VAL_PICK
    st = M.PS_VAL
    fill = 100 * st['fill']
    # 类内第 1 名（若选中的不是第 1）；余额连一本都买不起时 pv 为 None
    peers = (M.POW_CANDS.get(pv['_type1']) or []) if pv else []
    first = peers[0] if peers else None
    cmp_txt = ''
    if pv and first and first['bookId'] != pv['bookId']:
        cmp_txt = ('类内 V 第 1 是 %s（V %s，%s 元），但它只花掉余额的 %.1f%%；'
                   '本方案要「一次花光且买单本最值的」，所以按 V×标价 排序取 %s。'
                   % (bl(first, wrap=True), m1(first['_V']), m2(first['_price']),
                      100 * first['_price'] / COIN, bl(pv, wrap=True)))
    if pv:
        topv, topp = st['topV'], st['topPrice']
        body = f"""<div class="pickcard">
      <div>
        <span class="lab">最终选择</span>
        <h3>{bl(pv, wrap=True)}</h3>
        <p class="by">{E(pv.get('author'))}</p>
        <div class="tags">
          <span>{E(pv['_type1'])}</span>
          <span>类内 V 排名 {st['rank']} / {len(peers)}</span>
          <span class="tagv">价值评分 {m1(pv['_vscore'])}</span>
          <span>微信评分 {m1(pv['_rating'])}</span>
        </div>
        <div class="fillbar">
          <div class="fr"><span>积分花满度</span><span>{fill:.1f}%</span></div>
          <div class="ft"><i style="width:{fill:.1f}%"></i></div>
        </div>
      </div>
      <div class="nums">
        <div><b>{m2(pv['_price'])}<small>元</small></b><span>标价</span></div>
        <div><b>{m2(st['left'])}<small>元</small></b><span>买后剩余</span></div>
        <div><b>{m1(st['V'])}<small>分</small></b><span>综合价值 V</span></div>
        <div><b>{m1(pv['_sv'])}<small>分·元</small></b><span>单本价值产出 V×标价</span></div>
      </div>
    </div>

    <div class="alt">
      <div><h5>如果只追「书本身最值」</h5>
        <b>{bl(topv, wrap=True)}</b> · V {m1(topv['_V'])} · {m2(topv['_price'])} 元<br>
        但只花掉 <em>{100 * topv['_price'] / COIN:.0f}%</em> 的余额，剩下 {m2(COIN - topv['_price'])} 元闲置。</div>
      <div><h5>如果只追「剩余最少」</h5>
        <b>{bl(topp, wrap=True)}</b> · {m2(topp['_price'])} 元 · V {m1(topp['_V'])}<br>
        钱几乎花光，但 V 比选中书低 <em>{m1(pv['_V'] - topp['_V'])}</em> 分，性价比不划算。</div>
    </div>

    {('<p class="tnote" style="margin-top:14px">※ ' + cmp_txt + '</p>') if cmp_txt else ''}"""
    else:
        _nx = M.G_X.get('next_price')
        _why = ('待付费书里最便宜的一本《%s》要 %s 元'
                % (E(M.G_X.get('next_title') or '（未知书名）'), m2(_nx))) if _nx else '可买池为空'
        body = f"""<div class="pickcard">
      <div>
        <span class="lab">最终选择</span>
        <h3>本方案无书可选</h3>
        <p class="by">余额 {m2(COIN)} 元低于全部候选标价（{_why}）—— 不是没有候选书，是预算不够。</p>
        <div class="tags">
          <span>完成本周阅读任务 +20 书币后重跑即可</span>
        </div>
      </div>
      <div class="nums">
        <div><b>{m2(COIN)}<small>元</small></b><span>当前余额</span></div>
        <div><b>{m2(st['spent'])}<small>元</small></b><span>本方案花费</span></div>
        <div><b>0<small>本</small></b><span>可买本数</span></div>
        <div><b>0<small>个</small></b><span>候选位实到</span></div>
      </div>
    </div>"""
    return f"""<section id="s-p3">
  <div class="wrap">
    <div class="sec-h"><span class="no">方案 03</span><h2>绝对价值最高 · 只买一本</h2></div>
    <p class="sec-sub">目标不是数量，而是<b>把 {m2(COIN)} 元一次花在一本书上，且这本书本身足够值</b>。
    候选 = 「分类基准与类型评分」同一套 {len(M.TYPE1_STD)} 个一级分类 × 每类综合价值 V 前 2 名
    （{len(M.TYPE1_STD) * 2} 个候选位，实到 {st['nCand']} 本），
    再在单本标价 ≤ 余额的候选里按「V × 标价」取最大。</p>

    {body}

    <div class="tgh">
      <h3>分类后续高价值书目</h3>
      <span class="cnt">{len(M.TYPE1_STD)} 类 · {len(M.TYPE1_STD) * 2} 个候选位 · 实到 {st['nCand']} 本</span>
    </div>
    <p class="tgnote">按 {len(M.TYPE1_STD)} 个标准分类，选取每个分类前 2 名高价值书作为备选。</p>
    {type_grid()}
    <div style="height:46px"></div>
  </div>
</section>"""


def type_grid():
    pooln = {t: sum(1 for x in M.COIN_POOL_X if x['_type1'] == t and x['_vscore'])
             for t in M.TYPE1_STD}
    shelfn = {t: [x for x in M.BOOKS if x['_type1'] == t] for t in M.TYPE1_STD}
    glob = {b['bookId']: i + 1 for i, b in enumerate(M.VAL_CANDS)}
    pick_id = (M.VAL_PICK or {}).get('bookId')

    cards = []
    for t in M.TYPE1_STD:
        cands = M.POW_CANDS.get(t) or []
        hl = ' hl' if any(b['bookId'] == pick_id for b in cands) else ''
        head = f"""<div class="thead">
  <div class="tr">类排名 {M.TYPE1_RANK.get(t)} / {len(M.TYPE1_STD)}</div>
  <h4>{E(t)}</h4>
  <div class="tm">类均分 {m1(M.TYPE1_MEAN.get(t))} · 可买池有评分 {pooln.get(t, 0)} 本 · 候选 {len(cands)}/2</div>
</div>"""
        if not cands:
            n = len(shelfn[t])
            gone = sum(1 for x in shelfn[t] if x.get('soldout') == 1)
            zero = sum(1 for x in shelfn[t] if x['_price'] == 0)
            if not M.COIN_POOL_X:
                # 余额连一本都买不起：15 类全空，原因统一是预算，不能说成"全部已下架"
                body = f"""<div class="empty">
  <b>2 个候选位空缺。</b>余额 {m2(M.COIN)} 元低于全部待付费书标价，可买池为空 ——
  不是这类没有书，是这一轮一本都买不起。</div>"""
            elif gone == n and zero == n:
                body = f"""<div class="empty">
  <b>2 个候选位空缺。</b>书架里该类 {n} 本<b>全部已下架</b>（其中标价 0、永久可读的有 {zero} 本），
  不在待付费可买池中，所以无书可买。</div>"""
            else:
                body = f"""<div class="empty">
  <b>2 个候选位空缺。</b>书架里该类 {n} 本中 {gone} 本已下架、{zero} 本标价 0（免费/永久可读），
  余额 {m2(M.COIN)} 元内没有可买的付费书。</div>"""
        else:
            rows = []
            for j, b in enumerate(cands, 1):
                cls = ' pick' if b['bookId'] == pick_id else ''
                rows.append(f"""<div class="cand{cls}">
  <span class="cr">{j}</span>
  <div class="ci"><b>{bl(b, wrap=True)}</b><small>{E(b.get('author'))}</small></div>
  <div class="cn"><b>{m1(b['_V'])}</b><small class="rt">微信评分 {m1(b['_rating'])}</small><small>{m2(b['_price'])} 元</small></div>
</div>""")
            body = ''.join(rows)
        cards.append(f'<div class="tcard{hl}">{head}{body}</div>')
    return '<div class="tgrid">%s</div>' % ''.join(cards)


def board_table(rows, with_gap=False):
    heads = ['#', '书名 / 作者', '一级分类', '价值评分', '微信评分', '评分人数', '标价']
    if with_gap:
        heads.append('还需攒')
    heads += ['下架风险']
    th = ''.join('<th%s>%s</th>' % (' class="bk"' if h == '书名 / 作者' else '', h) for h in heads)
    body = []
    for i, b in enumerate(rows, 1):
        gap = ''
        if with_gap:
            gap = '<td class="plus">+%s</td>' % m2(b['_price'] - COIN)
        body.append(f"""<tr>
  <td class="idx">{i}</td>
  <td class="bk"><b>{bl(b, wrap=True)}</b><small>{E(b.get('author'))}</small></td>
  <td><span class="chip">{E(b['_type1'])}</span></td>
  <td><div class="vs"><i><s style="width:{min(100, b['_vscore'] or 0):.0f}%"></s></i><b>{m1(b['_vscore'])}</b></div></td>
  <td>{m1(b['_rating'])}</td>
  <td class="gray">{fmt_cnt(b['_rcnt'])}</td>
  <td class="money">{m2(b['_price'])}</td>
  {gap}
  <td><span class="risk {risk_cls(b['_risk'])}">{pct1(b['_risk'])}</span></td>
</tr>""")
    return f'<div class="scroll"><table class="tbl"><thead><tr>{th}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'


def _low_table(rows, cols):
    head = ''.join('<th>%s</th>' % h for h in cols)
    return ('<div class="scroll"><table class="tbl"><thead><tr>%s</tr></thead>'
            '<tbody>%s</tbody></table></div>' % (head, ''.join(rows)))


def _shelf_low_rows(pool, limit=None):
    out = []
    for i, b in enumerate(pool[:limit] if limit else pool, 1):
        can = b['_price'] <= COIN + 1e-9
        out.append(f"""<tr class="{'ok' if can else ''}">
  <td class="idx">{i}</td>
  <td class="bk"><b>{bl(b, wrap=True)}</b><small>{E(b.get('author'))}</small></td>
  <td><span class="chip">{E(b.get('_type1') or '未分类')}</span></td>
  <td class="money">{m2(b['_price'])}</td>
  <td class="v">{m1(b['_vscore'])}</td>
  <td>{m1(b.get('_rating') or 0)}</td>
  <td class="{'now' if can else 'gray'}">{'现在能买' if can else '差 %s 元' % m2(b['_price'] - COIN)}</td>
</tr>""")
    return out


def _seed_low_rows(books, limit=None):
    out = []
    for i, b in enumerate(books[:limit] if limit else books, 1):
        own = M.BY_ID.get(b.get('bookId'))
        if own and own['_cat'] == '永久可读':
            rel, cls = '已拥有', 'gray'
        elif own:
            rel, cls = '在书架未买断', 'hi'
        else:
            rel, cls = '未加书架', 'gray'
        out.append(f"""<tr>
  <td class="idx">{i}</td>
  <td class="bk"><b>{bl(b, wrap=True)}</b><small>{E(b.get('author'))}</small></td>
  <td><span class="chip">{E(M.type_disp(b))}</span></td>
  <td class="money">{m2(b.get('price') or 0)}</td>
  <td class="v">{m1(b.get('rating') or 0)}</td>
  <td class="gray">{fmt_cnt(b.get('rcnt') or 0)}</td>
  <td class="{cls}">{rel}</td>
</tr>""")
    return out


def sec_low():
    """零钱备选：书架的 ≤N 元书优先；书架一本都没有时，回落到全库采样清单。

    全库清单来自 data/lowprice_seed.json（fetch_lowprice.py 关键词采样），与用户书架无关，
    所以必须单独标注来源，不能让用户误以为是"你书架里的书"。
    """
    cap, shelf, seed = M.LOW_CAP, M.LOW_POOL, M.LOW_SEED
    if not shelf and not seed:
        return f"""<section id="s-low">
  <div class="wrap">
    <div class="sec-h"><span class="no">零钱</span><h2>≤ {m2(cap)} 元的高价值书</h2></div>
    <p class="sec-sub">你的书架里没有 ≤ {m2(cap)} 元、未下架、有评分的书，
    也还没有全库清单。跑一次 <code>python scripts/fetch_lowprice.py</code> 就能生成全库备选。</p>
  </div>
</section>"""

    blocks = []
    # ---------- 区 A：书架 ----------
    if shelf:
        aff = M.LOW_AFFORD
        if aff:
            sub = ('余额 %s 元，在「≤ %s 元」这一档里能买 <b>%d 本</b>。'
                   '下表按价值评分从高到低，<b>绿底 = 现在就能买</b>。'
                   % (m2(COIN), m2(cap), len(aff)))
        else:
            cheap = M.LOW_CHEAPEST
            sub = ('余额 %s 元还不够开张 —— 最便宜的一本 %s 要 %s 元，还差 <b>%s 元</b>。'
                   % (m2(COIN), bl(cheap, wrap=True), m2(cheap['_price']),
                      m2(cheap['_price'] - COIN)))
        blocks.append(f"""
    <div class="sec-h" style="margin-top:0"><span class="no">01</span><h3>你的书架 / 书单：%d 本</h3></div>
    <p class="sec-sub">{sub}</p>
    {_low_table(_shelf_low_rows(shelf), ['#', '书名 / 作者', '一级分类', '标价',
                                         '价值评分', '微信评分', '当前余额'])}""" % len(shelf))

    # ---------- 区 B：全库采样 ----------
    if seed:
        mr = M.LOW_SEED_META.get('minRating')
        mr_txt = ('%g' % mr) if isinstance(mr, (int, float)) else '70'
        meta_bits = []
        if M.LOW_SEED_META.get('keywordCount'):
            meta_bits.append('%d 个关键词' % M.LOW_SEED_META['keywordCount'])
        if M.LOW_SEED_META.get('requests'):
            meta_bits.append('%s 次请求' % M.LOW_SEED_META['requests'])
        if M.LOW_SEED_META.get('scanned'):
            meta_bits.append('扫描 %s 条' % M.LOW_SEED_META['scanned'])
        if M.LOW_SEED_META.get('elapsedMin'):
            meta_bits.append('抓取 %s 分钟' % M.LOW_SEED_META['elapsedMin'])
        if M.LOW_SEED_META.get('generatedAt'):
            meta_bits.append('生成于 %s' % M.LOW_SEED_META['generatedAt'])
        meta_txt = ('（%s）' % ' · '.join(meta_bits)) if meta_bits else ''
        _fixn = M.LOW_SEED_META.get('typeFixed') or 0
        _fix_txt = ('分类已按规则纠正 %d 本（平台原分类明显不当的，如《道德经》归到「历史」）。' % _fixn) if _fixn else ''
        _inf = getattr(M, 'SEED_INFERRED', 0)
        _inf_txt = ('其中 %d 本平台未给分类，分类名前的 <b>≈</b> 表示按书名 / 作者 / 出版社归入的近似类。' % _inf) if _inf else ''
        if shelf:
            lead = ('<b>备选：</b>如果你不想买上面那几本、或书架里这一档不够挑，'
                    '下面这份是<b>微信读书全库</b>的 ≤ %s 元高分书，与你书架无关，可以直接买。'
                    % m2(cap))
        else:
            lead = ('你的书架 / 书单里<b>没有任何 ≤ %s 元的可买书</b>，'
                    '所以这里用<b>微信读书全库</b>的 ≤ %s 元高分书当备选 —— '
                    '它们不在你书架里，但在微信读书能直接买。' % (m2(cap), m2(cap)))
        blocks.append(f"""
    <div class="sec-h" style="margin-top:36px"><span class="no">{'02' if shelf else '01'}</span>
      <h3>全库备选：≤ {m2(cap)} 元且微信评分 ≥ {mr_txt} 的 {len(seed)} 本</h3></div>
    <p class="sec-sub">{lead}<br>
    清单由关键词采样生成{meta_txt}。{_fix_txt}{_inf_txt}微信读书没有「按价格浏览全库」的接口，
    所以这是<b>采样而非穷举</b>；要扩大覆盖改 <code>fetch_lowprice.py</code> 顶部的 GROUPS 后重跑。</p>
    {_low_table(_seed_low_rows(seed, 12 if shelf else 20),
                ['#', '书名 / 作者', '一级分类', '标价', '微信评分', '评分人数', '与你书架'])}""")

    title = ('≤ %s 元的高价值书 · 你的书架' % m2(cap)) if shelf else \
            ('≤ %s 元的高价值书 · 全库备选' % m2(cap))
    kicker = ('余额不够时先看这里' if shelf
              else '你的书架这一档是空的，用全库清单兜底')
    tail = '两区都按评分排序' if (shelf and seed) else '按评分排序'
    return f"""<section id="s-low">
  <div class="wrap">
    <div class="sec-h"><span class="no">零钱</span><h2>{title}</h2></div>
    <p class="sec-sub">{kicker}；{tail}，书名可点进微信读书。</p>
    {''.join(blocks)}
    <div style="height:30px"></div>
  </div>
</section>"""


def sec_boards():
    pool_top = sorted([b for b in M.COIN_POOL_X if b['_vscore']],
                      key=lambda b: (-b['_vscore'], b['_price']))[:15]
    FUT = COIN + 20
    nexts = [b for b in M.NEED if not b['_sold'] and COIN < b['_price'] <= FUT and b['_vscore']]
    nexts.sort(key=lambda b: (-b['_vscore'], b['_price']))
    next_top = nexts[:15]

    n_shelf, n_sold, n_need_sold = sold_stats()
    sold_pct = 100.0 * n_sold / n_shelf if n_shelf else 0
    risk_sorted = sorted([b for b in M.COIN_POOL_X if b['_vscore']],
                         key=lambda b: (-b['_risk'], -(b['_vscore'] or 0)))
    risk_top = risk_sorted[:8]
    rmax = risk_top[0]['_risk'] if risk_top else 0

    rrows = []
    for i, b in enumerate(risk_top, 1):
        rrows.append(f"""<tr>
  <td class="idx">{i}</td>
  <td class="bk"><b>{bl(b, wrap=True)}</b><small>{E(b.get('author'))}</small></td>
  <td><span class="chip">{E(b['_type1'])}</span></td>
  <td class="v">{m1(b['_vscore'])}</td>
  <td>{m1(b['_rating'])}</td>
  <td class="money">{m2(b['_price'])}</td>
  <td><span class="risk {risk_cls(b['_risk'])}">{pct1(b['_risk'])}</span></td>
</tr>""")

    return f"""<section id="s-boards">
  <div class="wrap">
    <div class="sec-h"><span class="no">榜单</span><h2>现买好书榜 Top15 vs 下周好书榜 Top15</h2></div>
    <p class="sec-sub">目前账户书币可买好书榜前 15；若延迟满足、完成微信读书任务下周新增 20 书币，
    可购买的好书榜一并列出供比较参考。</p>

    <div class="board now">
      <div class="board-h">
        <h3>① 现买好书榜 Top15</h3><span class="tag now">预算 {m2(COIN)} 元</span>
        <span class="tip">可买池 {len(M.COIN_POOL_X)} 本内 · 有评分 {len([b for b in M.COIN_POOL_X if b['_vscore']])} 本</span>
      </div>
      <p class="board-note">单本标价都在余额之内，任选一本都付得起；按价值分降序取前 15。</p>
      {board_table(pool_top)}
    </div>

    <div class="board next">
      <div class="board-h">
        <h3>② 下周好书榜 Top15</h3><span class="tag next">预算 {m2(FUT)} 元</span>
        <span class="tip">{m2(COIN)}–{m2(FUT)} 元区间 · 有评分 {len(nexts)} 本</span>
      </div>
      <p class="board-note">区间内共 {len(nexts)} 本有评分的可购书，这里取价值分前 15。「还需攒」= 标价 − {m2(COIN)}。
      本周把阅读任务做满，预算就能从 {m2(COIN)} 元抬到 {m2(FUT)} 元（+50%），这批书一次性解锁。</p>
      {board_table(next_top, with_gap=True)}
    </div>

    <div class="rewardbox">
      <div class="rhead">
        <h4>本周还能领多少书币？</h4>
        <span class="pill">规则上限 20 书币 / 周</span>
      </div>
      <div class="rgrid">
        <div><h6>每周阅读奖励</h6><p>一周读满 7 天、累计 5 小时，在「我 → 福利 → 阅读奖励」选书币，<b>最高 20 书币</b>。</p></div>
        <div><h6>已领进度</h6><p>只能在 App 里看，网页端没有公开接口。<b>本次未读到你的已领记录</b>，所以只给上限、不报已领数。</p></div>
      </div>
    </div>

    <div style="height:34px"></div>
    <div class="board-h risk"><h3>下架风险：哪些书最该先买</h3></div>
    <p class="board-note">下架概率 = 0.5 × 权益类型实测下架率 + 0.3 × 出版年份段下架率 + 0.2 × 冷门度段下架率，
    全部来自你书架 {n_shelf} 本的实测下架比例，不是拍脑袋打分。</p>
    <div class="riskbox">
      <div class="rs"><b>{n_sold}<small>本</small></b><span>书架已下架，占 {sold_pct:.1f}%</span></div>
      <div class="rs b"><b>{n_need_sold}<small>本</small></b><span>待付费书还没买就已下架</span></div>
      <div class="rs g"><b>{pct1(rmax)}</b><span>可买池最高下架概率（均值 {pct1(M.RISK_AVG)}）</span></div>
    </div>
    <div style="height:14px"></div>
    <div class="scroll"><table class="tbl rk"><thead><tr><th>#</th><th class="bk">书名 / 作者</th><th>一级分类</th><th>价值评分</th><th>微信评分</th><th>标价</th><th>下架风险</th></tr></thead>
    <tbody>{''.join(rrows)}</tbody></table></div>
    <p class="board-note" style="margin-top:10px">这 8 本是可买池里风险最高的：多为老版经典与早期会员卡专享书。
    它们不是最该「收藏」的书，而是最该「别再拖」的书——真下架了就永久买不到了。</p>
    <div style="height:46px"></div>
  </div>
</section>"""


def sec_rule():
    _rated = [b for b in M.BOOKS if b['_vscore']]
    _full = [b for b in _rated if (b.get('_conf') or 0) >= 0.9999]
    n_rated, n_full = len(_rated), len(_full)
    pct_full = '%.1f' % (100.0 * n_full / max(1, n_rated))
    return f"""<section id="s-rule">
  <div class="wrap">
    <div class="sec-h"><span class="no">规则</span><h2>价值评分怎么算</h2></div>
    <p class="sec-sub">一句话：<b>以微信评分为起点，按「样本够不够、同类热不热、有没有人真的在读」逐项打折，只降不升</b>。
    这样算出来的分数一定不高于微信评分，属于「保守下界」——宁可低估，也不把冷门书吹成神作。</p>
    <div class="steps">
      <div class="step">
        <span class="sn">STEP 1</span>
        <h5>起点：微信评分</h5>
        <p>直接用平台给出的 0–100 分作为上限，不另外加权。</p>
      </div>
      <div class="step">
        <span class="sn">STEP 2</span>
        <h5>× 置信下界</h5>
        <p>评分人数少就往下压（Wilson 单侧 95% 下界）：<code>300 人 4.5★</code> 没有 <code>3 万人 4.3★</code> 可信。</p>
      </div>
      <div class="step">
        <span class="sn">STEP 3</span>
        <h5>× 类内参与率</h5>
        <p>同类书里它的评分/阅读人数排位。只做类内相对化，所以不会因为「整个类别都冷门」被冤枉。</p>
      </div>
      <div class="step">
        <span class="sn">STEP 4</span>
        <h5>× 互动证据</h5>
        <p>划线数、想法数（0.55 / 0.45 加权）。受众面只做 ±5% 微调，防止热门书碾压小众经典。</p>
      </div>
    </div>
    <div class="result">
      <span class="formula">价值分 = 微信评分 × min(置信下界 × 类内参与率 × 互动证据, √(微信评分 ÷ 9.1))</span>
      <span>⇒ <b>价值分恒 ≤ 微信评分</b>；封顶线随分数上升，只有 <b>9.1 分以上</b>才可能出现两者相等，
      8.5 分的书会被压到约 82。全库 {n_rated} 本有评分的书里只有 <b>{n_full} 本</b>（{pct_full}%）打满。</span>
    </div>
    <div style="height:44px"></div>
  </div>
</section>"""


def sec_notes():
    return f"""<section id="s-notes">
  <div class="wrap">
    <div class="sec-h"><span class="no">口径</span><h2>这些数字是怎么来的</h2></div>
    <p class="sec-sub">所有口径与 Excel《微信读书_权益分类与购书方案.xlsx》完全一致，明细见配套 Markdown 说明。</p>
    <div class="nlist">
      <div class="note r"><b>为什么剔掉「已读完未买断」</b>
        这 {len(M.COIN_POOL) - len(M.COIN_POOL_X)} 本已读到 100%、但没买断（靠体验卡/会员卡读完），
        买它只是拿个永久副本，没有阅读增量。四个方案都不含它们。</div>
      <div class="note r"><b>为什么剔掉「已下架」</b>
        标价虽在余额内，但已下架的书既买不到也读不了：这类有 {M.COIN_OK_SOLDOUT} 本（待付费书里已下架的一共 {M.NEED_SOLDOUT} 本），全部剔除。</div>
      <div class="note b"><b>价值分（0~100）</b>微信评分做统计下界修正 × 类内参与率 × 互动证据 × 受众面，
        封顶不高于微信评分，属于「保守下界」。
        <div class="formula">价值分 = 微信评分 × min(置信下界 × 类内参与率 × 互动证据, 1)</div></div>
      <div class="note b"><b>综合价值 V（只用于方案③）</b>把评分、参与率、阅读人数换算成同一量纲；
        <span class="em">价格不进 V</span>，价格只在最后一步决定「能花掉多少积分」。
        <div class="formula">V = 0.45×价值分 + 0.25×微信评分<br>&nbsp;&nbsp;&nbsp;+ 0.15×参与率百分位 + 0.15×阅读人数百分位</div></div>
      <div class="note"><b>「性价比」= 总价值分 ÷ 总花费</b>
        天然偏爱低价书，所以方案①的书目偏薄；想「只买一本花光积分」用方案③。
        方案③只有 1 本，不与①②④ 横向相加或比较。</div>
      <div class="note"><b>下架风险（实测 0–43%）</b>综合上架时长、下架概率与稀缺度。
        候选池平均 {pct1(M.RISK_AVG)}，作为判断某个方案「是否偏向快下架的书」的基准线。</div>
    </div>
    <div style="height:44px"></div>
  </div>
</section>"""


def footer():
    snap = M.SNAP
    return f"""<footer><div class="wrap fwrap">
  <div>数据快照：账户资产 {E(snap.get('账户资产'))}　·　书目/评分/阅读人数 {E(snap.get('书目/评分/阅读人数'))}　·　阅读进度 {E(snap.get('阅读进度'))}</div>
  <div>完整 15 张表见 <b>微信读书_权益分类与购书方案.xlsx</b>　·　口径说明见 <b>微信读书_方法与口径说明.md</b></div>
</div></footer>"""


JS = r"""
<script>
(function(){
  var cards = document.querySelectorAll('.pcard');
  cards.forEach(function(c){
    c.addEventListener('click', function(){
      var el = document.getElementById('s-' + c.dataset.go);
      if(el) el.scrollIntoView({behavior:'smooth', block:'start'});
    });
  });

  if (window.innerWidth <= 760) {
    document.querySelectorAll('.scroll').forEach(function(el){
      var p = document.createElement('p');
      p.className = 'swipe';
      p.textContent = '← 表格可左右滑动查看 →';
      el.parentNode.insertBefore(p, el);
    });
  }

  var links = Array.prototype.slice.call(document.querySelectorAll('#nav a'));
  var secs = links.map(function(a){ return document.querySelector(a.getAttribute('href')); });
  function sync(){
    var y = window.scrollY + 120, cur = 0;
    secs.forEach(function(s, i){ if(s && s.offsetTop <= y) cur = i; });
    links.forEach(function(a, i){ a.classList.toggle('on', i === cur); });
    document.getElementById('top').classList.toggle('show', window.scrollY > 600);
  }
  window.addEventListener('scroll', sync, {passive:true});
  sync();

  document.getElementById('top').addEventListener('click', function(){
    window.scrollTo({top:0, behavior:'smooth'});
  });
})();
</script>
"""


def build():
    if M.RATIO_PLAN:
        _lo = min(b['_price'] for b in M.RATIO_PLAN)
        _nlo = sum(1 for b in M.RATIO_PLAN if abs(b['_price'] - _lo) < 1e-9)
        f1 = ('※ 其中 %d 本标价仅 %s 元，「性价比」被极小分母拉到 %.0f 分/元，'
              '但它们合计只花 %s 元，不影响其余选择。'
              % (_nlo, m2(_lo), M.PS_RATIO['per'], m2(_nlo * _lo)))
    else:
        f1 = '※ 余额 %s 元买不到任何一本，本方案为空（不是没有候选书，是预算不够）。' % m2(COIN)
    p1 = plan_table(M.RATIO_PLAN, M.PS_RATIO, COIN, f1)

    if M.PS_CNT['n'] and M.G_X.get('next_price'):
        f2 = ('※ %d 本已是本数上限（第 %d 便宜的也要 %s 元 > 剩余 %s 元）。'
              '代价是掺进几本低分书，性价比只有 %.1f 分/元。'
              % (M.PS_CNT['n'], M.PS_CNT['n'] + 1, m2(M.G_X['next_price']),
                 m2(M.PS_CNT['left']), M.PS_CNT['per']))
    elif M.PS_CNT['n']:
        f2 = ('※ 可买池内 %d 本已全部买下（余额还剩 %s 元，没有更便宜的书了）。'
              % (M.PS_CNT['n'], m2(M.PS_CNT['left'])))
    else:
        f2 = '※ ' + M.cap_note()
    p2 = plan_table(M.MAIN_PLAN, M.PS_CNT, COIN, f2)
    p4 = plan_table(M.RISK_PLAN, M.PS_RISK, COIN,
                    '※ 只看价值分 ≥ 70 的高价值书，目标是「在它们下架前先拿到」。'
                    '它和方案①重合 %d 本：最值的书往往也是最容易没的书。' % M.OVERLAP_RR)

    doc = (
        '<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        '<title>购书方案 · %.2f 元书币怎么花</title>\n<style>' % COIN
        + CSS +
        '</style>\n</head>\n<body>\n'
        + hero() + nav()
        + '<main>' + sec_lowmode() + sec_plans()
        + sec_plan('方案 01', 's-p1', '性价比最高 · 逐本清单',
                   '按「价值评分 ÷ 标价」从高到低装，钱够就买。这是<b>每一元买到的价值最多</b>的解，'
                   '也是推荐方案——同样 %.2f 元，它的性价比 %.1f 分/元，高于方案② 的 %.1f 分/元。'
                   % (COIN, M.PS_RATIO['per'], M.PS_CNT['per']), p1)
        + sec_plan('方案 02', 's-p2', '本数最多 · 逐本清单',
                   '0/1 背包本数优先、同本数下价值分合计最高。<b>适合想把书架堆厚</b>，'
                   '代价是性价比被低价书拉低。', p2)
        + sec_p3()
        + sec_plan('方案 04', 's-p4', '抢下架 · 逐本清单',
                   '只考虑价值分 ≥ 70 的书，最大化「下架风险 × 价值分」的合计。'
                   '<b>适合“好书怕没了”的焦虑型买法</b>，平均下架概率 %.1f%%（候选池均值 %.1f%%）。'
                   % (M.PS_RISK['risk'], M.RISK_AVG), p4)
        + sec_boards()
        + sec_low()
        + sec_rule()
        + sec_notes()
        + '</main>\n' + footer()
        + '<button id="top" title="回到顶部">↑</button>\n' + JS
        + '\n</body>\n</html>\n'
    )
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(doc)
    print('saved:', OUT, len(doc), 'bytes')


if __name__ == '__main__':
    build()
