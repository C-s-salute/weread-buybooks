---
name: weread-buybooks
version: 1.1.2
description: 微信读书书架资产分析与购书方案生成。抓取书架、书币余额、书目元数据与互动数据，用「统计下界 × 类内参与率 × 互动证据 × 受众面」四层模型给每本书打价值评分，再产出四套书币购书方案（Excel 工作簿 + 单文件 HTML 报告 + 口径说明 Markdown + 图表）。余额 ≤2 元时自动切「零钱模式」，改用随包分发的全库 ≤2 元高分书清单（474 本），给精选 3 本 / 只买 1 本 / 花光预算三种买法。当用户提到微信读书书架、书币或赠币余额、赠币不足怎么买、一两块钱能买什么书、低价好书清单、买书方案、书架价值评分、待付费书单、下架风险、清理书架、要不要买某本书、书单对比时使用。WeRead shelf analytics and book-buying plan generator (value score, coin budget, low-balance mode, delisting risk, Excel/HTML report).
agent_created: true
---

# 微信读书书架资产分析与购书方案

把「书架 → 该买什么、先买什么、钱够不够」这条链路自动化：抓数据 → 打价值评分 → 生成四套购书方案 → 出 Excel / HTML / Markdown 三份同源报告。

## 何时使用

任一情形即触发：

- 「我还有多少书币 / 赠币能买什么书」「帮我出个购书方案」
- 「书架里哪些书值得先买」「按价值给我书架排个序」
- 「哪些书快下架了」「书架清理该清哪些」
- 「这个书单值不值得加」「书单里的书和书架比怎么样」
- 已有报告需要按新数据重算，或只改报告样式重出

不适用：只查单本书信息、只管理书架增删 → 用 `weread-skills`。

## 硬性口径（不可违反）

这四条来自真实使用中的反复纠正，写错会被直接打回：

1. **书币 ≠ 现金**：不能提现，也**不能用来开/付会员卡**。任何"折现""抵扣会员费"的表述都是错的。
2. **「单本 ≤ 余额的 N 本」≠「N 本都买得起」**：必须显式说"任选其一"，不能算成总额。四方案里的预算都是"逐本累加到花完"。
3. **价值评分不出现版本号**：Excel、Markdown、HTML 里都不写 v2/v3/v4/v5，sheet 名就叫「价值评分」，表头叫「价值评分」。
4. **下架风险是概率，不是分数**：底层值 = `100 × (0.5×权益类型下架率 + 0.3×年份段下架率 + 0.2×冷门度段下架率)`，展示一律 `%.1f%%`，文案写「下架概率」不写「风险分」。

另外两条通用要求：

- 同一解释只出现一次。已在卡片/图表里说过的原因，正文不再复述。
- 不编数字。网页端没有兑换进度接口，已领书币只能在 App「我 → 福利」看；拿不到就写"未读到"，不要估。

## 环境准备

**硬性要求：Python ≥ 3.12 且必须是 64 位（推荐 3.13 64 位）。** 不要用 3.10/3.11 或 32 位解释器，也不要试图让它"凑合装"——最新 `numpy`（2.5）的 `requires-python` 已是 `>=3.12`、`matplotlib`（3.11）是 `>=3.11`，版本不够时 pip 只能降级装老版本；32 位则没有官方 wheel，pip 会退去编译源码并失败。

```bash
python -m pip install -U pip                # 老 pip 解析不到新 wheel tag
pip install -r requirements.txt             # 带下限/上限，适合日常
pip install -r requirements-lock.txt        # 精确锁定，复现同一环境（推荐首次安装用这个）
python scripts/check_environment.py          # 自检：版本/位数/依赖/浏览器/node，一次说清
python scripts/login.py                      # 扫码登录，登录态存进 profile 目录
```

实测通过的版本组合（`requirements-lock.txt`）：

| 包 | 版本 | 用途 |
|---|---|---|
| playwright | 1.64.0 | 浏览器自动化抓取 |
| openpyxl | 3.1.5 | Excel 工作簿 |
| matplotlib | 3.11.2 | 互动证据图表 |
| numpy | 2.5.0 | matplotlib 依赖 + 数值计算 |
| requests | 2.34.2 | 抓公开书单页 |

安装踩坑的唯一正解：**换 64 位 3.12+ 重建虚拟环境**，不要在原环境里反复试老版本。
`check_environment.py` 会把「Python 32 位 / 版本过低 / pip 过旧 / 缺依赖」分别报出来，并直接给修复命令，退出码 1 表示必须先修。

- 浏览器优先用本机 Edge（`channel=msedge`），没装则回退 Chrome / 内置 chromium，由 `scripts/wconfig.py` 自动判定。
- 登录态只存在 profile 目录（默认 `~/.weread-buybooks/profile`），**不落任何明文 cookie 文件**。扫码有效期有限，过期重跑 `login.py`。
- 公开书单解析需要 Node（`node` 在 PATH 里即可，可用 `WEREAD_NODE` 指定）；缺 node 只影响「书单」阶段，不影响书架分析与购书方案。
- 环境变量：`WEREAD_WORK`（数据目录，设置后脚本 import 即切到该目录，因此单脚本也能直接跑）、`WEREAD_PROFILE`（profile 目录；给相对路径时按 `WEREAD_WORK` 解析，可把登录态一并放进数据目录）、`WEREAD_BROWSER`、`WEREAD_NODE`、`WEREAD_FONT`（图表中文字体文件）、`WEREAD_LOW_PRICE`（「零钱备选」的价格上限，默认 2 元）。
- 移植到其他机器或交给其他 agent：见 `README.md` 的「安装到其他环境 / 交给其他 Agent」。**技能目录内不含任何个人路径、账号或数据文件**，`.gitignore` 已排除 `weread_data/`、`profile/`、`_*.json`、`微信读书_*`。

### 自检结果怎么走

`check_environment.py` 输出四种状态，按状态决定下一步，别跳过：

| 状态 | 含义 | 动作 |
|---|---|---|
| `ready` | 依赖齐备 | 直接抓取出报告，不再讲安装步骤 |
| `partial` | 核心可用、可选缺失（典型是缺 node） | 照常跑，只在结果里标注「书单部分未生成」 |
| `needs_setup` | 有阻塞项 | 只讲缺失项的修法，修完重检，不要继续跑 |
| `unavailable` | 登录失效 / 限流 / 域名不可达（运行时才出现） | 用已有数据出报告并标注数据时点；恢复步骤见 `references/setup-guide.md` |

登录态过期不重试到底：`login.py` 拿不到 `wr_vid` 就停，让用户扫码。缺依赖时各功能还能做什么，见 `references/setup-guide.md` 末尾的对照表。

## 执行流程

首次使用先自检环境（缺依赖/版本不对时它会直接给修复命令）：

```bash
python scripts/check_environment.py
```

然后一条命令跑全链，数据落在 `--work` 目录（默认 `./weread_data`）：

```bash
python scripts/run_pipeline.py --work ./weread_data
```

分阶段控制：

```bash
python scripts/run_pipeline.py --work ./weread_data --only fetch       # 只刷账户/书架/元数据/进度
python scripts/run_pipeline.py --work ./weread_data --only score,excel,html,charts   # 只重算报告
python scripts/run_pipeline.py --work ./weread_data --skip fetch       # 用已有数据出报告
```

| 阶段 | 脚本 | 做什么 | 产出 |
|---|---|---|---|
| `fetch` | `fetch_all.py` | 账户资产、书架全量、书目元数据、累计阅读人数、想法数、热门划线、阅读进度。本数越多越慢（1000 本约 10–15 分钟），可断点续跑 | `_wallet3.json` `_web_shelf2.json` `_bookmeta_v5.json` `_progress2.json` |
| `booklist` | `fetch_booklist.py` → `fetch_bookdetail.py` → `fetch_extra.py` | 公开书单（免登录）→ 书单内未加架书的详情 → 互动数据 | `_booklists.json` `_bookdetail_extra.json` `_bookmeta_v5_extra.json` |
| `score` | `score_value.py` | 四层价值评分，纯本地计算（<1 秒） | `_valuescore5.json` |
| `excel` | `build_excel.py` | 工作簿（常规 16 个 sheet，零钱模式触发时 17 个），顺带写 Markdown 口径说明 | `微信读书_权益分类与购书方案.xlsx`、`微信读书_方法与口径说明.md` |
| `html` | `build_html.py` | 单文件自包含报告，数据全部 import 自 `build_excel.py` | `微信读书_购书方案.html` |
| `charts` | `make_charts.py` | 互动证据映射曲线 + 各类型分差 | `fig_v5_interaction.png` |

**三份产物必须同源**：HTML 直接 `import build_excel`，不重算指标。改数字只改 `score_value.py` 与 `build_excel.py`，重跑下游即可，绝不在 HTML 里手写常量。

`build_excel.py` 只做数据装载、主要 sheet 与编排；零钱模式那两张表在 `report_low.py`（只管排版），选书算法在 `lowplan.py`（纯函数、显式传参）。改算法只动 `lowplan.py`，改表样式只动 `report_low.py`。

**余额极低也必须出完整报告**：余额低于最便宜的书时四套方案本数为 0，但流程不中断——Excel 的「零钱备选(≤N元)」sheet 与 HTML 的「零钱备选」区块会给出两个来源：
- **① 你的书架/书单**里 ≤N 元、未下架、有评分的书（能直接下单）
- **② 微信读书全库**的 ≤N 元高分书（`data/lowprice_seed.json`，与你书架无关）。书架这一档为空时，② 自动升为主内容。

价格上限由环境变量 `WEREAD_LOW_PRICE` 控制（默认 2 元，可设 1 / 3 / 5）。`next_price`、`VAL_PICK` 在空池时为 `None`，所有下游文案必须走分支函数（`cap_note()` 等），禁止直接进 `%.2f`。

### 零钱模式（余额 ≤ 2 元自动触发）

余额 ≤ `WEREAD_LOW_TRIGGER`（默认 **2 元**，环境变量可调）时，报告自动切换：

| 位置 | 变化 |
|---|---|
| Excel | 多一个 sheet「**零钱购书方案**」（排在「积分购买方案」之后） |
| HTML | 首屏多一个「**零钱模式**」区块，三张方案卡 |
| 候选池 | 换成**微信读书全库**的 ≤2 元清单（不是用户书架） |

给三个买法 + 一个参考，每本都带直达链接：

1. **精选 3 本**（`WEREAD_LOW_MAXBOOKS`，默认 3）：同一系列只留评分最高的一期，再按微信评分从高到低挑
2. **只买 1 本**：买得起的范围里评分最高的
3. **花光预算**：在评分 ≥75 的书里挑最贵的一本，剩余最少
4. **再攒 20 书币**：按 `余额+20` 重算，告诉用户"能多买几本 / 或本数不变但换成更好的书"

**为什么必须设本数上限**：清单里有大量 0.01 元的「雪球漫画投资课」「私募专刊」小册子，纯按「本数最多」做 0/1 背包会刷出 **39 本**——数学最优，但业务上荒谬（1.11 元买 39 本零碎小册）。所以先做同系列去重（标题去数字/期号后取前 8 字为 key），再限制最多 3 本。

**价值函数差异**：零钱模式用**微信评分**当价值（全库书缺阅读人数、划线等数据，算不了四层价值分）。报告文案里已标注这一点，不要和书架书的「价值评分」混着比。

### 刷新「全库 ≤N 元高分书」清单

`data/lowprice_seed.json` 随技能分发，是一份采样清单。想扩容或更新价格时重跑：

```bash
python scripts/fetch_lowprice.py                        # 默认 ≤2 元、评分 ≥70，按分组页数抓
python scripts/fetch_lowprice.py --cap 5 --min-rating 75
python scripts/fetch_lowprice.py --pages-scale 2         # 所有分组页数翻倍（更慢、更全）
python scripts/fetch_lowprice.py --pages 5               # 统一每词 5 页（覆盖分组设置）
python scripts/fetch_lowprice.py --keywords 论语,专刊 --pages 2 --out ./seed.json
```

- 数据源是 `/web/search/name_books?keyword=&maxIdx=&count=`（JSON、可翻页），返回里直接带 `price` 与 `newRating`，**不必逐本查价**；已抽验与 `/web/book/info` 一致。
- 关键词按 `GROUPS` 分组，**每组自带页数**（热门 5 页 / 中频 4 页 / 长尾 3 页 / 泛搜索 2 页），当前 222 词、15 组，全量一轮 **9.4 分钟**（904 请求，实测）。
- **改词表前先算预算**：`python -c "import sys;sys.path.insert(0,'scripts');from fetch_lowprice import GROUPS;print(sum(len(w)*p for _,p,w in GROUPS))"`，超过 950 请求就会破 10 分钟。
- 微信读书**没有「按价格浏览全库」的接口**，所以这是**关键词采样**、不是全库穷举：能否命中某本低价书，取决于 `GROUPS` 是否覆盖到它。
- 各组的实测产出、书名与作者名的对照结论、哪些方向（AI / 经济思想 / 武侠）在这个价位根本没有书、分类纠偏与推断规则，见 `references/keyword-research.md`。

### 体检词表 / 挖新关键词

```bash
python scripts/analyze_keywords.py                  # 贡献排行 + 零贡献清单 + 书名公约数
python scripts/analyze_keywords.py --min-cover 3    # 只列能覆盖 ≥3 本的书名片段
```

它给出三样东西，用来决定「删哪些词、加哪些词」：

1. 每个关键词实际命中几本（从种子的 `kw` 字段反推），零贡献词单独列出。当前 222 词里只剩 4 个（人工智能、大模型、张恨水、沟通），新增词前先用它确认没有白花请求的词
2. **书名公约数**：对已命中书名做 2–6 字 n-gram，再用贪心集合覆盖挑出「一个片段囊括最多书」的词，并标出哪些还没进词表
3. 未被任何片段覆盖的书（只能靠独立关键词捞）

加词前先用 `--keywords 词1,词2 --no-enrich --out ./t.json` 小跑验证命中率，别直接全量重抓（`--out` 要给工作目录内的相对路径，Windows 上 `/tmp` 不可用）。

### 加书单

把书单分享链接逐行写进 `weread_data/booklist_urls.txt`，再跑 `--only booklist,score,excel,html,charts`。书单页服务端渲染，`_booklists.json` 从 `__NUXT__` 解析，不需要登录。

### 可选：与 App 数字对照

把 App 书架页显示的四类数量写进 `weread_data/_app_ref.json`：

```json
{"付费会员卡可读": 0, "体验卡可读": 0, "永久可读": 0, "免费": 0}
```

工作簿第二节会多出「App显示 / 差异」两列并附差异说明；不放这个文件则只呈现本表统计，不做对照。该文件属个人数据，已在 `.gitignore` 里排除。

## 产出物

| 文件 | 内容 |
|---|---|
| `微信读书_权益分类与购书方案.xlsx` | 概览 / 积分购买方案（余额、四方案对比、逐本、全量、口径）/ 零钱购书方案（仅零钱模式）/ 待付费书单 / 永久可读 / 免费 / 已读完未买断 / 在读 / 价值评分 / 分类基准与类型评分 / 热书排名 Top150 / 书架清理候选 / 零钱备选(≤2元) / 书单总览 / 书单明细 / 书架分组明细 / 全量书籍 |
| `微信读书_购书方案.html` | 首屏（余额 + 下架预警 + 5 级漏斗）→ 零钱模式（仅零钱模式）→ 四方案卡 → 逐本清单 → 方案③候选网格 → 双榜 → 零钱备选 → 评分规则 → 下架风险 Top8 → 口径 |
| `微信读书_方法与口径说明.md` | 可获取信息与可做功能 / 互动证据 / 口径说明（这三块**不放进工作簿**） |

四套方案的定位（都在"排除已读完"的池子上算）：

1. **性价比最高** — 只在有评分的书里，按「价值评分 ÷ 标价」降序贪心
2. **本数最多** — 本数优先的贪心背包
3. **绝对价值最高** — 一次只买 1 本，在 ≤ 余额的候选里按「综合价值 V × 标价」取最大
4. **抢下架** — 价值评分 ≥ 70 且「下架概率 × 价值评分」最大

方案 ② 是**本数最多**、③ 是**绝对价值最高**，不要互换。

## 结果核验

改完必须自查，不能口头声称"已验证"：

- 三份产物在同一目录、时间戳一致
- HTML 无外链、可离线打开；≤760px 时表格走横向滚动容器，页面本身不溢出
- 浏览器控制台零报错；A4 打印不截断

改动 `score_value.py` 的曲线参数会连动所有下游（分类基准、四方案、方案③选中书、清理候选），**必须整链重跑**，不能只重出 HTML。

## 参考文档

| 文件 | 何时读 |
|---|---|
| `references/methodology.md` | 要解释或修改价值评分公式、四方案算法、清理候选阈值 |
| `references/data-sources.md` | 要改抓取逻辑、加字段、排查登录/限流/接口变更 |
| `references/report-design.md` | 要改 HTML/Excel 的配色、文案、表格结构 |
| `references/pitfalls.md` | 踩坑速查：openpyxl、字体、对比度、格式化 |
| `references/keyword-research.md` | 关键词实测结论、时间预算算法、分类纠偏与推断规则——改词表或改分类前读它 |
| `references/setup-guide.md` | 装环境、登录、凭据存储与撤销、缺依赖时的降级对照 |
| `skill-dependencies.json` | 机器可读的依赖清单与功能降级映射（改依赖时同步改它） |
