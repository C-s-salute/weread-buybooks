# weread-buybooks

**把微信读书书架变成一份可执行的购书方案。** 抓书架与书币余额 → 给每本书打一个抗样本量偏差的价值分 → 用四套互不重叠的算法排出「买哪些、先买哪些、钱够不够」→ 出 Excel / HTML / Markdown 三份同源报告。

原生支持 **Agent Skill** 协议：装进 WorkBuddy、Claude Code 等客户端后，说一句「帮我出个购书方案」就能跑完全程，不必记命令。

![release](https://img.shields.io/badge/release-v1.1.2-blue)
![license](https://img.shields.io/badge/license-MIT-green)
![python](https://img.shields.io/badge/python-3.12%2B%20(64--bit)-brightgreen)

---

## 它解决什么问题

平台推荐值是**样本量无关的算术均值**。一本 9.0 分的书可能只有 27 个评分人，也可能有 26 万人——直接按它排序买书，等于让「小样本高分」和「大样本高分」平起平坐，结果就是钱花在了证据最薄弱的书上。

再加上三个真实痛点：

- **书架越囤越多，买了没读**：分不清哪些是「已付费待读」、哪些是「免费可读」，清理时无从下手
- **书币有限，想花在刀刃上**：余额只有几十币时，是先买 3 本 8 分的，还是 1 本 9.3 分的？
- **快下架了才发现没买断**：权益类型（会员卡 / 体验卡 / 永久 / 免费）一变，书就读不了了，但没有预警

这四件事，本技能一次算清。

## 主要功能

### 1. 抓取

账户书币/赠币余额、书架全量（权益类型、标价、是否下架）、每本书的平台评分 / 评分人数 / 累计阅读人数 / 想法数 / 热门划线 / 阅读进度。断点续跑，中断不丢进度。1000 本约 10–15 分钟。

### 2. 四层价值评分

每层只修正一种偏差，参数集中在 `scripts/score_value.py` 顶部：

| 层级 | 修正什么 | 方法 |
|---|---|---|
| 统计下界 | 样本量差异 | Wilson 单侧 95% 置信下界（LCB） |
| 类内相对度 | 类型基线差异（心理类普遍 8.5+，技术类普遍 7.5） | 除以同类型中位数 |
| 互动证据 | 「只投票不动笔」的虚高 | 人均划线密度 / 人均想法密度 |
| 受众面 | 小众书误判 | Hill 曲线，权重仅 ±5% |
| 封顶 | 四乘数叠乘把置信度推到 1 | `min(1, (raw/91)^0.5)` |

> 早期版本用阅读人数直接放大可信度，把一本 83 分的书打到 23.7——样本量问题只能用统计下界解决，不能用热度当权重。

### 3. 四套购书方案

都在「排除已读完」的池子上计算，互不重叠：

| 方案 | 目标 | 算法 |
|---|---|---|
| ① 性价比最高 | 单位花费的价值最大 | 只在有评分的书里，按 `价值评分 ÷ 标价` 降序贪心 |
| ② 本数最多 | 买到最多本 | 0/1 背包，本数优先、同本数取价值分和最大 |
| ③ 绝对价值最高 | 把预算压在一本好书 | 15 个标准分类各取综合价值前 2，再挑 `V × 标价` 最大 |
| ④ 抢下架 | 先抢快没了的书 | 价值 ≥70 的书里最大化 `Σ(下架概率 × 价值评分)` |

### 4. 零钱模式（余额 ≤2 元自动触发）

赠币只剩一两块、连一本正价书都买不起时，自动切换候选池——**从微信读书全库**（不是你的书架）的 ≤2 元、评分 ≥70 清单里选，随包分发 474 本：

- **精选 3 本**：同系列只留评分最高的一期，再按评分降序
- **只买 1 本**：买得起的范围里评分最高的
- **花光预算**：评分 ≥75 里挑最贵的，剩余最少
- 外加「再攒 20 书币」对比参考

必须设本数上限——清单里有大量 0.01 元的雪球小册子，纯按「本数最多」做 0/1 背包会算出 **1.11 元买 39 本**，数学最优但业务上荒谬。

### 5. 三份同源报告

`16–17 sheet` 工作簿 + 单文件 HTML（自包含、可离线、可打印、移动端适配）+ 口径说明 Markdown + 图表。HTML 直接 `import build_excel`，不重算指标，改数字只改计算脚本。

## 适合谁用

### 典型场景

| 场景 | 你会得到 |
|---|---|
| 「我还有多少书币，能买什么书」 | 余额 + 四方案对比 + 逐本清单 + 总花费/剩余 |
| 「书架里哪些书值得先买」 | 全量价值评分排序，含评分/标价/性价比 |
| 「哪些书快下架了」 | 下架风险 Top8，带权益类型与概率 |
| 「书架清理该清哪些」 | 清理候选表（已读完未买断 / 低分 / 长期未读） |
| 「这个书单值不值得加」 | 书单逐本评分 + 与书架对比（免登录抓公开书单） |
| 「赠币只剩一块多」 | 零钱模式三方案，全库低价高分书 |
| 「改报告样式/换新数据」 | 分阶段重跑，改样式不必重新抓数据（<20 秒） |

### 不适用

- 只查单本书信息、只管理书架增删 → 用 `weread-skills`
- 想要平台全站榜单、他人书架、社交数据 → 网页端没有这些接口
- 兑换进度 / 已领书币明细 → 只能在 App「我 → 福利」查看，本项目不会编造该数字

---

## 安装

三种方式，选一种。**方式一最省事**（不碰命令行）。

### 方式一：在 WorkBuddy 里对话式安装（推荐）

`SKILL.md` 是 Agent Skill 标准入口，WorkBuddy 装了技能后**靠 description 自动判断何时触发**，所以装完不用记任何命令。

**第 1 步** 把仓库地址发给 WorkBuddy，说清楚装哪一级：

```
帮我把 https://github.com/C-s-salute/weread-buybooks 安装成用户级技能
```

> 用户级 = `~/.workbuddy/skills/`，所有项目可用；项目级 = `<项目>/.workbuddy/skills/`，只在该项目生效。说「装成项目级技能」即可切换。

**第 2 步** 让它把环境一起装好（这一步是关键，别省）：

```
装完依赖后跑 scripts/check_environment.py 自检，把结果告诉我
```

它会做四件事，你可以逐项核对它有没有跳过：

1. 克隆仓库到 `~/.workbuddy/skills/weread-buybooks/`
2. `pip install -r requirements-lock.txt`（实测精确版本）+ `playwright install chromium`（若本机已有 Edge/Chrome 可跳过）
3. 跑 `check_environment.py`，输出 `ready` / `partial` / `needs_setup`
4. 报 `needs_setup` 时给出缺失项的修复命令，而不是硬跑

**第 3 步** 扫码登录一次：

```
帮我跑 scripts/login.py 登录微信读书
```

登录态存在 `~/.weread-buybooks/profile`（仓库之外，不会被误提交）。

**第 4 步** 重启 WorkBuddy，技能列表里出现 `weread-buybooks` 就绪。之后直接说：

> 「帮我出个购书方案」·「我还有多少书币」·「书架哪些书值得先买」
> 「哪些书快下架了」·「赠币只剩一块多能买什么」

**手动兜底**（AI 装不了时）：

```bash
# macOS / Linux
cp -r weread-buybooks ~/.workbuddy/skills/

# Windows PowerShell
Copy-Item -Recurse weread-buybooks "$env:USERPROFILE\.workbuddy\skills\"
```

解压后的路径必须是 `skills/weread-buybooks/SKILL.md`（多套一层目录会识别不到）。

### 方式二：下载 Release 包

不想用 git 就走这个：

1. 打开 [Releases](https://github.com/C-s-salute/weread-buybooks/releases/latest)，下载 `weread-buybooks-1.1.2.zip`
2. 解压，把 `weread-buybooks/` 整个目录放进目标 `skills/` 目录（见上表路径）
3. 按「环境要求」装依赖、跑自检、扫码登录

### 方式三：克隆仓库，用 CLI 直接跑

不需要任何 Agent 客户端，纯命令行：

```bash
git clone https://github.com/C-s-salute/weread-buybooks.git
cd weread-buybooks
python -m pip install -U pip
pip install -r requirements-lock.txt       # 实测精确版本；日常也可用 requirements.txt
python -m playwright install chromium      # 若已装 Edge / Chrome 可跳过

python scripts/check_environment.py        # 自检：版本/位数/依赖/浏览器/node
python scripts/login.py                    # 微信扫码
python scripts/run_pipeline.py --work ./weread_data
```

分阶段跑（改样式时不必重新抓数据）：

```bash
python scripts/run_pipeline.py --work ./weread_data --only fetch                    # 只刷数据（10–15 分钟 / 1000 本）
python scripts/run_pipeline.py --work ./weread_data --only score,excel,html,charts  # 只重算报告（<20 秒）
python scripts/run_pipeline.py --work ./weread_data --skip fetch                    # 用已有数据出报告
```

跑完在 `./weread_data` 里得到：

```
微信读书_权益分类与购书方案.xlsx     16 个 sheet（零钱模式 17 个）
微信读书_购书方案.html               单文件报告，可离线/可打印
微信读书_方法与口径说明.md           可获取信息 / 互动证据 / 口径说明
fig_v5_interaction.png               互动证据映射曲线
```

### 环境要求（三种方式都适用）

**硬性要求：Python ≥ 3.12 且必须是 64 位（推荐 3.13 64 位）。**

用 3.10/3.11 或 32 位解释器会在装依赖时卡住：最新 `numpy`（2.5）的 `requires-python` 已是 `>=3.12`，`matplotlib`（3.11）是 `>=3.11`；32 位没有官方 wheel，pip 只能退去编译源码然后失败。**唯一正解是换 64 位 3.12+ 重建虚拟环境**，不要在原环境里反复试老版本。

| 依赖 | 版本 | 必需性 | 用途 |
|---|---|---|---|
| playwright | 1.64.0 | 必需 | 浏览器自动化抓取 |
| openpyxl | 3.1.5 | 必需 | Excel 工作簿 |
| matplotlib | 3.11.2 | 必需 | 互动证据图表 |
| numpy | 2.5.0 | 必需 | matplotlib 依赖 + 数值计算 |
| requests | 2.34.2 | 必需 | 抓公开书单页 |
| Node.js | 任意 LTS | 可选 | 仅解析公开书单，缺了不影响购书方案 |
| Edge / Chrome | 任意近期版 | 可选 | 有则复用，无则用内置 chromium |

`check_environment.py` 会把「32 位 / 版本过低 / pip 过旧 / 缺依赖 / 缺浏览器」分别报出来并直接给修复命令，**退出码 1 表示必须先修再往下走**。缺 node/浏览器时各功能还能做什么，见 `references/setup-guide.md` 末尾的降级对照表。

### 换机器后的最小清单

1. Python ≥ 3.12 64 位 → `python -m pip install -U pip && pip install -r requirements-lock.txt`
2. `python scripts/check_environment.py` 全绿
3. `python scripts/login.py` 扫码一次
4. `python scripts/run_pipeline.py --work ./weread_data`

### 便携模式：数据 + 登录态同目录

```bash
export WEREAD_WORK=./weread_data
export WEREAD_PROFILE=./.profile     # 相对路径，按 WEREAD_WORK 解析成 ./weread_data/.profile
python scripts/login.py
python scripts/run_pipeline.py --work ./weread_data
```

`.gitignore` 已排除 `profile/`、`.weread_*`、`_*.json`、`微信读书_*`，这个目录可整体拷到 U 盘或另一台机器继续用。

### 常用环境变量

| 变量 | 作用 |
|---|---|
| `WEREAD_WORK` | 数据目录；设置后脚本 import 即切到该目录，单脚本也能直接跑 |
| `WEREAD_PROFILE` | 登录态目录；给相对路径时按 `WEREAD_WORK` 解析 |
| `WEREAD_BROWSER` | 浏览器通道（msedge / chrome / chromium） |
| `WEREAD_NODE` | 指定 node 路径 |
| `WEREAD_FONT` | 图表中文字体文件 |
| `WEREAD_LOW_PRICE` | 零钱备选价格上限，默认 2 元（可设 1/3/5） |
| `WEREAD_LOW_TRIGGER` | 零钱模式触发阈值，默认 2 元 |
| `WEREAD_LOW_MAXBOOKS` | 零钱模式精选本数上限，默认 3 |

---

## 目录结构

```
weread-buybooks/
├── SKILL.md                  Agent Skill 入口（YAML frontmatter + 使用流程）
├── README.md                 本文件
├── LICENSE                   MIT
├── requirements.txt          依赖下限/上限（日常用）
├── requirements-lock.txt     实测通过的精确版本（复现环境用）
├── skill-dependencies.json   机器可读的依赖与功能降级映射
├── .gitattributes            * text=auto eol=lf（避免 Windows 假 diff）
├── .gitignore                抓取数据/产物/登录态/凭证，全部排除
├── data/
│   └── lowprice_seed.json    全库 ≤2 元、评分 ≥70 的 474 本书（随包分发）
├── references/               按需加载的参考文档
│   ├── methodology.md        评分公式、四方案算法、关键教训
│   ├── data-sources.md       接口清单、字段口径、权益类型判定
│   ├── report-design.md      配色、字体、文案红线
│   ├── keyword-research.md   关键词实测结论、时间预算、分类纠偏规则
│   ├── pitfalls.md           踩坑速查
│   └── setup-guide.md        装环境、登录、凭据撤销、缺依赖降级对照
└── scripts/
    ├── run_pipeline.py       一条命令跑全链（分阶段可组合）
    ├── check_environment.py  环境自检（版本/位数/依赖/pip/浏览器/node）
    ├── wconfig.py            统一配置：work/profile/浏览器/node/python
    ├── login.py              扫码登录，写入 profile
    ├── fetch_all.py          账户 + 书架 + 元数据 + 进度（断点续跑）
    ├── fetch_booklist.py     公开书单解析（免登录，需 node）
    ├── fetch_bookdetail.py   书单内未加架书的详情
    ├── fetch_extra.py        书单内未加架书的互动数据
    ├── fetch_lowprice.py     全库采样 ≤N 元高分书 → data/lowprice_seed.json
    ├── analyze_keywords.py   词表体检：贡献排行 / 零贡献词 / 书名公约数
    ├── score_value.py        四层价值评分（参数集中在文件顶部）
    ├── build_excel.py        数据装载 + 主要 sheet + 编排 + Markdown
    ├── report_low.py         零钱模式两张 sheet 的排版
    ├── lowplan.py            零钱模式选书算法（纯函数，可单测）
    ├── docs_md.py            口径说明 Markdown 生成
    ├── build_html.py         单文件 HTML（import build_excel 保证同源）
    ├── make_charts.py        图表
    └── nuxt_extract.js       从书单页 __NUXT__ 抽数据（沙箱受限执行）
```

## 文件功能明细

### 入口与编排

| 文件 | 说明 |
|---|---|
| `SKILL.md` | 技能定义。frontmatter 的 `description` 决定何时被触发；正文是硬性口径、执行流程、自检状态分支 |
| `scripts/run_pipeline.py` | 串起 fetch → booklist → score → excel → html → charts。`--only` 选阶段、`--skip` 跳过、`--work` 指定目录 |
| `scripts/wconfig.py` | 所有路径的唯一来源（`WEREAD_WORK` / `PROFILE` / `BROWSER` / `NODE`）。可单独运行打印实际取值 |

### 抓取层

| 文件 | 说明 |
|---|---|
| `fetch_all.py` | 主抓取。账户资产、书架全量、书目元数据、累计阅读人数、想法数、热门划线、阅读进度 |
| `fetch_booklist.py` | 公开书单页是服务端渲染的，直接解 `__NUXT__`，**无需登录** |
| `fetch_bookdetail.py` | 补书单里未加架书的详情 |
| `fetch_extra.py` | 补书单里未加架书的互动数据 |
| `fetch_lowprice.py` | 关键词采样全库低价书。`GROUPS` 15 组 222 词、每组带页数；`--reclassify` 秒级重算分类不重抓 |

### 计算层

| 文件 | 说明 |
|---|---|
| `score_value.py` | 四层评分唯一实现。参数写入产物 JSON 的 `params`，下游一律从 `params` 读 |
| `lowplan.py` | 零钱模式选书：同系列去重 + 0/1 背包 + 本数上限。纯函数、显式传参，可脱离全链单测 |
| `analyze_keywords.py` | 词表体检工具，输出贡献排行、零贡献词、书名公约数（n-gram + 贪心集合覆盖） |

### 输出层

| 文件 | 说明 |
|---|---|
| `build_excel.py` | 数据装载 + 主要 sheet + 编排（16–17 sheet），顺带写 Markdown |
| `report_low.py` | 零钱模式「零钱购书方案」「零钱备选」两张表的版式，靠 `configure()` 注入依赖避免循环引用 |
| `docs_md.py` | 生成口径说明 Markdown（这三块不进工作簿） |
| `build_html.py` | 单文件自包含 HTML，数据全部来自 `build_excel` |
| `make_charts.py` | 互动证据映射曲线、各类型分差 |
| `nuxt_extract.js` | 在受限上下文里解析书单页（禁 require/import/eval/Function，关掉 wasm 代码生成） |

### 参考文档

| 文件 | 何时读 |
|---|---|
| `references/methodology.md` | 要解释或修改评分公式、四方案算法、清理候选阈值 |
| `references/data-sources.md` | 要改抓取逻辑、加字段、排查登录/限流/接口变更 |
| `references/report-design.md` | 要改 HTML/Excel 的配色、文案、表格结构 |
| `references/keyword-research.md` | 要改词表或改分类——含实测产出、时间预算、单字边界坑 |
| `references/pitfalls.md` | 踩坑速查：openpyxl、字体、对比度、格式化 |
| `references/setup-guide.md` | 装环境、登录、凭据存储与撤销、缺依赖降级 |
| `skill-dependencies.json` | 机器可读的依赖与降级映射（改依赖时同步改它） |

## 扩展：加自己的书单 / 换低价清单

**加书单**——把微信读书公开书单分享链接逐行写进 `weread_data/booklist_urls.txt`：

```bash
python scripts/run_pipeline.py --work ./weread_data --only booklist,score,excel,html,charts
```

**扩容低价清单**——`data/lowprice_seed.json` 是采样结果，想更新价格或扩容就重跑：

```bash
python scripts/fetch_lowprice.py --cap 5 --min-rating 75    # 默认 ≤2 元、评分 ≥70
python scripts/fetch_lowprice.py --pages-scale 2            # 分组页数翻倍（更慢更全）
python scripts/analyze_keywords.py                          # 体检词表，看哪些词白花请求
```

改词表前先算请求预算，超过 950 请求就会破 10 分钟：

```bash
python -c "import sys;sys.path.insert(0,'scripts');from fetch_lowprice import GROUPS;print(sum(len(w)*p for _,p,w in GROUPS))"
```

微信读书**没有「按价格浏览全库」的接口**，所以这是关键词采样、不是全库穷举——能否命中某本低价书取决于 `GROUPS` 是否覆盖到它。

## 设计上的几个选择

- **登录态只存在浏览器 profile 里**，不落任何明文 cookie / token 文件
- **三份产物同源**：HTML 直接 `import build_excel`，不重算指标
- **参数写进产物**：评分参数随 JSON 的 `params` 一起落盘，避免两处各写一份系数
- **可断点续跑**：抓取中断后重跑跳过已完成项，失败项单独记表
- **权重不追热度**：样本量问题用统计下界解决，不用阅读人数放大可信度
- **硬编码零容忍**：路径一律走 `wconfig.py`，文案里禁止写死本数/书名/比例

## 合规与边界

本项目仅用于**个人阅读管理**，使用前请阅读并遵守以下边界：

- **接口来源**：所有接口均为逆向自网页端（微信读书无公开 API 文档），**可能随时失效**，不保证可用性
- **数据范围**：仅抓取**自己账号可见**的数据，不涉及他人书架、社交关系或评论
- **数据用途**：仅用于个人阅读决策。**请勿二次分发原始抓取结果**、勿用于商业用途或训练数据集
- **不编造数据**：微信读书网页端没有兑换进度接口，已领书币只能在 App「我 → 福利」查看；拿不到就写「未读到」，绝不估算
- **不做的事**：不绕过付费墙、不批量注册账号、不做大规模爬取、不提供任何绕过平台限制的手段
- **请求节制**：默认请求间隔 0.35 秒，已限制并发；**请勿调低**，调低导致的限流/封号由使用者自负
- **账号风险**：自动化访问可能违反平台用户协议。**是否使用、如何使用由你自行判断并承担风险**，作者不提供任何形式的担保
- **停止使用**：若平台条款不允许此类访问，或接口变更导致无法使用，请**立即停止使用**，不要试图绕过
- **授权与许可**：`data/lowprice_seed.json` 中的书目信息（书名/作者/出版社/价格/评分）版权归微信读书及相应版权方所有，此处仅作技术演示与个人检索索引，不含任何书籍内容

## 贡献

Issue / PR 都欢迎，但请先读 `references/methodology.md` 与 `references/pitfalls.md`——很多「看起来是 bug」的行为是刻意设计（例如本数上限、封顶函数）。

**提交前请确认**：没有把 `weread_data/`、`profile/`、`_*.json`、`微信读书_*` 或任何账号凭证带进提交。

## License

[MIT](LICENSE)
