# 踩坑速查

## 路径与运行环境

| 坑 | 现象 | 解法 |
|---|---|---|
| 相对路径 | `FileNotFoundError: _web_shelf2.json` | `build_excel.py` 在**模块顶层**就读数据文件。正路是 `run_pipeline.py`（`cwd=work` 起子进程）；单脚本直跑请设 `WEREAD_WORK=<数据目录>`，`wconfig` 会在 import 时自动 `chdir` |
| `import build_excel` 失败 | `ModuleNotFoundError` | 已改为按**脚本自身所在目录**插入 `sys.path`，脱离 `PYTHONPATH` 也能跑；`run_pipeline.py` 仍会额外设置 `PYTHONPATH=scripts` 兜底 |
| profile 目录 | 换目录后登录态"消失" | 默认 `~/.weread-buybooks/profile`（在仓库外）；`WEREAD_PROFILE` 可用绝对路径，或用相对路径由 `WEREAD_WORK` 解析成绝对路径后再使用 |
| 缺浏览器通道 | `Executable doesn't exist at ...msedge.exe` | `wconfig.channel()` 会探测本机安装位置并自动降级到 chrome / 内置 chromium；换机器先跑 `python scripts/wconfig.py` 看判定结果 |
| node 找不到 | 书单解析 `nuxt eval fail` | `WEREAD_NODE` 指定绝对路径，或把 node 加进 PATH |

## Python / 格式化

| 坑 | 现象 | 解法 |
|---|---|---|
| openpyxl 居中 | `ValueError` | `Alignment(CENTER)` 不行，写 `Alignment(horizontal='center', vertical='center')` |
| `%` 格式化遇中文 | `ValueError: unsupported format character` | 文本里出现单个 `%`（如"36.9%的书"）必须转义成 `%%`，或干脆改用 `.format()` / f-string。HTML 用占位符 `.replace()` 最稳，避开 CSS 花括号与 `%` 双重冲突 |
| 大 JSON 写一半 | 文件损坏、下次读不出来 | 先写 `.tmp` 再 `os.replace()`；`ensure_ascii=False` 保留中文 |
| `json.load` 无默认 | 可选文件缺失直接崩 | 一律 `try/except` 给 `{}`，或缺文件时明确报错并提示先跑哪个阶段 |
| 旧数据文件缺失 | `score_value.py` 崩在 `_bookmeta.json` | 已修：`_bookmeta*.json` 只作补缺，缺失时给 `{}`；`_bookmeta_v5.json` 缺失则明确提示先 `--only fetch` |

## HTML / CSS

| 坑 | 现象 | 解法 |
|---|---|---|
| class 名冲突 | 已选书名变白、看不见 | `.pick` 被 `.cand.pick` 复用污染 → 大卡改名 `.pickcard`，并给 `.cand .ci b` 显式 `color:var(--ink)` |
| `opacity` 表达"剔除" | 文字发灰不好读 | 用边框/颜色表达语义，别降透明度 |
| 行楷加粗 | 字糊成一团 | 行楷无粗体，`font-weight:400` |
| 对比度不足 | 亮绿/青蓝/亮金当小字色 | 小字降档到 `#2C9921` / `#1B6FC4` / `#9A6B15`；逐 stop 实算相对亮度 |
| 书名链接丢失 | 手写 `《{}》` 后不可点 | 一律走 `bl(b, wrap=True)` |
| URL 里的 `&` | HTML 结构被截断 | `html.escape` 处理 href（`&` → `&amp;`） |
| 移动端横向溢出 | 页面左右能拖 | 表格包 `.scroll`，不要给容器设 `overflow:visible` |
| 打印截断 | A4 少半张表 | 加 `@media print` 规则，去下划线、去阴影、避免 `break-inside` 切表 |

## 抓取

| 坑 | 现象 | 解法 |
|---|---|---|
| 并行请求 | 返回空对象 / 403 | 严格串行，每本 `sleep(0.3~0.5s)` |
| 无登录态硬跑 | 抓回一堆空记录，后续全错 | 启动先查 `wr_vid`，缺失立即 `LOGIN_LOST` 退出 |
| 断点续跑状态错乱 | 跳过了没抓成功的书 | 状态文件记 `done` 与 `fails` 两张表，重试只针对 `fails` |
| SSR 页面解析 | `__NUXT__` 里取不到 `readingCount` | 用 `nuxt_extract.js` 在 Node VM 里求值，不要用正则硬切 |
| 字段口径混用 | "阅读人数"和"当前在读"混为一谈 | 累计阅读人数只在 `/book-detail?type=1&v=` 的 `readingStat.readingCount` |
| 接口改版 | 某天开始全 `fails` | 看 `_refresh_meta.json` 的失败率，>5% 就单本调试 `r.status` / `r.text()[:200]` |

## 计算与一致性

| 坑 | 现象 | 解法 |
|---|---|---|
| 只重出 HTML | 数字和 Excel 对不上 | HTML `import build_excel`，同一进程同一份内存对象；禁止在 HTML 里写常量 |
| 只改评分参数 | 四方案、清理候选全变但没重跑 | 参数属全局，改了必须 `--only score,excel,html,charts` 整链 |
| 两处同类逻辑 | 方案③候选的 15 类 与 分类基准表的 15 类不一致 | 同一口径出现两处，改一处必须同步 |
| **余额极低 / 空池** | 生成 Excel 时 `TypeError: must be real number, not NoneType`（卡在「逐本方案」）、`'NoneType' object is not subscriptable` | 余额低于最便宜的书时 `COIN_POOL` 为空、`next_price`/`VAL_PICK` 为 `None`。`make_group` 已用 `OUT_POOL`（单本超预算的书）给 `next_price` 兜底；Excel / HTML / MD 三处一律用 `cap_note()` 一类的分支函数，禁止把 `None` 直接塞 `%.2f`；`VAL_PICK` 取值前必判空。测试时把 `_wallet3.json` 的 `android.giftBalance` 改成 `0.05` 与 `1.11` 各跑一遍 |
| 硬编码"上一轮的结果" | 数据一变就自相矛盾（如余额 0 元时仍写"实到 28 本""两本雪球专刊 0.1 元"） | 文案里的本数、书名、比例全部从变量取，不许抄某次运行的具体值 |
| 修复 bug 后忘更新记忆 | 下次重蹈覆辙 | 踩到新坑就补进本文件 |

## 交付

| 坑 | 现象 | 解法 |
|---|---|---|
| 文件名被占用 | `PermissionError` 保存失败 | 依次退到 `_new.xlsx` 再退旧名；提示用户关掉 Excel |
| 中文路径在 PowerShell | 文件名乱码 | 用 Python 脚本或编辑工具处理文件，避免在 shell 里拼中文路径 |
| 临时截图残留 | 项目目录一堆 png | 核验完立即删临时脚本与截图，不要留在工作目录 |
