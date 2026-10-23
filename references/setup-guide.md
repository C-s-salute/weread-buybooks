# 依赖配置与恢复

本技能依赖本地运行环境与微信读书网页登录态，不依赖 MCP 或第三方 API Key。先跑一次只读自检，按输出决定是否需要看本文件：

```bash
python scripts/check_environment.py
```

输出三种状态：`ready` 直接进入抓取与出报告；`partial` 核心可用、可选功能缺失（缺 node 只影响书单）；`needs_setup` 有阻塞项，按提示修完再重跑。`unavailable` 由运行时判定——登录态失效或被限流时，抓取阶段会直接报出。

---

## python-runtime

- 用途：全流程（抓取、评分、出报告）。matplotlib / numpy 依赖官方预编译 wheel，必须是 64 位。
- 支持版本：Python 3.12 或 3.13，64 位；推荐 3.13。
- 官方主页：https://www.python.org/downloads/
- 官方文档：https://docs.python.org/3/
- 下载地址：https://www.python.org/downloads/windows/
- 登录：不需要。
- 权限：无。
- 配置步骤：
  1. 安装 64 位 Python 3.12 / 3.13（安装包名含 `64-bit`，不要选 32-bit）。
  2. `py -3.13 -m venv .venv` 建独立虚拟环境。
  3. `.venv\Scripts\activate` 激活。
  4. `python -m pip install -U pip`。
  5. `pip install -r requirements-lock.txt`（精确版本）或 `requirements.txt`（带上下限）。
  6. `python scripts/check_environment.py` 应输出 `ready`。
- 验证：`python -c "import sys,platform;print(sys.version.split()[0],platform.architecture()[0])"` 输出 `3.12+` 与 `64bit`。
- 常见故障：版本过低时 pip 只能装老版 numpy/matplotlib；32 位下没有 wheel，pip 会退去编译源码并失败。**唯一正解是换 64 位 3.12+ 重建虚拟环境**，不要在原环境反复试老版本。
- 轮换与撤销：删除 `.venv` 目录即完全清除；系统级 Python 通过系统卸载入口移除。
- 核验日期：2026-10-21。

## playwright-browser

- 用途：抓取书架、账户资产、书目互动数据。缺失时只能出已有数据的报告。
- 支持版本：playwright 1.64.0（`>=1.40,<2`）。
- 官方主页：https://playwright.dev/
- 官方文档：https://playwright.dev/python/docs/intro
- 下载地址：https://pypi.org/project/playwright/
- 登录：不需要（浏览器通道本身不需要账号）。
- 权限：无系统级权限。
- 配置步骤：
  1. `pip install "playwright>=1.40,<2"`。
  2. 本机有 Edge 或 Chrome 即可直接复用；没有则 `python -m playwright install chromium`。
  3. `python scripts/check_environment.py`，确认出现 `[OK] 浏览器通道`。
- 验证：自检里浏览器通道一行显示 `msedge` / `chrome` / `chromium`。
- 故障处理：通道探测失败时可用 `WEREAD_BROWSER` 显式指定；`python -m playwright install` 下载失败多为网络问题，可换镜像后重试。
- 轮换与撤销：`pip uninstall playwright` 即可；下载的浏览器缓存位于用户目录下的 playwright 缓存文件夹，可整目录删除。
- 核验日期：2026-10-21。

## weread-account

- 用途：抓取书架与账户资产。这是唯一需要登录的依赖。
- 支持版本：微信读书网页版（2026-10 界面）。
- 官方主页：https://weread.qq.com/
- 官方文档：https://weread.qq.com/ （微信读书没有公开开发者文档，本技能只使用网页端自身接口）
- 登录网址：https://weread.qq.com/
- 登录后路径：浏览器打开 https://weread.qq.com/ → 右上角「登录」→ 微信扫码 → 回到书架页。终端执行 `python scripts/login.py`，脚本复用同一 profile 目录完成落盘。
- 权限与 scope：仅需普通用户登录态；不申请任何开放平台权限。
- 凭据存储：登录态只存在浏览器 profile 目录（默认 `~/.weread-buybooks/profile`），**不落明文 cookie / token 文件，不写日志、不写对话**。该目录已在 `.gitignore` 中排除，不随技能分发。
- 配置步骤：
  1. `python scripts/login.py`。
  2. 微信扫码确认。
  3. 脚本检测到 `wr_vid` cookie 后自动关闭浏览器并保存 profile。
- 验证：`login.py` 结束打印登录态已保存；抓取阶段不再提示未登录。
- 过期与失效：扫码有效期有限，过期后重跑 `login.py`。接口返回空或提示未登录时按失效处理，不要反复重试。
- 轮换与撤销：网页版右上角菜单退出登录，或删除 profile 目录；微信侧可在「微信 → 设置 → 个人信息与权限 → 授权管理」撤销网页授权。
- 核验日期：2026-10-21。

## node-for-booklist

- 用途：仅解析公开书单页的 `__NUXT__`。缺失只影响书单，不影响书架分析与购书方案。
- 支持版本：Node 18 及以上。
- 官方主页：https://nodejs.org/
- 官方文档：https://nodejs.org/docs/latest/api/
- 登录：不需要。
- 权限：无。
- 配置步骤：安装 Node LTS，确保 `node` 在 PATH 中；或用 `WEREAD_NODE` 指定绝对路径。
- 验证：`node -v` 有版本号；自检显示 `[OK] node`。
- 安全说明：`nuxt_extract.js` 在**不含 require / process** 的 vm 上下文中执行页面里的赋值语句，限 8 秒超时；输入只来自已下载到本地的 weread.qq.com 页面，脚本不联网、不读凭据。vm 不是安全沙箱，因此只在受控输入下使用，不要把它指向来源不明的 HTML。
- 轮换与撤销：系统卸载入口移除 Node。
- 核验日期：2026-10-21。

## network-weread

- 用途：所有抓取阶段都要访问微信读书域名。
- 官方主页：https://weread.qq.com/
- 官方文档：https://weread.qq.com/
- 登录：不需要（域名连通性检查不涉及账号）。
- 权限：企业网络或代理需放行该域名。
- 配置步骤：确认能打开 https://weread.qq.com/ → 代理环境放行 → 先跑一次小量抓取验证。
- 验证：访问 https://weread.qq.com/ 返回 200。
- 故障处理：超时或非 200 时做有限次重试后停止，不无限循环；命中限流时降低频率或等一段时间再跑。
- 轮换与撤销：不适用。
- 核验日期：2026-10-21。

---

## 缺依赖时怎么办

| 情况 | 能继续做什么 | 不能做什么 |
|---|---|---|
| 缺 node | 书架分析、价值评分、四方案、零钱模式、Excel/HTML | 书单总览与书单明细 |
| 未登录 / 登录过期 | 用已有数据出报告 | 刷新书架与价格 |
| 缺 matplotlib | 除图表外全部 | `fig_v5_interaction.png` |
| Python 版本或位数不达标 | 无 | 全流程 |
| 域名不可达 | 用已有数据出报告 | 一切抓取 |

任何降级都必须在结果里标注：数据停在哪次抓取、哪部分没生成。不要用估计值填补缺失字段。
