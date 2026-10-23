# 数据源与字段口径

## 一、登录与浏览器

所有需要登录的接口都靠 **持久化浏览器 profile 里的 cookie**（`wr_vid` / `wr_skey` / `wr_rt` / `wr_fp` / `wr_ql` / `wr_gid`）。`login.py` 扫码后 profile 目录即可复用，脚本用 `launch_persistent_context(profile, headless=True)` 直接带态请求。

- **不要**把 cookie 落成明文文件（早期版本写过 `.weread_cookie`，是泄露点，已弃用）。
- `wr_rt` 是刷新令牌，比 `wr_skey` 更敏感；两者都随 App 退出登录或长时间不活跃失效。
- `fetch_*` 脚本启动时先检查 `wr_vid` 是否存在，缺失直接 `LOGIN_LOST` 退出——**不要**在无登录态下继续跑，否则会抓回一堆空数据。
- 浏览器通道：优先本机 Edge（`msedge`），其次 Chrome，再退回内置 chromium。由 `wconfig.channel()` 判定，找不到指定通道会自动降级。

## 二、接口清单

| 用途 | 请求 | 关键字段 |
|---|---|---|
| 账户资产 | `POST /web/pay/balance`（body 带 `pf=ios`） | `balance`（书币）/ `giftBalance`（赠币） |
| 会员卡 | `GET /web/pay/memberCardSummary?pf=ios` | 到期时间、卡类型 |
| 书架全量 | `GET /web/shelf/sync` | `books[]`、`archive[]`；每本含 `bookId` `title` `author` `paid` `payingStatus` `price` `centPrice` `soldout` `category` `publishTime` `readUpdateTime` `finishReading` |
| 书目详情 | `GET /web/book/info?bookId=` | `newRating`（平台评分）、`ratingCount`、`newRatingDetail`、`category`、`centPrice`、`soldout`、`totalWords`、`deepLink`、`encodeId` |
| 累计阅读人数 | `GET /book-detail?type=1&v=<encodeId>`（SSR，取 `window.__NUXT__`） | `readingStat.readingCount` |
| 想法数 | `GET /web/review/list?listType=2&count=1` | `totalCount` |
| 热门划线 | `GET /web/book/bestbookmarks?bookId=` | `items[].totalCount`（取 Top10 求和） |
| 阅读进度 | `GET /web/book/getProgress?bookId=` | `finishTime` / `startReadingTime` / `canFreeRead` |
| 公开书单 | `GET /misc/booklist/<id>`（免登录） | `__NUXT__` → `booklist` / `bookEntities` / `bookIds` |
| **全库搜索（分页）** | `GET /web/search/name_books?keyword=<kw>&maxIdx=0&count=40` | `books[].bookInfo` 含 `bookId` `title` `author` `publisher` `price`（元）`newRating`（×100）`newRatingCount` `soldout` `payType` `deepLink`；另有 `hasMore` `totalCount` |

搜索接口要点（`fetch_lowprice.py` 用它整理全库低价书单）：
- 返回**直接带价格与评分**，不用再逐本查 `/web/book/info`；已抽验 `price` 与 `centPrice/100` 完全一致、`newRating/10` 与详情页评分一致。
- `maxIdx` 步进 40 即可翻页（`?keyword=论语&maxIdx=40&count=40`），`hasMore=0` 即到底。
- 搜索页 `GET /web/search/books?keyword=` 是 SSR，数据在 `searchBooksStoreModule.bookInfos`（HTML 里正则切 JSON），首屏 40 条，**不能直接翻页**；要翻页必须用 `name_books`。
- `price=0` 是免费书、`price=-1` 是异常值，都要过滤；`soldout=1` 已下架也过滤。

## 三、权益类型判定（已用付费墙文案实地验证）

| `paid` | `payingStatus` | 判定 | 付费墙文案 |
|---|---|---|---|
| 1 | — | 永久可读 | 无（已拥有） |
| 0 | 2 | 付费会员卡可读 | "会员卡 19元/月 · 开通后即可阅读" |
| 0 | 1 | 体验卡可读 | "你正在使用体验卡阅读付费部分" |
| 0 | 0 | 免费 | 无 |

代码里统一用 `cat()` 函数：`P`（永久）/ `M`（会员卡）/ `T`（体验卡）/ `F`（免费）/ `X`（未知）。

## 四、字段易混点

| 字段 | 含义 | 别混成 |
|---|---|---|
| `ratingCount`（`_rcnt`） | **评分人数** | 阅读人数 |
| `readingCount`（`_readcnt`） | **累计阅读人数**（分享页 `readingStat` 才有；`/web/book/readInfo` 给的是"当前在读"） | 销量、在读人数 |
| `newRating` | 平台推荐值/评分（0–100 尺度，界面上显示为 9.0 这种） | 自己的价值评分 |
| `totalCount`（listType=2） | 想法数 | 星评人数（listType=3 才是） |

价格单位：`centPrice` 是**分**，`price` 是**元**。展示一律用 `centPrice / 100` 或 `price` 的一种，不要混用。

## 五、抓取节奏与稳定性

- **必须串行**，一次一本。并行会被限流（返回空或 403）。
- 每本之间 `sleep(0.3~0.5s)`。1000 本全量约 10–15 分钟，耗时随书架本数线性增长。
- **可断点续跑**：`fetch_all.py` 把已完成 bookId 写进 `_refresh_state.json`，中断后重跑会跳过已完成；`fetch_extra.py` 同理用 `_refresh_extra_state.json`。
- 单项失败不中断整体，记进 `fails` 列表，末尾汇总。
- 大响应（如 `mx_*` 之外的 SSR 页面）注意 `_bookmeta_v5.json` 会到 1MB 量级，用 `json.dump(..., ensure_ascii=False)` 且**先写 .tmp 再 os.replace**，避免中断留下半截文件。

## 六、接口变更时的排查顺序

1. 先跑 `python scripts/login.py` 确认登录态还有效
2. 用 `fetch_all.py --meta` 单独刷元数据（不动钱包/书架），看是否还是 `LOGIN_LOST`
3. 单本调试：`curl` 换成 playwright 的 `ctx.request.get(...)`，打印 `r.status` 与 `r.text()[:200]`
4. 若页面改版：`/book-detail` 的 SSR 结构变了要同步改 `encodeId` 提取与 `__NUXT__` 求值（`nuxt_extract.js`）
5. 建议每次抓完看一眼 `_refresh_meta.json` 里的 `fails` 数量，超过 5% 说明接口或风控有变

## 七、合规边界

- 仅抓取**自己账号可见**的数据，不做他人书架/社交数据。
- 微信读书网页端无公开 API 文档，所有接口均为逆向所得，可能随时失效。
- 数据仅用于个人阅读管理，不要二次分发原始抓取结果。
