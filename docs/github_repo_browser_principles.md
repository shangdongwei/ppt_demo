# GitHub 链接获取项目目录结构的原理解析（含“豆包式”实现要点）

本文以“输入 GitHub 仓库链接 → 获取完整目录结构 → 前端渲染 → 点击文件展示内容”为目标，系统梳理其核心技术原理，并明确实现时依赖的 GitHub 官方 API 能力、鉴权与权限校验机制、分页/截断处理、目录树转换、缓存与限流等关键细节。

---

## 1. GitHub API 调用逻辑：获取目录树的核心路径

### 1.1 REST：Git Trees API（核心）

#### 1.1.1 关键接口与参数

- 获取分支引用（得到 commit SHA）
  - `GET /repos/{owner}/{repo}/git/ref/heads/{branch}`
  - 返回关键字段：
    - `object.sha`：分支当前指向的 commit SHA

- 获取 commit（得到根 tree SHA）
  - `GET /repos/{owner}/{repo}/git/commits/{commit_sha}`
  - 返回关键字段：
    - `tree.sha`：该 commit 的根目录 tree SHA

- 获取 tree 内容（得到目录下的条目列表）
  - `GET /repos/{owner}/{repo}/git/trees/{tree_sha}`
  - 可选参数（重要）：
    - `?recursive=1`：一次性递归拉平返回整个仓库树（大仓库会截断，见 1.1.3）
  - 返回字段（核心结构）：
    - `sha`：该 tree 的 SHA
    - `truncated`：当递归结果过大时为 `true`
    - `tree[]`：条目数组，每个条目包含：
      - `path`：相对当前 tree 的路径段（非全路径）
      - `type`：`tree`（目录）/ `blob`（文件）/ `commit`（子模块）
      - `sha`：条目对象 SHA（目录/文件后续拉取用）
      - `size`：文件大小（仅 blob 可能有，且并非所有场景保证）
      - `url`：该对象 API 地址

#### 1.1.2 嵌套资源遍历规则（非 recursive 的安全做法）

“一次递归拉平”适合中小仓库，但为了覆盖大仓库与避免 `truncated`，更稳妥的做法是“按目录逐层拉取 tree”：

1. 先定位根 tree SHA。
2. 请求 `GET /git/trees/{tree_sha}` 得到当前目录下的条目。
3. 对于 `type=tree` 的条目，取其 `sha` 继续请求对应 tree。
4. 如需“完整目录结构”，对所有目录进行 BFS/DFS，直到遍历完所有子 tree。

这种方式的优点：
- 每个 tree 响应只包含“该目录的一层条目”，单次 payload 更可控；
- 不依赖 `recursive=1`，规避大仓库递归截断；
- 天然支持“懒加载”：只有用户展开目录时才请求该目录的 tree。

#### 1.1.3 超大仓库与 >100KB 的处理：`truncated` 与递归拆分策略

Git Trees API 在使用 `recursive=1` 时，当返回内容过大时会：
- `truncated: true`
- `tree[]` 只返回部分结果

因此“豆包式”要实现“完整目录结构”必须具备以下兜底逻辑：

- 优先策略：
  - 小仓库：`GET /git/trees/{root}?recursive=1`，一次拿到扁平树，减少 API 次数。
- 兜底策略（推荐默认采用）：
  - 一律使用非递归 tree 拉取（逐目录展开/遍历），彻底绕过 `truncated` 风险。
- 混合策略：
  - 先尝试 `recursive=1`；
  - 若 `truncated=true`，立刻切换为“逐目录遍历”。

逐目录遍历不依赖“响应大小阈值”，理论上可处理任意规模仓库，但必须配合：
- 请求队列与限流（见 4.2、4.3）
- 缓存与断点复用（见 4.1）

### 1.2 REST：文件内容拉取的两条常用链路

#### 1.2.1 Git Blob API（适合已知 SHA 的文件节点）

- `GET /repos/{owner}/{repo}/git/blobs/{blob_sha}`
- 返回：
  - `encoding` 常为 `base64`
  - `content` base64 文本（可能带换行）
  - `size`

优点：
- 目录树节点里已经有 `sha`，点击文件可直接拉取；
- 对私有仓库也适用（在请求里带 token 即可）。

注意：
- 超大文件不适合直接拉取与解码（前端内存与渲染会卡顿），需要“大小阈值”策略与下载入口。

#### 1.2.2 Contents API（适合按 path 拉取、或需要 download_url）

- `GET /repos/{owner}/{repo}/contents/{path}?ref={branch}`
- 返回中包含：
  - `type`：`file`/`dir`
  - `encoding`、`content`（小文件）
  - `download_url`（通常用于下载/预览）

注意：
- 用 path 拉取时，路径编码、分支 ref 都需要正确处理；
- 对私有仓库，`download_url` 的可用性与跨域取决于 GitHub 返回的 URL 形态与浏览器策略；通常更稳妥的仍是用 API 获取内容并在前端生成 Blob URL。

### 1.3 GraphQL：tree 查询（补充能力）

GraphQL 可通过 `repository.object(expression: "branch:path")` 获取对象：
- 当对象为 Tree 时，可取 `entries`（目录一层条目）
- 当对象为 Blob 时，可取 `text`（文本内容）或 `byteSize`

典型查询（目录一层）：

```graphql
query($owner:String!, $name:String!, $expr:String!) {
  repository(owner:$owner, name:$name) {
    object(expression:$expr) {
      ... on Tree {
        entries {
          name
          type
          object {
            ... on Blob { byteSize }
            ... on Tree { oid }
          }
        }
      }
    }
  }
  rateLimit { limit remaining resetAt }
}
```

关键点：
- GraphQL 更适合“按需查询一层目录”，与“懒加载”天然契合；
- 仍需递归遍历才能拿到全量目录；
- 受查询复杂度与嵌套深度限制，不建议用一次查询硬拉全量。

### 1.4 分页处理逻辑（REST Link / GraphQL Connection）

获取“目录结构”本身时：
- Git Trees API 的 `GET /git/trees/{sha}` 不使用传统分页（一个 tree 就是一层目录条目集合）；
- “超大仓库”问题主要体现为 `recursive=1` 的 `truncated`（见 1.1.3），而不是分页。

但完整系统通常还会调用其它需要分页的 GitHub API（例如列出仓库、列出 issue/pr 等），此时必须实现通用分页机制：

- REST API 分页：
  - 常用参数：`per_page`（最大 100）、`page`
  - 响应头 `Link` 会给出 `rel="next"`、`rel="last"` 等链接
  - 处理方式：循环请求 `next` 直到不存在，或达到上限/限流阈值

- GraphQL 分页（Connection）：
  - 常用参数：`first` + `after`（cursor）
  - 返回：`pageInfo { hasNextPage endCursor }`
  - 处理方式：while `hasNextPage` 为 true，携带 `after=endCursor` 继续拉取

---

## 2. 身份鉴权原理：公开仓库 vs 私有仓库

### 2.1 公开仓库（无需鉴权）

- REST API 匿名访问受严格速率限制：
  - 每小时 60 次（以 IP 维度为主）
- 适合：
  - 目录懒加载（只拉少量 tree）
  - 小仓库一次性拉平（谨慎使用）

风险：
- 任意“全量递归遍历”都可能快速耗尽 60 次额度，必须做缓存与限流。

### 2.2 私有仓库（必须鉴权）

#### 2.2.1 Personal Access Token（PAT）

两类 PAT：
- Classic PAT：需要勾选 scopes
- Fine-grained PAT：需要选择可访问的仓库与权限细粒度

目录/内容读取的最小权限要求（实践建议）：
- 私有仓库读取：Classic PAT 需 `repo`（包含私有仓库读权限）
- 公开仓库仅读：可用 `public_repo`（Classic）或 fine-grained 的 “Contents: Read”

校验令牌有效性：
- `GET /user`
  - 200：token 有效
  - 401：token 无效/过期
  - 403：可能权限不足或触发速率/风控

#### 2.2.2 OAuth 授权流程（适合面向大众用户的产品形态）

核心步骤：
1. 前端跳转 GitHub OAuth authorize：
   - `https://github.com/login/oauth/authorize?client_id=...&scope=...&redirect_uri=...&state=...`
2. 授权回调带 `code` 到你控制的后端
3. 后端用 `client_secret` 换 token
4. 前端仅保存 access token（或只保存 session），后端不落盘

注意：
- 若要求“服务端不得存储任何用户令牌”，则后端只能做“换 token 的瞬时中转”，禁止写数据库/日志；
- 若需要刷新机制：
  - GitHub OAuth token 是否可刷新取决于应用配置与 GitHub 提供的 token 类型；
  - 工程上常见做法是短期 token + 重新授权（或在浏览器 session 存活期间使用）。

### 2.3 令牌存储与刷新策略（合规优先）

强约束：“不得在服务端存储任何用户 GitHub 令牌”时：
- 令牌只能存放在浏览器本地：
  - `sessionStorage`：更安全（随标签关闭自动清除）
  - `localStorage`：更易用但风险更高（持久化，易被 XSS 窃取）
- 推荐默认：
  - 只存 `sessionStorage`；
  - 提供“取消授权/清除本地存储”按钮；
  - 严格避免把 token 写入 URL、日志、报错栈、分析埋点。

---

## 3. 目录结构扁平化/层级化转换原理

### 3.1 Git Trees API 的两种结构形态

- `recursive=1` 返回的是“扁平数组”（flat tree）：
  - 每条记录的 `path` 是相对于仓库根的完整路径（如 `src/utils/a.js`）
- 非递归 `GET /git/trees/{sha}` 返回的是“当前目录的一层 entries”：
  - 每条记录的 `path` 是相对当前目录的一段（如 `utils`、`a.js`）

### 3.2 从扁平数组构建嵌套目录树（算法）

输入：`items = [{path:"a/b/c.txt", type:"blob", size:...}, ...]`

构建过程（典型字典树 / Trie）：
1. 建立根节点 `root`
2. 对每条 item：
   - `parts = path.split("/")`
   - 逐段向下：
     - 若段节点不存在则创建目录节点
   - 最后一段创建文件节点（携带 size/type/sha）
3. 同级节点按“目录优先 + 名称排序”输出给前端渲染

复杂度：
- 时间：O(总路径段数)
- 空间：O(节点数)

### 3.3 大目录懒加载触发机制

懒加载的关键点是“节点是否已加载 children”：
- 初始只加载根目录；
- 用户点击展开某目录：
  1. 若 `loaded=false`：调用 `GET /git/trees/{dir_sha}` 拉取一层 children，写入缓存并渲染；
  2. 若 `loaded=true`：直接展开 DOM；
- 提供“加载完整目录”：
  - 对所有目录节点做 BFS/DFS，依次拉取 children（必须走请求队列与限流）。

---

## 4. 资源缓存与速率限制规避原理

### 4.1 缓存对象与建议策略

建议区分两类缓存：

1. 目录树缓存（高价值）
- Key：`tree:{owner}/{repo}:{tree_sha}` 或 `tree:{owner}/{repo}@{branch}:{path}`
- Value：该目录一层 entries（或转换后的 children）
- TTL：6~24 小时（目录结构变化相对不频繁）
- 存储：`localStorage`（或 IndexedDB）

2. 文件内容缓存（谨慎）
- 文件内容可能很大，不适合持久化写 localStorage
- 推荐：
  - 仅内存缓存（LRU，最多 N 个文件）
  - 或 sessionStorage + 严格大小上限 + 短 TTL（例如 30 分钟）

### 4.2 GitHub 速率限制识别

所有 REST 响应都会带：
- `X-RateLimit-Limit`
- `X-RateLimit-Remaining`
- `X-RateLimit-Reset`（Unix epoch seconds）

实现上应当：
- 每次响应更新 UI 的剩余额度；
- 当 remaining 低于阈值（如 50）时预警；
- 当 remaining = 0 且返回 403 时，提示用户等待到 reset。

### 4.3 请求队列与限流（避免触发风控/429）

典型策略：
- 并发限制：concurrency 1~4（浏览器端建议 2）
- 最小间隔：每次请求启动至少间隔 150~300ms
- 若执行“全量遍历”：
  - BFS/DFS 把目录 sha 加入队列
  - 队列按限流策略逐个执行
  - 可在 remaining 过低时提前中止并提示用户

---

## 5. 权限校验、错误处理与用户可理解提示

必须覆盖的典型错误：
- 仓库不存在：404
- 权限不足（私有仓库/缺 scopes）：404 或 403（GitHub 有时对私有仓库返回 404）
- token 无效：401
- 触发速率限制：403 + `X-RateLimit-Remaining: 0`
- 子模块：`type=commit`，需提示“子模块未展开或需额外处理”
- 文件过大：提示无法直接展示，提供 GitHub 打开/下载入口
- 网络异常：捕获 fetch 失败，提示重试

---

## 6. 与本项目实现的对应关系（简要）

本仓库的实现选择“逐目录 tree 拉取 + 懒加载”为默认策略，以规避 `recursive=1` 的截断问题，并在前端实现：
- GitHub 链接解析（owner/repo/branch/path）
- PAT 本地存储与校验
- 目录树交互（展开/折叠/全量加载）
- 文件内容按类型渲染（代码高亮/文本/图片/媒体/二进制下载）
- 目录 tree 缓存与请求队列限流
