const $ = (sel) => document.querySelector(sel);

const els = {
  repoUrl: $("#repoUrl"),
  repoHint: $("#repoHint"),
  loadBtn: $("#loadBtn"),
  loadAllBtn: $("#loadAllBtn"),
  clearCacheBtn: $("#clearCacheBtn"),
  tokenInput: $("#tokenInput"),
  saveSessionBtn: $("#saveSessionBtn"),
  saveLocalBtn: $("#saveLocalBtn"),
  clearAuthBtn: $("#clearAuthBtn"),
  rateLimitText: $("#rateLimitText"),
  repoMetaText: $("#repoMetaText"),
  treeRoot: $("#treeRoot"),
  errorBox: $("#errorBox"),
  fileViewer: $("#fileViewer"),
  filePathText: $("#filePathText"),
};

const CACHE_PREFIX = "gh_repo_browser:";

function nowMs() {
  return Date.now();
}

function formatBytes(n) {
  if (typeof n !== "number" || Number.isNaN(n)) return "";
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  if (n < 1024 * 1024 * 1024) return `${(n / 1024 / 1024).toFixed(1)} MB`;
  return `${(n / 1024 / 1024 / 1024).toFixed(2)} GB`;
}

function showError(msg) {
  els.errorBox.hidden = !msg;
  els.errorBox.textContent = msg || "";
}

function setRepoHint(msg) {
  els.repoHint.textContent = msg || "";
}

function setRepoMetaText(msg) {
  els.repoMetaText.textContent = msg || "未加载";
}

function setRateLimitText(msg) {
  els.rateLimitText.textContent = msg || "未知";
}

function cacheGet(key) {
  const raw = localStorage.getItem(CACHE_PREFIX + key);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (parsed.expiresAt && parsed.expiresAt < nowMs()) {
      localStorage.removeItem(CACHE_PREFIX + key);
      return null;
    }
    return parsed.value ?? null;
  } catch {
    localStorage.removeItem(CACHE_PREFIX + key);
    return null;
  }
}

function cacheSet(key, value, ttlMs) {
  try {
    localStorage.setItem(
      CACHE_PREFIX + key,
      JSON.stringify({ value, expiresAt: nowMs() + ttlMs })
    );
  } catch {
    return;
  }
}

function cacheClearAll() {
  const keys = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k && k.startsWith(CACHE_PREFIX)) keys.push(k);
  }
  for (const k of keys) localStorage.removeItem(k);
}

function getStoredToken() {
  return sessionStorage.getItem("gh_token") || localStorage.getItem("gh_token") || "";
}

function storeToken(token, where) {
  sessionStorage.removeItem("gh_token");
  localStorage.removeItem("gh_token");
  if (!token) return;
  if (where === "local") localStorage.setItem("gh_token", token);
  else sessionStorage.setItem("gh_token", token);
}

function clearToken() {
  sessionStorage.removeItem("gh_token");
  localStorage.removeItem("gh_token");
  els.tokenInput.value = "";
}

function parseGitHubUrl(rawUrl) {
  let u;
  try {
    u = new URL(rawUrl);
  } catch {
    throw new Error("链接格式无效：不是合法的 URL");
  }
  if (u.hostname !== "github.com") {
    throw new Error("链接格式无效：仅支持 github.com 域名");
  }

  const segs = u.pathname.split("/").filter(Boolean);
  if (segs.length < 2) {
    throw new Error("链接格式无效：缺少 owner/repo");
  }
  const owner = segs[0];
  const repo = segs[1];

  const kind = segs[2] === "tree" ? "tree" : segs[2] === "blob" ? "blob" : "repo";
  const refPathParts =
    kind === "repo" ? [] : segs.slice(3).map((s) => decodeURIComponent(s));
  if (kind !== "repo" && refPathParts.length === 0) {
    throw new Error(`链接格式无效：${kind} 链接缺少分支名`);
  }

  return { owner, repo, kind, refPathParts };
}

class RequestQueue {
  constructor({ concurrency, minIntervalMs }) {
    this.concurrency = concurrency;
    this.minIntervalMs = minIntervalMs;
    this.active = 0;
    this.queue = [];
    this.lastStartMs = 0;
  }

  enqueue(fn) {
    return new Promise((resolve, reject) => {
      this.queue.push({ fn, resolve, reject });
      this._pump();
    });
  }

  _pump() {
    if (this.active >= this.concurrency) return;
    const item = this.queue.shift();
    if (!item) return;

    const start = async () => {
      this.active += 1;
      this.lastStartMs = nowMs();
      try {
        const res = await item.fn();
        item.resolve(res);
      } catch (e) {
        item.reject(e);
      } finally {
        this.active -= 1;
        this._pump();
      }
    };

    const wait = Math.max(0, this.minIntervalMs - (nowMs() - this.lastStartMs));
    if (wait > 0) setTimeout(start, wait);
    else start();
  }
}

const queue = new RequestQueue({ concurrency: 2, minIntervalMs: 180 });

const rateLimitState = {
  limit: null,
  remaining: null,
  resetEpochSec: null,
};

function updateRateLimitFromHeaders(headers) {
  const limit = headers.get("x-ratelimit-limit");
  const remaining = headers.get("x-ratelimit-remaining");
  const reset = headers.get("x-ratelimit-reset");

  if (limit) rateLimitState.limit = Number(limit);
  if (remaining) rateLimitState.remaining = Number(remaining);
  if (reset) rateLimitState.resetEpochSec = Number(reset);

  if (rateLimitState.limit != null && rateLimitState.remaining != null) {
    let extra = "";
    if (rateLimitState.remaining <= 50 && rateLimitState.resetEpochSec) {
      const resetMs = rateLimitState.resetEpochSec * 1000;
      const mins = Math.max(0, Math.ceil((resetMs - nowMs()) / 60000));
      extra = `（接近上限，约 ${mins} 分钟后重置）`;
    }
    setRateLimitText(
      `${rateLimitState.remaining}/${rateLimitState.limit} ${extra}`.trim()
    );
  }
}

async function ghRequestJson(path, { token } = {}) {
  const url = `https://api.github.com${path}`;
  return queue.enqueue(async () => {
    const headers = {
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
    };
    if (token) headers.Authorization = `Bearer ${token}`;

    const resp = await fetch(url, { headers });
    updateRateLimitFromHeaders(resp.headers);
    const text = await resp.text();
    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch {
      data = null;
    }

    if (!resp.ok) {
      const msg =
        (data && (data.message || data.error)) ||
        `${resp.status} ${resp.statusText}`;
      const docUrl = data && data.documentation_url ? `\n${data.documentation_url}` : "";
      throw new Error(`GitHub API 请求失败：${msg}${docUrl}`);
    }
    return data;
  });
}

async function ghRequestBlob(sha, { owner, repo, token }) {
  const data = await ghRequestJson(`/repos/${owner}/${repo}/git/blobs/${sha}`, { token });
  const size = data.size ?? null;
  const encoding = data.encoding ?? null;
  const content = data.content ?? null;
  return { size, encoding, content };
}

async function getRepoInfo({ owner, repo, token }) {
  return ghRequestJson(`/repos/${owner}/${repo}`, { token });
}

async function getRefSha({ owner, repo, branch, token }) {
  const data = await ghRequestJson(
    `/repos/${owner}/${repo}/git/ref/heads/${encodeURIComponent(branch)}`,
    { token }
  );
  return data.object?.sha;
}

async function getCommit({ owner, repo, sha, token }) {
  return ghRequestJson(`/repos/${owner}/${repo}/git/commits/${sha}`, { token });
}

async function getTree({ owner, repo, treeSha, token }) {
  const cacheKey = `tree:${owner}/${repo}:${treeSha}`;
  const cached = cacheGet(cacheKey);
  if (cached) return cached;

  const data = await ghRequestJson(`/repos/${owner}/${repo}/git/trees/${treeSha}`, { token });
  const tree = Array.isArray(data.tree) ? data.tree : [];
  cacheSet(cacheKey, tree, 6 * 60 * 60 * 1000);
  return tree;
}

function joinPath(a, b) {
  if (!a) return b || "";
  if (!b) return a;
  return `${a.replace(/\/+$/, "")}/${b.replace(/^\/+/, "")}`;
}

function ensureLeadingSlash(p) {
  if (!p) return "";
  return p.startsWith("/") ? p : `/${p}`;
}

function extLower(path) {
  const idx = path.lastIndexOf(".");
  if (idx === -1) return "";
  return path.slice(idx + 1).toLowerCase();
}

function isProbablyTextByExt(path) {
  const ext = extLower(path);
  if (!ext) return true;
  const textExts = new Set([
    "txt",
    "md",
    "markdown",
    "json",
    "yml",
    "yaml",
    "toml",
    "ini",
    "cfg",
    "conf",
    "xml",
    "html",
    "htm",
    "css",
    "js",
    "ts",
    "tsx",
    "jsx",
    "py",
    "go",
    "rs",
    "java",
    "kt",
    "swift",
    "c",
    "h",
    "cpp",
    "hpp",
    "m",
    "mm",
    "sh",
    "bat",
    "ps1",
    "sql",
    "dockerfile",
    "make",
    "gradle",
    "properties",
    "gitignore",
    "env",
  ]);
  return textExts.has(ext);
}

function isImageByExt(path) {
  const ext = extLower(path);
  return new Set(["png", "jpg", "jpeg", "gif", "webp", "svg", "bmp", "ico"]).has(ext);
}

function isMediaByExt(path) {
  const ext = extLower(path);
  return new Set(["mp4", "webm", "mp3", "wav", "ogg"]).has(ext);
}

function guessMime(path) {
  const ext = extLower(path);
  const map = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    webp: "image/webp",
    svg: "image/svg+xml",
    mp4: "video/mp4",
    webm: "video/webm",
    mp3: "audio/mpeg",
    wav: "audio/wav",
    ogg: "audio/ogg",
    pdf: "application/pdf",
    zip: "application/zip",
  };
  return map[ext] || "application/octet-stream";
}

function decodeBase64ToBytes(b64) {
  const clean = b64.replace(/\s+/g, "");
  const bin = atob(clean);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function bytesLooksBinary(bytes) {
  const scan = Math.min(bytes.length, 4096);
  for (let i = 0; i < scan; i++) {
    if (bytes[i] === 0) return true;
  }
  return false;
}

function clearViewer() {
  els.fileViewer.innerHTML = `<div class="viewer__empty">点击左侧文件节点查看内容</div>`;
  els.filePathText.textContent = "";
}

function setViewerLoading(path) {
  els.filePathText.textContent = path ? ensureLeadingSlash(path) : "";
  els.fileViewer.innerHTML = `<div class="viewer__empty">正在加载...</div>`;
}

function setViewerError(path, msg) {
  els.filePathText.textContent = path ? ensureLeadingSlash(path) : "";
  els.fileViewer.innerHTML = `<div class="viewer__empty">${msg}</div>`;
}

function setViewerText({ path, text, size, actions }) {
  const meta = [
    size != null ? `大小：${formatBytes(size)}` : null,
    `类型：文本`,
  ].filter(Boolean);

  const actionsHtml = (actions || [])
    .map((a) => `<a href="${a.href}" target="_blank" rel="noreferrer">${a.label}</a>`)
    .join("");

  els.filePathText.textContent = ensureLeadingSlash(path);
  els.fileViewer.innerHTML = `
    <div class="viewer__filemeta">${meta.map((m) => `<div>${m}</div>`).join("")}</div>
    <div class="viewer__actions">${actionsHtml}</div>
    <pre><code class="hljs"></code></pre>
  `;
  const codeEl = els.fileViewer.querySelector("code");
  codeEl.textContent = text;
  if (window.hljs) window.hljs.highlightElement(codeEl);
}

function setViewerBinary({ path, size, blobUrl, mime, actions }) {
  const meta = [
    size != null ? `大小：${formatBytes(size)}` : null,
    `类型：二进制`,
    mime ? `MIME：${mime}` : null,
  ].filter(Boolean);

  const actionsList = [...(actions || [])];
  if (blobUrl) actionsList.unshift({ label: "下载", href: blobUrl });

  const actionsHtml = actionsList
    .map((a) => `<a href="${a.href}" target="_blank" rel="noreferrer">${a.label}</a>`)
    .join("");

  els.filePathText.textContent = ensureLeadingSlash(path);
  els.fileViewer.innerHTML = `
    <div class="viewer__filemeta">${meta.map((m) => `<div>${m}</div>`).join("")}</div>
    <div class="viewer__actions">${actionsHtml}</div>
    <div class="viewer__empty">该文件为二进制内容，建议下载或在 GitHub 打开。</div>
  `;
}

function setViewerImage({ path, size, blobUrl, actions }) {
  const meta = [
    size != null ? `大小：${formatBytes(size)}` : null,
    `类型：图片`,
  ].filter(Boolean);

  const actionsHtml = (actions || [])
    .map((a) => `<a href="${a.href}" target="_blank" rel="noreferrer">${a.label}</a>`)
    .join("");

  els.filePathText.textContent = ensureLeadingSlash(path);
  els.fileViewer.innerHTML = `
    <div class="viewer__filemeta">${meta.map((m) => `<div>${m}</div>`).join("")}</div>
    <div class="viewer__actions">${actionsHtml}</div>
    <div><img src="${blobUrl}" alt="" /></div>
  `;
}

function setViewerMedia({ path, size, blobUrl, kind, actions }) {
  const meta = [
    size != null ? `大小：${formatBytes(size)}` : null,
    `类型：媒体`,
  ].filter(Boolean);

  const actionsHtml = (actions || [])
    .map((a) => `<a href="${a.href}" target="_blank" rel="noreferrer">${a.label}</a>`)
    .join("");

  const tag = kind === "video" ? "video" : "audio";
  const attrs = kind === "video" ? 'controls playsinline' : "controls";

  els.filePathText.textContent = ensureLeadingSlash(path);
  els.fileViewer.innerHTML = `
    <div class="viewer__filemeta">${meta.map((m) => `<div>${m}</div>`).join("")}</div>
    <div class="viewer__actions">${actionsHtml}</div>
    <${tag} ${attrs} src="${blobUrl}"></${tag}>
  `;
}

const appState = {
  owner: null,
  repo: null,
  branch: null,
  basePath: "",
  baseTreeSha: null,
  token: "",
  rootNodeEl: null,
  shaToObjectUrl: new Map(),
};

function revokeObjectUrls() {
  for (const url of appState.shaToObjectUrl.values()) URL.revokeObjectURL(url);
  appState.shaToObjectUrl.clear();
}

function getGitHubWebUrl({ owner, repo, branch, path }) {
  const safePath = path ? path.split("/").map(encodeURIComponent).join("/") : "";
  return `https://github.com/${encodeURIComponent(owner)}/${encodeURIComponent(
    repo
  )}/blob/${encodeURIComponent(branch)}/${safePath}`;
}

async function initRepoFromUrl(rawUrl) {
  showError("");
  revokeObjectUrls();
  clearViewer();
  els.treeRoot.innerHTML = "";
  setRepoMetaText("加载中...");

  const token = els.tokenInput.value.trim() || getStoredToken();
  appState.token = token;

  const parsed = parseGitHubUrl(rawUrl);
  const repoInfo = await getRepoInfo({
    owner: parsed.owner,
    repo: parsed.repo,
    token,
  });

  let branch = repoInfo.default_branch;
  let initialPath = "";
  let initialPathIsFile = false;
  let refSha = null;

  if (parsed.kind === "repo") {
    branch = repoInfo.default_branch;
  } else {
    const parts = parsed.refPathParts;
    let resolved = null;
    for (let i = parts.length; i >= 1; i--) {
      const candidate = parts.slice(0, i).join("/");
      try {
        const sha = await getRefSha({
          owner: parsed.owner,
          repo: parsed.repo,
          branch: candidate,
          token,
        });
        if (sha) {
          resolved = { branch: candidate, refSha: sha, pathParts: parts.slice(i) };
          break;
        }
      } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        if (msg.includes("Not Found") || msg.includes("404")) continue;
        throw e;
      }
    }
    if (!resolved) {
      throw new Error(
        "无法解析分支名：请确认链接中的分支是否存在（若分支名包含 /，需使用 GitHub 的 tree/blob 链接格式）"
      );
    }

    branch = resolved.branch;
    refSha = resolved.refSha;
    initialPath = resolved.pathParts.join("/");
    initialPathIsFile = parsed.kind === "blob";
  }

  if (!branch) throw new Error("无法确定目标分支（default_branch 为空）");

  if (!refSha) {
    refSha = await getRefSha({ owner: parsed.owner, repo: parsed.repo, branch, token });
  }
  if (!refSha) throw new Error("无法获取分支引用（ref sha 为空）");

  const commit = await getCommit({
    owner: parsed.owner,
    repo: parsed.repo,
    sha: refSha,
    token,
  });
  const rootTreeSha = commit.tree?.sha;
  if (!rootTreeSha) throw new Error("无法获取根目录 tree sha");

  let baseTreeSha = rootTreeSha;
  let basePath = "";

  const dirPath = initialPathIsFile
    ? initialPath.split("/").slice(0, -1).join("/")
    : initialPath;
  const targetFileName = initialPathIsFile
    ? initialPath.split("/").filter(Boolean).slice(-1)[0] || ""
    : "";

  if (dirPath) {
    const segs = dirPath.split("/").filter(Boolean);
    for (const seg of segs) {
      const entries = await getTree({
        owner: parsed.owner,
        repo: parsed.repo,
        treeSha: baseTreeSha,
        token,
      });
      const next = entries.find((e) => e.type === "tree" && e.path === seg);
      if (!next) {
        throw new Error(`初始路径不存在或不是目录：${dirPath}`);
      }
      baseTreeSha = next.sha;
      basePath = joinPath(basePath, seg);
    }
  }

  appState.owner = parsed.owner;
  appState.repo = parsed.repo;
  appState.branch = branch;
  appState.baseTreeSha = baseTreeSha;
  appState.basePath = basePath;

  setRepoMetaText(`${parsed.owner}/${parsed.repo}@${branch}${basePath ? `:${basePath}` : ""}`);
  setRepoHint("");

  const rootLabel = basePath ? basePath.split("/").filter(Boolean).slice(-1)[0] : parsed.repo;
  const rootNode = createDirNode({
    name: rootLabel,
    fullPath: basePath,
    treeSha: baseTreeSha,
    expanded: true,
  });
  const rootUl = document.createElement("ul");
  rootUl.appendChild(rootNode.li);
  els.treeRoot.appendChild(rootUl);

  await expandDirNode(rootNode);

  if (targetFileName) {
    const fileLi = els.treeRoot.querySelector(
      `li[data-kind="file"][data-path="${CSS.escape(joinPath(basePath, targetFileName))}"]`
    );
    if (fileLi) {
      await openFile({ path: fileLi.dataset.path, sha: fileLi.dataset.sha });
    }
  }
}

function createDirNode({ name, fullPath, treeSha, expanded }) {
  const li = document.createElement("li");
  li.dataset.kind = "dir";
  li.dataset.path = fullPath || "";
  li.dataset.sha = treeSha;
  li.dataset.loaded = "0";
  li.dataset.expanded = expanded ? "1" : "0";

  const row = document.createElement("div");
  row.className = "node";
  row.dataset.action = "toggle";

  const icon = document.createElement("div");
  icon.className = "node__icon";
  icon.textContent = expanded ? "▾" : "▸";

  const nameEl = document.createElement("div");
  nameEl.className = "node__name";
  nameEl.textContent = name;

  const meta = document.createElement("div");
  meta.className = "node__meta";
  meta.textContent = "目录";

  row.appendChild(icon);
  row.appendChild(nameEl);
  row.appendChild(meta);

  const children = document.createElement("ul");
  children.hidden = !expanded;

  li.appendChild(row);
  li.appendChild(children);

  return { li, row, icon, children };
}

function createFileNode({ name, fullPath, sha, size }) {
  const li = document.createElement("li");
  li.dataset.kind = "file";
  li.dataset.path = fullPath || "";
  li.dataset.sha = sha;

  const row = document.createElement("div");
  row.className = "node";
  row.dataset.action = "open-file";

  const icon = document.createElement("div");
  icon.className = "node__icon";
  icon.textContent = "•";

  const nameEl = document.createElement("div");
  nameEl.className = "node__name";
  nameEl.textContent = name;

  const meta = document.createElement("div");
  meta.className = "node__meta";
  meta.textContent = size != null ? formatBytes(size) : "文件";

  row.appendChild(icon);
  row.appendChild(nameEl);
  row.appendChild(meta);
  li.appendChild(row);

  return { li, row };
}

async function expandDirNode(node) {
  const li = node.li;
  const isLoaded = li.dataset.loaded === "1";
  if (!isLoaded) {
    const owner = appState.owner;
    const repo = appState.repo;
    const token = appState.token;
    const treeSha = li.dataset.sha;

    const entries = await getTree({ owner, repo, treeSha, token });
    const dirs = [];
    const files = [];
    for (const e of entries) {
      if (e.type === "tree") dirs.push(e);
      else if (e.type === "blob") files.push(e);
    }
    dirs.sort((a, b) => a.path.localeCompare(b.path));
    files.sort((a, b) => a.path.localeCompare(b.path));

    const base = li.dataset.path || "";
    for (const d of dirs) {
      const child = createDirNode({
        name: d.path,
        fullPath: joinPath(base, d.path),
        treeSha: d.sha,
        expanded: false,
      });
      node.children.appendChild(child.li);
    }
    for (const f of files) {
      const child = createFileNode({
        name: f.path,
        fullPath: joinPath(base, f.path),
        sha: f.sha,
        size: typeof f.size === "number" ? f.size : null,
      });
      node.children.appendChild(child.li);
    }
    li.dataset.loaded = "1";
  }

  li.dataset.expanded = "1";
  node.icon.textContent = "▾";
  node.children.hidden = false;
}

function collapseDirNode(node) {
  node.li.dataset.expanded = "0";
  node.icon.textContent = "▸";
  node.children.hidden = true;
}

async function openFile({ path, sha }) {
  showError("");
  setViewerLoading(path);

  const owner = appState.owner;
  const repo = appState.repo;
  const token = appState.token;

  const actions = [
    {
      label: "在 GitHub 打开",
      href: getGitHubWebUrl({ owner, repo, branch: appState.branch, path }),
    },
  ];

  let blob;
  try {
    blob = await ghRequestBlob(sha, { owner, repo, token });
  } catch (e) {
    setViewerError(path, e.message || String(e));
    return;
  }

  const size = blob.size ?? null;
  if (size != null && size > 8 * 1024 * 1024 && !isProbablyTextByExt(path)) {
    setViewerBinary({ path, size, mime: guessMime(path), actions });
    return;
  }

  if (blob.encoding !== "base64" || typeof blob.content !== "string") {
    setViewerError(path, "无法解析文件内容（encoding 非 base64）");
    return;
  }

  const bytes = decodeBase64ToBytes(blob.content);
  const binary = bytesLooksBinary(bytes);
  const mime = guessMime(path);

  if (!binary && isProbablyTextByExt(path)) {
    if (size != null && size > 2 * 1024 * 1024) {
      setViewerError(path, "文件过大（>2MB），为避免卡顿请在 GitHub 侧查看或下载。");
      return;
    }
    const text = new TextDecoder("utf-8", { fatal: false }).decode(bytes);
    setViewerText({ path, text, size, actions });
    return;
  }

  const blobObj = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blobObj);
  appState.shaToObjectUrl.set(sha, url);

  if (isImageByExt(path)) {
    setViewerImage({ path, size, blobUrl: url, actions });
    return;
  }
  if (isMediaByExt(path)) {
    const kind = new Set(["mp4", "webm"]).has(extLower(path)) ? "video" : "audio";
    setViewerMedia({ path, size, blobUrl: url, kind, actions });
    return;
  }

  setViewerBinary({ path, size, blobUrl: url, mime, actions });
}

async function loadAllUnderCurrentRoot() {
  showError("");
  const owner = appState.owner;
  const repo = appState.repo;
  const token = appState.token;
  if (!owner || !repo || !appState.baseTreeSha) {
    showError("请先加载仓库目录");
    return;
  }

  const rootLi = els.treeRoot.querySelector('li[data-kind="dir"]');
  if (!rootLi) return;

  const stack = [rootLi];
  while (stack.length > 0) {
    const li = stack.pop();
    const expanded = li.dataset.expanded === "1";
    const loaded = li.dataset.loaded === "1";

    const icon = li.querySelector(".node__icon");
    const childrenUl = li.querySelector("ul");
    const node = { li, icon, children: childrenUl };
    if (!expanded) {
      li.dataset.expanded = "1";
      childrenUl.hidden = false;
      icon.textContent = "▾";
    }
    if (!loaded) {
      const treeSha = li.dataset.sha;
      const entries = await getTree({ owner, repo, treeSha, token });
      const dirs = entries.filter((e) => e.type === "tree").sort((a, b) => a.path.localeCompare(b.path));
      const files = entries.filter((e) => e.type === "blob").sort((a, b) => a.path.localeCompare(b.path));
      const base = li.dataset.path || "";
      for (const d of dirs) {
        const child = createDirNode({
          name: d.path,
          fullPath: joinPath(base, d.path),
          treeSha: d.sha,
          expanded: true,
        });
        node.children.appendChild(child.li);
      }
      for (const f of files) {
        const child = createFileNode({
          name: f.path,
          fullPath: joinPath(base, f.path),
          sha: f.sha,
          size: typeof f.size === "number" ? f.size : null,
        });
        node.children.appendChild(child.li);
      }
      li.dataset.loaded = "1";
    }

    const dirChildren = Array.from(li.querySelectorAll(":scope > ul > li[data-kind='dir']"));
    for (const child of dirChildren) stack.push(child);
  }
}

function bindEvents() {
  els.loadBtn.addEventListener("click", async () => {
    try {
      const url = els.repoUrl.value.trim();
      if (!url) throw new Error("请输入 GitHub 链接");
      await initRepoFromUrl(url);
    } catch (e) {
      showError(e.message || String(e));
      setRepoMetaText("未加载");
    }
  });

  els.loadAllBtn.addEventListener("click", async () => {
    try {
      await loadAllUnderCurrentRoot();
    } catch (e) {
      showError(e.message || String(e));
    }
  });

  els.clearCacheBtn.addEventListener("click", () => {
    cacheClearAll();
    showError("");
    setRepoHint("缓存已清除");
  });

  els.saveSessionBtn.addEventListener("click", async () => {
    const token = els.tokenInput.value.trim();
    storeToken(token, "session");
    showError("");
    try {
      if (token) await ghRequestJson("/user", { token });
      setRepoHint(token ? "授权有效（会话存储）" : "已清空会话授权");
    } catch (e) {
      showError(e.message || String(e));
    }
  });

  els.saveLocalBtn.addEventListener("click", async () => {
    const token = els.tokenInput.value.trim();
    storeToken(token, "local");
    showError("");
    try {
      if (token) await ghRequestJson("/user", { token });
      setRepoHint(token ? "授权有效（本地存储）" : "已清空本地授权");
    } catch (e) {
      showError(e.message || String(e));
    }
  });

  els.clearAuthBtn.addEventListener("click", () => {
    clearToken();
    showError("");
    setRepoHint("已取消授权并清除本地存储的令牌");
  });

  els.treeRoot.addEventListener("click", async (ev) => {
    const row = ev.target.closest(".node");
    if (!row) return;
    const li = row.closest("li");
    if (!li) return;

    const kind = li.dataset.kind;
    if (kind === "dir") {
      const icon = row.querySelector(".node__icon");
      const children = li.querySelector("ul");
      const node = { li, icon, children };
      const expanded = li.dataset.expanded === "1";
      try {
        if (expanded) collapseDirNode(node);
        else await expandDirNode(node);
      } catch (e) {
        showError(e.message || String(e));
      }
      return;
    }

    if (kind === "file") {
      const path = li.dataset.path;
      const sha = li.dataset.sha;
      await openFile({ path, sha });
    }
  });

  els.repoUrl.addEventListener("input", () => {
    const v = els.repoUrl.value.trim();
    if (!v) {
      setRepoHint("");
      return;
    }
    try {
      const p = parseGitHubUrl(v);
      const extra =
        p.kind === "repo"
          ? ""
          : `（${p.kind}：${p.refPathParts.join("/")}）`;
      const msg = `${p.owner}/${p.repo}${extra}`;
      setRepoHint(msg);
    } catch (e) {
      setRepoHint(e.message || String(e));
    }
  });
}

function init() {
  els.tokenInput.value = getStoredToken();
  bindEvents();
  setRateLimitText("未知");
  clearViewer();
}

init();
