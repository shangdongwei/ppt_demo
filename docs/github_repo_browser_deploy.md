# GitHub 仓库目录浏览器：部署说明

本系统为“纯前端直连 GitHub API + 本地静态资源服务端”的形态：
- 服务端只负责提供静态页面与资源，不接触、不转发、不存储用户令牌；
- 所有 GitHub API 请求均由浏览器直接发起，PAT 仅保存在浏览器本地存储中。

---

## 1. 环境要求

- Python 3.10+（仅用于启动静态文件服务）
- 浏览器：Chrome / Edge / Safari / Firefox 任一现代版本

---

## 2. 启动方式

在仓库根目录运行：

```bash
python -m github_browser.backend --host 127.0.0.1 --port 8000
```

打开浏览器访问：

```text
http://127.0.0.1:8000/
```

---

## 3. 使用说明

### 3.1 加载公开仓库

1. 在“GitHub 链接”输入框粘贴：
   - `https://github.com/owner/repo`
   - 或 `https://github.com/owner/repo/tree/branch/path`
2. 点击“加载目录”
3. 展开目录、点击文件查看内容

### 3.2 授权访问私有仓库（PAT）

1. 在 GitHub 创建 PAT：
   - Classic PAT：需要 `repo` scope 才能读取私有仓库
   - Fine-grained PAT：为目标仓库授予 “Contents: Read” 等读取权限（并确保该 token 可访问该私有仓库）
2. 将 PAT 粘贴到“GitHub Token（PAT）”
3. 点击：
   - “保存到会话”：关闭标签页后自动清除（推荐）
   - 或 “保存到本地”：持久化（风险更高）
4. 再加载私有仓库链接即可浏览

### 3.3 取消授权 / 清除令牌

- 点击“取消授权”，会清除 sessionStorage 与 localStorage 中的 token。

---

## 4. 速率限制与建议

- 匿名调用：约 60 次/小时，适合按需展开目录，不建议直接“加载完整目录”
- 授权调用：约 5000 次/小时，更适合大仓库与全量遍历

页面会展示 `remaining/limit`，并在剩余额度较低时提示接近上限。

