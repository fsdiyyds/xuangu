# GitHub 建仓 + Streamlit 可视化部署完整指南

从零到在浏览器查看 **B1战法 + LSTM Top50 可视化报告**，约 20 分钟。

---

## 目录

1. [GitHub 建仓与推送](#一github-建仓与推送)
2. [GitHub Actions 每日自动选股](#二github-actions-每日自动选股)
3. [Streamlit Cloud 可视化部署（重点）](#三streamlit-cloud-可视化部署重点)
4. [本地 Streamlit 预览](#四本地-streamlit-预览)
5. [每日使用流程](#五每日使用流程)
6. [常见问题](#六常见问题)

---

## 一、GitHub 建仓与推送

### 1.1 创建空仓库

1. 打开 https://github.com/new
2. **Repository name**: `waterRPA`
3. **Public**（公开仓库，Actions / Streamlit 免费）
4. **不要**勾选 README / .gitignore
5. **Create repository**

### 1.2 本地推送

```powershell
cd c:\Users\fenglei\PycharmProjects\pythonProject\waterRPA

# 首次需配置身份
git config --global user.name "你的GitHub用户名"
git config --global user.email "你的邮箱"

git init
git add .
git commit -m "feat: B1战法+LSTM量化选股与可视化"
git branch -M main
git remote add origin https://github.com/你的用户名/waterRPA.git
git push -u origin main
```

> **密码**处填 Personal Access Token（非登录密码）：  
> GitHub → Settings → Developer settings → Personal access tokens → Generate (classic) → 勾选 `repo`

---

## 二、GitHub Actions 每日自动选股

云端自动跑选股（你本地网络不通也能用）。

1. 仓库 → **Actions** → 启用 workflows
2. 左侧 **B1 LSTM Daily Pick** → **Run workflow**
3. 等待 30~90 分钟
4. 进入该次 run → **Artifacts** → 下载 `b1-lstm-results`

**产出文件：**

| 文件 | 说明 |
|------|------|
| `top50_latest.csv` | Top50，含名称、购买理由 |
| `b1_pool_latest.csv` | B1 初选 520 池 |
| `report_latest.md` | 文字报告 |
| `report_visual.html` | **可视化 HTML 报告**（价量+KDJ+LSTM预测图） |

Actions 跑完后会自动 commit 到 `quant_ashare/output/latest/`，Streamlit 可直接读取。

---

## 三、Streamlit Cloud 可视化部署（重点）

部署完成后，手机/电脑浏览器即可查看：**推荐列表、个股 K 线图、LSTM 预测路径、购买理由**。

### 3.1 前置条件

- 代码已推送到 GitHub 公开仓库
- 建议先跑通一次 GitHub Actions（有 `output/latest/top50_latest.csv` 数据）

### 3.2 部署步骤（图文流程）

**Step 1 — 打开 Streamlit Cloud**

访问 https://share.streamlit.io ，用 **GitHub 账号**登录。

**Step 2 — 创建新应用**

点击 **Create app** → **Yup, I have an app** → **Continue**

**Step 3 — 填写仓库信息**

| 配置项 | 填写内容 |
|--------|----------|
| Repository | `你的用户名/waterRPA` |
| Branch | `main` |
| Main file path | `quant_ashare/streamlit_app.py` |
| App URL (optional) | 自定义，如 `b1-lstm-quant` |

**Step 4 — 高级设置（重要）**

点击 **Advanced settings**：

**Python version**: `3.10`

**Secrets**（TOML 格式，粘贴以下内容）：

```toml
QUANT_DATA_SOURCE = "sina"
```

**Dependencies**：Streamlit 会自动读取 `quant_ashare/requirements.txt`，无需额外配置。

**Step 5 — Deploy**

点击 **Deploy!**，首次构建约 5~15 分钟（需安装 TensorFlow，较慢）。

**Step 6 — 访问应用**

部署成功后地址类似：

```
https://b1-lstm-quant.streamlit.app
```

或：

```
https://你的应用名-用户名.streamlit.app
```

### 3.3 页面功能说明

| 标签页 | 功能 |
|--------|------|
| **推荐列表** | Top50 展开卡片：股票全名、购买理由、B1/LSTM 指标 |
| **个股可视化** | 选择股票 → 四联图（价格+MA60+BBI、成交量、换手率、KDJ-J）+ LSTM 预测虚线 + 特征热力图 |
| **运行任务** | 在线触发选股（限制 500 只，约 10 分钟） |

### 3.4 Streamlit 与 GitHub Actions 联动

推荐工作流：

```
GitHub Actions（每天 18:00 自动跑）
        ↓
结果 commit 到 output/latest/
        ↓
Streamlit Cloud 打开 → 自动显示最新 Top50
        ↓
「个股可视化」页选股票看 K 线 + LSTM 预测
```

本地同步：

```powershell
git pull
streamlit run quant_ashare/streamlit_app.py
```

### 3.5 Streamlit 部署注意事项

| 问题 | 说明 |
|------|------|
| 构建失败 / 内存不足 | TensorFlow 较大，免费档可能 OOM；可只用「查看结果」页，不在云端跑 LSTM |
| 首次打开无数据 | 先手动触发 GitHub Actions，等 commit 后再刷新 Streamlit |
| 在线运行超时 | Streamlit 免费档有执行时间限制；大批量选股请用 GitHub Actions |
| 图表不显示 | 确认 `pip install plotly` 在 requirements.txt 中 |

### 3.6 仅查看模式（轻量部署）

若 TensorFlow 导致部署失败，可临时从 `requirements.txt` 注释掉 `tensorflow`，  
Streamlit 仍可展示 GitHub Actions 已生成的 CSV 和 HTML 报告（「推荐列表」页），  
只是「运行任务」和「LSTM 热力图」不可用。

---

## 四、本地 Streamlit 预览

不部署云端，本地浏览器查看：

```powershell
cd c:\Users\fenglei\PycharmProjects\pythonProject\waterRPA\quant_ashare
pip install streamlit plotly tensorflow -r requirements.txt
$env:QUANT_DATA_SOURCE="sina"
streamlit run streamlit_app.py
```

浏览器自动打开 http://localhost:8501

---

## 五、每日使用流程

```
1. 等待 GitHub Actions 跑完（或手动 Run workflow）
2. git pull 拉取最新 top50_latest.csv
3. 打开 Streamlit 网页 → 「推荐列表」查看购买理由
4. 「个股可视化」→ 选股票看历史走势 + LSTM 预测
5. 下载 CSV 或打开 report_visual.html 离线查看
```

---

## 六、常见问题

**Q: push 403 / 认证失败**  
用 Personal Access Token，不用 GitHub 密码。

**Q: Actions 无 workflow**  
确认 `.github/workflows/daily_b1_lstm.yml` 已 push。

**Q: Streamlit Deploy 失败**  
查看 Logs；常见原因是 TensorFlow 安装超时，见 3.6 轻量模式。

**Q: 股票名显示为代码**  
云端首次需拉名称缓存；在 Actions 环境跑时会自动生成 `data/cache/stock_names.json`。

**Q: 改扫描数量**  
编辑 `quant_ashare/config/cloud_settings.yaml` → `data.max_stocks` → push。

**Q: 改定时**  
编辑 workflow 中 `cron: "0 10 * * 1-5"`（UTC 10:00 = 北京 18:00）。

---

## 相关文档

- [SERVER_DEPLOY.md](SERVER_DEPLOY.md) — **开源服务器每日自动训练 + 推送（推荐对照操作）**
- [DEPLOY.md](DEPLOY.md) — 多平台部署（Colab / Render / Docker）
- [README.md](README.md) — 功能说明与命令参考

---

**免责声明**：仅供学习研究，不构成投资建议。股市有风险，投资需谨慎。
