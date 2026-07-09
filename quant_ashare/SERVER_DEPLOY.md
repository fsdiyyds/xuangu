# 开源服务器部署教程：每日自动训练 + 推送结果

本教程带你把 `quant_ashare` 部署到**免费/开源云端**，实现：

1. **每个交易日自动**拉取最新 A 股成交数据  
2. **自动训练**（原生模型 + Qlib 风格模型组合）  
3. **自动推送结果**到仓库 `output/latest/`，并在 Streamlit 网页查看  

推荐组合：**GitHub Actions（训练） + Streamlit Cloud（展示）**，全程免费。

---

## 目录

1. [架构说明](#一架构说明)
2. [准备：推送到 GitHub](#二准备推送到-github)
3. [GitHub Actions 每日自动训练](#三github-actions-每日自动训练)
4. [结果推送与查看](#四结果推送与查看)
5. [Streamlit Cloud 可视化](#五streamlit-cloud-可视化)
6. [邮件 / 微信推送（可选）](#六邮件--微信推送可选)
7. [自建 VPS / Docker（可选）](#七自建-vps--docker可选)
8. [日常操作清单](#八日常操作清单)
9. [常见问题](#九常见问题)

---

## 一、架构说明

```
┌─────────────────────┐     每天 18:00(北京)      ┌──────────────────────┐
│  新浪行情 API        │ ◄────────────────────── │  GitHub Actions      │
│  (QUANT_DATA_SOURCE │     增量拉取最新成交      │  ubuntu-latest       │
│   =sina)            │                          │  b1_lstm_daily.py    │
└─────────────────────┘                          └──────────┬───────────┘
                                                            │
                                                            ▼
                                                 commit → output/latest/
                                                 top50_latest.csv
                                                 report_latest.md
                                                 backtest_metrics.json
                                                            │
                                                            ▼
                                                 ┌──────────────────────┐
                                                 │  Streamlit Cloud     │
                                                 │  浏览器看推荐/K线     │
                                                 └──────────────────────┘
```

| 组件 | 作用 | 费用 |
|------|------|------|
| GitHub Actions | 定时训练、写回结果 | 公开仓库免费 |
| Streamlit Cloud | Web 展示与手动试跑 | 免费档可用 |
| 新浪数据源 | 云端可访问的行情 | 免费 |

---

## 二、准备：推送到 GitHub

### 2.1 创建仓库

1. 打开 https://github.com/new  
2. Repository name：`waterRPA`（或任意名）  
3. 选 **Public**（公开仓库 Actions 免费）  
4. **不要**勾选 README  
5. Create repository  

### 2.2 本地推送

```powershell
cd c:\Users\fenglei\PycharmProjects\pythonProject\waterRPA

# 若尚未配置远程
git remote add origin https://github.com/你的用户名/waterRPA.git

git add quant_ashare .github/workflows/daily_b1_lstm.yml
git commit -m "feat: 最新行情刷新 + Qlib 多模型 + 每日自动化"
git branch -M main
git push -u origin main
```

> 密码处填 **Personal Access Token**（GitHub → Settings → Developer settings → PAT → 勾选 `repo` + `workflow`）。

### 2.3 确认工作流文件存在

仓库中应有：

```
.github/workflows/daily_b1_lstm.yml
quant_ashare/b1_lstm_daily.py
quant_ashare/config/cloud_settings.yaml
quant_ashare/streamlit_app.py
```

---

## 三、GitHub Actions 每日自动训练

### 3.1 启用 Actions

1. 打开仓库 → **Actions**  
2. 若提示 Enable，点击启用  
3. 左侧找到 **B1 LSTM Daily Pick**  

### 3.2 定时规则

工作流默认：

```yaml
cron: "0 10 * * 1-5"   # UTC 10:00 = 北京时间 18:00，周一~周五
```

| 你想改成 | cron（UTC） |
|----------|-------------|
| 北京 17:00 | `0 9 * * 1-5` |
| 北京 20:00 | `0 12 * * 1-5` |
| 含周六 | `0 10 * * 1-6` |

修改后 `git push` 即可。

### 3.3 手动跑一次（强烈建议先做）

1. Actions → **B1 LSTM Daily Pick** → **Run workflow**  
2. 可选参数：  
   - `max_stocks`：如 `300`（调试用，全量约 2500）  
   - `models`：如 `b1,b2,lgb,lstm,ridge`  
3. 等待 30~90 分钟（取决于股票数与模型）  
4. 成功后绿色勾  

### 3.4 训练时做了什么

每次运行会：

1. `force_refresh=true`：**增量拉取最新历史成交**并合并缓存  
2. 按 `config/cloud_settings.yaml` 启用模型（默认含 B1/B2/LGB/LSTM/Ridge）  
3. 训练 → 回测 → 输出 Top50  
4. 写入 `quant_ashare/output/latest/`  
5. **自动 commit + push** 到仓库（消息含 `[skip ci]`，不会死循环）  
6. 上传 Artifacts 保留 30 天  

### 3.5 调整云端模型组合

编辑 `quant_ashare/config/cloud_settings.yaml`：

```yaml
ensemble:
  enabled:
    b1: true
    b2: true
    lgb: true
    lstm: true
    ridge: true      # Qlib 风格
    xgb: false       # 需 xgboost（requirements 已含）
    gru: false       # 时序，更耗时
  weights:
    b1: 0.20
    b2: 0.15
    lgb: 0.20
    lstm: 0.25
    ridge: 0.20
```

或手动触发时填 `models=b1,lgb,xgb,gru`。

---

## 四、结果推送与查看

### 4.1 仓库内文件

训练成功后，仓库会出现/更新：

| 路径 | 说明 |
|------|------|
| `quant_ashare/output/latest/top50_latest.csv` | 今日 Top 推荐 |
| `quant_ashare/output/latest/b1_pool_latest.csv` | B1 初选池 |
| `quant_ashare/output/latest/report_latest.md` | Markdown 报告 |
| `quant_ashare/output/latest/backtest_metrics.json` | 回测指标 |
| `quant_ashare/output/latest/data_asof.txt` | 行情截至日期 |
| `quant_ashare/output/latest/model_config.json` | 本次模型组合 |

### 4.2 本地下载

```powershell
git pull
# 打开
# quant_ashare\output\latest\top50_latest.csv
# quant_ashare\output\latest\report_latest.md
```

### 4.3 Actions Artifacts

Run 详情页 → **Artifacts** → 下载 `b1-lstm-results`（含完整 `output/b1_lstm/`）。

---

## 五、Streamlit Cloud 可视化

### 5.1 部署

1. 打开 https://share.streamlit.io ，用 GitHub 登录  
2. **Create app**  
3. 填写：  

| 项 | 值 |
|----|-----|
| Repository | `你的用户名/waterRPA` |
| Branch | `main` |
| Main file path | `quant_ashare/streamlit_app.py` |
| Python version | `3.10` |

4. **Advanced settings → Secrets**：

```toml
QUANT_DATA_SOURCE = "sina"
```

5. Deploy（首次约 10~15 分钟，需装 TensorFlow）

### 5.2 页面功能

| 标签 | 功能 |
|------|------|
| 推荐列表 | Top 推荐、回测指标、购买理由 |
| 个股可视化 | K 线 + LSTM 回测图 |
| 模型组合 & 运行 | 勾选原生/Qlib 模型、设权重、训练 |
| 数据管理 | 查看/刷新最新成交缓存 |

侧边栏可单独点 **「立即拉取最新成交数据」**。

### 5.3 与 Actions 联动

```
每天 18:00 Actions 训练并 push latest
        ↓
Streamlit 自动读仓库最新 CSV
        ↓
你打开网页即可看今日推荐
```

本地预览：

```powershell
cd quant_ashare
$env:QUANT_DATA_SOURCE="sina"
streamlit run streamlit_app.py
```

---

## 六、邮件 / 微信推送（可选）

### 6.1 邮件（GitHub Actions + SMTP）

在仓库 **Settings → Secrets and variables → Actions** 添加：

- `MAIL_USERNAME`  
- `MAIL_PASSWORD`（授权码）  
- `MAIL_TO`  

然后在 `.github/workflows/daily_b1_lstm.yml` 末尾增加一步（示例）：

```yaml
      - name: Email report
        if: success()
        env:
          MAIL_USERNAME: ${{ secrets.MAIL_USERNAME }}
          MAIL_PASSWORD: ${{ secrets.MAIL_PASSWORD }}
          MAIL_TO: ${{ secrets.MAIL_TO }}
        working-directory: quant_ashare
        run: |
          python - <<'PY'
          import os, smtplib
          from email.mime.text import MIMEText
          from pathlib import Path
          body = Path("output/latest/report_latest.md").read_text(encoding="utf-8")
          msg = MIMEText(body, "plain", "utf-8")
          msg["Subject"] = "每日量化选股报告"
          msg["From"] = os.environ["MAIL_USERNAME"]
          msg["To"] = os.environ["MAIL_TO"]
          with smtplib.SMTP_SSL("smtp.qq.com", 465) as s:
              s.login(os.environ["MAIL_USERNAME"], os.environ["MAIL_PASSWORD"])
              s.send_message(msg)
          print("mail sent")
          PY
```

> QQ 邮箱需开启 SMTP 并使用授权码；其他邮箱改 `smtp` 主机即可。

### 6.2 企业微信 / Server酱（可选）

用 webhook 推送 `top50_latest.csv` 前 10 行摘要即可，逻辑与邮件类似：在 Actions 最后一步 `curl` 你的 webhook。

---

## 七、自建 VPS / Docker（可选）

适合有自己的 Linux 服务器（如 Oracle Cloud 免费档、校园机）。

### 7.1 Docker 跑 Streamlit

```bash
cd quant_ashare
docker build -t quant-ashare .
docker run -d -p 8501:8501 \
  -e QUANT_DATA_SOURCE=sina \
  -v $(pwd)/output:/app/output \
  -v $(pwd)/data:/app/data \
  quant-ashare
```

浏览器打开 `http://服务器IP:8501`。

### 7.2 crontab 每日训练

```bash
# 编辑 crontab
crontab -e

# 北京时间工作日 18:05（服务器需设为 Asia/Shanghai 或换算 UTC）
5 18 * * 1-5 cd /opt/waterRPA/quant_ashare && \
  QUANT_DATA_SOURCE=sina /usr/bin/python3 -u b1_lstm_daily.py \
  --config config/cloud_settings.yaml --force-refresh --skip-backtest-gate \
  >> /var/log/quant_daily.log 2>&1
```

训练完后可用 `scp` / `rclone` / `git push` 把 `output/latest/` 同步到你本机或对象存储。

### 7.3 systemd 定时（更稳）

`/etc/systemd/system/quant-daily.service`：

```ini
[Unit]
Description=Quant AShare Daily Train

[Service]
Type=oneshot
WorkingDirectory=/opt/waterRPA/quant_ashare
Environment=QUANT_DATA_SOURCE=sina
Environment=TF_CPP_MIN_LOG_LEVEL=2
ExecStart=/usr/bin/python3 -u b1_lstm_daily.py --config config/cloud_settings.yaml --force-refresh --skip-backtest-gate
```

`/etc/systemd/system/quant-daily.timer`：

```ini
[Unit]
Description=Run quant daily on weekdays 18:00

[Timer]
OnCalendar=Mon..Fri 18:00:00
Persistent=true

[Install]
WantedBy=timers.target
```

```bash
sudo systemctl enable --now quant-daily.timer
sudo systemctl list-timers | grep quant
```

---

## 八、日常操作清单

### 每天（自动化后）

1. 等 Actions 跑完（或看邮件）  
2. 打开 Streamlit 网页 → **推荐列表**  
3. 需要时进 **个股可视化** 看 K 线  

### 改模型组合

1. 改 `config/cloud_settings.yaml` 的 `ensemble.enabled/weights`  
2. `git push`  
3. 次日自动生效；或手动 Run workflow  

### 本地调试

```powershell
cd quant_ashare
$env:QUANT_DATA_SOURCE="sina"
pip install -r requirements.txt

# 只刷最新数据
python -c "from data_fetcher import refresh_market_data; print(refresh_market_data(max_stocks=100)['date'].max())"

# 小规模训练
python -u b1_lstm_daily.py --max-stocks 100 --models b1,lgb,ridge --force-refresh --skip-backtest-gate

# Web
streamlit run streamlit_app.py
```

---

## 九、常见问题

**Q: Actions 超时？**  
免费 runner 约 6 小时上限，本工作流设 120 分钟。把 `cloud_settings.yaml` 的 `max_stocks` 降到 `800~1500`，或关掉 `lstm/gru`。

**Q: 本地拉不到东方财富？**  
始终设置 `QUANT_DATA_SOURCE=sina`。

**Q: Streamlit 装 TensorFlow 失败？**  
用「仅查看」：依赖 Actions 产出的 CSV；或从 `requirements.txt` 临时注释 `tensorflow`。

**Q: commit 没有更新？**  
看 Actions 日志里 `Commit latest results` 是否有变更；若结果与上次完全相同可能无 commit。

**Q: 如何确认用了最新行情？**  
看 `output/latest/data_asof.txt` 或报告里的「行情截至」；Streamlit 侧边栏也会显示。

**Q: Qlib 官方 pyqlib 要装吗？**  
不强制。本项目已内置 Qlib 风格 Zoo（Linear/Ridge/XGB/CatBoost/GRU/ALSTM/Transformer 等）。若要完整 Qlib 平台：`pip install pyqlib`。

---

## 相关文档

- [README.md](README.md) — 功能与命令  
- [GITHUB_SETUP.md](GITHUB_SETUP.md) — 建仓与 Streamlit 图文  
- [DEPLOY.md](DEPLOY.md) — Colab / Render 补充  

**免责声明**：仅供学习研究，不构成投资建议。股市有风险，投资需谨慎。
