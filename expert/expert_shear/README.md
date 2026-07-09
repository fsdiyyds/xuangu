# Expert 抗剪验算自动化

## 准备

1. 打开 **EXPERT RC - Shear and torsion** → **Design** 页
2. 将 `force_aggregate_result.xlsx` 的 **抗剪** sheet 复制为 `input_VN.xlsx`（或从 `force_aggregate` 导出），并补充列：
   - `n1`, `d_bar1`, `n2`, `d_bar2`（箍筋根数、直径 mm）
3. 运行：

```bash
cd expert/expert_shear
python expert_try_shear.py
```

## 输入 `input_VN.xlsx`

| 列 | 说明 |
|----|------|
| 区域、编号 | 与 force_aggregate 抗剪表一致 |
| 名称 | ULS MAX / ALS MAX 等 → Expert Type 取首词 ULS/ALS |
| 类型 | N / Qz（同行可兼有 N、Qz 数值） |
| N(kN)、Qz(kNm) | 轴力、剪力 |
| b_cm、h_cm、d1_cm | 截面与保护层（**仅 d1 填入界面 d**） |
| n1、d_bar1、n2、d_bar2 | 箍筋 Cadres：`2φ16+10φ12`；**先填 n2 再选 d_bar2**（φ2 在 n2 输入后才启用） |

## 荷载填写（Excel → Expert）

抗剪 sheet 每个点位通常 **8 行 Excel**（4 名称 × N/Qz 各一行）：
- 程序按 **8 步** 顺序填写（先 N 后 Qz，依 Excel 行序）
- 同名 N/Qz **合并到 Expert 同一行** → 最终 **4 行**（ULS/ALS MAX/MIN，每行含 V+N）
- Expert 界面仅显示 4 行，第 5 行起自动翻页（Down 无效时尝试展开按钮）

日志示例：

```
→ Excel 8 行 → 8 次填写（N/Qz）→ Expert 4 行荷载
→ Expert 第 1 行 N=1161.14 (ULS MAX, 步骤 1/8)
→ Expert 第 1 行 V=-659.59 (ULS MAX, 步骤 2/8)
...
```

## 荷载表翻页（抗剪专用）

抗剪 **Loads (kN)** 表格界面一次约 **4 行可见**（仅决定翻页方式，**不限制** Excel 荷载组数；读多少填多少）。

| 逻辑行 | 填表方式 |
|--------|----------|
| 第 1～4 行 | 直接点击；**第 4 行**（最底可见行）用物理点击 + 原位提交 |
| 第 5 行及以后 | ① 物理点**第 4 行** Load type → ② **SendInput Down**（真实按键，非 SendMessage）→ ③ 同 Y 填 Type/V/N |

翻页原理：Down 后第 1 行隐藏，窗口仍 4 行，**新空行出现在最底行**（Y 与翻页前第 4 行相同）。若第 5 组仍覆盖第 4 行，说明 Down 未生效，请检查 Expert 窗口是否在前台。

## 输出

- `notes/{区域_编号}.rtf` — 各构件 Note 计算书
- **`result_shear_summary.xlsx`** — 结果汇总表（sheet「结果汇总」），每构件一行：

| 列 | 含义 |
|----|------|
| 区域、编号 | 构件标识 |
| b/h/d、n1/φ1、n2/φ2 | 输入截面与箍筋 |
| 荷载组数、荷载类型 | ULS MAX 等组合 |
| V_control、N_control | 控制剪力工况 |
| St、St,max | Expert 计算间距 (cm) |
| 计算状态、结果摘要 | OK/NON OK 及 Results 区首行摘要 |
| RTF路径 | 对应 Note 文件 |

## 荷载表坐标校准

在 Expert **Design** 页打开抗剪窗口后：

```bash
python calibrate_load_shear.py
python calibrate_load_shear.py --test
```

若荷载表已校准、仅需补录箍筋 4 个输入框：

```bash
python calibrate_load_shear.py --stirrup-only
```

校准顺序：荷载表 6 点 → 箍筋 **n1 / φ1 / n2 / φ2** 4 点（坐标写入 `load_coords_shear.json`）。  
校准时 **φ2 框需先在界面填 n2≠0 使其启用**，再记录第 4 点。

或：

```bash
python expert_try_shear.py --calibrate
python expert_try_shear.py --test-coords
```

依次校准：行号、Type、**V**、**N**、第 2 行行高、右侧展开按钮 → 生成 `load_coords_shear.json`。

## 荷载表坐标（手动）

默认偏移若填表错位，请用上述校准脚本，勿手改 JSON。
