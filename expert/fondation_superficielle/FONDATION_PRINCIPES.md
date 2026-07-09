# 9.1.6 浅基础验算原理说明

依据您提供的 **计算原理.docx**（Fascicule 62 Titre V）与样例计算书整理。

参考文档：本目录下 `计算原理.docx`

---

## 1) 计算参数

### 基础埋深 De（Profondeur d'encastrement）

**De = 地面至基础底面的竖向深度**（扩大基础），**不是**墩柱直径。

用于：
- **q′₀** = σ′v0(De) ≈ γ×D（分水位上下用 γ / γ′）
- **iδβ** 中 B/(2De) 判定 Φ1

### 极限承载力（有效应力）
```
q'u = q'0 + kp × ple*
```

### q′₀ 自动计算
```
无地下水：q'0 = γ × De
有地下水（Zw < De）：q'0 = γ × Zw + γ' × (De − Zw)
γ' 默认 = γ − 10 kN/m³（可改）
```

### ple* / Ec / Ed 自旁压试验（Fascicule 62 / DTU 13-12）

| 参数 | 深度区间（自地面，基底深度 = De） | 方法 |
|------|----------------------------------|------|
| pl* | 各试验点 | pl* = plm − σ′v0(z) |
| **ple*** | De ～ De+1.5B | pl* 算术平均（pl* ≤ 1.5×pl*min） |
| **Ec** | De ～ De+B/2 | EM 调和平均 → E1 |
| **Ed** | De ～ De+4.58B（五层） | 4/Ed = Σ1/(w_i·E_i) |

input 各墩 sheet【三、旁压试验】填 `depth_m`, `plm_MPa`, `EM_MPa`；`q0`/`ple*`/`Ec`/`Ed` 留空则自动计算。

### 承台底换填（D2/D3 砂砾石）

若基础底以下换填压实砂砾（如「remplacement de sol en matériaux D2/D3」）：

| 输入 | 含义 |
|------|------|
| `replacement_thickness_m` | 换填厚度（自基底向下，m），常取基础厚度 H |
| `replacement_plm_MPa` | 换填层旁压极限压力 plm（试验或压实控制值） |
| `ple_star_zone` | `replacement`：仅在换填厚度内（限 1.5B）用换填 plm 平均 **ple***，不计下方原位粘土 |
| | `mixed`：换填区用换填 plm，其下 1.5B 仍计原位旁压点 |

**墩柱1 手填 ple*≈3135 kPa** 与仅读原位粘土旁压（≈1436 kPa）差异大，主因是换填后承载层为 D2/D3；
在 `replacement` 模式下，用换填 plm≈3.2 MPa、厚度 1.3 m 可复现约 3135 kPa。

### 土工参数
| 符号 | 含义 |
|------|------|
| q'0 | 有效覆土压力 (kPa) |
| φ | 内摩擦角 (°) |
| C' | 有效粘聚力 (kPa)，滑动验算限 ≤75 kPa |
| γ | 土体重度 (kN/m³) |

---

## 2) 地基承载力（Mobilisation du sol）

### 2.1 Meyerhof 参考应力 q'ref1
```
B' = B - 2ex
L' = L - 2ey
q'ref1 = N / (B' × L')
```

### 2.2 线性模型边缘应力
```
q_mean = N / (B × L)
qmin = q_mean × (1 - 6ex/B - 6ey/L)
qmax = q_mean × (1 + 6ex/B + 6ey/L)
```

### 2.3 参考应力 q'ref（Fascicule 62 B.2.2.2）
```
q'ref2 = (3×qmax + qmin) / 4
q'ref = max(q'ref1, q'ref2)
```
Meyerhof 模型下亦可取均匀应力平衡荷载。

### 2.4 偏心折减 iδβ（Fascicule 62 Titre V Annexe F.1）

**按土质选用公式：**

| 土质 | iδβ | 说明 |
|------|-----|------|
| 粘性土（argile / coherent） | Φ1(δ) = (1 − 2δ/π)² | δ 为荷载相对竖向的倾角 |
| 砂性土（sable / frictional） | Φ2(δ, De, B) | δ < π/4 与 δ ≥ π/4 两式，含 e^(−De/B) |
| 混合土 | Φ2 + (Φ1−Φ2)·(1−exp(−0,6c/(γB tanφ))) | NF P94-261，α=0,6 |

**δ 的取值（竖向偏心荷载）：** δ = arctan(2·ex/B)（ey 由 Meyerhof B′、L′ 体现）。
程序输入 `soil_category`（argile / sable / mixte）或根据 φ、C 自动判定。

**说明：** `Φ1_sand = 5.3/1.0`（B/(2De) 与 2.5 比较）仅用于**配筋**有效高度 d 折减，**不再**用于 iδβ。
承载力验算中的 iδβ 严格按 Annexe F.1 的 Φ1/Φ2。

qmin/qmax 线性模型仅计 ex（ey 由 Meyerhof B′、L′ 体现）。

### 2.5 容许承载力
```
qu = (q'u - q'0) × iδβ / γq + q'0
```
| 工况 | γq |
|------|-----|
| ELU | 2.0 |
| ELS | 3.0 |

验算：`q'ref < qu`

---

## 3) 倾覆 / 解压（Renversement）

依据 Fascicule 62 **B.3.2**、**B.3.3**，采用 Meyerhof 受压面积：

```
B' = B - 2ex
L' = L - 2ey
As = B' × L'        （受压面积 m²）
A  = B × L
As/A × 100%       （受压面积比 %）
```

| 工况 | 要求 |
|------|------|
| ELU 倾覆 | As/A ≥ **10%** |
| ELS 准永久 | 全截面受压，**σmin = qmin > 0** |
| ELS 稀有 | As/A ≥ **75%** |

---

## 4) 滑动（Glissement）

Fascicule 62 **B.3.4**：

```
Hd ≤ Fp + Vd×tan(φ')/γg1 + C'×A'/γg2
γg1 = 1.2，γg2 = 1.5
A' = As（Meyerhof 受压面积）
C' ≤ 75 kPa

样例计算书未计入 C' 项；默认 `glissement_include_C=0`。若需完整 Fascicule 公式，在 Commun 中设为 `1`。
```

### 被动土压力 Fp（Rankine，Word 8.2.2.4）
```
Fp = ½ × γ × hp² × Kp × L
Kp = tan²(45° + φ/2)
```

---

## 5) 沉降（Tassement — Ménard）

Word 文档 / Fascicule 62 Annexe F.2，均质土：

```
sf = sc + sd

sc = α / (9·Ec) × (q' - σ'v0) × λc × B        (mm，×1000)

sd = 2 / (9·Ed) × (q' - σ'v0) × B0 × (λd·B/B0)^α   (mm，×1000)
```

| 符号 | 含义 |
|------|------|
| q' | ELS 工况 q'ref |
| σ'v0 | q'0 |
| Ec | 球形影响区等效 Ménard 模量（sc） |
| Ed | 偏应力影响区等效 Ménard 模量（sd） |
| B0 | 参考宽度 0.60 m |
| α | 流变系数（与土质有关） |
| λc, λd | 形状系数（与 L/B 有关） |

---

## 6) 配筋（Ferraillage）

Word 8.2.3.1，截面 A-A / B-B：

```
VAA = (qA + qmax)/2 × B1 - γB × h × B1
MAA = (qA + 2qmax)/6 × B1² - γB × h × B1²/2
```

全截面受压时：
```
qA = qmax × (B0 - B1) / B
qB = qmax × B'2 / B
```

部分受压时：
```
qA = qmin + (qmax - qmin) × (B - B1) / B
qB = qmin + (qmax - qmin) × B2 / B
```

抗震最小配筋率：底缘 0.15%，顶缘 0.1%。

**程序配筋 As（与计算书对齐）：**

```
A-A：Mu = |Mmax_ELU|，d = H − e
B-B：Mu = |Mmax_ELU| × 1.2 × L/B；桥墩 d=H−e；仅桥台大基础（B/(2De)>2.5 且 De−H>30cm）时 d=(H−e)×H/De

As (cm²/m) = Mu × 10⁶ / (0.9 × d_mm × fe) / 100
τu (kPa) = Vd / d，d = (H − 上保护层) / 100  (m)
```

算例：墩柱1 B-B，M=334.1，d=min(110,85)=85 cm → As≈12.3 cm²/m。

---

## 运行

```bash
cd expert/fondation_superficielle
python fondation_superficielle.py
```

输入 `input.xlsx` → 输出 `output.xlsx`

---

## 7) 参数取值表

完整 **取值来源 / 规范依据** 见：

| 文件 | Sheet |
|------|--------|
| `input.xlsx` | **参数取值说明** — 通用参考表 + MIDAS 提取说明 |
| `output.xlsx` | **计算说明** — 各墩实际取值 + 派生参数 + 依据 |

### 数据来源概要

| 类别 | 主要来源 |
|------|----------|
| 几何 H/B/L/De | 结构总体设计、基础施工图 |
| 土工 q₀/kp/ple*/φ/C/γ | 地勘报告、Ménard 压密仪试验 |
| 沉降 Ec/Ed/α/λc/λd | 地勘报告、Fascicule 62 Annexe F.2 |
| 荷载 N/ex/ey/Hd | **MIDAS Civil** 各 ELU/ELS 组合 |
| 配筋 M/V | **MIDAS** 或基础局部计算（Word 8.2.3） |
| 系数 γq/γg1/Φ1 等 | 规范固定值，程序内置 |
