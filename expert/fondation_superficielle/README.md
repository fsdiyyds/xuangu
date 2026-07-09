# 浅基础验算 9.1.6 — fondation superficielle

法国公路桥梁浅基础验算（承载力、倾覆、滑动、沉降、配筋）。

## 目录结构

```
fondation_superficielle/
├── fondation_superficielle.py   # 主程序
├── input.xlsx                   # 输入数据（填写后运行）
├── output.xlsx                  # 计算结果（运行后生成）
├── FONDATION_PRINCIPES.md       # 公式原理说明
├── 计算原理.docx                # 规范原文（Fascicule 62）
└── README.md                    # 本文件
```

## 输入表配色

| 颜色 | 含义 |
|------|------|
| **黄色** | 需要填写的数值 |
| 绿/橙/蓝分区标题 | 几何、土工、滑动、沉降、配筋 |
| 灰色行 | 中文说明、单位、列名（勿改） |

## 快速开始

```bash
cd expert/fondation_superficielle
python fondation_superficielle.py --create-template   # 生成分区配色 input 模板
python fondation_superficielle.py                       # 读取 input.xlsx → 输出 output.xlsx
```

`output.xlsx` 版式与计算书 9.1.6 一致，可直接复制到 Word。

可选参数：

```bash
python fondation_superficielle.py --input input.xlsx --output output.xlsx
```

## 输入说明

| Sheet | 内容 |
|-------|------|
| 说明 | 字段说明 |
| Commun | 几何、土工、沉降、配筋公共参数 |
| 墩柱1 … 墩柱5 | 底部【MIDAS 内力粘贴】+ `load_mapping.txt` 自动分类 |

每个墩柱 sheet：**上方**为几何/土工/旁压参数，**最下方**黄色区粘贴 Midas 表（单元、荷载、位置、轴力、剪力、弯矩）。在 `midas_element_ids` 填写基础单元号过滤行。

`load_mapping.txt` 将「荷载」列映射为 ULS/ALS/SLS（及 QP/RARE），程序自动得到 ELU、ELS QP、ELS RARE 各验算工况。

详细公式见 `FONDATION_PRINCIPES.md`。若项目有正式计算书 Word，可将文档放在本目录供对照校准。

## 依赖

- Python 3
- pandas
- openpyxl
