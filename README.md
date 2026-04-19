# ZPY Benchmark Calculation

投入产出表（IOT）与碳排放数据对齐、转换及基准值计算的 MATLAB/Python 项目。

## 项目目的

本项目实现以下目标：
1. 将多部门投入产出表（50部门）映射转换为统一部门结构（22部门）
2. 对齐碳排放数据与投入产出表
3. 计算各行业的碳排放强度基准值

## 目录结构

```
ZPY_Benchmark_Calculation/
├── .gitignore                    # Git忽略配置
├── README.md                     # 本文档
│
├── Benchmark/                    # 基准值计算模块
│   ├── Benchmark_calculation.m   # 主程序：计算直接消耗系数→列昂惕夫逆矩阵→基准值
│   ├── IO_emission_new.xlsx      # 输入数据（投入产出表+碳排放）
│   └── benchmark2022.xlsx        # 输出结果（各行业基准值）
│
├── Consistency_Trans/            # 数据对齐与转换模块
│   ├── WIOT_IEAemission_Align.m  # 主程序：投入产出表映射与碳排放转换
│   ├── ICIO_T_WIOT.m             # 辅助程序：国际投入产出表压缩
│   ├── Align_Table.xlsx          # 映射矩阵（50→22部门）
│   ├── WIOT2022_USDPPP.xlsx      # 2022年世界投入产出表（USD PPP）
│   ├── WIOT2022_USD.xlsx         # 2022年世界投入产出表（USD）
│   ├── Handmade_IEA_EDGAR_2022_world_CO2emission.xlsx  # 2022年世界碳排放数据
│   └── IO_emission_new.xlsx      # 输出：转换后的投入产出表和碳排放向量
│
└── ICIO_PPP_Rebalance/           # ICIO重平衡模块（Python）
    ├── icio_PLI_transform.py     # 用PPP修改ICIO
    ├── gras_icio_rebalance.py    # GRAS重平衡算法
    ├── rebalance_verification.py # 验证重平衡结果
    ├── .venv/                    # Python虚拟环境
    └── README.md                 # 模块详细说明
```

## 使用说明

### 模块一：数据对齐（Consistency_Trans）

将50部门投入产出表和碳排放数据转换为22部门结构。

**前置数据**（放在 `Consistency_Trans/` 目录）：
- `WIOT2022_USDPPP.xlsx` - 2022年世界投入产出表
- `Align_Table.xlsx` - 映射矩阵表
- `Handmade_IEA_EDGAR_2022_world_CO2emission.xlsx` - 碳排放数据

**运行**：
```matlab
% 在 MATLAB 中进入 Consistency_Trans 目录
cd Consistency_Trans
WIOT_IEAemission_Align
```

**输出**：`IO_emission_new.xlsx`

---

### 模块二：基准值计算（Benchmark）

基于转换后的投入产出表计算各行业碳排放强度基准值。

**输入**：`IO_emission_new.xlsx`（来自模块一）

**运行**：
```matlab
% 在 MATLAB 中进入 Benchmark 目录
cd Benchmark
Benchmark_calculation
```

**计算流程**：
1. 读取投入产出表和碳排放数据
2. 计算直接消耗系数矩阵 `A`
3. 计算列昂惕夫逆矩阵 `L = (I - A)^(-1)`
4. 计算生产部门碳排放强度
5. 计算各行业基准值 `baseline = Emission_intensity × L`

**输出**：`benchmark2022.xlsx`

---

### 模块三：ICIO重平衡（ICIO_PPP_Rebalance）

用PPP修正国际投入产出表并应用GRAS算法重平衡。

**运行环境**：Python 3.x（已在 `.venv` 中配置依赖）

**运行**：
```bash
cd ICIO_PPP_Rebalance
.venv\Scripts\activate
python gras_icio_rebalance.py
```

详见 `ICIO_PPP_Rebalance/README.md`

## 数据依赖关系

```
WIOT2022_USDPPP.xlsx ─┬─→ Align_Table.xlsx ─→ WIOT_IEAemission_Align.m
                      │                          ↓
Handmade_IEA_EDGAR...─┴─→ IO_emission_new.xlsx ─→ Benchmark_calculation.m
                                                        ↓
                                              benchmark2022.xlsx (最终结果)
```

## 环境要求

- **MATLAB** R2019b 或更新版本（用于模块一、二）
- **Python** 3.x + pandas, numpy（用于模块三，已配置 `.venv`）

