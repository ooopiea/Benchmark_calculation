# Benchmark Calculation - Input-Output Table & Carbon Emission Alignment

## 项目概述

本项目通过映射矩阵实现投入产出表（Input-Output Table, IOT）的维度转换和碳排放数据的对齐。主要目标是将 50 个部门的投入产出表和碳排放数据转换为 22 个部门的结构，便于进行碳责任核算和国际比较分析。

## 文件结构

### 文件清单

- **WIOT_IEAemission_Align.m** - 主要程序：投入产出表映射与碳排放转换
- **ICIO_T_WIOT.m** - 辅助程序：国际投入产出表（ICIO）数据压缩
- **README.MD** - 本文档

### 数据文件（需要放在同级目录）

| 文件名 | 说明 |
|--------|------|
| `WIOT2022_USDPPP.xlsx` | 2022年世界投入产出表（USD购买力平价） |
| `Align_Table.xlsx` | 映射矩阵表（50部门→22部门，34部门→22部门） |
| `Handmade_IEA_EDGAR_2022_world_CO2emission.xlsx` | 2022年世界碳排放数据 |

### 输出文件

- **IO_emission_new.xlsx** - 转换后的投入产出表和碳排放向量
  - Sheet: `IO_new` - 转换后的投入产出表矩阵
  - Sheet: `Emission_new` - 转换后的部门碳排放向量

## 程序说明

### WIOT_IEAemission_Align.m 主要功能

#### 1. 数据读取
```matlab
path_IOT = 'WIOT2022_USDPPP.xlsx';        % 投入产出表
path_AlignTable = 'Align_Table.xlsx';      % 映射矩阵
path_Emission = 'Handmade_IEA_EDGAR_2022_world_CO2emission.xlsx'; % 碳排放
```

#### 2. 数据维度
- **原始表维度**：50 部门 × 56 列（50 部门 + 6 个最终需求列）
- **转换后维度**：22 部门 × 28 列（22 部门 + 6 个最终需求列）
- **增加值行数**：3 行（工资、营业盈余、税收等）

#### 3. 核心转换公式

投入产出表转换遵循以下矩阵运算：

$$Z_{new} = M \times Z_{old} \times M^T$$

$$F_{new} = M \times F_{old}$$

$$VA_{new} = VA_{old} \times M^T$$

其中 $M$ 是转换矩阵（22×50）

#### 4. 碳排放转换
```
排放向量转换：Emission_new = T_matrix_emission' × Emission_old
```

#### 5. 数据验证
程序包含两项验证机制：
- **维度检查**：输入输出表的行列维度验证
- **总量守恒**：验证转换后的总产出是否与原表映射值一致

## 使用方法

### 环境要求
- MATLAB R2019b 或更新版本
- 数据处理工具箱（用于 `readmatrix`、`writematrix`）

### 运行步骤

1. **准备数据文件**
   - 将所需的三个 Excel 文件放在与 `.m` 文件同一目录下

2. **运行主程序**
   ```matlab
   % 在 MATLAB 命令窗口中运行
   WIOT_IEAemission_Align
   ```

3. **查看结果**
   - 控制台输出维度信息和总量守恒检查结果
   - 生成的 `IO_emission_new.xlsx` 包含转换后的数据

### 输出示例

```
原始表维度: 50 x 56
新表维度: 22 x 28
转换成功：行向总量平衡。
原始 50 部门总排放: 12345.67 Mt
转换后 22 部门总排放: 12345.67 Mt
结果已成功输出到 IO_emission_new.xlsx
```

## 技术细节

### 映射矩阵结构

映射矩阵存储在 `Align_Table.xlsx` 中，需要两个工作表：
- **50TO22_OECD-IEA**：50 部门到 22 部门的映射（50×22 或 22×50）
- **34TO22_IEA-OECD**：34 个国家/地区到 22 个OECD标准部门的映射

### 缺失值处理

程序使用 `fillmissing` 函数自动处理映射矩阵中的缺失值，填充为 0：
```matlab
T_matrix = fillmissing(T_matrix, 'constant', 0);
```

### 数据一致性

投入产出表遵循基本恒等式：
$$X = Z \cdot \mathbf{1} + V^T + TLS$$

其中：
- $X$ = 总产出向量
- $Z$ = 中间投入矩阵
- $\mathbf{1}$ = 单位列向量
- $V$ = 增加值向量
- $TLS$ = 净税收

## 文件 ICIO_T_WIOT.m 说明

该程序用于处理国际投入产出表（ICIO）的压缩：
- 输入：81 国家 × 50 部门的全球投入产出表 (4050×4050)
- 输出：世界投入产出表 WIOT (50×50)
- 方法：使用求和算子 (Summation Operator) 进行国家维度的聚合

## 常见问题

### Q: 如何修改部门分类？
**A**: 修改 `n_old`、`n_new`、`fd_cols` 等参数，并更新相应的映射矩阵。

### Q: 转换过程中如何处理负值？
**A**: 当前程序保留所有值（包括负值）。如需特殊处理，可添加后续步骤。

### Q: 内存不足怎么办？
**A**: 考虑将数据分块处理或使用稀疏矩阵表示。

## 参考文献

- 投入产出分析理论基础：Leontief (1936)
- 国际投入产出表标准：ECCOINOMICS 标准
- 碳排放核算方法：IPCC 指南

## 许可证

本项目代码用于学术研究和教学目的。

## 联系方式

如有问题或改进建议，欢迎反馈。
