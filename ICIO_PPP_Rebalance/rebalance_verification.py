import pandas as pd
import numpy as np
import os

def analyze_icio_diff(mod_path, bal_path, output_path):
    print("正在加载文件，请稍候...")
    
    # 1. 读取数据
    # 假设文件为 Excel 格式，第一行为表头，第一列为索引
    # 如果是 CSV 文件，请将 read_excel 改为 read_csv(mod_path, index_col=0)
    df_mod = pd.read_excel(mod_path, index_col=0)
    df_bal = pd.read_excel(bal_path, index_col=0)

    # 2. 提取 Z 矩阵和 FD 矩阵 (Z: 4050x4050, FD: 4050x486)
    # 总列数为 4050 + 486 = 4536
    # 使用 .iloc 提取前 4050 行和前 4536 列的数据部分
    z_fd_mod = df_mod.iloc[0:4050, 0:4536].values
    z_fd_bal = df_bal.iloc[0:4050, 0:4536].values
    
    # 获取对应的行名和列名，用于保存结果
    rows = df_mod.index[0:4050]
    cols = df_mod.columns[0:4536]

    print("正在计算差异...")

    # 3. 计算绝对值差异: |Modified - Balanced|
    diff_abs = np.abs(z_fd_mod - z_fd_bal)

    # 4. 计算相对比例: |差异| / |Modified|
    # 为避免除以 0，使用 np.where 处理
    # 如果 modified 的值为 0 且存在差异，则比例设为 NaN 或 0
    diff_rel = np.divide(diff_abs, np.abs(z_fd_mod), 
                         out=np.zeros_like(diff_abs, dtype=float), 
                         where=z_fd_mod!=0)

    # 5. 打印简单的统计信息辅助验证
    print(f"最大绝对差异: {np.max(diff_abs)}")
    print(f"最大相对比例: {np.max(diff_rel)}")
    print(f"平均绝对差异: {np.mean(diff_abs)}")

    # 6. 筛选相对比例 > 0.5 的条目，提取行列标签 + 差异数据
    mask = diff_rel > 0.5
    row_idx, col_idx = np.where(mask)
    df_large = pd.DataFrame({
        "row_label":  [rows[r] for r in row_idx],
        "col_label":  [cols[c] for c in col_idx],
        "abs_diff":   diff_abs[row_idx, col_idx],
        "rel_ratio":  diff_rel[row_idx, col_idx],
        "mod_value":  z_fd_mod[row_idx, col_idx],
        "bal_value":  z_fd_bal[row_idx, col_idx],
    }).sort_values("rel_ratio", ascending=False).reset_index(drop=True)
    print(f"相对比例 > 0.5 的条目数：{len(df_large)}")

    # 7. 写入结果
    # 筛选汇总表单独写入小文件
    summary_path = output_path.replace('.xlsx', '_large_diff.xlsx')
    df_large.to_excel(summary_path, index=False)
    print(f"筛选汇总已保存至 {summary_path}（共 {len(df_large)} 条）")

    # 完整差异矩阵写入主文件
    print(f"正在将完整矩阵写入 {output_path}...")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_diff_abs = pd.DataFrame(diff_abs, index=rows, columns=cols)
        df_diff_abs.to_excel(writer, sheet_name='Absolute_Difference')

        df_diff_rel = pd.DataFrame(diff_rel, index=rows, columns=cols)
        df_diff_rel.to_excel(writer, sheet_name='Relative_Proportion')

    print("任务完成！")

def verify_r2(mod_path, bal_path,
              n_rows=4050, n_z_cols=4050, n_fd_cols=486):
    """
    计算 balanced vs modified 的决定系数 R²，不写任何文件。

    R² = 1 - SS_res / SS_tot
        SS_res = Σ(balanced - modified)²
        SS_tot = Σ(modified - mean(modified))²

    R² 接近 1 表示 GRAS 平衡结果与修正前结构高度一致；
    R² 显著低于 1 说明平衡过程对矩阵结构改动较大。
    分别报告 Z 块、FD 块、全局三个维度。
    """
    print("正在加载文件（仅计算 R²，不写出任何文件）...")
    df_mod = pd.read_excel(mod_path, index_col=0)
    df_bal = pd.read_excel(bal_path, index_col=0)

    n_total_cols = n_z_cols + n_fd_cols
    A = df_mod.iloc[0:n_rows, 0:n_total_cols].values.astype(float)
    B = df_bal.iloc[0:n_rows, 0:n_total_cols].values.astype(float)

    def _r2(a, b):
        ss_res = np.sum((b - a) ** 2)
        ss_tot = np.sum((a - np.mean(a)) ** 2)
        return 1.0 - ss_res / ss_tot if ss_tot > 0 else float('nan')

    def _stats(a, b, label):
        r2   = _r2(a, b)
        rmse = np.sqrt(np.mean((b - a) ** 2))
        corr = np.corrcoef(a.ravel(), b.ravel())[0, 1]
        print(f"  [{label}]  R²={r2:.6f}  RMSE={rmse:.4e}  Pearson r={corr:.6f}")
        return r2

    print("\n========== GRAS 平衡结果验证 ==========")
    _stats(A[:, :n_z_cols],           B[:, :n_z_cols],           "Z 块 (中间投入)")
    _stats(A[:, n_z_cols:n_total_cols], B[:, n_z_cols:n_total_cols], "FD 块 (最终需求)")
    _stats(A,                           B,                           "全局 Z+FD     ")
    print("========================================\n")


# ==========================================
# 主函数入口
# ==========================================
if __name__ == "__main__":
    FILE_MODIFIED = "modified_ICIO.xlsx"
    FILE_BALANCED = "balanced_ICIO.xlsx"
    FILE_OUTPUT   = "rebalance_diff.xlsx"

    if not (os.path.exists(FILE_MODIFIED) and os.path.exists(FILE_BALANCED)):
        print(f"提示：请确保目录下存在 {FILE_MODIFIED} 和 {FILE_BALANCED}")
    else:
        # 只验证 R²（快，不写文件）
        verify_r2(FILE_MODIFIED, FILE_BALANCED)

        # 若还需要生成差异矩阵 Excel，取消下行注释
        # analyze_icio_diff(FILE_MODIFIED, FILE_BALANCED, FILE_OUTPUT)
        