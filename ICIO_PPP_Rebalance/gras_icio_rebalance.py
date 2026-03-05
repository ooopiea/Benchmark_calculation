# -*- coding: utf-8 -*-
"""
============================================================================
 GRAS-based ICIO Table Rebalancing  (OECD 2022 × World Bank PPP)
============================================================================
 目标:对经 PLI(Price Level Index)修正过行侧(Z、FD、OUT_col)的 modified_ICIO
      表进行再平衡,得到以 USD PPP 为单位的世界投入产出平衡表。

 核心算法:Generalized RAS (GRAS) + ROW 残差策略
   参考文献:
     Junius, T., & Oosterhaven, J. (2003).
       The solution of updating or regionalizing a matrix with both
       positive and negative entries. Economic Systems Research, 15(1).
     Lenzen, M., Wood, R., & Gallego, B. (2007).
       Some comments on the GRAS method. Economic Systems Research, 19(4).
     Temurshoev, U., Miller, R. E., & Bouwmeester, M. C. (2013).
       A note on the GRAS method. Economic Systems Research, 25(3).
     OECD (2021). Development of the OECD Inter-Country Input-Output Database.
       ROW 作为"最后替代项":ROW 列固定,非 ROW 子矩阵运行 GRAS,
       ROW 行/列通过残差方式确定,防止 GRAS 乘子对小值项产生爆炸式放大。

 防爆炸机制(OECD ROW 残差策略):
   1. ROW 的列(作为消费方,即进口侧)固定在 PLI 调整后的初始值,不参与 GRAS 缩放。
   2. 从各行目标中减去 ROW 固定列贡献,仅对非 ROW 列的子矩阵运行 GRAS。
   3. GRAS 收敛后,ROW 各行的非 ROW 列值 = 各列目标 − 非 ROW 行的已平衡列和(列残差)。
   4. ROW Z 列的 VA 按列平衡条件反推调整,保持会计恒等式。

 作者:ZBY
 Python>=3.9, numpy, pandas, openpyxl
============================================================================
"""

from __future__ import annotations

import warnings
from typing import Dict, Optional, Tuple

import numpy as np
import pandas as pd


# ============================================================================
# 1. GRAS 核心算法
# ============================================================================

def _solve_quadratic_roots(t: np.ndarray, p: np.ndarray, n: np.ndarray,
                           eps: float = 1e-15) -> np.ndarray:
    """
    逐元素求解 `x * p - n / x = t` 的正根,等价于 `p*x^2 - t*x - n = 0`。

    对应 GRAS 每次更新行乘子 r_i(或列乘子 s_j)时需要求解的方程:
        r_i * p_i(s) - n_i(s) / r_i = u_i
    取正根:
        x = ( t + sqrt(t^2 + 4 p n) ) / (2 p)          当 p>0, n>0
        x = t / p                                      当 p>0, n=0 (且 t>0)
        x = -n / t                                     当 p=0, n>0 (且 t<0)
        x = 1                                          其余(行/列全零等)
    `eps` 用于区分"实质为零"与数值噪声。
    """
    t = np.asarray(t, dtype=np.float64)
    p = np.asarray(p, dtype=np.float64)
    n = np.asarray(n, dtype=np.float64)

    x = np.ones_like(t, dtype=np.float64)

    # ---- 主分支:P、N 同时有正质量 ----
    both = (p > eps) & (n > eps)
    disc = t[both] ** 2 + 4.0 * p[both] * n[both]         # 判别式 ≥ 0
    x[both] = (t[both] + np.sqrt(disc)) / (2.0 * p[both])

    # ---- 只有 P(全部正项),形式上对应经典 RAS ----
    p_only = (p > eps) & (n <= eps)
    x[p_only & (t > eps)] = (t / np.maximum(p, eps))[p_only & (t > eps)]

    # ---- 只有 N(全部负项) ----
    n_only = (p <= eps) & (n > eps)
    x[n_only & (t < -eps)] = (-n / np.minimum(t, -eps))[n_only & (t < -eps)]

    # 数值保护:避免出现非正或 Inf/NaN 的 x
    x = np.where((x > 0) & np.isfinite(x), x, 1.0)
    return x


def gras(
    M: np.ndarray,
    u: np.ndarray,
    v: np.ndarray,
    tol: float = 1e-7,
    max_iter: int = 500,
    stall_patience: int = 20,
    verbose: bool = True,
    eps: float = 1e-15,
) -> Tuple[np.ndarray, Dict]:
    """
    GRAS 矩阵平衡(Temurshoev et al. 2013 版本的闭式迭代)。

    Parameters
    ----------
    M : (m, n) ndarray,原矩阵(可含负数)。
    u : (m,)   ndarray,目标行和。
    v : (n,)   ndarray,目标列和。
    tol : 收敛相对误差阈值(相对于 sum 规模)。
    max_iter : 最大迭代次数。
    stall_patience : 连续多少次误差几乎不降则判停滞。
    verbose : 打印过程。
    eps : 数值零阈值。

    Returns
    -------
    M_new : (m, n) ndarray,平衡后矩阵。
    info  : dict,包含 iterations、error、rel_error、r、s、converged。

    Raises
    ------
    ValueError   : 目标不可行(总和不匹配、符号不兼容、零行列目标非零)。
    RuntimeError : 达到 max_iter 或停滞仍未收敛。
    """
    # -------- 0. 输入规整 --------
    M = np.asarray(M, dtype=np.float64)
    u = np.asarray(u, dtype=np.float64).ravel()
    v = np.asarray(v, dtype=np.float64).ravel()
    m, n = M.shape

    if u.shape[0] != m or v.shape[0] != n:
        raise ValueError(f"维度不匹配: M{M.shape}, u{u.shape}, v{v.shape}")
    if not np.all(np.isfinite(M)):
        raise ValueError("M 含非有限值(NaN/Inf),请先清洗。")

    # -------- 1. 总量可行性 --------
    total_u, total_v = u.sum(), v.sum()
    scale = max(abs(total_u), abs(total_v), 1.0)
    if abs(total_u - total_v) > 1e-6 * scale:
        raise ValueError(
            f"[GRAS] 不可行:sum(u)={total_u:.6g} ≠ sum(v)={total_v:.6g},"
            f"相对差 {abs(total_u - total_v)/scale:.2e}。"
            "请先修正宏观约束(例如调整 VA/TLS 使行列总量相等)。"
        )

    # -------- 2. 分解 M = P - N,符号可行性检查 --------
    P = np.where(M > 0, M, 0.0)
    N = np.where(M < 0, -M, 0.0)

    P_row, N_row = P.sum(axis=1), N.sum(axis=1)
    P_col, N_col = P.sum(axis=0), N.sum(axis=0)

    def _check(sums_p, sums_n, tgt, axis_name):
        bad = []
        for k in range(len(tgt)):
            if sums_p[k] < eps and sums_n[k] < eps:
                if abs(tgt[k]) > eps * scale:
                    bad.append((k, "全零", tgt[k]))
            elif sums_p[k] < eps and tgt[k] > eps * scale:
                bad.append((k, "无正项", tgt[k]))
            elif sums_n[k] < eps and tgt[k] < -eps * scale:
                bad.append((k, "无负项", tgt[k]))
        if bad:
            sample = "\n".join(
                f"    {axis_name}={k:>5d}: {reason} 但 target={t:.4g}"
                for k, reason, t in bad[:10]
            )
            raise ValueError(
                f"[GRAS] 符号不可行的 {axis_name} 共 {len(bad)} 条,示例:\n{sample}"
            )

    _check(P_row, N_row, u, "row")
    _check(P_col, N_col, v, "col")

    # -------- 3. 迭代 --------
    r = np.ones(m)
    s = np.ones(n)
    prev_err = np.inf
    stall = 0

    for it in range(1, max_iter + 1):
        # ---- 3.1 更新 s:对每列 j 解二次方程 ----
        r_inv = 1.0 / r
        p_j = P.T @ r        # 形状 (n,) ;  p_j[j] = Σ_i r_i * P[i,j]
        n_j = N.T @ r_inv    # 形状 (n,) ;  n_j[j] = Σ_i N[i,j] / r_i
        s = _solve_quadratic_roots(v, p_j, n_j, eps)

        # ---- 3.2 更新 r:对每行 i 解二次方程 ----
        s_inv = 1.0 / s
        p_i = P @ s          # 形状 (m,)
        n_i = N @ s_inv      # 形状 (m,)
        r = _solve_quadratic_roots(u, p_i, n_i, eps)

        # ---- 3.3 收敛判定 ----
        r_inv = 1.0 / r
        col_sums = s * (P.T @ r) - (N.T @ r_inv) / s
        col_err = np.max(np.abs(col_sums - v))
        row_sums = r * (P @ s) - (N @ (1.0 / s)) / r
        row_err = np.max(np.abs(row_sums - u))
        err = max(row_err, col_err)
        rel_err = err / scale

        if verbose and (it <= 3 or it % 20 == 0):
            print(f"[GRAS] iter={it:4d}  row_err={row_err:.3e}  "
                  f"col_err={col_err:.3e}  rel_err={rel_err:.3e}")

        if rel_err < tol:
            if verbose:
                print(f"[GRAS] ✓ 收敛于第 {it} 步,rel_err={rel_err:.3e}")
            M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
            return M_new, dict(iterations=it, error=err, rel_error=rel_err,
                               r=r, s=s, converged=True)

        if abs(prev_err - err) < tol * scale * 1e-3:
            stall += 1
            if stall >= stall_patience:
                raise RuntimeError(
                    f"[GRAS] 第 {it} 步停滞,误差几乎不再下降 "
                    f"(rel_err={rel_err:.3e} > tol={tol:.1e})。\n"
                    "  建议:\n"
                    "   (1) 放宽 tol 到 1e-5 ~ 1e-6;\n"
                    "   (2) 检查数据是否有极小/极大的数量级混杂;\n"
                    "   (3) 核对行列总量是否已严格对齐。"
                )
        else:
            stall = 0
        prev_err = err

    raise RuntimeError(
        f"[GRAS] {max_iter} 次迭代后未收敛,rel_err={rel_err:.3e}。\n"
        "  建议:(1) 增大 max_iter;(2) 放宽 tol;(3) 检查 sign 可行性。"
    )


# ============================================================================
# 2. 不可行列修正(子函数,供主流程复用)
# ============================================================================

def _fix_infeasible_z_cols(
    M: np.ndarray,
    v_Z: np.ndarray,
    v_FD: np.ndarray,
    col_x_target: np.ndarray,
    u_total: float,
    VA_new: np.ndarray,
    TLS_z_new: np.ndarray,
    n_zcols: int,
    eps: float = 1e-15,
    verbose: bool = True,
    label: str = "",
) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    """
    修正 Z 部分列的符号不可行性(全零/无负项/无正项列目标不匹配)。

    Parameters
    ----------
    M            : [Z_part | FD_part] 子矩阵。
    v_Z          : Z 部分各列目标(长度 n_zcols)。
    v_FD         : FD 部分各列目标。
    col_x_target : 各 Z 列对应产业的总产出(用作 VA+TLS 上限),长度 n_zcols。
    u_total      : sum(u),用于校验全局平衡。
    VA_new       : 各 Z 列的 VA(长度 n_zcols)。
    TLS_z_new    : 各 Z 列的 TLS(长度 n_zcols)。
    n_zcols      : M 中 Z 部分的列数。

    返回修正后的 (M, v_Z, v_FD, VA_new, TLS_z_new)。
    v_FD 按权重吸收残余 gap,确保 sum(u)==sum(v_Z)+sum(v_FD)。
    """
    Z_part = M[:, :n_zcols]
    P_zcol = np.where(Z_part > 0, Z_part, 0.0).sum(axis=0)
    N_zcol = np.where(Z_part < 0, -Z_part, 0.0).sum(axis=0)

    mask_A = (P_zcol < eps) & (N_zcol < eps) & (np.abs(v_Z) > 1e-12)
    mask_B = (N_zcol < eps) & (v_Z < -1e-12) & ~mask_A
    mask_C = (P_zcol < eps) & (v_Z > 1e-12) & ~mask_A
    infeasible = mask_A | mask_B | mask_C

    if not np.any(infeasible):
        return M, v_Z, v_FD, VA_new, TLS_z_new

    if verbose:
        pfx = f"[{label}] " if label else ""
        print(f"       {pfx}不可行 Z 列: {np.sum(infeasible)} 条 "
              f"(全零={np.sum(mask_A)}, 无负项={np.sum(mask_B)}, "
              f"无正项={np.sum(mask_C)}),修正 v_Z → 0 ...")

    v_Z = v_Z.copy()
    v_FD = v_FD.copy()
    VA_new = VA_new.copy()
    TLS_z_new = TLS_z_new.copy()
    M = M.copy()

    for j in np.where(infeasible)[0]:
        if v_Z[j] < 0:
            # 列无法产生负和:压缩 VA+TLS 使其不超过总产出
            total_vt = VA_new[j] + TLS_z_new[j]
            if total_vt > eps:
                cap = max(float(col_x_target[j]), 0.0)
                VA_new[j] *= cap / total_vt
                TLS_z_new[j] *= cap / total_vt
            v_Z[j] = 0.0
        else:
            # 列无法产生正和:超额补入 VA
            VA_new[j] += v_Z[j]
            v_Z[j] = 0.0
        M[:, j] = 0.0

    # 将修正产生的 gap 按权重分配到 v_FD
    gap2 = u_total - (v_Z.sum() + v_FD.sum())
    if abs(gap2) > 1e-9:
        v_FD_sum = v_FD.sum()
        if v_FD_sum > eps:
            v_FD = v_FD + gap2 * (v_FD / v_FD_sum)
        else:
            v_FD = v_FD + gap2 / max(len(v_FD), 1)

    if verbose:
        pfx = f"[{label}] " if label else ""
        print(f"       {pfx}修正后 |sum(u)-sum(v)| = "
              f"{abs(u_total - v_Z.sum() - v_FD.sum()):.3e}")

    return M, v_Z, v_FD, VA_new, TLS_z_new


# ============================================================================
# 3. ICIO 再平衡主流程(OECD ROW 残差策略)
# ============================================================================

def rebalance_icio(
    input_path: str,
    output_path: str,
    n_countries: int = 81,
    n_industries: int = 50,
    n_fd_categories: int = 6,
    row_country_idx: int = 80,
    tol: float = 1e-6,
    max_iter: int = 2000,
    verbose: bool = True,
) -> pd.DataFrame:
    """
    读取 modified_ICIO.xlsx,用 GRAS + OECD ROW 残差策略再平衡,写出平衡表。

    OECD ROW 残差策略
    -----------------
    ROW(rest of world,0-based 国家编号 row_country_idx)的列(进口侧)固定在
    PLI 调整后的初始值,不参与 GRAS 缩放。仅对非 ROW 列的子矩阵运行 GRAS。
    GRAS 收敛后:
      - 非 ROW 列的列和由 GRAS 精确满足;
      - ROW 行的非 ROW 列贡献 = 各列目标 − 非 ROW 行的已平衡列和(列残差);
      - ROW Z 列的 VA 反推调整,保持列平衡会计恒等式。

    防爆炸效果
    ----------
    非 ROW 条目的行目标已扣除 ROW 固定列贡献,GRAS 乘子只需满足
    非 ROW 子系统的约束,消除了因 ROW 隐含的大残差导致的乘子爆炸。

    Parameters
    ----------
    input_path       : modified_ICIO.xlsx 路径。
    output_path      : 输出文件路径。
    n_countries      : 国家数(含 ROW),默认 81。
    n_industries     : 产业数,默认 50。
    n_fd_categories  : 最终需求分类数,默认 6。
    row_country_idx  : ROW 在国家列表中的 0-based 位置,默认 80(最后)。
    tol              : GRAS 收敛阈值。
    max_iter         : GRAS 最大迭代次数。
    verbose          : 打印过程。

    结构假设:
        Rows = [ n_ci 条 country_industry ;  TLS ;  VA ;  OUT ]
        Cols = [ n_ci 条 country_industry ;  n_fd 条 country_finaldemand ;  OUT ]
        n_ci = n_countries * n_industries = 4050
        n_fd = n_countries * n_fd_categories = 486
    """
    n_ci = n_countries * n_industries
    n_fd = n_countries * n_fd_categories

    # -----------------------------------------------------------------------
    # Step 1: 读取 Excel
    # -----------------------------------------------------------------------
    if verbose:
        print(f"[1/6] 读取 {input_path} ...")
    df = pd.read_excel(input_path, header=0, index_col=0)
    if verbose:
        print(f"       原表形状: {df.shape} (期望 ({n_ci+3}, {n_ci+n_fd+1}))")

    if df.shape != (n_ci + 3, n_ci + n_fd + 1):
        warnings.warn(
            f"实际形状 {df.shape} 与期望 ({n_ci+3}, {n_ci+n_fd+1}) 不符,"
            "继续按位置切片,请确认表结构。"
        )

    # -----------------------------------------------------------------------
    # Step 2: 切分各块
    # -----------------------------------------------------------------------
    if verbose:
        print("[2/6] 切分 Z / FD / 行列边缘 ...")

    Z       = df.iloc[:n_ci,          :n_ci          ].to_numpy(dtype=np.float64)
    FD      = df.iloc[:n_ci,          n_ci:n_ci+n_fd ].to_numpy(dtype=np.float64)
    OUT_col = df.iloc[:n_ci,          n_ci+n_fd      ].to_numpy(dtype=np.float64)

    TLS_z   = df.iloc[n_ci,           :n_ci          ].to_numpy(dtype=np.float64)
    TLS_fd  = df.iloc[n_ci,           n_ci:n_ci+n_fd ].to_numpy(dtype=np.float64)
    VA      = df.iloc[n_ci + 1,       :n_ci          ].to_numpy(dtype=np.float64)
    OUT_row = df.iloc[n_ci + 2,       :n_ci          ].to_numpy(dtype=np.float64)

    for name, arr in [("Z", Z), ("FD", FD), ("TLS_z", TLS_z), ("TLS_fd", TLS_fd),
                      ("VA", VA), ("OUT_col", OUT_col), ("OUT_row", OUT_row)]:
        if np.any(~np.isfinite(arr)):
            n_bad = np.sum(~np.isfinite(arr))
            warnings.warn(f"{name} 含 {n_bad} 个 NaN/Inf,已替换为 0")
            arr[~np.isfinite(arr)] = 0.0

    # -----------------------------------------------------------------------
    # Step 3: 预处理 —— 反推 PLI,缩放 VA/TLS,宏观对齐
    # -----------------------------------------------------------------------
    if verbose:
        print("[3/6] 预处理:反推 PLI 并缩放 VA/TLS ...")

    with np.errstate(divide="ignore", invalid="ignore"):
        pli_per_ci = np.where((OUT_row > 0) & np.isfinite(OUT_row),
                              OUT_col / OUT_row, np.nan)

    pli_per_country = np.zeros(n_countries)
    for c in range(n_countries):
        sl = slice(c * n_industries, (c + 1) * n_industries)
        vec = pli_per_ci[sl]
        vec = vec[np.isfinite(vec) & (vec > 0)]
        pli_per_country[c] = vec.mean() if len(vec) > 0 else 1.0

    if verbose:
        print(f"       PLI 范围: [{pli_per_country.min():.4f}, "
              f"{pli_per_country.max():.4f}],均值 {pli_per_country.mean():.4f}")

    pli_ci = np.repeat(pli_per_country, n_industries)
    pli_fd = np.repeat(pli_per_country, n_fd_categories)

    VA_scaled     = VA     * pli_ci
    TLS_z_scaled  = TLS_z  * pli_ci
    TLS_fd_scaled = TLS_fd * pli_fd

    fd_total = float(FD.sum())
    vt_total = float(VA_scaled.sum() + TLS_z_scaled.sum())
    if vt_total <= 0:
        raise ValueError("缩放后 VA+TLS_z 总量 ≤ 0,无法做宏观对齐")

    k = fd_total / vt_total
    if verbose:
        print(f"       宏观对齐系数 k = {k:.6f}  "
              f"(PLI 缩放后净差 {(k-1)*100:+.4f}%)")

    VA_new    = VA_scaled    * k
    TLS_z_new = TLS_z_scaled * k

    # -----------------------------------------------------------------------
    # Step 4: ROW 分区 —— 识别 ROW 的行/列索引
    # -----------------------------------------------------------------------
    if verbose:
        print(f"[4/6] ROW 残差策略:固定国家 {row_country_idx} (ROW) 的列 ...")

    # ROW 在 Z 列中的切片(Z 是方阵,ROW 行列号相同)
    row_ci_s = row_country_idx * n_industries          # e.g. 4000
    row_ci_e = (row_country_idx + 1) * n_industries    # e.g. 4050
    # ROW 在 FD 列中的切片
    row_fd_s = row_country_idx * n_fd_categories       # e.g. 480
    row_fd_e = (row_country_idx + 1) * n_fd_categories # e.g. 486

    row_zcol_idx  = np.arange(row_ci_s, row_ci_e)      # ROW 的 Z 列(n_industries 条)
    row_row_idx   = row_zcol_idx                        # ROW 的行索引(同一国家)
    non_zcol_idx  = np.r_[0:row_ci_s, row_ci_e:n_ci]   # 非 ROW 的 Z 列
    row_fdcol_idx = np.arange(row_fd_s, row_fd_e)       # ROW 的 FD 列
    non_fdcol_idx = np.r_[0:row_fd_s, row_fd_e:n_fd]   # 非 ROW 的 FD 列

    n_non_zcol  = len(non_zcol_idx)   # n_ci - n_industries
    n_non_fdcol = len(non_fdcol_idx)  # n_fd - n_fd_categories

    # -----------------------------------------------------------------------
    # Step 4a: 构造全量 GRAS 输入变量(行/列目标)
    # -----------------------------------------------------------------------
    X_target = OUT_col.copy()
    v_Z      = X_target - VA_new - TLS_z_new   # Z+FD 各列的总目标
    v_FD     = FD.sum(axis=0)                  # FD 列目标

    # -----------------------------------------------------------------------
    # Step 4b: 固定 ROW 列 —— 不参与 GRAS,从行目标中扣除
    # -----------------------------------------------------------------------
    # ROW 固定列值(PLI 调整后,Z 和 FD 各自的 ROW 列)
    Z_row_fix  = Z[:, row_zcol_idx].copy()   # shape (n_ci, n_industries)
    FD_row_fix = FD[:, row_fdcol_idx].copy() # shape (n_ci, n_fd_categories)

    # 每行从 ROW 固定列获得的贡献
    row_fix_contrib = Z_row_fix.sum(axis=1) + FD_row_fix.sum(axis=1)  # (n_ci,)

    # 非 ROW 子矩阵的行目标
    u_sub = X_target - row_fix_contrib  # (n_ci,)

    # 若某行 u_sub < 0(ROW 固定列已超出该行总产出目标):按比例缩减固定值
    neg_mask = u_sub < -1e-9
    if np.any(neg_mask):
        n_neg = int(np.sum(neg_mask))
        if verbose:
            print(f"       警告: {n_neg} 行 ROW 固定列贡献超过行目标,按比例缩减 ...")
        for i in np.where(neg_mask)[0]:
            fix_sum = row_fix_contrib[i]
            if fix_sum > 1e-15:
                scale_fix = max(X_target[i] / fix_sum, 0.0)
                Z_row_fix[i, :]  *= scale_fix
                FD_row_fix[i, :] *= scale_fix
        row_fix_contrib = Z_row_fix.sum(axis=1) + FD_row_fix.sum(axis=1)
        u_sub = X_target - row_fix_contrib

    # -----------------------------------------------------------------------
    # Step 4c: 构造非 ROW 子矩阵及其列目标
    # -----------------------------------------------------------------------
    # 非 ROW 子矩阵:[Z 非ROW列 | FD 非ROW列]
    M_sub = np.hstack([Z[:, non_zcol_idx], FD[:, non_fdcol_idx]])
    # shape: (n_ci, n_non_zcol + n_non_fdcol)

    v_Z_sub  = v_Z[non_zcol_idx]    # 各非 ROW Z 列的目标列和
    v_FD_sub = v_FD[non_fdcol_idx]  # 各非 ROW FD 列的目标列和

    # 微调 v_FD_sub 使 sum(u_sub) == sum(v_Z_sub) + sum(v_FD_sub)
    gap = u_sub.sum() - (v_Z_sub.sum() + v_FD_sub.sum())
    if verbose:
        print(f"       非 ROW 子矩阵:形状 {M_sub.shape}, "
              f"sum(u_sub)={u_sub.sum():.6g}, gap={gap:.3e}")
    if abs(gap) > 1e-9:
        v_FD_sub_sum = v_FD_sub.sum()
        if v_FD_sub_sum > 1e-15:
            v_FD_sub = v_FD_sub + gap * (v_FD_sub / v_FD_sub_sum)
        else:
            v_FD_sub = v_FD_sub + gap / max(len(v_FD_sub), 1)

    # -----------------------------------------------------------------------
    # Step 4d: 修正 Z 部分的符号不可行列(只在 M_sub 的 Z 部分)
    # -----------------------------------------------------------------------
    VA_sub  = VA_new[non_zcol_idx]
    TLS_sub = TLS_z_new[non_zcol_idx]
    # col_x_target: 各非 ROW Z 列对应产业的总产出(正确的 cap 基准)
    col_x_target_sub = X_target[non_zcol_idx]

    M_sub, v_Z_sub, v_FD_sub, VA_sub, TLS_sub = _fix_infeasible_z_cols(
        M_sub, v_Z_sub, v_FD_sub,
        col_x_target=col_x_target_sub,
        u_total=float(u_sub.sum()),
        VA_new=VA_sub,
        TLS_z_new=TLS_sub,
        n_zcols=n_non_zcol,
        verbose=verbose,
        label="非ROW子矩阵",
    )
    VA_new[non_zcol_idx]    = VA_sub
    TLS_z_new[non_zcol_idx] = TLS_sub
    v_sub = np.concatenate([v_Z_sub, v_FD_sub])

    if verbose:
        print(f"       GRAS 前 |sum(u_sub)-sum(v_sub)| = {abs(u_sub.sum()-v_sub.sum()):.3e}")

    # -----------------------------------------------------------------------
    # Step 5: 运行 GRAS(仅非 ROW 子矩阵)
    # -----------------------------------------------------------------------
    if verbose:
        print("[5/6] 运行 GRAS(非 ROW 子矩阵) ...")
    M_sub_bal, info = gras(M_sub, u_sub, v_sub, tol=tol, max_iter=max_iter, verbose=verbose)
    if verbose:
        print(f"       收敛: iter={info['iterations']}  "
              f"rel_err={info['rel_error']:.3e}")

    # -----------------------------------------------------------------------
    # Step 5b: 从平衡后子矩阵提取 Z_bal 和 FD_bal
    # -----------------------------------------------------------------------
    Z_sub_bal  = M_sub_bal[:, :n_non_zcol]    # (n_ci, n_non_zcol)
    FD_sub_bal = M_sub_bal[:, n_non_zcol:]    # (n_ci, n_non_fdcol)

    # 完整 Z_bal 和 FD_bal
    Z_bal  = np.zeros((n_ci, n_ci))
    FD_bal = np.zeros((n_ci, n_fd))

    Z_bal[:, non_zcol_idx]  = Z_sub_bal
    Z_bal[:, row_zcol_idx]  = Z_row_fix          # ROW Z 列:保持初始值

    FD_bal[:, non_fdcol_idx] = FD_sub_bal
    FD_bal[:, row_fdcol_idx] = FD_row_fix         # ROW FD 列:保持初始值

    # 说明:GRAS 在 M_sub 上对全部 n_ci 行(含 ROW 行)同时求解,收敛后
    # Z_sub_bal 和 FD_sub_bal 中的 ROW 行值已经是平衡解的一部分,无需再做
    # "列残差"覆写。ROW 行通过行目标 u_sub[row_rows] 在 GRAS 内自然平衡。

    # -----------------------------------------------------------------------
    # Step 5c: 调整 ROW Z 列的 VA,恢复列平衡会计恒等式
    #   balance: Z_bal[:,j].sum() + VA_new[j] + TLS_z_new[j] = X_target[j]
    # -----------------------------------------------------------------------
    for j in row_zcol_idx:
        needed_va = X_target[j] - Z_bal[:, j].sum() - TLS_z_new[j]
        if needed_va < 0.0:
            # VA 不能为负:优先吸收到 TLS,再吸收到 VA=0
            TLS_z_new[j] = max(X_target[j] - Z_bal[:, j].sum(), 0.0)
            VA_new[j] = 0.0
            if verbose and abs(needed_va) > 1e6:
                print(f"       ROW Z col {j}: VA 调整后为 0 "
                      f"(缺口 {needed_va/1e6:.2f}M, TLS 已调整)")
        else:
            VA_new[j] = needed_va

    # -----------------------------------------------------------------------
    # Step 6: 重组完整表并写出
    # -----------------------------------------------------------------------
    if verbose:
        print(f"[6/6] 重组并保存到 {output_path} ...")

    out = pd.DataFrame(
        index=list(df.index[:n_ci]) + ["TLS", "VA", "OUT"],
        columns=list(df.columns),
        dtype=np.float64,
    )
    out.iloc[:n_ci, :n_ci]                = Z_bal
    out.iloc[:n_ci, n_ci:n_ci + n_fd]     = FD_bal
    out.iloc[:n_ci, n_ci + n_fd]          = X_target

    out.iloc[n_ci,     :n_ci]             = TLS_z_new
    out.iloc[n_ci,     n_ci:n_ci + n_fd]  = TLS_fd_scaled
    out.iloc[n_ci,     n_ci + n_fd]       = 0.0

    out.iloc[n_ci + 1, :n_ci]             = VA_new
    out.iloc[n_ci + 1, n_ci:n_ci + n_fd]  = 0.0
    out.iloc[n_ci + 1, n_ci + n_fd]       = 0.0

    out.iloc[n_ci + 2, :n_ci]             = X_target
    out.iloc[n_ci + 2, n_ci:n_ci + n_fd]  = 0.0
    out.iloc[n_ci + 2, n_ci + n_fd]       = 0.0

    # ---- 事后诊断 ----
    row_sum_check = Z_bal.sum(axis=1) + FD_bal.sum(axis=1)
    row_sum_err   = np.max(np.abs(row_sum_check - X_target))
    col_sum_check = Z_bal.sum(axis=0) + VA_new + TLS_z_new
    col_sum_err   = np.max(np.abs(col_sum_check - X_target))
    scale_diag    = float(X_target.sum())

    # 变化量诊断(识别异常放大)
    Z_orig  = df.iloc[:n_ci, :n_ci].to_numpy(dtype=np.float64)
    FD_orig = df.iloc[:n_ci, n_ci:n_ci+n_fd].to_numpy(dtype=np.float64)
    Z_orig[~np.isfinite(Z_orig)]   = 0.0
    FD_orig[~np.isfinite(FD_orig)] = 0.0

    Z_abs_change  = np.abs(Z_bal - Z_orig)
    FD_abs_change = np.abs(FD_bal - FD_orig)
    max_z_change  = Z_abs_change.max()
    max_fd_change = FD_abs_change.max()
    n_large_z     = int(np.sum(Z_abs_change > 100e6))
    n_large_fd    = int(np.sum(FD_abs_change > 100e6))

    if verbose:
        print(f"       平衡后最大行和误差 = {row_sum_err:.3e}  "
              f"(相对 {row_sum_err/scale_diag:.2e})")
        print(f"       平衡后最大列和误差 = {col_sum_err:.3e}  "
              f"(相对 {col_sum_err/scale_diag:.2e})")
        print(f"       Z 最大绝对变化   = {max_z_change/1e6:.1f}M USD  "
              f"(>100M 的格子数: {n_large_z})")
        print(f"       FD 最大绝对变化  = {max_fd_change/1e6:.1f}M USD  "
              f"(>100M 的格子数: {n_large_fd})")

        # ROW 列误差(不强制满足,记录供参考)
        row_zcol_err = np.max(np.abs(
            Z_bal[:, row_zcol_idx].sum(axis=0) + VA_new[row_zcol_idx]
            + TLS_z_new[row_zcol_idx] - X_target[row_zcol_idx]))
        print(f"       ROW Z 列最大会计误差 = {row_zcol_err:.3e}")

    out.to_excel(output_path)
    if verbose:
        print("[✓] 完成")
    return out


# ============================================================================
# 4. 命令行入口
# ============================================================================

if __name__ == "__main__":
    rebalance_icio(
        input_path="modified_ICIO.xlsx",
        output_path="balanced_ICIO.xlsx",
        n_countries=81,
        n_industries=50,
        n_fd_categories=6,
        row_country_idx=80,   # ROW = 最后一个国家(0-based)
        tol=1e-6,
        max_iter=2000,
        verbose=True,
    )
