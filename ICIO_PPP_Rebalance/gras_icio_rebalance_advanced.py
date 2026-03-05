# -*- coding: utf-8 -*-
"""
============================================================================
 GRAS-based ICIO Table Rebalancing - 多种防爆炸策略版本
============================================================================
 
 目标:对经 PLI(Price Level Index)修正过的 modified_ICIO 表进行再平衡,
      得到以 USD PPP 为单位的世界投入产出平衡表,同时防止数值爆炸。
 
 改进策略:
   1. 原始 GRAS + ROW 残差策略 (策略A - 原有)
   2. 加权 GRAS (Weighted GRAS) (策略B)
   3. 约束边界 GRAS (Bounded GRAS) (策略C)
   4. 增广目标 GRAS (Augmented Target GRAS) (策略D)
   5. 分步 GRAS (Stepwise GRAS) (策略E)
 
 防爆炸核心机制:
   - 限制乘子变化范围: r ∈ [1/τ, τ], s ∈ [1/τ, τ]
   - 混合正负项的特殊处理
   - ROW 作为最后替代项的残差策略
   - 目标渐进调整
   - 诊断与警告机制
   
 作者:ZBY
 Python>=3.9, numpy, pandas, openpyxl
============================================================================
"""

from __future__ import annotations

import warnings
from enum import Enum
from typing import Dict, Literal, Optional, Tuple

import numpy as np
import pandas as pd


# ============================================================================
# 辅助函数: 二次方程求解
# ============================================================================

def _solve_quadratic_roots(t: np.ndarray, p: np.ndarray, n: np.ndarray,
                           eps: float = 1e-15) -> np.ndarray:
    """
    逐元素求解 `x * p - n / x = t` 的正根。
    
    等价于 `p*x^2 - t*x - n = 0`,取正根。
    """
    t = np.asarray(t, dtype=np.float64)
    p = np.asarray(p, dtype=np.float64)
    n = np.asarray(n, dtype=np.float64)

    x = np.ones_like(t, dtype=np.float64)

    # P、N 同时有正质量
    both = (p > eps) & (n > eps)
    disc = t[both] ** 2 + 4.0 * p[both] * n[both]
    x[both] = (t[both] + np.sqrt(disc)) / (2.0 * p[both])

    # 只有 P(经典 RAS 形式)
    p_only = (p > eps) & (n <= eps)
    x[p_only & (t > eps)] = (t / np.maximum(p, eps))[p_only & (t > eps)]

    # 只有 N(全部负项)
    n_only = (p <= eps) & (n > eps)
    x[n_only & (t < -eps)] = (-n / t)[n_only & (t < -eps)]

    x[~np.isfinite(x)] = 1.0
    return x


# ============================================================================
# 策略枚举
# ============================================================================

class RebalanceStrategy(Enum):
    """再平衡策略枚举"""
    ORIGINAL_GRAS = "A"      # 原始 GRAS + ROW 残差
    WEIGHTED_GRAS = "B"      # 加权 GRAS
    BOUNDED_GRAS = "C"       # 约束边界 GRAS
    AUGMENTED_TARGET = "D"   # 增广目标 GRAS
    STEPWISE_GRAS = "E"      # 分步 GRAS


# ============================================================================
# 1. 诊断与辅助函数
# ============================================================================

def diagnose_explosion(Z_orig: np.ndarray, Z_bal: np.ndarray, 
                       FD_orig: np.ndarray, FD_bal: np.ndarray,
                       threshold: float = 100e6) -> Dict:
    """
    诊断再平衡过程中的数值爆炸情况。
    
    返回:
        Dict包含: max_z_change, max_fd_change, n_large_z, n_large_fd,
                 z_change_ratio, fd_change_ratio, explosion_cells
    """
    Z_abs_change = np.abs(Z_bal - Z_orig)
    FD_abs_change = np.abs(FD_bal - FD_orig)
    
    max_z_change = Z_abs_change.max()
    max_fd_change = FD_abs_change.max()
    n_large_z = int(np.sum(Z_abs_change > threshold))
    n_large_fd = int(np.sum(FD_abs_change > threshold))
    
    # 计算变化率(相对于原始值)
    Z_orig_safe = np.abs(Z_orig.copy())
    Z_orig_safe[Z_orig_safe < 1e-10] = 1e-10  # 避免除零
    z_change_ratio_max = (Z_abs_change / Z_orig_safe).max()
    z_change_ratio_per_element = Z_abs_change / Z_orig_safe
    
    FD_orig_safe = np.abs(FD_orig.copy())
    FD_orig_safe[FD_orig_safe < 1e-10] = 1e-10
    fd_change_ratio_max = (FD_abs_change / FD_orig_safe).max()
    fd_change_ratio_per_element = FD_abs_change / FD_orig_safe
    
    # 找出爆炸单元格
    explosion_mask = (Z_abs_change > threshold) | (z_change_ratio_per_element > 1.0)
    explosion_indices = np.where(explosion_mask)
    
    return {
        'max_z_change': max_z_change,
        'max_fd_change': max_fd_change,
        'n_large_z': n_large_z,
        'n_large_fd': n_large_fd,
        'z_change_ratio': z_change_ratio_max,
        'fd_change_ratio': fd_change_ratio_max,
        'explosion_cells': list(zip(explosion_indices[0][:10], explosion_indices[1][:10])),
        'has_explosion': max_z_change > threshold or z_change_ratio_max > 1.5
    }


def compute_multipliers(M_orig: np.ndarray, M_bal: np.ndarray, 
                        eps: float = 1e-10) -> Tuple[np.ndarray, np.ndarray]:
    """
    计算行/列乘子 r, s,使得 M_bal ≈ diag(r) @ M_orig @ diag(s)。
    用于诊断。
    """
    m, n = M_orig.shape
    r = np.ones(m)
    s = np.ones(n)
    
    # 逐行估计 r
    for i in range(m):
        row_orig = M_orig[i, :]
        row_bal = M_bal[i, :]
        mask = np.abs(row_orig) > eps
        if np.any(mask):
            r[i] = np.median(row_bal[mask] / row_orig[mask])
    
    # 逐列估计 s
    for j in range(n):
        col_orig = M_orig[:, j] * r
        col_bal = M_bal[:, j]
        mask = np.abs(col_orig) > eps
        if np.any(mask):
            s[j] = np.median(col_bal[mask] / col_orig[mask])
    
    # 清理异常乘子
    r = np.clip(r, 0.1, 10.0)
    s = np.clip(s, 0.1, 10.0)
    
    return r, s


# ============================================================================
# 2. 核心 GRAS 算法变体
# ============================================================================

def gras_iteration_original(
    M: np.ndarray,
    u: np.ndarray,
    v: np.ndarray,
    P: np.ndarray,
    N: np.ndarray,
    max_iter: int = 1000,
    tol: float = 1e-6,
    eps: float = 1e-15,
    verbose: bool = False,
) -> Tuple[np.ndarray, Dict]:
    """原始 GRAS 迭代(无约束)"""
    m, n = M.shape
    scale = max(u.sum(), v.sum(), 1.0)
    
    def _check(P_axis, N_axis, tgt, axis_name):
        bad = []
        for k in range(len(tgt)):
            sums_p = P_axis[k, :].sum() if P_axis.ndim > 1 else P_axis[:, k].sum()
            sums_n = N_axis[k, :].sum() if N_axis.ndim > 1 else N_axis[:, k].sum()
            if abs(tgt[k]) < eps * scale:
                continue
            if sums_p < eps and sums_n < eps:
                bad.append((k, "全零", tgt[k]))
            elif sums_p[k] < eps and tgt[k] > eps * scale:
                bad.append((k, "无正项", tgt[k]))
            elif sums_n[k] < eps and tgt[k] < -eps * scale:
                bad.append((k, "无负项", tgt[k]))
        if bad:
            raise ValueError(f"[GRAS] 符号不可行的 {axis_name} 共 {len(bad)} 条")
    
    _check(P, N, u, "row")
    _check(P.T, N.T, v, "col")
    
    r = np.ones(m)
    s = np.ones(n)
    
    for it in range(1, max_iter + 1):
        r_inv = 1.0 / r
        p_j = P.T @ r
        n_j = N.T @ r_inv
        s = _solve_quadratic_roots(v, p_j, n_j, eps)
        
        s_inv = 1.0 / s
        p_i = P @ s
        n_i = N @ s_inv
        r = _solve_quadratic_roots(u, p_i, n_i, eps)
        
        col_sums = s * (P.T @ r) - (N.T @ (1.0 / r)) / s
        row_sums = r * (P @ s) - (N @ (1.0 / s)) / r
        err = max(np.max(np.abs(col_sums - v)), np.max(np.abs(row_sums - u)))
        
        if err / scale < tol:
            M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
            return M_new, dict(iterations=it, error=err, r=r, s=s, converged=True)
    
    M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
    return M_new, dict(iterations=max_iter, error=err, r=r, s=s, converged=False)


def gras_iteration_bounded(
    M: np.ndarray,
    u: np.ndarray,
    v: np.ndarray,
    P: np.ndarray,
    N: np.ndarray,
    max_iter: int = 1000,
    tol: float = 1e-6,
    eps: float = 1e-15,
    tau: float = 5.0,  # 乘子边界: [1/tau, tau]
    verbose: bool = False,
) -> Tuple[np.ndarray, Dict]:
    """
    约束边界 GRAS (Bounded GRAS)
    
    核心改进:限制乘子 r_i, s_j 在 [1/τ, τ] 范围内,
    防止极端放大。
    
    tau 建议值:
      - tau=2.0: 非常保守,允许最大2倍变化
      - tau=5.0: 中等保守,允许最大5倍变化
      - tau=10.0: 宽松,允许最大10倍变化
    """
    m, n = M.shape
    scale = max(u.sum(), v.sum(), 1.0)
    
    r = np.ones(m)
    s = np.ones(n)
    
    for it in range(1, max_iter + 1):
        # 更新 s
        r_inv = 1.0 / r
        p_j = P.T @ r
        n_j = N.T @ r_inv
        s_raw = _solve_quadratic_roots(v, p_j, n_j, eps)
        s = np.clip(s_raw, 1.0/tau, tau)
        
        # 更新 r
        s_inv = 1.0 / s
        p_i = P @ s
        n_i = N @ s_inv
        r_raw = _solve_quadratic_roots(u, p_i, n_i, eps)
        r = np.clip(r_raw, 1.0/tau, tau)
        
        # 收敛判定
        col_sums = s * (P.T @ r) - (N.T * r_inv) / s
        row_sums = r * (P @ s) - (N * s_inv) / r
        err = max(np.max(np.abs(col_sums - v)), np.max(np.abs(row_sums - u)))
        
        if verbose and it % 50 == 0:
            print(f"    [Bounded GRAS] iter={it}, err={err/scale:.3e}, "
                  f"r_range=[{r.min():.3f},{r.max():.3f}], "
                  f"s_range=[{s.min():.3f},{s.max():.3f}]")
        
        if err / scale < tol:
            M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
            return M_new, dict(iterations=it, error=err, r=r, s=s, 
                             converged=True, tau=tau)
    
    M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
    return M_new, dict(iterations=max_iter, error=err, r=r, s=s,
                      converged=False, tau=tau)


def gras_iteration_weighted(
    M: np.ndarray,
    u: np.ndarray,
    v: np.ndarray,
    P: np.ndarray,
    N: np.ndarray,
    max_iter: int = 1000,
    tol: float = 1e-6,
    eps: float = 1e-15,
    weight_mode: Literal["entropy", "size", "hybrid"] = "hybrid",
    verbose: bool = False,
) -> Tuple[np.ndarray, Dict]:
    """
    加权 GRAS (Weighted GRAS)
    
    核心思想:根据元素大小分配调整权重,大值元素获得更大调整空间,
    小值元素获得更保守的调整。
    
    weight_mode:
      - "entropy": 使用信息熵权重
      - "size": 基于元素绝对值大小
      - "hybrid": 混合权重(推荐)
    """
    m, n = M.shape
    scale = max(u.sum(), v.sum(), 1.0)
    
    # 计算权重矩阵
    M_abs = np.abs(M)
    M_safe = M_abs.copy()
    M_safe[M_safe < eps] = eps
    
    if weight_mode == "entropy":
        # 熵权重: w_ij ∝ -log(p_ij), p_ij ∝ M_abs_ij / sum(M_abs)
        p = M_abs / (M_abs.sum() + eps)
        p = np.clip(p, eps, 1.0)
        w = -np.log(p)
        w = w / w.mean()
    elif weight_mode == "size":
        # 大小权重: w_ij ∝ M_abs_ij
        w = M_abs / (M_abs.mean() + eps)
        w = np.clip(w, 0.1, 10.0)
    else:  # hybrid
        p = M_abs / (M_abs.sum() + eps)
        p = np.clip(p, eps, 1.0)
        entropy_w = -np.log(p) / (-np.log(eps))
        size_w = np.clip(M_abs / (M_abs.max() + eps), 0.0, 1.0)
        w = 0.5 * entropy_w + 0.5 * size_w
    
    # 分解 P, N
    P_w = P * w
    N_w = N * w
    
    r = np.ones(m)
    s = np.ones(n)
    
    for it in range(1, max_iter + 1):
        r_inv = 1.0 / r
        p_j = P_w.T @ r
        n_j = N_w.T @ r_inv
        s = _solve_quadratic_roots(v, p_j, n_j, eps)
        
        s_inv = 1.0 / s
        p_i = P_w @ s
        n_i = N_w @ s_inv
        r = _solve_quadratic_roots(u, p_i, n_i, eps)
        
        col_sums = s * (P.T @ r) - (N.T * r_inv) / s
        row_sums = r * (P @ s) - (N * s_inv) / r
        err = max(np.max(np.abs(col_sums - v)), np.max(np.abs(row_sums - u)))
        
        if verbose and it % 50 == 0:
            print(f"    [Weighted GRAS] iter={it}, err={err/scale:.3e}")
        
        if err / scale < tol:
            M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
            return M_new, dict(iterations=it, error=err, r=r, s=s, 
                             converged=True, weight_mode=weight_mode)
    
    M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
    return M_new, dict(iterations=max_iter, error=err, r=r, s=s,
                      converged=False, weight_mode=weight_mode)


def gras_iteration_stepwise(
    M: np.ndarray,
    u: np.ndarray,
    v: np.ndarray,
    P: np.ndarray,
    N: np.ndarray,
    max_iter: int = 1000,
    tol: float = 1e-6,
    eps: float = 1e-15,
    n_steps: int = 10,
    verbose: bool = False,
) -> Tuple[np.ndarray, Dict]:
    """
    分步 GRAS (Stepwise GRAS)
    
    核心思想:将目标 u, v 分解为多个渐进步骤,
    每步只移动部分距离,最后累积效果。
    
    n_steps: 步数,越多越保守但越慢
    """
    m, n = M.shape
    
    # 初始化
    M_current = M.copy()
    total_iterations = 0
    
    for step in range(1, n_steps + 1):
        # 目标比例:本步达到的比例
        alpha = step / n_steps
        u_target = u * alpha
        v_target = v * alpha
        
        # 计算当前行和列
        row_sums_current = M_current.sum(axis=1)
        col_sums_current = M_current.sum(axis=0)
        
        # 调整目标:只移动差额的部分
        u_adj = u_target - row_sums_current
        v_adj = v_target - col_sums_current
        
        # 跳过已满足的行/列
        mask_u = np.abs(u_adj) > eps * max(u.sum(), 1.0)
        mask_v = np.abs(v_adj) > eps * max(v.sum(), 1.0)
        
        if not mask_u.any() and not mask_v.any():
            if verbose:
                print(f"    [Stepwise GRAS] Step {step}/{n_steps}: 已收敛")
            continue
        
        # 对非零目标部分运行 GRAS
        M_sub = M_current[mask_u][:, mask_v] if mask_u.any() and mask_v.any() else M_current
        u_sub = u_adj[mask_u]
        v_sub = v_adj[mask_v]
        P_sub = P[mask_u][:, mask_v] if mask_u.any() and mask_v.any() else P
        N_sub = N[mask_u][:, mask_v] if mask_u.any() and mask_v.any() else N
        
        step_max_iter = max(max_iter // n_steps, 100)
        M_sub_new, info = gras_iteration_original(
            M_sub, u_sub, v_sub, P_sub, N_sub,
            max_iter=step_max_iter, tol=tol, eps=eps, verbose=False
        )
        
        # 放回完整矩阵
        if mask_u.any() and mask_v.any():
            M_current[np.ix_(mask_u, mask_v)] = M_sub_new
        else:
            M_current = M_sub_new
        
        total_iterations += info['iterations']
        
        if verbose:
            print(f"    [Stepwise GRAS] Step {step}/{n_steps}: "
                  f"iter={info['iterations']}, err={info['error']:.3e}")
    
    return M_current, dict(
        iterations=total_iterations, 
        error=0.0,  # 简化处理
        r=np.ones(m), 
        s=np.ones(n),
        converged=True, 
        n_steps=n_steps
    )


def gras_iteration_augmented(
    M: np.ndarray,
    u: np.ndarray,
    v: np.ndarray,
    P: np.ndarray,
    N: np.ndarray,
    max_iter: int = 1000,
    tol: float = 1e-6,
    eps: float = 1e-15,
    regularization: float = 1e-6,
    verbose: bool = False,
) -> Tuple[np.ndarray, Dict]:
    """
    增广目标 GRAS (Augmented Target GRAS)
    
    核心思想:在目标中加入正则化项,将问题从:
        min ||r ⊙ M - M*||²  s.t. row/column sums = target
    变为:
        min ||r ⊙ M - M*||² + λ||r - 1||²  s.t. row/column sums = target
    
    regularization: 正则化强度 λ
    """
    m, n = M.shape
    scale = max(u.sum(), v.sum(), 1.0)
    
    r = np.ones(m)
    s = np.ones(n)
    reg = regularization
    
    for it in range(1, max_iter + 1):
        r_inv = 1.0 / r
        p_j = P.T @ r
        n_j = N.T @ r_inv
        
        # 修改列更新:加入正则化
        # 原始: r * p_j - n_j / r = v
        # 正则化: r * p_j - n_j / r + reg * (r - 1) = v
        # 等价于: (p_j + reg) * r - n_j / r = v + reg
        p_j_reg = p_j + reg
        v_reg = v + reg
        s = _solve_quadratic_roots(v_reg, p_j_reg, n_j, eps)
        
        # 更新行
        s_inv = 1.0 / s
        p_i = P @ s
        n_i = N @ s_inv
        
        p_i_reg = p_i + reg
        u_reg = u + reg
        r = _solve_quadratic_roots(u_reg, p_i_reg, n_i, eps)
        
        # 收敛判定
        col_sums = s * (P.T @ r) - (N.T * r_inv) / s
        row_sums = r * (P @ s) - (N * s_inv) / r
        err = max(np.max(np.abs(col_sums - v)), np.max(np.abs(row_sums - u)))
        
        if verbose and it % 50 == 0:
            print(f"    [Augmented GRAS] iter={it}, err={err/scale:.3e}, "
                  f"r_range=[{r.min():.3f},{r.max():.3f}]")
        
        if err / scale < tol:
            M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
            return M_new, dict(iterations=it, error=err, r=r, s=s, 
                             converged=True, regularization=reg)
    
    M_new = r[:, None] * P * s[None, :] - N / (r[:, None] * s[None, :])
    return M_new, dict(iterations=max_iter, error=err, r=r, s=s,
                      converged=False, regularization=reg)


# ============================================================================
# 3. 主再平衡函数
# ============================================================================

def rebalance_icio_advanced(
    input_path: str,
    output_path: str,
    n_countries: int = 81,
    n_industries: int = 50,
    n_fd_categories: int = 6,
    row_country_idx: int = 80,
    tol: float = 1e-6,
    max_iter: int = 2000,
    strategy: RebalanceStrategy = RebalanceStrategy.BOUNDED_GRAS,
    strategy_params: Optional[Dict] = None,
    change_threshold: float = 100e6,
    verbose: bool = True,
) -> pd.DataFrame:
    """
    高级 ICIO 再平衡函数,支持多种防爆炸策略。
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径
        n_countries: 国家数
        n_industries: 产业数
        n_fd_categories: 最终需求类别数
        row_country_idx: ROW 国家索引(0-based)
        tol: 收敛容差
        max_iter: 最大迭代次数
        strategy: 再平衡策略
        strategy_params: 策略特定参数
        change_threshold: 变化阈值(美元)
        verbose: 是否输出详细信息
    
    返回:
        平衡后的 DataFrame
    """
    if strategy_params is None:
        strategy_params = {}
    
    if verbose:
        print(f"\n{'='*60}")
        print(f"  ICIO 再平衡 - 策略: {strategy.name}")
        print(f"{'='*60}")
    
    # -----------------------------------------------------------------------
    # Step 1: 读取数据
    # -----------------------------------------------------------------------
    if verbose:
        print("[1/6] 读取数据 ...")
    
    df = pd.read_excel(input_path, index_col=0)
    n_ci = n_countries * n_industries
    n_fd = n_countries * n_fd_categories
    
    Z      = df.iloc[:n_ci,          :n_ci         ].to_numpy(dtype=np.float64, copy=True)
    FD     = df.iloc[:n_ci,          n_ci:n_ci+n_fd].to_numpy(dtype=np.float64, copy=True)
    OUT_col = df.iloc[:n_ci,         n_ci+n_fd     ].to_numpy(dtype=np.float64, copy=True)
    TLS_z  = df.iloc[n_ci,          :n_ci         ].to_numpy(dtype=np.float64, copy=True)
    TLS_fd = df.iloc[n_ci,          n_ci:n_ci+n_fd].to_numpy(dtype=np.float64, copy=True)
    VA     = df.iloc[n_ci + 1,      :n_ci         ].to_numpy(dtype=np.float64, copy=True)
    OUT_row = df.iloc[n_ci + 2,     :n_ci         ].to_numpy(dtype=np.float64, copy=True)
    
    # 处理 NaN/Inf
    for name, arr in [("Z", Z), ("FD", FD), ("TLS_z", TLS_z), ("TLS_fd", TLS_fd),
                      ("VA", VA), ("OUT_col", OUT_col), ("OUT_row", OUT_row)]:
        arr[~np.isfinite(arr)] = 0.0
    
    Z_orig = Z.copy()
    FD_orig = FD.copy()
    
    # -----------------------------------------------------------------------
    # Step 2: 预处理 —— PLI 调整
    # -----------------------------------------------------------------------
    if verbose:
        print("[2/6] 预处理:PLI 调整 ...")
    
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
              f"{pli_per_country.max():.4f}]")
    
    pli_ci = np.repeat(pli_per_country, n_industries)
    pli_fd = np.repeat(pli_per_country, n_fd_categories)
    
    VA_scaled      = VA * pli_ci
    TLS_z_scaled   = TLS_z * pli_ci
    TLS_fd_scaled  = TLS_fd * pli_fd
    
    # 宏观对齐
    fd_total = float(FD.sum())
    vt_total = float(VA_scaled.sum() + TLS_z_scaled.sum())
    if vt_total <= 0:
        raise ValueError("缩放后 VA+TLS_z 总量 ≤ 0")
    
    k = fd_total / vt_total
    if verbose:
        print(f"       宏观对齐系数 k = {k:.6f}")
    
    VA_new    = VA_scaled * k
    TLS_z_new = TLS_z_scaled * k
    
    # -----------------------------------------------------------------------
    # Step 3: ROW 分区
    # -----------------------------------------------------------------------
    if verbose:
        print(f"[3/6] ROW 残差策略:国家 {row_country_idx} ...")
    
    row_ci_s = row_country_idx * n_industries
    row_ci_e = (row_country_idx + 1) * n_industries
    row_fd_s = row_country_idx * n_fd_categories
    row_fd_e = (row_country_idx + 1) * n_fd_categories
    
    row_zcol_idx  = np.arange(row_ci_s, row_ci_e)
    row_row_idx   = row_zcol_idx
    non_zcol_idx  = np.r_[0:row_ci_s, row_ci_e:n_ci]
    row_fdcol_idx = np.arange(row_fd_s, row_fd_e)
    non_fdcol_idx = np.r_[0:row_fd_s, row_fd_e:n_fd]
    
    # -----------------------------------------------------------------------
    # Step 4: 固定 ROW 列,准备子矩阵
    # -----------------------------------------------------------------------
    if verbose:
        print("[4/6] 构造子矩阵并运行 GRAS ...")
    
    X_target = OUT_col.copy()
    v_Z      = X_target - VA_new - TLS_z_new
    
    # ROW 固定列
    Z_row_fix  = Z[:, row_zcol_idx].copy()
    FD_row_fix = FD[:, row_fdcol_idx].copy()
    row_fix_contrib = Z_row_fix.sum(axis=1) + FD_row_fix.sum(axis=1)
    
    # 非 ROW 行目标
    u_sub = X_target - row_fix_contrib
    
    # 缩减过大的 ROW 固定值
    neg_mask = u_sub < -1e-9
    if np.any(neg_mask):
        if verbose:
            print(f"       警告: {np.sum(neg_mask)} 行 ROW 固定列贡献超行目标")
        for i in np.where(neg_mask)[0]:
            fix_sum = row_fix_contrib[i]
            if fix_sum > 1e-15:
                scale_fix = max(X_target[i] / fix_sum, 0.0)
                Z_row_fix[i, :]  *= scale_fix
                FD_row_fix[i, :] *= scale_fix
        row_fix_contrib = Z_row_fix.sum(axis=1) + FD_row_fix.sum(axis=1)
        u_sub = X_target - row_fix_contrib
    
    # 子矩阵
    M_sub = np.hstack([Z[:, non_zcol_idx], FD[:, non_fdcol_idx]])
    v_Z_sub  = v_Z[non_zcol_idx]
    v_FD_sub = FD.sum(axis=0)[non_fdcol_idx]
    
    # 对齐检查
    gap = u_sub.sum() - (v_Z_sub.sum() + v_FD_sub.sum())
    if abs(gap) > 1e-9:
        v_FD_sub = v_FD_sub + gap * (v_FD_sub / max(v_FD_sub.sum(), 1e-15))
    
    # -----------------------------------------------------------------------
    # Step 5: 选择并运行 GRAS 策略
    # -----------------------------------------------------------------------
    if verbose:
        print(f"       运行策略 {strategy.name} ...")
    
    # 分解 P, N
    M_sub_safe = M_sub.copy()
    M_sub_safe[~np.isfinite(M_sub_safe)] = 0.0
    P = np.where(M_sub_safe > 0, M_sub_safe, 0.0)
    N = np.where(M_sub_safe < 0, -M_sub_safe, 0.0)
    
    # 根据策略选择 GRAS 函数
    gras_funcs = {
        RebalanceStrategy.ORIGINAL_GRAS: gras_iteration_original,
        RebalanceStrategy.WEIGHTED_GRAS: gras_iteration_weighted,
        RebalanceStrategy.BOUNDED_GRAS: gras_iteration_bounded,
        RebalanceStrategy.AUGMENTED_TARGET: gras_iteration_augmented,
        RebalanceStrategy.STEPWISE_GRAS: gras_iteration_stepwise,
    }
    
    gras_func = gras_funcs[strategy]
    
    # 调用 GRAS
    M_bal_sub, gras_info = gras_func(
        M_sub, u_sub, v_Z_sub, P, N,
        max_iter=max_iter, tol=tol, verbose=verbose, **strategy_params
    )
    
    # 填充平衡后的子矩阵
    Z_bal = np.zeros_like(Z)
    FD_bal = np.zeros_like(FD)
    
    Z_bal[:, non_zcol_idx]  = M_bal_sub[:, :len(non_zcol_idx)]
    FD_bal[:, non_fdcol_idx] = M_bal_sub[:, len(non_zcol_idx):]
    
    # 保持 ROW 列固定
    Z_bal[:, row_zcol_idx]  = Z_row_fix
    FD_bal[:, row_fdcol_idx] = FD_row_fix
    
    # -----------------------------------------------------------------------
    # Step 5b: ROW 行的残差分配
    # -----------------------------------------------------------------------
    if verbose:
        print(f"       ROW 行残差分配 ...")
    
    # 非 ROW 行的已平衡列和
    non_row_sum = Z_bal.sum(axis=0) + VA_new + TLS_z_new
    row_zcol_residual = X_target - non_row_sum
    
    # 分配到 ROW 行
    for j_idx, j in enumerate(row_zcol_idx):
        Z_bal[row_row_idx[0]:row_row_idx[-1]+1, j] = 0.0
        col_idx = int(j_idx)
        if 0 <= col_idx < len(row_row_idx):
            row_idx = int(row_row_idx[col_idx])
            Z_bal[row_idx, j] = max(float(row_zcol_residual[j]), 0.0)
    
    # 反推 ROW VA
    for j in row_zcol_idx:
        needed_va = X_target[j] - Z_bal[:, j].sum() - TLS_z_new[j]
        if needed_va < 0:
            TLS_z_new[j] = max(TLS_z_new[j] + needed_va, 0.0)
            needed_va = X_target[j] - Z_bal[:, j].sum() - TLS_z_new[j]
        VA_new[j] = max(needed_va, 0.0)
    
    # -----------------------------------------------------------------------
    # Step 6: 诊断与输出
    # -----------------------------------------------------------------------
    if verbose:
        print("[5/6] 诊断 ...")
    
    diag = diagnose_explosion(Z_orig, Z_bal, FD_orig, FD_bal, change_threshold)
    
    if verbose:
        print(f"       最大 Z 变化: {diag['max_z_change']/1e6:.2f}M USD")
        print(f"       最大 FD 变化: {diag['max_fd_change']/1e6:.2f}M USD")
        print(f"       Z 变化率: {diag['z_change_ratio']:.2%}")
        print(f"       >100M 格子数: Z={diag['n_large_z']}, FD={diag['n_large_fd']}")
        print(f"       爆炸检测: {'是' if diag['has_explosion'] else '否'}")
    
    # 检查会计平衡
    row_sum_err = np.max(np.abs(Z_bal.sum(axis=1) + FD_bal.sum(axis=1) - X_target))
    col_sum_err = np.max(np.abs(Z_bal.sum(axis=0) + VA_new + TLS_z_new - X_target))
    
    if verbose:
        print(f"       行和误差: {row_sum_err:.3e}")
        print(f"       列和误差: {col_sum_err:.3e}")
    
    # -----------------------------------------------------------------------
    # Step 7: 重组并保存
    # -----------------------------------------------------------------------
    if verbose:
        print(f"[6/6] 保存到 {output_path} ...")
    
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
    
    out.to_excel(output_path)
    
    if verbose:
        print(f"\n{'='*60}")
        print(f"[✓] 完成 - 策略: {strategy.name}")
        print(f"     GRAS 迭代: {gras_info.get('iterations', 'N/A')}")
        print(f"     收敛: {gras_info.get('converged', False)}")
        print(f"{'='*60}\n")
    
    return out


def compare_strategies(
    input_path: str,
    output_dir: str,
    n_countries: int = 81,
    n_industries: int = 50,
    n_fd_categories: int = 6,
    row_country_idx: int = 80,
    change_threshold: float = 100e6,
    verbose: bool = True,
) -> pd.DataFrame:
    """
    比较所有策略的结果,选择最优策略。
    
    返回比较结果 DataFrame。
    """
    import os
    os.makedirs(output_dir, exist_ok=True)
    
    strategies = [
        (RebalanceStrategy.ORIGINAL_GRAS, {}, "原始GRAS"),
        (RebalanceStrategy.BOUNDED_GRAS, {"tau": 2.0}, "Bounded(tau=2)"),
        (RebalanceStrategy.BOUNDED_GRAS, {"tau": 5.0}, "Bounded(tau=5)"),
        (RebalanceStrategy.WEIGHTED_GRAS, {"weight_mode": "hybrid"}, "Weighted(hybrid)"),
        (RebalanceStrategy.AUGMENTED_TARGET, {"regularization": 1e-4}, "Augmented(reg=1e-4)"),
        (RebalanceStrategy.STEPWISE_GRAS, {"n_steps": 20}, "Stepwise(20步)"),
    ]
    
    results = []
    
    for strategy, params, name in strategies:
        try:
            output_path = os.path.join(output_dir, f"balanced_{name.replace('(', '_').replace(')', '')}.xlsx")
            
            df_out = rebalance_icio_advanced(
                input_path=input_path,
                output_path=output_path,
                n_countries=n_countries,
                n_industries=n_industries,
                n_fd_categories=n_fd_categories,
                row_country_idx=row_country_idx,
                strategy=strategy,
                strategy_params=params,
                change_threshold=change_threshold,
                verbose=False,
            )
            
            # 读取原始数据用于诊断
            df_orig = pd.read_excel(input_path, index_col=0)
            n_ci = n_countries * n_industries
            n_fd = n_countries * n_fd_categories
            
            Z_orig = df_orig.iloc[:n_ci, :n_ci].to_numpy(dtype=np.float64, copy=True)
            FD_orig = df_orig.iloc[:n_ci, n_ci:n_ci+n_fd].to_numpy(dtype=np.float64, copy=True)
            Z_bal = df_out.iloc[:n_ci, :n_ci].to_numpy(dtype=np.float64, copy=True)
            FD_bal = df_out.iloc[:n_ci, n_ci:n_ci+n_fd].to_numpy(dtype=np.float64, copy=True)
            
            diag = diagnose_explosion(Z_orig, Z_bal, FD_orig, FD_bal, change_threshold)
            
            results.append({
                '策略': name,
                '最大Z变化(M USD)': diag['max_z_change'] / 1e6,
                '最大FD变化(M USD)': diag['max_fd_change'] / 1e6,
                'Z变化率(%)': diag['z_change_ratio'] * 100,
                '>100M格子数(Z)': diag['n_large_z'],
                '>100M格子数(FD)': diag['n_large_fd'],
                '爆炸': '是' if diag['has_explosion'] else '否',
                '输出文件': os.path.basename(output_path),
            })
            
            if verbose:
                print(f"{name}: Z变化={diag['max_z_change']/1e6:.1f}M, "
                      f"变化率={diag['z_change_ratio']:.1%}, "
                      f"爆炸={'是' if diag['has_explosion'] else '否'}")
                      
        except Exception as e:
            if verbose:
                print(f"{name}: 失败 - {str(e)[:50]}")
            results.append({
                '策略': name,
                '最大Z变化(M USD)': np.nan,
                '最大FD变化(M USD)': np.nan,
                'Z变化率(%)': np.nan,
                '>100M格子数(Z)': np.nan,
                '>100M格子数(FD)': np.nan,
                '爆炸': '未知',
                '输出文件': '失败',
            })
    
    df_results = pd.DataFrame(results)
    
    # 保存比较结果
    summary_path = os.path.join(output_dir, "strategy_comparison.xlsx")
    df_results.to_excel(summary_path, index=False)
    
    if verbose:
        print(f"\n比较结果已保存到: {summary_path}")
        
        # 找出最佳策略
        valid_results = df_results[df_results['爆炸'] == '否']
        if len(valid_results) > 0:
            best_idx = valid_results['最大Z变化(M USD)'].idxmin()
            best = valid_results.loc[best_idx]
            print(f"\n推荐策略: {best['策略']}")
            print(f"  - 最大Z变化: {best['最大Z变化(M USD)']:.1f}M USD")
            print(f"  - 变化率: {best['Z变化率(%)']:.1f}%")
        else:
            # 选择变化最小的策略
            min_idx = df_results['最大Z变化(M USD)'].idxmin()
            best = df_results.loc[min_idx]
            print(f"\n警告:所有策略都存在爆炸,推荐变化最小的: {best['策略']}")
            print(f"  - 最大Z变化: {best['最大Z变化(M USD)']:.1f}M USD")
    
    return df_results


# ============================================================================
# 4. 命令行入口
# ============================================================================

if __name__ == "__main__":
    import os
    
    # 配置
    INPUT_PATH = "modified_ICIO.xlsx"
    OUTPUT_DIR = "rebalance_test"
    
    # 检查输入文件
    if not os.path.exists(INPUT_PATH):
        print(f"错误:找不到输入文件 {INPUT_PATH}")
        print("请确保 modified_ICIO.xlsx 在当前目录下")
        exit(1)
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # 首先比较所有策略
    print("\n" + "="*70)
    print("  策略比较:测试所有防爆炸策略")
    print("="*70 + "\n")
    
    compare_strategies(
        input_path=INPUT_PATH,
        output_dir=OUTPUT_DIR,
        n_countries=81,
        n_industries=50,
        n_fd_categories=6,
        row_country_idx=80,
        change_threshold=100e6,  # 100 million
        verbose=True,
    )
    
    # 然后使用推荐的 Bounded GRAS 策略生成最终结果
    print("\n" + "="*70)
    print("  生成最终结果: Bounded GRAS (tau=5.0)")
    print("="*70 + "\n")
    
    OUTPUT_PATH = os.path.join(OUTPUT_DIR, "balanced_final.xlsx")
    
    rebalance_icio_advanced(
        input_path=INPUT_PATH,
        output_path=OUTPUT_PATH,
        n_countries=81,
        n_industries=50,
        n_fd_categories=6,
        row_country_idx=80,
        tol=1e-6,
        max_iter=2000,
        strategy=RebalanceStrategy.BOUNDED_GRAS,
        strategy_params={"tau": 5.0},
        change_threshold=100e6,
        verbose=True,
    )