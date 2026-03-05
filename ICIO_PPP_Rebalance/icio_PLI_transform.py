"""
icio_pli_transform.py
=====================
功能：
    1. 读取 OECD-ICIO 表格（2022_SML.xlsx 的首个 sheet）
    2. 读取 PLI 数据（PLO.xlsx 的 "OECD_PLI" sheet，E 列，共 81 个国家）
    3. 对 ICIO 表中每个国家的所有行（Z 区域 + FD 区域）乘以该国对应的 PLI 因子
    4. 输出转化后的 modified_ICIO（保存为 Excel）

ICIO 表结构（参见示意图）：
    - 行/列索引格式：COUNTRY_INDUSTRY（如 AUS_C01）
    - Z  区域：(4050 × 4050)  中间投入矩阵，81国 × 50行业
    - FD 区域：(4050 × 486)   最终需求矩阵，81国 × 6个最终需求部门
      最终需求列名格式：COUNTRY_FINALDEMAND（如 AUS_HFCE）
      六个部门：HFCE / NPISH / GGFC / GFCF / INVNT / R（调整项）
    - X  行：(1 × 4050)  各行业总产出（行标题 "OUTPUT"）
    - TLS 行：(1 × 4050 + 1 × 486)  税收净补贴（中间 + 最终需求）
    - VA 行：(1 × 4050)  增加值
    - 右下角 0：最终需求区域下方的税收 / 增加值块为 0

PLI（Price Level Index）：
    - 来自 PLO.xlsx，"OECD_PLI" sheet，E 列
    - 共 81 个国家，每行对应一个国家的 PLI 值
    - PLI 用作行缩放因子：将某国所有生产行（Z 行 + 对应 FD 行）乘以该国 PLI

转化逻辑：
    对于 ICIO 表中属于国家 c 的每一行（行索引前缀为 c_），
    该行的所有数值 × PLI[c]
    （X / VA / TLS 等汇总行不参与缩放，或可根据需求单独处理）
"""

import re
from pathlib import Path
import pandas as pd
import numpy as np


# ============================================================
# 1. 辅助函数
# ============================================================

def extract_country_from_label(label: str) -> str:
    """
    从行/列标签中提取国家代码。
    标签格式均为 COUNTRY_XXX，取下划线前的部分作为国家代码。
    示例：
        "AUS_C01"   -> "AUS"
        "USA_HFCE"  -> "USA"
    """
    parts = str(label).split("_", maxsplit=1)
    return parts[0] if len(parts) > 1 else label


def read_icio_table(filepath: str) -> pd.DataFrame:
    """
    读取 OECD-ICIO Excel 文件的首个 sheet。

    参数：
        filepath : ICIO Excel 文件路径（如 "2022_SML.xlsx"）

    返回：
        df : 以行标签为索引、列标签为列名的 DataFrame
             （行列标签均为字符串，如 "AUS_C01"、"AUS_HFCE" 等）

    注意：
        - 首行为列标题（header=0）
        - 首列为行标题（index_col=0）
        - 数值区域可能含 NaN（对应示意图中的 0 块），读取后填充为 0
    """
    print(f"[读取 ICIO] 文件：{filepath}")
    df = pd.read_excel(
        filepath,
        sheet_name=0,          # 首个 sheet
        header=0,              # 第 0 行为列标题
        index_col=0,           # 第 0 列为行索引
        dtype={0: str},        # 行标签强制为字符串
        engine="openpyxl",
    )

    # 确保行/列标签均为字符串（防止数字标签干扰后续匹配）
    df.index = df.index.astype(str)
    df.columns = df.columns.astype(str)

    # 将缺失值（NaN）填充为 0，与示意图中空白块含义一致
    df = df.fillna(0)

    print(f"  -> 读取完成，表格尺寸：{df.shape}（行 × 列）")
    return df


def read_pli_data(filepath: str, sheet_name: str = "OECD_PLI",
                  country_col_letter: str = "E") -> dict:
    """
    读取 PLI 数据文件。

    参数：
        filepath          : PLI Excel 文件路径（如 "PLO.xlsx"）
        sheet_name        : 目标 sheet 名称，默认 "OECD_PLI"
        country_col_letter: PLI 数值所在列字母，默认 "E"

    返回：
        pli_dict : {国家代码(str) -> PLI值(float)} 的字典
                   共 81 个国家

    假设：
        - PLI sheet 中存在一列国家代码（用于对应 ICIO 行标签的国家前缀）
          通常为 A 列或 B 列；脚本自动查找含有 3 位大写字母代码的列
        - E 列为 PLI 数值

    注意：
        - 若文件布局与假设不符，请调整下方 `country_col_idx` 的赋值逻辑
    """
    print(f"[读取 PLI] 文件：{filepath}，Sheet：{sheet_name}")

    # 读取整个 sheet（不指定 index_col，便于灵活查找国家列）
    raw = pd.read_excel(
        filepath,
        sheet_name=sheet_name,
        header=0,
        engine="openpyxl",
    )

    print(f"  -> 原始 sheet 列名：{list(raw.columns)}")
    print(f"  -> 原始 sheet 尺寸：{raw.shape}")

    # ---- 定位国家代码列 ----
    # 策略：遍历各列，找到含 3 位纯大写字母值最多的列作为国家代码列
    country_col_name = None
    max_match = 0
    iso3_pattern = re.compile(r"^[A-Z]{3}$")

    for col in raw.columns:
        col_series = raw[col].dropna().astype(str)
        match_count = int((col_series.str.match(iso3_pattern)).sum())
        # 确保类型一致以避免 numpy 类型比较错误
        if int(match_count) > int(max_match):
            max_match = int(match_count)
            country_col_name = col

    if country_col_name is None:
        raise ValueError(
            "无法自动定位国家代码列，请检查 PLI sheet 结构并手动指定。"
        )
    print(f"  -> 自动识别国家代码列：'{country_col_name}'（匹配 {max_match} 行）")

    # ---- 定位 PLI 数值列 ----
    # 将列字母（如 "E"）转换为 pandas 列位置索引（0-based）
    # Excel 列字母：A=0, B=1, ..., E=4
    col_letter_upper = country_col_letter.upper()
    col_idx = ord(col_letter_upper) - ord("A")  # E -> 4

    if col_idx >= len(raw.columns):
        raise ValueError(
            f"列 '{country_col_letter}' 超出 sheet 列数范围，请检查 PLI 文件。"
        )
    pli_col_name = raw.columns[col_idx]
    print(f"  -> PLI 数值列（第 {col_idx+1} 列）：'{pli_col_name}'")

    # ---- 构建国家 -> PLI 映射字典 ----
    pli_df = raw[[country_col_name, pli_col_name]].dropna(subset=[country_col_name])
    pli_df = pli_df[pli_df[country_col_name].astype(str).str.match(iso3_pattern)]
    pli_df = pli_df.rename(columns={country_col_name: "country", pli_col_name: "pli"})
    pli_df["country"] = pli_df["country"].astype(str).str.strip()
    pli_df["pli"] = pd.to_numeric(pli_df["pli"], errors="coerce")

    # 检查是否有缺失 PLI 值
    missing_pli = pli_df[pli_df["pli"].isna()]
    if not missing_pli.empty:
        print(f"  [警告] 以下国家的 PLI 值缺失，将默认填充为 1.0（不缩放）：")
        print(f"         {list(missing_pli['country'])}")
        pli_df["pli"] = pli_df["pli"].fillna(1.0)

    pli_dict = dict(zip(pli_df["country"], pli_df["pli"]))
    print(f"  -> 成功读取 {len(pli_dict)} 个国家的 PLI 值")
    return pli_dict


# ============================================================
# 2. 核心转化函数
# ============================================================

def identify_icio_row_sections(df: pd.DataFrame) -> dict:
    """
    根据行标签识别 ICIO 表各功能区段。

    ICIO 表行标签规律：
        - 中间投入行（Z）：格式 COUNTRY_INDUSTRY（如 "AUS_C01"），共 4050 行
        - 税收净补贴行（TLS）：标签含 "TLS"（如 "TLS"、"TAXSUB" 等）
        - 增加值行（VA）：标签含 "VA"（如 "VA"、"VALU" 等）
        - 总产出行（X）：标签含 "OUTPUT" 或 "TOTAL" 等

    返回：
        sections : {
            "Z_rows"   : [行标签列表] — 中间投入行（国家_行业格式）
            "TLS_rows" : [行标签列表] — 税收净补贴行
            "VA_rows"  : [行标签列表] — 增加值行
            "X_rows"   : [行标签列表] — 总产出行
            "other_rows": [行标签列表] — 其他行
        }

    注意：此函数通过启发式规则判断，如实际文件标签格式不同，请调整 patterns。
    """
    iso3_pattern = re.compile(r"^[A-Z]{3}_")   # 以 3 位国家码 + 下划线开头

    Z_rows, TLS_rows, VA_rows, X_rows, other_rows = [], [], [], [], []

    for label in df.index:
        label_up = str(label).upper()
        if iso3_pattern.match(str(label)):
            Z_rows.append(label)                # 国家_行业 格式 -> 中间投入行
        elif "TLS" in label_up or "TAXSUB" in label_up:
            TLS_rows.append(label)
        elif "VA" in label_up or "VALU" in label_up:
            VA_rows.append(label)
        elif "OUTPUT" in label_up or "TOTAL" in label_up or label_up == "X":
            X_rows.append(label)
        else:
            other_rows.append(label)

    sections = {
        "Z_rows": Z_rows,
        "TLS_rows": TLS_rows,
        "VA_rows": VA_rows,
        "X_rows": X_rows,
        "other_rows": other_rows,
    }

    print(f"\n[行区段识别]")
    for k, v in sections.items():
        print(f"  {k:<12}: {len(v)} 行")
    return sections


def apply_pli_to_icio(df: pd.DataFrame, pli_dict: dict,
                      scale_z_only: bool = True) -> pd.DataFrame:
    """
    对 ICIO 表应用 PLI 因子，返回 modified_ICIO。

    转化规则（按照价格水平指数调整生产侧）：
        对于 ICIO 中属于国家 c 的每一条生产行（行标签前缀为 c）：
            该行所有列的数值  ×=  PLI[c]

        "生产行"指中间投入矩阵 Z 的行（格式 COUNTRY_INDUSTRY）。
        X / VA / TLS 等汇总行默认不参与缩放（scale_z_only=True）；
        若需一并缩放，设 scale_z_only=False。

    参数：
        df           : 原始 ICIO DataFrame（行列均以字符串标签）
        pli_dict     : {国家代码 -> PLI值} 字典
        scale_z_only : True  -> 仅缩放 Z 区域行（推荐）
                       False -> 缩放所有以国家代码开头的行（含 VA / X 等）

    返回：
        modified_df : 缩放后的 ICIO DataFrame（与原表同形）
    """
    print(f"\n[应用 PLI 因子] scale_z_only={scale_z_only}")

    # 复制一份，避免修改原始数据
    modified_df = df.copy()

    # 识别行区段
    sections = identify_icio_row_sections(df)
    rows_to_scale = sections["Z_rows"] if scale_z_only else (
        sections["Z_rows"] + sections["other_rows"]
    )

    # 按国家分组缩放
    iso3_pattern = re.compile(r"^[A-Z]{3}_")
    scaled_countries = set()
    skipped_countries = set()

    for row_label in rows_to_scale:
        country = extract_country_from_label(row_label)

        if country not in pli_dict:
            # 该国家在 PLI 字典中找不到，记录并跳过（不缩放）
            skipped_countries.add(country)
            continue

        pli_value = pli_dict[country]
        # 将该行所有数值乘以 PLI 因子（in-place 乘法，效率更高）
        modified_df.loc[row_label] *= pli_value
        scaled_countries.add(country)

    print(f"  -> 已缩放国家数：{len(scaled_countries)}")
    if skipped_countries:
        print(f"  [警告] 以下国家在 PLI 字典中未找到，未执行缩放：")
        print(f"         {sorted(skipped_countries)}")

    return modified_df


# ============================================================
# 3. 输出函数
# ============================================================

def save_modified_icio(modified_df: pd.DataFrame, output_path: str) -> None:
    """
    将 modified_ICIO 保存为 Excel 文件。

    参数：
        modified_df  : apply_pli_to_icio() 返回的 DataFrame
        output_path  : 输出文件路径（如 "modified_ICIO.xlsx"）
    """
    print(f"\n[保存结果] 输出文件：{output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        modified_df.to_excel(writer, sheet_name="modified_ICIO")
    print(f"  -> 已保存，尺寸：{modified_df.shape}")


def print_sanity_check(original_df: pd.DataFrame,
                       modified_df: pd.DataFrame,
                       pli_dict: dict,
                       sample_country: str = None) -> None:
    """
    简单的健全性检查：
        选取一个示例国家，打印该国第一行在缩放前后的数值差异，
        并验证缩放比例是否等于 PLI 值。

    参数：
        original_df    : 原始 ICIO
        modified_df    : 缩放后 ICIO
        pli_dict       : PLI 字典
        sample_country : 抽查的国家代码（默认取 pli_dict 中第一个）
    """
    if not pli_dict:
        print("[健全性检查] PLI 字典为空，跳过。")
        return

    if sample_country is None:
        sample_country = next(iter(pli_dict))

    # 找到该国家的第一个 Z 行
    target_rows = [r for r in original_df.index
                   if str(r).startswith(f"{sample_country}_")]
    if not target_rows:
        print(f"[健全性检查] 未找到国家 '{sample_country}' 的行，跳过。")
        return

    sample_row = target_rows[0]
    orig_vals = original_df.loc[sample_row].values.astype(float)
    mod_vals = modified_df.loc[sample_row].values.astype(float)
    pli_val = pli_dict[sample_country]

    # 仅比较非零元素，防止 0/0
    nonzero_mask = orig_vals != 0
    if nonzero_mask.sum() == 0:
        print(f"[健全性检查] 行 '{sample_row}' 全为 0，无法验证比值。")
        return

    ratios = mod_vals[nonzero_mask] / orig_vals[nonzero_mask]
    ratio_mean = np.mean(ratios)
    ratio_std = np.std(ratios)

    print(f"\n[健全性检查] 抽查国家：{sample_country}，PLI = {pli_val:.4f}")
    print(f"  行标签：{sample_row}")
    print(f"  缩放比例（modified / original）均值：{ratio_mean:.6f}，标准差：{ratio_std:.2e}")
    if abs(ratio_mean - pli_val) < 1e-6:
        print(f"  ✓ 缩放比例与 PLI 值一致，转化正确。")
    else:
        print(f"  ✗ 缩放比例与 PLI 值不一致！请检查代码。")


# ============================================================
# 4. 主流程
# ============================================================

def main():
    # ---- 文件路径配置 ----
    # 获取脚本所在目录
    script_dir = Path(__file__).parent
    
    ICIO_FILE = script_dir / "2022_SML.xlsx"      # ICIO 表文件（首 sheet 即为 ICIO）
    PLI_FILE  = script_dir / "PLI.xlsx"           # PLI 数据文件
    PLI_SHEET = "OECD_PLI"                        # PLI 所在 sheet 名
    PLI_COL   = "E"                               # PLI 数值所在列（Excel 列字母）
    OUTPUT_FILE = script_dir / "modified_ICIO.xlsx"

    # ---- Step 1：读取 ICIO 表 ----
    icio_df = read_icio_table(ICIO_FILE)

    # ---- Step 2：读取 PLI 数据 ----
    pli_dict = read_pli_data(PLI_FILE, sheet_name=PLI_SHEET,
                             country_col_letter=PLI_COL)

    # ---- Step 3：应用 PLI 因子，生成 modified_ICIO ----
    #   scale_z_only=True：仅对中间投入矩阵 Z 的行（格式 COUNTRY_INDUSTRY）进行缩放
    #   若需同时缩放 VA / TLS 等行，改为 scale_z_only=False
    modified_icio_df = apply_pli_to_icio(icio_df, pli_dict, scale_z_only=True)

    # ---- Step 4：健全性检查 ----
    print_sanity_check(icio_df, modified_icio_df, pli_dict)

    # ---- Step 5：保存结果 ----
    save_modified_icio(modified_icio_df, OUTPUT_FILE)

    print("\n✓ 全部流程完成！")


if __name__ == "__main__":
    main()
