# -*- coding: utf-8 -*-
"""
质检数据整理工具 v1.1
=================================================================
功能：
  1. 读取《服务推诿.xlsx》  -> 提取字段（规则名：省统管控-服务推诿）
  2. 读取《开新降旧.xlsx》  -> 追加（规则名：省统管控-开新降旧（优化试点））
  3. 读取《xxxx有效营业厅清单.xlsx》 -> 按门店名匹配，回填 分公司/网格/所属大商/特权
=================================================================
"""

import os
import sys

# ===== 第一时间输出友好提示 =====
sys.stdout.write("=" * 60 + "\n")
sys.stdout.write("  质检数据整理工具 v1.1\n")
sys.stdout.write("=" * 60 + "\n")
sys.stdout.write(">> 程序正在初始化，请稍候...\n")
sys.stdout.write("   (正在解压依赖库，启动可能需要 20-30 秒)\n")
sys.stdout.write("-" * 60 + "\n")
sys.stdout.flush()


# ===== 现在开始导入标准库和第三方库 =====
import glob
import re
import time
import shutil
import tempfile
from datetime import datetime, timedelta
from difflib import get_close_matches

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ------------------------------------------------------------------
# 路径工具（兼容源码运行和 PyInstaller 打包）
# ------------------------------------------------------------------
def get_base_dir():
    """获取可执行文件所在目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


# ------------------------------------------------------------------
# 配置区
# ------------------------------------------------------------------
SERVICE_FILE       = "服务推诿.xlsx"
OPENNEW_FILE       = "开新降旧.xlsx"
STORE_FILE_PATTERN = "*有效营业厅清单*.xlsx"
OUTPUT_SUFFIX      = "开新降旧+服务推诿.xlsx"

RULE_SERVICE = "省统管控-服务推诿"
RULE_OPENNEW = "省统管控-开新降旧（优化试点）"

OUTPUT_COLUMNS = [
    "所在门店",
    "所属大商",
    "岗位名称",
    "声音来源",
    "员工",
    "分公司",
    "网格",
    "质检记录编号",
    "规则名称",
    "省公司判定违规推荐时间",
    "特权",
    "标记",          # L 列：仅用于填充浅红标记，表头在格式化时置空
]

MARK_COL_NAME  = "标记"
MARK_COL_WIDTH = 6

QC_HEADER_ROW = 3

MIN_COL_WIDTH = 12
MAX_COL_WIDTH = 60
HEADER_HEIGHT = 30

LIGHT_RED = "FFC7CE"


# ------------------------------------------------------------------
# 工具函数
# ------------------------------------------------------------------
def clean_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "nat", "none") else s


def norm_store_name(name: str) -> str:
    """门店名称归一化，用于提高匹配成功率"""
    s = clean_text(name)
    if not s:
        return ""
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[\s\-—_()（）【】\[\]、,，.。/\\]", "", s)
    s = s.replace("服营厅", "营业厅").replace("服务厅", "营业厅")
    s = s.replace("中山市", "").replace("中山", "")
    return s


def find_col(cols, target):
    """在列名列表中返回第一个匹配列的下标"""
    for i, c in enumerate(cols):
        if str(c).strip() == target:
            return i
    return None


def display_width(text: str) -> int:
    """中文按 2 个字符宽算，英文/数字按 1 个算"""
    w = 0
    for ch in str(text):
        w += 2 if ord(ch) > 127 else 1
    return w


# ------------------------------------------------------------------
# 1. 读取质检明细表
# ------------------------------------------------------------------
def read_qc_sheet(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(f"文件不存在：{path}")
    df = pd.read_excel(path, sheet_name=0, header=QC_HEADER_ROW, dtype=object)
    df = df.dropna(how="all").reset_index(drop=True)
    if "质检记录编号" in df.columns:
        df = df[df["质检记录编号"].notna()].reset_index(drop=True)
    return df


def build_part(df: pd.DataFrame, rule_name: str) -> pd.DataFrame:
    """按输出列规则，从质检明细里抽取出结果片段"""
    out = pd.DataFrame(index=df.index)

    def get(col):
        return df[col] if col in df.columns else pd.Series([""] * len(df), index=df.index)

    out["所在门店"]   = get("所在门店").map(clean_text)
    out["所属大商"]   = ""                                   # 后续由营业厅清单回填
    out["岗位名称"]   = get("岗位名称").map(clean_text)
    out["声音来源"]   = get("声音来源").map(clean_text)
    out["员工"]       = "员工"                               # 统一填“员工”
    out["分公司"]     = ""                                   # 后续回填
    out["网格"]       = ""                                   # 后续回填
    out["质检记录编号"] = get("质检记录编号").map(clean_text)
    out["规则名称"]   = rule_name

    # 日期后追加空括号，如 2026-09-10()
    date_series = (
        pd.to_datetime(get("录音日期"), errors="coerce")
          .dt.strftime("%Y-%m-%d")
          .fillna("")
    )
    out["省公司判定违规推荐时间"] = date_series.apply(lambda x: f"{x}()" if x else "")

    out["特权"]       = ""                                   # 后续回填
    out["标记"]       = ""                                   # L 列，仅用于样式标记

    return out[OUTPUT_COLUMNS]


# ------------------------------------------------------------------
# 2. 读取营业厅清单，建立“门店 -> 信息”索引
# ------------------------------------------------------------------
def build_store_index(store_file: str) -> dict:
    """
    读取《xxxx有效营业厅清单》全部 Sheet，建立索引：
        {归一化门店名: {分公司, 网格, 所属大商, 特权}}
    在营表优先（含 所属大商 / 可降档0813），失效表仅补 分公司 / 网格。
    """
    index = {}
    xl = pd.ExcelFile(store_file)

    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(store_file, sheet_name=sheet, header=0, dtype=object)
        except Exception as e:
            print(f"  [跳过] Sheet「{sheet}」读取失败：{e}")
            continue

        df = df.dropna(how="all")
        cols = [str(c).strip() for c in df.columns]

        i_name   = find_col(cols, "营业厅名称")
        i_branch = find_col(cols, "分公司")
        i_grid   = find_col(cols, "网格")
        i_ds     = find_col(cols, "所属大商")
        i_priv   = find_col(cols, "可降档0813")

        if i_name is None:
            print(f"  [跳过] Sheet「{sheet}」没有『营业厅名称』列")
            continue

        cnt = 0
        for _, row in df.iterrows():
            key = norm_store_name(row.iloc[i_name])
            if not key or key in index:      # 已存在则不覆盖（在营优先）
                continue
            index[key] = {
                "分公司":   clean_text(row.iloc[i_branch]) if i_branch is not None else "",
                "网格":     clean_text(row.iloc[i_grid])   if i_grid   is not None else "",
                "所属大商": clean_text(row.iloc[i_ds])     if i_ds     is not None else "",
                "特权":     clean_text(row.iloc[i_priv])   if i_priv   is not None else "",
            }
            cnt += 1
        print(f"  Sheet「{sheet}」载入门店 {cnt} 条")

    return index


def lookup_store(store_name: str, index: dict):
    """用门店名去索引里查（精确 -> 模糊 -> 包含）"""
    key = norm_store_name(store_name)
    if not key:
        return None

    # ① 精确匹配
    if key in index:
        return index[key]

    # ② 模糊匹配
    keys = list(index.keys())
    hit = get_close_matches(key, keys, n=1, cutoff=0.85)
    if hit:
        return index[hit[0]]

    # ③ 包含匹配（长度保护）
    for k, v in index.items():
        if len(k) >= 4 and len(key) >= 4 and (k in key or key in k):
            return v

    return None


def get_output_mmdd(service_df: pd.DataFrame) -> str:
    """优先取录音日期的月份日，取不到则用昨天"""
    try:
        d = pd.to_datetime(service_df.get("录音日期"), errors="coerce").dropna()
        if len(d):
            return d.max().strftime("%m%d")
    except Exception:
        pass
    return (datetime.now() - timedelta(days=1)).strftime("%m%d")


# ------------------------------------------------------------------
# 3. 格式化 Excel
# ------------------------------------------------------------------
def format_worksheet(ws, df: pd.DataFrame):
    """对写入后的工作表做美化"""
    n_rows, n_cols = df.shape

    # ---- 定位列索引 ----
    rule_col_idx = None
    mark_col_idx = None
    for idx, col_name in enumerate(df.columns, start=1):
        name = str(col_name).strip()
        if name == "规则名称":
            rule_col_idx = idx
        elif name == MARK_COL_NAME:
            mark_col_idx = idx

    # ---- 表头行高 ----
    ws.row_dimensions[1].height = HEADER_HEIGHT

    # ---- 表头样式 ----
    header_font  = Font(bold=True, size=11)
    header_fill  = PatternFill("solid", fgColor="DDEBF7")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for c in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # 标记列的表头置空（只保留样式）
    if mark_col_idx is not None:
        ws.cell(row=1, column=mark_col_idx).value = ""

    # ---- 列宽自适应 ----
    for c in range(1, n_cols + 1):
        col_letter = get_column_letter(c)

        # 标记列固定窄宽
        if c == mark_col_idx:
            ws.column_dimensions[col_letter].width = MARK_COL_WIDTH
            continue

        header_text = str(df.columns[c - 1])
        max_w = display_width(header_text)
        for r in range(1, min(n_rows, 500) + 1):
            val = df.iloc[r - 1, c - 1] if r - 1 < n_rows else ""
            w = display_width(str(val))
            if w > max_w:
                max_w = w

        width = max_w + 2
        width = max(MIN_COL_WIDTH, min(MAX_COL_WIDTH, width))
        ws.column_dimensions[col_letter].width = width

    light_red_fill = PatternFill("solid", fgColor=LIGHT_RED)
    data_align = Alignment(vertical="center", wrap_text=True, horizontal="left")

    # ---- 数据单元格 ----
    for r in range(2, n_rows + 2):
        ws.row_dimensions[r].height = 22

        # 判断该行是否为开新降旧
        is_opennew = False
        if rule_col_idx is not None:
            rule_val = str(df.iloc[r - 2, rule_col_idx - 1]).strip()
            is_opennew = (rule_val == RULE_OPENNEW)

        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment = data_align

        # 只给标记列的开新降旧行填充浅红色
        if is_opennew and mark_col_idx is not None:
            ws.cell(row=r, column=mark_col_idx).fill = light_red_fill

    # ---- 冻结首行 ----
    ws.freeze_panes = "A2"


# ------------------------------------------------------------------
# 主流程
# ------------------------------------------------------------------
def main():
    start_time = time.time()

    # 1. 切换工作目录（关键：解决打包后找不到文件/路径问题）
    base_dir = get_base_dir()
    os.chdir(base_dir)
    print(f"\n[工作目录] {base_dir}\n")

    # ===== 重定向 openpyxl 临时目录 =====
    temp_dir = os.path.join(base_dir, "openpyxl_temp")
    try:
        os.makedirs(temp_dir, exist_ok=True)
        os.environ['OPENPYXL_TMP_DIR'] = temp_dir
        tempfile.tempdir = temp_dir
    except Exception as e:
        print(f"[警告] 设置临时目录失败，将使用系统默认：{e}")

    # 2. 检查必需的输入文件
    service_path = os.path.join(base_dir, SERVICE_FILE)
    opennew_path = os.path.join(base_dir, OPENNEW_FILE)

    if not os.path.exists(service_path):
        print(f"[错误] 未找到文件：{SERVICE_FILE}")
        print(f"   请将该文件与 EXE 放在同一目录下。")
        input("按 Enter 键退出...")
        return

    if not os.path.exists(opennew_path):
        print(f"[错误] 未找到文件：{OPENNEW_FILE}")
        print(f"   请将该文件与 EXE 放在同一目录下。")
        input("按 Enter 键退出...")
        return

    # 3. 读取服务推诿
    print("=" * 60)
    print("步骤 1/3：读取《服务推诿.xlsx》")
    print("=" * 60)
    try:
        service_df = read_qc_sheet(service_path)
        part_service = build_part(service_df, RULE_SERVICE)
        print(f"  [成功] 提取 {len(part_service)} 条记录\n")
    except Exception as e:
        print(f"[错误] 读取服务推诿失败：{e}")
        input("按 Enter 键退出...")
        return

    # 4. 读取开新降旧
    print("=" * 60)
    print("步骤 2/3：读取《开新降旧.xlsx》并追加")
    print("=" * 60)
    try:
        opennew_df = read_qc_sheet(opennew_path)
        part_opennew = build_part(opennew_df, RULE_OPENNEW)
        print(f"  [成功] 提取 {len(part_opennew)} 条记录\n")
    except Exception as e:
        print(f"[错误] 读取开新降旧失败：{e}")
        input("按 Enter 键退出...")
        return

    result = pd.concat([part_service, part_opennew], ignore_index=True)
    print(f"[汇总] 合并后共 {len(result)} 条记录\n")

    # 5. 读取营业厅清单并回填
    print("=" * 60)
    print("步骤 3/3：匹配《有效营业厅清单》，回填 分公司/网格/所属大商/特权")
    print("=" * 60)
    store_files = glob.glob(os.path.join(base_dir, STORE_FILE_PATTERN))
    if not store_files:
        print(f"  [!] 未找到《xxxx有效营业厅清单.xlsx》，跳过回填\n")
    else:
        store_file = store_files[0]
        print(f"  使用文件：{os.path.basename(store_file)}")
        try:
            index = build_store_index(store_file)
        except Exception as e:
            print(f"[错误] 读取营业厅清单失败：{e}")
            input("按 Enter 键退出...")
            return

        filled, missing = 0, []
        for i, row in result.iterrows():
            info = lookup_store(row["所在门店"], index)
            if info:
                result.at[i, "所属大商"] = info["所属大商"]
                result.at[i, "分公司"]   = info["分公司"]
                result.at[i, "网格"]     = info["网格"]
                result.at[i, "特权"]     = info["特权"]
                filled += 1
            else:
                missing.append(row["所在门店"])

        print(f"  [成功] 匹配成功 {filled} / {len(result)} 条")
        if missing:
            uniq = sorted(set(missing))
            print(f"  [!] 未匹配到的门店（{len(uniq)} 个）：")
            for m in uniq:
                print(f"      - {m}")
        print()

    # 6. 写出 Excel
    result = result.astype(str).replace({"nan": "", "NaT": "", "None": ""})

    mmdd = get_output_mmdd(service_df)
    out_name = f"{mmdd}{OUTPUT_SUFFIX}"
    out_path = os.path.join(base_dir, out_name)

    print("=" * 60)
    print("正在写出 Excel 并应用格式化...")
    print("=" * 60)
    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            result.to_excel(writer, sheet_name="Sheet1", index=False)
            ws = writer.sheets["Sheet1"]
            format_worksheet(ws, result)
        print(f"  [成功] 输出文件：{out_name}\n")
    except PermissionError:
        print(f"[错误] 无法写入 {out_name}，请先关闭该文件再重试。")
        input("按 Enter 键退出...")
        return
    except Exception as e:
        print(f"[错误] 写出失败：{e}")
        input("按 Enter 键退出...")
        return

    # 7. 清理临时目录
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
    except Exception as e:
        print(f"[警告] 清理临时目录失败：{e}")

    # 8. 完成
    elapsed = time.time() - start_time
    print("\n" + "=" * 60)
    print("[完成] 所有任务执行完成！")
    print(f"[耗时] {elapsed:.2f} 秒")
    print(f"[输出] {out_name}")
    print("=" * 60 + "\n")
    input("按 Enter 键退出...")


if __name__ == "__main__":
    main()