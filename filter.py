# -*- coding: utf-8 -*-
"""
质检数据整理工具 v3.1
=================================================================
功能：
  1. 从程序所在目录读取 ad-config.json 获取规则清单；
     读不到/读失败/结构无效 -> 自动回退内置默认配置
  2. 扫描目录下以规则名开头的明细表（前缀模糊匹配），有则处理、无则跳过
  3. 读取《xxxx有效营业厅清单.xlsx》 -> 按门店名匹配，回填 分公司/网格/所属大商/特权
  4. 规则名称列按规则用不同浅色填充区分
=================================================================
"""

import os
import sys

sys.stdout.write("=" * 60 + "\n")
sys.stdout.write("  质检数据整理工具 v2.3\n")
sys.stdout.write("=" * 60 + "\n")
sys.stdout.write(">> 程序正在初始化，请稍候...\n")
sys.stdout.write("   (正在解压依赖库，启动可能需要 20-30 秒)\n")
sys.stdout.write("-" * 60 + "\n")
sys.stdout.flush()

import glob
import json
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
# 路径工具
# ------------------------------------------------------------------
def get_base_dir() -> str:
    """
    EXE 所在目录（打包后）或脚本所在目录（源码运行）。
    PyInstaller --onefile 下 sys.executable 指向用户双击的 EXE 真实路径，
    不受临时解压目录和 cwd 影响。
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))


# ------------------------------------------------------------------
# 配置区
# ------------------------------------------------------------------
CONFIG_FILE = "ad-config.json"

# ★ 内置默认规则清单（配置文件读取失败时使用）
DEFAULT_RULE_FILES = [
    "省统-营销-保险推荐",
    "中山-营销-亲情网",
    "省统-回捞-降档挽留",
    "省统-营销-推荐升档",
    "省统-营销-推升千兆",
    "中山-营销-自动充",
    "省统-服务-服务推诿",
    "省统-服务-服务质量",
    "省统-营销-开新降旧",
    "省统-营销-开新拆旧",
    "服务禁用语",
]

STORE_FILE_PATTERN = "*有效营业厅清单*.xlsx"
OUTPUT_SUFFIX      = "多规则汇总.xlsx"

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
    "标记",
]

MARK_COL_NAME  = "标记"
MARK_COL_WIDTH = 6

QC_HEADER_ROW = 3

MIN_COL_WIDTH = 12
MAX_COL_WIDTH = 60
HEADER_HEIGHT = 30

LIGHT_RED = "FFC7CE"

STORE_NAME_REPLACE = {
    "中山市全球通大厦服营厅": "全球通大厦营业厅",
}

# ★ 规则填充色（浅色系，白底/黑字可读）
RULE_FILL_COLORS = [
    "FCE4D6",  # 浅橙
    "DDEBF7",  # 浅蓝
    "E2EFDA",  # 浅绿
    "FFF2CC",  # 浅黄
    "EDE7F6",  # 浅紫
    "FDE9E9",  # 浅粉
    "DEEAF6",  # 浅天蓝
    "F2F2F2",  # 浅灰
    "FBE5D6",  # 浅杏
    "E4DFEC",  # 浅藕荷
    "D9E1F2",  # 浅蓝灰
]


# ------------------------------------------------------------------
# 配置加载：只从程序所在目录读 ad-config.json，失败一律回退默认
# ------------------------------------------------------------------
def load_rule_files(base_dir: str):
    """
    返回 (rule_files, source_desc)
      - 只读取 <base_dir>/ad-config.json
      - 不存在 / JSON 语法错 / 顶层非对象 / 缺 rule_files / 数组清洗后为空
        -> 一律回退 DEFAULT_RULE_FILES
    """
    path = os.path.join(base_dir, CONFIG_FILE)

    # ---- 情况 1：文件不存在 ----
    if not os.path.isfile(path):
        print(f"[配置] 未找到配置文件：{path}")
        print(f"[配置] 使用内置默认规则清单")
        return DEFAULT_RULE_FILES[:], "内置默认"

    # ---- 情况 2：读取 / 解析失败 ----
    print(f"[配置] 读取：{path}")
    try:
        # utf-8-sig 兼容记事本/VS Code 保存的带 BOM 文件
        with open(path, "r", encoding="utf-8-sig") as f:
            cfg = json.load(f)
    except Exception as e:
        print(f"[配置] 读取失败：{e}")
        print(f"[配置] 已回退到内置默认规则清单")
        return DEFAULT_RULE_FILES[:], "内置默认（读取失败）"

    # ---- 情况 3：结构无效 ----
    if not isinstance(cfg, dict):
        print(f"[配置] 顶层不是对象，已回退内置默认")
        return DEFAULT_RULE_FILES[:], "内置默认（结构错误）"

    rules = cfg.get("rule_files")
    if not isinstance(rules, list):
        print(f"[配置] 缺少 rule_files 数组，已回退内置默认")
        return DEFAULT_RULE_FILES[:], "内置默认（缺少 rule_files）"

    # ---- 情况 4：清洗（去空、去重、转字符串）----
    cleaned, seen = [], set()
    for r in rules:
        s = str(r).strip()
        if s and s not in seen:
            seen.add(s)
            cleaned.append(s)

    if not cleaned:
        print(f"[配置] rule_files 清洗后为空，已回退内置默认")
        return DEFAULT_RULE_FILES[:], "内置默认（rule_files 为空）"

    return cleaned, path


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


def clean_store_display(name: str) -> str:
    s = clean_text(name)
    if not s:
        return ""
    return STORE_NAME_REPLACE.get(s, s)


def norm_store_name(name: str) -> str:
    s = clean_store_display(name)
    if not s:
        return ""
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[\s\-—_()（）【】\[\]、,，.。/\\]", "", s)
    s = s.replace("服营厅", "营业厅").replace("服务厅", "营业厅")
    return s


def find_col(cols, target):
    for i, c in enumerate(cols):
        if str(c).strip() == target:
            return i
    return None


def display_width(text: str) -> int:
    w = 0
    for ch in str(text):
        w += 2 if ord(ch) > 127 else 1
    return w


# ------------------------------------------------------------------
# 1. 读取质检明细表 & 抽取片段
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
    out = pd.DataFrame(index=df.index)

    def get(col):
        return df[col] if col in df.columns else pd.Series([""] * len(df), index=df.index)

    out["所在门店"]   = get("所在门店").map(clean_store_display)
    out["所属大商"]   = ""
    out["岗位名称"]   = get("岗位名称").map(clean_text)
    out["声音来源"]   = get("声音来源").map(clean_text)
    out["员工"]       = "员工"
    out["分公司"]     = ""
    out["网格"]       = ""
    out["质检记录编号"] = get("质检记录编号").map(clean_text)
    out["规则名称"]   = rule_name

    date_series = (
        pd.to_datetime(get("录音日期"), errors="coerce")
          .dt.strftime("%Y-%m-%d")
          .fillna("")
    )
    out["省公司判定违规推荐时间"] = date_series.apply(lambda x: f"{x}()" if x else "")

    out["特权"]       = ""
    out["标记"]       = ""

    return out[OUTPUT_COLUMNS]


# ------------------------------------------------------------------
# 2. 营业厅清单索引
# ------------------------------------------------------------------
def build_store_index(store_file: str) -> dict:
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
            if not key or key in index:
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
    key = norm_store_name(store_name)
    if not key:
        return None
    if key in index:
        return index[key]
    keys = list(index.keys())
    hit = get_close_matches(key, keys, n=1, cutoff=0.85)
    if hit:
        return index[hit[0]]
    for k, v in index.items():
        if len(k) >= 4 and len(key) >= 4 and (k in key or key in k):
            return v
    return None


# ------------------------------------------------------------------
# 3. 扫描规则文件（前缀模糊匹配）
# ------------------------------------------------------------------
def discover_rule_files(base_dir: str, rule_files: list) -> list:
    """
    扫描 base_dir 下所有 .xlsx，文件名以某个规则名开头即视为该规则。
    返回 [(规则名, 文件路径), ...]，按 rule_files 顺序排列。
    """
    all_xlsx = []
    for p in glob.glob(os.path.join(base_dir, "*.xlsx")):
        fname = os.path.basename(p)
        if fname.startswith("~$"):          # Excel 临时文件
            continue
        if fname.endswith("有效营业厅清单.xlsx"):
            continue
        if fname.endswith(OUTPUT_SUFFIX):
            continue
        all_xlsx.append(p)
    all_xlsx.sort()

    ordered = []
    for rule in rule_files:
        matched = [p for p in all_xlsx
                   if os.path.splitext(os.path.basename(p))[0].startswith(rule)]
        if matched:
            ordered.append((rule, matched[0]))
    return ordered


def get_output_mmdd(parts: list) -> str:
    """从所有规则片段里取最大录音日期，取不到用昨天"""
    best = None
    for df in parts:
        try:
            s = df.get("省公司判定违规推荐时间", pd.Series(dtype=str))
            d = pd.to_datetime(s.str.replace("()", "", regex=False), errors="coerce").dropna()
            if len(d):
                m = d.max()
                if best is None or m > best:
                    best = m
        except Exception:
            pass
    if best is not None:
        return best.strftime("%m%d")
    return (datetime.now() - timedelta(days=1)).strftime("%m%d")


# ------------------------------------------------------------------
# 4. 格式化（规则列填充着色）
# ------------------------------------------------------------------
def format_worksheet(ws, df: pd.DataFrame, rule_color_map: dict):
    n_rows, n_cols = df.shape

    rule_col_idx = None
    mark_col_idx = None
    for idx, col_name in enumerate(df.columns, start=1):
        name = str(col_name).strip()
        if name == "规则名称":
            rule_col_idx = idx
        elif name == MARK_COL_NAME:
            mark_col_idx = idx

    ws.row_dimensions[1].height = HEADER_HEIGHT

    header_font  = Font(bold=True, size=11)
    header_fill  = PatternFill("solid", fgColor="DDEBF7")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    if mark_col_idx is not None:
        ws.cell(row=1, column=mark_col_idx).value = ""

    for c in range(1, n_cols + 1):
        col_letter = get_column_letter(c)
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

    rule_fills = {
        name: PatternFill("solid", fgColor=color)
        for name, color in rule_color_map.items()
    }

    for r in range(2, n_rows + 2):
        ws.row_dimensions[r].height = 22

        rule_val = ""
        if rule_col_idx is not None:
            rule_val = str(df.iloc[r - 2, rule_col_idx - 1]).strip()

        is_opennew = (rule_val == "省统-营销-开新降旧")

        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment = data_align

        # ★ 规则名称列填充着色
        if rule_col_idx is not None and rule_val in rule_fills:
            ws.cell(row=r, column=rule_col_idx).fill = rule_fills[rule_val]

        if is_opennew and mark_col_idx is not None:
            ws.cell(row=r, column=mark_col_idx).fill = light_red_fill

    ws.freeze_panes = "A2"


# ------------------------------------------------------------------
# 临时目录（避开中文路径风险）
# ------------------------------------------------------------------
def setup_temp_dir():
    """把 openpyxl / tempfile 的临时目录指到系统 temp（英文路径最稳）"""
    try:
        sys_temp = tempfile.gettempdir()
        app_temp = os.path.join(sys_temp, "qc_tool_tmp")
        os.makedirs(app_temp, exist_ok=True)
        os.environ['OPENPYXL_TMP_DIR'] = app_temp
        tempfile.tempdir = app_temp
        return app_temp
    except Exception as e:
        print(f"[警告] 设置临时目录失败，使用系统默认：{e}")
        return None


# ------------------------------------------------------------------
# 主流程
# ------------------------------------------------------------------
def main():
    start_time = time.time()

    base_dir = get_base_dir()
    os.chdir(base_dir)
    print(f"\n[工作目录] {base_dir}\n")

    # ============ 加载规则清单（优先配置文件） ============
    rule_files_cfg, cfg_source = load_rule_files(base_dir)
    print(f"[配置] 规则清单共 {len(rule_files_cfg)} 项，来源：{cfg_source}\n")

    # ============ 临时目录 ============
    temp_dir = setup_temp_dir()

    # ============ 步骤 1：扫描规则文件 ============
    print("=" * 60)
    print("步骤 1/3：扫描规则明细表（前缀模糊匹配）")
    print("=" * 60)
    rule_files = discover_rule_files(base_dir, rule_files_cfg)
    if not rule_files:
        print("[错误] 未找到任何规则明细表。")
        print(f"   请在 EXE 同目录放置以下列名称开头的 xlsx 文件：")
        for n in rule_files_cfg:
            print(f"     - {n}*.xlsx")
        input("按 Enter 键退出...")
        return

    print(f"  发现 {len(rule_files)} 个规则文件：")
    for name, path in rule_files:
        print(f"    - {os.path.basename(path)}  →  规则：{name}")
    print()

    # 规则颜色分配：按规则名排序取模，保证稳定
    all_rule_names = sorted({name for name, _ in rule_files})
    rule_color_map = {
        name: RULE_FILL_COLORS[i % len(RULE_FILL_COLORS)]
        for i, name in enumerate(all_rule_names)
    }

    # ============ 步骤 2：逐个读取并抽取 ============
    print("=" * 60)
    print("步骤 2/3：读取各规则明细表并抽取")
    print("=" * 60)
    parts = []
    for name, path in rule_files:
        try:
            df = read_qc_sheet(path)
            part = build_part(df, name)
            parts.append(part)
            print(f"  [成功] {name}：{len(part)} 条")
        except Exception as e:
            print(f"  [失败] {name}：{e}")
    if not parts:
        print("[错误] 所有规则文件都读取失败。")
        input("按 Enter 键退出...")
        return

    result = pd.concat(parts, ignore_index=True)
    print(f"\n[汇总] 合并后共 {len(result)} 条记录\n")

    # ============ 步骤 3：营业厅清单回填 ============
    print("=" * 60)
    print("步骤 3/3：匹配《有效营业厅清单》，回填 分公司/网格/所属大商/特权")
    print("=" * 60)
    store_files = glob.glob(os.path.join(base_dir, STORE_FILE_PATTERN))
    if not store_files:
        print("  [!] 未找到《xxxx有效营业厅清单.xlsx》，跳过回填\n")
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

    # ============ 写出 ============
    result = result.astype(str).replace({"nan": "", "NaT": "", "None": ""})

    mmdd = get_output_mmdd(parts)
    out_name = f"{mmdd}{OUTPUT_SUFFIX}"
    out_path = os.path.join(base_dir, out_name)

    print("=" * 60)
    print("正在写出 Excel 并应用格式化...")
    print("=" * 60)
    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            result.to_excel(writer, sheet_name="Sheet1", index=False)
            ws = writer.sheets["Sheet1"]
            format_worksheet(ws, result, rule_color_map)
        print(f"  [成功] 输出文件：{out_name}\n")
    except PermissionError:
        print(f"[错误] 无法写入 {out_name}，请先关闭该文件再重试。")
        input("按 Enter 键退出...")
        return
    except Exception as e:
        print(f"[错误] 写出失败：{e}")
        input("按 Enter 键退出...")
        return

    # 清理临时目录（系统 temp 下，删不掉也不影响）
    try:
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass

    elapsed = time.time() - start_time
    print("\n" + "=" * 60)
    print("[完成] 所有任务执行完成！")
    print(f"[耗时] {elapsed:.2f} 秒")
    print(f"[输出] {out_name}")
    print("=" * 60 + "\n")
    input("按 Enter 键退出...")


if __name__ == "__main__":
    main()
