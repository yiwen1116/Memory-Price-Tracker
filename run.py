import io
import os
import re
import glob
import requests
import pandas as pd
from datetime import datetime
import openpyxl

# =======================================================
FOLDER_DIR = "."
# =======================================================

def get_web_data_and_date(url):
    """提取网页数据，统一返回 YYYY-MM-DD 用于内部定位"""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"
    }
    session = requests.Session()
    response = session.get(url, headers=headers, timeout=20)
    response.raise_for_status()
    response.encoding = 'utf-8'
    html_text = response.text

    date_match = re.search(r"更新[^\d]{0,20}(\d{4}[-/]\d{1,2}[-/]\d{1,2})", html_text)
    if date_match:
        real_date = date_match.group(1).replace("/", "-")
    else:
        real_date = datetime.now().strftime("%Y-%m-%d")

    html_data = io.StringIO(html_text)
    tables = pd.read_html(html_data)
    if not tables:
        return [], real_date

    valid_tables = [t for t in tables if len(t.columns) >= 6]
    return valid_tables, real_date


def clean_string(s):
    s = str(s).upper().replace(" ", "").replace("*", "X").replace("（", "(").replace("）", ")").strip()
    s = s.replace("($USD)", "").replace("(USD)", "")
    return s


def get_row_data_dict(tables_list, df_history):
    row_data = {}
    for df_today in tables_list:
        for _, row in df_today.iterrows():
            item_name = str(row.iloc[0]).strip()
            if not item_name or item_name == 'nan' or '日期' in item_name:
                continue

            clean_item = clean_string(item_name)

            for i, col in enumerate(df_history.columns):
                col_group = col[0]
                if 'Unnamed' in str(col_group):
                    continue
                clean_group = clean_string(col_group)

                if clean_item == clean_group or clean_item in clean_group or clean_group in clean_item:
                    indicator = col[1]
                    val = None
                    try:
                        if indicator == '日高点':
                            val = str(row.iloc[1]).strip()
                        elif indicator == '日低点':
                            val = str(row.iloc[2]).strip()
                        elif indicator == '盘高点':
                            val = str(row.iloc[3]).strip()
                        elif indicator == '盘低点':
                            val = str(row.iloc[4]).strip()
                        elif indicator == '盘平均':
                            val = str(row.iloc[5]).strip()
                    except:
                        continue

                    if val and val != 'nan' and val != 'None' and val != '-':
                        row_data[i] = val
    return row_data


def write_to_excel_cell(ws, df_history, row_data_dict, real_date_str):
    """【钢筋混凝土纯物理防线】：完全抛弃 Pandas 多级表头索引，用纯物理空行雷达进行定位追加，绝不误杀覆盖"""
    target_row = None
    system_today_str = datetime.now().strftime("%Y-%m-%d")

    # 1. 🔍 固定死日期列在你的 Excel 中就是第一列 (A列)
    date_col_idx = 1

    # 2. 🔍 用纯物理扫描全表，看看今天系统日期 (2026-05-29) 是否已经存在（满足 18:10 覆写 14:40 保持一行的去重刚需）
    for r in range(3, ws.max_row + 1):
        raw_val = ws.cell(row=r, column=date_col_idx).value
        if raw_val is not None:
            if isinstance(raw_val, datetime):
                cell_val = raw_val.strftime("%Y-%m-%d")
            else:
                try:
                    clean_str = str(raw_val).strip().split()[0].replace("/", "-")
                    cell_val = datetime.strptime(clean_str, "%Y-%m-%d").strftime("%Y-%m-%d")
                except:
                    cell_val = str(raw_val)

            # 如果今天已经跑过一次了，直接锁定该行准备原位更新
            if cell_val == system_today_str:
                target_row = r
                break

    # 3. 🛡️ 【核心硬修复】：如果是全新的今天，用物理空行雷达往下数，直到找到真正的空白格，绝对不相信 ws.max_row
    if target_row is None:
        check_row = 3
        while True:
            cell_content = ws.cell(row=check_row, column=date_col_idx).value
            if cell_content is None or str(cell_content).strip() == "" or str(cell_content).strip() == "None":
                target_row = check_row
                break
            check_row += 1
        print(f"\n 📅 工作表【{ws.title}】：物理防线已启动 -> 避开了所有历史行，锁定了真正的空白物理尾行: 第 {target_row} 行")
    else:
        print(f"\n 📅 工作表【{ws.title}】：今日下午已运行过 -> 正在执行【原位更新】覆盖第 {target_row} 行数据。")

    # 4. ✍️ 统一以 YYYY/M/D 规范将今天的系统日期写入对应的“日期”单元格
    dt_obj = datetime.strptime(system_today_str, "%Y-%m-%d")
    formatted_date = f"{dt_obj.year}/{dt_obj.month}/{dt_obj.day}"
    ws.cell(row=target_row, column=date_col_idx).value = formatted_date

    # 5. 💾 安全填入抓取到的最新价格数据
    for col_idx, value in row_data_dict.items():
        if value is not None and str(value) != 'nan':
            try:
                ws.cell(row=target_row, column=col_idx + 1).value = float(value)
            except ValueError:
                ws.cell(row=target_row, column=col_idx + 1).value = value


def main():
    print("\n==================================================")

    # ---------------- 自动化模块：寻找文件 ---------------- #
    print("🔍 [0/4] 自动雷达：正在定位本地最新的 Excel 追踪表...")
    search_pattern = os.path.join(FOLDER_DIR, "内存价格每日追踪_*.xlsx")
    matched_files = glob.glob(search_pattern)

    if not matched_files:
        print(f"❌ 错误：在 {FOLDER_DIR} 下找不到名字包含【内存价格每日追踪_】的表格！")
        return

    # 按文件名大到小排序，确保永远锁定最新最全的那张表
    current_excel = max(matched_files)
    print(f"📂 成功锁定唯一历史火种文件：【{os.path.basename(current_excel)}】\n")
    # -------------------------------------------------- #

    print("🚀 [1/4] 开始同步 TrendForce 官方数据...")
    try:
        tables_dram, dram_date = get_web_data_and_date("https://www.trendforce.cn/price")
        tables_flash, flash_date = get_web_data_and_date("https://www.trendforce.cn/price/flash")

        print("\n📦 [2/4] 正在进行多级表头数据结构对齐...")
        df_dram_hist = pd.read_excel(current_excel, sheet_name='DRAM Spot Price', header=[0, 1], engine='openpyxl')
        df_flash_hist = pd.read_excel(current_excel, sheet_name='NAND Flash', header=[0, 1], engine='openpyxl')

        dram_row_dict = get_row_data_dict(tables_dram, df_dram_hist)
        flash_row_dict = get_row_data_dict(tables_flash, df_flash_hist)

        print("💾 [3/4] 正在安全注入 Excel 底层单元格...")
        wb = openpyxl.load_workbook(current_excel)
        write_to_excel_cell(wb['DRAM Spot Price'], df_dram_hist, dram_row_dict, dram_date)
        write_to_excel_cell(wb['NAND Flash'], df_flash_hist, flash_row_dict, flash_date)
        wb.save(current_excel)
        wb.close()

        # ---------------- 自动化模块：修改文件名 ---------------- #
        print("🏷️ [4/4] 正在执行智能文件升级与迭代...")
        today_compact = datetime.now().strftime("%Y%m%d")  # 例如：20260529
        new_filename = f"内存价格每日追踪_{today_compact}.xlsx"
        new_filepath = os.path.join(FOLDER_DIR, new_filename)

        if os.path.abspath(current_excel) != os.path.abspath(new_filepath):
            if os.path.exists(new_filepath):
                os.replace(current_excel, new_filepath)
            else:
                os.rename(current_excel, new_filepath)
            print(f"✅ 文件已全自动进化更名为 ──> 【{new_filename}】")
        else:
            print(f"✅ 数据已在今日最新表格内完成原位热更新 ──> 【{new_filename}】")
        # -------------------------------------------------- #

        print("\n✨✨✨ 微信/网盘全自动托管模式运行成功！历史链路无缝闭环！ ✨✨✨")
        print("==========================================================================\n")

    except PermissionError:
        print(f"\n❌ 写入失败！系统报错：Permission denied")
    except Exception as e:
        print(f"❌ 运行出错了: {e}")


if __name__ == "__main__":
    main()
