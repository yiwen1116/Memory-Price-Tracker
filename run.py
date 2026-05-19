import io
import os
import re
import glob
import requests
import pandas as pd
from datetime import datetime
import openpyxl

# =======================================================
# 🎯 配置区：请确认你的 Excel 所在的【文件夹路径】（不要带文件名！）
# 例如放在 D 盘根目录就是 r"D:\"，放在某个文件夹就是 r"D:\内存数据\"
FOLDER_DIR = r"D:\TrendForce_Data"


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
    target_row = None

    for r in range(3, ws.max_row + 1):
        for c_idx, col in enumerate(df_history.columns):
            if col[1] == '日期':
                raw_val = ws.cell(row=r, column=c_idx + 1).value
                if raw_val is not None:
                    if isinstance(raw_val, datetime):
                        cell_val = raw_val.strftime("%Y-%m-%d")
                    else:
                        try:
                            clean_str = str(raw_val).strip().split()[0].replace("/", "-")
                            cell_val = datetime.strptime(clean_str, "%Y-%m-%d").strftime("%Y-%m-%d")
                        except ValueError:
                            cell_val = str(raw_val)

                    standard_real_date = datetime.strptime(real_date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
                    if cell_val == standard_real_date:
                        target_row = r
                        break
        if target_row:
            break

    if target_row is None:
        target_row = ws.max_row + 1
        print(f"\n 📅 工作表【{ws.title}】：将在末尾新增行 -> {real_date_str}")
    else:
        print(f"\n 📅 工作表【{ws.title}】：将在第 {target_row} 行原位更新数据 -> {real_date_str}")

    dt_obj = datetime.strptime(real_date_str, "%Y-%m-%d")
    formatted_date = f"{dt_obj.year}/{dt_obj.month}/{dt_obj.day}"

    for c_idx, col in enumerate(df_history.columns):
        if col[1] == '日期':
            ws.cell(row=target_row, column=c_idx + 1).value = formatted_date

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
        input("\n按回车键退出程序...")
        return

    # 按文件修改时间排序，自动锁定最新的一份文件
    current_excel = max(matched_files, key=os.path.getmtime)
    print(f"📂 成功锁定目标文件：【{os.path.basename(current_excel)}】\n")
    # -------------------------------------------------- #

    print("🚀 [1/4] 开始同步 TrendForce 官方数据...")
    try:
        tables_dram, dram_date = get_web_data_and_date("https://www.trendforce.cn/price")
        tables_flash, flash_date = get_web_data_and_date("https://www.trendforce.cn/price/flash")

        print("\n📦 [2/4] 正在进行内存数据对齐...")
        df_dram_hist = pd.read_excel(current_excel, sheet_name='DRAM Spot Price', header=[0, 1], engine='openpyxl')
        df_flash_hist = pd.read_excel(current_excel, sheet_name='NAND Flash', header=[0, 1], engine='openpyxl')

        dram_row_dict = get_row_data_dict(tables_dram, df_dram_hist)
        flash_row_dict = get_row_data_dict(tables_flash, df_flash_hist)

        print("💾 [3/4] 正在安全写入底层单元格...")
        wb = openpyxl.load_workbook(current_excel)
        write_to_excel_cell(wb['DRAM Spot Price'], df_dram_hist, dram_row_dict, dram_date)
        write_to_excel_cell(wb['NAND Flash'], df_flash_hist, flash_row_dict, flash_date)
        wb.save(current_excel)
        wb.close()

        # ---------------- 自动化模块：修改文件名 ---------------- #
        print("🏷️ [4/4] 正在执行智能文件重命名...")
        today_compact = datetime.now().strftime("%Y%m%d")  # 比如 20260519
        new_filename = f"内存价格每日追踪_{today_compact}.xlsx"
        new_filepath = os.path.join(FOLDER_DIR, new_filename)

        # 如果当前文件名不是今天的文件名，就执行重命名
        if os.path.abspath(current_excel) != os.path.abspath(new_filepath):
            if os.path.exists(new_filepath):
                # 如果今天的文件已存在，说明跑过第二次，直接静默覆盖替换
                os.replace(current_excel, new_filepath)
            else:
                os.rename(current_excel, new_filepath)
            print(f"✅ 文件已自动换上新衣：改名为 ──> 【{new_filename}】")
        else:
            print(f"✅ 文件名已经是今日最新 ──> 【{new_filename}】")
        # -------------------------------------------------- #

        print("\n✨✨✨ 全自动托管模式大功告成！今天的工作结束啦！ ✨✨✨")
        print("==========================================================================\n")

    except PermissionError:
        print(f"\n❌ 写入失败！系统报错：Permission denied")
        print(f"💡 解决办法：请关闭正在打开的 Excel 表格，然后再试一次！")
    except Exception as e:
        print(f"❌ 运行出错了: {e}")

    input("按回车键退出程序...")


if __name__ == "__main__":
    main()