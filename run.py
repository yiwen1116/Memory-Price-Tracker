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
    """
    更正后的核心写入逻辑：
    1. 【同一天跑多次】：如果真实系统今天这一行已存在，直接原位更新，最新一次覆盖。
    2. 【网页未更新】：如果网站抓出的日期与 Excel 里的最后一行日期相同，则绝对不添加新行，静默退出！
    """
    target_row = None
    system_today_str = datetime.now().strftime("%Y-%m-%d")

    # 1. 🔍 获取当前 Excel 里的最后一行日期，用来和网页做对比
    last_row_idx = ws.max_row
    last_row_date_str = None
    
    # 寻找最后一行的日期所在位置
    for c_idx, col in enumerate(df_history.columns):
        if col[1] == '日期':
            last_cell_val = ws.cell(row=last_row_idx, column=c_idx + 1).value
            if last_cell_val is not None:
                if isinstance(last_cell_val, datetime):
                    last_row_date_str = last_cell_val.strftime("%Y-%m-%d")
                else:
                    try:
                        clean_str = str(last_cell_val).strip().split()[0].replace("/", "-")
                        last_row_date_str = datetime.strptime(clean_str, "%Y-%m-%d").strftime("%Y-%m-%d")
                    except:
                        last_row_date_str = str(last_cell_val)
            break

    # 2. 🛡️ 【拦截器】如果网页返回的更新日期和 Excel 最后一行一模一样 -> 判定为未更新，坚决不加行！
    standard_web_date = datetime.strptime(real_date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
    if last_row_date_str == standard_web_date and system_today_str != last_row_date_str:
        print(f"\n 🛑 工作表【{ws.title}】：检测到网站上依然是老数据 ({real_date_str}) -> 触发【防空跑机制】：不添加任何新行。")
        return False # 返回 False 代表今天没有发生任何数据写入行为

    # 3. 🔍 扫描 Excel，看今天系统这一行是否在下午已经生成过
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

                    if cell_val == system_today_str:
                        target_row = r
                        break
        if target_row:
            break

    # 4. 🔀 决定是原位更新，还是新增今日行
    if target_row is not None:
        print(f"\n 📅 工作表【{ws.title}】：今日下午已跑过 -> 正在执行【原位更新】，18:10 最新数据覆盖 14:40 旧数据。")
    else:
        target_row = ws.max_row + 1
        print(f"\n 📅 工作表【{ws.title}】：检测到网页今日已发布新数据 -> 正在末尾【新增今日数据行】: {system_today_str}")

    # 5. 写入今天的日期
    dt_obj = datetime.strptime(system_today_str, "%Y-%m-%d")
    formatted_date = f"{dt_obj.year}/{dt_obj.month}/{dt_obj.day}"

    for c_idx, col in enumerate(df_history.columns):
        if col[1] == '日期':
            ws.cell(row=target_row, column=c_idx + 1).value = formatted_date

    # 6. 填入数据
    for col_idx, value in row_data_dict.items():
        if value is not None and str(value) != 'nan':
            try:
                ws.cell(row=target_row, column=col_idx + 1).value = float(value)
            except ValueError:
                ws.cell(row=target_row, column=col_idx + 1).value = value
    
    return True # 返回 True 代表发生了有效的修改


def main():
    print("\n==================================================")

    # ---------------- 自动化模块：寻找文件 ---------------- #
    print("🔍 [0/4] 自动雷达：正在定位本地最新的 Excel 追踪表...")
    search_pattern = os.path.join(FOLDER_DIR, "内存价格每日追踪_*.xlsx")
    matched_files = glob.glob(search_pattern)

    if not matched_files:
        print(f"❌ 错误：在 {FOLDER_DIR} 下找不到名字包含【内存价格每日追踪_】的表格！")
        return

    current_excel = max(matched_files, key=os.path.getmtime)
    print(f"📂 成功继承历史火种，锁定目标文件：【{os.path.basename(current_excel)}】\n")
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
        
        # 运行写入并获取是否有有效的数据修改
        changed_dram = write_to_excel_cell(wb['DRAM Spot Price'], df_dram_hist, dram_row_dict, dram_date)
        changed_flash = write_to_excel_cell(wb['NAND Flash'], df_flash_hist, flash_row_dict, flash_date)
        
        if not changed_dram and not changed_flash:
            wb.close()
            print("\n ☕ 【智能休眠】：由于网站上未更新任何价格数据，本次运行未做任何修改，不重命名文件。")
            print("✨✨✨ 自动化流程提前安全退出。 ✨✨✨\n")
            return

        wb.save(current_excel)
        wb.close()

        # ---------------- 自动化模块：修改文件名 ---------------- #
        print("🏷️ [4/4] 正在执行智能文件升级与迭代...")
        today_compact = datetime.now().strftime("%Y%m%d")  # 例如：20260528
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
        print(f"💡 解决办法：请关闭你电脑里正在打开的 Excel 表格，然后再运行一次！")
    except Exception as e:
        print(f"❌ 运行出错了: {e}")


if __name__ == "__main__":
    main()
