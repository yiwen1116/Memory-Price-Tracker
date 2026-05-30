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
    """【Pandas刚性计算行号版】：彻底抛弃 openpyxl 测空行，用 Pandas 绝对公正的矩阵行数锁死行尾，全面粉碎环境变动 Bug"""
    target_row = None
    standard_web_date = datetime.strptime(real_date_str, "%Y-%m-%d").strftime("%Y-%m-%d")

    # 1. 🔍 固定使用第一列（A列）作为时间基准中轴
    base_date_col_idx = 1

    # 2. 🔍 扫描 Excel 现有的数据行，看网页返回的这个日期（比如 2026-05-29）在表里是否已经存在
    # 💡 这里的循环终点严格按照 Pandas 读出来的真实数据矩阵长度（len(df_history)）来算，绝不上当受骗！
    for idx in range(len(df_history)):
        # Excel 中的实际物理行号 = Pandas 矩阵索引 + 3 (因为前 2 行是双表头)
        current_excel_row = idx + 3
        raw_val = ws.cell(row=current_excel_row, column=base_date_col_idx).value
        
        if raw_val is not None:
            if isinstance(raw_val, datetime):
                cell_val = raw_val.strftime("%Y-%m-%d")
            else:
                try:
                    clean_str = str(raw_val).strip().split()[0].replace("/", "-")
                    cell_val = datetime.strptime(clean_str, "%Y-%m-%d").strftime("%Y-%m-%d")
                except:
                    cell_val = str(raw_val)

            # 如果在 Excel 里找到了和网页日期完全吻合的一行
            if cell_val == standard_web_date:
                target_row = current_excel_row
                break

    # 3. 🔀 刚性判定写入与加行位置
    if target_row is not None:
        # 情况一：网页是29号，表里有29号。下午第二次跑或者周末空跑，直接原地刷新最新价格，绝不产生30号的多余新行
        print(f"\n 📅 工作表【{ws.title}】：检测到数据日期 {standard_web_date} 已在表中存在 -> 执行【原位热更新】，避免多余行。")
    else:
        # 情况二：网页出现了全新的开盘日期！
        # 💡【终极刚性解】：全新空行 = Pandas 矩阵里现有的有效数据行数 + 双表头 2 行 + 1 (新的一行)
        target_row = len(df_history) + 2 + 1
        print(f"\n 📅 工作表【{ws.title}】：检测到全新的网页数据日期 -> 刚性计算锁启动，锁定全新空行追加: 第 {target_row} 行 ({standard_web_date})")

    # 4. ✍️ 填满这一行中所有的“日期”单元格
    dt_obj = datetime.strptime(standard_web_date, "%Y-%m-%d")
    formatted_date = f"{dt_obj.year}/{dt_obj.month}/{dt_obj.day}"

    for c_idx, col in enumerate(df_history.columns):
        if col[1] == '日期':
            ws.cell(row=target_row, column=c_idx + 1).value = formatted_date

    # 5. 💾 写入今日抓到的各个颗粒度具体价格
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
        today_compact = datetime.now().strftime("%Y%m%d")  # 例如：20260530
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
