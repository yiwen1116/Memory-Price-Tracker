import os
import glob
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

def load_latest_data():
    search_pattern = os.path.join(".", "内存价格每日追踪_*.xlsx")
    matched_files = [f for f in glob.glob(search_pattern) if not os.path.basename(f).startswith('~')]
    if not matched_files:
        raise FileNotFoundError("未找到任何 Excel 数据文件。")
    latest_excel = max(matched_files)
    
    df_dram = pd.read_excel(latest_excel, sheet_name='DRAM Spot Price', header=[0, 1], engine='openpyxl')
    df_flash = pd.read_excel(latest_excel, sheet_name='NAND Flash', header=[0, 1], engine='openpyxl')
    return df_dram, df_flash, os.path.basename(latest_excel)

def preprocess_df(df):
    # 展平多级表头并提取日期
    df.columns = [f"{col[0]} - {col[1]}" if 'Unnamed' not in str(col[0]) else col[1] for col in df.columns]
    date_col = [col for col in df.columns if '日期' in col][0]
    df[date_col] = pd.to_datetime(df[date_col])
    df = df.sort_values(date_col)
    df.set_index(date_col, inplace=True)
    return df

def generate_html():
    df_dram_raw, df_flash_raw, filename = load_latest_data()
    df_dram = preprocess_df(df_dram_raw)
    df_flash = preprocess_df(df_flash_raw)

    # 提取带有“盘平均”的列进行展示
    dram_cols = [col for col in df_dram.columns if '盘平均' in col]
    flash_cols = [col for col in df_flash.columns if '盘平均' in col]

    fig = make_subplots(rows=2, cols=1, subplot_titles=(f"DRAM 盘平均价格走势", f"NAND Flash 盘平均价格走势"), vertical_spacing=0.1)

    for col in dram_cols:
        fig.add_trace(go.Scatter(x=df_dram.index, y=df_dram[col], mode='lines+markers', name=col), row=1, col=1)
        
    for col in flash_cols:
        fig.add_trace(go.Scatter(x=df_flash.index, y=df_flash[col], mode='lines+markers', name=col), row=2, col=1)

    fig.update_layout(
        title_text=f"半导体存储器价格高频追踪 (数据源: {filename})",
        height=1000,
        hovermode="x unified",
        template="plotly_white"
    )
    
    # 导出为静态网页
    fig.write_html("index.html")
    print("✅ 静态看板 index.html 生成成功！")

if __name__ == "__main__":
    generate_html()
