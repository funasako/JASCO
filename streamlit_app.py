import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import xlsxwriter
import datetime
import pytz


# タイトル等
st.set_page_config(page_title="JASCO UV-vis")
st.title("JASCO UV-vis Spectra Formatter")
st.write("1. スペクトルマネージャーでテキスト形式でエクスポート（ファイル名をしっかりつけておく）")
st.write("2. 以下にアップロードする。複数アップロードでプロット重ね書き")
st.write("3. Excelファイルをダウンロード")
st.write("")

# ファイルアップロード
uploaded_files = st.file_uploader(
    "テキストファイルをアップロード（複数可）",
    type=["txt"], 
    accept_multiple_files=True                         
)

# Excel列ずらし対応
def col_num_to_excel_col(n):
    """Convert a 0-based column number to Excel-style column label (e.g., 0 -> 'A', 27 -> 'AB')"""
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result

def convert_files_to_excel(files):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Data")
        writer.sheets["Data"] = worksheet

        # フォント設定用のフォーマットを作成
        cell_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12})
        border_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'border': 1})

        # Excelグラフの初期設定
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        chart.set_size({'width': 533, 'height': 377})
        chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
        chart.set_plotarea({'border': {'color': 'black', 'width': 1.5}, 'fill': {'none': True}})
        chart.set_legend({'none': True})
        chart.set_x_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'major_unit': 100,
            'min': 300,
            'max': 700,
            'reverse': False,
            'name': 'Wavelength / nm',
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},
        })
        chart.set_y_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'min': 0,
            'name': 'Absorbance',
            'major_gridlines': {'visible': False},
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},
        })
        
        start_col = 11  # 初期列（L列 = インデックス11）
        
        global_max_x = 700


        # すべてのデータを格納するリスト
        data_frames = []
        
        # 表示用グラフの作成
        fig, ax = plt.subplots(figsize=(8, 6))
        

        
        for file in files:
            # ファイル名とデータの読み取り
            content = file.read().decode("shift_jis").splitlines()
            xy_start = content.index("XYDATA") + 1
            xy_end = content.index("##### Extended Information") - 2
            xy_data_lines = content[xy_start:xy_end + 1]

            data = [line.split() for line in xy_data_lines if line.strip()]
            df = pd.DataFrame(data, columns=["X", "Y"]).astype(float)
            # データフレーム化
            data_frames.append(df)
    
            # グラフにプロットを追加
            ax.plot(df["X"], df["Y"], label=file.name, linewidth=1.5)

            
            # セルの高さを設定
            for i in range(len(df) + 3):
                worksheet.set_row(i, 20, cell_format)

            # ファイル名を記入
            worksheet.write(0, start_col, file.name, cell_format)

            # データを書き込む
            worksheet.write(1, start_col, 'WL', cell_format)
            worksheet.write(1, start_col + 1, 'Abs', cell_format)
            for i, (x, y) in enumerate(zip(df["X"], df["Y"])):
                worksheet.write(i + 2, start_col, x, cell_format)
                worksheet.write(i + 2, start_col + 1, y, cell_format)
                
            # 最大Xを更新（グラフの横軸最大値の設定）
            global_max_x = max(global_max_x, df["X"].max())

            # N列の計算式を設定
            worksheet.write(0, start_col + 2, 1, border_format)
            worksheet.write(1, start_col + 2, 0, border_format)
            worksheet.write_formula(2, start_col + 2, f"={col_num_to_excel_col(start_col + 1)}3*${col_num_to_excel_col(start_col + 2)}$1+${col_num_to_excel_col(start_col + 2)}$2", cell_format)
            for i in range(1, len(df)):
                worksheet.write_formula(i + 2, start_col + 2, f"={col_num_to_excel_col(start_col + 1)}{i+3}*${col_num_to_excel_col(start_col + 2)}$1+${col_num_to_excel_col(start_col + 2)}$2", cell_format)

            # グラフを作成

            chart.add_series({
                'categories': f"=Data!${col_num_to_excel_col(start_col)}$3:${col_num_to_excel_col(start_col)}${len(df) + 2}",
                'values': f"=Data!${col_num_to_excel_col(start_col + 2)}$3:${col_num_to_excel_col(start_col + 2)}${len(df) + 2}",
                'marker': {'type': 'none'},
                'line': {'color': '#008EC0', 'width': 1.5},
            })

            start_col += 4  # 次のファイルは右に4列ずらして書き込み

        chart.set_x_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'major_unit': 100,
            'min': 300,
            'max': global_max_x,
            'reverse': False,
            'name': 'Wavelength / nm',
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},
        })
        chart.set_size({'width': 460, 'height': 370})
        worksheet.insert_chart("A3", chart)

        # 表示用グラフの装飾
        ax.set_xlabel("Wavelength / nm", fontsize=12)
        ax.set_ylabel("Absorbance", fontsize=12)
        ax.set_xlim(300, global_max_x)  # Xの最大値を動的に設定
        ax.legend(loc="upper right", fontsize=10)
        ax.grid(True)

        # Streamlitでグラフを表示
        st.pyplot(fig)
    
    return output.getvalue()

if uploaded_files:
    # 現在の日本時間（JST）を取得
    japan_tz = pytz.timezone('Asia/Tokyo')
    current_time = datetime.datetime.now(japan_tz).strftime("%Y%m%d_%H%M%S")
    file_name = f"spectra_{current_time}.xlsx"
    
    # Excel変換とデータ保存
    excel_data = convert_files_to_excel(uploaded_files)
    st.download_button(
        label="Excelファイルをダウンロード",
        data=excel_data,
        file_name=file_name,  # 動的に生成したファイル名を指定
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

    
