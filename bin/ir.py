import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import xlsxwriter
import datetime
import pytz
import os
import math


# タイトル等
st.set_page_config(page_title="IR | JASCO Spectra Formatter", page_icon=":bar_chart:", )
st.title("IR | JASCO Spectra Formatter")
st.write("※動作にはインターネット接続が必要です。")
st.write("1. 装置が書き出したテキスト形式ファイルを用意する、もしくは、スペクトルマネージャーでテキストファイルをエクスポートする（ファイル名をしっかりつけておく）")
st.write("2. 以下にドラッグ&ドロップしてグラフ表示。複数ファイルからプロット重ね書きも可能")
st.write("3. Excelファイルをダウンロード")
st.write("4. 別のExcelファイルを作成する場合は、ページを再読込するかアップロード済みファイルをすべて✕ボタンで削除する")
st.write("")

# 変数初期化
xlsxYmin = None

# ファイルアップロード
uploaded_files = st.file_uploader(
    "エクスポートしたtxtファイルをアップロード（複数可）",
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

# テキストファイルのデータ構造分析
def extract_xy_data(content):
    xy_start = content.index("XYDATA") + 1 #開始行は固定
    xy_end = None #終了行は以下のように分岐    
    # '##### Extended Information'があれば、その2行上
    extended_info_index = next((i for i, line in enumerate(content) if '##### Extended Information' in line), None)
    if extended_info_index is not None:
        xy_end = extended_info_index - 2  # 2行上にする
    else:
        # 空行があれば、その1行上
        empty_line_index = next((i for i, line in enumerate(content) if line.strip() == ""), None)
        if empty_line_index is not None:
            xy_end = empty_line_index - 1  # 1行上にする
        else:
            # 上記どちらでもない場合は、ファイルの最終行
            xy_end = len(content) - 1  # 最終行
    
    # データを抽出
    xy_data_lines = content[xy_start:xy_end + 1]
    
    return xy_data_lines

def convert_files_to_excel(files):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Data")
        writer.sheets["Data"] = worksheet

        # フォント設定用のフォーマットを作成
        cell_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 11})
        border_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 11, 'border': 1})
        filename_format = workbook.add_format({'font_color': 'blue', 'font_name': 'Times New Roman', 'font_size': 11})

        # Excelグラフの初期設定
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        chart.set_size({'width': 533, 'height': 377})
        chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
        chart.set_plotarea({'border': {'color': 'black', 'width': 1.5}, 'fill': {'none': True}})
        chart.set_legend({'none': True})
        chart.set_title({'none': True}) 
        
        start_col = 11  # 初期列（L列 = インデックス11）
   
        # すべてのデータを格納するリスト
        data_frames = []
        
        # 表示用グラフの作成
        fig, ax = plt.subplots(figsize=(8, 8))
        
        # ファイル処理ループの前に、ファイル数を取得
        num_files = len(uploaded_files)
        now_file = 1
        
        for i, file in enumerate(uploaded_files):
            # ファイル名とデータの読み取り
            content = file.read().decode("shift_jis").splitlines()
            
            # XYデータを抽出してDataFrameに追加
            xy_data_lines = extract_xy_data(content)
            data = [line.split() for line in xy_data_lines if line.strip()]
            df = pd.DataFrame(data, columns=["X", "Y"]).astype(float)
            data_frames.append(df)

            # %T最小値を保持
            xlsxYmin = math.floor(df.loc[104:, "Y"].min() / 10) * 10 - 10
                       
            # セルの高さを設定
            for i in range(len(df) + 3):
                worksheet.set_row(i, 20, cell_format)

            # ファイル名を記入
            worksheet.write(0, start_col, file.name, filename_format)

            # データを書き込む
            worksheet.write(1, start_col, 'WL', cell_format) # X項目名
            worksheet.write(1, start_col + 1, '%T', cell_format) # Y項目名
            for i, (x, y) in enumerate(zip(df["X"], df["Y"])):
                worksheet.write(i + 2, start_col, x, cell_format) # Xデータ
                worksheet.write(i + 2, start_col + 1, y, cell_format) # Yデータ

            # グラフ表示用Yデータ列の補正定数の設定：先に処理したものから上へ40ずつずらす
            overlayconst = (num_files - now_file) * 40
            
            # グラフ表示用Yデータ列の計算式を設定
            worksheet.write(0, start_col + 2, 1, border_format)
            worksheet.write(1, start_col + 2, overlayconst, border_format)

            # 順序カウンタ増加
            now_file += 1
            
            # stlite対応の文字列操作
            col1 = col_num_to_excel_col(start_col + 1)
            col2 = col_num_to_excel_col(start_col + 2)
            formula = "=" + col1 + "3*$" + col2 + "$1+$" + col2 + "$2"
            worksheet.write_formula(2, start_col + 2, formula, cell_format)
            
            # 列名を事前に計算
            col1 = col_num_to_excel_col(start_col + 1)
            col2 = col_num_to_excel_col(start_col + 2)
            col_categories = col_num_to_excel_col(start_col)

            # ループ内のformula
            for i in range(1, len(df)):
                formula = "=" + col1 + str(i + 3) + "*$" + col2 + "$1+$" + col2 + "$2"
                worksheet.write_formula(i + 2, start_col + 2, formula, cell_format)
                
            # グラフを作成
            # stlite対応の文字列操作
            categories_range = "=Data!$" + col_categories + "$3:$" + col_categories + "$" + str(len(df) + 2)
            values_range = "=Data!$" + col2 + "$3:$" + col2 + "$" + str(len(df) + 2)

            # ファイル名から拡張子を除去
            filename_noext = os.path.splitext(file.name)[0]
            # Excelにプロット追加
            chart.add_series({
                'categories': categories_range,
                'values': values_range,
                'name': filename_noext,
                'marker': {'type': 'none'},
                'line': {'color': '#008EC0', 'width': 1.5},
                'title': None,
            })
            
            # 表示用グラフにプロットを追加
            df["Y"] = df["Y"] + overlayconst
            ax.plot(df["X"], df["Y"], label=file.name, linewidth=1.5)
            
            start_col += 4  # 次のファイルは右に4列ずらして書き込み
            # ループ終了

        # Excelグラフ書式修正
        chart.set_x_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'major_unit': 500,
            'min': 500,
            'max': 4000,
            'reverse': True,
            'crossing': 'max',
            'name': 'Wavenumber / cm–1',
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},
        })
        chart.set_y_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'none',  # 主目盛を非表示
            'minor_tick_mark': 'none',  # 補助目盛を非表示
            'label_position': 'none',  # ラベルを非表示
            'max': (num_files - 1) * 40 + 110,
            'min': xlsxYmin,
            'crossing': -1000,
            'name': 'Transmittance (%)',
            'major_gridlines': {'visible': False},
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},
        })
        # プロットエリアの位置を調整
        chart.set_plotarea({
            'layout': {'x': 0.1},
            'border': {'color': 'black', 'width': 1.5}, 
            'fill': {'none': True}
        })
        chart.set_size({'width': 460, 'height': 370 + 80 * num_files})
        chart.set_legend({'none': True})
        worksheet.insert_chart("A4", chart)

        # 表示用グラフの装飾
        ax.set_xlabel(r'$\mathrm{Wavenumber / cm^{-1}}$', fontsize=12)
        ax.set_ylabel("Transmittance (%)", fontsize=12)
        ax.set_xlim(500, 4000) 
        ax.set_ylim(xlsxYmin, (num_files - 1) * 40 + 110) 
        ax.legend(loc="lower left", fontsize=10)
        ax.grid(True)
        ax.invert_xaxis()
        # Streamlitでグラフを表示
        st.pyplot(fig)
    
    return output.getvalue()

if uploaded_files:
    # 現在の日本時間（JST）を取得
    japan_tz = pytz.timezone('Asia/Tokyo')
    current_time = datetime.datetime.now(japan_tz).strftime("%Y%m%d_%H%M%S")
    file_name = f"IR_{current_time}.xlsx"
    
    # Excel変換とデータ保存
    excel_data = convert_files_to_excel(uploaded_files)
    st.download_button(
        label="Excelファイルをダウンロード",
        data=excel_data,
        file_name=file_name,  # 動的に生成したファイル名を指定
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

