import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import xlsxwriter

# ファイルアップロード
uploaded_file = st.file_uploader("テキストファイルをアップロードしてください", type=["txt"])

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Data")
        writer.sheets["Data"] = worksheet

        # データを書き込む (L3セルとM3セルから)
        worksheet.write('L2', 'Xデータ')
        worksheet.write('M2', 'Yデータ')
        for i, (x, y) in enumerate(zip(df["X"], df["Y"])):
            worksheet.write(i + 2, 11, x)  # L列 (インデックス11)
            worksheet.write(i + 2, 12, y)  # M列 (インデックス12)

        # セルの高さを設定
        for i in range(len(df) + 3):
            worksheet.set_row(i, 20)  # セルの高さを20ptに変更

        # グラフを作成
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        chart.add_series({
            'categories': f"=Data!$L$3:$L${len(df) + 2}",
            'values': f"=Data!$M$3:$M${len(df) + 2}",
            'marker': {'type': 'none'},
            'line': {'color': 'blue'},  # 単色の線
        })

        # グラフのプロパティ設定
        chart.set_size({'width': 533, 'height': 377})  # 幅13.3cm, 高さ10cm
        chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
        chart.set_plotarea({'border': {'color': 'black', 'width': 1.5}, 'fill': {'none': True}})
        
        # 凡例を非表示
        chart.set_legend({'none': True})

        # 横軸設定 (開始範囲を300に固定)
        chart.set_x_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'major_unit': 100,  # 横軸の目盛り間隔
            'min': 300,  # 横軸の開始範囲を300に固定
            'max': df['X'].max(),  # 横軸の最大値をXデータの最大値に設定
            'reverse': False,
            'name': 'Wavelength / nm',  # 横軸ラベルを設定
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},  # 数値のフォント設定
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},  # ラベルのフォント設定（太字解除）
        })

        # 縦軸設定 (0.1刻み)
        chart.set_y_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'name': 'Absorbance',  # 縦軸ラベルを設定
            'major_gridlines': {'visible': False},  # 縦軸の目盛線を削除
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},  # 数値のフォント設定
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},  # ラベルのフォント設定（太字解除）
            'major_unit': 0.1,  # 縦軸の目盛り間隔を0.1に設定
        })

        # グラフを配置
        worksheet.insert_chart('A3', chart)

    return output.getvalue()  # 自動的に保存されるので、明示的な保存は不要

if uploaded_file is not None:
    content = uploaded_file.read().decode("shift_jis").splitlines()
    xy_start = content.index("XYDATA") + 1
    xy_end = content.index("##### Extended Information") - 2
    xy_data_lines = content[xy_start:xy_end + 1]

    data = [line.split() for line in xy_data_lines if line.strip()]
    df = pd.DataFrame(data, columns=["X", "Y"]).astype(float)

    # グラフを描画
    st.write("### グラフ表示")
    fig, ax = plt.subplots()
    ax.plot(df["X"], df["Y"], linewidth=1.5, color='blue')  # 線の描画
    ax.set_xlabel("Wavelength / nm")
    ax.set_ylabel("Absorbance")
    ax.set_xlim(300, df["X"].max())  # 横軸の開始範囲を300に固定
    st.pyplot(fig)

    # データをテーブル表示
    st.write("### 抽出したデータ")
    st.dataframe(df)

    # Excelデータを作成しダウンロード
    excel_data = convert_df_to_excel(df)
    st.download_button(
        label="Excelファイルをダウンロード",
        data=excel_data,
        file_name='xydata.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
