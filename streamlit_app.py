import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import xlsxwriter

# 複数ファイルアップロード
uploaded_files = st.file_uploader("テキストファイルをアップロードしてください", type=["txt"], accept_multiple_files=True)

def convert_df_to_excel(files_data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Data")
        writer.sheets["Data"] = worksheet

        # フォント設定用のフォーマットを作成
        cell_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12})
        border_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'border': 1})

        x_offset = 11  # 初期のL列（Xデータの列）

        # 各ファイルに対して処理を行う
        for file_idx, (file_name, df) in enumerate(files_data):
            # L1セルにファイル名を記入
            worksheet.write(f'L{file_idx * 3 + 1}', file_name, cell_format)

            # Xデータ、Yデータを記入
            worksheet.write(f'L{file_idx * 3 + 2}', 'WL', cell_format)  # Xデータの列名
            worksheet.write(f'M{file_idx * 3 + 2}', 'Abs', cell_format)  # Yデータの列名

            for i, (x, y) in enumerate(zip(df["X"], df["Y"])):
                # x と y を数値に変換する
                try:
                    x = float(x)
                    y = float(y)
                except ValueError:
                    continue  # 変換できない場合はスキップ

                worksheet.write(i + 2 + file_idx * len(df), x, cell_format)  # Xデータ
                worksheet.write(i + 2 + file_idx * len(df), x_offset + 1, y, cell_format)  # Yデータ

            # N列（補正値）の計算式を設定
            worksheet.write(f'N{file_idx * 3 + 2}', 1, border_format)  # N1セルに1
            worksheet.write(f'N{file_idx * 3 + 3}', 0, border_format)  # N2セルに0
            worksheet.write_formula(f'N{file_idx * 3 + 4}', f"=M{file_idx * 3 + 4}*$N${file_idx * 3 + 2}+$N${file_idx * 3 + 3}", cell_format)  # N3セル以降
            for i in range(1, len(df)):
                worksheet.write_formula(f'N{i + 2 + file_idx * len(df)}', f"=M{i + 3 + file_idx * len(df)}*$N${file_idx * 3 + 2}+$N${file_idx * 3 + 3}", cell_format)

            # 次のデータが追加される列（O, P, Q...）
            x_offset += 4  # 1列分（L, M）を使用した後、次はO, P, Qにデータを追加

        # グラフを作成
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})

        # 各ファイルのXデータと補正されたYデータをグラフに追加
        for file_idx, (file_name, df) in enumerate(files_data):
            x_start = 11 + 3 * file_idx  # 各ファイルに対応するXデータの列
            y_start = 12 + 3 * file_idx  # 各ファイルに対応する補正後のYデータの列
            chart.add_series({
                'categories': f"=Data!${chr(65 + x_start)}$3:${chr(65 + x_start)}${len(df) + 2}",
                'values': f"=Data!${chr(65 + y_start)}$3:${chr(65 + y_start)}${len(df) + 2}",
                'marker': {'type': 'none'},
                'line': {'color': '#008EC0', 'width': 1.5},  # 線の色と太さ
            })

        # グラフのプロパティ設定
        chart.set_size({'width': 533, 'height': 377})  # 幅13.3cm, 高さ10cm
        chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
        chart.set_plotarea({'border': {'color': 'black', 'width': 1.5}, 'fill': {'none': True}})
        chart.set_legend({'none': True})  # 凡例を非表示

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

    return output.getvalue()



if uploaded_files:
    files_data = []
    for uploaded_file in uploaded_files:
        content = uploaded_file.read().decode("shift_jis").splitlines()
        xy_start = content.index("XYDATA") + 1
        xy_end = content.index("##### Extended Information") - 2
        xy_data_lines = content[xy_start:xy_end + 1]
        data = [line.split() for line in xy_data_lines if line.strip()]
        df = pd.DataFrame(data, columns=["X", "Y"]).astype(float)
        files_data.append((uploaded_file.name, df))

    # グラフを描画
    st.write("### グラフ表示")
    fig, ax = plt.subplots()
    for file_idx, (file_name, df) in enumerate(files_data):
        ax.plot(df["X"], df["Y"], linewidth=1.5, color='#008EC0')  # プロットの線の色を#008EC0に変更
    ax.set_xlabel("Wavelength / nm")
    ax.set_ylabel("Absorbance")
    ax.set_xlim(300, df["X"].max())  # 横軸の開始範囲を300に固定
    st.pyplot(fig)

    # Excelデータを作成しダウンロード
    excel_filename = uploaded_files[0].name.replace(".txt", ".xlsx")  # 1つ目のファイル名を使用
    excel_data = convert_df_to_excel(files_data)
    st.download_button(
        label="Excelファイルをダウンロード",
        data=excel_data,
        file_name=excel_filename,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    
