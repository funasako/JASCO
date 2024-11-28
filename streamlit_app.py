import streamlit as st
import pandas as pd
import io

# ファイルアップロード
uploaded_file = st.file_uploader("テキストファイルをアップロードしてください", type=["txt"])

if uploaded_file is not None:
    # ファイルを読み込む
    content = uploaded_file.read().decode("shift_jis").splitlines()

    # XYデータを抽出
    xy_start = content.index("XYDATA") + 1
    xy_end = content.index("##### Extended Information") - 2
    xy_data_lines = content[xy_start:xy_end + 1]

    # データをデータフレームに変換
    data = [line.split() for line in xy_data_lines if line.strip()]
    df = pd.DataFrame(data, columns=["X", "Y"]).astype(float)

    # Excelファイルにデータとグラフを書き出す関数
    def convert_df_to_excel(df):
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Data', startrow=2, startcol=9)  # J3, K3セルにデータを配置

        workbook  = writer.book
        worksheet = writer.sheets['Data']

        # セルの高さを設定
        for i in range(0, len(df) + 3):  # データ行数分+ヘッダ行
            worksheet.set_row(i, 21)  # すべての行を21ピクセルに設定

        # グラフ作成
        chart = workbook.add_chart({'type': 'line'})

        # グラフデータを指定
        chart.add_series({
            'categories': ['Data', 2, 9, 2 + len(df) - 1, 9],  # X軸データ (J3から)
            'values':     ['Data', 2, 10, 2 + len(df) - 1, 10],  # Y軸データ (K3から)
            'name':       None,
        })

        # グラフスタイル設定
        chart.set_size({'width': 533, 'height': 377})  # 13.3cm x 10cm
        chart.set_chartarea({'fill': {'none': True}, 'border': {'none': True}})
        chart.set_plotarea({'fill': {'none': True}, 'border': {'color': 'black', 'width': 1.5}})
        
        # 軸設定
        chart.set_x_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'interval_unit': 100,
            'reverse': False,  # 軸の向き：右側が大きく
        })

        chart.set_y_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
        })

        # 凡例とタイトルを非表示
        chart.set_legend({'none': True})
        chart.set_title({'none': True})

        # グラフを挿入
        worksheet.insert_chart('A3', chart)

        writer.close()
        output.seek(0)
        return output

    # ダウンロードボタンの作成
    excel_data = convert_df_to_excel(df)
    st.download_button(label="Excelファイルをダウンロード", data=excel_data, file_name="data_with_chart.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # グラフを描画
    st.write("### グラフ表示")
    st.line_chart(df.set_index("X"))
