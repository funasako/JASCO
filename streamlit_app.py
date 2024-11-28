import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# ファイルアップロード
uploaded_file = st.file_uploader("テキストファイルをアップロードしてください", type=["txt"])

if uploaded_file is not None:
    # ファイルを読み込む
    content = uploaded_file.read().decode("shift_jis").splitlines()

    # XYデータを抽出
    xy_start = content.index("XYDATA") + 1
    xy_end = content.index("##### Extended Information") - 2  # 2行上まで含める
    xy_data_lines = content[xy_start:xy_end + 1]

    # データをデータフレームに変換
    data = [line.split() for line in xy_data_lines if line.strip()]
    df = pd.DataFrame(data, columns=["X", "Y"]).astype(float)

    # グラフを描画
    st.write("### グラフ表示")
    fig, ax = plt.subplots()
    ax.plot(df["X"], df["Y"], label="XY Data")
    ax.set_xlabel("Wavelength / nm")
    ax.set_ylabel("Absorbance")
    st.pyplot(fig)

    # データをテーブル表示
    st.write("### 抽出したデータ")
    st.dataframe(df)

    # Excelファイルにデータとグラフを書き出す
    def convert_df_to_excel(df):
        output = pd.ExcelWriter("output.xlsx", engine='xlsxwriter')
        df.to_excel(output, index=False, sheet_name='Data')

        workbook  = output.book
        worksheet = output.sheets['Data']

        # グラフを作成
        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({
            'categories': ['Data', 1, 0, len(df), 0],  # X軸（1列目）
            'values':     ['Data', 1, 1, len(df), 1],  # Y軸（2列目）
            'name':       'XY Data',
        })

        # グラフをシートに挿入
        worksheet.insert_chart('D2', chart)  # グラフをD2セルに挿入

        output.close()

        # Excelファイルを読み込み、バイト形式に変換
        with open("output.xlsx", "rb") as file:
            return file.read()

    # ダウンロードボタンの作成
    excel_data = convert_df_to_excel(df)
    st.download_button(label="Excelファイルをダウンロード", data=excel_data, file_name="data_with_chart.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
