import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

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

    # エクセルファイルに保存
    def convert_df_to_excel(dataframe):
        """データフレームをエクセルファイルに変換"""
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            dataframe.to_excel(writer, index=False, sheet_name="Sheet1")
        processed_data = output.getvalue()
        return processed_data

    # エクセルファイルのダウンロードリンクを作成
    excel_data = convert_df_to_excel(df)
    st.download_button(
        label="エクセルファイルをダウンロード",
        data=excel_data,
        file_name="xy_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
