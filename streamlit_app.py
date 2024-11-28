import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO  # バイトデータ用のライブラリ

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

    # エクセルファイルにデータを書き込む
    output = BytesIO()  # バイトストリームを作成
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='XY Data')  # Excelに書き込み
        writer.save()
    
    # ダウンロードボタンの作成
    st.download_button(
        label="エクセルファイルをダウンロード",
        data=output.getvalue(),
        file_name="xy_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
